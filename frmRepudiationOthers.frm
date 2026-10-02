VERSION 5.00
Begin VB.Form frmRepudiationOthers 
   Caption         =   "RC6-Others (Specify)"
   ClientHeight    =   3495
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   3510
   LinkTopic       =   "Form2"
   ScaleHeight     =   3495
   ScaleWidth      =   3510
   StartUpPosition =   2  'CenterScreen
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
         Caption         =   "RC6-Others (Specify)"
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
   Begin VB.CommandButton cmdCancel 
      Caption         =   "Cancel"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1920
      TabIndex        =   8
      Top             =   3000
      Width           =   1455
   End
   Begin VB.CommandButton cmdGenerate 
      Caption         =   "Generate"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   120
      TabIndex        =   7
      Top             =   3000
      Width           =   1455
   End
   Begin VB.TextBox txtTitlecode 
      Height          =   285
      Left            =   240
      TabIndex        =   6
      Top             =   960
      Width           =   3135
   End
   Begin VB.OptionButton opt31 
      Caption         =   "Within 31 days"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   1920
      TabIndex        =   5
      Top             =   2280
      Width           =   1215
   End
   Begin VB.OptionButton optOut 
      Caption         =   "Out-patient"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1920
      TabIndex        =   4
      Top             =   1800
      Width           =   1335
   End
   Begin VB.OptionButton optDoc 
      Caption         =   "Letter to doctor"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   240
      TabIndex        =   3
      Top             =   2280
      Width           =   1095
   End
   Begin VB.OptionButton optCase 
      Caption         =   "Case Before post effective date"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   240
      TabIndex        =   2
      Top             =   1800
      Width           =   1815
   End
   Begin VB.Label Label2 
      Caption         =   "Please select available document template"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   240
      TabIndex        =   1
      Top             =   1440
      Width           =   3495
   End
   Begin VB.Label Label1 
      Caption         =   "Please specify the title of document"
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
      TabIndex        =   0
      Top             =   720
      Width           =   2895
   End
End
Attribute VB_Name = "frmRepudiationOthers"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim objword As Word.Application
Dim objdoc As Word.Document
Dim bookmarktype As String 'to see which document to be edited




Private Sub cmdCancel_Click()
Me.Hide
frmRepudiation.Show
End Sub

Private Sub cmdGenerate_Click()
Call funcword
End Sub

Private Sub funcword()
On Error GoTo openerror 'if you generate a doc then close it then generate again error occurs
                        ' so this trap will rectify this problem
Dim path As String



If objword Is Nothing Then
    Set objword = New Word.Application
    objword.Visible = True
    path = openword() 'to determine the type of document template
    Set objdoc = objword.Documents.Add(App.path & "\template\" & path & ".dot")

ElseIf Not objdoc Is Nothing Then
        If MsgBox("Previous document " & objdoc & " not closed. Press OK to save and close it now or cancel to abort", _
            vbExclamation + vbOKCancel, DialogBoxTitle) = vbOK Then
            objdoc.Save
            objdoc.Close
        ElseIf vbCancel Then
          Exit Sub
        End If
       Set objdoc = Nothing
       path = openword()
       Set objdoc = objword.Documents.Add(App.path & "\template\" & path & ".dot")
End If
Call bookmark 'call word then pump in the data
  
  
form_load_exit:
Exit Sub

openerror:

    If Err.Number = "5356" Then
    MsgBox Err.Description & ". Please close the previous document first using Word or save it to other name", vbExclamation + vbOKOnly, DialogBoxTitle
    
    ElseIf Err.Number = "462" Then
        Set objword = New Word.Application
        objword.Visible = True
        path = openword()
        Set objdoc = objword.Documents.Add(App.path & "\template\" & path & ".dot")
        Call bookmark
          
    ElseIf Err.Number = "-2147417848" Then 'this is automation error..when you close the doc(not word) then generate again
        path = openword()
        Set objdoc = objword.Documents.Add(App.path & "\template\" & path & ".dot")
        Call bookmark
    ElseIf Err.Number = "4198" Then
    
    Else
        MsgBox Err.Number & "-" & Err.Description, vbOKOnly + vbExclamation, DialogBoxTitle
    End If
    
End Sub

Private Sub bookmark()
Select Case bookmarktype
  Case "CASE"
    If bookmarkCASE = True Then
    Call printDoc
    End If
  Case "DOC"
    If bookmarkDOC = True Then
    Call printDoc
    End If
  
  Case "OUT"
    If bookmarkOUT = True Then
    Call printDoc
    End If
  
  Case "WITHIN"
    If bookmarkWITHIN = True Then
    Call printDoc
    End If
  
  End Select
  
  
  

End Sub
Private Sub printDoc()

            objdoc.SaveAs (App.path & "\template\repudiation\" & Trim(frmRepudiation.txtCAScode) & "-" & Trim(frmRepudiation.txtRepudiationCode) & ".doc")
           Me.Show
           If MsgBox("Do you want to print this document?", vbQuestion + vbYesNo, DialogBoxTitle) = vbYes Then
                'objdoc.SaveAs (App.path & "\repudiation\" & Trim(txtCAScode) & ".doc")
                'objdoc.Close
                'objdoc.PrintOut
                MsgBox "printed"
                frmRepudiationOthers.Hide
                frmRepudiation.Show
                Exit Sub
           Else
               frmRepudiationOthers.Hide
               frmRepudiation.Show
               
               Exit Sub
           End If


End Sub

Private Function openword() As String
    If optCase.Value = True Then
        openword = "CASE BEFORE PO. EFFEC DATE"
        bookmarktype = "CASE"
    ElseIf optDoc.Value = True Then
        openword = "Letter to doc"
        bookmarktype = "DOC"
    ElseIf optOut.Value = True Then
        openword = "Out-Patient"
        bookmarktype = "OUT"
    ElseIf opt31.Value = True Then
        openword = "Within 31 days"
        bookmarktype = "WITHIN"
    End If


End Function


Private Function bookmarkCASE()
    objdoc.Bookmarks("mbmnumber").Range.Text = frmRepudiation.txtMBMcode
    objdoc.Bookmarks("policynumber").Range.Text = frmRepudiation.txtPolicyCode
    objdoc.Bookmarks("insuredname").Range.Text = frmRepudiation.txtINScode
    objdoc.Bookmarks("patientname").Range.Text = frmRepudiation.txtPatientCode
    objdoc.Bookmarks("hospital").Range.Text = frmRepudiation.txtHospitalCode
    objdoc.Bookmarks("reason").Range.Text = frmRepudiation.txtReasonCode
    objdoc.Bookmarks("title").Range.Text = UCase(txtTitlecode)
    bookmarkCASE = True
End Function


Private Function bookmarkDOC()
Dim list As New ADODB.Recordset
Dim sql As String
sql = "select * from view_cas_repudiation where casnumber='" & frmRepudiation.txtCAScode & "'"
list.Open sql, medixmasterconn, adOpenKeyset, adLockOptimistic


objdoc.Bookmarks("patientname").Range.Text = frmRepudiation.txtPatientCode
objdoc.Bookmarks("date").Range.Text = frmRepudiation.txtDateCode
'objdoc.Bookmarks("invoicenumber").Range.Text
objdoc.Bookmarks("ic").Range.Text = list!MBMIcBcPp
objdoc.Bookmarks("reason").Range.Text = frmRepudiation.txtReasonCode
objdoc.Bookmarks("title").Range.Text = UCase(txtTitlecode)
bookmarkDOC = True
list.Close
End Function


Private Function bookmarkOUT()
objdoc.Bookmarks("mbmnumber").Range.Text = frmRepudiation.txtMBMcode
objdoc.Bookmarks("policynumber").Range.Text = frmRepudiation.txtPolicyCode
objdoc.Bookmarks("insuredname").Range.Text = frmRepudiation.txtMBMcode
objdoc.Bookmarks("patientname").Range.Text = frmRepudiation.txtPatientCode
objdoc.Bookmarks("hospital").Range.Text = frmRepudiation.txtHospitalCode
objdoc.Bookmarks("date").Range.Text = frmRepudiation.txtDateCode
objdoc.Bookmarks("title").Range.Text = UCase(txtTitlecode)
bookmarkOUT = True
End Function


Private Function bookmarkWITHIN()
objdoc.Bookmarks("mbmnumber").Range.Text = frmRepudiation.txtMBMcode
objdoc.Bookmarks("policynumber").Range.Text = frmRepudiation.txtPolicyCode
objdoc.Bookmarks("insuredname").Range.Text = frmRepudiation.txtINScode
objdoc.Bookmarks("patientname").Range.Text = frmRepudiation.txtPatientCode
objdoc.Bookmarks("reason").Range.Text = frmRepudiation.txtReasonCode
objdoc.Bookmarks("hospital").Range.Text = frmRepudiation.txtHospitalCode
objdoc.Bookmarks("title").Range.Text = UCase(txtTitlecode)
bookmarkWITHIN = True

End Function

Private Sub Form_activate()
txtTitlecode.SetFocus
End Sub
