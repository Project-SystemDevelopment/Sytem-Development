VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Begin VB.Form frmZurichCardPrinting 
   Caption         =   "frmZurichCardPrinting"
   ClientHeight    =   4350
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   10155
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   4350
   ScaleWidth      =   10155
   WindowState     =   2  'Maximized
   Begin VB.Frame Frame13 
      Height          =   615
      Left            =   4680
      TabIndex        =   3
      Top             =   0
      Width           =   5295
      Begin VB.CommandButton Exit 
         Caption         =   "Exit"
         Height          =   315
         Left            =   3120
         TabIndex        =   10
         Top             =   200
         Width           =   840
      End
      Begin VB.CommandButton CmdPrint 
         Caption         =   "&Print Err List"
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
         Index           =   1
         Left            =   3960
         TabIndex        =   9
         Top             =   200
         Width           =   1080
      End
      Begin VB.CommandButton Clear 
         Caption         =   "Clear"
         Height          =   315
         Left            =   2040
         TabIndex        =   8
         Top             =   200
         Width           =   1080
      End
      Begin VB.CommandButton cmdVerify 
         Caption         =   "&Verify"
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
         Left            =   120
         TabIndex        =   7
         Top             =   200
         Width           =   1080
      End
      Begin VB.CommandButton cmdUpload 
         Caption         =   "&Upload"
         Enabled         =   0   'False
         Height          =   315
         Left            =   1080
         TabIndex        =   6
         Top             =   200
         Width           =   1080
      End
      Begin VB.CommandButton CmdPrint 
         Caption         =   "&Print Err List"
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
         Index           =   0
         Left            =   1920
         TabIndex        =   5
         Top             =   600
         Width           =   1080
      End
   End
   Begin VB.CommandButton cmdFileLocation 
      Caption         =   " >"
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
      Left            =   4125
      TabIndex        =   1
      Top             =   150
      Width           =   320
   End
   Begin VB.TextBox txtImportFilename 
      Enabled         =   0   'False
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
      Left            =   1395
      TabIndex        =   0
      Top             =   120
      Width           =   2625
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   8400
      Top             =   120
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin MSComctlLib.ListView lstERRList 
      Height          =   3375
      Left            =   0
      TabIndex        =   4
      Top             =   600
      Width           =   9975
      _ExtentX        =   17595
      _ExtentY        =   5953
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
         Text            =   "Line"
         Object.Width           =   970
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
         Object.Width           =   8819
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   4
         Object.Width           =   2540
      EndProperty
   End
   Begin VB.Label Label1 
      Caption         =   "Filename"
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
      Left            =   120
      TabIndex        =   2
      Top             =   195
      Width           =   975
   End
End
Attribute VB_Name = "frmZurichCardPrinting"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim NoOfRecLine As Integer
Dim RecordArray() As String
Dim vRecordArray() As String
Dim totcount As Integer
'lstERRList.ListItems.Clear


Private Sub Clear_Click()
txtImportFilename.Text = ""
lstERRList.ListItems.Clear
End Sub

Private Sub cmdFileLocation_Click()
 CommonDialog1.CancelError = True
  On Error GoTo ErrHandler
  CommonDialog1.InitDir = "\\svr-doc\" & "Bor_In"
  CommonDialog1.Flags = cdlOFNHideReadOnly
  CommonDialog1.Filter = "All Files (*.*)|*.*|Text Files (*.txt)|*.txt|Dat Files (*.Dat)|*.Dat"
  CommonDialog1.ShowOpen
  txtImportFilename = CommonDialog1.FileTitle
  Exit Sub
  
ErrHandler:
  Exit Sub
End Sub



Private Sub CmdPrint_Click(Index As Integer)
Screen.MousePointer = vbHourglass
        
    If lstERRList.ListItems.Count <> 0 Then
        Call ExportListViewtoExcel(lstERRList)
    Else
        MsgBox "Report cannot be printed without any record.", vbInformation
    End If
        
    Screen.MousePointer = vbDefault
End Sub

Private Sub cmdUpload_Click()
 Dim filename As String
   filename = Format(Now, "DD") & "-" & Format(Now, "MM") & "-" & Format(Now, "YYYY") & "-" & Format(Time, "HH") & "-" & Format(Time, "NN") & "-" & Format(Time, "SS") & ".csv"
    Dim fso, txtfile
    Set fso = CreateObject("Scripting.FileSystemObject")
    Set txtfile = fso.CreateTextFile("\\svr-doc\Membership\Bor_in\Zurich Card printing\" & filename, True)
    Dim loopcount As Integer
    loopcount = 0
    Do While loopcount < totcount
     loopcount = loopcount + 1
     txtfile.WriteLine (vRecordArray(loopcount - 1))
    Loop
    txtfile.Close
    
 If intExit = 0 Then
            MsgBox "Uploaded Sucessfully... Text file save @ \\svr-doc\Membership\Bor_in\Zurich Card printing", vbOKOnly + vbInformation, DialogBoxTitle
            Screen.MousePointer = vbDefault
        Else
            Screen.MousePointer = vbDefault
        End If
End Sub

Private Sub cmdVerify_Click()
Dim PolicyDetails As New ADODB.Recordset
Dim strsql As String
Dim vNewName As String
Dim VNewPolicyNo As String
Dim vDate As String
Dim vPlandesc As String
Dim vplan As String
Dim vAmount As String
Dim myLines() As String

Dim vEncoding As String
Erase RecordArray
lstERRList.ListItems.Clear

Open txtImportFilename For Input As #1   ' Open file.
        Do
        Do While Not EOF(1)
        Line Input #1, TextLine
        
        myLines() = Split(TextLine, "|")
        If Trim(myLines(0)) = "" Then
           Exit Do
        Else
           totcount = totcount + 1
        End If
        Loop
            Loop Until Check = False
        Close #1   ' Close file.
   
ReDim vRecordArray(totcount - 1) As String
   Open txtImportFilename For Input As #1   ' Open file.
        Do
        Do While Not EOF(1)  ' Loop until end of file.
           Line Input #1, TextLine   ' Read line into variable.
            NoOfRecLine = NoOfRecLine + 1
                RecordArray() = Split(TextLine, "|")
                VNewPolicyNo = ""
                    vNewName = ""
                    vDate = ""
                    vPlandesc = ""
                    vplan = ""
                    vAmount = ""
                vEncoding = ""
            Dim I As Integer
                
                    If CInt(UBound(RecordArray)) = 0 Then
                        VNewPolicyNo = Trim(RecordArray(0))
                    ElseIf CInt(UBound(RecordArray)) = 1 Then
                        VNewPolicyNo = Trim(RecordArray(0))
                        vNewName = Trim(RecordArray(1))
                    ElseIf CInt(UBound(RecordArray)) = 2 Then
                    VNewPolicyNo = Trim(RecordArray(0))
                        vNewName = Trim(RecordArray(1))
                        vDate = Trim(RecordArray(2))
                    ElseIf CInt(UBound(RecordArray)) = 3 Then
                        VNewPolicyNo = Trim(RecordArray(0))
                        vNewName = Trim(RecordArray(1))
                        vDate = Trim(RecordArray(2))
                        vPlandesc = Trim(RecordArray(3))
                    ElseIf CInt(UBound(RecordArray)) = 4 Then
                    VNewPolicyNo = Trim(RecordArray(0))
                        vNewName = Trim(RecordArray(1))
                        vDate = Trim(RecordArray(2))
                        vPlandesc = Trim(RecordArray(3))
                        vplan = Trim(RecordArray(4))
                    ElseIf CInt(UBound(RecordArray)) = 5 Then
                        VNewPolicyNo = Trim(RecordArray(0))
                        vNewName = Trim(RecordArray(1))
                        vDate = Trim(RecordArray(2))
                        vPlandesc = Trim(RecordArray(3))
                        vplan = Trim(RecordArray(4))
                        vAmount = Trim(RecordArray(5))
                    End If
                
                 '   VNewPolicyNo = Trim(RecordArray(0))
                 '   vNewName = Trim(RecordArray(1))
                 '   vDate = Trim(RecordArray(2))
                 '   vPlandesc = Trim(RecordArray(3))
                 '   vplan = Trim(RecordArray(4))
                 '   vAmount = Trim(RecordArray(5))
                    
                    Dim Status As Boolean
                    Status = False
                    If Not vNewName = "" And Not VNewPolicyNo = "" Then
                      strsql = "SELECT MBMName, MBMPolicyNo, Encoding FROM View_zurichcardprinting where MBMPolicyNo = '" & VNewPolicyNo & "' and MBMName= '" & Replace(vNewName, "'", "''") & "'"
                    PolicyDetails.Open strsql, medixmasterconn, adOpenKeyset, adLockOptimistic
                If PolicyDetails.recordCount Then
                Do Until PolicyDetails.EOF
                Status = True
           With PolicyDetails
                vEncoding = IIf(Len(PolicyDetails!Encoding), PolicyDetails!Encoding, "")
           End With
           PolicyDetails.MoveNext
       Loop
       If Status = False Then
        Call ErrorListing(NoOfRecLine, VNewPolicyNo, vNewName, "Not Found")
       End If
                Else
                    vEncoding = ""
                End If
                PolicyDetails.Close
               
                vRecordArray(NoOfRecLine - 1) = VNewPolicyNo & "|" & vNewName & "|" & vDate & "|" & vPlandesc & "|" & vplan & "|" & vAmount & "|" & vEncoding
                
                Set PolicyDetails = Nothing
                      ' Call ErrorListing(NoOfRecLine, VNewPolicyNo, vNewName, "Not Found")
                    End If
                   
                DoEvents
                If intExit = 1 Then
                    Check = False
                    Exit Do
                End If
            Loop
            Loop Until Check = False
        Close #1   ' Close file.
        If intExit = 0 Then
            MsgBox "Verified Sucessfully...", vbOKOnly + vbInformation, DialogBoxTitle
            Screen.MousePointer = vbDefault
            cmdUpload.Enabled = True
        Else
            Screen.MousePointer = vbDefault
        End If

        
  
End Sub

Private Sub ErrorListing(tmpLineNo As Integer, tmpPol As String, tmpMBMName As String, message As String)
Dim ListError As ListItem
    txtTotalError = txtTotalError + 1
    Set ListError = lstERRList.ListItems.Add(, , tmpLineNo)
        ListError.SubItems(1) = Trim(tmpPol)
        ListError.SubItems(2) = Trim(tmpMBMName)
        ListError.SubItems(3) = Trim(message)
        'ListError.SubItems(4) = Trim()
End Sub




Private Sub Exit_Click()
Unload Me
End Sub
