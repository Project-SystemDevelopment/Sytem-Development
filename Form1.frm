VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   4650
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   8910
   LinkTopic       =   "Form1"
   ScaleHeight     =   4650
   ScaleWidth      =   8910
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdExit 
      Caption         =   "&Exit"
      Height          =   375
      Left            =   3840
      TabIndex        =   4
      Top             =   2640
      Width           =   975
   End
   Begin MSComctlLib.ListView ListView1 
      Height          =   4335
      Left            =   120
      TabIndex        =   2
      Top             =   120
      Width           =   3615
      _ExtentX        =   6376
      _ExtentY        =   7646
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
      NumItems        =   4
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "ICD Code"
         Object.Width           =   1764
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "ICD Description"
         Object.Width           =   97009
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Object.Width           =   0
      EndProperty
   End
   Begin VB.CommandButton cmdRemove 
      Caption         =   "&Remove"
      Height          =   375
      Left            =   3840
      TabIndex        =   1
      Top             =   2160
      Width           =   975
   End
   Begin VB.CommandButton cmdAdd 
      Caption         =   "&Add >"
      Height          =   375
      Left            =   3840
      TabIndex        =   0
      Top             =   1680
      Width           =   975
   End
   Begin MSComctlLib.ListView ListView2 
      Height          =   4335
      Left            =   5040
      TabIndex        =   3
      Top             =   120
      Width           =   3615
      _ExtentX        =   6376
      _ExtentY        =   7646
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      AllowReorder    =   -1  'True
      FullRowSelect   =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   4
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "ICD Code"
         Object.Width           =   1764
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "ICD Description"
         Object.Width           =   97009
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Object.Width           =   0
      EndProperty
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Addrecords()

    Dim ListOriginal As New ADODB.Recordset
    Dim ListDestination As New ADODB.Recordset
    Dim sql1 As String
    Dim sql2 As String
    Dim ICD As ListItem
    
          sql1 = " select ICDCode, ICDDescription, " & _
                 " ICDLastUpdateUser, ICDLastUpdateDate" & _
                 " from ICDOriginal where ICDCode = '" & Trim(ListView1.SelectedItem.Text) & "'"
         
          sql2 = "select * from ICDTest"
          ListOriginal.Open sql1, medixmasterconn, adOpenKeyset, adLockOptimistic
          ListDestination.Open sql2, medixmasterconn, adOpenKeyset, adLockOptimistic
       
          ListDestination.AddNew
          With ListDestination
              !ICDCode = ListOriginal!ICDCode
              !ICDDescription = ListOriginal!ICDDescription
              !ICDLastUpdateUser = Logon
              !ICDLastUpdateDate = Date
          End With
          ListDestination.Update
          Set ICD = ListView2.ListItems.Add(, , ListOriginal!ICDCode)
            ICD.SubItems(1) = Trim(ListOriginal!ICDDescription)
            ICD.SubItems(2) = Logon
            ICD.SubItems(3) = Date
 
          ListDestination.Close
          ListOriginal.Close
          Set ListOriginal = Nothing
          Set ListDestination = Nothing
End Sub
Private Sub cmdAdd_Click()

    Dim CheckRecord
    Dim ICDTest As New ADODB.Recordset
    Dim strsql As String
    Dim RecordConfirm
 
    If Len(Trim(ListView1.SelectedItem.Text)) Then
        Screen.MousePointer = vbHourglass
   
            strsql = "Select ICDCode From ICDTest Where ICDCODE ='" & Trim(ListView1.SelectedItem.Text) & "'"
            ICDTest.Open strsql, medixmasterconn, adOpenKeyset, adLockOptimistic
            
           If ICDTest.RecordCount >= 1 Then
                msgstr = "ICD Code Already Exist in This Table..!" & Chr(13) & "ICD Code: " & ICDTest!ICDCode
                RecordConfirm = MsgBox(msgstr, vbOKOnly + vbCritical, DialogBoxTitle)
                
                Else
                
                Call Addrecords
           
           End If
            ICDTest.Close
            Set ICDTest = Nothing
       
        Screen.MousePointer = vbDefault
    End If
End Sub

Private Sub cmdExit_Click()
End
End Sub
Private Sub cmdremove_click()
  
    
    Dim ListDestination As New ADODB.Recordset
    Dim sql1 As String
    Dim ForLoop As Integer
    Dim ListCount As Integer
    Dim RecordConfirm
    
          
          sql1 = "Select ICDCode, ICDDescription, " & _
                " ICDLastUpdateUser, ICDLastUpdateDate" & _
                " from ICDTest where ICDCode = '" & Trim(ListView2.SelectedItem.Text) & "'"
                  
          ListDestination.Open sql1, medixmasterconn, adOpenKeyset, adLockOptimistic
    
                msgstr = "Are you sure you want to delete this record..!" & Chr(13) & "ICD Code : " & ListDestination!ICDCode
                RecordConfirm = MsgBox(msgstr, vbOKCancel + vbCritical, DialogBoxTitle)
                If RecordConfirm = vbOK Then

                    ListCount = ListView2.ListItems.Count
                    For ForLoop = 1 To ListCount
                        If ListView2.ListItems(ForLoop).Selected = True Then
                           ListView2.ListItems.Remove (ForLoop)
                    Exit For
                        End If
                    Next ForLoop

         
                    ListDestination.Delete
                    ListDestination.UpdateBatch
          
                End If
          ListDestination.Close
          Set ListDestination = Nothing

End Sub

Private Sub Form_Load()
Call ListOriginal
Call ListDestination

End Sub
Private Sub ListOriginal()
    Dim ListOriginal As New ADODB.Recordset
    Dim ICDOriginal As ListItem
    Dim sql As String

        sql = " select ICDCode, ICDDescription, " & _
          " ICDLastUpdateUser, ICDLastUpdateDate" & _
          " from ICDOriginal "
        
        ListOriginal.Open sql, medixmasterconn, adOpenKeyset, adLockOptimistic
        
        Do Until ListOriginal.EOF
            Set ICDOriginal = ListView1.ListItems.Add(, , ListOriginal!ICDCode)
                ICDOriginal.SubItems(1) = Trim(ListOriginal!ICDDescription)
                'ICDOriginal.SubItems(2) = Trim(ListOriginal!ICDLastUpdateUser)
                'ICDOriginal.SubItems(3) = Trim(ListOriginal!ICDLastUpdateDate)
                ListOriginal.MoveNext
        Loop
                     
        ListOriginal.Close
        Set ListOriginal = Nothing
      
End Sub

Private Sub ListDestination(Optional header As String)
    Dim sql1 As String
    Dim ListDestination As New ADODB.Recordset
    Dim ICDTest As ListItem
    
    sql1 = " select ICDCode, ICDDescription, " & _
      " ICDLastUpdateUser, ICDLastUpdateDate" & _
      " from ICDTest "

    ListDestination.Open sql1, medixmasterconn, adOpenKeyset, adLockOptimistic
  
        If Len(Trim(header)) Then
            If ChkStr(header) = "ICD CODE" Then
            sql1 = sql1 + "order by ICDCODe"
            End If
        End If
 
        Do Until ListDestination.EOF
            Set ICDTest = ListView2.ListItems.Add(, , ListDestination!ICDCode)
                ICDTest.SubItems(1) = Trim(ListDestination!ICDDescription)
                'ICDTest.SubItems(2) = Trim(ListDestination!ICDLastUpdateUser)
                'ICDTest.SubItems(3) = Trim(ListDestination!ICDLastUpdateDate)
                ListDestination.MoveNext
        Loop
        
      ListDestination.Close
      Set ListDestination = Nothing
        
End Sub

Private Sub ListView2_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
        ListView2.ListItems.Clear
        Call ListDestination(ColumnHeader.Text)
End Sub
