VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Begin VB.Form frmMemberClinics 
   Caption         =   "Member Panel Clinics Maintenance"
   ClientHeight    =   5940
   ClientLeft      =   60
   ClientTop       =   510
   ClientWidth     =   10425
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   5940
   ScaleWidth      =   10425
   WindowState     =   2  'Maximized
   Begin VB.Frame Frame1 
      Caption         =   "Selection Criteria"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   120
      TabIndex        =   8
      Top             =   240
      Width           =   10215
      Begin VB.TextBox txtMBMNumber 
         Height          =   285
         Left            =   1200
         TabIndex        =   0
         Top             =   240
         Width           =   1695
      End
      Begin VB.Label lblName 
         Height          =   255
         Left            =   3840
         TabIndex        =   12
         Top             =   240
         Width           =   5055
      End
      Begin VB.Label Label1 
         Caption         =   "Member No"
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   9
         Top             =   260
         Width           =   1335
      End
   End
   Begin VB.Frame Frame2 
      Height          =   615
      Left            =   120
      TabIndex        =   1
      Top             =   5280
      Width           =   10215
      Begin VB.CommandButton cmdEdit 
         Caption         =   "Edit"
         Height          =   315
         Left            =   6600
         TabIndex        =   7
         Top             =   180
         Width           =   855
      End
      Begin VB.CommandButton cmdAdd 
         Caption         =   "Add"
         Default         =   -1  'True
         Height          =   315
         Left            =   5760
         TabIndex        =   6
         Top             =   180
         Width           =   855
      End
      Begin VB.CommandButton CmdExit 
         Caption         =   "Exit"
         Height          =   315
         Left            =   9240
         TabIndex        =   5
         Top             =   180
         Width           =   855
      End
      Begin VB.CommandButton CmdClear 
         Caption         =   "Clear"
         Height          =   315
         Left            =   8280
         TabIndex        =   4
         Top             =   180
         Width           =   975
      End
      Begin VB.CommandButton CmdSearch 
         Caption         =   "Search"
         Height          =   315
         Left            =   7440
         TabIndex        =   3
         Top             =   180
         Width           =   855
      End
      Begin VB.CommandButton cmdUpdate 
         Caption         =   "Update"
         Height          =   315
         Left            =   6600
         TabIndex        =   2
         Top             =   180
         Width           =   855
      End
   End
   Begin MSComctlLib.ListView lstPanelHosp 
      Height          =   4335
      Left            =   120
      TabIndex        =   10
      Top             =   960
      Width           =   10215
      _ExtentX        =   18018
      _ExtentY        =   7646
      View            =   3
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      Checkboxes      =   -1  'True
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   13
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Selected"
         Object.Width           =   1764
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Member Number"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "Clinic Code"
         Object.Width           =   1411
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Text            =   "Clinic Name"
         Object.Width           =   4410
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   4
         Text            =   "Panel"
         Object.Width           =   1411
      EndProperty
      BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   5
         Text            =   "Address 1"
         Object.Width           =   2117
      EndProperty
      BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   6
         Text            =   "Address 2"
         Object.Width           =   2117
      EndProperty
      BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   7
         Text            =   "Address 3"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(9) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   8
         Text            =   "State"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(10) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   9
         Text            =   "Post Code"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(11) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   10
         Text            =   "Location"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(12) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   11
         Text            =   "Tel No."
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(13) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   12
         Text            =   "Fax No."
         Object.Width           =   2540
      EndProperty
   End
   Begin MSComctlLib.ListView lstPanelClinics 
      Height          =   4335
      Left            =   120
      TabIndex        =   13
      Top             =   960
      Width           =   10215
      _ExtentX        =   18018
      _ExtentY        =   7646
      View            =   3
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      Checkboxes      =   -1  'True
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   12
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Selected"
         Object.Width           =   1764
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Clinic Code"
         Object.Width           =   1411
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "Clinic Name"
         Object.Width           =   4410
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Text            =   "Panel"
         Object.Width           =   1411
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   4
         Text            =   "Address 1"
         Object.Width           =   2117
      EndProperty
      BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   5
         Text            =   "Address 2"
         Object.Width           =   2117
      EndProperty
      BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   6
         Text            =   "Address 3"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   7
         Text            =   "State"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(9) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   8
         Text            =   "Post Code"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(10) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   9
         Text            =   "Location"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(11) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   10
         Text            =   "Tel No."
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(12) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   11
         Text            =   "Fax No."
         Object.Width           =   2540
      EndProperty
   End
   Begin VB.Label Label29 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00404040&
      Caption         =   "Member Panel Clinics Maintenance"
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
      TabIndex        =   11
      Top             =   0
      Width           =   11805
   End
End
Attribute VB_Name = "frmMemberClinics"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim HOS As New ADODB.Recordset
Dim PanelHOS As New ADODB.Recordset
Dim PanelHOSList As ListItem
Dim HOSArr As Collection
Dim StrAppend As String
Dim tmpPlan As String
Dim tmpPay As String
Dim tmpHOS As String
Dim RecordArray() As String
Private m_blnDirection(1 To 12) As Boolean
Dim ItemChecked As Boolean
Dim EditFlag As Boolean




Private Sub cmdAdd_Click()
    cmdEdit.Visible = False
    cmdUpdate.Visible = True
    lstPanelHosp.Visible = False
    lstPanelClinics.Visible = True
    EditFlag = False
    Call PanelCLNListing
End Sub

Private Sub CmdClear_Click()
    lstPanelHosp.ListItems.Clear
    lstPanelClinics.ListItems.Clear
End Sub

Private Sub cmdEdit_Click()
    cmdEdit.Visible = False
    cmdUpdate.Visible = True
    EditFlag = True
End Sub

Private Sub CmdExit_Click()
        Unload Me
        Set frmPanelHospMaintenance = Nothing
End Sub


Public Sub CmdSearch_Click()
    lstPanelHosp.Visible = True
    lstPanelClinics.Visible = False

        Call PanelListing
End Sub

Private Sub cmdUpdate_Click()
On Error GoTo errSave
    Dim RecordSaved
    Dim PayPanel As New ADODB.Recordset
    Dim strsqlPanel As String
    Dim strsql As String
    Dim tmpHOSCode As String
    Dim i As Integer
    Dim startTrans As Boolean
    Dim strMBMClinic As String
     startTrans = False
     ItemChecked = False
     strMBMClinic = ""
     
     'Add New
      If EditFlag = False Then
            For i = 1 To lstPanelClinics.ListItems.Count
                If lstPanelClinics.ListItems(i).Checked = True Then
                    ItemChecked = True
                    Exit For
                Else
                    ItemChecked = False
                End If
            Next i
        Else
        'Update Edit
            For i = 1 To lstPanelHosp.ListItems.Count
                If lstPanelHosp.ListItems(i).Checked = False Then
                    ItemChecked = False
                    Exit For
                Else
                    ItemChecked = True
                End If
            Next i
        End If
            
               
     If Trim(txtMBMNumber) = "" Then
            MsgBox "Please enter membership number!", vbCritical + vbOKOnly, DialogBoxTitle
            txtMBMNumber.SetFocus
     Else

         RecordSaved = MsgBox("Do you want to update the record?", vbYesNo + vbExclamation)
         If RecordSaved = vbYes Then

             '   medixmasterconn.Open "File Name=" & BOSPath & "Conn\mediconnect.UDL"
                Screen.MousePointer = vbHourglass
                startTrans = True
                medixmasterconn.beginTrans
                
                If EditFlag = False Then
                        'Add New
                        For i = 1 To lstPanelClinics.ListItems.Count
                            If lstPanelClinics.ListItems(i).Checked = True Then
                                If strMBMClinic = "" Then
                                    strMBMClinic = Trim(lstPanelClinics.ListItems(i).SubItems(1))
                                Else
                                
                                    strMBMClinic = strMBMClinic & "," & Trim(lstPanelClinics.ListItems(i).SubItems(1))
                                End If
                                
                            End If
                        Next i
                
                        strsql = "Select * from View_Member_PanelClinics where MBMNumber = '" & Trim(txtMBMNumber) & "'"
                        PayPanel.Open strsql, medixmasterconn, adOpenKeyset, adLockOptimistic
                        
                        If PayPanel.recordCount <> 0 Then
                            If PayPanel!MBMPanelCLNCodes <> "" Then
                                strMBMClinic = PayPanel!MBMPanelCLNCodes & "," & strMBMClinic
                            Else
                                strMBMClinic = strMBMClinic
                            End If
                            'MBMPanelCLN
                        PayPanel.Close
                        Set PayPanel = Nothing
                                                strsql = " update MBMCLNOthers set " & _
                                 " MBMPanelCLNCodes = '" & strMBMClinic & "', " & _
                                 " MBMPanelCLN = 'Y' " & _
                                 " Where MBMNumber = '" & Trim(txtMBMNumber) & "'"
                        
                        medixmasterconn.Execute strsql
                        End If
                        
                            

                        
                Else
                        
                       For i = 1 To lstPanelHosp.ListItems.Count
                            If lstPanelHosp.ListItems(i).Checked = True Then
                                If strMBMClinic = "" Then
                                    strMBMClinic = Trim(lstPanelHosp.ListItems(i).SubItems(2))
                                Else
                                    strMBMClinic = strMBMClinic & "," & Trim(lstPanelHosp.ListItems(i).SubItems(2))
                                End If
                            End If
                        Next i
                
                
                
                        strsql = " update MBMCLNOthers set " & _
                                 " MBMPanelCLNCodes = '" & strMBMClinic & "', " & _
                                 " MBMPanelCLN = 'Y' " & _
                                 " Where MBMNumber = '" & Trim(txtMBMNumber) & "'"
                        
                        medixmasterconn.Execute strsql
                
                End If
                
                medixmasterconn.commitTrans
                MsgBox "Record updated !", vbOKOnly
                Screen.MousePointer = Default
                lstPanelHosp.ListItems.Clear
                lstPanelClinics.ListItems.Clear
                lstPanelHosp.Visible = True
                lstPanelClinics.Visible = False

                Call PanelListing
                cmdEdit.Visible = True
                cmdUpdate.Visible = False
         End If
         
    End If

Exit Sub

errSave:
    MsgBox Err.Description
  '  medixmasterconn.Close
  '  Set medixmasterconn = Nothing
    If startTrans = True Then
        medixmasterconn.RollbackTrans
    End If
    Screen.MousePointer = Default
    
End Sub

Private Sub lstPanelClinics_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
        Select Case ColumnHeader.Index
            Case 1
                SortListView lstPanelClinics, ColumnHeader.Index, ldtString, True
            Case 2
                SortListView lstPanelClinics, ColumnHeader.Index, ldtString, True
            Case 3
                SortListView lstPanelClinics, ColumnHeader.Index, ldtString, True
            Case 4
                SortListView lstPanelClinics, ColumnHeader.Index, ldtString, True
            Case 5
                SortListView lstPanelClinics, ColumnHeader.Index, ldtString, True
            Case 6
                SortListView lstPanelClinics, ColumnHeader.Index, ldtString, True
            Case 7
                SortListView lstPanelClinics, ColumnHeader.Index, ldtString, True
            Case 8
                SortListView lstPanelClinics, ColumnHeader.Index, ldtString, True
            Case 9
                SortListView lstPanelClinics, ColumnHeader.Index, ldtString, True
            Case 10
                SortListView lstPanelClinics, ColumnHeader.Index, ldtString, True
            Case 11
                SortListView lstPanelClinics, ColumnHeader.Index, ldtString, True
            Case 12
                SortListView lstPanelClinics, ColumnHeader.Index, ldtString, True
            
        End Select

End Sub

Private Sub lstPanelClinics_ItemCheck(ByVal Item As MSComctlLib.ListItem)
    ItemChecked = True
    Call ProcessCheckedItem(Item)
End Sub

Private Sub lstPanelHosp_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
        Select Case ColumnHeader.Index
            Case 1
                SortListView lstPanelHosp, ColumnHeader.Index, ldtString, m_blnDirection(1)
            Case 2
                SortListView lstPanelHosp, ColumnHeader.Index, ldtString, m_blnDirection(2)
            Case 3
                SortListView lstPanelHosp, ColumnHeader.Index, ldtString, m_blnDirection(3)
            Case 4
                SortListView lstPanelHosp, ColumnHeader.Index, ldtString, m_blnDirection(4)
            Case 5
                SortListView lstPanelHosp, ColumnHeader.Index, ldtString, m_blnDirection(5)
            Case 6
                SortListView lstPanelHosp, ColumnHeader.Index, ldtString, m_blnDirection(6)
            Case 7
                SortListView lstPanelHosp, ColumnHeader.Index, ldtString, m_blnDirection(7)
            Case 8
                SortListView lstPanelHosp, ColumnHeader.Index, ldtString, m_blnDirection(8)
            Case 9
                SortListView lstPanelHosp, ColumnHeader.Index, ldtString, m_blnDirection(9)
            Case 10
                SortListView lstPanelHosp, ColumnHeader.Index, ldtString, m_blnDirection(10)
            Case 11
                SortListView lstPanelHosp, ColumnHeader.Index, ldtString, m_blnDirection(11)
            Case 12
                SortListView lstPanelHosp, ColumnHeader.Index, ldtString, m_blnDirection(12)
            
        End Select

End Sub

Private Sub lstPanelHosp_ItemCheck(ByVal Item As MSComctlLib.ListItem)
    ItemChecked = True
    Call ProcessCheckedItem(Item)
End Sub

Public Function ProcessCheckedItem(Item As ListItem)

    If Item.Checked = True Then
        Item.Text = "Panel"
    Else
        Item.Text = ""
    End If
    
End Function


Private Sub PanelListing()
Dim strsql As String
Dim HOS As New ADODB.Recordset
Dim PanelHOS As New ADODB.Recordset
Dim PanelHOSList As ListItem
Dim cmd As New ADODB.Command
Dim Prm1 As ADODB.Parameter
Dim HospNameList As ListItem
Dim tmpHOSCode As String
Dim i As Integer
Dim tmpcnt As Integer
Dim strsqlCLN As String
Dim PanelCLN As New ADODB.Recordset
     lstPanelHosp.ListItems.Clear
    Erase RecordArray
    tmpcnt = 0
    i = 1
    lblName = ""
    strsql = "Select * from View_Member_PanelClinics where MBMNumber = '" & Trim(txtMBMNumber) & "'"
    PanelHOS.Open strsql, medixmasterconn, adOpenKeyset, adLockOptimistic
    
    
        If PanelHOS.EOF = False Then
            Do While Not PanelHOS.EOF
            
                'With PanelHOS
                    'MBMNumber
                    'MBMName
                    'MBMPanelCLNCodes
                    'Do While Not tmpcnt = i
                    If Len(PanelHOS!MBMPanelCLNCodes) Then
                        RecordArray() = Split(PanelHOS!MBMPanelCLNCodes, ",")
                    Else
                        RecordArray() = Split("", ",")
                    End If
                    'i = 1
                    'If UBound(RecordArray) <> 0 Then
                        For i = 0 To UBound(RecordArray)
                            'tmpcnt = 0
                            If RecordArray(tmpcnt) <> "" Then
                                strsqlCLN = "Select * from CLN where CLNCode = '" & Trim(RecordArray(tmpcnt)) & "'"
                                PanelCLN.Open strsqlCLN, medixmasterconn, adOpenKeyset, adLockOptimistic
                                                        
                                If PanelCLN.recordCount Then
                                    Set HospNameList = lstPanelHosp.ListItems.Add(, , "Panel")
                                    HospNameList.Checked = True
                                    lblName = IIf(Len(PanelHOS!MBMName), PanelHOS!MBMName, "")
                                    HospNameList.SubItems(1) = IIf(Len(PanelHOS!MBMNumber), PanelHOS!MBMNumber, "")
                                    HospNameList.SubItems(2) = IIf(Len(PanelCLN!CLNCode), PanelCLN!CLNCode, "")
                                    HospNameList.SubItems(3) = IIf(Len(PanelCLN!CLNname), PanelCLN!CLNname, "")
                                    HospNameList.SubItems(4) = IIf(Len(PanelCLN!CLNPanel), PanelCLN!CLNPanel, "")
                                    HospNameList.SubItems(5) = IIf(Len(PanelCLN!CLNAdd1), PanelCLN!CLNAdd1, "")
                                    HospNameList.SubItems(6) = IIf(Len(PanelCLN!CLNAdd2), PanelCLN!CLNAdd2, "")
                                    HospNameList.SubItems(7) = IIf(Len(PanelCLN!CLNAdd3), PanelCLN!CLNAdd3, "")
                                    HospNameList.SubItems(8) = IIf(Len(PanelCLN!CLNState), PanelCLN!CLNState, "")
                                    HospNameList.SubItems(9) = IIf(Len(PanelCLN!CLNPostCode), PanelCLN!CLNPostCode, "")
                                    HospNameList.SubItems(10) = IIf(Len(PanelCLN!CLNLocation), PanelCLN!CLNLocation, "")
                                    HospNameList.SubItems(11) = IIf(Len(PanelCLN!CLNPhone), PanelCLN!CLNPhone, "")
                                    HospNameList.SubItems(12) = IIf(Len(PanelCLN!CLNFax), PanelCLN!CLNFax, "")
                                End If
                                                        
                                PanelCLN.Close
                                Set PanelCLN = Nothing
                            End If
                            'i = i + 1
                            tmpcnt = tmpcnt + 1
                        Next
                'End With
                'End If
            PanelHOS.MoveNext
            Loop
         
        ElseIf PanelHOS.EOF = True Then
              MsgBox "No Record Found!", vbCritical
        End If
    
        PanelHOS.Close
        Set PanelHOS = Nothing
        
       ' medixmasterconn.Close
       ' Set medixmasterconn = Nothing
End Sub

Private Sub PanelCLNListing(Optional StrAppend As String)
Dim strsql As String
Dim HOS As New ADODB.Recordset
Dim PanelHOS As New ADODB.Recordset
Dim PanelHOSList As ListItem
Dim cmd As New ADODB.Command
Dim Prm1 As ADODB.Parameter
Dim HospNameList As ListItem
Dim tmpHOSCode As String
Dim i As Integer
Dim tmpcnt As Integer
Dim strsqlCLN As String
Dim PanelCLN As New ADODB.Recordset
     lstPanelClinics.ListItems.Clear
    Erase RecordArray
    tmpcnt = 0
    i = 1
            strsqlCLN = "Select * from CLN order by CLNName"
            PanelCLN.Open strsqlCLN, medixmasterconn, adOpenKeyset, adLockOptimistic
                                    
            Do While Not PanelCLN.EOF
            
                Set HospNameList = lstPanelClinics.ListItems.Add(, , "")
                'HospNameList.SubItems(1) = IIf(Len(!MBMNumber), !MBMNumber, "")
                'lblName = IIf(Len(PanelHOS!MBMName), PanelHOS!MBMName, "")
                
                HospNameList.SubItems(1) = IIf(Len(PanelCLN!CLNCode), PanelCLN!CLNCode, "")
                HospNameList.SubItems(2) = IIf(Len(PanelCLN!CLNname), PanelCLN!CLNname, "")
                HospNameList.SubItems(3) = IIf(Len(PanelCLN!CLNPanel), PanelCLN!CLNPanel, "")
                HospNameList.SubItems(4) = IIf(Len(PanelCLN!CLNAdd1), PanelCLN!CLNAdd1, "")
                HospNameList.SubItems(5) = IIf(Len(PanelCLN!CLNAdd2), PanelCLN!CLNAdd2, "")
                HospNameList.SubItems(6) = IIf(Len(PanelCLN!CLNAdd3), PanelCLN!CLNAdd3, "")
                HospNameList.SubItems(7) = IIf(Len(PanelCLN!CLNState), PanelCLN!CLNState, "")
                HospNameList.SubItems(8) = IIf(Len(PanelCLN!CLNPostCode), PanelCLN!CLNPostCode, "")
                HospNameList.SubItems(9) = IIf(Len(PanelCLN!CLNLocation), PanelCLN!CLNLocation, "")
                HospNameList.SubItems(10) = IIf(Len(PanelCLN!CLNPhone), PanelCLN!CLNPhone, "")
                HospNameList.SubItems(11) = IIf(Len(PanelCLN!CLNFax), PanelCLN!CLNFax, "")
                                                    
            PanelCLN.MoveNext
            Loop

       ' ElseIf PanelHOS.EOF = True Then
       '       MsgBox "No Record Found!", vbCritical
    '    End If
    
        
        PanelCLN.Close
        Set PanelCLN = Nothing
       ' medixmasterconn.Close
       ' Set medixmasterconn = Nothing
End Sub

