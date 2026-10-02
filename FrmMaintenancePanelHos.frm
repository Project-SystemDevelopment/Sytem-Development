VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Begin VB.Form FrmMaintenancePanelHos 
   Caption         =   "Panel Hospital Maintenance"
   ClientHeight    =   6870
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   10815
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   6870
   ScaleWidth      =   10815
   WindowState     =   2  'Maximized
   Begin VB.Frame Frame2 
      Height          =   615
      Left            =   0
      TabIndex        =   5
      Top             =   5760
      Width           =   10215
      Begin VB.CommandButton CmdClear 
         Caption         =   "Clear"
         Height          =   315
         Left            =   8280
         TabIndex        =   9
         Top             =   180
         Width           =   975
      End
      Begin VB.CommandButton CmdExit 
         Caption         =   "Exit"
         Height          =   315
         Left            =   9240
         TabIndex        =   8
         Top             =   180
         Width           =   855
      End
      Begin VB.CommandButton cmdAdd 
         Caption         =   "Add"
         Default         =   -1  'True
         Height          =   315
         Left            =   6600
         TabIndex        =   7
         Top             =   180
         Width           =   855
      End
      Begin VB.CommandButton cmdEdit 
         Caption         =   "Edit"
         Height          =   315
         Left            =   7440
         TabIndex        =   6
         Top             =   180
         Width           =   855
      End
      Begin VB.CommandButton cmdUpdate 
         Caption         =   "Update"
         Height          =   315
         Left            =   7440
         TabIndex        =   11
         Top             =   180
         Width           =   855
      End
      Begin VB.CommandButton CmdSearch 
         Caption         =   "Search"
         Height          =   315
         Left            =   8400
         TabIndex        =   10
         Top             =   180
         Visible         =   0   'False
         Width           =   855
      End
      Begin VB.Label lblCounter 
         BackColor       =   &H00FFFFFF&
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
         Left            =   120
         TabIndex        =   13
         Top             =   240
         Width           =   735
      End
      Begin VB.Label Label3 
         Caption         =   "Record(s)"
         Height          =   255
         Left            =   960
         TabIndex        =   12
         Top             =   240
         Width           =   1455
      End
   End
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
      Height          =   735
      Left            =   0
      TabIndex        =   1
      Top             =   240
      Width           =   10215
      Begin VB.ComboBox cboPLNPayor 
         Height          =   315
         Left            =   1680
         TabIndex        =   2
         Top             =   240
         Width           =   3255
      End
      Begin VB.Label Label1 
         Caption         =   "Hospital Panel Code"
         Height          =   255
         Index           =   1
         Left            =   120
         TabIndex        =   3
         Top             =   255
         Width           =   1575
      End
   End
   Begin MSComctlLib.ListView lstPanelHosp 
      Height          =   4695
      Left            =   0
      TabIndex        =   4
      Top             =   1080
      Width           =   10215
      _ExtentX        =   18018
      _ExtentY        =   8281
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
      NumItems        =   10
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Object.Width           =   2117
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Hosp Code"
         Object.Width           =   2117
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "Hosp Name"
         Object.Width           =   6174
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Text            =   "Tel No"
         Object.Width           =   2117
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   4
         Text            =   "Fax No"
         Object.Width           =   2117
      EndProperty
      BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   5
         Text            =   "Add1"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   6
         Text            =   "Add2"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   7
         Text            =   "Add3"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(9) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   8
         Text            =   "State"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(10) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   9
         Text            =   "Town"
         Object.Width           =   2540
      EndProperty
   End
   Begin MSComctlLib.ListView lstHospitalListing 
      Height          =   4695
      Left            =   0
      TabIndex        =   14
      Top             =   1080
      Width           =   10215
      _ExtentX        =   18018
      _ExtentY        =   8281
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
      NumItems        =   3
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Object.Width           =   2117
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Hospital Code"
         Object.Width           =   2646
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "Hospital Name"
         Object.Width           =   2646
      EndProperty
   End
   Begin VB.Label Label29 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00404040&
      Caption         =   "Panel Hospital Maintenance"
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
      Width           =   10485
   End
End
Attribute VB_Name = "FrmMaintenancePanelHos"
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
Dim StrAppend2 As String
Dim tmpPlan As String
Dim tmpPay As String
Dim EditFlag As Boolean
Dim ItemChecked As Boolean
Dim tmpPNLHOSStatus As String
Dim tmpCount As Double
Dim tmpPayor As String

Private Sub cboPLNPayor_KeyPress(KeyAscii As Integer)
Dim NewText As String
Dim ValidCount As Integer
Dim ValidValue As String
Dim i As Integer
    If KeyAscii >= 32 And KeyAscii <> 127 Then
    NewText = LCase(Left(cboPLNPayor.Text, cboPLNPayor.SelStart) + Chr(KeyAscii) + Mid(cboPLNPayor.Text, cboPLNPayor.SelStart + cboPLNPayor.SelLength + 1))
 
    ValidCount = 0
    ValidValue = ""
 
    For i = 0 To cboPLNPayor.ListCount - 1
 
        If NewText = LCase(Left(cboPLNPayor.List(i), Len(NewText))) Then
                ValidCount = ValidCount + 1
                ValidValue = cboPLNPayor.List(i)
        End If
        Next
        If ValidCount <= 1 Then KeyAscii = 0
            If ValidCount = 1 Then
                cboPLNPayor.Text = ValidValue
                cboPLNPayor.SelStart = 0
                cboPLNPayor.SelLength = Len(ValidValue)
            End If
        End If
End Sub



Private Sub Form_Load()
    EditFlag = True
    GetPanelHosCode
End Sub

Private Sub GetPanelHosCode()
Dim HosPanel As New ADODB.Recordset
Dim sqlINS As String

    cboPLNPayor.Clear
    lstPanelHosp.ListItems.Clear
  ' medixmasterconn.Open "File Name=" & BOSPath & "Conn\MediConnHISMaint.UDL"
    
     '   HosPanel.Open "SELECT HosPanelCode, HosPanelDesc From HosPanelCode", medixmasterconn, adOpenForwardOnly, adLockBatchOptimistic
        
        
    sqlINS = "SELECT HosPanelCode, HosPanelDesc From HosPanelCode"
    
 '   If Len(Trim(StrAppend)) Then
 '       sqlINS = sqlINS + StrAppend
 '   End If
    
    HosPanel.Open sqlINS, medixmasterconnMaint, adOpenForwardOnly, adLockBatchOptimistic
        
        Do Until HosPanel.EOF
            'cboPLNPayor.AddItem HosPanel!HosPanelCode & " - " & HosPanel!HosPanelDesc
            cboPLNPayor.AddItem HosPanel!HosPanelCode
            HosPanel.MoveNext
        Loop
        HosPanel.Close
        Set HosPanel = Nothing

   ' medixmasterconn.Close
   ' Set medixmasterconn = Nothing
    
End Sub

Private Sub cboPLNPayor_Click()
    'tmpPayor = Trim(Mid(cboPLNPayor, 1, InStr(1, cboPLNPayor, "-") - 1))
    tmpPayor = Trim(cboPLNPayor)
        If Trim(cboPLNPayor) = "" Then
            MsgBox "Please Select Hospital Panel Code!", vbCritical + vbOKOnly, DialogBoxTitle
            cboPLNPayor.SetFocus
        Else
            StrAppend = ""
            cmdEdit.Visible = True
            cmdUpdate.Visible = True
                      
            If cboPLNPayor <> "" Then
                StrAppend = StrAppend + " and PanelHospitals.HosPanelCode = '" & tmpPayor & "' "
               
            End If
            Call PanelListing(StrAppend)
        End If
End Sub

Private Sub CmdExit_Click()
    Unload Me
    Set frmPanelHospMaintenance = Nothing
End Sub

Private Sub CmdClear_Click()
    lstPanelHosp.ListItems.Clear
    lstHospitalListing.ListItems.Clear
End Sub

Public Sub CmdSearch_Click()

        If Trim(cboPLNPayor) = "" Then
            MsgBox "Please Select Hospital Panel Code!", vbCritical + vbOKOnly, DialogBoxTitle
            cboPLNPayor.SetFocus
        Else
            StrAppend = ""
            cmdEdit.Visible = True
            cmdUpdate.Visible = True
                      
            If cboPLNPayor <> "" Then
                StrAppend = StrAppend + " and PanelHospitals.HosPanelCode = '" & tmpPayor & "' "
               
            End If
    
            Call PanelListing(StrAppend)
        End If

End Sub

Private Sub PanelListing(Optional StrAppend As String)
Dim strsql As String
Dim HOS As New ADODB.Recordset
Dim PanelHOS As New ADODB.Recordset
Dim PanelHOSList As ListItem
Dim cmd As New ADODB.Command
Dim Prm1 As ADODB.Parameter
Dim HospNameList As ListItem
Dim tmpHOSCode As String

      lblCounter = 0
      lstPanelHosp.Visible = True
      lstHospitalListing.Visible = False
      Screen.MousePointer = vbHourglass
    '  medixmasterconn.Open "File Name=" & BOSPath & "Conn\mediconnect.UDL"
      
      Set cmd.ActiveConnection = medixmasterconn
      cmd.CommandText = "Search_PanelHospitals2"
      cmd.CommandType = adCmdStoredProc
      cmd.CommandTimeout = 300
      Set Prm1 = cmd.CreateParameter("StrAppend", adVarChar, adParamInput, 2000, StrAppend)
      cmd.Parameters.Append Prm1
      HOS.CursorLocation = adUseClient
      HOS.CursorType = adOpenForwardOnly
      HOS.CacheSize = 100
      
      Set PanelHOS = cmd.Execute
      FrmMaintenancePanelHos.lstPanelHosp.ListItems.Clear
    
        If PanelHOS.EOF = False Then
            Do While Not PanelHOS.EOF
                With PanelHOS
                    Set HospNameList = FrmMaintenancePanelHos.lstPanelHosp.ListItems.Add(, , "Panel")
                    HospNameList.Checked = True
                    HospNameList.SubItems(1) = IIf(Len(!HOSCode), !HOSCode, "")
                    HospNameList.SubItems(2) = IIf(Len(!HosName), !HosName, "")
                    HospNameList.SubItems(3) = IIf(Len(!HOSTelNo), !HOSTelNo, "")
                    HospNameList.SubItems(4) = IIf(Len(!HOSFaxNo), !HOSFaxNo, "")
                    HospNameList.SubItems(5) = IIf(Len(!HOSAdd1), !HOSAdd1, "")
                    HospNameList.SubItems(6) = IIf(Len(!HOSAdd2), !HOSAdd2, "")
                    HospNameList.SubItems(7) = IIf(Len(!HOSAdd3), !HOSAdd3, "")
                    HospNameList.SubItems(8) = IIf(Len(!HOSState), !HOSState, "")
                    HospNameList.SubItems(9) = IIf(Len(!HOSTown), !HOSTown, "")
                    
                End With
                PanelHOS.MoveNext
                lblCounter = lblCounter + 1
            Loop
         
        ElseIf PanelHOS.EOF = True Then
              MsgBox "No Record Found!", vbCritical
        End If
    
        PanelHOS.Close
        Set PanelHOS = Nothing
        
     '   medixmasterconn.Close
       ' Set medixmasterconn = Nothing
        Screen.MousePointer = Default
        
End Sub

Private Sub cmdEdit_Click()
    cmdEdit.Visible = False
    cmdUpdate.Visible = True
    cboPLNPayor.Enabled = True
    lstPanelHosp.Visible = True
    lstHospitalListing.Visible = False
    EditFlag = True
    
End Sub

Private Sub cmdAdd_Click()
    cmdEdit.Visible = False
    cmdUpdate.Visible = True
    cboPLNPayor.Enabled = True
    lstPanelHosp.Visible = False
    lstHospitalListing.Visible = True
    EditFlag = False
    
    
      If Trim(cboPLNPayor) = "" Then
            MsgBox "Please Select Hospital Panel Code!", vbCritical + vbOKOnly, DialogBoxTitle
            cboPLNPayor.SetFocus
      Else
            
            Call HOSListing
      End If

End Sub

Private Sub HOSListing()
Dim HOS As New ADODB.Recordset
Dim cmd As New ADODB.Command
Dim Prm1 As ADODB.Parameter
Dim HospList As ListItem

    lstHospitalListing.ListItems.Clear
    lblCounter = 0
   ' medixmasterconn.Open "File Name=" & BOSPath & "Conn\mediconnect.UDL"
    Set cmd.ActiveConnection = medixmasterconn
    cmd.CommandText = "SearchHOS2"
    cmd.CommandType = adCmdStoredProc
    cmd.CommandTimeout = 300
        
    Set Prm1 = cmd.CreateParameter("tmpPayor", adVarChar, adParamInput, 100, tmpPayor)
    cmd.Parameters.Append Prm1
   
    HOS.CursorLocation = adUseClient
    HOS.CursorType = adOpenForwardOnly
    HOS.CacheSize = 100
    Set HOS = cmd.Execute
    
    Do Until HOS.EOF
        With HOS
            Set HospList = FrmMaintenancePanelHos.lstHospitalListing.ListItems.Add(, , "")
            HospList.SubItems(1) = IIf(Len(!HosName), !HOSCode, "")
            HospList.SubItems(2) = IIf(Len(!HosName), !HosName, "")
        End With
        HOS.MoveNext
        lblCounter = lblCounter + 1
    Loop
    HOS.Close
    Set HOS = Nothing
    'medixmasterconn.Close
   ' Set medixmasterconn = Nothing
 '   Screen.MousePointer = Default
    
End Sub

Private Sub lstHospitalListing_ItemCheck(ByVal Item As MSComctlLib.ListItem)
    ItemChecked = True
    Call ProcessCheckedItem(Item)
End Sub

Private Sub lstPanelHosp_ItemCheck(ByVal Item As MSComctlLib.ListItem)
    ItemChecked = True
    Call ProcessCheckedItem(Item)
End Sub

Public Function ProcessCheckedItem(Item As ListItem)
Dim i, newbound  As Integer
Dim listloop As Integer
Dim strsqlCLM As String
Dim CLM As New ADODB.Recordset
Dim strsqlPAY As String
Dim PAY As New ADODB.Recordset
Dim CaseNo As String
Dim caseversion As String
Dim Proc As Boolean
Dim MsgAsk

    If Item.Checked = True Then
        Item.Text = "Panel"
    Else
        Item.Text = ""
    End If
    
End Function

Private Sub cmdUpdate_Click()
On Error GoTo errSave
    Dim RecordSaved
    Dim PayPanel As New ADODB.Recordset
    Dim strsqlPanel As String
    Dim strsql As String
    Dim tmpHOSCode As String
    Dim i As Integer
    Dim startTrans As Boolean
    
     startTrans = False
     ItemChecked = False
     
     'Add New
      If EditFlag = False Then
            For i = 1 To lstHospitalListing.ListItems.Count
                If lstHospitalListing.ListItems(i).Checked = True Then
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
            
               
     If Trim(cboPLNPayor) = "" Then
            MsgBox "Please Select Hospital Panel Code!", vbCritical + vbOKOnly, DialogBoxTitle
            cboPLNPayor.SetFocus
     Else

         RecordSaved = MsgBox("Do you want to update the record?", vbYesNo + vbExclamation)
         If RecordSaved = vbYes Then

             '   medixmasterconn.Open "File Name=" & BOSPath & "Conn\mediconnect.UDL"
                Screen.MousePointer = vbHourglass
                startTrans = True
                medixmasterconn.beginTrans
                
                If EditFlag = False Then
                        'Add New
                        For i = 1 To lstHospitalListing.ListItems.Count
                            If lstHospitalListing.ListItems(i).Checked = True Then

                                strsql = "Select * From PanelHospitals Where HosPanelCode = '" & tmpPayor & "' And HOSCode = '" & Trim(lstHospitalListing.ListItems(i).SubItems(1)) & "'"
                                PayPanel.Open strsql, medixmasterconn, adOpenKeyset, adLockOptimistic
                                
                                If PayPanel.recordCount = 0 Then
                                    PayPanel.AddNew
                                End If
                            
                                With PayPanel
                                    !HosPanelCode = Trim(tmpPayor)
                                    !HOSCode = Trim(lstHospitalListing.ListItems(i).SubItems(1))
                                    !LastUpdateUSR = Logon
                                    !LastUpdateDate = Date
                                End With
                                PayPanel.Update
                                PayPanel.Close
                                Set PayPanel = Nothing
                                
                            End If
                        Next i
                        
                Else
                
                       For i = 1 To lstPanelHosp.ListItems.Count
                            If lstPanelHosp.ListItems(i).Checked = False Then
                                strsql = "Delete From PanelHospitals Where HosPanelCode = '" & Trim(tmpPayor) & "' And HOSCode = '" & Trim(lstPanelHosp.ListItems(i).SubItems(1)) & "'"
                                PayPanel.Open strsql, medixmasterconn, adOpenKeyset, adLockOptimistic
                            End If
                        Next i
                
                End If
                
                medixmasterconn.commitTrans
                MsgBox "Record updated !", vbOKOnly
             '   medixmasterconn.Close
           '     Set medixmasterconn = Nothing
                Screen.MousePointer = Default
                
                lstPanelHosp.ListItems.Clear
                strsqlPanel = "  and PanelHospitals.HosPanelCode = '" & tmpPayor & "' "
                Call PanelListing(strsqlPanel)
                cmdEdit.Visible = True
                cmdUpdate.Visible = False
                cboPLNPayor.Enabled = True
            
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


