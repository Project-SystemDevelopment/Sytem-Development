VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Begin VB.Form frmMBMClinics 
   Caption         =   "Member Panel Clinics Maintenance"
   ClientHeight    =   5940
   ClientLeft      =   60
   ClientTop       =   510
   ClientWidth     =   10410
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   5940
   ScaleWidth      =   10410
   WindowState     =   2  'Maximized
   Begin VB.Frame Frame2 
      Height          =   615
      Left            =   120
      TabIndex        =   3
      Top             =   5280
      Width           =   10215
      Begin VB.CommandButton cmdUpdate 
         Caption         =   "Update"
         Height          =   315
         Left            =   6000
         TabIndex        =   9
         Top             =   240
         Width           =   855
      End
      Begin VB.CommandButton CmdSearch 
         Caption         =   "Search"
         Height          =   315
         Left            =   7440
         TabIndex        =   8
         Top             =   180
         Width           =   855
      End
      Begin VB.CommandButton CmdClear 
         Caption         =   "Clear"
         Height          =   315
         Left            =   8280
         TabIndex        =   7
         Top             =   180
         Width           =   975
      End
      Begin VB.CommandButton CmdExit 
         Caption         =   "Exit"
         Height          =   315
         Left            =   9240
         TabIndex        =   6
         Top             =   180
         Width           =   855
      End
      Begin VB.CommandButton cmdAdd 
         Caption         =   "Add"
         Default         =   -1  'True
         Height          =   315
         Left            =   5040
         TabIndex        =   5
         Top             =   180
         Width           =   855
      End
      Begin VB.CommandButton cmdEdit 
         Caption         =   "Edit"
         Height          =   315
         Left            =   6600
         TabIndex        =   4
         Top             =   180
         Width           =   855
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
      Height          =   615
      Left            =   120
      TabIndex        =   0
      Top             =   240
      Width           =   10215
      Begin VB.TextBox txtMBMNumber 
         Height          =   285
         Left            =   1200
         MaxLength       =   4
         TabIndex        =   1
         Top             =   240
         Width           =   1695
      End
      Begin VB.Label Label1 
         Caption         =   "Member No"
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   2
         Top             =   260
         Width           =   1335
      End
   End
   Begin MSComctlLib.ListView lstPanelHosp 
      Height          =   4335
      Left            =   1080
      TabIndex        =   10
      Top             =   960
      Width           =   10215
      _ExtentX        =   18018
      _ExtentY        =   7646
      View            =   3
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   13
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Member No"
         Object.Width           =   1411
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Member Name"
         Object.Width           =   4762
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "Clinic Code"
         Object.Width           =   2117
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Text            =   "Clinic Name"
         Object.Width           =   4410
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   4
         Text            =   "Panel"
         Object.Width           =   2540
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
   Begin MSComctlLib.ListView LstPanelClinics 
      Height          =   4335
      Left            =   600
      TabIndex        =   12
      Top             =   2040
      Width           =   10215
      _ExtentX        =   18018
      _ExtentY        =   7646
      View            =   3
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   13
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Member No"
         Object.Width           =   1411
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Member Name"
         Object.Width           =   4762
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "Clinic Code"
         Object.Width           =   2117
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Text            =   "Clinic Name"
         Object.Width           =   4410
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   4
         Text            =   "Panel"
         Object.Width           =   2540
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
Attribute VB_Name = "frmMBMClinics"
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

Private Sub cboHosp_Change()
If InStr(cboHosp.Text, "null") Then
       cboHosp.Enabled = False
End If
End Sub

Private Sub cboHosp_KeyPress(KeyAscii As Integer)
Dim NewText As String
Dim ValidCount As Integer
Dim ValidValue As String
Dim i As Integer
    Clipboard.Clear
    If KeyAscii >= 32 And KeyAscii <> 127 Then
    NewText = LCase(Left(cboHosp, cboHosp.SelStart) + Chr(KeyAscii) + Mid(cboHosp, cboHosp.SelStart + cboHosp.SelLength + 1))
    
    ValidCount = 0
    ValidValue = ""
    
    For i = 0 To cboHosp.ListCount - 1

        If NewText = LCase(Left(cboHosp.List(i), Len(NewText))) Then
                ValidCount = ValidCount + 1
                ValidValue = cboHosp.List(i)
        End If
        Next
        If ValidCount <= 1 Then KeyAscii = 0
            If ValidCount = 1 Then
               cboHosp = ValidValue
               cboHosp.SelStart = 0
               cboHosp.SelLength = Len(ValidValue)
            End If
        End If
End Sub



Private Sub cboPLNPayor_change()
If InStr(cboPLNPayor.Text, "null") Then
       cboPLNPayor.Enabled = False
    End If
End Sub

Private Sub cboPLNPayor_KeyPress(KeyAscii As Integer)
Dim NewText As String
    Dim ValidCount As Integer
    Dim ValidValue As String
    Dim i As Integer
    
    If KeyAscii > 32 And KeyAscii < 127 Then
        NewText = ChkStr(Left(cboPLNPayor.Text, cboPLNPayor.SelStart) & _
                Chr(KeyAscii) + Mid(cboPLNPayor.Text, cboPLNPayor.SelStart + cboPLNPayor.SelLength + 1))
        
        ValidCount = 0
        ValidValue = ""
        
        For i = 0 To cboPLNPayor.ListCount - 1
            
            If NewText = ChkStr(Left(cboPLNPayor.List(i), Len(NewText))) Then
                    ValidCount = ValidCount + 1
                    ValidValue = cboPLNPayor.List(i)
            End If
        Next
            If ValidCount <= 1 Then KeyAscii = 0
            End If
            If ValidCount = 1 Then
              cboPLNPayor.Text = ValidValue
              cboPLNPayor.SelStart = 0
              cboPLNPayor.SelLength = Len(ValidValue)
              End If
End Sub

Private Sub chkHOS_Click()
Dim HOS As New ADODB.Recordset
Dim HOSIndex As Integer
Dim tmpHOSCode As String
Dim tmpHosName As String

    cboHosp.Clear
    Set HOSArr = New Collection
  '  medixmasterconn.Open "File Name=" & BOSPath & "Conn\mediconnect.UDL"
    If chkHOS.Value = 1 Then
        HOS.Open "select HOSCode, HOSName from HOS", medixmasterconn, adOpenForwardOnly, adLockBatchOptimistic
        Do Until HOS.EOF
            'ReDim Preserve HOSArr(HOSIndex)
            tmpHOSCode = HOS!HOSCode
            tmpHosName = HOS!HosName
            cboHosp.AddItem HOS!HosName
            HOSArr.Add tmpHOSCode, tmpHosName
            'cboHosp.ItemData(cboHosp.NewIndex) = HOSIndex
            HOS.MoveNext
        Loop
        HOS.Close
        Set HOS = Nothing
    End If
  '  medixmasterconn.Close
  '  Set medixmasterconn = Nothing
End Sub

Private Sub chkPLNPayor_Click()
Dim PAY As New ADODB.Recordset

    cboPLNPayor.Clear
    'medixmasterconn.Open "File Name=" & BOSPath & "Conn\mediconnect.UDL"
    If chkPLNPayor.Value = 1 Then
        PAY.Open "select PAYCOde, PAYCompanyName from PAY", medixmasterconn, adOpenForwardOnly, adLockBatchOptimistic
        Do Until PAY.EOF
            cboPLNPayor.AddItem PAY!PAYCode & " - " & PAY!PAYCompanyName
            PAY.MoveNext
        Loop
        PAY.Close
        Set PAY = Nothing
    End If

    'medixmasterconn.Close
    'Set medixmasterconn = Nothing
End Sub

Private Sub cmdAdd_Click()
    cmdEdit.Visible = False
    cmdUpdate.Visible = True
    lstPanelHosp.Visible = False
    lstPanelClinics.Visible = True
    EditFlag = False
    
    
      If Trim(txtMBMNumber) = "" Then
            MsgBox "Please enter membership number!", vbCritical + vbOKOnly, DialogBoxTitle
            txtMBMNumber.SetFocus
      Else
            
            Call CLNListing
      End If

End Sub

Private Sub CmdClear_Click()
    txtPlanNo.Text = ""
    lstPanelHosp.ListItems.Clear
    chkPLNPayor.Value = False
    chkHOS.Value = False
End Sub

Private Sub cmdEdit_Click()
    cmdEdit.Visible = False
    cmdUpdate.Visible = True
    txtPlanNo.Enabled = True
    cboPLNPayor.Enabled = True
    cboHosp.Enabled = True
End Sub

Private Sub CmdExit_Click()
        Unload Me
        Set frmPanelHospMaintenance = Nothing
End Sub


Public Sub CmdSearch_Click()
  
        Call PanelListing
End Sub



Private Sub lstPanelHosp_Click()
Dim lstPanel As ListItem

    cmdEdit.Visible = True
    cmdUpdate.Visible = False
    txtPlanNo.Enabled = False
    cboPLNPayor.Enabled = False
    cboHosp.Enabled = False
    
    If lstPanelHosp.ListItems.Count > 0 Then
        Set lstPanel = lstPanelHosp.SelectedItem
            cboPLNPayor = IIf(Len(Trim(lstPanel.Text)), lstPanel.Text, "")
            tmpPay = IIf(Len(Trim(lstPanel.Text)), lstPanel.Text, "")
            txtPlanNo = IIf(Len(Trim(lstPanel.SubItems(2))), lstPanel.SubItems(2), "")
            tmpPlan = IIf(Len(Trim(lstPanel.SubItems(2))), lstPanel.SubItems(2), "")
            cboHosp = IIf(Len(Trim(lstPanel.SubItems(3))), lstPanel.SubItems(3), "")
            tmpHOS = IIf(Len(Trim(lstPanel.SubItems(3))), lstPanel.SubItems(3), "")
    End If
End Sub

Private Sub lstPanelHosp_KeyDown(KeyCode As Integer, Shift As Integer)
Dim lstPanel As ListItem

    cmdEdit.Visible = True
    cmdUpdate.Visible = False
    txtPlanNo.Enabled = False
    cboPLNPayor.Enabled = False
    cboHosp.Enabled = False
    
    If lstPanelHosp.ListItems.Count > 0 Then
        Set lstPanel = lstPanelHosp.SelectedItem
            cboPLNPayor = IIf(Len(Trim(lstPanel.Text)), lstPanel.Text, "")
            tmpPay = IIf(Len(Trim(lstPanel.Text)), lstPanel.Text, "")
            txtPlanNo = IIf(Len(Trim(lstPanel.SubItems(2))), lstPanel.SubItems(2), "")
            tmpPlan = IIf(Len(Trim(lstPanel.SubItems(2))), lstPanel.SubItems(2), "")
            cboHosp = IIf(Len(Trim(lstPanel.SubItems(3))), lstPanel.SubItems(3), "")
            tmpHOS = IIf(Len(Trim(lstPanel.SubItems(3))), lstPanel.SubItems(3), "")
    End If
End Sub

Private Sub lstPanelHosp_KeyUp(KeyCode As Integer, Shift As Integer)
Dim lstPanel As ListItem

    cmdEdit.Visible = True
    cmdUpdate.Visible = False
    txtPlanNo.Enabled = False
    cboPLNPayor.Enabled = False
    cboHosp.Enabled = False
    
    If lstPanelHosp.ListItems.Count > 0 Then
        Set lstPanel = lstPanelHosp.SelectedItem
            cboPLNPayor = IIf(Len(Trim(lstPanel.Text)), lstPanel.Text, "")
            tmpPay = IIf(Len(Trim(lstPanel.Text)), lstPanel.Text, "")
            txtPlanNo = IIf(Len(Trim(lstPanel.SubItems(2))), lstPanel.SubItems(2), "")
            tmpPlan = IIf(Len(Trim(lstPanel.SubItems(2))), lstPanel.SubItems(2), "")
            cboHosp = IIf(Len(Trim(lstPanel.SubItems(3))), lstPanel.SubItems(3), "")
            tmpHOS = IIf(Len(Trim(lstPanel.SubItems(3))), lstPanel.SubItems(3), "")
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
Dim i As Integer
Dim tmpcnt As Integer
Dim strsqlCLN As String
Dim PanelCLN As New ADODB.Recordset

    Erase RecordArray
    tmpcnt = 1
    i = 0
    strsql = "Select * from View_Member_PanelClinics where MBMNumber = '" & Trim(txtMBMNumber) & "'"
    PanelHOS.Open strsql, medixmasterconn, adOpenKeyset, adLockOptimistic
    
    
        If PanelHOS.EOF = False Then
            Do While Not PanelHOS.EOF
            
                With PanelHOS
                    'MBMNumber
                    'MBMName
                    'MBMPanelCLNCodes
                    'Do While Not tmpcnt = i
                    
                    RecordArray() = Split(PanelHOS!MBMPanelCLNCodes, ",")
                    
                    For i = 0 To UBound(BordxNoArray)
                        tmpcnt = tmpcnt + 1
                        If RecordArray(tmpcnt) <> "" Then
                            strsqlCLN = "Select * from CLN where CLNCode = '" & Trim(RecordArray(tmpcnt)) & "'"
                            PanelCLN.Open strsql, medixmasterconn, adOpenKeyset, adLockOptimistic
                                                    
                            If PanelCLN.recordCount Then
                                Set HospNameList = frmPanelHospMaintenance.lstPanelHosp.ListItems.Add(, , !MBMNumber)
                                HospNameList.SubItems(1) = IIf(Len(!MBMName), !MBMName, "")
                                HospNameList.SubItems(2) = IIf(Len(!PLNCode), !PLNCode, "")
                                HospNameList.SubItems(3) = IIf(Len(!HOSCode), !HOSCode, "")
                                HospNameList.SubItems(4) = IIf(Len(!HosName), !HosName, "")
                                HospNameList.SubItems(5) = IIf(Len(!HOSTelNo), !HOSTelNo, "")
                                HospNameList.SubItems(6) = IIf(Len(!HOSFaxNo), !HOSFaxNo, "")
                            End If
                                                    
                            PanelCLN.Close
                            Set PanelCLN = Nothing
                        End If
                    Next
                End With
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


Private Sub PanelCLNListing()
    strsqlCLN = "Select * from CLN order by CLNName"
    PanelCLN.Open strsql, medixmasterconn, adOpenKeyset, adLockOptimistic
                            
    If PanelCLN.recordCount Then
        Set HospNameList = frmPanelHospMaintenance.lstPanelHosp.ListItems.Add(, , !MBMNumber)
        HospNameList.SubItems(1) = IIf(Len(!MBMName), !MBMName, "")
        HospNameList.SubItems(2) = IIf(Len(!PLNCode), !PLNCode, "")
        HospNameList.SubItems(3) = IIf(Len(!HOSCode), !HOSCode, "")
        HospNameList.SubItems(4) = IIf(Len(!HosName), !HosName, "")
        HospNameList.SubItems(5) = IIf(Len(!HOSTelNo), !HOSTelNo, "")
        HospNameList.SubItems(6) = IIf(Len(!HOSFaxNo), !HOSFaxNo, "")
    End If
                            
    PanelCLN.Close
    Set PanelCLN = Nothing

End Sub
