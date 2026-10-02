VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmSearchCASTracking 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Search by Case"
   ClientHeight    =   6165
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8355
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6165
   ScaleWidth      =   8355
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame11 
      Height          =   1215
      Left            =   120
      TabIndex        =   7
      Top             =   0
      Width           =   8115
      Begin VB.OptionButton optClose 
         Caption         =   "Close By"
         Height          =   240
         Left            =   6600
         TabIndex        =   16
         Top             =   720
         Width           =   1500
      End
      Begin VB.OptionButton optOpen 
         Caption         =   "Open By"
         Height          =   240
         Left            =   5160
         TabIndex        =   15
         Top             =   720
         Width           =   1500
      End
      Begin VB.OptionButton optDept 
         Caption         =   "Department"
         Height          =   240
         Left            =   3600
         TabIndex        =   14
         Top             =   720
         Width           =   1500
      End
      Begin VB.TextBox txtSearch 
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
         Left            =   795
         MaxLength       =   15
         TabIndex        =   11
         Top             =   300
         Width           =   4185
      End
      Begin VB.OptionButton optMBMNo 
         Caption         =   "Membership No"
         Height          =   240
         Left            =   1680
         TabIndex        =   10
         Top             =   720
         Width           =   1500
      End
      Begin VB.OptionButton optCaseNo 
         Caption         =   "Case No"
         Height          =   240
         Left            =   240
         TabIndex        =   9
         Top             =   720
         Value           =   -1  'True
         Width           =   1365
      End
      Begin VB.CommandButton cmdSearch 
         Caption         =   "&Search"
         Default         =   -1  'True
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
         Left            =   5520
         TabIndex        =   8
         Top             =   285
         Width           =   960
      End
      Begin VB.Label Label16 
         Caption         =   "Search "
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
         TabIndex        =   12
         Top             =   285
         Width           =   810
      End
   End
   Begin VB.Frame Frame1 
      Height          =   615
      Left            =   120
      TabIndex        =   0
      Top             =   5400
      Width           =   8175
      Begin VB.CommandButton cmdExit 
         Cancel          =   -1  'True
         Caption         =   "E&xit"
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
         Left            =   6840
         TabIndex        =   2
         Top             =   240
         Width           =   960
      End
      Begin VB.CommandButton cmdStop 
         Caption         =   "S&top"
         Height          =   315
         Left            =   5880
         TabIndex        =   1
         Top             =   240
         Width           =   960
      End
      Begin VB.Label Label1 
         Caption         =   "of"
         Height          =   255
         Left            =   960
         TabIndex        =   6
         Top             =   240
         Width           =   375
      End
      Begin VB.Label lblTotal 
         Alignment       =   2  'Center
         BackColor       =   &H80000009&
         Height          =   255
         Left            =   1200
         TabIndex        =   5
         Top             =   240
         Width           =   735
      End
      Begin VB.Label lblCurrent 
         Alignment       =   2  'Center
         BackColor       =   &H80000009&
         Height          =   255
         Left            =   120
         TabIndex        =   4
         Top             =   240
         Width           =   735
      End
      Begin VB.Label Label2 
         Caption         =   "Records"
         Height          =   255
         Left            =   2040
         TabIndex        =   3
         Top             =   240
         Width           =   1095
      End
   End
   Begin MSComctlLib.ListView lstCaseTrackingList 
      Height          =   3960
      Left            =   120
      TabIndex        =   13
      Top             =   1320
      Width           =   8055
      _ExtentX        =   14208
      _ExtentY        =   6985
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
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      NumItems        =   10
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Case No"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Member No"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "Member Name"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Text            =   "Member IC No"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   4
         Text            =   "Case Open By"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   5
         Text            =   "Case Close By"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   6
         Text            =   "Case Discharge Date"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   7
         Text            =   "Case Reg Date"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(9) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   8
         Text            =   "Department"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(10) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   9
         Text            =   "Case Last Update Date"
         Object.Width           =   2540
      EndProperty
   End
End
Attribute VB_Name = "frmSearchCASTracking"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public intExit As Integer
Public Source As Integer
Public intCaseTracking As Integer

Private Sub Form_Load()
    intExit = 0
    intCaseTracking = 1
    Call CenterForm(Me)
End Sub

' #################################

Private Sub CaseTrackingList(Optional str As String)
    counter = 0
    Dim CAS As New ADODB.Recordset
    Dim sqlstr As String
    Dim CASList As ListItem
    Dim TotalCASCount As Integer
        
        Screen.MousePointer = vbHourglass
        sqlstr = "Select CASNumber, MBMNumber, CaseOpenBy, CaseCloseBy, " & _
             "CASDischargeDate, CaseRegDate, CASLastUpdateDate, " & _
             "DPTCode, USRCode, MBMName, MBMIcBcPp "
        sqlstr = sqlstr & " From VIEW_CAS_Department Where 0=0 "
                 
        If Len(Trim(txtSearch)) Then
            If optMBMNo.Value = True Then
                    sqlstr = sqlstr & " AND  MBMNumber LIKE  '" & txtSearch & "%' "
            ElseIf optCaseNo.Value = True Then
                sqlstr = sqlstr & " AND CASNumber LIKE '%" & txtSearch & "%' "
            ElseIf optDept.Value = True Then
                sqlstr = sqlstr & " AND DPTCode LIKE '%" & txtSearch & "%' "
            ElseIf optOpen.Value = True Then
                sqlstr = sqlstr & " AND CaseOpenBy LIKE '%" & txtSearch & "%' "
            ElseIf optClose.Value = True Then
                sqlstr = sqlstr & " AND CaseCloseBy LIKE '%" & txtSearch & "%' "
            End If
        End If
          
        lblCurrent = 0
        CAS.Open sqlstr, medixmasterconn, adOpenKeyset, adLockOptimistic
        If CAS.recordCount = 0 Then
           txtTotalRecord = 0
           lblTotal = 0
           lstCaseTrackingList.listitems.Clear
           MsgBox "No record found!", vbCritical + vbOKOnly
           Screen.MousePointer = vbDefault
        Else
           lstCaseTrackingList.listitems.Clear
            lblTotal = CAS.recordCount
            Do
                
                Do Until CAS.EOF
                    Set CASList = lstCaseTrackingList.listitems.Add(, , CAS!CASNumber)
               
                        If Len(Trim(CAS!MBMNumber)) Then
                            CASList.SubItems(1) = Trim(CAS!MBMNumber)
                        End If
                        
                        If Len(Trim(CAS!MBMName)) Then
                            CASList.SubItems(2) = Trim(CAS!MBMName)
                        End If
                        
                        If Len(Trim(CAS!MBMIcBcPp)) Then
                            CASList.SubItems(3) = Trim(CAS!MBMIcBcPp)
                        End If
                        
                        If Len(Trim(CAS!CaseOpenBy)) Then
                            CASList.SubItems(4) = Trim(CAS!CaseOpenBy)
                        End If
                        
                        If Len(Trim(CAS!CaseCloseBy)) Then
                            CASList.SubItems(5) = Trim(CAS!CaseCloseBy)
                        End If
                        
                        If Len(Trim(CAS!CASDischargeDate)) Then
                            CASList.SubItems(6) = Trim(CAS!CASDischargeDate)
                        End If
                        
                        If Len(Trim(CAS!CaseRegDate)) Then
                            CASList.SubItems(7) = Trim(CAS!CaseRegDate)
                        End If
                        
                        If Len(Trim(CAS!DPTCode)) Then
                            CASList.SubItems(8) = Trim(CAS!DPTCode)
                        End If
                        
                        If Len(Trim(CAS!CASLastUpdateDate)) Then
                            CASList.SubItems(9) = Trim(CAS!CASLastUpdateDate)
                        End If
                                 
                    CAS.MoveNext
               
                    TotalCASCount = TotalCASCount + 1
                    txtTotalRecord = TotalCASCount
                    lblCurrent = lblCurrent + 1
                    DoEvents
                    
                    If intExit = 1 Then
                        check = False
                        Exit Do
                    End If
                
                Loop
            Loop Until check = False
    
    Screen.MousePointer = vbDefault
    End If
    intExit = 0
    CAS.Close
    Set CAS = Nothing
    
End Sub


Private Sub lstCaseTrackingList_DblClick()
    If lstCaseTrackingList.listitems.Count >= 1 Then
      frmCASTracking.txtCaseNo = Trim(lstCaseTrackingList.SelectedItem.Text)
      'frmCASTracking.lstCaseTrackingList.SelectedItem.Text = Trim(lstCaseTrackingList.SelectedItem.Text)
      'frmCASTracking.TxtMBMName = Trim(lstCaseTrackingList.SelectedItem.SubItems(2))
      'frmCASTracking.TxtMBMIC = Trim(lstCaseTrackingList.SelectedItem.SubItems(3))
      frmCASTracking.CboOpen = Trim(lstCaseTrackingList.SelectedItem.SubItems(4))
      'frmCASTracking.TxtOpen2 = Trim(lstCaseTrackingList.SelectedItem.SubItems(4))
      frmCASTracking.CboClose = Trim(lstCaseTrackingList.SelectedItem.SubItems(5))
      'frmCASTracking.TxtClose2 = Trim(lstCaseTrackingList.SelectedItem.SubItems(5))
      'frmCASTracking.txtDischargeDate = Trim(lstCaseTrackingList.SelectedItem.SubItems(6))
      'frmCASTracking.txtCasRegDate = Trim(lstCaseTrackingList.SelectedItem.SubItems(7))
      frmCASTracking.CboDept = Trim(lstCaseTrackingList.SelectedItem.SubItems(8))
            
      'If Len(Trim(lstCaseTrackingList.SelectedItem.SubItems(9))) Then
        'frmCASTracking.txtCasLastUpdate = Trim(lstCaseTrackingList.SelectedItem.SubItems(9))
      'End If
              
      Unload Me
      End If
End Sub

' ########################################################
Private Sub cmdSearch_Click()
    cmdStop.Enabled = True
    Call CaseTrackingList
End Sub

Private Sub cmdListAll_Click()
    Call CaseTrackingList("All")
End Sub

Private Sub CmdExit_Click()
    Unload Me
End Sub

Private Sub cmdStop_Click()
    intExit = 1
    cmdExit.Enabled = True
End Sub
' #####################################################


