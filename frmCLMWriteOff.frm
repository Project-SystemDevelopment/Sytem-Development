VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Begin VB.Form frmCLMWriteOff 
   Caption         =   "Claim Write Off"
   ClientHeight    =   8595
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   11760
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   8595
   ScaleWidth      =   11760
   Begin MSComctlLib.ListView lstWorksheetTrackList 
      Height          =   5775
      Left            =   60
      TabIndex        =   30
      Top             =   1440
      Width           =   11655
      _ExtentX        =   20558
      _ExtentY        =   10186
      View            =   3
      LabelEdit       =   1
      Sorted          =   -1  'True
      MultiSelect     =   -1  'True
      LabelWrap       =   0   'False
      HideSelection   =   0   'False
      AllowReorder    =   -1  'True
      Checkboxes      =   -1  'True
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   16
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Check"
         Object.Width           =   1411
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Status"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "Wks No"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   3
         Text            =   "Wks Status"
         Object.Width           =   1764
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   4
         Text            =   "Approved Amt"
         Object.Width           =   2117
      EndProperty
      BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   5
         Text            =   "Payable to Hosp."
         Object.Width           =   2117
      EndProperty
      BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   6
         Text            =   "Payable to MBM"
         Object.Width           =   2117
      EndProperty
      BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   7
         Text            =   "DOR (Case)"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(9) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   8
         Text            =   "DOA"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(10) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   9
         Text            =   "DOD"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(11) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   10
         Text            =   "Patient"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(12) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   11
         Text            =   "Hospital"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(13) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   12
         Text            =   "Bordx No"
         Object.Width           =   2117
      EndProperty
      BeginProperty ColumnHeader(14) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   13
         Text            =   "Bordx Amt"
         Object.Width           =   2117
      EndProperty
      BeginProperty ColumnHeader(15) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   14
         Text            =   "Bordx Date"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(16) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   15
         Text            =   "WriteOffAmt"
         Object.Width           =   2540
      EndProperty
   End
   Begin VB.Frame Frame1 
      Height          =   1185
      Left            =   60
      TabIndex        =   11
      Top             =   240
      Width           =   11655
      Begin VB.TextBox TxtCaseNo2 
         Height          =   315
         Left            =   3240
         MaxLength       =   15
         TabIndex        =   1
         Top             =   240
         Visible         =   0   'False
         Width           =   1575
      End
      Begin VB.TextBox TxtCaseNo1 
         Height          =   315
         Left            =   1200
         MaxLength       =   15
         TabIndex        =   0
         Top             =   240
         Width           =   1575
      End
      Begin VB.CheckBox ChkScan 
         Caption         =   "Use Scan"
         Height          =   255
         Left            =   2880
         TabIndex        =   12
         Top             =   240
         Value           =   1  'Checked
         Width           =   1095
      End
      Begin VB.ComboBox CboPayor 
         Height          =   315
         Left            =   1200
         Style           =   2  'Dropdown List
         TabIndex        =   2
         Top             =   525
         Width           =   3615
      End
      Begin VB.TextBox TxtMBMNo1 
         Height          =   315
         Left            =   1200
         MaxLength       =   15
         TabIndex        =   3
         Top             =   765
         Width           =   1575
      End
      Begin VB.TextBox TxtMBMNo2 
         Height          =   315
         Left            =   3240
         MaxLength       =   15
         TabIndex        =   4
         Top             =   765
         Width           =   1575
      End
      Begin VB.TextBox TxtPolicyNo1 
         Height          =   315
         Left            =   7200
         MaxLength       =   25
         TabIndex        =   5
         Top             =   195
         Width           =   1575
      End
      Begin VB.TextBox TxtPolicyNo2 
         Height          =   315
         Left            =   9240
         MaxLength       =   25
         TabIndex        =   6
         Top             =   195
         Width           =   1575
      End
      Begin VB.TextBox TxtBordxNo1 
         Height          =   315
         Left            =   7200
         MaxLength       =   20
         TabIndex        =   7
         Top             =   480
         Width           =   1575
      End
      Begin VB.TextBox TxtBordxNo2 
         Height          =   315
         Left            =   9240
         MaxLength       =   20
         TabIndex        =   8
         Top             =   480
         Width           =   1575
      End
      Begin VB.Label Label32 
         Caption         =   "To"
         Height          =   255
         Left            =   2880
         TabIndex        =   21
         Top             =   240
         Width           =   360
      End
      Begin VB.Label Label23 
         Caption         =   "To"
         Height          =   255
         Left            =   8880
         TabIndex        =   20
         Top             =   525
         Width           =   360
      End
      Begin VB.Label Label20 
         Caption         =   "To"
         Height          =   255
         Left            =   8880
         TabIndex        =   19
         Top             =   270
         Width           =   360
      End
      Begin VB.Label Label18 
         Caption         =   "To"
         Height          =   255
         Left            =   2880
         TabIndex        =   18
         Top             =   855
         Width           =   360
      End
      Begin VB.Label Label17 
         Caption         =   "Bordx No"
         Height          =   255
         Left            =   6120
         TabIndex        =   17
         Top             =   570
         Width           =   840
      End
      Begin VB.Label Label13 
         Caption         =   "Policy No"
         Height          =   255
         Left            =   6120
         TabIndex        =   16
         Top             =   270
         Width           =   840
      End
      Begin VB.Label Label7 
         Caption         =   "Member No"
         Height          =   255
         Left            =   240
         TabIndex        =   15
         Top             =   855
         Width           =   840
      End
      Begin VB.Label Label2 
         Caption         =   "Wks. No"
         Height          =   255
         Left            =   240
         TabIndex        =   14
         Top             =   270
         Width           =   840
      End
      Begin VB.Label Label1 
         Caption         =   "Payor"
         Height          =   255
         Left            =   240
         TabIndex        =   13
         Top             =   560
         Width           =   1215
      End
   End
   Begin VB.Frame Frame2 
      Height          =   735
      Left            =   60
      TabIndex        =   22
      Top             =   7200
      Width           =   11655
      Begin VB.TextBox LblTotalAmt 
         Height          =   285
         Left            =   3960
         TabIndex        =   34
         Top             =   240
         Width           =   1575
      End
      Begin VB.TextBox lblTotal 
         Height          =   285
         Left            =   1560
         TabIndex        =   33
         Top             =   360
         Width           =   735
      End
      Begin VB.TextBox lblCurrent 
         Height          =   285
         Left            =   120
         TabIndex        =   32
         Top             =   240
         Width           =   855
      End
      Begin VB.CommandButton CmdCheck 
         Caption         =   "&Check All"
         Height          =   375
         Left            =   7080
         TabIndex        =   9
         Top             =   240
         Width           =   855
      End
      Begin VB.CommandButton CmdUpdate 
         Caption         =   "&Update"
         Height          =   375
         Left            =   8640
         TabIndex        =   10
         Top             =   240
         Width           =   735
      End
      Begin VB.CommandButton CmdPrintFile 
         Caption         =   "Print to &File"
         Height          =   375
         Left            =   9360
         TabIndex        =   26
         Top             =   240
         Width           =   975
      End
      Begin VB.CommandButton CmdSearch 
         Caption         =   "&Search"
         Default         =   -1  'True
         Height          =   375
         Left            =   7920
         TabIndex        =   25
         Top             =   240
         Width           =   735
      End
      Begin VB.CommandButton CmdClear 
         Caption         =   "&Clear"
         Height          =   375
         Left            =   10320
         TabIndex        =   24
         Top             =   240
         Width           =   615
      End
      Begin VB.CommandButton CmdExit 
         Cancel          =   -1  'True
         Caption         =   "E&xit"
         Height          =   375
         Left            =   10920
         TabIndex        =   23
         Top             =   240
         Width           =   615
      End
      Begin VB.Label Label29 
         Caption         =   "Total Amt"
         Height          =   255
         Left            =   3120
         TabIndex        =   29
         Top             =   360
         Width           =   855
      End
      Begin VB.Label Label22 
         Caption         =   "of"
         Height          =   255
         Left            =   1200
         TabIndex        =   28
         Top             =   360
         Width           =   375
      End
      Begin VB.Label Label21 
         Caption         =   "Records"
         Height          =   255
         Left            =   2400
         TabIndex        =   27
         Top             =   360
         Width           =   1095
      End
   End
   Begin VB.Label Label28 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00404040&
      Caption         =   "Claim Write Off"
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
      TabIndex        =   31
      Top             =   0
      Width           =   11805
   End
End
Attribute VB_Name = "frmCLMWriteOff"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub ChkScan_Click()
    If ChkScan.Value = 1 Then
        lstWorksheetTrackList.ListItems.Clear
        Label32.Visible = False
        TxtCaseNo2.Visible = False
        TxtCaseNo1.SetFocus
        ChkScan.Left = 2950
    Else
        lstWorksheetTrackList.ListItems.Clear
        Label32.Visible = True
        TxtCaseNo2.Visible = True
        TxtCaseNo1.SetFocus
        ChkScan.Left = 4900
    End If
End Sub

Private Sub CmdCheck_Click()
    SetCheckAllItems True
End Sub

Private Sub CmdClear_Click()
    Call ClearForm(Me)
    lstWorksheetTrackList.ListItems.Clear
    CboPayor.Clear
    Current2 = 0
    TotalAmt2 = 0
    TxtCaseNo1.SetFocus
    ChkScan.Value = 1
    Check = False
    Check2 = False
    lblCurrent = 0
    lblTotal = 0
    LblTotalAmt = 0
    ReDim ClaimNoArray(0)
    Call ComboListing

End Sub

Private Sub CmdExit_Click()
    Unload Me
End Sub

Private Sub CmdPrintFile_Click()
    'Call ExportToExcelAnalysis15(lstWorksheetTrackList)
    Screen.MousePointer = vbHourglass
        If lstWorksheetTrackList.ListItems.Count = 0 Then
            MsgBox "There aren't any items In the listview selected." _
            , vbOKOnly + vbInformation, "Export Failed"
        Else
            Call ExportListViewtoExcel(lstWorksheetTrackList)
        End If
    Screen.MousePointer = vbDefault
End Sub

Private Sub CmdSearch_Click()
ScanCode = TxtCaseNo1
    CmdClear.Enabled = False
    CmdExit.Enabled = False
    CmdPrintFile.Enabled = False
    CmdSearch.Enabled = False
    ItemChecked = False
    ReDim ClaimNoArray(0)
    
    If ChkScan.Value = 1 Then
    
        If TxtCaseNo1 = "" Then
            MsgBox "Please Enter Worksheet No", vbOKOnly + vbCritical, DialogBoxTitle
            TxtCaseNo1.SetFocus
            CmdClear.Enabled = True
            CmdExit.Enabled = True
            CmdPrintFile.Enabled = True
        Else
        
            Screen.MousePointer = vbHourglass
            
            listrecord = lstWorksheetTrackList.ListItems.Count
            If listrecord <> 0 Then
                
                For listloop1 = 1 To listrecord
                    Call CaseNoCheck
                Next listloop1
            End If
                If Check = True Then
                    MsgBox "Data has been uploaded!", vbOKOnly + vbCritical
                ElseIf Check = False Then
                    CmdClear.Enabled = True
                    CmdExit.Enabled = True
                    CmdPrintFile.Enabled = True
                    Call SearchRecord
                End If
                Check = False
                CmdClear.Enabled = True
                CmdExit.Enabled = True
                CmdPrintFile.Enabled = True
            Screen.MousePointer = Default
            TxtCaseNo1 = ""
            TxtCaseNo1.SetFocus
        End If
        TxtCaseNo1.SetFocus
    Else
        Screen.MousePointer = vbHourglass
            'medixmasterconnWS.Open "File Name=" & BOSPath & "Conn\mediconnectWS.UDL"
                Call SearchRecord
           ' medixmasterconnWS.Close
            'Set medixmasterconnWS = Nothing
        Screen.MousePointer = Default
            
        CmdClear.Enabled = True
        CmdExit.Enabled = True
        CmdPrintFile.Enabled = True
        
    End If
    CmdSearch.Enabled = True
End Sub

Private Function CaseNoCheck()
Dim CasNumber As String
    CasNumber = Trim(Mid(lstWorksheetTrackList.ListItems(listloop1).SubItems(2), 1, IIf(InStr(1, lstWorksheetTrackList.ListItems(listloop1).SubItems(2), "-") = 0, 10, InStr(1, lstWorksheetTrackList.ListItems(listloop1).SubItems(2), "-") - 1)))
    If Trim(ChkStr(TxtCaseNo1)) = ChkStr(CasNumber) Then
        Check = True
    End If
End Function

Private Function SearchRecord()
Dim strsql As String
Dim strAppend As String
Dim Sql As String

    
    strAppend = ""

    If Len(Trim(CboPayor)) Then
        strAppend = strAppend + " And PAYCode = '" & Trim(Mid(CboPayor, 1, IIf(InStr(1, CboPayor, "-") = 0, 4, InStr(1, CboPayor, "-") - 1))) & "'"
    End If
    
    If ChkScan.Value = 1 Then
        If Len(Trim(TxtCaseNo1)) Then
            strAppend = strAppend + " AND CASNumber = '" & Trim(TxtCaseNo1) & "'"
        End If
    Else
        If Len(Trim(TxtCaseNo1)) Then
            strAppend = strAppend + " AND CASNumber >= '" & Trim(TxtCaseNo1) & "'"
        End If
        
        If Len(Trim(TxtCaseNo2)) Then
            strAppend = strAppend + " AND CASNumber <= '" & Trim(TxtCaseNo2) & "'"
        End If
    End If
        
    If Len(Trim(TxtPolicyNo1)) Then
        strAppend = strAppend + " AND MBMPolicyNo >= '" & Trim(TxtPolicyNo1) & "'"
    End If
    
    If Len(Trim(TxtPolicyNo2)) Then
        strAppend = strAppend + " AND MBMPolicyNo <= '" & Trim(TxtPolicyNo2) & "'"
    End If
    
    If Len(Trim(TxtMBMNo1)) Then
        strAppend = strAppend + " AND MBMNumber >= '" & Trim(TxtMBMNo1) & "'"
    End If
    
    If Len(Trim(TxtMBMNo2)) Then
        strAppend = strAppend + " AND MBMNumber <= '" & Trim(TxtMBMNo2) & "'"
    End If
    
    If Len(Trim(TxtBordxNo1)) Then
        strAppend = strAppend + " AND BordxRefNo >= '" & Trim(TxtBordxNo1) & "'"
    End If
    
    If Len(Trim(TxtBordxNo2)) Then
        strAppend = strAppend + " AND BordxRefNo <= '" & Trim(TxtBordxNo2) & "'"
    End If
    

    searchCriteria = strAppend
    
    Call WorksheetTrackingListing(strAppend)

End Function

Private Sub WorksheetTrackingListing(Optional strAppend As String)
Dim rst2 As New ADODB.Recordset
Dim Record As ListItem
Dim Sql As String
Dim total As String
Dim Current As String
Dim TotalAmt As Variant
Dim strCount

    total = 0
    Current = 0
    
    If ChkScan.Value = 0 Then
        lstWorksheetTrackList.ListItems.Clear
    End If
        
    Sql = " Select  CASNumber, ClaimBordxStatus, claimversion," & _
          " ClaimStatus, CLMTotalApproved, CLMPayableByMedix2H, " & _
          " CLMPayableByMedix2M,  PatientName, CaseRegDate, CASDischargeDate, " & _
          " HOSName, BordxRefNo, BordxClaimAmt, BordxDate, CASDateAdmitted " & _
          " From VIEW_WKSTrackingNew where ClaimBordxStatus = '0' "
    
    strCount = "select count(*) as totalRecord From VIEW_WKSTrackingNew where 0 = 0 "
    If Len(Trim(strAppend)) Then
        Sql = Sql + strAppend
        strCount = strCount + strAppend
    End If

    Set strCount = medixmasterconnWS.Execute(strCount)
    rst2.Open Sql, medixmasterconnWS, adOpenKeyset, adLockOptimistic
    
    If strCount!TotalRecord = 0 Then
        MsgBox "No record found!", vbOKOnly + vbCritical, DialogBoxTitle
        CmdClear.Enabled = True
        CmdExit.Enabled = True
        CmdPrintFile.Enabled = True
    Else
         
    Do
        Do Until rst2.EOF
         total = strCount!TotalRecord
    
    
        lblTotal = total
        
            With rst2
                Set Record = lstWorksheetTrackList.ListItems.Add(, , "")
                If Len(!CasNumber & "-" & !claimversion) Then
                    Record.SubItems(1) = (!CasNumber & "-" & !claimversion)
                End If
                
                If Len(!CasNumber & "-" & !claimversion) Then
                    Record.SubItems(2) = (!CasNumber & "-" & !claimversion)
                End If
                If Len(!ClaimStatus) Then
                   Record.SubItems(3) = !ClaimStatus
                End If
                If Len(!CLMTotalApproved) Then
                    Record.SubItems(4) = Format(!CLMTotalApproved, "#,##0.00")
                End If
                If Len(!CLMPayableByMedix2H) Then
                    Record.SubItems(5) = Format(!CLMPayableByMedix2H, "#,##0.00")
                End If
                If Len(!CLMPayableByMedix2M) Then
                    Record.SubItems(6) = Format(!CLMPayableByMedix2M, "#,##0.00")
                End If
                If Len(rst2!CaseRegDate) Then
                    Record.SubItems(7) = Format(rst2!CaseRegDate, "dd/mm/yyyy")
                End If
                If Len(rst2!CASDateAdmitted) Then
                   Record.SubItems(8) = Format(rst2!CASDateAdmitted, "dd/mm/yyyy")
                End If
                If Len(rst2!CASDischargeDate) Then
                    Record.SubItems(9) = Format(rst2!CASDischargeDate, "dd/mm/yyyy")
                End If
                If Len(rst2!PatientName) Then
                    Record.SubItems(10) = rst2!PatientName
                End If
                If Len(rst2!HosName) Then
                    Record.SubItems(11) = rst2!HosName
                End If
                If Len(rst2!BordxRefNo) Then
                    Record.SubItems(12) = rst2!BordxRefNo
                End If
                If Len(rst2!BordxClaimAmt) Then
                    Record.SubItems(13) = Format(rst2!BordxClaimAmt, "#,##0.00")
                End If
                If Len(rst2!BordxDate) Then
                    Record.SubItems(14) = Format(rst2!BordxDate, "dd/mm/yyyy")
                End If
                                
                Current = Current + 1
                lblCurrent = Current
                If Len(rst2!CLMTotalApproved) Then
                    TotalAmt = TotalAmt + rst2!CLMTotalApproved
                End If
                LblTotalAmt = Format(TotalAmt, "#,##0.00")
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
    CmdClear.Enabled = True
    CmdExit.Enabled = True
    CmdPrintFile.Enabled = True
End Sub


Private Sub CmdUpdate_Click()
Dim CASNo As String
Dim CASVer As String
Dim AppAmt As String
Dim status As String
Dim Update
Dim ItemChecked As Boolean
Dim i As Integer
    
ItemChecked = False
CASNo = ""
CASVer = ""
status = ""
AppAmt = ""
    For i = 1 To lstWorksheetTrackList.ListItems.Count
        If lstWorksheetTrackList.ListItems(i).Checked = True Then
            ItemChecked = True
        'Exit For
        End If
    Next i
    
    If lstWorksheetTrackList.ListItems.Count = 0 Or ItemChecked = False Then
        MsgBox "Please select the record", vbOKOnly + vbCritical
    Else
       ' medixmasterconnWS.Open "File Name=" & BOSPath & "Conn\mediconnectWS.UDL"
        Update = MsgBox("Do you want to update the record?", vbYesNo + vbExclamation)
            If Update = vbYes Then
                i = 0
                ItemChecked = False
                For i = 1 To lstWorksheetTrackList.ListItems.Count
                    If lstWorksheetTrackList.ListItems(i).Checked = True Then
                        ItemChecked = True
                    Exit For
                    End If
                Next i
                
                If ItemChecked = True Then
                    For i = 1 To lstWorksheetTrackList.ListItems.Count
                        If lstWorksheetTrackList.ListItems(i).Checked = True Then
                            CASNo = Trim(Mid(lstWorksheetTrackList.ListItems(i).ListSubItems(1).Text, 1, IIf(InStr(1, lstWorksheetTrackList.ListItems(i).ListSubItems(1).Text, "-") = 0, 10, InStr(1, lstWorksheetTrackList.ListItems(i).ListSubItems(1).Text, "-") - 1)))
                            CASVer = Trim(Mid(lstWorksheetTrackList.ListItems(i).ListSubItems(1).Text, InStr(1, lstWorksheetTrackList.ListItems(i).ListSubItems(1).Text, "-") + 1, Len(lstWorksheetTrackList.ListItems(i).ListSubItems(1).Text) - InStr(1, lstWorksheetTrackList.ListItems(i).ListSubItems(1).Text, "-")))
                            AppAmt = Trim(lstWorksheetTrackList.ListItems(i).SubItems(3))
                            Call UpdateRecord(CASNo, CInt(CASVer), CInt(AppAmt))
                            
                            CASNo = ""
                            CASVer = ""
                        End If
                    Next i
                    
                    MsgBox "Record Updated, vbOKOnly + vbInformation, DialogBoxTitle"

                    lstWorksheetTrackList.ListItems.Clear
                End If
            End If
        'medixmasterconnWS.Close
       ' Set medixmasterconnWS = Nothing
    End If

End Sub

Private Sub Form_Load()
    Call ComboListing
End Sub


Public Function ProcessCheckedItem(Item As ListItem, Optional Partial As Boolean)
                
    If Item.Checked = True Then
        If Trim(Item.SubItems(1)) = "Checked" Then
            MsgBox "Record has been checked!", vbCritical
            Item.Checked = False
        Else
            Item.SubItems(1) = "Checked"
        End If
    Else
        Item.SubItems(1) = ""
    End If
End Function

Private Sub UpdateRecord(CASNo As String, CASVer As Integer, AppAmt As Double)
Dim strsql As String
Dim CLM As New ADODB.Recordset
    
    strsql = " Select CASNumber,ClaimVersion,CLMWOAmt,CLMWODate,CLMWOBy from Claim " & _
                " Where CASNumber = '" & Trim(CASNo) & "' AND ClaimVersion = '" & Trim(CASVer) & "' AND CLMWOAmt IS NULL"
    CLM.Open strsql, medixmasterconnWS, adOpenKeyset, adLockOptimistic
    
     If CLM.recordCount > 0 Then
        With CLM
            !CLMWOAmt = AppAmt
            !CLMWODate = Date
            !CLMWOBy = User
        End With
        CLM.Update
    End If
    CLM.Close
    Set CLM = Nothing

End Sub

Private Sub SetCheckAllItems(bState As Boolean)
Dim LV As LVITEM
Dim lvCount As Long
Dim lvIndex As Long
Dim lvState As Long
Dim r As Long
    lvState = IIf(bState, &H2000, &H1000)
    lvCount = lstWorksheetTrackList.ListItems.Count - 1
    Do
        With LV
        .mask = LVIF_STATE
        .state = lvState
        .stateMask = LVIS_STATEIMAGEMASK
        End With
        r = SendMessageAny(lstWorksheetTrackList.hwnd, LVM_SETITEMSTATE, lvIndex, LV)
        lvIndex = lvIndex + 1
    Loop Until lvIndex > lvCount
End Sub

Private Sub ComboListing()
Dim PAY As New ADODB.Recordset
Dim cmd As New ADODB.Command

   ' medixmasterconn.Open "File Name=" & BOSPath & "Conn\mediconnect.UDL"
    
    Set cmd.ActiveConnection = medixmasterconn
    cmd.CommandText = "SearchPayor"
    cmd.CommandType = adCmdStoredProc
    cmd.CommandTimeout = 15
    PAY.CursorLocation = adUseClient
    PAY.CursorType = adOpenForwardOnly
    PAY.CacheSize = 100
    Set PAY = cmd.Execute
        
    Do Until PAY.EOF
        With PAY
            CboPayor.AddItem !PAYCode & " - " & Trim(!PAYCompanyName)
        End With
        PAY.MoveNext
    Loop
    PAY.Close
    Set PAY = Nothing
        
   ' medixmasterconn.Close
  '  Set medixmasterconn = Nothing

End Sub

Private Sub lstWorksheetTrackList_ItemCheck(ByVal Item As MSComctlLib.ListItem)
    Call ProcessCheckedItem(Item)
End Sub

