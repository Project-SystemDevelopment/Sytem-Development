VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmContraAllocation 
   Caption         =   "Contra Allocation"
   ClientHeight    =   7740
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   11880
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   7740
   ScaleWidth      =   11880
   Begin VB.Frame Frame2 
      Height          =   615
      Left            =   120
      TabIndex        =   19
      Top             =   7080
      Width           =   11775
      Begin VB.CommandButton cmdCheckALL 
         Caption         =   "&Chechk All"
         Height          =   375
         Left            =   8520
         TabIndex        =   30
         Top             =   150
         Width           =   975
      End
      Begin VB.CommandButton cmdSearch 
         Caption         =   "&Search"
         Default         =   -1  'True
         Height          =   375
         Left            =   9480
         TabIndex        =   22
         Top             =   150
         Width           =   735
      End
      Begin VB.CommandButton cmdUpdate 
         Caption         =   "&Update"
         Height          =   375
         Left            =   10200
         TabIndex        =   21
         Top             =   150
         Width           =   735
      End
      Begin VB.CommandButton cmdExit 
         Cancel          =   -1  'True
         Caption         =   "&Exit"
         Height          =   375
         Left            =   10920
         TabIndex        =   20
         Top             =   150
         Width           =   735
      End
      Begin VB.Label lblAllocationNo 
         Caption         =   "Allocation No :"
         Height          =   255
         Left            =   6720
         TabIndex        =   32
         Top             =   240
         Visible         =   0   'False
         Width           =   1095
      End
      Begin VB.Label lblNumber 
         Caption         =   "Number"
         ForeColor       =   &H000000FF&
         Height          =   255
         Left            =   7800
         TabIndex        =   31
         Top             =   240
         Visible         =   0   'False
         Width           =   1215
      End
   End
   Begin VB.Frame Frame1 
      Height          =   1095
      Left            =   120
      TabIndex        =   11
      Top             =   240
      Width           =   11655
      Begin VB.TextBox txtContraCASNoFR 
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
         Left            =   7080
         MaxLength       =   15
         TabIndex        =   5
         Top             =   195
         Width           =   1215
      End
      Begin VB.TextBox txtContraCASNoTo 
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
         Left            =   8640
         MaxLength       =   15
         TabIndex        =   6
         Top             =   195
         Width           =   1215
      End
      Begin VB.TextBox txtCaseNoTo 
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
         Left            =   2760
         MaxLength       =   15
         TabIndex        =   1
         Top             =   195
         Width           =   1575
      End
      Begin VB.TextBox txtALLNoTo 
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
         Left            =   8640
         MaxLength       =   15
         TabIndex        =   8
         Top             =   435
         Width           =   1215
      End
      Begin VB.CheckBox ChkScan 
         Caption         =   "Use Scan"
         Height          =   255
         Left            =   4440
         TabIndex        =   23
         Top             =   180
         Width           =   1095
      End
      Begin VB.TextBox txtCaseNoFr 
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
         Left            =   840
         MaxLength       =   15
         TabIndex        =   0
         Top             =   195
         Width           =   1575
      End
      Begin VB.TextBox txtALLNoFr 
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
         Left            =   7080
         MaxLength       =   15
         TabIndex        =   7
         Top             =   435
         Width           =   1215
      End
      Begin MSMask.MaskEdBox txtALLDateFR 
         Height          =   285
         Left            =   7080
         TabIndex        =   9
         Top             =   690
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   503
         _Version        =   393216
         MaxLength       =   10
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Mask            =   "##/##/####"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox txtALLDateTo 
         Height          =   285
         Left            =   8640
         TabIndex        =   10
         Top             =   690
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   503
         _Version        =   393216
         MaxLength       =   10
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Mask            =   "##/##/####"
         PromptChar      =   "_"
      End
      Begin VB.TextBox txtREFNoFR 
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
         Left            =   840
         MaxLength       =   15
         TabIndex        =   2
         Top             =   450
         Width           =   1575
      End
      Begin VB.TextBox txtREFNoTo 
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
         Left            =   2760
         MaxLength       =   15
         TabIndex        =   3
         Top             =   450
         Width           =   1575
      End
      Begin VB.ComboBox cboStatus 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   840
         Sorted          =   -1  'True
         TabIndex        =   4
         Tag             =   "Thanks"
         Text            =   "0 - ALL"
         Top             =   690
         Width           =   3495
      End
      Begin VB.Label Label9 
         Caption         =   "To"
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
         Left            =   8400
         TabIndex        =   29
         Top             =   240
         Width           =   555
      End
      Begin VB.Label Label8 
         Caption         =   "Contra Case No"
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
         Left            =   5880
         TabIndex        =   28
         Top             =   240
         Width           =   1245
      End
      Begin VB.Label Label7 
         Caption         =   "To"
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
         Left            =   2520
         TabIndex        =   27
         Top             =   450
         Width           =   285
      End
      Begin VB.Label Label6 
         Caption         =   "To"
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
         Left            =   2520
         TabIndex        =   26
         Top             =   240
         Width           =   555
      End
      Begin VB.Label Label5 
         Caption         =   "To"
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
         Left            =   8400
         TabIndex        =   25
         Top             =   480
         Width           =   555
      End
      Begin VB.Label Label4 
         Caption         =   "Ref No"
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
         TabIndex        =   24
         Top             =   480
         Width           =   1005
      End
      Begin VB.Label Label2 
         Caption         =   "To"
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
         Left            =   8400
         TabIndex        =   16
         Top             =   720
         Width           =   555
      End
      Begin VB.Label Label22 
         Caption         =   "Alloc. Date"
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
         Left            =   5880
         TabIndex        =   15
         Top             =   720
         Width           =   1275
      End
      Begin VB.Label Label1 
         Caption         =   "Alloc.  No"
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
         Left            =   5880
         TabIndex        =   14
         Top             =   480
         Width           =   1005
      End
      Begin VB.Label Label19 
         Caption         =   "Case No"
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
         Top             =   225
         Width           =   1005
      End
      Begin VB.Label Label10 
         Caption         =   "Status"
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
         TabIndex        =   12
         Top             =   720
         Width           =   1770
      End
   End
   Begin MSComctlLib.ListView lstContraAllocation 
      Height          =   5535
      Left            =   120
      TabIndex        =   18
      Top             =   1560
      Width           =   11820
      _ExtentX        =   20849
      _ExtentY        =   9763
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
      NumItems        =   8
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Check"
         Object.Width           =   1764
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Contra Status"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "Claim No"
         Object.Width           =   2117
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   3
         Text            =   "Claim Amount"
         Object.Width           =   2469
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   4
         Text            =   "Contra Claim No"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   5
         Text            =   "Contra Claim Amount"
         Object.Width           =   2999
      EndProperty
      BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   6
         Text            =   "Allocation No"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   7
         Text            =   "Allocation Date"
         Object.Width           =   2540
      EndProperty
   End
   Begin VB.Label Label28 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00404040&
      Caption         =   "Contra Allocation"
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
      TabIndex        =   33
      Top             =   0
      Width           =   11805
   End
   Begin VB.Label Label3 
      Caption         =   "Listing"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   225
      Left            =   120
      TabIndex        =   17
      Top             =   1320
      Width           =   720
   End
End
Attribute VB_Name = "frmContraAllocation"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public Source As Integer
Public intMBMAccount As Integer
Public intExit As Integer
Dim one As Boolean
Private Hospitals As String
Dim ScanCode As String
Dim Check As Boolean
Dim ContraNo As String

Private Sub ChkScan_Click()
    lstContraAllocation.ListItems.Clear
    If ChkScan.Value = 1 Then
        Label6.Visible = False
        txtCaseNoTo.Visible = False
        ChkScan.Left = 2500
        txtCaseNoFr = ""
        txtCaseNoFr.SetFocus
    Else
        Label6.Visible = True
        txtCaseNoTo.Visible = True
        txtCaseNoTo.Left = 2760
        ChkScan.Left = 4450
        txtCaseNoFr.SetFocus
    End If
End Sub

Private Sub cmdCheckALL_Click()
    SetCheckAllItems True
End Sub
Private Sub SetCheckAllItems(bState As Boolean)
Dim LV As LVITEM
Dim lvCount As Long
Dim lvIndex As Long
Dim lvState As Long
Dim R As Long
    lvState = IIf(bState, &H2000, &H1000)
    lvCount = lstContraAllocation.ListItems.Count - 1
    Do
        With LV
        .mask = LVIF_STATE
        .state = lvState
        .stateMask = LVIS_STATEIMAGEMASK
        End With
        R = SendMessageAny(lstContraAllocation.hwnd, LVM_SETITEMSTATE, lvIndex, LV)
        lvIndex = lvIndex + 1
    Loop Until lvIndex > lvCount
End Sub


Private Sub cmdExit_Click()
    Unload Me
End Sub

Private Sub cmdSearch_Click()
    lblAllocationNo.Visible = False
    lblNumber.Visible = False
    Screen.MousePointer = vbHourglass
    LblAmt = 0
    one = False
    'Erase CaseNoArray
    ReDim CaseNoArray(0)
    cmdSearch.Enabled = False
   ' medixmasterconnPayment.Open "File Name=" & BOSPath & "Conn\mediconnectPayment.UDL"

    If ChkScan.Value = 1 Then
        If txtCaseNoFr = "" Then
            MsgBox "Please Enter Case No", vbOKOnly + vbCritical, DialogBoxTitle
            txtCaseNoFr.SetFocus
        Else
            listrecord = lstContraAllocation.ListItems.Count
            If listrecord <> 0 Then
                
                For listloop1 = 1 To listrecord
                    Call CaseNoCheck
                Next listloop1
            End If
                If Check = True Then
                    MsgBox "Data has been uploaded!", vbOKOnly + vbCritical
                ElseIf Check = False Then
                    Call SearchRecord
                End If
                Check = False
            Screen.MousePointer = Default
            txtCaseNoFr = ""
            txtCaseNoFr.SetFocus
        End If
        'txtCaseNoFr.SetFocus
    Else
        Call SearchRecord
    End If
    'End If
 '   medixmasterconnPayment.Close
 '   Set medixmasterconnPayment = Nothing

    Screen.MousePointer = Default
    cmdSearch.Enabled = True
End Sub

Private Function CaseNoCheck()
    If Trim(ChkStr(txtCaseNoFr)) = Trim(Mid(lstContraAllocation.ListItems(listloop1).SubItems(10), 1, IIf(InStr(1, lstContraAllocation.ListItems(listloop1).SubItems(10), "-") = 0, 10, InStr(1, lstContraAllocation.ListItems(listloop1).SubItems(10), "-") - 1))) Then
        Check = True
    End If
End Function

Private Sub cmdUpdate_Click()
Dim CASNo As String
Dim CASVer As String
Dim status As String
Dim Update
Dim ItemChecked As Boolean
Dim i As Integer
    
    ItemChecked = False
    CASNo = ""
    CASVer = ""
    status = ""
    cmdUpdate.Enabled = False
    For i = 1 To lstContraAllocation.ListItems.Count
        If lstContraAllocation.ListItems(i).Checked = True Then
            ItemChecked = True
        'Exit For
        End If
    Next i
    
    If lstContraAllocation.ListItems.Count = 0 Or ItemChecked = False Then
        MsgBox "Please select the record", vbOKOnly + vbCritical
    Else
        Update = MsgBox("Do you want to update the record?", vbYesNo + vbExclamation)
            If Update = vbYes Then
                i = 0
                ItemChecked = False
                ContraNo = GenerateContraNo
                For i = 1 To lstContraAllocation.ListItems.Count
                    If lstContraAllocation.ListItems(i).Checked = True Then
                        ItemChecked = True
                    Exit For
                    End If
                Next i
                
                If ItemChecked = True Then
                    For i = 1 To lstContraAllocation.ListItems.Count
                        If lstContraAllocation.ListItems(i).Checked = True Then
                            CASNo = Trim(Mid(lstContraAllocation.ListItems(i).ListSubItems(2).Text, 1, IIf(InStr(1, lstContraAllocation.ListItems(i).ListSubItems(2).Text, "-") = 0, 10, InStr(1, lstContraAllocation.ListItems(i).ListSubItems(2).Text, "-") - 1)))
                            CASVer = Trim(Mid(lstContraAllocation.ListItems(i).ListSubItems(2).Text, InStr(1, lstContraAllocation.ListItems(i).ListSubItems(2).Text, "-") + 1, Len(lstContraAllocation.ListItems(i).ListSubItems(2).Text) - InStr(1, lstContraAllocation.ListItems(i).ListSubItems(2).Text, "-")))
                            
                           ' medixmasterconnPayment.Open "File Name=" & BOSPath & "Conn\mediconnectPayment.UDL"
                                Call UpdateRecord(CASNo, CInt(CASVer))
                          '  medixmasterconnPayment.Close
                          '  Set medixmasterconnPayment = Nothing
                            CASNo = ""
                            CASVer = ""
                        End If
                    Next i
                    
                    lblAllocationNo.Visible = True
                    lblNumber = ""
                    lblNumber = ContraNo
                    lblNumber.Visible = True
                    
                    MsgBox "Record Updated, Your Allocation Number is : " & ContraNo, vbOKOnly + vbInformation, DialogBoxTitle

                    lstContraAllocation.ListItems.Clear
                End If
            End If
    End If
    cmdUpdate.Enabled = True
End Sub

Private Sub Form_Load()
    Call ComboList
End Sub

Private Sub Form_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Shift = 0 Then
        Select Case KeyCode
            Case vbKeyF6
                KeyCode = 0
                PrintScreenSnapShot
        End Select
    End If
End Sub

Private Function SearchRecord()
Dim strsql As String
Dim strAppend As String
Dim sql As String
    
    strAppend = ""
        
    If Len(Trim(txtREFNoFR)) Then
        strAppend = strAppend + " And RefNo >= '" & Trim(txtREFNoFR) & "'"
    End If
    
    If Len(Trim(txtREFNoFR)) Then
        strAppend = strAppend + " And RefNo <= '" & Trim(txtREFNoTo) & "'"
    End If
    
    
    If Mid(cboStatus, 1, 1) = 1 Then
        strAppend = strAppend + " and ContraStatus = 'A' "
    End If
 
    If Mid(cboStatus, 1, 1) = 2 Then
        strAppend = strAppend + " and (ContraStatus = '' or ContraStatus is null) "
    End If
 
            
    If ChkScan.Value = 1 Then
        strAppend = strAppend + " AND CaseNo = '" & Trim(txtCaseNoFr) & "'"
    Else
        If Len(txtCaseNoFr) Then
            strAppend = strAppend + " AND CaseNo >= '" & Trim(txtCaseNoFr) & "'"
        End If
    
        If Len(txtCaseNoTo) Then
            strAppend = strAppend + " AND CaseNo <= '" & Trim(txtCaseNoTo) & "'"
        End If
    End If
    
    If Len(txtContraCASNoFR) Then
        strAppend = strAppend + " AND ContraCaseNo >= '" & Trim(txtContraCASNoFR) & "'"
    End If

    If Len(txtContraCASNoTo) Then
        strAppend = strAppend + " AND ContraCaseNo <= '" & Trim(txtContraCASNoTo) & "'"
    End If
    
    If Len(txtALLNoFr) Then
        strAppend = strAppend + " AND AllocationNo >= '" & Trim(txtALLNoFr) & "'"
    End If

    If Len(txtALLNoTo) Then
        strAppend = strAppend + " AND AllocationNo <= '" & Trim(txtALLNoTo) & "'"
    End If
    
    If Len(Trim(txtALLDateFR)) > 0 And IsDate(txtALLDateFR) = True Then
        strAppend = strAppend + " And AllocationDate >= '" & Format(txtALLDateFR, "mm/dd/yyyy") & "'"
    End If
    
    If Len(Trim(txtALLDateTo)) > 0 And IsDate(txtALLDateTo) = True Then
        strAppend = strAppend + " And AllocationDate <= '" & Format(txtALLDateTo, "mm/dd/yyyy") & "'"
    End If
    
    Call ContraListing(strAppend)

End Function

Private Sub ContraListing(Optional strAppend As String)
Dim rst2 As New ADODB.Recordset
Dim Record As ListItem
Dim sql As String
Dim strsqlPAY As String
Dim PAY As New ADODB.Recordset
Dim INVAmt As Variant
Dim total As String
Dim Current As String
Dim TotalAmt As Variant
Dim Check As Boolean
Dim RECFull As Boolean
Dim RECPartial As String
Dim CASNumber As String
Dim WSVersion As String

   total = 0
   Current = 0
   If ChkScan.Value = 0 Then
        lstContraAllocation.ListItems.Clear
   End If

    sql = "SELECT * From ContraAllocation where 0 = 0 "
    If Len(Trim(strAppend)) Then
        sql = sql + strAppend
    End If
    
    rst2.Open sql, medixmasterconnPayment, adOpenKeyset, adLockOptimistic
    
    If rst2.recordCount = 0 Then
        MsgBox "No record found!", vbOKOnly + vbCritical, DialogBoxTitle
    Else
    
    Do
        Do Until rst2.EOF
        total = rst2.recordCount
        lblTotal = total
        With rst2
            Set Record = lstContraAllocation.ListItems.Add(, , "")
            
            If Trim(!ContraStatus) = "A" Then
                Record.SubItems(1) = "ALLOCATED"
            Else
                Record.SubItems(1) = ""
            End If

            If Len(rst2!CaseNo) Then
                Record.SubItems(2) = rst2!CaseNo & "-" & rst2!VersionNo
            End If
            
            If Len(rst2!CaseAmt) Then
                Record.SubItems(3) = rst2!CaseAmt
            End If
                                          
            If Len(rst2!ContraCaseNo) Then
                Record.SubItems(4) = Trim(rst2!ContraCaseNo) & "-" & Trim(rst2!ContraVersionNo)
            End If
            
            If Len(rst2!ContraCaseAmt) Then
                Record.SubItems(5) = Format(rst2!ContraCaseAmt, "#,##0.00")
            End If
            
            If Len(rst2!AllocationNo) Then
                Record.SubItems(6) = rst2!AllocationNo
            End If
            
            If Len(rst2!AllocationDate) Then
                Record.SubItems(7) = Trim(rst2!AllocationDate)
            End If
            
            Current = Current + 1
            lblCurrent = Current
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

Public Function ProcessCheckedItem(Item As ListItem, Optional Partial As Boolean)
                
    If Item.Checked = True Then
        'lstContraAllocation.SelectedItem.Selected = False
        'Item.Selected = True
        If Trim(Item.SubItems(1)) = "ALLOCATED" Then
            MsgBox "Record has been allocated!", vbCritical
            Item.Checked = False
        Else
            Item.SubItems(1) = "ALLOCATED"
        End If
    Else
        Item.SubItems(1) = ""
    End If
End Function

Private Sub lstContraAllocation_ItemCheck(ByVal Item As MSComctlLib.ListItem)
    Call ProcessCheckedItem(Item)
End Sub

Private Sub UpdateRecord(CASNo As String, CASVer As Integer)
Dim Update
Dim strsql As String
Dim CON As New ADODB.Recordset
    
    strsql = " Select * from ContraAllocation " & _
                " Where CaseNo = '" & Trim(CASNo) & "' AND VersionNo = '" & Trim(CASVer) & "'"
    CON.Open strsql, medixmasterconnPayment, adOpenKeyset, adLockOptimistic
    
     If CON.recordCount > 0 Then
        With CON
            !AllocationNo = ContraNo
            !AllocationDate = Date
            !ContraStatus = "A"
        End With
        CON.Update
    End If
    CON.Close
    Set CON = Nothing

End Sub

Private Sub ComboList()

    cboStatus.AddItem "0" & " - " & "ALL"
    cboStatus.AddItem "1" & " - " & "ALLOCATED"
    cboStatus.AddItem "2" & " - " & "NOT ALLOCATED"
End Sub
