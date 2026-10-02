VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "crystl32.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Begin VB.Form FrmPrintBordxEx 
   Caption         =   "Print Bordx (Group Company Excess)"
   ClientHeight    =   8595
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   11880
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   8595
   ScaleWidth      =   11880
   WindowState     =   2  'Maximized
   Begin VB.Frame Frame1 
      Height          =   900
      Left            =   120
      TabIndex        =   13
      Top             =   240
      Width           =   11655
      Begin VB.TextBox TxtPlan2 
         Height          =   285
         Left            =   8880
         MaxLength       =   4
         TabIndex        =   5
         Top             =   150
         Width           =   1095
      End
      Begin VB.TextBox TxtPlan1 
         Height          =   285
         Left            =   7200
         MaxLength       =   4
         TabIndex        =   4
         Top             =   150
         Width           =   1095
      End
      Begin VB.CheckBox ChkScan1 
         Caption         =   "Use Scan"
         Height          =   255
         Left            =   6120
         TabIndex        =   3
         Top             =   530
         Value           =   1  'Checked
         Width           =   1095
      End
      Begin VB.TextBox TxtBordxNo2 
         Height          =   285
         Left            =   3960
         MaxLength       =   20
         TabIndex        =   1
         Top             =   480
         Width           =   2055
      End
      Begin VB.ComboBox CboPayor 
         Height          =   315
         ItemData        =   "FrmPrintBordxEx.frx":0000
         Left            =   1320
         List            =   "FrmPrintBordxEx.frx":0002
         Style           =   2  'Dropdown List
         TabIndex        =   2
         Top             =   150
         Width           =   4695
      End
      Begin VB.TextBox TxtBordxNo1 
         Height          =   285
         Left            =   1320
         MaxLength       =   20
         TabIndex        =   0
         Top             =   480
         Width           =   2055
      End
      Begin VB.Label Label10 
         Caption         =   "To"
         Height          =   255
         Left            =   8520
         TabIndex        =   24
         Top             =   180
         Width           =   255
      End
      Begin VB.Label Label9 
         Caption         =   "Plan Code"
         Height          =   255
         Left            =   6360
         TabIndex        =   23
         Top             =   210
         Width           =   855
      End
      Begin VB.Label Label2 
         Caption         =   "To"
         Height          =   255
         Left            =   3480
         TabIndex        =   17
         Top             =   600
         Width           =   375
      End
      Begin VB.Label Label1 
         Caption         =   "Payor"
         Height          =   255
         Left            =   480
         TabIndex        =   15
         Top             =   240
         Width           =   615
      End
      Begin VB.Label Label4 
         Caption         =   "Bordx No"
         Height          =   255
         Left            =   480
         TabIndex        =   14
         Top             =   600
         Width           =   975
      End
   End
   Begin VB.Frame Frame2 
      Height          =   855
      Left            =   120
      TabIndex        =   11
      Top             =   7200
      Width           =   11655
      Begin VB.TextBox lblCurrent 
         Height          =   375
         Left            =   240
         TabIndex        =   26
         Top             =   240
         Width           =   975
      End
      Begin VB.CommandButton cmdStop 
         Caption         =   "S&top"
         Height          =   495
         Left            =   7680
         TabIndex        =   8
         Top             =   240
         Width           =   855
      End
      Begin VB.CommandButton CmdSearch 
         Caption         =   "&Search"
         Height          =   495
         Left            =   6840
         TabIndex        =   7
         Top             =   240
         Width           =   855
      End
      Begin VB.CommandButton CmdClear 
         Caption         =   "&Clear"
         Height          =   495
         Left            =   8520
         TabIndex        =   9
         Top             =   240
         Width           =   855
      End
      Begin VB.CommandButton CmdExit 
         Cancel          =   -1  'True
         Caption         =   "E&xit"
         Height          =   495
         Left            =   10200
         TabIndex        =   10
         Top             =   240
         Width           =   855
      End
      Begin VB.CommandButton cmdPrintDetails 
         Caption         =   "&Print"
         Height          =   495
         Left            =   9360
         TabIndex        =   6
         Top             =   240
         Width           =   855
      End
      Begin Crystal.CrystalReport crptSummary 
         Left            =   960
         Top             =   720
         _ExtentX        =   741
         _ExtentY        =   741
         _Version        =   348160
         PrintFileLinesPerPage=   60
      End
      Begin MSComDlg.CommonDialog dlgCommonDialog_PrintBordx 
         Left            =   1440
         Top             =   720
         _ExtentX        =   847
         _ExtentY        =   847
         _Version        =   393216
      End
      Begin VB.Label Label8 
         Alignment       =   1  'Right Justify
         BackColor       =   &H80000009&
         Height          =   255
         Left            =   2280
         TabIndex        =   22
         Top             =   600
         Visible         =   0   'False
         Width           =   1215
      End
      Begin VB.Label Label7 
         Alignment       =   1  'Right Justify
         BackColor       =   &H80000009&
         Height          =   255
         Left            =   4920
         TabIndex        =   21
         Top             =   600
         Visible         =   0   'False
         Width           =   1215
      End
      Begin VB.Label Label6 
         Alignment       =   1  'Right Justify
         BackColor       =   &H80000009&
         Height          =   255
         Left            =   3600
         TabIndex        =   20
         Top             =   600
         Visible         =   0   'False
         Width           =   1215
      End
      Begin VB.Label Label5 
         Alignment       =   1  'Right Justify
         BackColor       =   &H80000009&
         Height          =   255
         Left            =   3600
         TabIndex        =   19
         Top             =   240
         Visible         =   0   'False
         Width           =   1215
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         BackColor       =   &H80000009&
         Height          =   255
         Left            =   2280
         TabIndex        =   18
         Top             =   240
         Visible         =   0   'False
         Width           =   1215
      End
      Begin VB.Label Label21 
         Caption         =   "Records"
         Height          =   255
         Left            =   1560
         TabIndex        =   12
         Top             =   360
         Width           =   615
      End
   End
   Begin MSComctlLib.ListView lstPrintBordxExc 
      Height          =   6015
      Left            =   120
      TabIndex        =   16
      Top             =   1200
      Width           =   11655
      _ExtentX        =   20558
      _ExtentY        =   10610
      View            =   3
      LabelEdit       =   1
      Sorted          =   -1  'True
      MultiSelect     =   -1  'True
      LabelWrap       =   0   'False
      HideSelection   =   0   'False
      AllowReorder    =   -1  'True
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   11
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "No"
         Object.Width           =   882
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Bordx No"
         Object.Width           =   1764
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "Claim No"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Text            =   "Group Company"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   4
         Text            =   "Policy No"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   5
         Text            =   "Patient Name"
         Object.Width           =   6174
      EndProperty
      BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   6
         Text            =   "Act. Amt"
         Object.Width           =   2293
      EndProperty
      BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   7
         Text            =   "App. Amt"
         Object.Width           =   2293
      EndProperty
      BeginProperty ColumnHeader(9) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   8
         Text            =   "App. Comp. Amt"
         Object.Width           =   2293
      EndProperty
      BeginProperty ColumnHeader(10) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   9
         Text            =   "MBM Excess"
         Object.Width           =   2117
      EndProperty
      BeginProperty ColumnHeader(11) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   10
         Text            =   "Comp. Excess"
         Object.Width           =   2117
      EndProperty
   End
   Begin VB.Label Label28 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00404040&
      Caption         =   "Print Bordx (Group Company Excess)"
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
      TabIndex        =   25
      Top             =   0
      Width           =   11805
   End
End
Attribute VB_Name = "FrmPrintBordxEx"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim BORDXNO As String
Dim ItemChecked As Boolean
Dim listloop1 As Integer
Dim listloop2 As Integer
Dim Check As Boolean
Dim Check2 As Boolean
Public intExit As Integer
Public BordxExCriteria As String
Dim ScanCode As String
Dim SumActualAmt As String
Dim SumAppMBM As String
Dim SumAppCom As String
Dim SumMBMEx As String
Dim SumComEx As String
Dim SumActualAmt2 As String
Dim SumAppMBM2 As String
Dim SumAppCom2 As String
Dim SumMBMEx2 As String
Dim SumComEx2 As String
Dim Current2 As String
Private m_blnDirection(1 To 11) As Boolean

Private Sub ChkScan1_Click()
    
    If ChkScan1.Value = 1 Then
        TxtBordxNo1.SetFocus
        Label2.Visible = False
        TxtBordxNo2.Visible = False
    Else
        Label2.Visible = True
        TxtBordxNo2.Visible = True
        
    End If
    
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    '--- Enable PrintScreen feature
    '--- IMPORTANT - Don't forget to set .KeyPreview = True for this form at design time!!
    If Shift = 0 Then
        Select Case KeyCode
            Case vbKeyF6
                KeyCode = 0
                PrintScreenSnapShot
        End Select
    End If
End Sub


Private Sub DetailPrint(PrintMode As String)
Dim Cnt As Integer
Dim ii As Integer

Dim tmpBordxNo As String
Dim tmpBordxDate As String
Dim tmpPayor As String
Dim tmpPlan As String

Dim rsUpdate As New ADODB.Recordset
Dim strsql As String
Dim Sql As String

Dim i As Integer
Dim A As Integer
Dim B As Integer

bPrintPass = False
CancelPrinting = False
Screen.MousePointer = vbHourglass

A = 0
B = 0

For i = 1 To lstPrintBordxExc.ListItems.Count
    Set lstPrintBordx.SelectedItem = lstPrintBordx.ListItems(i)
        'PrintM = lstPrintBordx.SelectedItem.SubItems(11)
        PrintMode = lstPrintBordx.SelectedItem.SubItems(10)
        If PrintMode = "YES" Then
            A = A + 1
        Else
            B = B + 1
        End If
Next i

Select Case CboAmt.ListIndex
        Case 0
            Cnt = 0
            If PrintMode = "ALL" Then
                For ii = 1 To lstPrintBordx.ListItems.Count
                    Set lstPrintBordx.SelectedItem = lstPrintBordx.ListItems(ii)   'lvwLeaflet.ListItems(ii)
                        tmpBordxNo = lstPrintBordx.SelectedItem.SubItems(PRINTBORDX_BORDXNO)
                        tmpHLTCode = lstPrintBordx.SelectedItem.SubItems(PRINTBORDX_HLTCode)
                Next ii
            Else
                If (A = 0 And B <> 0) And cboPrintOpt.ListIndex = 1 Then
                    MsgBox "Please select 'Print New' option for this report", vbOKOnly + vbCritical, DialogBoxTitle
                Else
                    If (A <> 0 And B = 0) And cboPrintOpt.ListIndex = 0 Then
                        MsgBox "This report already printed, please select 'Reprint' ", vbOKOnly + vbCritical, DialogBoxTitle
                    Else
                        For ii = 1 To lstPrintBordx.ListItems.Count
                            Set lstPrintBordx.SelectedItem = lstPrintBordx.ListItems(ii)   'lvwLeaflet.ListItems(ii)
                                If lstPrintBordx.SelectedItem.SubItems(PRINTBORDX_PRINTSTATUS) = PrintMode Then
                                    tmpBordxNo = lstPrintBordx.SelectedItem.SubItems(PRINTBORDX_BORDXNO)
                                    tmpHLTCode = lstPrintBordx.SelectedItem.SubItems(PRINTBORDX_HLTCode)
                                End If
                        Next ii
                    End If
                End If
            End If
            
        Case 1
            Cnt = 0
            If PrintMode = "ALL" Then
                For ii = 1 To lstPrintBordx.ListItems.Count
                     Set lstPrintBordx.SelectedItem = lstPrintBordx.ListItems(ii)   'lvwLeaflet.ListItems(ii)
                        tmpBordxNo = lstPrintBordx.SelectedItem.SubItems(PRINTBORDX_BORDXNO)
                        tmpHLTCode = lstPrintBordx.SelectedItem.SubItems(PRINTBORDX_HLTCode)
                Next ii
            Else
                If (A = 0 And B <> 0) And cboPrintOpt.ListIndex = 1 Then
                    MsgBox "Please select 'Print New' option for this report", vbOKOnly + vbCritical, DialogBoxTitle
                Else
                    If (A <> 0 And B = 0) And cboPrintOpt.ListIndex = 0 Then
                        MsgBox "This report already printed, please select 'Reprint' ", vbOKOnly + vbCritical, DialogBoxTitle
                    Else
                        For ii = 1 To lstPrintBordx.ListItems.Count
                            Set lstPrintBordx.SelectedItem = lstPrintBordx.ListItems(ii)   'lvwLeaflet.ListItems(ii)
                                If lstPrintBordx.SelectedItem.SubItems(PRINTBORDX_PRINTSTATUS) = PrintMode Then
                                    tmpBordxNo = lstPrintBordx.SelectedItem.SubItems(PRINTBORDX_BORDXNO)
                                    tmpHLTCode = lstPrintBordx.SelectedItem.SubItems(PRINTBORDX_HLTCode)
                                End If
                        Next ii
                    End If
                End If
            End If
        
        Case 2
            Cnt = 0
            If PrintMode = "ALL" Then
                For ii = 1 To lstPrintBordx.ListItems.Count
                    Set lstPrintBordx.SelectedItem = lstPrintBordx.ListItems(ii)   'lvwLeaflet.ListItems(ii)
                        tmpBordxNo = lstPrintBordx.SelectedItem.SubItems(PRINTBORDX_BORDXNO)
                        tmpHLTCode = lstPrintBordx.SelectedItem.SubItems(PRINTBORDX_HLTCode)
                Next ii
            Else
                If (A = 0 And B <> 0) And cboPrintOpt.ListIndex = 1 Then
                    MsgBox "Please select 'Print New' option for this report", vbOKOnly + vbCritical, DialogBoxTitle
                Else
                    If (A <> 0 And B = 0) And cboPrintOpt.ListIndex = 0 Then
                        MsgBox "This report already printed, please select 'Reprint' ", vbOKOnly + vbCritical, DialogBoxTitle
                    Else
                        For ii = 1 To lstPrintBordx.ListItems.Count
                            Set lstPrintBordx.SelectedItem = lstPrintBordx.ListItems(ii)   'lvwLeaflet.ListItems(ii)
                                If lstPrintBordx.SelectedItem.SubItems(PRINTBORDX_PRINTSTATUS) = PrintMode Then
                                    tmpBordxNo = lstPrintBordx.SelectedItem.SubItems(PRINTBORDX_BORDXNO)
                                    tmpHLTCode = lstPrintBordx.SelectedItem.SubItems(PRINTBORDX_HLTCode)
                                End If
                        Next ii
                    End If
                End If
            End If
End Select
        
If cmdDetailPreview Then
    If Trim(tmpHLTCode) = "S" Then

        Select Case CboAmt.ListIndex
            Case 0
                crptSummary.ReportFileName = crdPath & "\Report\ClaimBordx_Sihat_Detail.rpt"
                crptSummary.Destination = crptToWindow
                crptSummary.WindowState = crptMaximized
                crptSummary.SelectionFormula = "{VIEW_ClaimBordx.BordxRefNo} = '" & Trim(txtBordxNO) & "'"
                crptSummary.SortFields(0) = "+{VIEW_ClaimBordx.INSCode}"
                crptSummary.SortFields(1) = "+{VIEW_ClaimBordx.PLNCode}"
                crptSummary.ParameterFields(0) = "txtPayor;" & CboPayor & ";true"
                crptSummary.Action = 1
    
            Case 1
                crptSummary.ReportFileName = crdPath & "\Report\ClaimBordx_Sihat_Detail_Above.rpt"
                crptSummary.Destination = crptToWindow
                crptSummary.WindowState = crptMaximized
                'crptSummary.FetchSelectionFormula = "order by {VIEW_ClaimBordx_Above.INSCode}, {VIEW_ClaimBordx_Above.PLNCode}"
                crptSummary.SelectionFormula = "{VIEW_ClaimBordx_Above.BordxRefNo} = '" & Trim(txtBordxNO) & "'"
                crptSummary.SortFields(0) = "+{VIEW_ClaimBordx_Above.INSCode}"
                crptSummary.SortFields(1) = "+{VIEW_ClaimBordx_Above.PLNCode}"
                crptSummary.ParameterFields(0) = "txtPayor;" & CboPayor & ";true"
                crptSummary.Action = 1
            
            Case 2
                crptSummary.ReportFileName = crdPath & "\Report\ClaimBordx_Sihat_Detail_All.rpt"
                crptSummary.Destination = crptToWindow
                crptSummary.WindowState = crptMaximized
                crptSummary.SelectionFormula = "{VIEW_ClaimBordx_All.BordxRefNo} = '" & Trim(txtBordxNO) & "'"
                crptSummary.SortFields(0) = "+{VIEW_ClaimBordx_All.INSCode}"
                crptSummary.SortFields(1) = "+{VIEW_ClaimBordx_All.PLNCode}"
                crptSummary.ParameterFields(0) = "txtPayor;" & CboPayor & ";true"
                crptSummary.Action = 1
        End Select
    
    Else
        If Trim(tmpHLTCode) = "N" Then
            Select Case CboAmt.ListIndex
                Case 0
                    crptSummary.ReportFileName = crdPath & "\Report\ClaimBordx_Sihat_Detail_NS.rpt"
                    crptSummary.Destination = crptToWindow
                    crptSummary.WindowState = crptMaximized
                    crptSummary.SelectionFormula = "{VIEW_ClaimBordx_NS.BordxRefNo} = '" & Trim(txtBordxNO) & "'"
                    crptSummary.SortFields(0) = "+{VIEW_ClaimBordx_NS.INSCode}"
                    crptSummary.SortFields(1) = "+{VIEW_ClaimBordx_NS.PLNCode}"
                    crptSummary.ParameterFields(0) = "txtPayor;" & CboPayor & ";true"
                    crptSummary.Action = 1
        
                Case 1
                    crptSummary.ReportFileName = crdPath & "\Report\ClaimBordx_Sihat_Detail_Above_NS.rpt"
                    crptSummary.Destination = crptToWindow
                    crptSummary.WindowState = crptMaximized
                    crptSummary.SelectionFormula = "{VIEW_ClaimBordx_Above_NS.BordxRefNo} = '" & Trim(txtBordxNO) & "'"
                    crptSummary.SortFields(0) = "+{VIEW_ClaimBordx_Above_NS.INSCode}"
                    crptSummary.SortFields(1) = "+{VIEW_ClaimBordx_Above_NS.PLNCode}"
                    crptSummary.ParameterFields(0) = "txtPayor;" & CboPayor & ";true"
                    crptSummary.Action = 1
                
                Case 2
                    crptSummary.ReportFileName = crdPath & "\Report\ClaimBordx_Sihat_Detail_All_NS.rpt"
                    crptSummary.Destination = crptToWindow
                    crptSummary.WindowState = crptMaximized
                    crptSummary.SelectionFormula = "{VIEW_ClaimBordx_All_NS.BordxRefNo} = '" & Trim(txtBordxNO) & "'"
                    crptSummary.SortFields(0) = "+{VIEW_ClaimBordx_All_NS.INSCode}"
                    crptSummary.SortFields(1) = "+{VIEW_ClaimBordx_All_NS.PLNCode}"
                    crptSummary.ParameterFields(0) = "txtPayor;" & CboPayor & ";true"
                    crptSummary.Action = 1
            End Select
        End If
    End If
    
Else
    
    If cmdPrintDetails Then
        If Trim(tmpHLTCode) = "S" Then
            Select Case CboAmt.ListIndex
                Case 0
                    crptSummary.ReportFileName = crdPath & "\Report\ClaimBordx_Sihat_Detail.rpt"
                    crptSummary.Destination = crptToPrinter
                    crptSummary.SelectionFormula = "{VIEW_ClaimBordx.BordxRefNo} = '" & Trim(txtBordxNO) & "'"
                    crptSummary.SortFields(0) = "+{VIEW_ClaimBordx.INSCode}"
                    crptSummary.SortFields(1) = "+{VIEW_ClaimBordx.PLNCode}"
                    crptSummary.ParameterFields(0) = "txtPayor;" & CboPayor & ";true"
                    
                    If bPrintPass = False And CancelPrinting = False Then
                        bPrintPass = PrintFunction
                        If bPrintPass = False Then
                            Screen.MousePointer = Default
                            Exit Sub
                        End If
                    Else
                        crptSummary.PrintReport
                    End If
                                
                    If bPrintPass = True And CancelPrinting = False Then
                        'Update file to Bordx
                        Sql = "SELECT * FROM Bordx"
                        Sql = Sql & " WHERE BordxRefNo = '" & Trim(txtBordxNO) & "'"
                                               
                        rsUpdate.Open Sql, medixmasterconnWS, adOpenKeyset, adLockOptimistic, adCmdText
                        rsUpdate!BordxPrintedStatus = "1"
                        rsUpdate.Update
                        rsUpdate.Close
                    End If
                    Call cmdSearch_Click
                
                Case 1
                    crptSummary.ReportFileName = crdPath & "\Report\ClaimBordx_Sihat_Detail_Above.rpt"
                    crptSummary.Destination = crptToPrinter
                    crptSummary.SelectionFormula = "{VIEW_ClaimBordx_Above.BordxRefNo} = '" & Trim(txtBordxNO) & "'"
                    crptSummary.SortFields(0) = "+{VIEW_ClaimBordx_Above.INSCode}"
                    crptSummary.SortFields(1) = "+{VIEW_ClaimBordx_Above.PLNCode}"
                    crptSummary.ParameterFields(0) = "txtPayor;" & CboPayor & ";true"
                    
                    If bPrintPass = False And CancelPrinting = False Then
                        bPrintPass = PrintFunction
                        If bPrintPass = False Then
                            Screen.MousePointer = Default
                            Exit Sub
                        End If
                    Else
                        crptSummary.PrintReport
                    End If
                                
                    If bPrintPass = True And CancelPrinting = False Then
                        'Update file to Bordx
                        Sql = "SELECT * FROM Bordx"
                        Sql = Sql & " WHERE BordxRefNo = '" & Trim(txtBordxNO) & "'"
                                               
                        rsUpdate.Open Sql, medixmasterconnWS, adOpenKeyset, adLockOptimistic, adCmdText
                        rsUpdate!BordxPrintedStatus = "1"
                        rsUpdate.Update
                        rsUpdate.Close
                    End If
                    Call cmdSearch_Click
                    
                Case 2
                    crptSummary.ReportFileName = crdPath & "\Report\ClaimBordx_Sihat_Detail_All.rpt"
                    crptSummary.Destination = crptToPrinter
                    crptSummary.SelectionFormula = "{VIEW_ClaimBordx_All.BordxRefNo} = '" & Trim(txtBordxNO) & "'"
                    crptSummary.SortFields(0) = "+{VIEW_ClaimBordx_All.INSCode}"
                    crptSummary.SortFields(1) = "+{VIEW_ClaimBordx_All.PLNCode}"
                    crptSummary.ParameterFields(0) = "txtPayor;" & CboPayor & ";true"
                    
                    If bPrintPass = False And CancelPrinting = False Then
                        bPrintPass = PrintFunction
                        If bPrintPass = False Then
                            Screen.MousePointer = Default
                            Exit Sub
                        End If
                    Else
                        crptSummary.PrintReport
                    End If
                                
                    If bPrintPass = True And CancelPrinting = False Then
                        'Update file to Bordx
                        Sql = "SELECT * FROM Bordx"
                        Sql = Sql & " WHERE BordxRefNo = '" & Trim(txtBordxNO) & "'"
                                               
                        rsUpdate.Open Sql, medixmasterconnWS, adOpenKeyset, adLockOptimistic, adCmdText
                        rsUpdate!BordxPrintedStatus = "1"
                        rsUpdate.Update
                        rsUpdate.Close
                    End If
                    Call cmdSearch_Click
            End Select
                
        Else
            If Trim(tmpHLTCode) = "N" Then
                Select Case CboAmt.ListIndex
                    Case 0
                        crptSummary.ReportFileName = crdPath & "\Report\ClaimBordx_Sihat_Detail_NS.rpt"
                        crptSummary.Destination = crptToPrinter
                        crptSummary.SelectionFormula = "{VIEW_ClaimBordx_NS.BordxRefNo} = '" & Trim(txtBordxNO) & "'"
                        crptSummary.SortFields(0) = "+{VIEW_ClaimBordx_NS.INSCode}"
                        crptSummary.SortFields(1) = "+{VIEW_ClaimBordx_NS.PLNCode}"
                        crptSummary.ParameterFields(0) = "txtPayor;" & CboPayor & ";true"
                        
                        If bPrintPass = False And CancelPrinting = False Then
                            bPrintPass = PrintFunction
                            If bPrintPass = False Then
                                Screen.MousePointer = Default
                                Exit Sub
                            End If
                        Else
                            crptSummary.PrintReport
                        End If
                                    
                        If bPrintPass = True And CancelPrinting = False Then
                            'Update file to Bordx
                            Sql = "SELECT * FROM Bordx"
                            Sql = Sql & " WHERE BordxRefNo = '" & Trim(txtBordxNO) & "'"
                                                   
                            rsUpdate.Open Sql, medixmasterconnWS, adOpenKeyset, adLockOptimistic, adCmdText
                            rsUpdate!BordxPrintedStatus = "1"
                            rsUpdate.Update
                            rsUpdate.Close
                        End If
                        Call cmdSearch_Click
                    
                    Case 1
                        crptSummary.ReportFileName = crdPath & "\Report\ClaimBordx_Sihat_Detail_Above_NS.rpt"
                        crptSummary.Destination = crptToPrinter
                        crptSummary.SelectionFormula = "{VIEW_ClaimBordx_Above_NS.BordxRefNo} = '" & Trim(txtBordxNO) & "'"
                        crptSummary.SortFields(0) = "+{VIEW_ClaimBordx_Above_NS.INSCode}"
                        crptSummary.SortFields(1) = "+{VIEW_ClaimBordx_Above_NS.PLNCode}"
                        crptSummary.ParameterFields(0) = "txtPayor;" & CboPayor & ";true"
                        
                        If bPrintPass = False And CancelPrinting = False Then
                            bPrintPass = PrintFunction
                            If bPrintPass = False Then
                                Screen.MousePointer = Default
                                Exit Sub
                            End If
                        Else
                            crptSummary.PrintReport
                        End If
                                    
                        If bPrintPass = True And CancelPrinting = False Then
                            'Update file to Bordx
                            Sql = "SELECT * FROM Bordx"
                            Sql = Sql & " WHERE BordxRefNo = '" & Trim(txtBordxNO) & "'"
                                                   
                            rsUpdate.Open Sql, medixmasterconnWS, adOpenKeyset, adLockOptimistic, adCmdText
                            rsUpdate!BordxPrintedStatus = "1"
                            rsUpdate.Update
                            rsUpdate.Close
                        End If
                        Call cmdSearch_Click
                        
                    Case 2
                        crptSummary.ReportFileName = crdPath & "\Report\ClaimBordx_Sihat_Detail_All_NS.rpt"
                        crptSummary.Destination = crptToPrinter
                        crptSummary.SelectionFormula = "{VIEW_ClaimBordx_All_NS.BordxRefNo} = '" & Trim(txtBordxNO) & "'"
                        crptSummary.SortFields(0) = "+{VIEW_ClaimBordx_All_NS.INSCode}"
                        crptSummary.SortFields(1) = "+{VIEW_ClaimBordx_All_NS.PLNCode}"
                        crptSummary.ParameterFields(0) = "txtPayor;" & CboPayor & ";true"
                        
                        If bPrintPass = False And CancelPrinting = False Then
                            bPrintPass = PrintFunction
                            If bPrintPass = False Then
                                Screen.MousePointer = Default
                                Exit Sub
                            End If
                        Else
                            crptSummary.PrintReport
                        End If
                                    
                        If bPrintPass = True And CancelPrinting = False Then
                            'Update file to Bordx
                            Sql = "SELECT * FROM Bordx"
                            Sql = Sql & " WHERE BordxRefNo = '" & Trim(txtBordxNO) & "'"
                                                   
                            rsUpdate.Open Sql, medixmasterconnWS, adOpenKeyset, adLockOptimistic, adCmdText
                            rsUpdate!BordxPrintedStatus = "1"
                            rsUpdate!BordxPrintedBy = Logon
                            rsUpdate!BordxPrintedDate = Date
                            rsUpdate.Update
                            rsUpdate.Close
                        End If
                        Call cmdSearch_Click
                End Select
            End If
        End If
    End If
End If
Screen.MousePointer = Default

End Sub

Private Function PrintFunction() As Boolean
On Error GoTo cancelPrint
'Select Case SSTab1.Tab
'    Case 0
        If crptSummary Is Nothing Then Exit Function
        
        With dlgCommonDialog_PrintBordx
            .DialogTitle = "Print Crystal Report "
            .CancelError = True
            .Flags = cdlPDNoSelection + cdlPDNoPageNums
            .ShowPrinter
            
            
            If Err <> MSComDlg.cdlCancel Then
            
                crptSummary.PrintReport
            End If
        End With
        PrintFunction = True
        Exit Function

cancelPrint:
        PrintFunction = False
        CancelPrinting = True
   
    
End Function
Private Sub SummaryPrint(PrintMode As String)
Dim Cnt As Integer
Dim ii As Integer

Dim tmpBordxNo As String
Dim tmpBordxDate As String
Dim tmpPayor As String
Dim tmpPlan As String

Dim rsUpdate As New ADODB.Recordset
Dim strsql As String
Dim Sql As String

Dim i As Integer
Dim A As Integer
Dim B As Integer

bPrintPass = False
CancelPrinting = False
Screen.MousePointer = vbHourglass

A = 0
B = 0

For i = 1 To lstPrintBordx.ListItems.Count
    Set lstPrintBordx.SelectedItem = lstPrintBordx.ListItems(i)   'lvwMCO.ListItems(ii)
        PrintMode = lstPrintBordx.SelectedItem.SubItems(10)
        If PrintMode = "YES" Then
            A = A + 1
        Else
            B = B + 1
        End If
Next i

Select Case CboAmt.ListIndex
        Case 0
            Cnt = 0
            If PrintMode = "ALL" Then
                For ii = 1 To lstPrintBordx.ListItems.Count
                    Set lstPrintBordx.SelectedItem = lstPrintBordx.ListItems(ii)   'lvwLeaflet.ListItems(ii)
                        tmpBordxNo = lstPrintBordx.SelectedItem.SubItems(PRINTBORDX_BORDXNO)
                        tmpPayor = lstPrintBordx.SelectedItem.SubItems(PRINTBORDX_PAYOR)
                        tmpHLTCode = lstPrintBordx.SelectedItem.SubItems(PRINTBORDX_HLTCode)
                Next ii
            Else
                If (A = 0 And B <> 0) And cboPrintOpt.ListIndex = 1 Then
                    MsgBox "Please select 'Print New' option for this report", vbOKOnly + vbCritical, DialogBoxTitle
                Else
                    If (A <> 0 And B = 0) And cboPrintOpt.ListIndex = 0 Then
                        MsgBox "This report already printed, please select 'Reprint' ", vbOKOnly + vbCritical, DialogBoxTitle
                    Else
                        For ii = 1 To lstPrintBordx.ListItems.Count
                            Set lstPrintBordx.SelectedItem = lstPrintBordx.ListItems(ii)   'lvwLeaflet.ListItems(ii)
                                If lstPrintBordx.SelectedItem.SubItems(PRINTBORDX_PRINTSTATUS) = PrintMode Then
                                    tmpBordxNo = lstPrintBordx.SelectedItem.SubItems(PRINTBORDX_BORDXNO)
                                    'tmpPayor = lstPrintBordx.SelectedItem.SubItems(PRINTBORDX_PAYOR)
                                    tmpHLTCode = lstPrintBordx.SelectedItem.SubItems(PRINTBORDX_HLTCode)
                                End If
                        Next ii
                    End If
                End If
            End If
            
        Case 1
            Cnt = 0
            If PrintMode = "ALL" Then
                For ii = 1 To lstPrintBordx.ListItems.Count
                     Set lstPrintBordx.SelectedItem = lstPrintBordx.ListItems(ii)   'lvwLeaflet.ListItems(ii)
                        tmpBordxNo = lstPrintBordx.SelectedItem.SubItems(PRINTBORDX_BORDXNO)
                        'tmpPayor = lstPrintBordx.SelectedItem.SubItems(PRINTBORDX_PAYOR)
                        tmpHLTCode = lstPrintBordx.SelectedItem.SubItems(PRINTBORDX_HLTCode)
                Next ii
            Else
                If (A = 0 And B <> 0) And cboPrintOpt.ListIndex = 1 Then
                    MsgBox "Please select 'Print New' option for this report", vbOKOnly + vbCritical, DialogBoxTitle
                Else
                    If (A <> 0 And B = 0) And cboPrintOpt.ListIndex = 0 Then
                        MsgBox "This report already printed, please select 'Reprint' ", vbOKOnly + vbCritical, DialogBoxTitle
                    Else
                        For ii = 1 To lstPrintBordx.ListItems.Count
                            Set lstPrintBordx.SelectedItem = lstPrintBordx.ListItems(ii)   'lvwLeaflet.ListItems(ii)
                                If lstPrintBordx.SelectedItem.SubItems(PRINTBORDX_PRINTSTATUS) = PrintMode Then
                                    tmpBordxNo = lstPrintBordx.SelectedItem.SubItems(PRINTBORDX_BORDXNO)
                                    'tmpPayor = lstPrintBordx.SelectedItem.SubItems(PRINTBORDX_PAYOR)
                                    tmpHLTCode = lstPrintBordx.SelectedItem.SubItems(PRINTBORDX_HLTCode)
                                End If
                        Next ii
                    End If
                End If
            End If
        
        Case 2
            Cnt = 0
            If PrintMode = "ALL" Then
                For ii = 1 To lstPrintBordx.ListItems.Count
                    Set lstPrintBordx.SelectedItem = lstPrintBordx.ListItems(ii)   'lvwLeaflet.ListItems(ii)
                        tmpBordxNo = lstPrintBordx.SelectedItem.SubItems(PRINTBORDX_BORDXNO)
                        'tmpPayor = lstPrintBordx.SelectedItem.SubItems(PRINTBORDX_PAYOR)
                        tmpHLTCode = lstPrintBordx.SelectedItem.SubItems(PRINTBORDX_HLTCode)
                Next ii
            Else
                If (A = 0 And B <> 0) And cboPrintOpt.ListIndex = 1 Then
                    MsgBox "Please select 'Print New' option for this report", vbOKOnly + vbCritical, DialogBoxTitle
                Else
                    If (A <> 0 And B = 0) And cboPrintOpt.ListIndex = 0 Then
                        MsgBox "This report already printed, please select 'Reprint' ", vbOKOnly + vbCritical, DialogBoxTitle
                    Else
                        For ii = 1 To lstPrintBordx.ListItems.Count
                            Set lstPrintBordx.SelectedItem = lstPrintBordx.ListItems(ii)   'lvwLeaflet.ListItems(ii)
                                If lstPrintBordx.SelectedItem.SubItems(PRINTBORDX_PRINTSTATUS) = PrintMode Then
                                    tmpBordxNo = lstPrintBordx.SelectedItem.SubItems(PRINTBORDX_BORDXNO)
                                    'tmpPayor = lstPrintBordx.SelectedItem.SubItems(PRINTBORDX_PAYOR)
                                    tmpHLTCode = lstPrintBordx.SelectedItem.SubItems(PRINTBORDX_HLTCode)
                                End If
                        Next ii
                    End If
                End If
            End If
End Select
        
If CmdSummaryPreview Then
    If Trim(tmpHLTCode) = "S" Then
        Select Case CboAmt.ListIndex
            Case 0
                crptSummary.ReportFileName = crdPath & "\Report\ClaimBordx_Sihat_Below.rpt"
                crptSummary.Destination = crptToWindow
                crptSummary.WindowState = crptMaximized
                crptSummary.SelectionFormula = "{VIEW_ClaimBordx_Analysis.BordxRefNo} = '" & Trim(txtBordxNO) & "'"
                                
                crptSummary.ParameterFields(0) = "txtPayor;" & CboPayor & ";true"
                crptSummary.Action = 1
    
            Case 1
                crptSummary.ReportFileName = crdPath & "\Report\ClaimBordx_Sihat_Above.rpt"
                crptSummary.Destination = crptToWindow
                crptSummary.WindowState = crptMaximized
                crptSummary.SelectionFormula = "{VIEW_ClaimBordx_Analysis_Above.BordxRefNo} = '" & Trim(txtBordxNO) & "'"
                                
                crptSummary.ParameterFields(0) = "txtPayor;" & CboPayor & ";true"
                crptSummary.Action = 1
            
            Case 2
                crptSummary.ReportFileName = crdPath & "\Report\ClaimBordx_Sihat_All.rpt"
                crptSummary.Destination = crptToWindow
                crptSummary.WindowState = crptMaximized
                crptSummary.SelectionFormula = "{VIEW_ClaimBOrdx_Analysis_All.BordxRefNo} = '" & Trim(txtBordxNO) & "'"
                                
                crptSummary.ParameterFields(0) = "txtPayor;" & CboPayor & ";true"
                crptSummary.Action = 1
        End Select
    Else
        If Trim(tmpHLTCode) = "N" Then
            Select Case CboAmt.ListIndex
                Case 0
                    crptSummary.ReportFileName = crdPath & "\Report\ClaimBordx_NSihat_Below.rpt"
                    crptSummary.Destination = crptToWindow
                    crptSummary.WindowState = crptMaximized
                    crptSummary.SelectionFormula = "{VIEW_ClaimBOrdx_NS_Analysis_Below.BordxRefNo} = '" & Trim(txtBordxNO) & "'"
                                    
                    crptSummary.ParameterFields(0) = "txtPayor;" & CboPayor & ";true"
                    crptSummary.Action = 1
        
                Case 1
                    crptSummary.ReportFileName = crdPath & "\Report\ClaimBordx_NSihat_Above.rpt"
                    crptSummary.Destination = crptToWindow
                    crptSummary.WindowState = crptMaximized
                    crptSummary.SelectionFormula = "{VIEW_ClaimBOrdx_NS_Analysis_Above.BordxRefNo} = '" & Trim(txtBordxNO) & "'"
                                    
                    crptSummary.ParameterFields(0) = "txtPayor;" & CboPayor & ";true"
                    crptSummary.Action = 1
                
                Case 2
                    crptSummary.ReportFileName = crdPath & "\Report\ClaimBordx_NSihat_All.rpt"
                    crptSummary.Destination = crptToWindow
                    crptSummary.WindowState = crptMaximized
                    crptSummary.SelectionFormula = "{VIEW_ClaimBOrdx_NS_Analysis_All.BordxRefNo} = '" & Trim(txtBordxNO) & "'"
                                    
                    crptSummary.ParameterFields(0) = "txtPayor;" & CboPayor & ";true"
                    crptSummary.Action = 1
            End Select
           
        End If
    End If
Else
    If CmdPrintSummary Then
        If Trim(tmpHLTCode) = "S" Then
            Select Case CboAmt.ListIndex
                Case 0
                    crptSummary.ReportFileName = crdPath & "\Report\ClaimBordx_Sihat_Below.rpt"
                    crptSummary.Destination = crptToPrinter
                    crptSummary.SelectionFormula = "{VIEW_ClaimBordx_Analysis.BordxRefNo} = '" & Trim(txtBordxNO) & "'"
                    crptSummary.ParameterFields(0) = "txtPayor;" & CboPayor & ";true"
                    
                    If bPrintPass = False And CancelPrinting = False Then
                        bPrintPass = PrintFunction
                        If bPrintPass = False Then
                            Screen.MousePointer = Default
                            Exit Sub
                        End If
                    Else
                        crptSummary.PrintReport
                    End If
                                
                    If bPrintPass = True And CancelPrinting = False Then
                        'Update file to Bordx
                        Sql = "SELECT * FROM Bordx"
                        Sql = Sql & " WHERE BordxRefNo = '" & Trim(txtBordxNO) & "'"
                                               
                        rsUpdate.Open Sql, medixmasterconnWS, adOpenKeyset, adLockOptimistic, adCmdText
                        rsUpdate!BordxPrintedStatus = "1"
                        rsUpdate.Update
                        rsUpdate.Close
                    End If
                    Call cmdSearch_Click
                
                Case 1
                    crptSummary.ReportFileName = crdPath & "\Report\ClaimBordx_Sihat_Above.rpt"
                    crptSummary.Destination = crptToPrinter
                    crptSummary.SelectionFormula = "{VIEW_ClaimBordx_Analysis_Above.BordxRefNo} = '" & Trim(txtBordxNO) & "'"
                    crptSummary.ParameterFields(0) = "txtPayor;" & CboPayor & ";true"
                    
                    If bPrintPass = False And CancelPrinting = False Then
                        bPrintPass = PrintFunction
                        If bPrintPass = False Then
                            Screen.MousePointer = Default
                            Exit Sub
                        End If
                    Else
                        crptSummary.PrintReport
                    End If
                                
                    If bPrintPass = True And CancelPrinting = False Then
                        'Update file to Bordx
                        Sql = "SELECT * FROM Bordx"
                        Sql = Sql & " WHERE BordxRefNo = '" & Trim(txtBordxNO) & "'"
                                               
                        rsUpdate.Open Sql, medixmasterconnWS, adOpenKeyset, adLockOptimistic, adCmdText
                        rsUpdate!BordxPrintedStatus = "1"
                        rsUpdate.Update
                        rsUpdate.Close
                    End If
                    Call cmdSearch_Click
                
                Case 2
                    crptSummary.ReportFileName = crdPath & "\Report\ClaimBordx_Sihat_All.rpt"
                    crptSummary.Destination = crptToPrinter
                    crptSummary.SelectionFormula = "{VIEW_ClaimBordx_Analysis_All.BordxRefNo} = '" & Trim(txtBordxNO) & "'"
                    crptSummary.ParameterFields(0) = "txtPayor;" & CboPayor & ";true"
                    
                    If bPrintPass = False And CancelPrinting = False Then
                        bPrintPass = PrintFunction
                        If bPrintPass = False Then
                            Screen.MousePointer = Default
                            Exit Sub
                        End If
                    Else
                        crptSummary.PrintReport
                    End If
                                
                    If bPrintPass = True And CancelPrinting = False Then
                        'Update file to Bordx
                        Sql = "SELECT * FROM Bordx"
                        Sql = Sql & " WHERE BordxRefNo = '" & Trim(txtBordxNO) & "'"
                                               
                        rsUpdate.Open Sql, medixmasterconnWS, adOpenKeyset, adLockOptimistic, adCmdText
                        rsUpdate!BordxPrintedStatus = "1"
                        rsUpdate.Update
                        rsUpdate.Close
                    End If
                    Call cmdSearch_Click
            End Select
        
        Else
            If Trim(tmpHLTCode) = "N" Then
                Select Case CboAmt.ListIndex
                    Case 0
                        crptSummary.ReportFileName = crdPath & "\Report\ClaimBordx_NSihat_Below.rpt"
                        crptSummary.Destination = crptToPrinter
                        crptSummary.SelectionFormula = "{VIEW_ClaimBOrdx_NS_Analysis_Below.BordxRefNo} = '" & Trim(txtBordxNO) & "'"
                        crptSummary.ParameterFields(0) = "txtPayor;" & CboPayor & ";true"
                        
                        If bPrintPass = False And CancelPrinting = False Then
                            bPrintPass = PrintFunction
                            If bPrintPass = False Then
                                Screen.MousePointer = Default
                                Exit Sub
                            End If
                        Else
                            crptSummary.PrintReport
                        End If
                                    
                        If bPrintPass = True And CancelPrinting = False Then
                            'Update file to Bordx
                            Sql = "SELECT * FROM Bordx"
                            Sql = Sql & " WHERE BordxRefNo = '" & Trim(txtBordxNO) & "'"
                                                   
                            rsUpdate.Open Sql, medixmasterconnWS, adOpenKeyset, adLockOptimistic, adCmdText
                            rsUpdate!BordxPrintedStatus = "1"
                            rsUpdate.Update
                            rsUpdate.Close
                        End If
                        Call cmdSearch_Click
                    
                    Case 1
                        crptSummary.ReportFileName = crdPath & "\Report\ClaimBordx_NSihat_Above.rpt"
                        crptSummary.Destination = crptToPrinter
                        crptSummary.SelectionFormula = "{VIEW_ClaimBOrdx_NS_Analysis_Above.BordxRefNo} = '" & Trim(txtBordxNO) & "'"
                        crptSummary.ParameterFields(0) = "txtPayor;" & CboPayor & ";true"
                        
                        If bPrintPass = False And CancelPrinting = False Then
                            bPrintPass = PrintFunction
                            If bPrintPass = False Then
                                Screen.MousePointer = Default
                                Exit Sub
                            End If
                        Else
                            crptSummary.PrintReport
                        End If
                                    
                        If bPrintPass = True And CancelPrinting = False Then
                            'Update file to Bordx
                            Sql = "SELECT * FROM Bordx"
                            Sql = Sql & " WHERE BordxRefNo = '" & Trim(txtBordxNO) & "'"
                                                   
                            rsUpdate.Open Sql, medixmasterconnWS, adOpenKeyset, adLockOptimistic, adCmdText
                            rsUpdate!BordxPrintedStatus = "1"
                            rsUpdate.Update
                            rsUpdate.Close
                        End If
                        Call cmdSearch_Click
                    
                    Case 2
                        crptSummary.ReportFileName = crdPath & "\Report\ClaimBordx_NSihat_All.rpt"
                        crptSummary.Destination = crptToPrinter
                        crptSummary.SelectionFormula = "{VIEW_ClaimBOrdx_NS_Analysis_All.BordxRefNo} = '" & Trim(txtBordxNO) & "'"
                        crptSummary.ParameterFields(0) = "txtPayor;" & CboPayor & ";true"
                        
                        If bPrintPass = False And CancelPrinting = False Then
                            bPrintPass = PrintFunction
                            If bPrintPass = False Then
                                Screen.MousePointer = Default
                                Exit Sub
                            End If
                        Else
                            crptSummary.PrintReport
                        End If
                                    
                        If bPrintPass = True And CancelPrinting = False Then
                            'Update file to Bordx
                            Sql = "SELECT * FROM Bordx"
                            Sql = Sql & " WHERE BordxRefNo = '" & Trim(txtBordxNO) & "'"
                                                   
                            rsUpdate.Open Sql, medixmasterconnWS, adOpenKeyset, adLockOptimistic, adCmdText
                            rsUpdate!BordxPrintedStatus = "1"
                            rsUpdate.Update
                            rsUpdate.Close
                        End If
                        Call cmdSearch_Click
                End Select
            End If
        End If
    End If
End If
Screen.MousePointer = Default

End Sub

Private Sub cmdPrintDetails_Click()
        
    Screen.MousePointer = vbHourglass
        Call ExportBordxExcess(lstPrintBordxExc)
    Screen.MousePointer = Default
    
    BORDXNO = ""
End Sub

Private Sub Form_Load()
    intExit = 0
    lblCurrent = 0
    SumActualAmt2 = 0
    SumAppMBM2 = 0
    SumAppCom2 = 0
    SumMBMEx2 = 0
    SumComEx2 = 0
    Current2 = 0
    Label3 = 0
    Label5 = 0
    Label6 = 0
    Label7 = 0
    Label8 = 0
    Label2.Visible = False
    TxtBordxNo2.Visible = False
   ' medixmasterconn.Open "File Name=" & BOSPath & "Conn\mediconnect.UDL"
    Call ComboListing
  '  medixmasterconn.Close
  '  Set medixmasterconn = Nothing
End Sub

'******************** Start Command Button ********************************
Private Sub CmdClear_Click()
    lstPrintBordxExc.ListItems.Clear
    CboPayor.Clear
    TxtBordxNo1.Text = ""
    TxtBordxNo2.Text = ""
    lblCurrent = 0
    Label3 = 0
    Label5 = 0
    Label6 = 0
    Label7 = 0
    Label8 = 0
    SumActualAmt = 0
    SumAppMBM = 0
    SumAppCom = 0
    SumMBMEx = 0
    SumComEx = 0
    SumActualAmt2 = 0
    SumAppMBM2 = 0
    SumAppCom2 = 0
    SumMBMEx2 = 0
    SumComEx2 = 0
    Current2 = 0
    Check = False
    Check2 = False
    BORDXNO = ""
    Call ComboListing
End Sub

Public Sub cmdSearch_Click()

Label3 = ""
Label5 = ""
Label6 = ""
Label7 = ""
Label8 = ""

    Screen.MousePointer = vbHourglass
    CmdSearch.Enabled = False
  '  medixmasterconnWS.Open "File Name=" & BOSPath & "Conn\mediconnectWS.UDL"
    
    If ChkScan1.Value = 1 Then
        If TxtBordxNo1 = "" Then
            MsgBox "Please enter Bordx Reference No", vbOKOnly + vbCritical, DialogBoxTitle
            TxtBordxNo1.SetFocus
        Else
            ScanCode = TxtBordxNo1
            If Mid(ScanCode, 1, 2) = "LO" Then
                TxtBordxNo1 = ""
                Call SearchForm(ScanCode)
            Else
                If CboPayor = "" Then
                    MsgBox "Please select Payor Code", vbOKOnly + vbCritical, DialogBoxTitle
                    CboPayor.SetFocus
                Else
                                                                       
                listrecord = lstPrintBordxExc.ListItems.Count
                    If listrecord <> 0 Then
            
                        For listloop1 = 1 To listrecord
                            Call BordxNoCheck
                        Next listloop1
                    End If
                    If Check = True Then
                        MsgBox "Data has been uploaded!", vbOKOnly + vbCritical
                    ElseIf Check = False Then
                        CmdClear.Enabled = False
                        CmdExit.Enabled = False
                        cmdPrintDetails.Enabled = False
                        Call SearchRecord
                    End If
                    Check = False
                    TxtBordxNo1 = ""
                    TxtBordxNo1.SetFocus
                End If
            End If
        End If
    
    Else
        CmdClear.Enabled = False
        CmdExit.Enabled = False
        cmdPrintDetails.Enabled = False
        Call SearchRecord
    
    End If
   ' medixmasterconnWS.Close
  '  Set medixmasterconnWS = Nothing
    CmdSearch.Enabled = True
    Screen.MousePointer = Default
End Sub

Private Sub CmdStop_Click()
    intExit = 1
    CmdClear.Enabled = True
    CmdExit.Enabled = True
    cmdPrintDetails.Enabled = True
End Sub

Private Sub CmdExit_Click()
    Unload Me
End Sub

Private Sub ComboListing()
Dim PAY As New ADODB.Recordset
Dim cmd As New ADODB.Command

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
            CboPayor.AddItem !PAYCode & " - " & !PAYCompanyName
            PAY.MoveNext
        End With
    Loop
    PAY.Close
    Set PAY = Nothing
    
End Sub

Private Function SearchRecord()
Dim strsql As String
Dim strAppend As String
Dim Sql As String

    
    strAppend = ""
    
    If Len(Trim(CboPayor)) Then
        strAppend = strAppend + " And PAYCode = '" & Trim(Mid(CboPayor, 1, IIf(InStr(1, CboPayor, "-") = 0, 4, InStr(1, CboPayor, "-") - 1))) & "'"
    End If
    
    If Len(Trim(TxtPlan1)) Then
        strAppend = strAppend + " And PLNCode >= '" & ChkStr(TxtPlan1) & "'"
    End If
    
    If Len(Trim(TxtPlan2)) Then
        strAppend = strAppend + " And PLNCode <= '" & ChkStr(TxtPlan2) & "'"
    End If
    
    If ChkScan1.Value = 1 Then
        If Len(Trim(TxtBordxNo1)) Then
            strAppend = strAppend + " And BordxRefNo = '" & ChkStr(TxtBordxNo1) & "'"
            BORDXNO = Trim(TxtBordxNo1)
        End If
    Else
        If Len(Trim(TxtBordxNo1)) Then
            strAppend = strAppend + " And BordxRefNo >= '" & ChkStr(TxtBordxNo1) & "'"
        End If
        
        If Len(Trim(TxtBordxNo2)) Then
            strAppend = strAppend + " And BordxRefNo <= '" & ChkStr(TxtBordxNo2) & "'"
        End If
    End If
            
    BordxExCriteria = strAppend
    
    If ChkScan1.Value = 1 Then
        Call PrintBordxExListing2(strAppend)
    Else
        Call PrintBordxExListing(strAppend)
    End If

End Function

Private Sub PrintBordxExListing(Optional strAppend As String)
Dim rst2 As New ADODB.Recordset
Dim rstCount As New ADODB.Recordset
Dim Record As ListItem
Dim Sql As String
Dim Current As String
Dim intloop As Integer
    
    Current = 0
    SumActualAmt = 0
    SumAppMBM = 0
    SumAppCom = 0
    SumMBMEx = 0
    SumComEx = 0
    
    lstPrintBordxExc.ListItems.Clear
               
    Sql = " Select CASNumber, ClaimVersion, MBMExcess, CLMCompExcess, " & _
          " PatientName, AppMBM, CLMTotalCompApp, ActualAmt, " & _
          " PAYCode, BordxRefNo, MBMPolicyNo, MBMGroupCompany From View_PrintBordxExcess Where 0 = 0 "
    
    If Len(Trim(strAppend)) Then
        Sql = Sql + strAppend
    End If
           
    rst2.Open Sql, medixmasterconnWS, adOpenForwardOnly, adLockBatchOptimistic
   
    Do
        Do Until rst2.EOF
        
             With rst2
             
                intloop = intloop + 1
            
                Set Record = lstPrintBordxExc.ListItems.Add(, , intloop)
                
                If Len(!BordxRefNo) Then
                    Record.SubItems(1) = Trim(!BordxRefNo)
                End If
                
                If Len(!CasNumber & "-" & !claimversion) Then
                    Record.SubItems(2) = (!CasNumber & "-" & !claimversion)
                End If
                
                If Len(!MBMGroupCompany) Then
                    Record.SubItems(3) = Trim(!MBMGroupCompany)
                End If
                
                If Len(!MBMPolicyNo) Then
                    Record.SubItems(4) = Trim(!MBMPolicyNo)
                End If
                
                If Len(!PatientName) Then
                    Record.SubItems(5) = Trim(!PatientName)
                End If
                
                Record.SubItems(6) = IIf(IsNumeric(!ActualAmt) = True, Format(!ActualAmt, "#,##0.00;(#,##0.00)"), 0)
                SumActualAmt = CDbl(SumActualAmt) + Format(Record.SubItems(6), "#,##0.00;(#,##0.00)")
                Label3 = Format(SumActualAmt, "#,##0.00;(#,##0.00)")
                
                Record.SubItems(7) = IIf(IsNumeric(!AppMBM) = True, Format(!AppMBM, "#,##0.00;(#,##0.00)"), 0)
                SumAppMBM = CDbl(SumAppMBM) + Format(Record.SubItems(7), "#,##0.00;(#,##0.00)")
                Label8 = Format(SumAppMBM, "#,##0.00;(#,##0.00)")
                    
                Record.SubItems(8) = IIf(IsNumeric(!CLMTotalCompApp) = True, Format(!CLMTotalCompApp, "#,##0.00;(#,##0.00)"), 0)
                SumAppCom = CDbl(SumAppCom) + Format(Record.SubItems(8), "#,##0.00;(#,##0.00)")
                Label5 = Format(SumAppCom, "#,##0.00;(#,##0.00)")
                    
                Record.SubItems(9) = IIf(IsNumeric(!MBMExcess) = True, Format(!MBMExcess, "#,##0.00;(#,##0.00)"), 0)
                SumMBMEx = CDbl(SumMBMEx) + Format(Record.SubItems(9), "#,##0.00;(#,##0.00)")
                Label6 = Format(SumMBMEx, "#,##0.00;(#,##0.00)")
                
                Record.SubItems(10) = IIf(IsNumeric(!CLMCompExcess) = True, Format(!CLMCompExcess, "#,##0.00;(#,##0.00)"), 0)
                SumComEx = CDbl(SumComEx) + Format(Record.SubItems(10), "#,##0.00;(#,##0.00)")
                Label7 = Format(SumComEx, "#,##0.00;(#,##0.00)")
                
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
    
    If lstPrintBordxExc.ListItems.Count = 0 Then
        MsgBox "No record found!", vbOKOnly + vbCritical, DialogBoxTitle
        cmdPrintDetails.Enabled = True
        CmdClear.Enabled = True
        CmdExit.Enabled = True
        lblCurrent = 0
    End If
    
    intExit = 0
    rst2.Close
    Set rst2 = Nothing
    CmdClear.Enabled = True
    CmdExit.Enabled = True
    cmdPrintDetails.Enabled = True
    
End Sub

Private Sub PrintBordxExListing2(Optional strAppend As String)
Dim rst2 As New ADODB.Recordset
Dim rstCount As New ADODB.Recordset
Dim Record As ListItem
Dim Sql As String
Dim intloop As Integer
    
               
    Sql = " Select CASNumber, ClaimVersion, MBMExcess, CLMCompExcess, " & _
          " PatientName, AppMBM, CLMTotalCompApp, ActualAmt, " & _
          " PAYCode, BordxRefNo, MBMPolicyNo, MBMGroupCompany From View_PrintBordxExcess where 0=0 "
    
    If Len(Trim(strAppend)) Then
        Sql = Sql + strAppend
    End If
           
    rst2.Open Sql, medixmasterconnWS, adOpenForwardOnly, adLockBatchOptimistic
   
    Do
        Do Until rst2.EOF
        
             With rst2
                
                intloop = intloop + 1
            
                Set Record = lstPrintBordxExc.ListItems.Add(, , intloop)
                
                If Len(!BordxRefNo) Then
                    Record.SubItems(1) = Trim(!BordxRefNo)
                End If
                
                If Len(!CasNumber & "-" & !claimversion) Then
                    Record.SubItems(2) = (!CasNumber & "-" & !claimversion)
                End If
                
                If Len(!MBMGroupCompany) Then
                    Record.SubItems(3) = Trim(!MBMGroupCompany)
                End If
                
                If Len(!MBMPolicyNo) Then
                    Record.SubItems(4) = Trim(!MBMPolicyNo)
                End If
                
                If Len(!PatientName) Then
                    Record.SubItems(5) = Trim(!PatientName)
                End If
                
                Record.SubItems(6) = IIf(IsNumeric(!ActualAmt) = True, Format(!ActualAmt, "#,##0.00;(#,##0.00)"), 0)
                SumActualAmt2 = CDbl(SumActualAmt2) + Format(Record.SubItems(6), "#,##0.00;(#,##0.00)")
                Label3 = Format(SumActualAmt2, "#,##0.00;(#,##0.00)")
                
                Record.SubItems(7) = IIf(IsNumeric(!AppMBM) = True, Format(!AppMBM, "#,##0.00;(#,##0.00)"), 0)
                SumAppMBM2 = CDbl(SumAppMBM2) + Format(Record.SubItems(7), "#,##0.00;(#,##0.00)")
                Label8 = Format(SumAppMBM2, "#,##0.00;(#,##0.00)")
                    
                Record.SubItems(8) = IIf(IsNumeric(!CLMTotalCompApp) = True, Format(!CLMTotalCompApp, "#,##0.00;(#,##0.00)"), 0)
                SumAppCom2 = CDbl(SumAppCom2) + Format(Record.SubItems(8), "#,##0.00;(#,##0.00)")
                Label5 = Format(SumAppCom2, "#,##0.00;(#,##0.00)")
                    
                Record.SubItems(9) = IIf(IsNumeric(!MBMExcess) = True, Format(!MBMExcess, "#,##0.00;(#,##0.00)"), 0)
                SumMBMEx2 = CDbl(SumMBMEx2) + Format(Record.SubItems(9), "#,##0.00;(#,##0.00)")
                Label6 = Format(SumMBMEx2, "#,##0.00;(#,##0.00)")
                
                Record.SubItems(10) = IIf(IsNumeric(!CLMCompExcess) = True, Format(!CLMCompExcess, "#,##0.00;(#,##0.00)"), 0)
                SumComEx2 = CDbl(SumComEx2) + Format(Record.SubItems(10), "#,##0.00;(#,##0.00)")
                Label7 = Format(SumComEx2, "#,##0.00;(#,##0.00)")
                
                Current2 = Current2 + 1
                lblCurrent = Current2
                
            End With
            rst2.MoveNext
            DoEvents
                If intExit = 1 Then
                   Check = False
                   Exit Do
                End If
        Loop
    Loop Until Check = False
    
    If lstPrintBordxExc.ListItems.Count = 0 Then
        MsgBox "No record found!", vbOKOnly + vbCritical, DialogBoxTitle
        cmdPrintDetails.Enabled = True
        CmdClear.Enabled = True
        CmdExit.Enabled = True
        lblCurrent = 0
    End If
    
    intExit = 0
    rst2.Close
    Set rst2 = Nothing
    CmdClear.Enabled = True
    CmdExit.Enabled = True
    cmdPrintDetails.Enabled = True
End Sub

Private Sub lstPrintBordxExc_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
    m_blnDirection(ColumnHeader.Index) = Not m_blnDirection(ColumnHeader.Index)
    Select Case ColumnHeader.Index
    Case 1
        SortListView lstPrintBordxExc, ColumnHeader.Index, ldtNumber, m_blnDirection(1)
    Case 2
        SortListView lstPrintBordxExc, ColumnHeader.Index, ldtString, m_blnDirection(2)
    Case 3
        SortListView lstPrintBordxExc, ColumnHeader.Index, ldtString, m_blnDirection(3)
    Case 4
        SortListView lstPrintBordxExc, ColumnHeader.Index, ldtString, m_blnDirection(4)
    Case 5
        SortListView lstPrintBordxExc, ColumnHeader.Index, ldtString, m_blnDirection(5)
    Case 6
        SortListView lstPrintBordxExc, ColumnHeader.Index, ldtString, m_blnDirection(6)
    Case 7
        SortListView lstPrintBordxExc, ColumnHeader.Index, ldtNumber, m_blnDirection(7)
    Case 8
        SortListView lstPrintBordxExc, ColumnHeader.Index, ldtNumber, m_blnDirection(8)
    Case 9
        SortListView lstPrintBordxExc, ColumnHeader.Index, ldtNumber, m_blnDirection(9)
    Case 10
        SortListView lstPrintBordxExc, ColumnHeader.Index, ldtNumber, m_blnDirection(10)
    Case 11
        SortListView lstPrintBordxExc, ColumnHeader.Index, ldtNumber, m_blnDirection(11)
    End Select
End Sub

Private Sub CboPayor_Click()
    If CboPayor.Text = "" Then
        If InStr(CboPayor.Text, "null") Then
            CboPayor.Enabled = False
        End If
        TxtBordxNo1.SetFocus
    End If
    TxtBordxNo1.SetFocus
End Sub

Private Sub TxtBordxNo1_GotFocus()
    CmdSearch.Default = True
End Sub

Private Sub TxtBordxNo1_LostFocus()
    TxtBordxNo1 = ChkStr(TxtBordxNo1)
    CmdSearch.Default = False
End Sub

Private Function BordxNoCheck()
    If Trim(ChkStr(TxtBordxNo1)) = ChkStr(Trim(lstPrintBordxExc.ListItems(listloop1).Text)) Then
        Check = True
    End If
End Function

