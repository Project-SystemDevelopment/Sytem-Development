VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form FrmReceipt 
   Caption         =   "Receipt from MNRB/PAYOR"
   ClientHeight    =   8595
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   11880
   LinkTopic       =   "Form2"
   MDIChild        =   -1  'True
   ScaleHeight     =   8595
   ScaleWidth      =   11880
   WindowState     =   2  'Maximized
   Begin MSComctlLib.ListView lstWorksheet 
      Height          =   4680
      Left            =   120
      TabIndex        =   42
      Top             =   2040
      Width           =   11415
      _ExtentX        =   20135
      _ExtentY        =   8255
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
      NumItems        =   5
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Check"
         Object.Width           =   1764
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Payor"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "Bordx No"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Text            =   "Worksheet No"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   4
         Text            =   "Amount"
         Object.Width           =   2540
      EndProperty
   End
   Begin VB.Frame Frame1 
      Height          =   1620
      Left            =   120
      TabIndex        =   25
      Top             =   360
      Width           =   11415
      Begin VB.OptionButton OptBordx 
         Caption         =   "By Bordx"
         Height          =   255
         Left            =   1680
         TabIndex        =   44
         Top             =   200
         Value           =   -1  'True
         Width           =   1335
      End
      Begin VB.OptionButton OptWks 
         Caption         =   "By Worksheet"
         Height          =   255
         Left            =   3360
         TabIndex        =   43
         Top             =   200
         Width           =   1335
      End
      Begin VB.TextBox TxtBordx1 
         Height          =   285
         Left            =   1680
         MaxLength       =   20
         TabIndex        =   36
         Top             =   825
         Width           =   2295
      End
      Begin VB.ComboBox CboGRPCom 
         Height          =   315
         Left            =   1680
         TabIndex        =   32
         Top             =   1125
         Width           =   5295
      End
      Begin VB.ComboBox CboPayor 
         Height          =   315
         Left            =   1680
         TabIndex        =   31
         Top             =   495
         Width           =   5295
      End
      Begin VB.TextBox TxtBordx2 
         Height          =   285
         Left            =   4560
         MaxLength       =   20
         TabIndex        =   30
         Top             =   825
         Width           =   2415
      End
      Begin VB.Frame Frame5 
         Caption         =   "Payment Status"
         Height          =   1150
         Left            =   8760
         TabIndex        =   26
         Top             =   240
         Width           =   1935
         Begin VB.CheckBox ChkPartial2 
            Caption         =   "Partial"
            Height          =   315
            Left            =   120
            TabIndex        =   29
            Top             =   800
            Width           =   975
         End
         Begin VB.CheckBox Chkunpaid2 
            Caption         =   "unpaid"
            Height          =   360
            Left            =   120
            TabIndex        =   28
            Top             =   210
            Width           =   975
         End
         Begin VB.CheckBox ChkPaid2 
            Caption         =   "paid"
            Height          =   195
            Left            =   120
            TabIndex        =   27
            Top             =   550
            Width           =   975
         End
      End
      Begin VB.Frame Frame6 
         Caption         =   "Payment Status"
         Height          =   915
         Left            =   8760
         TabIndex        =   33
         Top             =   240
         Width           =   1935
         Begin VB.CheckBox Chkunpaid 
            Caption         =   "unpaid"
            Height          =   360
            Left            =   120
            TabIndex        =   35
            Top             =   210
            Width           =   975
         End
         Begin VB.CheckBox Chkpaid 
            Caption         =   "paid"
            Height          =   195
            Left            =   120
            TabIndex        =   34
            Top             =   550
            Width           =   975
         End
      End
      Begin VB.Label Label10 
         Caption         =   "Receipt Type"
         Height          =   255
         Left            =   480
         TabIndex        =   45
         Top             =   240
         Width           =   1215
      End
      Begin VB.Label Label1 
         Caption         =   "Group Company"
         Height          =   255
         Left            =   480
         TabIndex        =   40
         Top             =   1215
         Width           =   1215
      End
      Begin VB.Label Label2 
         Caption         =   "Payor"
         Height          =   255
         Left            =   480
         TabIndex        =   39
         Top             =   580
         Width           =   1215
      End
      Begin VB.Label Label4 
         Caption         =   "Bordx No from"
         Height          =   255
         Left            =   480
         TabIndex        =   38
         Top             =   915
         Width           =   1215
      End
      Begin VB.Label Label14 
         Caption         =   "To"
         Height          =   255
         Left            =   4200
         TabIndex        =   37
         Top             =   915
         Width           =   495
      End
   End
   Begin VB.Frame Frame2 
      Height          =   550
      Left            =   120
      TabIndex        =   17
      Top             =   7400
      Width           =   11655
      Begin VB.CommandButton CmdSave 
         Caption         =   "&Save"
         Height          =   330
         Left            =   8400
         TabIndex        =   6
         Top             =   150
         Width           =   615
      End
      Begin VB.CommandButton CmdSearch 
         Caption         =   "S&earch"
         Default         =   -1  'True
         Height          =   330
         Left            =   9000
         TabIndex        =   7
         Top             =   150
         Width           =   735
      End
      Begin VB.CommandButton CmdStop 
         Caption         =   "S&top"
         Height          =   330
         Left            =   9720
         TabIndex        =   8
         Top             =   150
         Width           =   615
      End
      Begin VB.CommandButton CmdExit 
         Cancel          =   -1  'True
         Caption         =   "E&xit"
         Height          =   330
         Left            =   10920
         TabIndex        =   10
         Top             =   150
         Width           =   615
      End
      Begin VB.CommandButton CmdClear 
         Caption         =   "&Clear"
         Height          =   330
         Left            =   10320
         TabIndex        =   9
         Top             =   150
         Width           =   615
      End
      Begin VB.Label lblSelectedAmt 
         Alignment       =   1  'Right Justify
         BackColor       =   &H80000009&
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "0.00"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   1
         EndProperty
         Height          =   255
         Left            =   5160
         TabIndex        =   22
         Top             =   200
         Width           =   1335
      End
      Begin VB.Label Label19 
         Caption         =   "Amount"
         Height          =   255
         Left            =   4440
         TabIndex        =   23
         Top             =   240
         Width           =   615
      End
      Begin VB.Label lbltotal 
         Alignment       =   2  'Center
         BackColor       =   &H80000009&
         Height          =   255
         Left            =   1080
         TabIndex        =   21
         Top             =   200
         Width           =   735
      End
      Begin VB.Label Label9 
         Caption         =   "of"
         Height          =   255
         Left            =   840
         TabIndex        =   20
         Top             =   240
         Width           =   375
      End
      Begin VB.Label lblcurrent 
         Alignment       =   2  'Center
         BackColor       =   &H80000009&
         Height          =   255
         Left            =   120
         TabIndex        =   19
         Top             =   200
         Width           =   615
      End
      Begin VB.Label Label11 
         Caption         =   "Records"
         Height          =   255
         Left            =   1920
         TabIndex        =   18
         Top             =   240
         Width           =   1095
      End
   End
   Begin VB.Frame Frame3 
      Height          =   690
      Left            =   120
      TabIndex        =   12
      Top             =   6720
      Width           =   11655
      Begin VB.TextBox txtRemark 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   9720
         MaxLength       =   500
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   5
         Top             =   200
         Width           =   1785
      End
      Begin VB.TextBox txtReceiptAmt 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   5160
         TabIndex        =   3
         Top             =   240
         Width           =   1455
      End
      Begin VB.TextBox TxtReceiptNo 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   970
         MaxLength       =   20
         TabIndex        =   1
         Top             =   240
         Width           =   1575
      End
      Begin VB.TextBox txtChequeNo 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   7560
         MaxLength       =   20
         TabIndex        =   4
         Top             =   240
         Width           =   1455
      End
      Begin MSMask.MaskEdBox txtReceiptDate 
         Height          =   255
         Left            =   3060
         TabIndex        =   2
         Top             =   240
         Width           =   1050
         _ExtentX        =   1852
         _ExtentY        =   450
         _Version        =   393216
         MaxLength       =   10
         Mask            =   "##/##/####"
         PromptChar      =   "_"
      End
      Begin VB.Label Label7 
         Caption         =   "Remark"
         Height          =   255
         Left            =   9120
         TabIndex        =   24
         Top             =   240
         Width           =   615
      End
      Begin VB.Label Label3 
         Caption         =   "Receipt No"
         Height          =   255
         Left            =   120
         TabIndex        =   16
         Top             =   240
         Width           =   1095
      End
      Begin VB.Label Label5 
         Caption         =   "Date"
         Height          =   255
         Left            =   2640
         TabIndex        =   15
         Top             =   240
         Width           =   495
      End
      Begin VB.Label Label13 
         Caption         =   "Cheque No"
         Height          =   255
         Left            =   6720
         TabIndex        =   14
         Top             =   240
         Width           =   975
      End
      Begin VB.Label Label6 
         Caption         =   "Receipt Amt"
         Height          =   255
         Left            =   4200
         TabIndex        =   13
         Top             =   240
         Width           =   1095
      End
   End
   Begin VB.Frame Frame8 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   375
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   11805
      Begin VB.Label Label28 
         Appearance      =   0  'Flat
         BackColor       =   &H00000000&
         Caption         =   "Receipt from MNRB/PAYOR "
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
         TabIndex        =   11
         Top             =   50
         Width           =   3285
      End
   End
   Begin MSComctlLib.ListView lstBordx 
      Height          =   4725
      Left            =   120
      TabIndex        =   41
      Top             =   2040
      Width           =   11415
      _ExtentX        =   20135
      _ExtentY        =   8334
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
      NumItems        =   5
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Check"
         Object.Width           =   1764
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   1
         Text            =   "Payor"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "Bordx No"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   3
         Text            =   "Amount"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   4
         Text            =   "Payment Status"
         Object.Width           =   2540
      EndProperty
   End
End
Attribute VB_Name = "FrmReceipt"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public intExit As Integer
Public ReceiptFRMNRBCriteria As String
Dim BordxNoArray() As String
Private m_blnDirection(1 To 8) As Boolean

Private Sub CmdClear_Click()
    lstBordx.listitems.clear
    lstWorksheet.listitems.clear
    CboPayor.text = ""
    TxtBordx1.text = ""
    TxtBordx2.text = ""
    CboGRPCom.text = ""
    Chkunpaid.value = 0
    Chkpaid.value = 0
    Chkunpaid2.value = 0
    ChkPaid2.value = 0
    ChkPartial2.value = 0
    lblcurrent = 0
    lbltotal = 0
    lblSelectedAmt = 0
End Sub

Private Sub CmdExit_Click()
Dim RecordExit
    RecordExit = MsgBox("Quit this form ? ", vbYesNo + vbExclamation)
    If RecordExit = vbYes Then
        bExit = True
        Unload Me
    End If
End Sub

Private Sub ComboListing()
Dim PAY As New ADODB.Recordset
Dim GRPCOM As New ADODB.Recordset

PAY.Open "Select PAYCode, PAYCompanyName from PAY", medixmasterconn, adOpenKeyset, adLockOptimistic
GRPCOM.Open "Select Distinct GRPCompany from PLN", medixmasterconn, adOpenKeyset, adLockOptimistic
    
    Do Until PAY.EOF
        CboPayor.AddItem PAY!PayCode & " - " & Trim(PAY!PAYCompanyName)
        PAY.MoveNext
    Loop
    PAY.Close
    Set PAY = Nothing
    
        
    Do Until GRPCOM.EOF
        If Len(GRPCOM!GRPCompany) Then
            CboGRPCom.AddItem GRPCOM!GRPCompany
        End If
        GRPCOM.MoveNext
    Loop
    GRPCOM.Close
    Set GRPCOM = Nothing
    
End Sub

Private Sub cmdSave_Click()
Dim ii As Integer
Dim Update
Dim Bordxno As String
Dim status As String
Dim delimiter1 As Integer
    
    If TxtReceiptNo = "" Then
       MsgBox "Please Enter Receipt No !", vbCritical
       TxtReceiptNo.SetFocus
    ElseIf txtReceiptDate = "__/__/____" Then
       MsgBox "Please Enter Receipt Date !", vbCritical
       txtReceiptDate.SetFocus
    ElseIf txtChequeNo = "" Then
       MsgBox "Please Enter Cheque No. !", vbCritical
       txtChequeNo.SetFocus
    ElseIf txtReceiptAmt <> lblSelectedAmt And txtRemark = "" Then
       MsgBox "Amount receipt is incorrect " & vbNewLine & "Please key in remark for reason.", vbCritical + vbOKOnly
       txtRemark.SetFocus
    Else
        Update = MsgBox("Do you want to update the record ?", vbYesNo + vbExclamation)
        If Update = vbYes Then
            For ii = 0 To UBound(BordxNoArray)
                If Trim(BordxNoArray(ii)) <> 0 Then
                    
                    Bordxno = Trim(BordxNoArray(ii))
                    ', 1, IIf(delimiter1 = 0, 6, delimiter1 - 1)))
                    'status = Trim(Mid(BordxNoArray(ii), delimiter1 + 8, Len(BordxNoArray(ii)) - delimiter1))
                    Call UpdateRecord(Bordxno) ', status)
                End If
            Next ii
            Erase BordxNoArray
            Arraycnt = 0
            ReDim BordxNoArray(0)
            Call ClearForm(Me)
            MsgBox "Record Updated...", vbOKOnly + vbInformation, DialogBoxTitle
            Call ReceiptBordxListing
        End If
    End If
End Sub

Private Sub cmdSearch_Click()
    Screen.MousePointer = vbHourglass
        Call SearchRecord
    Screen.MousePointer = Default
End Sub

Private Sub cmdStop_Click()
    intExit = 1
    CmdExit.Enabled = True
End Sub

Private Sub Form_Load()
    Call ComboListing
    ReDim BordxNoArray(0)
    lstWorksheet.Visible = False
    Frame5.Visible = False
    intExit = 0
    lblcurrent = 0
    lbltotal = 0
    lblSelectedAmt = 0
End Sub


Private Sub lstBordx_GotFocus()
Dim i, newbound As Integer
Dim tempArray() As String
Dim Found As Boolean
Dim listloop, ListCount As Integer
Dim TotalAmt As Variant
Dim RECStatus As String
Dim RECAmt As Integer
Dim strsqlBOR As String
Dim BOR As New ADODB.Recordset
Dim strsqlREC As String
Dim REC As New ADODB.Recordset
Dim Bordxno As String
Static Arraycnt As Integer

If Source = 1 Then
   Bordxno = Trim(lstBordx.SelectedItem.SubItems(2))
   
   strsqlBOR = "select BordxRefNo,BordxMNRBFrStatus from Bordx where BordxRefNo = '" & Trim(Bordxno) & "'"
   BOR.Open strsqlBOR, medixmasterconnWS, adOpenKeyset, adLockOptimistic
   
   strsqlREC = "select * from ReceiptfrMNRB where BordxNo='" & Trim(Bordxno)
   REC.Open strsqlREC, medixmasterconnPayment, adOpenKeyset, adLockOptimistic
   
   If lstBordx.SelectedItem.Checked = True Then
     
      If BOR!BordxMNRBFrStatus = True Then
         MsgBox "Record has been clear !", vbCritical
         lstBordx.SelectedItem.Checked = False
      Else
         ReDim Preserve BordxNoArray(Arraycnt)
         BordxNoArray(Arraycnt) = lstBordx.SelectedItem.SubItems(2) '& "~" & lstBordx.SelectedItem.SubItems(4)
         Arraycnt = Arraycnt + 1
      End If
   Else
      newbound = 0
      Bordxno = Trim(lstBordx.SelectedItem.SubItems(2))
      If UBound(BordxNoArray) <> 0 Then
         For i = 0 To UBound(BordxNoArray)
            If (lstBordx.SelectedItem.SubItems(2)) <> BordxNoArray(i) Then
                ReDim Preserve tempArray(newbound)
                tempArray(newbound) = BordxNoArray(i)
                newbound = newbound + 1
            End If
         Next i
         lblSelectedAmt = lblSelectedAmt - CDbl(Format(lstBordx.SelectedItem.SubItems(4), "#,##0.00"))
         Erase BordxNoArray
         If UBound(tempArray) = 0 Then
            Arraycnt = 1
         Else
            Arraycnt = UBound(tempArray) + 1
         End If
         ReDim BordxNoArray(UBound(tempArray))
         For i = 0 To UBound(tempArray)
             BordxNoArray(i) = tempArray(1)
         Next i
         Erase tempArray
      Else
         lblSelectedAmt = lblSelectedAmt - CDbl(Format(lstBordx.SelectedItem.SubItems(3), "#,##0.00"))
         Erase BordxNoArray
         ReDim BordxNoArray(0)
         Arraycnt = 0
      End If
      If Len(BOR!BordxMNRBFrStatus) Then
         lstBordx.SelectedItem.SubItems(4) = BOR!BordxMNRBFrStatus
      Else
         lstBordx.SelectedItem.SubItems(4) = ""
      End If
      If REC.recordCount Then
         lstBordx.SelectedItem.SubItems(3) = Format(REC!BordxClaimAmt, "#,##0.00")
      Else
         lstBordx.SelectedItem.SubItems(3) = ""
      End If
   End If
   BOR.Close
   Set BOR = Nothing
   REC.Close
   Set REC = Nothing
   Source = 0
End If
End Sub

Private Sub lstBordx_ItemCheck(ByVal Item As MSComctlLib.ListItem)
Dim i, newbound As Integer
Dim tempArray() As String
Dim Found As String
Dim listloop As Integer
Dim ListCount As Integer
Dim TotalAmt As Variant
Dim RECStatus As String
Dim RECAmt As Integer
Dim strsqlBOR As String
Dim strsqlREC As String
Dim BOR As New ADODB.Recordset
Dim REC As New ADODB.Recordset
Dim Bordxno As String
Static Arraycnt As Integer
 
   
    If Item.Checked = True Then
    
        Bordxno = Trim(Mid(Item.SubItems(2), 1, IIf(InStr(1, Item.SubItems(2), "-") = 0, 10, InStr(1, Item.SubItems(2), "-") - 1)))
            
        strsqlBOR = "select BordxRefNo, BordxMNRBFrStatus from Bordx where BordxRefNo = '" & Trim(Bordxno) & "'"
        BOR.Open strsqlBOR, medixmasterconnWS, adOpenKeyset, adLockOptimistic
              
        strsqlREC = "select * from ReceiptfrMNRB where BordxNo = '" & Trim(Bordxno) & "'"
        REC.Open strsqlREC, medixmasterconnPayment, adOpenKeyset, adLockOptimistic
           
        If BOR!BordxMNRBFrStatus = True Then
            MsgBox "Record has been paid !", vbCritical
            Item.Checked = False
        Else
          
            Item.SubItems(4) = "Full"
            lblSelectedAmt = lblSelectedAmt + CDbl(Format(Item.SubItems(3), "#,##0.00"))
            ReDim Preserve BordxNoArray(Arraycnt)
            BordxNoArray(Arraycnt) = Item.SubItems(2) '& "~" & Item.SubItems(4)
            Arraycnt = Arraycnt + 1
        End If
    BOR.Close
    Set BOR = Nothing
    REC.Close
    Set REC = Nothing

    Else
       
       newbound = 0
       
       If UBound(BordxNoArray) <> 0 Then
          For i = 0 To UBound(BordxNoArray)
              If (Item.SubItems(2)) <> BordxNoArray(i) Then
                  ReDim Preserve tempArray(newbound)
                  tempArray(newbound) = BordxNoArray(i)
                  newbound = newbound + 1
              End If
          Next
          
          If Item.SubItems(4) = "Full" Then
            lblSelectedAmt = lblSelectedAmt - CDbl(Format(Item.SubItems(3), "#,##0.00"))
            Item.SubItems(4) = ""
          End If
          
          Erase BordxNoArray
          ReDim BordxNoArray(0)
          ReDim tempArray(0)
          
          If UBound(tempArray) = 0 Then
             Arraycnt = 1
          Else
             Arraycnt = UBound(tempArray) + 1
          End If
          
          ReDim BordxNoArray(UBound(tempArray))
          For i = 0 To UBound(tempArray)
              BordxNoArray(i) = tempArray(i)
          Next
          Erase tempArray
          ReDim tempArray(0)
        
       Else
            If Item.SubItems(4) = "Full" Then
                lblSelectedAmt = lblSelectedAmt - CDbl(Format(Item.SubItems(3), "#,##0.00"))
                Item.SubItems(4) = ""
            Else
                Item.SubItems(4) = ""
            End If
          
            Erase BordxNoArray
            ReDim BordxNoArray(0)
            Arraycnt = 0

      
       End If
    End If
End Sub


Private Function SearchRecord()
Dim strsql As String
Dim strAppend As String
Dim Sql As String
 
 strAppend = ""
 If Len(Trim(CboPayor)) Then
    strAppend = strAppend + " And BordxPAYCode = '" & Trim(Mid(CboPayor, 1, IIf(InStr(1, CboPayor, "-") = 0, 4, InStr(1, CboPayor, "-") - 1))) & "'"
 End If
 If Len(Trim(TxtBordx1)) Then
    strAppend = strAppend + " and BordxRefNo >= '" & RemStr(Trim(TxtBordx1)) & "'"
 End If
 If Len(Trim(TxtBordx2)) Then
    strAppend = strAppend + " and BordxRefNo <= '" & RemStr(Trim(TxtBordx2)) & "'"
 End If
 If Len(Trim(CboGRPCom)) Then
    strAppend = strAppend + " and GRPCompany = '" & Trim(Mid(CboGRPCom, 1, IIf(InStr(1, CboGRPCom, 1, "-") = 0, 4, InStr(1, CboGRPCom, "-") - 1))) & "'"
 End If
 If Chkunpaid.value = 1 Then
    strAppend = strAppend + " and BordxMNRBFrStatus = '0'"
 End If
 If Chkpaid.value = 1 Then
    strAppend = strAppend + " and BordxMNRBFrStatus = '1'"
 End If
 If ChkPartial2.value = 1 Then
   strAppend = strAppend + " and PVPayMNRBStatus = 'P'"
 End If
 If ChkPaid2.value = 1 Then
   strAppend = strAppend + " and PVPayMNRBStatus = 'F'"
 End If
 If Chkunpaid2.value = 1 Then
   strAppend = strAppend + " and PVPayMNRBStatus Is null " & ""
 End If
 
 If OptBordx.value = True Then
       ReceiptFRMNRBCriteria = strAppend
       Call ReceiptBordxListing(strAppend)
 Else
       ReceiptFRMNRBCriteria = strAppend
       Call ReceiptWorksheetListing(strAppend)
 End If
 
End Function


Private Sub ReceiptBordxListing(Optional strAppend As String)
Dim rst2 As New ADODB.Recordset
Dim Sql As String
Dim Total As Integer
Dim Current As Integer
Dim check As Boolean
Dim RECFull As Boolean
Dim RECPartial As String

lstBordx.listitems.clear

Sql = "Select BordxRefNo, BordxClaimAmt,BordxPAYCode, " & _
      " BordxMNRBFrStatus From VIEW_BordxReceipt where 0 = 0"

If Len(Trim(strAppend)) Then
   Sql = Sql + strAppend
End If

rst2.Open Sql, medixmasterconnWS, adOpenKeyset, adLockOptimistic

If rst2.recordCount = 0 Then
    MsgBox "No record found !", vbOKOnly + vbCritical, DialogBoxTitle
End If
Do
    Do Until rst2.EOF
    Total = rst2.recordCount
    lbltotal = CInt(Total)
    With rst2
        Set Record = lstBordx.listitems.Add(, , "")
               
        If Len(!BordxPAYCode) Then
           Record.SubItems(1) = !BordxPAYCode
        End If
        If Len(!BordxRefNo) Then
           Record.SubItems(2) = !BordxRefNo
        End If
        If Len(!BordxClaimAmt) Then
           Record.SubItems(3) = Format(!BordxClaimAmt, "#,##0.00")
        End If
        If Len(!BordxMNRBFrStatus) Then
           Record.SubItems(4) = !BordxMNRBFrStatus
        End If
        
        Current = Current + 1
        lblcurrent = Current
    End With
    rst2.MoveNext
    DoEvents
        If intExit = 1 Then
            check = False
            Exit Do
        End If
    Loop
Loop Until check = False
intExit = 0
rst2.Close
Set rst2 = Nothing
End Sub

Private Sub ReceiptWorksheetListing(Optional strAppend As String)
Dim rst2 As New ADODB.Recordset
Dim Sql As String
Dim Total As Integer
Dim Current As Integer
Dim check As Boolean
Dim RECFull As Boolean
Dim RECPartial As String

lstWorksheet.listitems.clear

Sql = "Select BordxRefNo, BordxClaimAmt,BordxPAYCode, " & _
      " BordxMNRBFrStatus From VIEW_BordxReceipt where 0 = 0"

If Len(Trim(strAppend)) Then
   Sql = Sql + strAppend
End If

rst2.Open Sql, medixmasterconnWS, adOpenKeyset, adLockOptimistic

If rst2.recordCount = 0 Then
    MsgBox "No record found !", vbOKOnly + vbCritical, DialogBoxTitle
End If
Do
    Do Until rst2.EOF
    Total = rst2.recordCount
    lbltotal = CInt(Total)
    With rst2
        Set Record = lstWorksheet.listitems.Add(, , "")
               
        If Len(!BordxPAYCode) Then
           Record.SubItems(1) = !BordxPAYCode
        End If
        If Len(!BordxRefNo) Then
           Record.SubItems(2) = !BordxRefNo
        End If
        If Len(!BordxClaimAmt) Then
           Record.SubItems(3) = Format(!BordxClaimAmt, "#,##0.00")
        End If
        If Len(!BordxMNRBFrStatus) Then
           Record.SubItems(4) = !BordxMNRBFrStatus
        End If
        
        Current = Current + 1
        lblcurrent = Current
    End With
    rst2.MoveNext
    DoEvents
        If intExit = 1 Then
            check = False
            Exit Do
        End If
    Loop
Loop Until check = False
intExit = 0
rst2.Close
Set rst2 = Nothing
End Sub

'*************************** Start Diable Non Integer Value
Private Sub TxtReceiptAmt_KeyPress(KeyAscii As Integer)
    KeyAscii = KeyNumber(KeyAscii)
End Sub


Private Sub CboPayor_Change()
    If InStr(CboPayor.text, "null") Then
       CboPayor.Enabled = False
    End If
End Sub

Private Sub CboPayor_KeyPress(KeyAscii As Integer)
Dim NewText As String
Dim ValidCount As Integer
Dim ValidValue As String
Dim i As Integer
  If KeyAscii >= 32 And KeyAscii <> 127 Then
    NewText = LCase(Left(CboPayor.text, CboPayor.SelStart) + Chr(KeyAscii) + Mid(CboPayor.text, CboPayor.SelStart + CboPayor.SelLength + 1))
 
    ValidCount = 0
    ValidValue = ""
 
    For i = 0 To CboPayor.ListCount - 1
 
        If NewText = LCase(Left(CboPayor.list(i), Len(NewText))) Then
                ValidCount = ValidCount + 1
                ValidValue = CboPayor.list(i)
        End If
    Next i
    If ValidCount <= 1 Then KeyAscii = 0
        If ValidCount = 1 Then
            CboPayor.text = ValidValue
            CboPayor.SelStart = 0
            CboPayor.SelLength = Len(ValidValue)
        End If
  End If
End Sub
Private Sub CboGrpCom_Change()
    If InStr(CboGRPCom.text, "null") Then
       CboGRPCom.Enabled = False
    End If
End Sub

Private Sub CboGrpCom_KeyPress(KeyAscii As Integer)
Dim NewText As String
Dim ValidCount As Integer
Dim ValidValue As String
Dim i As Integer
 If KeyAscii >= 32 And KeyAscii <> 127 Then
    NewText = LCase(Left(CboGRPCom.text, CboGRPCom.SelStart) + Chr(KeyAscii) + Mid(CboGRPCom.text, CboGRPCom.SelStart + CboGRPCom.SelLength + 1))
 
    ValidCount = 0
    ValidValue = ""
 
    For i = 0 To CboGRPCom.ListCount - 1
 
        If NewText = LCase(Left(CboGRPCom.list(i), Len(NewText))) Then
                ValidCount = ValidCount + 1
                ValidValue = CboGRPCom.list(i)
        End If
    Next i
    If ValidCount <= 1 Then KeyAscii = 0
    If ValidCount = 1 Then
       CboGRPCom.text = ValidValue
       CboGRPCom.SelStart = 0
       CboGRPCom.SelLength = Len(ValidValue)
    End If
 End If
End Sub

Private Sub UpdateRecord(Bordxno As String) ', status As String)
Dim Update
Dim strsqlCLM As String
Dim strsqlREC As String
Dim strsqlBordx As String
Dim CLM As New ADODB.Recordset
Dim REC As New ADODB.Recordset
Dim Bordx As New ADODB.Recordset
 

strsqlREC = "Select RecMNRBIndex, * from ReceiptfrMNRB Where BordxNo = '" & Trim(Bordxno) & "'"
REC.Open strsqlREC, medixmasterconnPayment, adOpenKeyset, adLockOptimistic

If REC.recordCount > 0 Then
    MsgBox "Record Duplicated !", vbCritical, "Error"
    REC.Close
    TxtReceiptNo.SetFocus
    TxtReceiptNo.SelStart = 0
    TxtReceiptNo.SelLength = Len(TxtWorkCode)
    Set REC = Nothing
    Exit Sub
Else
    REC.AddNew
    With REC
        !ReceiptNo = TxtReceiptNo
        !ReceiptDate = txtReceiptDate
        !ReceiptAmt = txtReceiptAmt
        !ChequeNo = txtChequeNo
        !Bordxno = Bordxno
        !Remark = txtRemark
        !MNRBLastUpdatedate = Date
        !MNRBLastUpdateUser = Logon
    End With
    REC.Update
    REC.Close
    Set REC = Nothing
    
    'strsqlCLM = "select * from claim where BordxRefNo = '" & Trim(Bordxno) & "'"
    'CLM.Open strsqlCLM, medixmasterconnWS, adOpenKeyset, adLockOptimistic
    'With CLM
        '!PVPayMNRBStatus = Mid(status, 1, 1)
    'End With
    'CLM.Update
    'CLM.Close
    'Set CLM = Nothing
    
    strsqlBordx = "select * from Bordx where BordxrefNo = '" & Trim(Bordxno) & "'"
    Bordx.Open strsqlBordx, medixmasterconnWS, adOpenKeyset, adLockOptimistic
    With Bordx
        !BordxMNRBRecDate = txtReceiptDate
        !BordxMNRBFrStatus = True
    End With
    Bordx.Update
    Bordx.Close
    Set Bordx = Nothing
End If
lblSelectedAmt = 0
End Sub

Private Sub OptBordx_Click()
     If OptBordx.value = True Then
        lstBordx.Visible = True
        Frame6.Visible = True
        Frame5.Visible = False
        lstBordx.listitems.clear
        CboPayor.text = ""
        TxtBordx1.text = ""
        TxtBordx2.text = ""
        CboGRPCom.text = ""
        Chkunpaid.value = 0
        Chkpaid.value = 0
        lstWorksheet.Visible = False
     End If
End Sub

Private Sub OptWks_Click()
    If OptWks.value = True Then
       lstWorksheet.Visible = True
       Frame5.Visible = True
       Frame6.Visible = False
       lstBordx.listitems.clear
       CboPayor.text = ""
        TxtBordx1.text = ""
        TxtBordx2.text = ""
        CboGRPCom.text = ""
        ChkPaid2.value = 0
        Chkunpaid2.value = 0
        ChkPartial2.value = 0
       lstBordx.Visible = False
    End If
End Sub
