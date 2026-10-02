VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Begin VB.Form FrmReceiptOther 
   Caption         =   "Receipt (Others)"
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
   Begin VB.Frame Frame2 
      Height          =   550
      Left            =   0
      TabIndex        =   52
      Top             =   7200
      Width           =   11775
      Begin VB.TextBox txtCHKAmt 
         Height          =   225
         Left            =   3360
         TabIndex        =   61
         Top             =   210
         Width           =   1335
      End
      Begin VB.TextBox lblCurRecord 
         Height          =   285
         Left            =   240
         TabIndex        =   60
         Top             =   120
         Width           =   855
      End
      Begin VB.CommandButton CmdDelete 
         Caption         =   "&Delete"
         Height          =   330
         Left            =   6840
         TabIndex        =   25
         Top             =   180
         Width           =   615
      End
      Begin VB.CommandButton CmdUpdate 
         Caption         =   "&Update"
         Height          =   330
         Left            =   6120
         TabIndex        =   24
         Top             =   180
         Width           =   735
      End
      Begin VB.CommandButton CmdSave 
         Caption         =   "&Save"
         Height          =   330
         Left            =   4920
         TabIndex        =   22
         Top             =   180
         Width           =   615
      End
      Begin VB.CommandButton CmdSearch 
         Caption         =   "&Search"
         Height          =   330
         Left            =   7440
         TabIndex        =   26
         Top             =   180
         Width           =   735
      End
      Begin VB.CommandButton CmdExit 
         Cancel          =   -1  'True
         Caption         =   "E&xit"
         Height          =   330
         Left            =   11040
         TabIndex        =   31
         Top             =   180
         Width           =   615
      End
      Begin VB.CommandButton CmdClear 
         Caption         =   "&Clear"
         Height          =   330
         Left            =   10440
         TabIndex        =   30
         Top             =   180
         Width           =   615
      End
      Begin VB.CommandButton CmdStop 
         Caption         =   "S&top"
         Height          =   330
         Left            =   8160
         TabIndex        =   27
         Top             =   180
         Width           =   615
      End
      Begin VB.CommandButton CmdEdit 
         Caption         =   "E&dit"
         Height          =   330
         Left            =   5520
         TabIndex        =   23
         Top             =   180
         Width           =   615
      End
      Begin VB.CommandButton cmdEXPtoExcel 
         Caption         =   "Print to &File"
         Height          =   330
         Left            =   8760
         TabIndex        =   28
         Top             =   180
         Width           =   975
      End
      Begin VB.CommandButton cmdPrintPV 
         Caption         =   "Print &OR"
         Height          =   330
         Left            =   9720
         TabIndex        =   29
         Top             =   180
         Width           =   735
      End
      Begin VB.Label lblTotal 
         Alignment       =   1  'Right Justify
         BackColor       =   &H80000009&
         Height          =   255
         Left            =   1200
         TabIndex        =   54
         Top             =   600
         Visible         =   0   'False
         Width           =   735
      End
      Begin VB.Label Label9 
         Caption         =   "of"
         Height          =   255
         Left            =   960
         TabIndex        =   56
         Top             =   600
         Visible         =   0   'False
         Width           =   375
      End
      Begin VB.Label Label11 
         Caption         =   "Records"
         Height          =   255
         Left            =   1440
         TabIndex        =   55
         Top             =   240
         Width           =   735
      End
      Begin VB.Label Label15 
         Caption         =   "Receipt Amt"
         Height          =   255
         Left            =   2400
         TabIndex        =   53
         Top             =   240
         Width           =   975
      End
   End
   Begin VB.Frame Frame3 
      Height          =   1455
      Left            =   0
      TabIndex        =   43
      Top             =   5880
      Width           =   11775
      Begin VB.ComboBox CboHos2 
         Height          =   315
         Left            =   1080
         TabIndex        =   13
         Top             =   180
         Width           =   5760
      End
      Begin VB.TextBox TxtRecRemarks 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1080
         Left            =   7680
         MaxLength       =   200
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   21
         Top             =   240
         Width           =   3825
      End
      Begin VB.ComboBox cboAccPeriodMth 
         Height          =   315
         ItemData        =   "FrmReceiptHOS.frx":0000
         Left            =   5640
         List            =   "FrmReceiptHOS.frx":0002
         TabIndex        =   19
         Top             =   480
         Width           =   1215
      End
      Begin VB.TextBox txtChequeNo 
         Height          =   285
         Left            =   3240
         MaxLength       =   20
         TabIndex        =   17
         Top             =   480
         Width           =   1095
      End
      Begin VB.TextBox txtReceiptNo 
         Height          =   285
         Left            =   1080
         MaxLength       =   20
         TabIndex        =   14
         Top             =   480
         Width           =   1095
      End
      Begin MSMask.MaskEdBox txtReceiptDate 
         Height          =   285
         Left            =   1080
         TabIndex        =   15
         Top             =   720
         Width           =   1095
         _ExtentX        =   1931
         _ExtentY        =   503
         _Version        =   393216
         MaxLength       =   10
         Mask            =   "##/##/####"
         PromptChar      =   "_"
      End
      Begin VB.ComboBox cboAccPeriodYR 
         Height          =   315
         ItemData        =   "FrmReceiptHOS.frx":0004
         Left            =   5640
         List            =   "FrmReceiptHOS.frx":0006
         TabIndex        =   20
         Top             =   720
         Width           =   1215
      End
      Begin VB.TextBox TxtBankCharge 
         Height          =   285
         Left            =   3240
         MaxLength       =   8
         TabIndex        =   18
         Top             =   720
         Width           =   1095
      End
      Begin VB.TextBox txtAmt 
         Height          =   285
         Left            =   1080
         MaxLength       =   8
         TabIndex        =   16
         Top             =   975
         Width           =   1095
      End
      Begin VB.Label Label2 
         Caption         =   "Hospital"
         Height          =   255
         Left            =   120
         TabIndex        =   58
         Top             =   250
         Width           =   1215
      End
      Begin VB.Label Label3 
         Caption         =   "Receipt No"
         Height          =   255
         Left            =   120
         TabIndex        =   51
         Top             =   495
         Width           =   1095
      End
      Begin VB.Label Label5 
         Caption         =   "Receipt Date"
         Height          =   255
         Left            =   120
         TabIndex        =   50
         Top             =   765
         Width           =   1095
      End
      Begin VB.Label Label13 
         Caption         =   "Cheque No"
         Height          =   255
         Left            =   2280
         TabIndex        =   49
         Top             =   495
         Width           =   975
      End
      Begin VB.Label Label6 
         Caption         =   "Amount"
         Height          =   255
         Left            =   120
         TabIndex        =   48
         Top             =   1035
         Width           =   735
      End
      Begin VB.Label Label1 
         Caption         =   "Remarks"
         Height          =   255
         Left            =   6960
         TabIndex        =   47
         Top             =   240
         Width           =   1215
      End
      Begin VB.Label Label26 
         Caption         =   "A/C Period Yr"
         Height          =   255
         Left            =   4440
         TabIndex        =   46
         Top             =   720
         Width           =   1335
      End
      Begin VB.Label Label12 
         Caption         =   "A/C Period Mth"
         Height          =   255
         Left            =   4440
         TabIndex        =   45
         Top             =   480
         Width           =   1335
      End
      Begin VB.Label Label31 
         Caption         =   "Bank Charge"
         Height          =   255
         Left            =   2280
         TabIndex        =   44
         Top             =   780
         Width           =   975
      End
   End
   Begin VB.Frame Frame1 
      Height          =   1155
      Left            =   0
      TabIndex        =   32
      Top             =   240
      Width           =   11775
      Begin VB.ComboBox CboHos 
         Height          =   315
         Left            =   1800
         TabIndex        =   0
         Top             =   210
         Width           =   4635
      End
      Begin VB.ComboBox cboACPeriodMthTo 
         Height          =   315
         ItemData        =   "FrmReceiptHOS.frx":0008
         Left            =   9960
         List            =   "FrmReceiptHOS.frx":000A
         TabIndex        =   8
         Top             =   210
         Width           =   1215
      End
      Begin VB.ComboBox cboACPeriodMthFr 
         Height          =   315
         ItemData        =   "FrmReceiptHOS.frx":000C
         Left            =   8280
         List            =   "FrmReceiptHOS.frx":000E
         TabIndex        =   7
         Top             =   210
         Width           =   1215
      End
      Begin VB.TextBox txtRECNoFr 
         Height          =   285
         Left            =   1800
         MaxLength       =   20
         TabIndex        =   1
         Top             =   500
         Width           =   1455
      End
      Begin VB.CheckBox ChkRecpNo 
         Height          =   255
         Left            =   5160
         TabIndex        =   3
         Top             =   520
         Width           =   375
      End
      Begin VB.CheckBox chkPeriodMth 
         Height          =   255
         Left            =   11235
         TabIndex        =   9
         Top             =   270
         Width           =   375
      End
      Begin VB.CheckBox chkRecpDate 
         Height          =   255
         Left            =   5160
         TabIndex        =   6
         Top             =   800
         Width           =   375
      End
      Begin VB.CheckBox ChkPeriodYr 
         Height          =   255
         Left            =   11235
         TabIndex        =   12
         Top             =   585
         Width           =   375
      End
      Begin VB.TextBox txtRECNoTo 
         Height          =   285
         Left            =   3720
         MaxLength       =   20
         TabIndex        =   2
         Top             =   500
         Width           =   1335
      End
      Begin VB.ComboBox cboACPeriodYRFr 
         Height          =   315
         ItemData        =   "FrmReceiptHOS.frx":0010
         Left            =   8280
         List            =   "FrmReceiptHOS.frx":0012
         TabIndex        =   10
         Top             =   510
         Width           =   1215
      End
      Begin VB.ComboBox cboACPeriodYRTo 
         Height          =   315
         ItemData        =   "FrmReceiptHOS.frx":0014
         Left            =   9960
         List            =   "FrmReceiptHOS.frx":0016
         TabIndex        =   11
         Top             =   510
         Width           =   1215
      End
      Begin MSMask.MaskEdBox txtRECDateFr 
         Height          =   315
         Left            =   1800
         TabIndex        =   4
         Top             =   720
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   556
         _Version        =   393216
         MaxLength       =   10
         Mask            =   "##/##/####"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox txtRECDateTo 
         Height          =   315
         Left            =   3720
         TabIndex        =   5
         Top             =   720
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   556
         _Version        =   393216
         MaxLength       =   10
         Mask            =   "##/##/####"
         PromptChar      =   "_"
      End
      Begin VB.Label Label10 
         Caption         =   "Hospital"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   240
         TabIndex        =   41
         Top             =   240
         Width           =   1215
      End
      Begin VB.Label Label4 
         Caption         =   "To"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   3360
         TabIndex        =   40
         Top             =   840
         Width           =   375
      End
      Begin VB.Label Label17 
         Caption         =   "Recp Date from"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   240
         TabIndex        =   39
         Top             =   840
         Width           =   1455
      End
      Begin VB.Label Label16 
         Caption         =   "Recp No from"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   240
         TabIndex        =   38
         Top             =   550
         Width           =   1215
      End
      Begin VB.Label Label7 
         Caption         =   "To"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   0
         Left            =   3360
         TabIndex        =   37
         Top             =   555
         Width           =   375
      End
      Begin VB.Label Label25 
         Caption         =   "A/C Period Yr"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   6720
         TabIndex        =   36
         Top             =   600
         Width           =   1815
      End
      Begin VB.Label Label24 
         Caption         =   "To"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   9600
         TabIndex        =   35
         Top             =   600
         Width           =   375
      End
      Begin VB.Label Label23 
         Caption         =   "To"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   9600
         TabIndex        =   34
         Top             =   315
         Width           =   375
      End
      Begin VB.Label Label22 
         Caption         =   "A/C Period Mth"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   6720
         TabIndex        =   33
         Top             =   315
         Width           =   1815
      End
   End
   Begin MSComctlLib.ListView lstReceiptHOS 
      Height          =   4455
      Left            =   0
      TabIndex        =   42
      Top             =   1440
      Width           =   11775
      _ExtentX        =   20770
      _ExtentY        =   7858
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
      NumItems        =   8
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Hospital Name"
         Object.Width           =   7056
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Receipt No"
         Object.Width           =   2117
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   2
         Text            =   "Receipt Date"
         Object.Width           =   2117
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   3
         Text            =   "Receipt Amt"
         Object.Width           =   2117
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   4
         Text            =   "Bank Charge"
         Object.Width           =   2117
      EndProperty
      BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   5
         Text            =   "Cheque No"
         Object.Width           =   2469
      EndProperty
      BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   6
         Text            =   "A/C Period"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   7
         Text            =   "ReceiptIndex"
         Object.Width           =   0
      EndProperty
   End
   Begin VB.Label Label28 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00404040&
      Caption         =   "Receipt (Others)"
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
      TabIndex        =   59
      Top             =   0
      Width           =   11805
   End
   Begin VB.Label Label33 
      Caption         =   "Bold - Searchable"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   57
      Top             =   7770
      Width           =   1695
   End
End
Attribute VB_Name = "FrmReceiptOther"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public intExit As Integer
Dim EditFlag As Boolean
Private m_blnDirection(1 To 8) As Boolean


Private Sub CmdDelete_Click()
Dim rsDeleteOther As New ADODB.Recordset
Dim strsql As String
Dim DeleteRecord

    If lstReceiptHOS.ListItems.Count = 0 Then
       MsgBox "Please search the record", vbOKOnly + vbCritical, DialogBoxTitle
    Else
        
        DeleteRecord = MsgBox("Do you want to delete this record?", vbYesNo + vbExclamation)
        If DeleteRecord = vbYes Then
            medixmasterconnPayment.Open "File Name=" & BOSPath & "Conn\mediconnectPayment.UDL"

            strsql = "Delete from ReceiptFrHOS Where ReceiptIndex = '" & Trim(lstReceiptHOS.SelectedItem.SubItems(7)) & "'"
            rsDeleteOther.Open strsql, medixmasterconnPayment, adOpenDynamic, adLockOptimistic
            MsgBox "Record Deleted...", vbOKOnly + vbInformation, DialogBoxTitle
            Call ClearForm(Me)
            Call ReceiptFrHOSListing
            
            medixmasterconnPayment.Close
            Set medixmasterconnPayment = Nothing
            
        End If
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

Private Sub CmdClear_Click()
    Call ClearForm(Me)
    CmdSave.Enabled = True
    lstReceiptHOS.ListItems.Clear
    txtCHKAmt = 0
    lblCurRecord = 0
    lblTotal = 0
End Sub

Private Sub cmdEdit_Click()
    EditFlag = True
End Sub

Private Sub cmdExit_Click()
    Unload Me
End Sub

'************************Start ComboListing**************************

Private Sub ComboListing()
Dim HOS As New ADODB.Recordset
Dim monthstart
Dim yearstart
Dim cmd As New ADODB.Command


        Set cmd.ActiveConnection = medixmasterconn
        cmd.CommandText = "SearchHOS"
        cmd.CommandType = adCmdStoredProc
        cmd.CommandTimeout = 15
        HOS.CursorLocation = adUseClient
        HOS.CursorType = adOpenForwardOnly
        HOS.CacheSize = 100
        Set HOS = cmd.Execute
        
    Do Until HOS.EOF
        CboHos.AddItem Trim(HOS!HOSName)
        CboHos2.AddItem HOS!HosCode & " - " & Trim(HOS!HOSName)
        HOS.MoveNext
    Loop
    HOS.Close
    Set HOS = Nothing
    
    For monthstart = 1 To 12
        With cboAccPeriodMth
            .AddItem monthstart
        End With
        With cboACPeriodMthFr
            .AddItem monthstart
        End With
        With cboACPeriodMthTo
            .AddItem monthstart
        End With
    Next monthstart
    
    For yearstart = 1999 To 3000
        With cboAccPeriodYR
            .AddItem yearstart
        End With
        With cboACPeriodYRFr
            .AddItem yearstart
        End With
        With cboACPeriodYRTo
            .AddItem yearstart
        End With
    Next yearstart
    
End Sub
'************************Start ComboListing**************************

Private Function SearchRecord()
Dim strsql As String
Dim strAppend As String
Dim sql As String
Dim Hospitals As String
    
    strAppend = ""
    Hospitals = ""
        
    
    If Len(Trim(CboHos)) Then
        Hospitals = Replace(CboHos, "'", "''")
        strAppend = strAppend + " And HOSName = '" & Trim(Hospitals) & "'"
    End If
                
    If Len(Trim(txtRecNoFr)) Then
        strAppend = strAppend + " AND ReceiptNo >= '" & RemStr(Trim(txtRecNoFr)) & "'"
    End If
    
    If Len(Trim(txtRecNoFr)) Then
        strAppend = strAppend + " AND ReceiptNo <= '" & RemStr(Trim(txtRecNoFr)) & "'"
    End If
    
    If ChkRecpNo.Value = 1 Then
        sql = ""
        sql = " AND (ReceiptNo Is null OR ReceiptNo = '')" & ""
        strAppend = strAppend + sql
    End If
    
    If Len(Trim(txtRecDateFr)) > 0 And IsDate(txtRecDateFr) = True Then
        strAppend = strAppend + " And ReceiptDate >= '" & Format(txtRecDateFr, "mm/dd/yyyy") & "'"
    End If
    
    If Len(Trim(txtRecDateTo)) > 0 And IsDate(txtRecDateTo) = True Then
        strAppend = strAppend + " And ReceiptDate <= '" & Format(txtRecDateTo, "mm/dd/yyyy") & "'"
    End If
    
    If chkRecpDate.Value = 1 Then
        sql = ""
        sql = " AND (ReceiptDate Is null OR ReceiptDate = '')" & ""
        strAppend = strAppend + sql
    End If
    
    If Len(cboACPeriodMthFr) Then
        strAppend = strAppend + " AND RECMBMAccPeriodMth >= '" & Trim(cboACPeriodMthFr) & "'"
    End If
    
    If Len(cboACPeriodMthTo) Then
        strAppend = strAppend + " AND RECMBMAccPeriodMth <= '" & Trim(cboACPeriodMthTo) & "'"
    End If
    
    If chkPeriodMth.Value = 1 Then
        sql = ""
        sql = " AND (RECMBMAccPeriodMth Is null OR RECMBMAccPeriodMth = '')" & ""
        strAppend = strAppend + sql
    End If

    If Len(cboACPeriodYRFr) Then
        strAppend = strAppend + " AND RECMBMAccPeriodYr >= '" & Trim(cboACPeriodYRFr) & "'"
    End If
    
    If Len(cboACPeriodYRTo) Then
        strAppend = strAppend + " AND RECMBMAccPeriodYr <= '" & Trim(cboACPeriodYRTo) & "'"
    End If
    
    If ChkPeriodYr.Value = 1 Then
        sql = ""
        sql = " AND (RECMBMAccPeriodYr Is null OR RECMBMAccPeriodYr = '')" & ""
        strAppend = strAppend + sql
    End If

    
    Call ReceiptFrHOSListing(strAppend)
   
End Function

Private Sub ReceiptFrHOSListing(Optional strAppend As String)
Dim rst2 As New ADODB.Recordset
Dim REC As New ADODB.Recordset
Dim Record As ListItem
Dim sql As String
Dim strsqlREC As String
Dim total As String
Dim Current As String
Dim TotalAmt As String

    total = 0
    Current = 0
    TotalAmt = 0
    
    lstReceiptHOS.ListItems.Clear
    
    sql = " Select * From VIEW_ReceiptFrOthers Where 0 = 0 "
        
    If Len(Trim(strAppend)) Then
        sql = sql + strAppend
    End If

    rst2.Open sql, medixmasterconnPayment, adOpenForwardOnly, adLockBatchOptimistic
   
    If rst2.EOF = True Then
        MsgBox "No record found!", vbOKOnly + vbCritical, DialogBoxTitle
        CmdSearch.Enabled = True
    Else
    
    Do
        Do Until rst2.EOF
        'total = rst2.recordCount
        'lblTotal = total
            
            With rst2
                                
                If Len(!HOSName) Then
                    Set Record = lstReceiptHOS.ListItems.Add(, , Trim(!HOSName))
                Else
                    Set Record = lstReceiptHOS.ListItems.Add(, , "Empty")
                End If
                                
                If Len(!ReceiptNo) Then
                    Record.SubItems(1) = Trim(!ReceiptNo)
                End If
                
                
                If Len(!ReceiptDate) Then
                    Record.SubItems(2) = Format(!ReceiptDate, "dd/mm/yyyy")
                End If
                
                If Len(!ReceiptAmt) Then
                    Record.SubItems(3) = Format(!ReceiptAmt, "#,##0.00")
                End If
                
                If Len(!HOSBankCharge) Then
                    Record.SubItems(4) = Format(!HOSBankCharge, "#,##0.00")
                End If
                
                If Len(!ReceiptChequeNo) Then
                    Record.SubItems(5) = Trim(!ReceiptChequeNo)
                End If
                
                If Len(rst2!RECMBMAccPeriodMth) Or Len(rst2!RECMBMAccPeriodYr) Then
                    Record.SubItems(6) = rst2!RECMBMAccPeriodMth & "/" & rst2!RECMBMAccPeriodYr
                End If
                
                If Len(rst2!ReceiptIndex) Then
                    Record.SubItems(7) = rst2!ReceiptIndex
                End If
                
                If Len(!ReceiptAmt) Then
                    TotalAmt = TotalAmt + !ReceiptAmt
                End If
                txtCHKAmt = Format(TotalAmt, "#,##0.00")
                
                Current = Current + 1
                lblCurRecord = Current
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
    
    CmdSearch.Enabled = True
    
End Sub

Private Sub cmdEXPtoExcel_Click()
    'Screen.MousePointer = vbHourglass
        'Call ExportListViewtoExcel(lstReceiptHOS)
    'Screen.MousePointer = vbDefault
    
    Screen.MousePointer = vbHourglass
        
    If lstReceiptHOS.ListItems.Count <> 0 Then
        Call ExportListViewtoExcel(lstReceiptHOS)
    Else
        MsgBox "Report cannot be printed without any record.", vbInformation
    End If
        
    Screen.MousePointer = vbDefault
    
End Sub

Private Sub cmdSave_Click()
Dim RecordSaved
Dim SaveOthers As New ADODB.Recordset
Dim strsql As String
    CmdSave.Enabled = False
    medixmasterconnPayment.Open "File Name=" & BOSPath & "Conn\mediconnectPayment.UDL"

    If CboHos2 = "" Then
       MsgBox "Please Enter Hospital Code", vbOKOnly + vbCritical, DialogBoxTitle
       CboHos2.SetFocus
    ElseIf TxtReceiptNo = "" Then
        MsgBox "Please Enter Receipt No", vbOKOnly + vbCritical, DialogBoxTitle
        TxtReceiptNo.SetFocus
    ElseIf txtReceiptDate = "__/__/____" Then
        MsgBox "Please Enter Receipt Date", vbOKOnly + vbCritical, DialogBoxTitle
        txtReceiptDate.SetFocus
    ElseIf txtAmt = "" Then
        MsgBox "Please Enter Receipt Amount", vbOKOnly + vbCritical, DialogBoxTitle
        txtAmt.SetFocus
    ElseIf txtChequeNo = "" Then
        MsgBox "Please Enter Cheque No", vbOKOnly + vbCritical, DialogBoxTitle
        txtChequeNo.SetFocus
    Else
        
        RecordSaved = MsgBox("Do you want to save the record?", vbYesNo + vbExclamation)
        If RecordSaved = vbYes Then
            
                SaveOthers.Open "ReceiptFrHOS", medixmasterconnPayment, adOpenKeyset, adLockOptimistic
                SaveOthers.AddNew
                        
                With SaveOthers
                    !HOSName = Trim(Mid(CboHos2, 1, 5))
                    !ReceiptNo = ChkStr(TxtReceiptNo)
                    !ReceiptDate = Format(txtReceiptDate, "dd/mm/yyyy")
                    !ReceiptAmt = Format(txtAmt, "#,##0.00")
                    If TxtBankCharge <> "" Then
                        !HOSBankCharge = Format(TxtBankCharge, "#,##0.00")
                    End If
                    !ReceiptChequeNo = ChkStr(Trim(txtChequeNo))
                    If cboAccPeriodMth <> "" Then
                        !RECMBMAccPeriodMth = Trim(cboAccPeriodMth)
                    End If
                    If cboAccPeriodYR <> "" Then
                        !RECMBMAccPeriodYr = Trim(cboAccPeriodYR)
                    End If
                    !ReceiptRemarks = ChkStr(Trim(TxtRecRemarks))
                    !ReceiptLastUpdateUser = Logon
                    !ReceiptLastUpdateDate = Date
                End With
                SaveOthers.Update
                Call ClearForm(Me)
                Call ReceiptFrHOSListing
                
                    MsgBox "Record Saved...", vbOKOnly + vbInformation, DialogBoxTitle
                
            SaveOthers.Close
            Set SaveOthers = Nothing
        End If
    End If
    medixmasterconnPayment.Close
    Set medixmasterconnPayment = Nothing
    CmdSave.Enabled = True
End Sub

Private Sub cmdSearch_Click()
        
    Screen.MousePointer = vbHourglass
        CmdSearch.Enabled = False
        
        medixmasterconnPayment.Open "File Name=" & BOSPath & "Conn\mediconnectPayment.UDL"
        Call SearchRecord
        medixmasterconnPayment.Close
        Set medixmasterconnPayment = Nothing
        
        CmdSearch.Enabled = True
        
    Screen.MousePointer = vbDefault
        
End Sub

Private Sub CmdStop_Click()
    intExit = 1
    CmdExit.Enabled = True
    CmdSearch.Enabled = True
End Sub


Private Sub cmdUpdate_Click()
Dim RecordSaved
Dim RptMBM As New ADODB.Recordset
Dim UpdateRptMBM As New ADODB.Recordset
Dim strsql As String
    
    medixmasterconnPayment.Open "File Name=" & BOSPath & "Conn\mediconnectPayment.UDL"
    
    If EditFlag = False Then

    ElseIf lstReceiptHOS.ListItems.Count = 0 Then
        MsgBox "Please select the record", vbOKOnly + vbCritical
           
        ElseIf CboHos2 = "" Then
            MsgBox "Please Enter Hospital Name!", vbCritical
            CboHos2.SetFocus
        ElseIf TxtReceiptNo = "" Then
            MsgBox "Please Enter Receipt No!", vbCritical
            TxtReceiptNo.SetFocus
        ElseIf IsDate(txtReceiptDate) = False Then
            MsgBox "Invalid Receipt Date!", vbCritical
            txtReceiptDate.SetFocus
        ElseIf txtReceiptDate = "__/__/____" Then
            MsgBox "Please Enter Receipt Date", vbOKOnly + vbCritical, DialogBoxTitle
            txtReceiptDate.SetFocus
        ElseIf txtChequeNo = "" Then
            MsgBox "Please Enter Cheque No!", vbCritical
            txtChequeNo.SetFocus
        ElseIf txtAmt = "" Then
            MsgBox "Please Enter Cheque Amt!", vbCritical
            txtAmt.SetFocus
               
        Else
            strsql = "Select * From ReceiptFrHOS Where ReceiptIndex = '" & Trim(lstReceiptHOS.SelectedItem.SubItems(7)) & "'"
                     
            RptMBM.Open strsql, medixmasterconnPayment, adOpenKeyset, adLockOptimistic
            
            RecordSaved = MsgBox("Do you want to update the record?", vbYesNo + vbExclamation)
            If RecordSaved = vbYes Then
                
                If RptMBM.recordCount = 0 Then
                    RptMBM.AddNew
                End If
                
                With RptMBM
                    !ReceiptNo = TxtReceiptNo
                    !ReceiptDate = txtReceiptDate
                    !ReceiptAmt = txtAmt
                    If TxtBankCharge <> "" Then
                        !HOSBankCharge = Format(TxtBankCharge, "#,##0.00")
                    End If
                    !ReceiptChequeNo = txtChequeNo
                    !HOSName = Trim(Mid(CboHos2, 1, 5))
                    !ReceiptRemarks = Trim(TxtRecRemarks)
                    !ReceiptLastUpdateDate = Date
                    !ReceiptLastUpdateUser = Logon
                    If cboAccPeriodMth <> "" Then
                        !RECMBMAccPeriodMth = cboAccPeriodMth
                    End If
                    If cboAccPeriodYR <> "" Then
                        !RECMBMAccPeriodYr = cboAccPeriodYR
                    End If
            
                    .Update
                End With
                MsgBox "Record Updated...", vbOKOnly + vbInformation, "Information"
          
                Call SearchRecord
                EditFlag = False
            End If
            RptMBM.Close
            Set RptMBM = Nothing
            
        TxtReceiptNo.Text = ""
        txtReceiptDate = "__/__/____"
        txtAmt.Text = ""
        txtChequeNo.Text = ""
        TxtRecRemarks.Text = ""
        TxtBankCharge = ""
        cboAccPeriodMth = ""
        cboAccPeriodYR = ""
        CboHos2 = ""
        CmdSave.Enabled = True
    End If
    medixmasterconnPayment.Close
    Set medixmasterconnPayment = Nothing

End Sub

Private Sub Form_Load()
    EditFlag = False
    medixmasterconn.Open "File Name=" & BOSPath & "Conn\mediconnect.UDL"
    Call ComboListing
    medixmasterconn.Close
    Set medixmasterconn = Nothing

    txtCHKAmt = 0
    intExit = 0
End Sub

Private Sub lstReceiptHOS_Click()
Dim rst3 As New ADODB.Recordset
Dim strsql As String

medixmasterconnPayment.Open "File Name=" & BOSPath & "Conn\mediconnectPayment.UDL"

If EditFlag = True Then

    CmdSave.Enabled = False

    If lstReceiptHOS.ListItems.Count Then

        TxtReceiptNo = ""
        txtReceiptDate.mask = ""
        txtReceiptDate.Text = ""
        txtReceiptDate.mask = "##/##/####"
        TxtBankCharge = ""
        txtAmt = ""
        txtChequeNo = ""
        TxtRecRemarks = ""
        cboAccPeriodMth = ""
        cboAccPeriodYR = ""


        strsql = " SELECT * From ReceiptFrHOS Where ReceiptIndex = '" & lstReceiptHOS.SelectedItem.SubItems(7) & "'"
        rst3.Open strsql, medixmasterconnPayment, adOpenKeyset, adLockOptimistic
    
        If rst3.recordCount = 0 Then
            MsgBox "Payment hasn't been made!", vbOKOnly + vbCritical, DialogBoxTitle
        Else
            TxtReceiptNo.SetFocus
            
            If Len(rst3!HOSName) Then
                CboHos2 = rst3!HOSName
            End If
            If Len(rst3!ReceiptNo) Then
                TxtReceiptNo = rst3!ReceiptNo
            End If
            If Len(rst3!ReceiptDate) Then
                txtReceiptDate = rst3!ReceiptDate
            End If
            If Len(rst3!ReceiptAmt) Then
                txtAmt = rst3!ReceiptAmt
            End If
            If Len(rst3!HOSBankCharge) Then
                TxtBankCharge = rst3!HOSBankCharge
            End If
            If Len(rst3!ReceiptChequeNo) Then
                txtChequeNo = rst3!ReceiptChequeNo
            End If
            If Len(rst3!ReceiptRemarks) Then
                TxtRecRemarks = rst3!ReceiptRemarks
            End If
            If Len(rst3!RECMBMAccPeriodMth) Then
                cboAccPeriodMth = rst3!RECMBMAccPeriodMth
            End If
            If Len(rst3!RECMBMAccPeriodYr) Then
                cboAccPeriodYR = rst3!RECMBMAccPeriodYr
            End If
            CmdSave.Enabled = False
            rst3.Close
            Set rst3 = Nothing
        End If
    End If
End If

'medixmasterconnWS.Close
'Set medixmasterconnWS = Nothing
medixmasterconnPayment.Close
Set medixmasterconnPayment = Nothing

End Sub


Private Sub txtAmt_KeyPress(KeyAscii As Integer)
    KeyAscii = KeyNumber(KeyAscii)
End Sub

Private Sub TxtBankCharge_KeyPress(KeyAscii As Integer)
    KeyAscii = KeyNumber(KeyAscii)
End Sub

Private Sub txtReceiptDate_LostFocus()
    If IsDate(txtReceiptDate) Then
        cboAccPeriodMth.Text = Format(txtReceiptDate, "mm")
        cboAccPeriodYR.Text = Format(txtReceiptDate, "yyyy")
    Else
        If txtReceiptDate <> "__/__/____" Then
            MsgBox "Invalid Date Format! ", vbCritical
            txtReceiptDate.SetFocus
        End If
    End If
End Sub

Private Sub lstReceiptHOS_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
    m_blnDirection(ColumnHeader.Index) = Not m_blnDirection(ColumnHeader.Index)
    Select Case ColumnHeader.Index
    Case 1
        SortListView lstReceiptHOS, ColumnHeader.Index, ldtString, m_blnDirection(1)
    Case 2
        SortListView lstReceiptHOS, ColumnHeader.Index, ldtString, m_blnDirection(2)
    Case 3
        SortListView lstReceiptHOS, ColumnHeader.Index, ldtDateTime, m_blnDirection(3)
    Case 4
        SortListView lstReceiptHOS, ColumnHeader.Index, ldtNumber, m_blnDirection(4)
    Case 5
        SortListView lstReceiptHOS, ColumnHeader.Index, ldtNumber, m_blnDirection(5)
    Case 6
        SortListView lstReceiptHOS, ColumnHeader.Index, ldtString, m_blnDirection(6)
    Case 7
        SortListView lstReceiptHOS, ColumnHeader.Index, ldtNumber, m_blnDirection(7)
    Case 8
        SortListView lstReceiptHOS, ColumnHeader.Index, ldtNumber, m_blnDirection(8)
    End Select
End Sub


Private Sub cboHOS_KeyPress(KeyAscii As Integer)
Dim NewText As String
Dim ValidCount As Integer
Dim ValidValue As String
Dim i As Integer
    If KeyAscii >= 32 And KeyAscii <> 127 Then
    NewText = LCase(Left(CboHos.Text, CboHos.SelStart) + Chr(KeyAscii) + Mid(CboHos.Text, CboHos.SelStart + CboHos.SelLength + 1))
   
    ValidCount = 0
    ValidValue = ""
    
    For i = 0 To CboHos.ListCount - 1
       
        If NewText = LCase(Left(CboHos.List(i), Len(NewText))) Then
                ValidCount = ValidCount + 1
                ValidValue = CboHos.List(i)
        End If
        Next
        If ValidCount <= 1 Then KeyAscii = 0
            If ValidCount = 1 Then
               CboHos.Text = ValidValue
               CboHos.SelStart = 0
               CboHos.SelLength = Len(ValidValue)
            End If
        End If
End Sub

Private Sub cboHOS2_KeyPress(KeyAscii As Integer)
Dim NewText As String
Dim ValidCount As Integer
Dim ValidValue As String
Dim i As Integer
    If KeyAscii >= 32 And KeyAscii <> 127 Then
    NewText = LCase(Left(CboHos2.Text, CboHos2.SelStart) + Chr(KeyAscii) + Mid(CboHos2.Text, CboHos2.SelStart + CboHos2.SelLength + 1))
   
    ValidCount = 0
    ValidValue = ""
    
    For i = 0 To CboHos2.ListCount - 1
       
        If NewText = LCase(Left(CboHos2.List(i), Len(NewText))) Then
                ValidCount = ValidCount + 1
                ValidValue = CboHos2.List(i)
        End If
        Next
        If ValidCount <= 1 Then KeyAscii = 0
            If ValidCount = 1 Then
               CboHos2.Text = ValidValue
               CboHos2.SelStart = 0
               CboHos2.SelLength = Len(ValidValue)
            End If
        End If
End Sub
