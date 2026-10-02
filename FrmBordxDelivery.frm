VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form FrmBordxDelivery 
   Caption         =   "Bordx Delivery"
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
   Begin VB.Frame Frame8 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   450
      Left            =   0
      TabIndex        =   20
      Top             =   0
      Width           =   11925
      Begin VB.Label Label28 
         Appearance      =   0  'Flat
         BackColor       =   &H00000000&
         Caption         =   "Set Bordx Delivery/Account Date"
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
         TabIndex        =   21
         Top             =   120
         Width           =   4605
      End
   End
   Begin VB.Frame Frame1 
      Height          =   1275
      Left            =   0
      TabIndex        =   23
      Top             =   360
      Width           =   4695
      Begin VB.TextBox TxtBordx2 
         Height          =   285
         Left            =   3120
         MaxLength       =   20
         TabIndex        =   26
         Top             =   630
         Width           =   1455
      End
      Begin VB.ComboBox CboPayor 
         Height          =   315
         Left            =   1200
         Style           =   2  'Dropdown List
         TabIndex        =   25
         Top             =   270
         Width           =   3375
      End
      Begin VB.TextBox TxtBordx1 
         Height          =   285
         Left            =   1200
         MaxLength       =   20
         TabIndex        =   24
         Top             =   630
         Width           =   1335
      End
      Begin VB.Label Label14 
         Caption         =   "To"
         Height          =   255
         Left            =   2640
         TabIndex        =   29
         Top             =   630
         Width           =   255
      End
      Begin VB.Label Label8 
         Caption         =   "Bordx No from"
         Height          =   255
         Left            =   120
         TabIndex        =   28
         Top             =   630
         Width           =   1095
      End
      Begin VB.Label Label1 
         Caption         =   "Payor"
         Height          =   255
         Left            =   120
         TabIndex        =   27
         Top             =   240
         Width           =   1215
      End
   End
   Begin VB.Frame Frame3 
      Height          =   550
      Left            =   0
      TabIndex        =   13
      Top             =   6480
      Width           =   11775
      Begin VB.ComboBox CboPeriodYr 
         Height          =   315
         ItemData        =   "FrmBordxDelivery.frx":0000
         Left            =   6360
         List            =   "FrmBordxDelivery.frx":0002
         TabIndex        =   15
         Top             =   150
         Width           =   1335
      End
      Begin VB.ComboBox CboPeriodMth 
         Height          =   315
         ItemData        =   "FrmBordxDelivery.frx":0004
         Left            =   3480
         List            =   "FrmBordxDelivery.frx":0006
         TabIndex        =   14
         Top             =   150
         Width           =   1335
      End
      Begin MSMask.MaskEdBox TxtDate 
         Height          =   255
         Left            =   600
         TabIndex        =   16
         Top             =   195
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   450
         _Version        =   393216
         MaxLength       =   10
         Mask            =   "##/##/####"
         PromptChar      =   "_"
      End
      Begin VB.Label LblTotalAmt 
         Alignment       =   1  'Right Justify
         BackColor       =   &H80000009&
         Height          =   255
         Left            =   8640
         TabIndex        =   31
         Top             =   210
         Width           =   975
      End
      Begin VB.Label Label29 
         Caption         =   "Total Amt"
         Height          =   255
         Left            =   7920
         TabIndex        =   30
         Top             =   210
         Width           =   735
      End
      Begin VB.Label Label19 
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
         Left            =   5040
         TabIndex        =   19
         Top             =   195
         Width           =   1215
      End
      Begin VB.Label Label11 
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
         Left            =   2040
         TabIndex        =   18
         Top             =   195
         Width           =   1455
      End
      Begin VB.Label Label2 
         Caption         =   "Date"
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
         Left            =   120
         TabIndex        =   17
         Top             =   195
         Width           =   495
      End
   End
   Begin VB.Frame Frame2 
      Height          =   615
      Left            =   0
      TabIndex        =   0
      Top             =   7080
      Width           =   11775
      Begin VB.CommandButton cmdPrint 
         Caption         =   "Print to &File"
         Height          =   375
         Left            =   9240
         TabIndex        =   8
         Top             =   180
         Width           =   1095
      End
      Begin VB.CommandButton cmdStop 
         Caption         =   "S&top"
         Height          =   375
         Left            =   5160
         TabIndex        =   7
         Top             =   180
         Width           =   735
      End
      Begin VB.CommandButton CmdClear 
         Caption         =   "&Clear"
         Height          =   375
         Left            =   10320
         TabIndex        =   6
         Top             =   180
         Width           =   615
      End
      Begin VB.CommandButton CmdExit 
         Cancel          =   -1  'True
         Caption         =   "E&xit"
         Height          =   375
         Left            =   10920
         TabIndex        =   5
         Top             =   180
         Width           =   615
      End
      Begin VB.CommandButton CmdToMNRB 
         Caption         =   "To &MNRB"
         Height          =   375
         Left            =   8400
         TabIndex        =   4
         Top             =   180
         Width           =   855
      End
      Begin VB.CommandButton CmdFromPayor 
         Caption         =   "&From Payor"
         Height          =   375
         Left            =   7440
         TabIndex        =   3
         Top             =   180
         Width           =   975
      End
      Begin VB.CommandButton CmdSearch 
         Caption         =   "&Search"
         Default         =   -1  'True
         Height          =   375
         Left            =   4440
         TabIndex        =   2
         Top             =   180
         Width           =   735
      End
      Begin VB.CommandButton CmdToPayor 
         Caption         =   "To &Payor"
         Height          =   375
         Left            =   6600
         TabIndex        =   1
         Top             =   180
         Width           =   855
      End
      Begin VB.Label lblTotal 
         Alignment       =   1  'Right Justify
         BackColor       =   &H80000009&
         Height          =   255
         Left            =   1080
         TabIndex        =   12
         Top             =   240
         Width           =   615
      End
      Begin VB.Label Label21 
         Caption         =   "Records"
         Height          =   255
         Left            =   1800
         TabIndex        =   11
         Top             =   240
         Width           =   1095
      End
      Begin VB.Label lblCurrent 
         Alignment       =   1  'Right Justify
         BackColor       =   &H80000009&
         Height          =   255
         Left            =   120
         TabIndex        =   10
         Top             =   240
         Width           =   615
      End
      Begin VB.Label Label22 
         Caption         =   "of"
         Height          =   255
         Left            =   840
         TabIndex        =   9
         Top             =   240
         Width           =   375
      End
   End
   Begin MSComctlLib.ListView lstBordxDelivery 
      Height          =   4815
      Left            =   0
      TabIndex        =   22
      Top             =   1680
      Width           =   11895
      _ExtentX        =   20981
      _ExtentY        =   8493
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
      NumItems        =   14
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Check"
         Object.Width           =   1411
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   1
         Text            =   "Payor"
         Object.Width           =   1764
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   2
         Text            =   "Bordx No"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   3
         Text            =   "Bordx Date"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   4
         Text            =   "Bordx Amount"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   5
         Text            =   "Sent to Payor Status"
         Object.Width           =   1764
      EndProperty
      BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   6
         Text            =   "Date Sent to Payor"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   7
         Text            =   "A/C Period"
         Object.Width           =   2117
      EndProperty
      BeginProperty ColumnHeader(9) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   8
         Text            =   "Return Fr Payor Status"
         Object.Width           =   1764
      EndProperty
      BeginProperty ColumnHeader(10) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   9
         Text            =   "Date Returned Fr Payor"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(11) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   10
         Text            =   "A/C Period"
         Object.Width           =   2117
      EndProperty
      BeginProperty ColumnHeader(12) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   11
         Text            =   "Sent to MNRB Status"
         Object.Width           =   1764
      EndProperty
      BeginProperty ColumnHeader(13) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   12
         Text            =   "Date Sent to MNRB"
         Object.Width           =   1764
      EndProperty
      BeginProperty ColumnHeader(14) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   13
         Text            =   "A/C Period"
         Object.Width           =   2117
      EndProperty
   End
   Begin VB.Frame Frame4 
      Height          =   1215
      Left            =   4800
      TabIndex        =   32
      Top             =   400
      Width           =   7095
      Begin VB.Frame Frame7 
         Height          =   495
         Left            =   120
         TabIndex        =   43
         Top             =   720
         Width           =   3495
         Begin VB.OptionButton optBoth1 
            Caption         =   "Both"
            Height          =   255
            Left            =   2640
            TabIndex        =   46
            Top             =   120
            Value           =   -1  'True
            Width           =   735
         End
         Begin VB.OptionButton ChkNotReturn 
            Caption         =   "Not Yet Return"
            Height          =   255
            Left            =   1200
            TabIndex        =   45
            Top             =   120
            Width           =   1455
         End
         Begin VB.OptionButton ChkReturn 
            Caption         =   "Returned"
            Height          =   255
            Left            =   120
            TabIndex        =   44
            Top             =   120
            Width           =   1095
         End
      End
      Begin VB.Frame Frame5 
         Height          =   495
         Left            =   120
         TabIndex        =   33
         Top             =   280
         Width           =   3495
         Begin VB.OptionButton ChkSent 
            Caption         =   "Sent "
            Height          =   255
            Left            =   120
            TabIndex        =   36
            Top             =   160
            Width           =   975
         End
         Begin VB.OptionButton ChkNotSent 
            Caption         =   "Not Yet Sent"
            Height          =   255
            Left            =   1200
            TabIndex        =   35
            Top             =   160
            Width           =   1335
         End
         Begin VB.OptionButton optBoth 
            Caption         =   "Both"
            Height          =   255
            Left            =   2640
            TabIndex        =   34
            Top             =   160
            Value           =   -1  'True
            Width           =   735
         End
      End
      Begin VB.Frame Frame6 
         Height          =   495
         Left            =   3720
         TabIndex        =   37
         Top             =   280
         Width           =   3255
         Begin VB.OptionButton ChkMNRB 
            Caption         =   "Sent"
            Height          =   255
            Left            =   120
            TabIndex        =   40
            Top             =   200
            Width           =   855
         End
         Begin VB.OptionButton ChkNotMNRB 
            Caption         =   "Not Yet Sent"
            Height          =   195
            Left            =   960
            TabIndex        =   39
            Top             =   200
            Width           =   1455
         End
         Begin VB.OptionButton optBoth3 
            Caption         =   "Both"
            Height          =   255
            Left            =   2400
            TabIndex        =   38
            Top             =   200
            Value           =   -1  'True
            Width           =   735
         End
      End
      Begin VB.Label Label3 
         Caption         =   "Payor"
         Height          =   255
         Left            =   120
         TabIndex        =   42
         Top             =   120
         Width           =   1215
      End
      Begin VB.Label Label4 
         Caption         =   "MNRB"
         Height          =   255
         Left            =   3720
         TabIndex        =   41
         Top             =   120
         Width           =   1215
      End
   End
End
Attribute VB_Name = "FrmBordxDelivery"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public intExit As Integer
Dim BordxNoArray() As String
Dim searchCriteria As String
Dim ItemChecked As Boolean
Private m_blnDirection(1 To 14) As Boolean

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

Private Sub Form_Load()
    intExit = 0
    lblCurrent = 0
    lblTotal = 0
    LblTotalAmt = 0
    ReDim BordxNoArray(0)
    Call ComboListing
    ItemChecked = False
End Sub


'**********************Start Command Button *******************************

Private Sub CmdPrint_Click()
    Screen.MousePointer = vbHourglass
        Call ExportToExcelAnalysis1(lstBordxDelivery)
    Screen.MousePointer = vbDefault
End Sub

Private Sub CmdClear_Click()
    CboPayor.Clear
    TxtDate.Text = "__/__/____"
    CboPeriodMth.Text = ""
    CboPeriodYr.Text = ""
    TxtBordx1.Text = ""
    TxtBordx2.Text = ""
    lstBordxDelivery.listitems.Clear
    optBoth.value = True
    optBoth1.value = True
    optBoth3.value = True
    lblCurrent = 0
    lblTotal = 0
    LblTotalAmt = 0
    ReDim BordxNoArray(0)
    Call ComboListing
End Sub

Private Sub cmdSearch_Click()
    ItemChecked = False
    LblTotalAmt = ""
    lblTotal = ""
    lblCurrent = ""
    If CboPayor = "" Then
        MsgBox "Please Enter Payor Code.", vbCritical
        CboPayor.SetFocus
    Else
        CmdClear.Enabled = False
        Screen.MousePointer = vbHourglass
            Call SearchRecord
        Screen.MousePointer = Default
    End If
End Sub

Private Function SearchRecord()
Dim strsql As String
Dim strAppend As String
Dim sql As String
Dim CAS As New ADODB.Recordset

    strAppend = ""
    If Len(Trim(CboPayor)) Then
        strAppend = strAppend + " And BordxPAYCode = '" & Trim(Mid(CboPayor, 1, IIf(InStr(1, CboPayor, "-") = 0, 2, InStr(1, CboPayor, "-") - 1))) & "'"
    End If
    
    If Len(Trim(TxtBordx1)) Then
        strAppend = strAppend + " AND BordxRefNo >= '" & RemStr(Trim(TxtBordx1)) & "'"
    End If
    
    If Len(Trim(TxtBordx2)) Then
        strAppend = strAppend + " AND BordxRefNo <= '" & RemStr(Trim(TxtBordx2)) & "'"
    End If
    
    If ChkSent.value = True Then
        strAppend = strAppend + " AND  BordxPAYToStatus = '1' "
    End If
    
    If ChkNotSent.value = True Then
        strAppend = strAppend + " AND  BordxPAYToStatus = '0' "
    End If
    
    If ChkReturn.value = True Then
        strAppend = strAppend + " AND BordxPAYFrStatus = '1' "
    End If
    
    If ChkNotReturn.value = True Then
        strAppend = strAppend + " AND BordxPAYFrStatus = '0' "
    End If
    
    If ChkMNRB.value = True Then
        strAppend = strAppend + " AND BordxMNRBToStatus = '1' "
    End If
    
    If ChkNotMNRB.value = True Then
        strAppend = strAppend + " AND BordxMNRBToStatus = '0'  "
    End If
    
    searchCriteria = strAppend
    Call BordxDeliveryListing(strAppend)
   
End Function

Private Sub CmdStop_Click()
    intExit = 1
    CmdClear.Enabled = True
    CmdExit.Enabled = True
End Sub


Private Sub CmdExit_Click()
'Dim RecordExit
'    RecordExit = MsgBox("Quit this form?", vbYesNo + vbExclamation)
'    If RecordExit = vbYes Then
'        bExit = True
        Unload Me
'    End If
End Sub

Private Sub CmdToPayor_Click()
Dim ii As Integer
Dim cnt As Integer
Dim recordFound As Boolean
Dim recordPassed As Boolean
Dim pos1, pos2 As Integer
    'recordPassed = False
    If ItemChecked = True Then
        
        If TxtDate = "__/__/____" Then
            MsgBox "Please Enter in date", vbCritical
            TxtDate.SetFocus
        'ElseIf TxtDate > Date Then
            'MsgBox "Date to Payor can not greater than System Date! ", vbCritical
            'TxtDate.SetFocus
        'ElseIf CboPeriodMth = "" Then
            'MsgBox "Month cannot be Empty.", vbCritical
            'CboPeriodMth.SetFocus
        'ElseIf CboPeriodYr = "" Then
            'MsgBox "Year cannot be empty!", vbCritical
            'CboPeriodYr.SetFocus
        Else
            Update = MsgBox("Do you want to update the record?", vbYesNo + vbExclamation)
            
            If Update = vbYes Then
                cnt = 0
                recordFound = False
                If lstBordxDelivery.listitems.count <> 0 Then
                    For ii = 0 To UBound(BordxNoArray)
                        If Trim(BordxNoArray(ii)) <> "" Then
                            pos1 = InStr(1, BordxNoArray(ii), "|") + 1
                            pos2 = InStr(pos1, BordxNoArray(ii), "&")
                            'If Mid(BordxNoArray(ii), pos1, pos2 - pos1) = False Then
                                UpdateSentPayor Mid(BordxNoArray(ii), 1, pos1 - 2)
                                recordFound = True
                            'Else
                                'recordPassed = True
                            'End If
                            
                        End If
                    Next ii
                    If recordFound = True Then
                        MsgBox "Record updated...", vbOKOnly + vbInformation, DialogBoxTitle
                    End If
                    'If recordPassed = True Then
                        'MsgBox "There is at least one record has been passed to Payor before!", vbOKOnly + vbInformation, DialogBoxTitle
                    'End If
                    Erase BordxNoArray
                    Arraycnt = 0
                    ReDim BordxNoArray(0)
                    Call BordxDeliveryListing(searchCriteria)
                End If
            End If
        End If
    Else
        MsgBox "Please select at least one record to be updated! ", vbCritical
    End If
End Sub

Private Sub CmdFromPayor_Click()
Dim ii As Integer
Dim cnt As Integer
Dim recordFound As Boolean
Dim recordPassed As Boolean
Dim pos1, pos2, pos3 As Integer
    'recordPassed = False
    If ItemChecked = True Then
        
        If TxtDate = "__/__/____" Then
            MsgBox "Please Enter in date", vbCritical
            TxtDate.SetFocus
        'ElseIf TxtDate > Date Then
            'MsgBox "Date from Payor can not greater than System Date! ", vbCritical
            'TxtDate.SetFocus
        'ElseIf CboPeriodMth = "" Then
            'MsgBox "Month cannot be Empty.", vbCritical
            'CboPeriodMth.SetFocus
        'ElseIf CboPeriodYr = "" Then
            'MsgBox "Year cannot be empty!", vbCritical
            'CboPeriodYr.SetFocus
        Else
                       
            Update = MsgBox("Do you want to update the record?", vbYesNo + vbExclamation)
            
            If Update = vbYes Then
                
                cnt = 0
           
                recordFound = False
                If lstBordxDelivery.listitems.count <> 0 Then
           
                    For ii = 0 To UBound(BordxNoArray)
                        If BordxNoArray(ii) <> "" Then
                            pos1 = InStr(1, BordxNoArray(ii), "|") + 1
                            pos2 = InStr(pos1, BordxNoArray(ii), "&")
                            pos3 = InStr(pos2, BordxNoArray(ii), "+")
                            'If Mid(BordxNoArray(ii), pos2 + 1, pos3 - pos2 - 1) = False Then
                                UpdateReturnPayor (Mid(BordxNoArray(ii), 1, pos1 - 2))
                                recordFound = True
                            'Else
                                'recordPassed = True
                            'End If
                        End If
                            
                    Next ii
                    If recordFound = True Then
                        MsgBox "Record updated...", vbOKOnly + vbInformation, DialogBoxTitle
                    'Else
                        'MsgBox "Record cannot be updated...", vbOKOnly + vbCritical, DialogBoxTitle
                    End If
                    'If recordPassed = True Then
                        'MsgBox "There is at least one record has been passed to Payor before!", vbOKOnly + vbInformation, DialogBoxTitle
                    'End If
                    
                    Erase BordxNoArray
                    Arraycnt = 0
                    ReDim BordxNoArray(0)
                    Call BordxDeliveryListing(searchCriteria)
                End If
            End If
         End If
    Else
      MsgBox "Please select at least one record to be updated! ", vbCritical
    End If
End Sub

Private Sub CmdToMNRB_Click()
Dim ii As Integer
Dim cnt As Integer
Dim recordFound As Boolean
Dim recordPassed As Boolean
Dim pos1, pos2, pos3, pos4 As Integer
    
'recordPassed = False
If ItemChecked = True Then
     
    
    If TxtDate = "__/__/____" Then
        MsgBox "Please Enter in date", vbCritical
        TxtDate.SetFocus
    'ElseIf TxtDate > Date Then
        'MsgBox "Date to MNRB can not greater than System Date! ", vbCritical
        'TxtDate.SetFocus
    'ElseIf CboPeriodMth = "" Then
        'MsgBox "Month cannot be Empty.", vbCritical
        'CboPeriodMth.SetFocus
    'ElseIf CboPeriodYr = "" Then
        'MsgBox "Year cannot be empty!", vbCritical
        'CboPeriodYr.SetFocus
    Else
        Update = MsgBox("Do you want to update the record?", vbYesNo + vbExclamation)
        If Update = vbYes Then
    
            cnt = 0
           
            recordFound = False
            If lstBordxDelivery.listitems.count <> 0 Then
               
               For ii = 0 To UBound(BordxNoArray)
                 If BordxNoArray(ii) <> "" Then
                        pos1 = InStr(1, BordxNoArray(ii), "|") + 1
                        pos2 = InStr(pos1, BordxNoArray(ii), "&")
                        pos3 = InStr(pos2, BordxNoArray(ii), "+")
                        
                            'If Mid(BordxNoArray(ii), pos3 + 1, pos3 - pos1 - 1) = False Then
                                UpdateSentMNRB Mid(BordxNoArray(ii), 1, pos1 - 2)
                                recordFound = True
                            'Else
                                'recordPassed = True
                            'End If
                 End If
               Next ii
               If recordFound = True Then
                    MsgBox "Record updated...", vbOKOnly + vbInformation, DialogBoxTitle
               End If
               
               'If recordPassed = True Then
                    'MsgBox "There is at least one record has been passed to Payor before!", vbOKOnly + vbInformation, DialogBoxTitle
               'End If
               Erase BordxNoArray
               Arraycnt = 0
               ReDim BordxNoArray(0)
               Call BordxDeliveryListing(searchCriteria)
            End If
        End If
    End If
    Else
       MsgBox "Please select at least one record to be updated!", vbCritical
    End If
End Sub

'Private Sub CmdFrMNRB_Click()
'Dim ii As Integer
'Dim cnt As Integer
'Dim recordFound As Boolean
'Dim recordPassed As Boolean
'Dim pos1, pos2, pos3, pos4 As Integer
    'recordPassed = False
    'If ItemChecked = True Then
        
        'If TxtDate = "__/__/____" Then
            'MsgBox "Please Enter in date", vbCritical
           'TxtDate.SetFocus
        'Else
            'Update = MsgBox("Do you want to update the record?", vbYesNo + vbExclamation)
            
            'If Update = vbYes Then
                
                'cnt = 0
           
                'recordFound = False
                'If lstBordxDelivery.listitems.count <> 0 Then
           
                    'For ii = 0 To UBound(BordxNoArray)
                        'If BordxNoArray(ii) <> "" Then
                            'pos1 = InStr(1, BordxNoArray(ii), "|") + 1
                            'pos2 = InStr(pos1, BordxNoArray(ii), "&")
                            'pos3 = InStr(pos2, BordxNoArray(ii), "+")
                            'pos4 = InStr(pos3, BordxNoArray(ii), "*")
                            'If Mid(BordxNoArray(ii), pos4 + 1, Len(BordxNoArray(ii)) - pos4) = False Then
                                'UpdateReturnMNRB Mid(BordxNoArray(ii), 1, pos1 - 2)
                                'recordFound = True
                            'Else
                                'recordPassed = True
                            'End If
                        'End If
                            
                    'Next ii
                    'If recordFound = True Then
                        'MsgBox "Record updated...", vbOKOnly + vbInformation, DialogBoxTitle
                    'Else
                        'MsgBox "Record cannot be updated...", vbOKOnly + vbCritical, DialogBoxTitle
                    'End If
                    'If recordPassed = True Then
                        'MsgBox "There is at least one record has been passed to Payor before!", vbOKOnly + vbInformation, DialogBoxTitle
                    'End If
                    
                    'Erase BordxNoArray
                    'Arraycnt = 0
                    'ReDim BordxNoArray(0)
                    'Call BordxDeliveryListing(searchCriteria)
                'End If
            'End If
         'End If
    'Else
      'MsgBox "Please select at least one record to be updated! ", vbCritical
    'End If
'End Sub
'**********************End Command Button *******************************

'**********************Start BordxDelivery Listing  *******************************
Private Sub BordxDeliveryListing(Optional strAppend As String)
Dim rst2 As New ADODB.Recordset
Dim rstCount As New ADODB.Recordset
Dim Record As ListItem
Dim sql As String
Dim sqlCount As String
Dim ClaimAdmin As Boolean
Dim Total As String
Dim Current As String
Dim TotalAmt As Variant

    
    Total = 0
    Current = 0
    
    lstBordxDelivery.listitems.Clear

    sql = " Select BordxRefNo, BordxPAYCode,  BordxDate, BordxClaimAmt, " & _
          " BordxPAYSentDate, BordxPAYRecDate, BordxMNRBSentDate, " & _
          " BordxPAYToStatus, BordxPAYFrStatus, BordxMNRBToStatus, " & _
          " BordxMNRBRecDate, BordxMNRBFrStatus, PAYToACPeriodYR, PAYToACPeriodMth, " & _
          " PAYFrACPeriodMth, PAYFrACPeriodYR, MNRBToACPeriodMth, MNRBToACPeriodYR " & _
          " From BORDX where 0 = 0 "
          '" From VIEW_Bordx_Delivery where 0 = 0 "
    
    sqlCount = "select count(*) as TotalRecord From Bordx where 0=0 "
    
    If Len(Trim(strAppend)) Then
        sql = sql + strAppend
        sqlCount = sqlCount + strAppend
    End If
    
    Set rstCount = medixmasterconnWS.Execute(sqlCount)
    
    rst2.Open sql, medixmasterconnWS, adOpenKeyset, adLockOptimistic
   
    If rstCount!totalrecord = 0 Then
        MsgBox "No record found!", vbOKOnly + vbCritical, DialogBoxTitle
    Else
         
    Do
        Do Until rst2.EOF
        Total = rstCount!totalrecord
        lblTotal = Total
        
            With rst2
                Set Record = lstBordxDelivery.listitems.Add(, , "")
                
                If Len(rst2!BordxPAYCode) Then
                    Record.SubItems(1) = rst2!BordxPAYCode
                End If
                If Len(rst2!BordxRefNo) Then
                    Record.SubItems(2) = rst2!BordxRefNo
                End If
                If Len(rst2!BordxDate) Then
                    Record.SubItems(3) = Format(rst2!BordxDate, "dd/mm/yyyy")
                End If
                If Len(rst2!BordxClaimAmt) Then
                    Record.SubItems(4) = Format(rst2!BordxClaimAmt, "#,##0.00")
                End If
                If Len(rst2!BordxPayToStatus) Then
                   Record.SubItems(5) = rst2!BordxPayToStatus
                End If
                If Len(rst2!BordxPAYSentDate) Then
                    Record.SubItems(6) = Format(rst2!BordxPAYSentDate, "dd/mm/yyyy")
                End If
                If Len(rst2!PAYToACPeriodMth & "/" & rst2!PAYToACPeriodYR) Then
                    Record.SubItems(7) = Trim(rst2!PAYToACPeriodMth & "/" & rst2!PAYToACPeriodYR)
                End If
                If Len(rst2!BordxPayFrStatus) Then
                    Record.SubItems(8) = rst2!BordxPayFrStatus
                End If
                If Len(rst2!BordxPayRecDate) Then
                    Record.SubItems(9) = Format(rst2!BordxPayRecDate, "dd/mm/yyyy")
                End If
                If Len(rst2!PAYFrACPeriodMth & "/" & rst2!PAYFrACPeriodYR) Then
                    Record.SubItems(10) = Trim(rst2!PAYFrACPeriodMth & "/" & rst2!PAYFrACPeriodYR)
                End If
                If Len(rst2!BordxMNRBToStatus) Then
                    Record.SubItems(11) = rst2!BordxMNRBToStatus
                End If
                If Len(rst2!BordxMNRBSentDate) Then
                    Record.SubItems(12) = Format(rst2!BordxMNRBSentDate, "dd/mm/yyyy")
                End If
                If Len(rst2!MNRBToACPeriodMth & "/" & rst2!MNRBToACPeriodYR) Then
                    Record.SubItems(13) = Trim(rst2!MNRBToACPeriodMth & "/" & rst2!MNRBToACPeriodYR)
                End If
                                
                'If Len(rst2!BordxMNRBFrStatus) Then
                    'Record.SubItems(11) = rst2!BordxMNRBFrStatus
                'End If
                'If Len(rst2!BordxMNRBRecDate) Then
                    'Record.SubItems(12) = Format(rst2!BordxMNRBRecDate, "dd/mm/yyyy")
                'End If
                
                Current = Current + 1
                lblCurrent = Current
                If Len(rst2!BordxClaimAmt) Then
                    TotalAmt = TotalAmt + rst2!BordxClaimAmt
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
End Sub
'**********************End BordxDelivery Listing *******************************


'**********************Start Check ListView Record*******************************
Private Sub lstBordxDelivery_ItemCheck(ByVal Item As MSComctlLib.ListItem)
Dim i, newbound  As Integer
Dim tempArray() As String
Dim Found As Boolean
Dim temp
Static Arraycnt As Integer
    
        
    If Item.Checked = True Then
            ReDim Preserve BordxNoArray(Arraycnt)
            BordxNoArray(Arraycnt) = Item.SubItems(2) & "|" & Item.SubItems(5) & "&" & Item.SubItems(8) & "+" & Item.SubItems(11) '& "*" & Item.SubItems(11)
            Arraycnt = Arraycnt + 1
            ItemChecked = True
    Else
        ' Remove from the array
        newbound = 0
    
        If UBound(BordxNoArray) <> 0 Then
            For i = 0 To UBound(BordxNoArray)
                If Trim(BordxNoArray(i)) <> "" Then
                    If (Item.SubItems(2)) <> BordxNoArray(i) Then
                        ReDim Preserve tempArray(newbound)
                        tempArray(newbound) = BordxNoArray(i)
                        newbound = newbound + 1
                    End If
                End If
            Next
            Erase BordxNoArray
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
        Else
            Erase BordxNoArray
            ReDim BordxNoArray(0)
            Arraycnt = 0
            ItemChecked = False
        End If
    End If
End Sub
'**********************End Check ListView Record*******************************

'**********************Start Update all Record functions*******************************

Private Sub UpdateSentPayor(tBordxNo As String)
Dim Update
Dim strsql As String
Dim ToPayor As New ADODB.Recordset
           
    
  strsql = " Select BordxRefNo, BordxPaySentDate, BordxPayToStatus, " & _
           " PAYToACPeriodMth, PAYToACPeriodYR " & _
           " From Bordx where BordxRefNo = '" & Trim(tBordxNo) & "'"
           
  ToPayor.Open strsql, medixmasterconnWS, adOpenKeyset, adLockOptimistic
   
  If ToPayor.recordCount > 0 Then
            
     With ToPayor
     !BordxPAYSentDate = TxtDate
     !BordxPayToStatus = True
     '!PAYToACPeriodMth = Trim(CboPeriodMth)
     '!PAYToACPeriodYR = Trim(CboPeriodYr)
     End With
     ToPayor.Update
     ToPayor.Close
     Set ToPayor = Nothing
  End If
End Sub

Private Function UpdateReturnPayor(tBordxNo As String) As Boolean
Dim Update
Dim strsql As String
Dim FRPayor As New ADODB.Recordset

           
  strsql = "Select BordxRefNo, BordxPayRecDate, BordxPAYSentDate, " & _
  " BordxPayFrStatus, PAYFrACPeriodMth, PAYFrACPeriodYR " & _
  " From Bordx where BordxRefNo ='" & Trim(tBordxNo) & "'"
  
  FRPayor.Open strsql, medixmasterconnWS, adOpenKeyset, adLockOptimistic
  
  If FRPayor.recordCount > 0 Then
    If CDate(FRPayor!BordxPAYSentDate) > CDate(TxtDate) Then
        MsgBox "Date Received from Payor cannot be ealier than Date Sent to Payor!", vbCritical
        TxtDate.SetFocus
        UpdateReturnPayor = False
    Else
      With FRPayor
      !BordxPayRecDate = TxtDate
      !BordxPayFrStatus = True
      '!PAYFrACPeriodMth = Trim(CboPeriodMth)
      '!PAYFrACPeriodYR = Trim(CboPeriodYr)
      End With
      FRPayor.Update
      UpdateReturnPayor = True
    End If
  End If
  FRPayor.Close
  Set FRPayor = Nothing
End Function

Private Sub UpdateSentMNRB(tBordxNo As String)
Dim Update
Dim strsql As String
Dim ToMNRB As New ADODB.Recordset

   strsql = "Select BordxRefNo, BordxMNRBSentDate, BordxMNRBToStatus, " & _
   " MNRBToACPeriodMth, MNRBToACPeriodYR " & _
   " From Bordx where BordxRefNo ='" & Trim(tBordxNo) & "'"
   
   ToMNRB.Open strsql, medixmasterconnWS, adOpenKeyset, adLockOptimistic
    
    If ToMNRB.recordCount > 0 Then
            
      With ToMNRB
      !BordxMNRBSentDate = TxtDate
      !BordxMNRBToStatus = True
      '!MNRBToACPeriodMth = Trim(CboPeriodMth)
      '!MNRBToACPeriodYR = Trim(CboPeriodYr)
      End With
      ToMNRB.Update
      ToMNRB.Close
      Set ToMNRB = Nothing
    End If

End Sub

'Private Function UpdateReturnMNRB(tBordxNo As String) As Boolean
'Dim Update
'Dim strsql As String
'Dim FRMNRB As New ADODB.Recordset

           
  'strsql = "Select BordxRefNo, BordxMNRBSentDate, BordxMNRBRecDate, BordxMNRBFrStatus From Bordx where BordxRefNo ='" & Trim(tBordxNo) & "'"
  'FRMNRB.Open strsql, medixmasterconnWS, adOpenKeyset, adLockOptimistic
    
  'If FRMNRB.recordCount > 0 Then
    
        'If FRMNRB!BordxMNRBSentDate > TxtDate Then
            'MsgBox "Date Received from MNRB cannot be ealier than Date Sent to MNRB!", vbCritical
            'TxtDate.SetFocus
            'UpdateReturnMNRB = False
        'Else
            'With FRMNRB
            '!BordxMNRBRecDate = TxtDate
            '!BordxMNRBFrStatus = True
            'End With
            'FRMNRB.Update
            'FRMNRB.Close
            'Set FRMNRB = Nothing
        'End If
  'End If
'End Function
'**********************End Update all Record functions*******************************

'**********************Start Sorted ListView *******************************
Private Sub lstBordxDelivery_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
    m_blnDirection(ColumnHeader.Index) = Not m_blnDirection(ColumnHeader.Index)
    Select Case ColumnHeader.Index
    Case 1
        SortListView lstBordxDelivery, ColumnHeader.Index, ldtString, m_blnDirection(1)
    Case 2
        SortListView lstBordxDelivery, ColumnHeader.Index, ldtString, m_blnDirection(2)
    Case 3
        SortListView lstBordxDelivery, ColumnHeader.Index, ldtString, m_blnDirection(3)
    Case 4
        SortListView lstBordxDelivery, ColumnHeader.Index, ldtDateTime, m_blnDirection(4)
    Case 5
        SortListView lstBordxDelivery, ColumnHeader.Index, ldtNumber, m_blnDirection(5)
    Case 6
        SortListView lstBordxDelivery, ColumnHeader.Index, ldtString, m_blnDirection(6)
    Case 7
        SortListView lstBordxDelivery, ColumnHeader.Index, ldtDateTime, m_blnDirection(7)
    Case 8
        SortListView lstBordxDelivery, ColumnHeader.Index, ldtNumber, m_blnDirection(8)
    Case 9
        SortListView lstBordxDelivery, ColumnHeader.Index, ldtString, m_blnDirection(9)
    Case 10
        SortListView lstBordxDelivery, ColumnHeader.Index, ldtDateTime, m_blnDirection(10)
    Case 11
        SortListView lstBordxDelivery, ColumnHeader.Index, ldtNumber, m_blnDirection(11)
    Case 12
        SortListView lstBordxDelivery, ColumnHeader.Index, ldtString, m_blnDirection(12)
    Case 13
        SortListView lstBordxDelivery, ColumnHeader.Index, ldtDateTime, m_blnDirection(13)
    Case 14
        SortListView lstBordxDelivery, ColumnHeader.Index, ldtNumber, m_blnDirection(14)
    End Select
End Sub
'**********************End Sorted ListView *******************************

'**********************Start Combo Listing*******************************
Private Sub ComboListing()
Dim PAY As New ADODB.Recordset
Dim MonthStart
Dim YearStart

PAY.Open "Select PAYCode, PAYCompanyName from PAY", medixmasterconn, adOpenKeyset, adLockOptimistic
    
    Do Until PAY.EOF
        CboPayor.AddItem PAY!PAYCode & " - " & Trim(PAY!PAYCompanyName)
        PAY.MoveNext
    Loop
    PAY.Close
    Set PAY = Nothing

    For MonthStart = 1 To 12
        With CboPeriodMth
            .AddItem MonthStart
        End With
    Next MonthStart
    
    For YearStart = 1999 To 3000
        With CboPeriodYr
            .AddItem YearStart
        End With
    Next YearStart

End Sub
'**********************End Combo Listing*******************************

'**********************Start Combobox List *******************************
Private Sub CboPayor_Click()
Dim strsqlPAY As String
Dim PAY As New ADODB.Recordset

If CboPayor.Text = "" Then

    strsqlPAY = "Select PAYCode, HLTCode from PAY where PAYCode = '" & Mid(CboPayor, 1, 2) & "'"
    PAY.Open strsqlPAY, medixmasterconn, adOpenKeyset, adLockOptimistic
    
    If PAY!HLTCode = "N" Then
        CmdToMNRB.Enabled = False
        'CmdFrMNRB.Enabled = False
    Else
        CmdToMNRB.Enabled = True
        'CmdFrMNRB.Enabled = True
    End If
    PAY.Close
    Set PAY = Nothing
    
End If
End Sub

Private Sub CboPayor_Change()
Dim strsqlPAY As String
Dim PAY As New ADODB.Recordset

If CboPayor.Text = "" Then
    
    strsqlPAY = "Select PAYCode, HLTCode from PAY where PAYCode = '" & Mid(CboPayor, 1, 2) & "'"
    PAY.Open strsqlPAY, medixmasterconn, adOpenKeyset, adLockOptimistic
    
    If InStr(CboPayor.Text, "null") Then
       CboPayor.Enabled = False
    End If
    
    If PAY!HLTCode = "N" Then
        CmdToMNRB.Enabled = False
        'CmdFrMNRB.Enabled = False
    Else
        CmdToMNRB.Enabled = True
        'CmdFrMNRB.Enabled = True
    End If
    PAY.Close
    Set PAY = Nothing
    
End If
End Sub

'Private Sub CboPayor_KeyPress(KeyAscii As Integer)
'Dim NewText As String
'Dim ValidCount As Integer
'Dim ValidValue As String
'Dim i As Integer
 '   If KeyAscii >= 32 And KeyAscii <> 127 Then
  '  NewText = LCase(Left(CboPayor.Text, CboPayor.SelStart) + Chr(KeyAscii) + Mid(CboPayor.Text, CboPayor.SelStart + CboPayor.SelLength + 1))
 '
  '  ValidCount = 0
   ' ValidValue = ""
 '
  '  For i = 0 To CboPayor.ListCount - 1
 '
  '      If NewText = LCase(Left(CboPayor.list(i), Len(NewText))) Then
   '             ValidCount = ValidCount + 1
    '            ValidValue = CboPayor.list(i)
     '   End If
      '  Next
       ' If ValidCount <= 1 Then KeyAscii = 0
        '    If ValidCount = 1 Then
'                CboPayor.Text = ValidValue
 '               CboPayor.SelStart = 0
  '              CboPayor.SelLength = Len(ValidValue)
     '       End If
   ''     End If
'End Sub
'**********************End Combobox List *******************************

Private Sub txtDate_LostFocus()
    If IsDate(TxtDate) Then
    Else
        If TxtDate <> "__/__/____" Then
            MsgBox "Invalid Date Format! ", vbCritical
            TxtDate.SetFocus
        End If
    End If
End Sub





