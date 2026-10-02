VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "tabctl32.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmMMA 
   Caption         =   "MMA"
   ClientHeight    =   7245
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   8820
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   7245
   ScaleWidth      =   8820
   WindowState     =   2  'Maximized
   Begin VB.Frame Frame9 
      Height          =   735
      Left            =   360
      TabIndex        =   12
      Top             =   6720
      Width           =   10995
      Begin VB.CommandButton cmdSearch 
         Caption         =   "&Search"
         Default         =   -1  'True
         Height          =   315
         Left            =   3840
         TabIndex        =   21
         Top             =   240
         Width           =   840
      End
      Begin VB.CommandButton cmdAdd 
         Caption         =   "&Add"
         Height          =   315
         Left            =   1320
         TabIndex        =   20
         Top             =   240
         Width           =   840
      End
      Begin VB.CommandButton cmdEdit 
         Caption         =   "&Edit"
         Enabled         =   0   'False
         Height          =   315
         Left            =   2160
         TabIndex        =   19
         Top             =   240
         Width           =   840
      End
      Begin VB.CommandButton cmdDelete 
         Caption         =   "&Delete"
         Enabled         =   0   'False
         Height          =   315
         Left            =   3000
         TabIndex        =   18
         Top             =   240
         Width           =   840
      End
      Begin VB.CommandButton cmdSave 
         Caption         =   "Sa&ve"
         Enabled         =   0   'False
         Height          =   315
         Left            =   7800
         TabIndex        =   17
         Top             =   240
         Width           =   840
      End
      Begin VB.CommandButton cmdCancel 
         Cancel          =   -1  'True
         Caption         =   "E&xit"
         Height          =   315
         Left            =   8640
         TabIndex        =   16
         Top             =   240
         Width           =   840
      End
      Begin VB.CommandButton CmdRefresh 
         Caption         =   "&Refresh"
         Height          =   315
         Left            =   4680
         TabIndex        =   15
         Top             =   240
         Width           =   840
      End
      Begin VB.CommandButton CmdPrint 
         Caption         =   "&Print"
         Height          =   315
         Left            =   6960
         TabIndex        =   14
         Top             =   240
         Width           =   840
      End
      Begin VB.CommandButton CmdClear 
         Caption         =   "&Clear"
         Height          =   315
         Left            =   5520
         TabIndex        =   13
         Top             =   240
         Width           =   840
      End
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   7335
      Left            =   120
      TabIndex        =   25
      Top             =   240
      Width           =   11535
      _ExtentX        =   20346
      _ExtentY        =   12938
      _Version        =   393216
      Tabs            =   2
      TabsPerRow      =   2
      TabHeight       =   520
      TabCaption(0)   =   "Procedures"
      TabPicture(0)   =   "frmMMA.frx":0000
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Label4"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "lstProceduresList"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "Frame2"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).ControlCount=   3
      TabCaption(1)   =   "Look Up Table"
      TabPicture(1)   =   "frmMMA.frx":001C
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Frame1"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).Control(1)=   "lstLookUpList"
      Tab(1).Control(1).Enabled=   0   'False
      Tab(1).Control(2)=   "Label20"
      Tab(1).Control(2).Enabled=   0   'False
      Tab(1).ControlCount=   3
      Begin VB.Frame Frame1 
         Height          =   975
         Left            =   -74760
         TabIndex        =   39
         Top             =   480
         Width           =   11055
         Begin VB.TextBox txtAnaestheCat 
            Height          =   285
            Left            =   6600
            MaxLength       =   50
            TabIndex        =   9
            Top             =   240
            Width           =   2055
         End
         Begin VB.TextBox txtAnaestheFee 
            Height          =   285
            Left            =   6600
            MaxLength       =   8
            TabIndex        =   11
            Top             =   495
            Width           =   2055
         End
         Begin VB.TextBox txtSurgeonCat 
            Height          =   285
            Left            =   2160
            MaxLength       =   10
            TabIndex        =   8
            Top             =   300
            Width           =   2055
         End
         Begin VB.TextBox txtSurgeonFee 
            Height          =   285
            Left            =   2160
            MaxLength       =   8
            TabIndex        =   10
            Top             =   555
            Width           =   2055
         End
         Begin VB.Label Label13 
            Caption         =   "Anaesthetists Fees"
            Height          =   225
            Left            =   4440
            TabIndex        =   43
            Top             =   585
            Width           =   1980
         End
         Begin VB.Label Label12 
            Caption         =   "Anaesthetist's Category"
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
            Left            =   4440
            TabIndex        =   42
            Top             =   315
            Width           =   2100
         End
         Begin VB.Label Label15 
            Caption         =   "Surgeon's Fees"
            Height          =   225
            Left            =   120
            TabIndex        =   41
            Top             =   645
            Width           =   1980
         End
         Begin VB.Label Label16 
            Caption         =   "Surgeon's Category"
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
            TabIndex        =   40
            Top             =   375
            Width           =   1740
         End
      End
      Begin VB.Frame Frame2 
         Height          =   1455
         Left            =   240
         TabIndex        =   26
         Top             =   420
         Width           =   11055
         Begin VB.TextBox txtMainBP 
            Height          =   285
            Left            =   2160
            MaxLength       =   100
            TabIndex        =   0
            Top             =   240
            Width           =   5175
         End
         Begin VB.TextBox txtOPCSNarrative 
            Height          =   465
            Left            =   2160
            MaxLength       =   3000
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   2
            Top             =   500
            Width           =   5205
         End
         Begin VB.TextBox txtOPCSCode 
            Height          =   285
            Left            =   9000
            MaxLength       =   10
            TabIndex        =   1
            Top             =   240
            Width           =   1455
         End
         Begin VB.TextBox txtSurgCat 
            Height          =   285
            Left            =   2160
            MaxLength       =   30
            TabIndex        =   3
            Top             =   935
            Width           =   1695
         End
         Begin VB.TextBox txtAnaesCat 
            Height          =   285
            Left            =   5640
            MaxLength       =   30
            TabIndex        =   4
            Top             =   935
            Width           =   1695
         End
         Begin VB.TextBox txtPageNo 
            Height          =   285
            Left            =   9000
            MaxLength       =   10
            TabIndex        =   7
            Top             =   500
            Width           =   1455
         End
         Begin VB.ComboBox cboPart 
            Height          =   315
            Left            =   9000
            TabIndex        =   5
            Top             =   750
            Width           =   1455
         End
         Begin VB.TextBox txtPartACost 
            Height          =   285
            Left            =   9000
            MaxLength       =   5
            TabIndex        =   6
            Top             =   1040
            Width           =   1455
         End
         Begin VB.Label Label11 
            Caption         =   "Page No"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   225
            Left            =   7680
            TabIndex        =   36
            Top             =   525
            Width           =   780
         End
         Begin VB.Label Label10 
            Caption         =   "Part A Cost"
            Height          =   225
            Left            =   7680
            TabIndex        =   35
            Top             =   1030
            Width           =   1020
         End
         Begin VB.Label Label9 
            Caption         =   "Part"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   300
            Left            =   7680
            TabIndex        =   34
            Top             =   780
            Width           =   660
         End
         Begin VB.Label Label1 
            Caption         =   "Anaesthetist's Category"
            Height          =   225
            Left            =   3960
            TabIndex        =   31
            Top             =   960
            Width           =   1740
         End
         Begin VB.Label Label8 
            Caption         =   "Surgeon's Category"
            Height          =   255
            Left            =   120
            TabIndex        =   30
            Top             =   1005
            Width           =   1740
         End
         Begin VB.Label Label7 
            Caption         =   "OPCS Narrative  *"
            Height          =   255
            Left            =   120
            TabIndex        =   29
            Top             =   600
            Width           =   1620
         End
         Begin VB.Label Label5 
            Caption         =   "OPCS Code *"
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
            Left            =   7680
            TabIndex        =   28
            Top             =   240
            Width           =   1260
         End
         Begin VB.Label Label2 
            Caption         =   "Main Body Part   *"
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
            TabIndex        =   27
            Top             =   240
            Width           =   1620
         End
      End
      Begin MSComctlLib.ListView lstProceduresList 
         Height          =   4320
         Left            =   240
         TabIndex        =   32
         Top             =   2160
         Width           =   11055
         _ExtentX        =   19500
         _ExtentY        =   7620
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
         NumItems        =   9
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Index"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Main Body Part"
            Object.Width           =   3528
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "OPCS Code"
            Object.Width           =   3528
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   3
            Text            =   "OPCS Narrative"
            Object.Width           =   5292
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   4
            Text            =   "Surgeon's Category"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   5
            Text            =   "Anaesthetist's Category"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   6
            Text            =   "Part"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   7
            Text            =   "Part A Cost"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(9) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   8
            Text            =   "Page No"
            Object.Width           =   2540
         EndProperty
      End
      Begin MSComctlLib.ListView lstLookUpList 
         Height          =   4560
         Left            =   -74760
         TabIndex        =   37
         Top             =   1800
         Width           =   11055
         _ExtentX        =   19500
         _ExtentY        =   8043
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
         NumItems        =   5
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Index"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Surgeon's Category"
            Object.Width           =   3528
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "Surgeon's Fees"
            Object.Width           =   3528
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   3
            Text            =   "Anaesthetist's Category"
            Object.Width           =   3528
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   4
            Text            =   "Anaethetist's Fees"
            Object.Width           =   3528
         EndProperty
      End
      Begin VB.Label Label20 
         Alignment       =   2  'Center
         Caption         =   "LOOK UP LIST"
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
         Left            =   -74760
         TabIndex        =   38
         Top             =   1560
         Width           =   1395
      End
      Begin VB.Label Label4 
         Alignment       =   2  'Center
         Caption         =   "PROCEDURES LIST"
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
         Left            =   240
         TabIndex        =   33
         Top             =   1920
         Width           =   1635
      End
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00404040&
      Caption         =   "Malaysian Medical Association "
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
      TabIndex        =   44
      Top             =   0
      Width           =   11805
   End
   Begin VB.Label Label33 
      Caption         =   "Bold - Searchable        * - Mandatory Fields"
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
      TabIndex        =   24
      Top             =   7680
      Width           =   4095
   End
   Begin VB.Label Label6 
      Caption         =   "TOTAL RECORDS"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   375
      Left            =   5160
      TabIndex        =   23
      Top             =   7680
      Width           =   1455
   End
   Begin VB.Label txtTotalRecord 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00C0FFFF&
      Caption         =   "0"
      Enabled         =   0   'False
      Height          =   255
      Left            =   6600
      TabIndex        =   22
      Top             =   7680
      Width           =   1815
   End
End
Attribute VB_Name = "frmMMA"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim EditFlag As Boolean
Dim bExit As Boolean
Public ProcCriteria As String
Public LookCriteria As String
Private m_blnDirection(1 To 11) As Boolean

Private Sub cboPart_Change()
    If InStr(cboPart.Text, "null") Then
       cboPart.Enabled = False
    End If
End Sub

Private Sub cboPart_KeyPress(KeyAscii As Integer)
Dim NewText As String
Dim ValidCount As Integer
Dim ValidValue As String
Dim i As Integer
    
    Clipboard.Clear
    If KeyAscii >= 32 And KeyAscii <> 127 Then
    NewText = LCase(Left(cboPart.Text, cboPart.SelStart) + Chr(KeyAscii) + Mid(cboPart.Text, cboPart.SelStart + cboPart.SelLength + 1))

    ValidCount = 0
    ValidValue = ""

    For i = 0 To cboPart.ListCount - 1
 
        If NewText = LCase(Left(cboPart.List(i), Len(NewText))) Then
                ValidCount = ValidCount + 1
                ValidValue = cboPart.List(i)
        End If
        Next
        If ValidCount <= 1 Then KeyAscii = 0
             If ValidCount = 1 Then
               cboPart.Text = ValidValue
               cboPart.SelStart = 0
               cboPart.SelLength = Len(ValidValue)
             End If
        End If
End Sub

Private Sub CmdPrint_Click()
Screen.MousePointer = vbHourglass
    
    Select Case SSTab1.Tab
        Case 0
            If lstProceduresList.listitems.Count <> 0 Then
                Call ExportListViewtoExcel(lstProceduresList)
            Else
                MsgBox "Report cannot be printed without any record.", vbInformation
            End If
        Case 1
            If lstLookUpList.listitems.Count <> 0 Then
                Call ExportListViewtoExcel(lstLookUpList)
            Else
                MsgBox "Report cannot be printed without any record.", vbInformation
            End If
    End Select
    
    cmdAdd.Enabled = True
    cmdEdit.Enabled = False
    Screen.MousePointer = Default
End Sub

Private Sub Form_Activate()
    SSTab1.Tab = 0
    CmdRefresh.Enabled = True
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

Private Sub Form_KeyPress(KeyAscii As Integer)
    Call EnterToTab(KeyAscii)
End Sub

Private Sub Form_Load()
Call ProcComboListing
End Sub

'********************Start Add, Delete, Edit, Save, Search, Cancel buttons******************

Private Sub cmdAdd_Click()
EditFlag = False
Call ClearForm(Me)
Call DisableCmdButton(Me)
Call EntriesEnable(True)
cmdSave.Enabled = True
cmdAdd.Enabled = True
    Select Case SSTab1.Tab
        Case 0
            txtMainBP.SetFocus
        Case 1
            txtSurgeonCat.SetFocus
    End Select

End Sub

Private Sub CmdDelete_Click()
Screen.MousePointer = vbHourglass
    'medixmasterconnMaint.Open "File Name=" & BOSPath & "Conn\MediConnHISMaint.UDL"
    If SSTab1.Tab = 0 Then
        Call DeleteProc
    ElseIf SSTab1.Tab = 1 Then
        Call DeleteLook
    End If
    'medixmasterconnMaint.Close
    'Set medixmasterconnMaint = Nothing

cmdAdd.Enabled = True
Screen.MousePointer = Default
End Sub

Private Sub cmdEdit_Click()
    EditFlag = True
    Call EntriesEnable(True)
    txtMainBP.SetFocus
    cmdSave.Enabled = True
    cmdDelete.Enabled = False
    
    If txtMainBP <> "" Then
        Call EntriesEnable(True)
    End If
    
    If txtSurgeonCat <> "" Then
        Call EntriesEnable(True)
    End If
    
End Sub

Private Sub cmdSave_Click()
    Screen.MousePointer = vbHourglass
   ' medixmasterconnMaint.Open "File Name=" & BOSPath & "Conn\MediConnHISMaint.UDL"

    Select Case SSTab1.Tab
        Case 0
            Call SaveProc
        Case 1
            Call SaveLook
    End Select
    cmdAdd.Enabled = True
    cmdEdit.Enabled = False
    Screen.MousePointer = Default
 '   medixmasterconnMaint.Close
 '   Set medixmasterconnMaint = Nothing
End Sub

Private Sub cmdSearch_Click()

    Screen.MousePointer = vbHourglass
    cmdSearch.Enabled = False
   ' medixmasterconnMaint.Open "File Name=" & BOSPath & "Conn\MediConnHISMaint.UDL"

    Select Case SSTab1.Tab
         Case 0
             Call SearchProc
             txtMainBP.Enabled = True
             txtOPCSCode.Enabled = True
             cboPart.Enabled = True
             txtPageNo.Enabled = True
             txtOPCSNarrative.Enabled = False
         Case 1
             Call SearchLook
         End Select
   '  medixmasterconnMaint.Close
   '  Set medixmasterconnMaint = Nothing
    cmdSearch.Enabled = True
    Screen.MousePointer = Default
    cmdAdd.Enabled = True
End Sub

Private Sub CmdRefresh_Click()

Call ClearForm(Me)
Call EntriesEnable(True)

Select Case SSTab1.Tab

    Case 0
        Call ProcComboListing
               
End Select
End Sub

Private Sub CmdClear_Click()
    Call ClearForm(Me)
    lstProceduresList.listitems.Clear
    lstLookUpList.listitems.Clear
    txtTotalRecord = 0
    cmdEdit.Enabled = False
    Call ProcComboListing
End Sub

Private Sub cmdCancel_Click()
        Unload Me
End Sub
'********************End Add, Delete, Edit, Save, Search, Cancel buttons******************

'**************************Start EntriesEnable******************************
Private Sub EntriesEnable(status As Boolean)
    Dim ctl As Control

    For Each ctl In Me.Controls
        If TypeOf ctl Is TextBox Then
            ctl.Enabled = status
        ElseIf TypeOf ctl Is ComboBox Then
            ctl.Enabled = status
        ElseIf TypeOf ctl Is MaskEdBox Then
            ctl.Enabled = status
        End If
    Next ctl
End Sub
'**************************Start EntriesEnable******************************

'**********************************Start Procedures Part**************************************
Private Sub SaveProc()
Dim RecordSaved
Dim Proc As New ADODB.Recordset
Dim SaveProcedures As New ADODB.Recordset
Dim strsql As String
Dim ProcSeq As String
        
    If txtMainBP = "" Then
        MsgBox "Please Enter Main Body Part", vbOKOnly + vbCritical, DialogBoxTitle
        txtMainBP.SetFocus
    ElseIf txtOPCSNarrative = "" Then
        MsgBox "Please Enter OPCS Narrative", vbOKOnly + vbCritical, DialogBoxTitle
        txtOPCSNarrative.SetFocus
    ElseIf txtPageNo = "" Then
        MsgBox "Please Enter Page No", vbOKOnly + vbCritical, DialogBoxTitle
        txtPageNo.SetFocus
    Else
      RecordSaved = MsgBox("Do you want to save the record?", vbYesNo + vbExclamation)
        If RecordSaved = vbYes Then
            If EditFlag = False Then
                Proc.Open "MMAProcedures", medixmasterconnMaint, adOpenKeyset, adLockOptimistic
                Proc.AddNew
            Else
                strsql = "Select * From MMAProcedures Where ProcIndex = '" & ChkStr(lstProceduresList.SelectedItem.Text) & "'"
                Proc.Open strsql, medixmasterconnMaint, adOpenKeyset, adLockOptimistic
            End If
            
                With Proc
                    !ProcMainBPCode = ChkStr(txtMainBP)
                    !OPCSCode = IIf(Len(txtOPCSCode), ChkStr(txtOPCSCode), Null)
                    !OPCSNarrative = ChkStr(txtOPCSNarrative)
                    !ProcSurgCat = IIf(Len(txtSurgCat), ChkStr(txtSurgCat), Null)
                    !ProcAnaesCat = IIf(Len(txtAnaesCat), ChkStr(txtAnaesCat), Null)
                    !ProcPageNo = IIf(Len(txtPageNo), ChkStr(txtPageNo), Null)
                    !PartACost = IIf(Len(txtPartACost), ChkStr(txtPartACost), Null)
                    !ProcPart = Mid(cboPart, 1, 1)
                    !ProcLastUpdateUser = Logon
                    !ProcLastUpdateDate = Date
                End With
                Proc.Update
                Call ClearForm(Me)
                Call ProcListing
                If EditFlag = False Then
                    MsgBox "Record Saved...", vbOKOnly + vbInformation, DialogBoxTitle
                Else
                    MsgBox "Record Updated...", vbOKOnly + vbInformation, DialogBoxTitle
                End If
                
            Proc.Close
            Set SaveProcedures = Nothing
        End If
        EditFlag = False
    End If
'End If
End Sub

Private Sub DeleteProc()
Dim DeleteProc As New ADODB.Recordset
Dim rsChkProc As New ADODB.Recordset
Dim strsql, sqlProcedures As String
Dim DeleteRecord

Call EntriesEnable(True)
EditFlag = False
If txtMainBP = "" Then
   MsgBox "Please Enter Main Body Part", vbOKOnly + vbCritical, DialogBoxTitle
   txtMainBP.SetFocus
ElseIf txtOPCSCode = "" Then
   MsgBox "Please Enter OPCS Code", vbOKOnly + vbCritical, DialogBoxTitle
   txtOPCSCode.SetFocus
ElseIf lstProceduresList.listitems.Count = 0 Then
   MsgBox "Please search the record", vbOKOnly + vbCritical, DialogBoxTitle
'Else
'   sqlProcedures = "Select * from MMAProcedures where ProcIndex = '" & ChkStr(lstProceduresList.SelectedItem.SubItems(1)) & "'" & _
'   " AND OPCSCode = '" & ChkStr(lstProceduresList.SelectedItem.SubItems(2)) & "'"
   
'   rsChkProc.MaxRecords = 1
'   rsChkProc.Open sqlProcedures, medixmasterconnMaint, adOpenKeyset, adLockOptimistic
   
'   If rsChkProc.recordCount Then
'        MsgBox "This record cannot be deleted because it is used in Company Table!", vbOKOnly + vbCritical, "Error Message"
   Else
   DeleteRecord = MsgBox("Do you want to delete the record?", vbYesNo + vbExclamation)
     If DeleteRecord = vbYes Then ' And rsChkProc.recordCount = 0 Then
           strsql = "Delete From MMAProcedures Where ProcIndex = '" & ChkStr(lstProceduresList.SelectedItem.Text) & "'"
           DeleteProc.Open strsql, medixmasterconnMaint, adOpenKeyset, adLockOptimistic
           MsgBox "Record Deleted...", vbOKOnly + vbInformation, DialogBoxTitle
           cmdSave.Enabled = True
           Call ClearForm(Me)
           Call ProcListing
           Call DisableCmdButton(Me)
           Call EntriesEnable(True)
           txtMainBP.SetFocus
         End If
     End If

Exit Sub
End Sub

Private Sub ProcListing(Optional strAppend As String)
Dim rst2 As New ADODB.Recordset
Dim ProcedureMaster As ListItem
Dim sql As String
Dim TotalProcCount As Integer
    
    sql = "Select * from MMAProcedures where 0=0 "
   
    If Len(Trim(strAppend)) Then
        sql = sql + strAppend
    End If
              
    TotalProcCount = 0
    lstProceduresList.listitems.Clear
    
    rst2.Open sql, medixmasterconnMaint, adOpenForwardOnly, adLockBatchOptimistic
    If rst2.EOF = True Then
        MsgBox "Record not found", vbOKOnly + vbInformation, DialogBoxTitle
    Else
        Do Until rst2.EOF
            Set ProcedureMaster = lstProceduresList.listitems.Add(, , rst2!ProcIndex)
                ProcedureMaster.SubItems(1) = rst2!ProcMainBPCode
                If rst2!OPCSCode <> "" Then
                    ProcedureMaster.SubItems(2) = rst2!OPCSCode
                End If
                If rst2!OPCSNarrative <> "" Then
                    ProcedureMaster.SubItems(3) = rst2!OPCSNarrative
                End If
                If rst2!ProcSurgCat <> "" Then
                    ProcedureMaster.SubItems(4) = rst2!ProcSurgCat
                End If
                If rst2!ProcAnaesCat <> "" Then
                    ProcedureMaster.SubItems(5) = rst2!ProcAnaesCat
                End If
                If rst2!ProcPart <> "" Then
                    ProcedureMaster.SubItems(6) = rst2!ProcPart
                End If
                If rst2!PartACost <> "" Then
                    ProcedureMaster.SubItems(7) = rst2!PartACost
                End If
                If rst2!ProcPageNo <> "" Then
                    ProcedureMaster.SubItems(8) = rst2!ProcPageNo
                End If

                rst2.MoveNext
        TotalProcCount = TotalProcCount + 1
        txtTotalRecord = TotalProcCount
        Loop
    rst2.Close
    End If
End Sub

Private Sub ProcComboListing()
        
    cboPart.Clear
    cboPart.AddItem "A"
    cboPart.AddItem "B"
    cboPart.AddItem "C"
    
End Sub

Private Sub lstProceduresList_Click()
Dim rst3 As New ADODB.Recordset
Dim strsql As String
    
  '  medixmasterconnMaint.Open "File Name=" & BOSPath & "Conn\MediConnHISMaint.UDL"

    If lstProceduresList.listitems.Count <> 0 Then
        Call EnableCmdButton(Me)
        Call ClearForm(Me)
    
        strsql = "Select * from MMAProcedures Where ProcIndex = '" & lstProceduresList.SelectedItem.Text & "'"
        rst3.Open strsql, medixmasterconnMaint, adOpenKeyset, adLockOptimistic
        
        txtMainBP = rst3!ProcMainBPCode
        
        If Len(rst3!OPCSCode) Then
            txtOPCSCode = rst3!OPCSCode
        End If
        
        If Len(rst3!OPCSNarrative) Then
            txtOPCSNarrative = rst3!OPCSNarrative
        End If

        If Len(rst3!ProcSurgCat) Then
            txtSurgCat = rst3!ProcSurgCat
        End If
        If Len(rst3!ProcAnaesCat) Then
            txtAnaesCat = rst3!ProcAnaesCat
        End If
        If Len(rst3!ProcPageNo) Then
            txtPageNo = rst3!ProcPageNo
        End If
        If Len(rst3!ProcPart) Then
            cboPart = rst3!ProcPart
        End If
        If Len(rst3!PartACost) Then
            txtPartACost = rst3!PartACost
        End If
        
        Call EntriesEnable(False)
        cmdSave.Enabled = False
        rst3.Close
        Set rst3 = Nothing
    End If
   ' medixmasterconnMaint.Close
   ' Set medixmasterconnMaint = Nothing
End Sub

Private Sub lstProceduresList_KeyDown(KeyCode As Integer, Shift As Integer)
Dim rst3 As New ADODB.Recordset
Dim strsql As String
        
    If lstProceduresList.listitems.Count <> 0 Then
        Call EnableCmdButton(Me)
        Call ClearForm(Me)
    
        strsql = "Select * from MMAProcedures Where ProcIndex = '" & lstProceduresList.SelectedItem.Text & "'"
        rst3.Open strsql, medixmasterconnMaint, adOpenKeyset, adLockOptimistic
        
        txtMainBP = rst3!ProcMainBPCode
        
        If Len(rst3!OPCSCode) Then
            txtOPCSCode = rst3!OPCSCode
        End If
        
        If Len(rst3!OPCSNarrative) Then
            txtOPCSNarrative = rst3!OPCSNarrative
        End If

        If Len(rst3!ProcSurgCat) Then
            txtSurgCat = rst3!ProcSurgCat
        End If
        If Len(rst3!ProcAnaesCat) Then
            txtAnaesCat = rst3!ProcAnaesCat
        End If
        If Len(rst3!ProcPageNo) Then
            txtPageNo = rst3!ProcPageNo
        End If
        If Len(rst3!ProcPart) Then
            cboPart = rst3!ProcPart
        End If
        If Len(rst3!PartACost) Then
            txtPartACost = rst3!PartACost
        End If
        
        Call EntriesEnable(False)
        cmdSave.Enabled = False
        rst3.Close
        Set rst3 = Nothing
    End If
End Sub

Private Sub lstProceduresList_KeyUp(KeyCode As Integer, Shift As Integer)
Dim rst3 As New ADODB.Recordset
Dim strsql As String
        
    If lstProceduresList.listitems.Count <> 0 Then
        Call EnableCmdButton(Me)
        Call ClearForm(Me)
    
        strsql = "Select * from MMAProcedures Where ProcIndex = '" & lstProceduresList.SelectedItem.Text & "'"
        rst3.Open strsql, medixmasterconnMaint, adOpenKeyset, adLockOptimistic
        
        txtMainBP = rst3!ProcMainBPCode
        
        If Len(rst3!OPCSCode) Then
            txtOPCSCode = rst3!OPCSCode
        End If
        
        If Len(rst3!OPCSNarrative) Then
            txtOPCSNarrative = rst3!OPCSNarrative
        End If

        If Len(rst3!ProcSurgCat) Then
            txtSurgCat = rst3!ProcSurgCat
        End If
        If Len(rst3!ProcAnaesCat) Then
            txtAnaesCat = rst3!ProcAnaesCat
        End If
        If Len(rst3!ProcPageNo) Then
            txtPageNo = rst3!ProcPageNo
        End If
        If Len(rst3!ProcPart) Then
            cboPart = rst3!ProcPart
        End If
        If Len(rst3!PartACost) Then
            txtPartACost = rst3!PartACost
        End If
        
        Call EntriesEnable(False)
        cmdSave.Enabled = False
        rst3.Close
        Set rst3 = Nothing
    End If
End Sub

Private Sub SearchProc()
Dim strAppend As String
   
    Screen.MousePointer = vbHourglass
    strAppend = ""
    
    If Len(Trim(txtMainBP)) Then
        strAppend = strAppend + " And ProcMainBPCode like '%" & ChkStr(txtMainBP) & "%'"
    End If
    
    If Len(Trim(txtOPCSCode)) Then
        strAppend = strAppend + " And OPCSCode like '%" & ChkStr(txtOPCSCode) & "%'"
    End If
    
    ProcCriteria = strAppend
    Call ProcListing(strAppend)
    Screen.MousePointer = Default
End Sub

Private Sub lstProceduresList_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)

lstProceduresList.SortKey = ColumnHeader.Index - 1
lstProceduresList.Sorted = True
    If lstProceduresList.SortOrder = lvwAscending Then
        lstProceduresList.SortOrder = lvwDescending
    Else
        lstProceduresList.SortOrder = lvwAscending
    End If
End Sub

'Private Sub PrintProcedures()
'    RptCompanyGrpList.show
'End Sub

'**********************************End Procedures Part**************************************

'**********************************Start Look Up Table**************************************
Private Sub SaveLook()
Dim RecordSaved
Dim LookUp As New ADODB.Recordset
Dim SaveLook As New ADODB.Recordset
Dim strsql As String
Dim LookSeq As String
              
      RecordSaved = MsgBox("Do you want to save the record?", vbYesNo + vbExclamation)
        If RecordSaved = vbYes Then
            If EditFlag = False Then
                LookUp.Open "MMAFees", medixmasterconnMaint, adOpenKeyset, adLockOptimistic
                LookUp.AddNew
            Else
                strsql = "Select * From MMAFees Where FeeIndex = '" & ChkStr(lstLookUpList.SelectedItem.Text) & "'"
                LookUp.Open strsql, medixmasterconnMaint, adOpenKeyset, adLockOptimistic
            End If
            
                With LookUp
                    !SurgeonsCat = IIf(Len(txtSurgeonCat), ChkStr(txtSurgeonCat), Null)
                    !SurgeonsFees = IIf(Len(txtSurgeonFee), ChkStr(txtSurgeonFee), Null)
                    !AnaesthetistsCat = IIf(Len(txtAnaestheCat), ChkStr(txtAnaestheCat), Null)
                    !AnaesthetistsFees = IIf(Len(txtAnaestheFee), ChkStr(txtAnaestheFee), Null)
                    !FeeLastUpdateUser = Logon
                    !FeeLastUpdateDate = Date
                End With
                LookUp.Update
                Call ClearForm(Me)
                Call LookListing
                If EditFlag = False Then
                    MsgBox "Record Saved...", vbOKOnly + vbInformation, DialogBoxTitle
                Else
                    MsgBox "Record Updated...", vbOKOnly + vbInformation, DialogBoxTitle
                End If
                
            LookUp.Close
            Set SaveLook = Nothing
        End If
        EditFlag = False
End Sub

Private Sub DeleteLook()
Dim DeleteLookUp As New ADODB.Recordset
Dim rsChkLook As New ADODB.Recordset
Dim strsql, sqlLook As String
Dim DeleteRecord

EditFlag = False
'If txtMainBP = "" Then
'   MsgBox "Please Enter Main Body Part", vbOKOnly + vbCritical, DialogBoxTitle
'   txtMainBP.SetFocus
'ElseIf txtOPCSCode = "" Then
'   MsgBox "Please Enter OPCS Code", vbOKOnly + vbCritical, DialogBoxTitle
'   txtOPCSCode.SetFocus
If lstLookUpList.listitems.Count = 0 Then
   MsgBox "Please Enter The Record", vbOKOnly + vbCritical, DialogBoxTitle
Else
'   sqlProcedures = "Select * from MMAProcedures where ProcIndex = '" & ChkStr(lstProceduresList.SelectedItem.SubItems(1)) & "'" & _
'   " AND OPCSCode = '" & ChkStr(lstProceduresList.SelectedItem.SubItems(2)) & "'"
   
'   rsChkProc.MaxRecords = 1
'   rsChkProc.Open sqlProcedures, medixmasterconnMaint, adOpenKeyset, adLockOptimistic
   
'   If rsChkProc.recordCount Then
'        MsgBox "This record cannot be deleted because it is used in Company Table!", vbOKOnly + vbCritical, "Error Message"
   
   DeleteRecord = MsgBox("Do you want to delete the record?", vbYesNo + vbExclamation)
     If DeleteRecord = vbYes Then ' And rsChkProc.recordCount = 0 Then
           strsql = "Delete From MMAFees Where FeeIndex = '" & ChkStr(lstLookUpList.SelectedItem.Text) & "'"
           DeleteLookUp.Open strsql, medixmasterconnMaint, adOpenKeyset, adLockOptimistic
           MsgBox "Record Deleted...", vbOKOnly + vbInformation, DialogBoxTitle
           cmdSave.Enabled = True
           Call ClearForm(Me)
           Call LookListing
           Call DisableCmdButton(Me)
           Call EntriesEnable(True)
           txtSurgeonCat.SetFocus
         End If
End If
Exit Sub
End Sub

Private Sub LookListing(Optional strAppend As String)
Dim rst2 As New ADODB.Recordset
Dim LookMaster As ListItem
Dim sql As String
Dim TotalLookCount As Integer
    
    sql = "Select * from MMAFees where 0=0 "
   
    If Len(Trim(strAppend)) Then
        sql = sql + strAppend
    End If
              
    TotalLookCount = 0
    lstLookUpList.listitems.Clear
    
    rst2.Open sql, medixmasterconnMaint, adOpenForwardOnly, adLockBatchOptimistic
    If rst2.EOF = True Then
        MsgBox "Record not found", vbOKOnly + vbInformation, DialogBoxTitle
    Else
        Do Until rst2.EOF
            Set LookMaster = lstLookUpList.listitems.Add(, , rst2!FeeIndex)
                If rst2!SurgeonsCat <> "" Then
                    LookMaster.SubItems(1) = rst2!SurgeonsCat
                End If
                If rst2!SurgeonsFees <> "" Then
                    LookMaster.SubItems(2) = rst2!SurgeonsFees
                End If
                If rst2!AnaesthetistsCat <> "" Then
                    LookMaster.SubItems(3) = rst2!AnaesthetistsCat
                End If
                If rst2!AnaesthetistsFees <> "" Then
                    LookMaster.SubItems(4) = rst2!AnaesthetistsFees
                End If
                rst2.MoveNext
        TotalLookCount = TotalLookCount + 1
        txtTotalRecord = TotalLookCount
        Loop
    rst2.Close
    End If
End Sub


Private Sub lstLookUpList_Click()
Dim rst3 As New ADODB.Recordset
Dim strsql As String

   ' medixmasterconnMaint.Open "File Name=" & BOSPath & "Conn\MediConnHISMaint.UDL"
    If lstLookUpList.listitems.Count <> 0 Then
        Call EnableCmdButton(Me)
        Call ClearForm(Me)
    
        strsql = "Select * from MMAFees Where FeeIndex = '" & lstLookUpList.SelectedItem.Text & "'"
        rst3.Open strsql, medixmasterconnMaint, adOpenKeyset, adLockOptimistic
        
        If Len(rst3!SurgeonsCat) Then
            txtSurgeonCat = rst3!SurgeonsCat
        End If
        
        If Len(rst3!SurgeonsFees) Then
            txtSurgeonFee = rst3!SurgeonsFees
        End If

        If Len(rst3!AnaesthetistsCat) Then
            txtAnaestheCat = rst3!AnaesthetistsCat
        End If
        If Len(rst3!AnaesthetistsFees) Then
            txtAnaestheFee = rst3!AnaesthetistsFees
        End If
        
        Call EntriesEnable(False)
        cmdSave.Enabled = False
        rst3.Close
        Set rst3 = Nothing
    End If
  '  medixmasterconnMaint.Close
  '  Set medixmasterconnMaint = Nothing
End Sub

Private Sub lstLookUpList_KeyDown(KeyCode As Integer, Shift As Integer)
Dim rst3 As New ADODB.Recordset
Dim strsql As String
        
    If lstLookUpList.listitems.Count <> 0 Then
        Call EnableCmdButton(Me)
        Call ClearForm(Me)
    
        strsql = "Select * from MMAFees Where FeeIndex = '" & lstLookUpList.SelectedItem.Text & "'"
        rst3.Open strsql, medixmasterconnMaint, adOpenKeyset, adLockOptimistic
        
        If Len(rst3!SurgeonsCat) Then
            txtSurgeonCat = rst3!SurgeonsCat
        End If
        
        If Len(rst3!SurgeonsFees) Then
            txtSurgeonFee = rst3!SurgeonsFees
        End If

        If Len(rst3!AnaesthetistsCat) Then
            txtAnaestheCat = rst3!AnaesthetistsCat
        End If
        If Len(rst3!AnaesthetistsFees) Then
            txtAnaestheFee = rst3!AnaesthetistsFees
        End If
        
        Call EntriesEnable(False)
        cmdSave.Enabled = False
        rst3.Close
        Set rst3 = Nothing
    End If
End Sub

Private Sub lstLookUpList_KeyUp(KeyCode As Integer, Shift As Integer)
Dim rst3 As New ADODB.Recordset
Dim strsql As String
        
    If lstLookUpList.listitems.Count <> 0 Then
        Call EnableCmdButton(Me)
        Call ClearForm(Me)
    
        strsql = "Select * from MMAFees Where FeeIndex = '" & lstLookUpList.SelectedItem.Text & "'"
        rst3.Open strsql, medixmasterconnMaint, adOpenKeyset, adLockOptimistic
        
        If Len(rst3!SurgeonsCat) Then
            txtSurgeonCat = rst3!SurgeonsCat
        End If
        
        If Len(rst3!SurgeonsFees) Then
            txtSurgeonFee = rst3!SurgeonsFees
        End If

        If Len(rst3!AnaesthetistsCat) Then
            txtAnaestheCat = rst3!AnaesthetistsCat
        End If
        If Len(rst3!AnaesthetistsFees) Then
            txtAnaestheFee = rst3!AnaesthetistsFees
        End If
        
        Call EntriesEnable(False)
        cmdSave.Enabled = False
        rst3.Close
        Set rst3 = Nothing
    End If
End Sub

Private Sub SearchLook()
Dim strAppend As String
   
    Screen.MousePointer = vbHourglass
    strAppend = ""
    
    If Len(Trim(txtSurgeonCat)) Then
        strAppend = strAppend + " And SurgeonsCat like '%" & ChkStr(txtSurgeonCat) & "%'"
    End If
    
    If Len(Trim(txtAnaestheCat)) Then
        strAppend = strAppend + " And AnaesthetistsCat like '%" & ChkStr(txtAnaestheCat) & "%'"
    End If
    
    LookCriteria = strAppend
    Call LookListing(strAppend)
    Screen.MousePointer = Default
End Sub

Private Sub lstLookUpList_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)

lstLookUpList.SortKey = ColumnHeader.Index - 1
lstLookUpList.Sorted = True
    If lstLookUpList.SortOrder = lvwAscending Then
        lstLookUpList.SortOrder = lvwDescending
    Else
        lstLookUpList.SortOrder = lvwAscending
    End If
End Sub

'Private Sub PrintLook()
'    RptCompanyGrpList.show
'End Sub

'**********************************End Look Up Table**************************************

Private Sub txtPartACost_KeyPress(KeyAscii As Integer)
KeyAscii = KeyNumber(KeyAscii)
End Sub

Private Sub txtPageNo_KeyPress(KeyAscii As Integer)
KeyAscii = KeyNumber(KeyAscii)
End Sub

Private Sub txtAnaestheFee_KeyPress(KeyAscii As Integer)
KeyAscii = KeyNumber(KeyAscii)
End Sub

Private Sub txtSurgeonFee_KeyPress(KeyAscii As Integer)
KeyAscii = KeyNumber(KeyAscii)
End Sub

Private Sub SSTab1_Click(PreviousTab As Integer)
    cmdEdit.Enabled = False
    cmdDelete.Enabled = False
End Sub

