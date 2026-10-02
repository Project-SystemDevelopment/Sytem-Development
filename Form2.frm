VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmSeachMCOFee 
   Caption         =   "Search For MCO Fee"
   ClientHeight    =   5805
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   8340
   LinkTopic       =   "Form2"
   ScaleHeight     =   5805
   ScaleWidth      =   8340
   StartUpPosition =   3  'Windows Default
   Begin MSMask.MaskEdBox txtBordxDateTo 
      Height          =   285
      Left            =   4800
      TabIndex        =   13
      Top             =   960
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   503
      _Version        =   393216
      MaxLength       =   10
      Mask            =   "##/##/####"
      PromptChar      =   "_"
   End
   Begin VB.CommandButton cmdSearch 
      Caption         =   "&Search"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   3480
      TabIndex        =   10
      Top             =   1800
      Width           =   1335
   End
   Begin VB.CommandButton cmdReprint 
      Caption         =   "Reprint"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   4920
      TabIndex        =   8
      Top             =   5280
      Width           =   1335
   End
   Begin VB.CommandButton cmdPrint 
      Caption         =   "Print"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   3120
      TabIndex        =   7
      Top             =   5280
      Width           =   1335
   End
   Begin VB.Frame FrameListing 
      Caption         =   "Listing"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2775
      Left            =   360
      TabIndex        =   6
      Top             =   2640
      Width           =   7695
      Begin MSComctlLib.ListView ListView1 
         Height          =   2295
         Left            =   120
         TabIndex        =   14
         Top             =   240
         Width           =   7455
         _ExtentX        =   13150
         _ExtentY        =   4048
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
         NumItems        =   0
      End
   End
   Begin VB.CommandButton cmdClose 
      Caption         =   "C&lose"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   6600
      TabIndex        =   5
      Top             =   1800
      Width           =   1335
   End
   Begin VB.CommandButton cmdClear 
      Caption         =   "&Clear"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   5040
      TabIndex        =   4
      Top             =   1800
      Width           =   1335
   End
   Begin VB.Frame FrameSearch 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2295
      Left            =   360
      TabIndex        =   0
      Top             =   120
      Width           =   7695
      Begin MSMask.MaskEdBox txtBordxDateFrom 
         Height          =   285
         Left            =   2400
         TabIndex        =   12
         Top             =   840
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   503
         _Version        =   393216
         MaxLength       =   10
         Mask            =   "##/##/####"
         PromptChar      =   "_"
      End
      Begin VB.ComboBox comInsurercode 
         Height          =   315
         ItemData        =   "Form2.frx":0000
         Left            =   2400
         List            =   "Form2.frx":0002
         Sorted          =   -1  'True
         TabIndex        =   9
         Top             =   1200
         Width           =   5175
      End
      Begin VB.Label lblPayor 
         Caption         =   "Payor"
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
         Left            =   360
         TabIndex        =   11
         Top             =   1200
         Width           =   1935
      End
      Begin VB.Label lblSearch 
         Caption         =   "Enter criteria for searching"
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
         TabIndex        =   3
         Top             =   0
         Width           =   2295
      End
      Begin VB.Label lblBordxDateTo 
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
         Left            =   3960
         TabIndex        =   2
         Top             =   840
         Width           =   375
      End
      Begin VB.Label lblBordxDateFrom 
         Caption         =   "Bordereaux Date From"
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
         Left            =   360
         TabIndex        =   1
         Top             =   840
         Width           =   2055
      End
   End
End
Attribute VB_Name = "frmSeachMCOFee"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public SearchCriteria As String
Private Sub PayListing(Optional strAppend As String)
'On Error Resume Next
Dim rst2 As New ADODB.Recordset
Dim PayMaster As ListItem
Dim sql As String

ListView1.ListItems.Clear
    sql = "SELECT * FROM MBM"
        
    'If Len(strAppend) Then
    'sql = sql + strAppend
    'End If
    
    rst2.Open sql, masterconn, adOpenKeyset, adLockOptimistic
    If rst2.RecordCount = 0 Then
       MsgBox "Record not found", vbInformation + vbOKOnly
    Else
    Do Until rst2.EOF
        Set PayMaster = ListView1.ListItems.Add(, , rst2!PayCode)
            PayMaster.SubItems(1) = rst2!MBMNumber
            PayMaster.SubItems(2) = rst2!MBMBordxDate
            PayMaster.SubItems(3) = rst2!MBMPolicyNo
            PayMaster.SubItems(4) = rst2!MBMName
            PayMaster.SubItems(5) = rst2!RELCode
            PayMaster.SubItems(6) = rst2!MBMIcBcPp
            PayMaster.SubItems(7) = rst2!MBMGroupCompany
            PayMaster.SubItems(8) = rst2!MBMMCOFees
            rst2.MoveNext
        Loop
    
    rst2.Close
    Set rst2 = Nothing
    End If
End Sub

Private Sub cmdSearchListing(Optional strAppend As String)

strAppend = ""

If Len(comInsurercode) Then
strAppend = strAppend + " And PayCode like '%" & comInsurercode & "%'"
End If

'If Len(Trim(CblHealth)) Then
'strAppend = strAppend + " And HLTCode = '" & Mid(ChkStr(CblHealth), 1, 2) & "'"
'End If


SearchCriteria = strAppend
Call PayListing(strAppend)
    
End Sub

Private Sub cmdClear_Click()
mskBordxDateFrom = 0
mskBordxDateTo = 0
comInsurercode = ""
mskBordxDateFrom.Item
End Sub

Private Sub cmdClose_Click()

End Sub

Private Sub cmdSearch_Click()
Dim strSearch As String

'If txtBordxDateFrom = "__/__/____" Then
'    MsgBox "Please Enter Bordx Date", vbOKOnly + vbCritical, "Warning"
'    txtBordxDateFrom.SetFocus
'ElseIf IsDate(txtBordxDateFrom) = False Then
'    MsgBox "Please Enter a Valid Bordx Date", vbOKOnly + vbCritical, "Warning"
'    txtBordxDateFrom.SetFocus
'ElseIf txtBordxDateTo = "__/__/____" Then
'    MsgBox "Please Enter Bordx Date", vbOKOnly + vbCritical, "Warning"
'    txtBordxDateTo.SetFocus
'ElseIf IsDate(txtBordxDateTo) = False Then
'    MsgBox "Please Enter a Valid Bordx Date", vbOKOnly + vbCritical, "Warning"
'    txtBordxDateTo.SetFocus
'End If

Call cmdSearchListing
End Sub

Private Sub Form_Load()
Call ComboListing
End Sub
Sub ComboListing()
    Dim PAY As New ADODB.Recordset
    Dim sqlstr As String
   
    sqlstr = "select PAYCode , PAYCompanyName from PAY "
    
    PAY.Open sqlstr, masterconn, adOpenKeyset, adLockOptimistic
    
    Do Until PAY.EOF
        comInsurercode.AddItem (PAY!PayCode) & " - " & (PAY!PAYCompanyName)
        PAY.MoveNext
    Loop
    
    PAY.Close
    Set PAY = Nothing
End Sub

Private Sub ListView1_DblClick()
'On Error Resume Next
Dim rst3 As New ADODB.Recordset
Dim strsql As String
    
    'Call EnableCmdButton(Me)
    'Call ClearForm(Me)
    strsql = "SELECT MBMNumber, MBMBordxDate, MBMPolicyNo, MBMName, RELCode, MBMIcBcPp, MBMGroupCompany, MBMMCOFees " & _
    "from MBM Where PayCode = '" & ListView1.SelectedItem.Text & "'"
    rst3.Open strsql, masterconn, adOpenKeyset, adLockOptimistic


'TxtAgeCode = rst3!AGECode
'CblHealth = rst3!HLTCode
'TxtDesc = rst3!AGEDescription
'TxtAgeFr = rst3!agefrom
'TxtAgeTo = rst3!ageto
'Call EntriesEnable(False)
'CmdSave.Enabled = False
rst3.Close
End Sub
