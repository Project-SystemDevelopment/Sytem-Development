VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form FrmTempCaseTransfer 
   Caption         =   "Temporary Case Transfer"
   ClientHeight    =   3225
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   11370
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   3225
   ScaleWidth      =   11370
   WindowState     =   2  'Maximized
   Begin VB.Frame Frame1 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2895
      Left            =   0
      TabIndex        =   7
      Top             =   240
      Width           =   11415
      Begin VB.Frame Frame2 
         Caption         =   "Temp Member"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1575
         Left            =   120
         TabIndex        =   26
         Top             =   720
         Width           =   5535
         Begin VB.TextBox txtCovID1 
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
            Left            =   4200
            TabIndex        =   27
            Top             =   360
            Width           =   1215
         End
         Begin VB.TextBox txtMBMNo1 
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
            Left            =   1080
            TabIndex        =   28
            Top             =   360
            Width           =   2295
         End
         Begin VB.TextBox txtPolicyNo1 
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
            Left            =   1080
            TabIndex        =   29
            Top             =   620
            Width           =   2295
         End
         Begin VB.TextBox txtICNo1 
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
            Left            =   1080
            TabIndex        =   30
            Top             =   860
            Width           =   2295
         End
         Begin VB.TextBox txtMBMName1 
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
            Left            =   1080
            TabIndex        =   31
            Top             =   1100
            Width           =   2295
         End
         Begin VB.ComboBox cboPolStatus1 
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
            Left            =   4200
            TabIndex        =   32
            Top             =   600
            Width           =   1215
         End
         Begin MSMask.MaskEdBox txtEffDate1 
            Height          =   315
            Left            =   4200
            TabIndex        =   33
            Top             =   840
            Width           =   1215
            _ExtentX        =   2143
            _ExtentY        =   556
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
         Begin MSMask.MaskEdBox txtExpDate1 
            Height          =   315
            Left            =   4200
            TabIndex        =   34
            Top             =   1125
            Width           =   1215
            _ExtentX        =   2143
            _ExtentY        =   556
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
         Begin VB.Label Label3 
            Caption         =   "Member No."
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
            TabIndex        =   42
            Top             =   360
            Width           =   975
         End
         Begin VB.Label Label4 
            Caption         =   "Cov ID"
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
            Left            =   3480
            TabIndex        =   41
            Top             =   360
            Width           =   615
         End
         Begin VB.Label Label5 
            Caption         =   "Patient"
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
            TabIndex        =   40
            Top             =   1120
            Width           =   975
         End
         Begin VB.Label Label55 
            Caption         =   "Pol Status"
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
            Left            =   3480
            TabIndex        =   39
            Top             =   615
            Width           =   735
         End
         Begin VB.Label Label28 
            Caption         =   "Eff Date"
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
            Left            =   3480
            TabIndex        =   38
            Top             =   900
            Width           =   675
         End
         Begin VB.Label Label11 
            Caption         =   "Exp Date"
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
            Left            =   3480
            TabIndex        =   37
            Top             =   1155
            Width           =   750
         End
         Begin VB.Label Label9 
            Caption         =   "Policy No."
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
            TabIndex        =   36
            Top             =   620
            Width           =   975
         End
         Begin VB.Label Label14 
            Caption         =   "IC No."
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
            TabIndex        =   35
            Top             =   880
            Width           =   975
         End
      End
      Begin VB.Frame Frame4 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   530
         Left            =   120
         TabIndex        =   25
         Top             =   2280
         Width           =   11175
         Begin VB.CommandButton CmdCancel 
            Caption         =   "&Cancel"
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
            Left            =   9240
            TabIndex        =   5
            Top             =   160
            Width           =   735
         End
         Begin VB.CommandButton cmdSave 
            Caption         =   "&Transfer"
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
            Left            =   8400
            TabIndex        =   4
            Top             =   160
            Width           =   855
         End
         Begin VB.CommandButton CmdExit 
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
            Height          =   285
            Left            =   9960
            TabIndex        =   6
            Top             =   160
            Width           =   735
         End
      End
      Begin VB.Frame Frame5 
         Height          =   530
         Left            =   120
         TabIndex        =   23
         Top             =   120
         Width           =   11175
         Begin VB.TextBox txtNewCaseNo 
            Enabled         =   0   'False
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
            Left            =   6720
            TabIndex        =   44
            Top             =   160
            Width           =   2295
         End
         Begin VB.TextBox txtCaseNo 
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
            Left            =   1080
            TabIndex        =   0
            Top             =   160
            Width           =   2295
         End
         Begin VB.CommandButton CmdGo 
            Caption         =   "&Go"
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
            Height          =   285
            Left            =   3480
            TabIndex        =   1
            Top             =   160
            Width           =   735
         End
         Begin VB.Label Label17 
            Caption         =   "Case No."
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
            Left            =   5760
            TabIndex        =   45
            Top             =   195
            Width           =   1095
         End
         Begin VB.Label Label2 
            Caption         =   "Temp Case "
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
            Top             =   195
            Width           =   855
         End
      End
      Begin VB.Frame FrameTransfer 
         Caption         =   "Transfer Member"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1575
         Left            =   5760
         TabIndex        =   8
         Top             =   720
         Width           =   5535
         Begin VB.TextBox txtCovID2 
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
            Left            =   4200
            TabIndex        =   3
            Top             =   360
            Width           =   1215
         End
         Begin VB.TextBox txtMBMNo2 
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
            Left            =   1080
            TabIndex        =   2
            Top             =   360
            Width           =   2295
         End
         Begin VB.TextBox txtPolicyNo2 
            Enabled         =   0   'False
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
            Left            =   1080
            TabIndex        =   9
            Top             =   620
            Width           =   2295
         End
         Begin VB.TextBox txtICNo2 
            Enabled         =   0   'False
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
            Left            =   1080
            TabIndex        =   10
            Top             =   860
            Width           =   2295
         End
         Begin VB.TextBox txtMBMName2 
            Enabled         =   0   'False
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
            Left            =   1080
            TabIndex        =   11
            Top             =   1100
            Width           =   2295
         End
         Begin VB.ComboBox CboPolStatus2 
            Enabled         =   0   'False
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
            Left            =   4200
            TabIndex        =   12
            Top             =   600
            Width           =   1215
         End
         Begin MSMask.MaskEdBox txtEffDate2 
            Height          =   315
            Left            =   4200
            TabIndex        =   13
            Top             =   840
            Width           =   1215
            _ExtentX        =   2143
            _ExtentY        =   556
            _Version        =   393216
            Enabled         =   0   'False
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
         Begin MSMask.MaskEdBox txtExpDate2 
            Height          =   315
            Left            =   4200
            TabIndex        =   14
            Top             =   1125
            Width           =   1215
            _ExtentX        =   2143
            _ExtentY        =   556
            _Version        =   393216
            Enabled         =   0   'False
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
         Begin VB.Label Label6 
            Caption         =   "IC No."
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
            TabIndex        =   22
            Top             =   880
            Width           =   975
         End
         Begin VB.Label Label7 
            Caption         =   "Policy No."
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
            TabIndex        =   21
            Top             =   620
            Width           =   975
         End
         Begin VB.Label Label8 
            Caption         =   "Exp Date"
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
            Left            =   3480
            TabIndex        =   20
            Top             =   1155
            Width           =   750
         End
         Begin VB.Label Label10 
            Caption         =   "Eff Date"
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
            Left            =   3480
            TabIndex        =   19
            Top             =   900
            Width           =   675
         End
         Begin VB.Label Label12 
            Caption         =   "Pol Status"
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
            Left            =   3480
            TabIndex        =   18
            Top             =   615
            Width           =   735
         End
         Begin VB.Label Label13 
            Caption         =   "Patient"
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
            TabIndex        =   17
            Top             =   1120
            Width           =   975
         End
         Begin VB.Label Label15 
            Caption         =   "Cov ID"
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
            Left            =   3480
            TabIndex        =   16
            Top             =   360
            Width           =   615
         End
         Begin VB.Label Label16 
            Caption         =   "Transfer No."
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
            TabIndex        =   15
            Top             =   360
            Width           =   975
         End
      End
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00404040&
      Caption         =   "Temporary Case Transfer "
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
      TabIndex        =   43
      Top             =   0
      Width           =   11895
   End
End
Attribute VB_Name = "FrmTempCaseTransfer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim SuppNotFound As Boolean
Dim CaseNumber As String
Dim Check As Boolean

Private Sub cmdCancel_Click()
    Call ClearForm(Me)
    CaseNumber = ""
End Sub

Private Sub cmdExit_Click()
    Unload Me
End Sub

Private Sub CmdGo_Click()
Dim tempMBMNo As String
Dim TempCovID As String
Dim TempPOlicy As String
Dim TempStatus As String
Dim TempIcNo As String
Dim TempMBMName As String
Dim TempEffdate As String
Dim TempExpDAte As String
Dim CASXSql As String
Dim CasXRec As New ADODB.Recordset
Dim CLMRec As New ADODB.Recordset
Dim ClmSql As String
Dim Bordxt, CaseNotFound, MultiCase As Boolean
Dim cmd As New ADODB.Command
Dim cmd1 As New ADODB.Command
Dim cmd2 As New ADODB.Command
Dim Prm1 As ADODB.Parameter
Dim Prm2 As ADODB.Parameter
Dim strAppendCase As String
Dim strAppend As String

    txtCaseNo = ChkStr(txtCaseNo)
    Bordxt = False
    CaseNotFound = False
    MultiCase = False
    tempMBMNo = ""
    TempCovID = ""
    CaseNumber = ""
    CmdGo.Enabled = False
    cmdSave.Enabled = False
    'medixmasterconnMaint.Open "File Name=" & BOSPath & "Conn\MediConnHISMaint.UDL"

'--- chk for case not found
    'CASXSql = "SELECT MBMNumber, MBMPatientCovID from CASCrossReference WHERE CASNumber='" & Trim(txtCaseNo) & "'"
    'CasXRec.Open CASXSql, medixmasterconnMaint, adOpenForwardOnly, adLockReadOnly
    strAppendCase = "1"
    strAppend = " AND CASNumber='" & Trim(txtCaseNo) & "'"
    Set cmd.ActiveConnection = medixmasterconnMaint
    cmd.CommandText = "Case_TempTransfer"
    cmd.CommandType = adCmdStoredProc
    cmd.CommandTimeout = 300
    Set Prm1 = cmd.CreateParameter("StrAppend", adVarChar, adParamInput, 2000, strAppend)
    cmd.Parameters.Append Prm1
    Set Prm2 = cmd.CreateParameter("StrAppendCase", adVarChar, adParamInput, 2000, strAppendCase)
    cmd.Parameters.Append Prm2
    CasXRec.CursorLocation = adUseClient
    CasXRec.CursorType = adOpenForwardOnly
    CasXRec.CacheSize = 100
    Set CasXRec = cmd.Execute
        
    If CasXRec.EOF Then
        CaseNotFound = True
    Else
        tempMBMNo = CasXRec!MBMNumber
        TempCovID = CasXRec!MBMPatientCovID
    End If
    CasXRec.Close
    Set CasXRec = Nothing
' --- chk for how many cases under 1 membership
    'CASXSql = "SELECT MBMNumber from CASCrossReference WHERE MBMNumber='" & Trim(txtMBMNo1) & _
              "' AND CASNumber <> '" & Trim(txtCaseNo) & "'"
    'CasXRec.Open CASXSql, medixmasterconnMaint, adOpenForwardOnly, adLockReadOnly
    If tempMBMNo <> "" Then
        strAppendCase = "2"
        strAppend = " AND MBMNumber='" & Trim(tempMBMNo) & "' AND CASNumber <> '" & Trim(txtCaseNo) & "'"
        Set cmd1.ActiveConnection = medixmasterconnMaint
        cmd1.CommandText = "Case_TempTransfer"
        cmd1.CommandType = adCmdStoredProc
        cmd1.CommandTimeout = 300
        Set Prm1 = cmd1.CreateParameter("StrAppend", adVarChar, adParamInput, 2000, strAppend)
        cmd1.Parameters.Append Prm1
        Set Prm2 = cmd1.CreateParameter("StrAppendCase", adVarChar, adParamInput, 2000, strAppendCase)
        cmd1.Parameters.Append Prm2
        CasXRec.CursorLocation = adUseClient
        CasXRec.CursorType = adOpenForwardOnly
        CasXRec.CacheSize = 100
        Set CasXRec = cmd1.Execute
        
        If CasXRec.EOF = False Then
            MultiCase = True
        End If
        CasXRec.Close
        Set CasXRec = Nothing
    End If
   ' medixmasterconnMaint.Close
    'Set medixmasterconnMaint = Nothing

'--- chk for bordxt
    'medixmasterconnWS.Open "File Name=" & BOSPath & "Conn\mediconnectWS.UDL"

    'ClmSql = "SELECT CASNumber, BordxRefNo from Claim WHERE CASNumber='" & Trim(txtCaseNo) & "'"
    'ClmRec.Open ClmSql, medixmasterconnWS, adOpenForwardOnly, adLockReadOnly
'    If ClmRec.EOF = False Then
    strAppendCase = "1"
    strAppend = " AND CASNumber='" & Trim(txtCaseNo) & "'"
    Set cmd2.ActiveConnection = medixmasterconnWS
    cmd2.CommandText = "Case_TempTransfer"
    cmd2.CommandType = adCmdStoredProc
    cmd2.CommandTimeout = 300
    Set Prm1 = cmd2.CreateParameter("StrAppend", adVarChar, adParamInput, 2000, strAppend)
    cmd2.Parameters.Append Prm1
    Set Prm2 = cmd2.CreateParameter("StrAppendCase", adVarChar, adParamInput, 2000, strAppendCase)
    cmd2.Parameters.Append Prm2
    CLMRec.CursorLocation = adUseClient
    CLMRec.CursorType = adOpenForwardOnly
    CLMRec.CacheSize = 100
    Set CLMRec = cmd2.Execute
    
    Do While Not CLMRec.EOF
        If Len(Trim(CLMRec!BordxRefNo)) Then
            Bordxt = True
            Exit Do
        End If
        CLMRec.MoveNext
    Loop
    CLMRec.Close
    Set CLMRec = Nothing
    'medixmasterconnWS.Close
   ' Set medixmasterconnWS = Nothing

    If Trim(txtCaseNo) = "" Then
        MsgBox "Please keyin Case Number ! ", vbCritical, DialogBoxTitle
        txtCaseNo.SetFocus
    ElseIf CaseNotFound Then
        MsgBox "No Record found ! ", vbCritical, DialogBoxTitle
        txtCaseNo.SetFocus
    ElseIf Mid(txtCaseNo, 1, 1) <> "Z" Then
        MsgBox "This is not a Temporary Case Number ! ", vbCritical, DialogBoxTitle
        txtCaseNo.SetFocus
    ElseIf Mid(tempMBMNo, 1, 1) <> "Z" Then
        MsgBox "This Case is not registered under a Temporary Member ! ", vbCritical, DialogBoxTitle
        txtCaseNo.SetFocus
    ElseIf Bordxt Then
        MsgBox "This Case has been bordx ! ", vbCritical, DialogBoxTitle
        txtCaseNo.SetFocus
'    ElseIf MultiCase Then
'        MsgBox "Other case is reigetered under this membership ! ", vbCritical, DialogBoxTitle
'        txtCaseNo.SetFocus
    Else
        Call MBMInfo(tempMBMNo, TempCovID, TempPOlicy, TempStatus, TempIcNo, TempMBMName, TempEffdate, TempExpDAte)
        txtMBMNo1 = tempMBMNo
        txtCovID1 = TempCovID
        
        txtPolicyNo1 = TempPOlicy
        If Trim(TempStatus) = "A" Then
            cboPolStatus1.ListIndex = 0
        ElseIf Trim(TempStatus) = "C" Then
            cboPolStatus1.ListIndex = 1
        ElseIf Trim(TempStatus) = "D" Then
            cboPolStatus1.ListIndex = 2
        End If
        txtICNo1 = TempIcNo
        txtMBMName1 = TempMBMName
        txtEffDate1 = IIf(IsDate(TempEffdate), TempEffdate, "__/__/____")
        txtExpDate1 = IIf(IsDate(TempExpDAte), TempExpDAte, "__/__/____")
        txtMBMNo2 = tempMBMNo
        txtCovID2 = TempCovID
        Call DisplayTransMBMInfo
        cmdSave.Enabled = True
        FrameTransfer.Enabled = True
        txtMBMNo2.SetFocus
    End If
    CmdGo.Enabled = True
    cmdSave.Enabled = True
End Sub

Private Sub CmdSave_Click()
    CmdGo.Enabled = False
    cmdSave.Enabled = False

    If Trim(txtMBMNo2) = "" Then
        MsgBox "Please Keyin transfer Member No. ! ", vbCritical + vbOKOnly, DialogBoxTitle
        txtMBMNo2.SetFocus
    ElseIf Trim(txtCovID2) = "" Then
        MsgBox "Please Keyin transfer Cover ID ! ", vbCritical + vbOKOnly, DialogBoxTitle
        txtCovID2.SetFocus
    ElseIf Mid(txtMBMNo2, 1, 1) = "Z" Then
        MsgBox "Cannot transfer to the Temporary Membership ! ", vbCritical + vbOKOnly, DialogBoxTitle
        txtMBMNo2.SetFocus
    ElseIf Trim(txtPolicyNo2) = "" Then
        MsgBox "Member No. not found ! ", vbCritical, DialogBoxTitle
        txtMBMNo2.SetFocus
    ElseIf Trim(txtMBMName2) = "" Then
        MsgBox "Patient not found ! ", vbCritical, DialogBoxTitle
        txtMBMNo2.SetFocus
    ElseIf Mid(CboPolStatus2, 1, 1) = "D" Then
        MsgBox "Member status - 'DELETE'  ! ", vbCritical + vbOKOnly, DialogBoxTitle
        txtMBMNo2.SetFocus
    ElseIf SuppNotFound Then
        MsgBox "Invalid Cover ID ! Pls try again ! ", vbCritical
        txtCovID2.SetFocus
    ElseIf Trim(txtMBMNo1 & txtCovID1) = Trim(txtMBMNo2 & txtCovID2) Then
        MsgBox "Same Membership ! No transformation ! ", vbCritical
        txtMBMNo2.SetFocus
    Else
       ' medixmasterconn.Open "File Name=" & BOSPath & "Conn\mediconnect.UDL"
       ' medixmasterconnWS.Open "File Name=" & BOSPath & "Conn\mediconnectWS.UDL"
       ' medixmasterconnMaint.Open "File Name=" & BOSPath & "Conn\MediConnHISMaint.UDL"
        
        Call TransferProcess
        
       ' medixmasterconnMaint.Close
      '  medixmasterconn.Close
      '  medixmasterconnWS.Close
       ' Set medixmasterconnMaint = Nothing
       ' Set medixmasterconn = Nothing
       ' Set medixmasterconnWS = Nothing
    End If
    CmdGo.Enabled = True
    cmdSave.Enabled = True

End Sub

Private Sub Form_Load()
    cboPolStatus1.Clear
    cboPolStatus1.AddItem "A - Active"
    cboPolStatus1.AddItem "C - Cancelled"
    cboPolStatus1.AddItem "D - Delete"
    CboPolStatus2.Clear
    CboPolStatus2.AddItem "A - Active"
    CboPolStatus2.AddItem "C - Cancelled"
    CboPolStatus2.AddItem "D - Delete"
    SuppNotFound = False
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

Private Sub TxtCaseNo_LostFocus()
    txtCaseNo = ChkStr(txtCaseNo)
End Sub

Private Function MBMInfo(tempMBMNo, TempCovID, TempPOlicy, TempStatus, TempIcNo, TempMBMName, TempEffdate, TempExpDAte As String)
Dim MBMSql As String
Dim MBMRec As New ADODB.Recordset
Dim SuppSql As String
Dim SuppRec As New ADODB.Recordset
Dim cmd As New ADODB.Command
Dim cmd1 As New ADODB.Command
Dim Prm1 As ADODB.Parameter
Dim Prm2 As ADODB.Parameter
Dim strAppendCase As String
Dim strAppend As String

      '  medixmasterconnMaint.Open "File Name=" & BOSPath & "Conn\MediConnHISMaint.UDL"

        'MBMSql = "SELECT MBMNumber, MBMName, MBMCOVID, MBMPayorEffDate, MBMPayorEXPDate, MBMPolicyNo, MBMIcBcPp, MBMStatus FROM MBMCrossReference " & _
                " WHERE MBMNumber='" & Trim(TempMBMno) & "'"
        'MBMRec.Open MBMSql, medixmasterconnMaint, adOpenKeyset, adLockOptimistic
        strAppendCase = "3"
        strAppend = " AND MBMNumber='" & Trim(tempMBMNo) & "'"
        Set cmd.ActiveConnection = medixmasterconnMaint
        cmd.CommandText = "Case_TempTransfer"
        cmd.CommandType = adCmdStoredProc
        cmd.CommandTimeout = 300
        Set Prm1 = cmd.CreateParameter("StrAppend", adVarChar, adParamInput, 2000, strAppend)
        cmd.Parameters.Append Prm1
        Set Prm2 = cmd.CreateParameter("StrAppendCase", adVarChar, adParamInput, 2000, strAppendCase)
        cmd.Parameters.Append Prm2
        MBMRec.CursorLocation = adUseClient
        MBMRec.CursorType = adOpenForwardOnly
        MBMRec.CacheSize = 100
        Set MBMRec = cmd.Execute
        
        If MBMRec.EOF Then
            MsgBox "Member not found ! ", vbCritical, DialogBoxTitle
        Else
            SuppNotFound = False
            TempPOlicy = MBMRec!MBMPolicyNo
            TempStatus = MBMRec!MBMStatus
            TempIcNo = MBMRec!MBMIcBcPp
            TempMBMName = MBMRec!MBMName
            TempEffdate = MBMRec!MBMPayorEffDate
            TempExpDAte = MBMRec!MBMPayorEXPDate
            If Trim(TempCovID) <> "00" Then
                'medixmasterconn.Open "File Name=" & BOSPath & "Conn\mediconnect.UDL"

                'SuppSql = "SELECT MBMCName, MBMCICBCPPNo FROM MBMCoveredPersons WHERE MBMCNumber='" & Trim(TempMBMno) & "' AND MBMCCoverID='" & Trim(TempCovID) & "'"
                'SuppRec.Open SuppSql, medixmasterconn, adOpenKeyset, adLockOptimistic
                strAppendCase = "1"
                strAppend = " AND MBMCNumber='" & Trim(tempMBMNo) & "' AND MBMCCoverID='" & Trim(TempCovID) & "'"
                Set cmd1.ActiveConnection = medixmasterconn
                cmd1.CommandText = "Case_TempTransfer"
                cmd1.CommandType = adCmdStoredProc
                cmd1.CommandTimeout = 300
                Set Prm1 = cmd1.CreateParameter("StrAppend", adVarChar, adParamInput, 2000, strAppend)
                cmd1.Parameters.Append Prm1
                Set Prm2 = cmd.CreateParameter("StrAppendCase", adVarChar, adParamInput, 2000, strAppendCase)
                cmd1.Parameters.Append Prm2
                SuppRec.CursorLocation = adUseClient
                SuppRec.CursorType = adOpenForwardOnly
                SuppRec.CacheSize = 100
                Set SuppRec = cmd1.Execute
        
                If SuppRec.EOF Then
                    MsgBox "Supplementary not found ! ", vbCritical, DialogBoxTitle
                    SuppNotFound = True
                    'txtCaseNo.SetFocus
                Else
                    TempMBMName = SuppRec!MBMCName
                    TempIcNo = SuppRec!MBMCICBCPPNo
                End If
                SuppRec.Close
                Set SuppRec = Nothing
               ' medixmasterconn.Close
               ' Set medixmasterconn = Nothing

            End If
        End If
        MBMRec.Close
        Set MBMRec = Nothing
  '  medixmasterconnMaint.Close
   ' Set medixmasterconnMaint = Nothing
End Function

Private Sub txtCovID2_LostFocus()
    If Trim(txtMBMNo2) <> "" And Trim(txtCovID2) <> "" Then
        Call DisplayTransMBMInfo
    End If
End Sub

Private Sub txtMBMNo2_LostFocus()
    txtMBMNo2 = ChkStr(txtMBMNo2)
    If Trim(txtMBMNo2) <> "" And Trim(txtCovID2) <> "" Then
        Call DisplayTransMBMInfo
    End If
End Sub

Private Sub DisplayTransMBMInfo()
Dim tempMBMNo As String
Dim TempCovID As String
Dim TempPOlicy As String
Dim TempStatus As String
Dim TempIcNo As String
Dim TempMBMName As String
Dim TempEffdate As String
Dim TempExpDAte As String

    tempMBMNo = txtMBMNo2
    TempCovID = txtCovID2
    Call MBMInfo(tempMBMNo, TempCovID, TempPOlicy, TempStatus, TempIcNo, TempMBMName, TempEffdate, TempExpDAte)
    txtMBMNo2 = tempMBMNo
    txtCovID2 = TempCovID
    txtPolicyNo2 = TempPOlicy
    If Trim(TempStatus) = "A" Then
        CboPolStatus2.ListIndex = 0
    ElseIf Trim(TempStatus) = "C" Then
        CboPolStatus2.ListIndex = 1
    ElseIf Trim(TempStatus) = "D" Then
        CboPolStatus2.ListIndex = 2
    End If
    txtICNo2 = TempIcNo
    txtMBMName2 = TempMBMName
    txtEffDate2 = IIf(IsDate(TempEffdate), TempEffdate, "__/__/____")
    txtExpDate2 = IIf(IsDate(TempExpDAte), TempExpDAte, "__/__/____")
End Sub

Private Sub TransferProcess()
On Error GoTo ErrorSave
Dim startTrans As Boolean
Dim CAsREc As New ADODB.Recordset
Dim CasXRec As New ADODB.Recordset
Dim CLMRec As New ADODB.Recordset
Dim ClmSql As String
Dim MBMRec As New ADODB.Recordset
Dim MBMSql As String
Dim SuppSql As String
Dim SuppRec As New ADODB.Recordset
Dim strSptLmt As String
Dim SptLmtInd As String
Dim SptPlnLmt As New ADODB.Recordset
Dim TmpCovID As String
Dim CasSql As String
Dim ClmTotApprove As Double
Dim ClmHNursing As Double
Dim ClmKidney As Double
Dim ClmCancer As Double
Dim tmpPln As String
Dim SuppALBStatus, BnfSql, MBMBNFSql As String
Dim tmpHLCode, tmpINSCode, TmpEffDate As String
Dim MBMBnf2 As New ADODB.Recordset
Dim BNFLimitRec As New ADODB.Recordset
Dim tmpTransRem As String
Dim tmpAvLmt As Double
Dim CASExcess As New ADODB.Recordset
Dim strsqlExcess As String

    Call CheckMBMCAS
    CasSql = " WHERE CASNumber = '" & Trim(txtCaseNo) & "'"
    Screen.MousePointer = vbHourglass
    medixmasterconn.beginTrans
    medixmasterconnMaint.beginTrans
    medixmasterconnWS.beginTrans
    
    CaseNumber = GenerateCaseNo(Mid(txtMBMNo2, 1, 2))  'Payor code
    txtNewCaseNo = CaseNumber
    
    CasXRec.Open "SELECT * FROM tblRecord where CaseNo = '" & Trim(txtCaseNo) & "'", medixmasterconnMaint, adOpenKeyset, adLockOptimistic
        If CasXRec.EOF = False Then
            CasXRec!CaseNo = CaseNumber
            CasXRec.Update
        End If
    CasXRec.Close
    Set CasXRec = Nothing
    
    
    CasXRec.Open "SELECT CASNumber, MBMNumber, MBMPatientCovID, TransRemarks FROM CASCrossReference " & CasSql, medixmasterconnMaint, adOpenKeyset, adLockOptimistic
    tmpTransRem = IIf(Len(CasXRec!TransRemarks), CasXRec!TransRemarks, "")
    tmpTransRem = tmpTransRem & "~TransCAS by : " & Logon & " Date : " & Format(Date, "yyyymmdd") & "," & txtCaseNo & "->" & txtNewCaseNo
    tmpTransRem = tmpTransRem & " , " & txtMBMNo1 & "-" & txtCovID1 & "->" & txtMBMNo2 & "-" & txtCovID2
    CasXRec!TransRemarks = tmpTransRem
    CasXRec!CasNumber = CaseNumber
    CasXRec!MBMNumber = Trim(txtMBMNo2)
    CasXRec!MBMPatientCovID = Trim(txtCovID2)
    
    CasXRec.Update
    CasXRec.Close
    Set CasXRec = Nothing
    
    CAsREc.Open "SELECT CASNumber, MBMNumber,MBMPatientCovID, CASLastUpdateUser, CASLastUpdateDate FROM CAS " & CasSql, medixmasterconn, adOpenKeyset, adLockOptimistic
    CAsREc!CasNumber = CaseNumber
    CAsREc!MBMNumber = Trim(txtMBMNo2)
    CAsREc!MBMPatientCovID = Trim(txtCovID2)
    CAsREc!CASLastUpdateUser = Logon
    CAsREc!CASLastUpdateDate = Date
    CAsREc.Update
    
    CAsREc.Close
    Set CAsREc = Nothing
    
    CAsREc.Open "SELECT CASNumber, MBMNumber,MBMCoveredID FROM CASAmendHis " & CasSql, medixmasterconn, adOpenKeyset, adLockOptimistic
            
    If Not CAsREc.EOF Then
        CAsREc!CasNumber = CaseNumber
        CAsREc!MBMNumber = Trim(txtMBMNo2)
        CAsREc!MBMCoveredID = Trim(txtCovID2)
        CAsREc.Update
        
    End If
    CAsREc.Close
    Set CAsREc = Nothing
    
    ClmTotApprove = 0
    ClmSql = "SELECT CasNumber, CLMTotalApproved  From Claim WHERE CASNumber = '" & Trim(txtCaseNo) & "'"
    CLMRec.Open ClmSql, medixmasterconnWS, adOpenKeyset, adLockOptimistic
        
    If Not CLMRec.EOF Then
        Do Until CLMRec.EOF
            CLMRec!CasNumber = CaseNumber
            ClmTotApprove = ClmTotApprove + CLMRec!CLMTotalApproved
            CLMRec.Update
            CLMRec.MoveNext
        Loop
    End If
    CLMRec.Close
    Set CLMRec = Nothing
    
    ClmCancer = 0
    ClmKidney = 0
    ClmHNursing = 0
    ClmSql = "SELECT CLMIndCancerTreatment, CLMIndKidneyDial, CLMHomeNCAmt From ClaimApprove WHERE CASNumber = '" & Trim(txtCaseNo) & "'"
    CLMRec.Open ClmSql, medixmasterconnWS, adOpenKeyset, adLockOptimistic
    If Not CLMRec.EOF Then
        Do Until CLMRec.EOF
            ClmCancer = ClmCancer + IIf(Len(CLMRec!CLMIndCancerTreatment), CLMRec!CLMIndCancerTreatment, 0)
            ClmKidney = ClmKidney + IIf(Len(CLMRec!CLMIndKidneyDial), CLMRec!CLMIndKidneyDial, 0)
            ClmHNursing = ClmHNursing + IIf(IsNumeric(CLMRec!CLMHomeNCAmt), CLMRec!CLMHomeNCAmt, 0)
            CLMRec.MoveNext
        Loop
    End If
    CLMRec.Close
    Set CLMRec = Nothing

    strsqlExcess = "SELECT CASNumber, EXCNumber  From CAS_Excess WHERE CASNumber = '" & Trim(txtCaseNo) & "'"
    CASExcess.Open strsqlExcess, medixmasterconnWS, adOpenKeyset, adLockOptimistic
        
    If Not CASExcess.EOF Then
        Do Until CASExcess.EOF
            CASExcess!CasNumber = CaseNumber
            'CASExcess!EXCNumber = "E" & CaseNumber & "-1"
            'EDIT BY M.SHAN 23/08/2010
            CASExcess!EXCNumber = Mid(CASExcess!EXCNumber, 1, 1) & CaseNumber & Mid(CASExcess!EXCNumber, Len(CASExcess!EXCNumber) - 1, 2)
            CASExcess.Update
            CASExcess.MoveNext
        Loop
    End If
    CASExcess.Close
    Set CASExcess = Nothing

    Call UpdateCASRecs
    Call UpdateClmRecs
    If Check = True Then
        Call DeleteMembership
    End If
    
    If ClmTotApprove <> 0 Then
            
        MBMSql = "SELECT HLTCode, INSCode, PLNCode, MBMPayorEffDate, MBMAvailableLimit FROM MBM " & _
                " WHERE MBMNumber='" & Trim(txtMBMNo2) & "'"
        MBMRec.Open MBMSql, medixmasterconn, adOpenKeyset, adLockOptimistic
        tmpHLCode = MBMRec!HLTCode
        tmpPln = MBMRec!PLNCode
        tmpINSCode = MBMRec!INSCode
        If Trim(MBMRec!MBMPayorEffDate) <> "" Then
            TmpEffDate = MBMRec!MBMPayorEffDate
        End If
        SuppALBStatus = "A"
        strSptLmt = "Select SuppLimitStatus, PLNCode from AnnualLimit where PLNCode = '" & Trim(MBMRec!PLNCode) & "'"
        SptPlnLmt.Open strSptLmt, medixmasterconn, adOpenKeyset, adLockOptimistic
                
        SptLmtInd = "A"
        If Trim(SptPlnLmt!SuppLimitStatus) <> "" Then
            SptLmtInd = Trim(SptPlnLmt!SuppLimitStatus)
        End If
        SptPlnLmt.Close
        Set SptPlnLmt = Nothing
        
        If SptLmtInd = "A" Or Trim(txtCovID2) = "00" Then
            tmpAvLmt = IIf(IsNumeric(MBMRec!MBMAvailableLimit), MBMRec!MBMAvailableLimit, 0)
            If tmpAvLmt > 0 Then
                MBMRec!MBMAvailableLimit = MBMRec!MBMAvailableLimit - ClmTotApprove
                MBMRec.Update
            End If
        Else
        
            TmpCovID = txtCovID2
            If SptLmtInd = "B" Then
                TmpCovID = "01"
            End If
            SuppSql = "Select MBMCAvailableLimit from MBMCoveredPersons " & _
                "WHERE MBMCNumber='" & Trim(txtMBMNo2) & "' AND MBMCCoverID='" & Trim(TmpCovID) & "'"
            SuppRec.Open SuppSql, medixmasterconn, adOpenKeyset, adLockOptimistic
            tmpAvLmt = IIf(IsNumeric(SuppRec!MBMCAvailableLimit), SuppRec!MBMCAvailableLimit, 0)
            If tmpAvLmt > 0 Then
                SuppRec!MBMCAvailableLimit = SuppRec!MBMCAvailableLimit - ClmTotApprove
                SuppRec.Update
            End If
            SuppRec.Close
            Set SuppRec = Nothing
        End If
        MBMRec.Close
        Set MBMRec = Nothing
    End If
    ' ----- Life Time Batch
    BnfSql = "Select SuppALBStatus as SuppBNFStatus from BnfAnnLmt Where INSCode ='" & Trim(tmpINSCode) & "' And PLNCode ='" & Trim(tmpPln) & "' And HLTCode = '" & Trim(tmpHLCode) & "' " & _
             " AND ALBEffDate <= '" & Format(TmpEffDate, "mm/dd/yyyy") & "' order by ALBEffDate desc"
    BNFLimitRec.Open BnfSql, medixmasterconn, adOpenForwardOnly, adLockReadOnly
    If BNFLimitRec.EOF = False Then
        SuppALBStatus = BNFLimitRec!SuppBNFStatus
    End If
    BNFLimitRec.Close
    Set BNFLimitRec = Nothing
    TmpCovID = txtCovID2
    If SuppALBStatus = "B" Then
        TmpCovID = "01"
    End If

    MBMBNFSql = "SELECT MBMCancerBalLmt, MBMKidneyBalLmt, MBMHomeNCBalLmt FROM MBMBnfLmt " & _
            " WHERE MBMNumber='" & Trim(txtMBMNo2) & "' and MBMCoveredID ='" & Trim(TmpCovID) & "'"
    MBMBnf2.Open MBMBNFSql, medixmasterconn, adOpenKeyset, adLockOptimistic
    If MBMBnf2.recordCount Then
        MBMBnf2!MBMCancerBalLmt = IIf(IsNumeric(MBMBnf2!MBMCancerBalLmt), MBMBnf2!MBMCancerBalLmt, 0) - ClmCancer
        MBMBnf2!MBMKidneyBalLmt = IIf(IsNumeric(MBMBnf2!MBMKidneyBalLmt), MBMBnf2!MBMKidneyBalLmt, 0) - ClmKidney
        MBMBnf2!MBMHomeNCBalLmt = IIf(IsNumeric(MBMBnf2!MBMHomeNCBalLmt), MBMBnf2!MBMHomeNCBalLmt, 0) - ClmHNursing
        MBMBnf2.Update
    End If
    MBMBnf2.Close
    Set MBMBnf2 = Nothing

    medixmasterconn.commitTrans
    medixmasterconnMaint.commitTrans
    medixmasterconnWS.commitTrans
    startTrans = False
        
    Screen.MousePointer = vbDefault
    MsgBox "Transfered Sucessfully..., Your Case No. is : " & CaseNumber, vbOKOnly + vbInformation, DialogBoxTitle
    Call ClearForm(Me)
    CaseNumber = ""
    Exit Sub
    
ErrorSave:
    Screen.MousePointer = vbDefault
    If startTrans = True Then
        medixmasterconnMaint.RollbackTrans
        medixmasterconn.RollbackTrans
        medixmasterconnWS.RollbackTrans
        MsgBox Err.Description
    End If
End Sub

Private Sub CheckMBMCAS()
Dim rst2 As New ADODB.Recordset
Dim cmd As New ADODB.Command
Dim Prm1 As ADODB.Parameter
Dim Prm2 As ADODB.Parameter
Dim strAppendCase As String
Dim strAppend As String

strAppendCase = "4"
strAppend = " AND MBMNumber = '" & Trim(txtMBMNo1) & "'"
Set cmd.ActiveConnection = medixmasterconnMaint
    cmd.CommandText = "Case_TempTransfer"
    cmd.CommandType = adCmdStoredProc
    cmd.CommandTimeout = 300
    Set Prm1 = cmd.CreateParameter("StrAppend", adVarChar, adParamInput, 2000, strAppend)
    cmd.Parameters.Append Prm1
    Set Prm2 = cmd.CreateParameter("StrAppendCase", adVarChar, adParamInput, 2000, strAppendCase)
    cmd.Parameters.Append Prm2
    rst2.CursorLocation = adUseClient
    rst2.CursorType = adOpenForwardOnly
    rst2.CacheSize = 100
    Set rst2 = cmd.Execute

    rst2.MoveNext
    If rst2.EOF Then
        Check = True
    End If
        
End Sub

Private Sub DeleteMembership()
Dim MBMSql As String
Dim SuppSql As String
Dim DelSql As String
Dim MBMRec As New ADODB.Recordset
    
    MBMSql = "WHERE MBMNumber='" & Trim(txtMBMNo1) & "'"
    SuppSql = "WHERE MBMCNumber='" & Trim(txtMBMNo1) & "'"
    
' ----- MBMCrossReference table
        DelSql = "Delete From MBMCrossReference " & MBMSql
        MBMRec.Open DelSql, medixmasterconnMaint, adOpenKeyset, adLockOptimistic
' ----- MBMLifeTime table
        DelSql = "Delete From MBMLifeTime " & MBMSql
        MBMRec.Open DelSql, medixmasterconn, adOpenKeyset, adLockOptimistic
' ----- MBM table
        DelSql = ""
        DelSql = "Delete From MBM " & MBMSql
        Call DeleteRec(DelSql)
' ----- MBMTwo table
        DelSql = ""
        DelSql = "Delete From MBMTwo " & MBMSql
        Call DeleteRec(DelSql)
' ----- MBMOthers table
        DelSql = ""
        DelSql = "Delete From MBMOthers " & MBMSql
        Call DeleteRec(DelSql)
' ----- MBMOthersTwo table
        DelSql = ""
        DelSql = "Delete From MBMOthersTWo " & MBMSql
        Call DeleteRec(DelSql)
' ----- MBMCLNOthers table
        DelSql = ""
        DelSql = "Delete From MBMCLNOthers " & MBMSql
        Call DeleteRec(DelSql)
' ----- MBMCoveredPersons table
        DelSql = ""
        DelSql = "Delete From MBMCoveredPersons " & SuppSql
        Call DeleteRec(DelSql)
' ----- MBMCoveredPersonsTwo table
        DelSql = ""
        DelSql = "Delete From MBMCoveredPersonsTwo " & SuppSql
        Call DeleteRec(DelSql)
' ----- MBMCoveredPersonsCLN
        DelSql = ""
        DelSql = "Delete From MBMCoveredPersonsCLN " & SuppSql
        Call DeleteRec(DelSql)
End Sub

Private Sub DeleteRec(strsql As String)
Dim MBMRec As New ADODB.Recordset
    MBMRec.Open strsql, medixmasterconn, adOpenKeyset, adLockOptimistic
End Sub

Private Sub UpdateCASRecs()
Dim CASSql1, CASSql2 As String
Dim UpdSql As String
Dim MBMRec As New ADODB.Recordset

    CASSql1 = "SELECT CASNumber From "
    CASSql2 = "WHERE CASNumber='" & Trim(txtCaseNo) & "'"
    
' ----- CASOthers table
        UpdSql = ""
        UpdSql = CASSql1 & "CASOthers " & CASSql2
        Call UpdHISDBRecProc(UpdSql)
' ----- CASOthersTwo Table
        UpdSql = ""
        UpdSql = CASSql1 & "CASOthersTwo " & CASSql2
        Call UpdHISDBRecProc(UpdSql)
' ----- CAS_UD table
        UpdSql = ""
        UpdSql = CASSql1 & "CAS_UD " & CASSql2
        Call UpdHISDBRecProc(UpdSql)
' ----- CASAuditReview Table
        UpdSql = ""
        UpdSql = CASSql1 & "CASAuditReview " & CASSql2
        Call UpdHISDBRecProc(UpdSql)
' ----- CASRemarks table
        UpdSql = ""
        UpdSql = CASSql1 & "CASRemarks " & CASSql2
        Call UpdHISDBRecProc(UpdSql)
' ----- CASRepudiation Table
        UpdSql = ""
        UpdSql = CASSql1 & "CASRepudiation " & CASSql2
        Call UpdHISDBRecProc(UpdSql)
' ----- CASProcInv table
        UpdSql = ""
        UpdSql = CASSql1 & "CASProcInv " & CASSql2
        Call UpdHISDBRecProc(UpdSql)
' ----- DIAFinal table
        UpdSql = ""
        UpdSql = CASSql1 & "DIAFinal " & CASSql2
        Call UpdHISDBRecProc(UpdSql)
' ----- DIAFirst table
        UpdSql = ""
        UpdSql = CASSql1 & "DIAFirst " & CASSql2
        Call UpdHISDBRecProc(UpdSql)
' ----- CASLG table
        UpdSql = ""
        UpdSql = "SELECT CASENo AS CASNumber from CASLG WHERE CASENo='" & Trim(txtCaseNo) & "'"
        Call UpdMaintRecProc(UpdSql)
' ----- LGSeq table
        UpdSql = ""
        UpdSql = CASSql1 & "LGSeq " & CASSql2
        Call UpdMaintRecProc(UpdSql)
End Sub

Private Sub UpdateClmRecs()
Dim CASSql1, CASSql2 As String
Dim UpdSql As String
Dim MBMRec As New ADODB.Recordset

    CASSql1 = "SELECT CASNumber From "
    CASSql2 = "WHERE CASNumber='" & Trim(txtCaseNo) & "'"
        
' ----- ClaimActual Table
        UpdSql = ""
        UpdSql = CASSql1 & "ClaimActual " & CASSql2
        Call UpdWSRecProc(UpdSql)
' ----- ClaimApprove Table
        UpdSql = ""
        UpdSql = CASSql1 & "ClaimApprove " & CASSql2
        Call UpdWSRecProc(UpdSql)
' ----- ClaimHOS Table
        UpdSql = ""
        UpdSql = CASSql1 & "ClaimHOS " & CASSql2
        Call UpdWSRecProc(UpdSql)
' ----- ClaimInvoice Table
        UpdSql = ""
        UpdSql = CASSql1 & "ClaimInvoice " & CASSql2
        Call UpdWSRecProc(UpdSql)
' ----- ClaimLimit Table
        UpdSql = ""
        UpdSql = CASSql1 & "ClaimLimit " & CASSql2
        Call UpdWSRecProc(UpdSql)
' ----- ClaimMember Table
        UpdSql = ""
        UpdSql = CASSql1 & "ClaimMember " & CASSql2
        Call UpdWSRecProc(UpdSql)
' ----- ClaimPayMBMStatus Table
        UpdSql = ""
        UpdSql = CASSql1 & "ClaimPayMBMStatus " & CASSql2
        Call UpdWSRecProc(UpdSql)
' ----- ClaimPosting Table
        UpdSql = ""
        UpdSql = CASSql1 & "ClaimPosting " & CASSql2
        Call UpdWSRecProc(UpdSql)
' ----- ClaimSameLimit Table
        UpdSql = ""
        UpdSql = CASSql1 & "ClaimSameLimit " & CASSql2
        Call UpdWSRecProc(UpdSql)
' ----- CLMReminder Table
        UpdSql = ""
        UpdSql = CASSql1 & "CLMReminder " & CASSql2
        Call UpdWSRecProc(UpdSql)
End Sub

Private Sub UpdHISDBRecProc(strsql As String)
Dim CAsREc As New ADODB.Recordset
    
    CAsREc.Open strsql, medixmasterconn, adOpenKeyset, adLockOptimistic
    Do While Not CAsREc.EOF
        CAsREc!CasNumber = Trim(txtNewCaseNo)
        CAsREc.Update
        CAsREc.MoveNext
    Loop
    CAsREc.Close
    Set CAsREc = Nothing
End Sub

Private Sub UpdWSRecProc(strsql As String)
Dim CLMRec As New ADODB.Recordset
    
    CLMRec.Open strsql, medixmasterconnWS, adOpenKeyset, adLockOptimistic
    Do While Not CLMRec.EOF
        CLMRec!CasNumber = Trim(txtNewCaseNo)
        CLMRec.Update
        CLMRec.MoveNext
    Loop
    CLMRec.Close
    Set CLMRec = Nothing
End Sub

Private Sub UpdMaintRecProc(strsql As String)
Dim CLMRec As New ADODB.Recordset
    
    CLMRec.Open strsql, medixmasterconnMaint, adOpenKeyset, adLockOptimistic
    Do While Not CLMRec.EOF
        CLMRec!CasNumber = Trim(txtNewCaseNo)
        CLMRec.Update
        CLMRec.MoveNext
    Loop
    CLMRec.Close
    Set CLMRec = Nothing
End Sub
