VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "tabctl32.ocx"
Begin VB.Form frmMaintenanceMembers 
   Caption         =   "Members Maintenance"
   ClientHeight    =   7590
   ClientLeft      =   2220
   ClientTop       =   1875
   ClientWidth     =   8985
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   8595
   ScaleMode       =   0  'User
   ScaleWidth      =   18668.58
   WindowState     =   2  'Maximized
   Begin VB.CommandButton CmdExit 
      Cancel          =   -1  'True
      Caption         =   "E&xit"
      Height          =   555
      Left            =   7920
      TabIndex        =   8
      Top             =   7080
      Width           =   930
   End
   Begin VB.CommandButton CmdUpdate 
      Caption         =   "&Update"
      Height          =   555
      Left            =   6720
      TabIndex        =   7
      Top             =   7080
      Width           =   930
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   7965
      Left            =   120
      TabIndex        =   0
      Top             =   256
      Width           =   11775
      _ExtentX        =   20770
      _ExtentY        =   14049
      _Version        =   393216
      Tabs            =   1
      TabsPerRow      =   1
      TabHeight       =   561
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      TabCaption(0)   =   "Valid Date Range"
      TabPicture(0)   =   "frmMaintenanceMembers.frx":0000
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Frame2"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Frame1"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "Frame3"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).ControlCount=   3
      Begin VB.Frame Frame3 
         Caption         =   "Birth Date"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1455
         Left            =   2160
         TabIndex        =   15
         Top             =   4200
         Width           =   6525
         Begin MSMask.MaskEdBox txtMBMDOBStartDate 
            BeginProperty DataFormat 
               Type            =   1
               Format          =   "dd/MM/yyyy"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1033
               SubFormatType   =   3
            EndProperty
            Height          =   375
            Left            =   1560
            TabIndex        =   5
            Top             =   600
            Width           =   1095
            _ExtentX        =   1931
            _ExtentY        =   661
            _Version        =   393216
            MaxLength       =   10
            Mask            =   "##/##/####"
            PromptChar      =   "_"
         End
         Begin MSMask.MaskEdBox txtMBMDOBEndDate 
            BeginProperty DataFormat 
               Type            =   1
               Format          =   "dd/MM/yyyy"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1033
               SubFormatType   =   3
            EndProperty
            Height          =   375
            Left            =   4800
            TabIndex        =   6
            Top             =   480
            Width           =   1095
            _ExtentX        =   1931
            _ExtentY        =   661
            _Version        =   393216
            MaxLength       =   10
            Mask            =   "##/##/####"
            PromptChar      =   "_"
         End
         Begin VB.Label Label6 
            Caption         =   "Start Date :"
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
            Left            =   360
            TabIndex        =   17
            Top             =   600
            Width           =   2055
         End
         Begin VB.Label Label3 
            Caption         =   "End Date :"
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
            Left            =   3720
            TabIndex        =   16
            Top             =   600
            Width           =   1815
         End
      End
      Begin VB.Frame Frame1 
         Caption         =   "Effective Date"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1455
         Left            =   2160
         TabIndex        =   12
         Top             =   600
         Width           =   6525
         Begin MSMask.MaskEdBox txtMBMEffStartDate 
            BeginProperty DataFormat 
               Type            =   1
               Format          =   "dd/MM/yyyy"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1033
               SubFormatType   =   3
            EndProperty
            Height          =   375
            Left            =   1560
            TabIndex        =   1
            Top             =   600
            Width           =   1095
            _ExtentX        =   1931
            _ExtentY        =   661
            _Version        =   393216
            MaxLength       =   10
            Mask            =   "##/##/####"
            PromptChar      =   "_"
         End
         Begin MSMask.MaskEdBox txtMBMEffEndDate 
            BeginProperty DataFormat 
               Type            =   1
               Format          =   "dd/MM/yyyy"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1033
               SubFormatType   =   3
            EndProperty
            Height          =   375
            Left            =   4800
            TabIndex        =   2
            Top             =   600
            Width           =   1095
            _ExtentX        =   1931
            _ExtentY        =   661
            _Version        =   393216
            MaxLength       =   10
            Mask            =   "##/##/####"
            PromptChar      =   "_"
         End
         Begin VB.Label Label1 
            Caption         =   "Start Date :"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   360
            TabIndex        =   14
            Top             =   600
            Width           =   1815
         End
         Begin VB.Label Label2 
            Caption         =   "End Date :"
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
            Left            =   3720
            TabIndex        =   13
            Top             =   600
            Width           =   1815
         End
      End
      Begin VB.Frame Frame2 
         Caption         =   "Expiry Date"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1455
         Left            =   2160
         TabIndex        =   9
         Top             =   2400
         Width           =   6525
         Begin MSMask.MaskEdBox txtMBMExpStartDate 
            BeginProperty DataFormat 
               Type            =   1
               Format          =   "dd/MM/yyyy"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1033
               SubFormatType   =   3
            EndProperty
            Height          =   375
            Left            =   1560
            TabIndex        =   3
            Top             =   600
            Width           =   1095
            _ExtentX        =   1931
            _ExtentY        =   661
            _Version        =   393216
            MaxLength       =   10
            Mask            =   "##/##/####"
            PromptChar      =   "_"
         End
         Begin MSMask.MaskEdBox txtMBMExpEndDate 
            BeginProperty DataFormat 
               Type            =   1
               Format          =   "dd/MM/yyyy"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1033
               SubFormatType   =   3
            EndProperty
            Height          =   375
            Left            =   4800
            TabIndex        =   4
            Top             =   600
            Width           =   1095
            _ExtentX        =   1931
            _ExtentY        =   661
            _Version        =   393216
            MaxLength       =   10
            Mask            =   "##/##/####"
            PromptChar      =   "_"
         End
         Begin VB.Label Label4 
            Caption         =   "End Date :"
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
            Left            =   3720
            TabIndex        =   11
            Top             =   600
            Width           =   1815
         End
         Begin VB.Label Label5 
            Caption         =   "Start Date :"
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
            Left            =   360
            TabIndex        =   10
            Top             =   600
            Width           =   1815
         End
      End
   End
   Begin VB.Label Label7 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00404040&
      Caption         =   "Date Maintenance"
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
      Height          =   199
      Left            =   0
      TabIndex        =   18
      Top             =   0
      Width           =   11805
   End
End
Attribute VB_Name = "frmMaintenanceMembers"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'Private lngFormWidth As Long
'Private lngFormHeight As Long
Dim bExit As Boolean

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

'******************** Start Exit, Update Button **********************************************
Private Sub cmdExit_Click()
'    Dim RecordExit
'   RecordExit = MsgBox("Quit this form?", vbYesNo + vbExclamation)
'    If RecordExit = vbYes Then
'        bExit = True
        Unload Me
'    End If
End Sub

Private Sub cmdUpdate_Click()
Dim RecordSaved
Dim MBMValidDate As New ADODB.Recordset
Dim SaveValidDate As New ADODB.Recordset
Dim strsql As String
   ' medixmasterconn.Open "File Name=" & BOSPath & "Conn\mediconnect.UDL"

    If txtMBMEffStartDate = "__/__/____" Then
        MsgBox "Please Enter Effective Start Date", vbOKOnly + vbCritical, DialogBoxTitle
        txtMBMEffStartDate.SetFocus
    ElseIf txtMBMEffEndDate = "__/__/____" Then
        MsgBox "Please Enter Effective End Date", vbOKOnly + vbCritical, DialogBoxTitle
        txtMBMEffEndDate.SetFocus
    ElseIf txtMBMExpStartDate = "__/__/____" Then
        MsgBox "Please Enter Expiry Start Date", vbOKOnly + vbCritical, DialogBoxTitle
        txtMBMExpStartDate.SetFocus
    ElseIf txtMBMExpEndDate = "__/__/____" Then
        MsgBox "Please Enter Expiry End Date", vbOKOnly + vbCritical, DialogBoxTitle
        txtMBMExpEndDate.SetFocus
    ElseIf txtMBMDOBStartDate = "__/__/____" Then
        MsgBox "Please Enter DOB Start Date", vbOKOnly + vbCritical, DialogBoxTitle
        txtMBMDOBStartDate.SetFocus
    ElseIf txtMBMDOBEndDate = "__/__/____" Then
        MsgBox "Please Enter DOB End Date", vbOKOnly + vbCritical, DialogBoxTitle
        txtMBMDOBEndDate.SetFocus
    Else
        strsql = "Select ValidID  From MBMValidDateRange "
 
        MBMValidDate.Open strsql, medixmasterconn, adOpenKeyset, adLockOptimistic
        RecordSaved = MsgBox("Do you want to save the record?", vbYesNo + vbExclamation)
        If RecordSaved = vbYes Then
            SaveValidDate.Open "MBMValidDateRange", medixmasterconn, adOpenKeyset, adLockOptimistic
            If MBMValidDate.recordCount = 0 Then
                SaveValidDate.AddNew
            End If
            
            With SaveValidDate
                !MBMEffStart = txtMBMEffStartDate
                !MBMEffEnd = txtMBMEffEndDate
                !MBMExpStart = txtMBMExpStartDate
                !MBMExpEnd = txtMBMExpEndDate
                !MBMDOBStart = txtMBMDOBStartDate
                !MBMDOBEnd = txtMBMDOBEndDate
                !MBMValidDateLastUpdateUser = Logon
                !MBMValidDateLastUpdateDate = Date
                .Update
            End With
            MsgBox "Record Saved...", vbOKOnly + vbInformation, DialogBoxTitle
      
            SaveValidDate.Close
            Set SaveValidDate = Nothing
        End If
        MBMValidDate.Close
        Set MBMValidDate = Nothing
    '    medixmasterconn.Close
    '    Set medixmasterconn = Nothing

    End If
End Sub
'******************** End Exit, Update Button **********************************************

'******************** Start Load Date **********************************************
Private Sub LoadDate()
    Dim MBMValidDate As New ADODB.Recordset
    Dim sql As String
    
    sql = "Select ValidID, MBMEffStart, MBMEffEnd, MBMExpStart, " & _
          "MBMExpEnd, MBMDOBStart, MBMDOBEnd " & _
          "from MBMValidDateRange "
          
    MBMValidDate.Open sql, medixmasterconn, adOpenKeyset, adLockOptimistic
    If MBMValidDate.recordCount = 0 Then
        MsgBox "Record not found", vbOKOnly + vbInformation, DialogBoxTitle
    Else
        'LoadDate = 1
            With MBMValidDate
                'txtMBMEffStartDate.Mask = ""
                txtMBMEffStartDate = Trim(!MBMEffStart)
                'txtMBMEffStartDate.Mask = "__/__/____"
                'txtMBMEffEndDate.Mask = ""
                txtMBMEffEndDate = Trim(!MBMEffEnd)
                'txtMBMEffEndDate.Mask = "__/__/____"
                'txtMBMExpStartDate.Mask = ""
                txtMBMExpStartDate = Trim(!MBMExpStart)
                'txtMBMExpStartDate.Mask = "__/__/____"
                'txtMBMExpEndDate.Mask = ""
                txtMBMExpEndDate = Trim(!MBMExpEnd)
                'txtMBMExpEndDate.Mask = "__/__/____"
                'txtMBMDOBStartDate.Mask = ""
                txtMBMDOBStartDate = Trim(!MBMDOBStart)
                'txtMBMDOBStartDate.Mask = "__/__/____"
                'txtMBMDOBEndDate.Mask = ""
                txtMBMDOBEndDate = Trim(!MBMDOBEnd)
                'txtMBMDOBEndDate.Mask = "__/__/____"
            End With
        MBMValidDate.Close
        Set MBMValidDate = Nothing
    End If
End Sub
'******************** End Load Date **********************************************


Private Sub Form_Load()
 '   medixmasterconn.Open "File Name=" & BOSPath & "Conn\mediconnect.UDL"
    LoadDate
  '  medixmasterconn.Close
  '  Set medixmasterconn = Nothing

End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    Call EnterToTab(KeyAscii)
End Sub


Private Sub txtMBMEffStartDate_LostFocus()
    If IsDate(txtMBMEffStartDate) Then
    Else
        If txtMBMEffStartDate <> "__/__/____" Then
            MsgBox "Invalid Date Format! ", vbCritical
            txtMBMEffStartDate.SetFocus
        End If
    End If
End Sub


Private Sub txtMBMEffEndDate_LostFocus()
    If IsDate(txtMBMEffEndDate) Then
    Else
        If txtMBMEffEndDate <> "__/__/____" Then
            MsgBox "Invalid Date Format! ", vbCritical
            txtMBMEffEndDate.SetFocus
        End If
    End If
End Sub

Private Sub txtMBMExpStartDate_LostFocus()
    If IsDate(txtMBMExpStartDate) Then
    Else
        If txtMBMExpStartDate <> "__/__/____" Then
            MsgBox "Invalid Date Format! ", vbCritical
            txtMBMExpStartDate.SetFocus
        End If
    End If
End Sub

Private Sub txtMBMExpEndDate_LostFocus()
    If IsDate(txtMBMExpEndDate) Then
    Else
        If txtMBMExpEndDate <> "__/__/____" Then
            MsgBox "Invalid Date Format! ", vbCritical
            txtMBMExpEndDate.SetFocus
        End If
    End If
End Sub


Private Sub txtMBMDOBStartDate_LostFocus()
    If IsDate(txtMBMDOBStartDate) Then
    Else
        If txtMBMDOBStartDate <> "__/__/____" Then
            MsgBox "Invalid Date Format! ", vbCritical
            txtMBMDOBStartDate.SetFocus
        End If
    End If
End Sub

Private Sub txtMBMDOBEndDate_LostFocus()
    If IsDate(txtMBMDOBEndDate) Then
    Else
        If txtMBMDOBEndDate <> "__/__/____" Then
            MsgBox "Invalid Date Format! ", vbCritical
            txtMBMDOBEndDate.SetFocus
        End If
    End If
End Sub
