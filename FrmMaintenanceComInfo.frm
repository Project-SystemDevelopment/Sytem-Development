VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "tabctl32.ocx"
Begin VB.Form FrmMaintenanceComInfo 
   Caption         =   "Company Information"
   ClientHeight    =   6840
   ClientLeft      =   2160
   ClientTop       =   1980
   ClientWidth     =   7920
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   MDIChild        =   -1  'True
   ScaleHeight     =   27557.29
   ScaleMode       =   0  'User
   ScaleWidth      =   59665.14
   WindowState     =   2  'Maximized
   Begin TabDlg.SSTab SSTab1 
      Height          =   7095
      Left            =   120
      TabIndex        =   10
      Top             =   240
      Width           =   11685
      _ExtentX        =   20611
      _ExtentY        =   12515
      _Version        =   393216
      Tabs            =   1
      TabsPerRow      =   1
      TabHeight       =   741
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      TabCaption(0)   =   "MediExpress (M) Sdn Bhd"
      TabPicture(0)   =   "FrmMaintenanceComInfo.frx":0000
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Frame1"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "CmdUpdate"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "CmdExit"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).ControlCount=   3
      Begin VB.CommandButton CmdExit 
         Cancel          =   -1  'True
         Caption         =   "E&xit"
         Height          =   555
         Left            =   9720
         TabIndex        =   9
         Top             =   5880
         Width           =   1050
      End
      Begin VB.CommandButton CmdUpdate 
         Caption         =   "&Update"
         Height          =   555
         Left            =   8400
         TabIndex        =   8
         Top             =   5880
         Width           =   1050
      End
      Begin VB.Frame Frame1 
         Height          =   4815
         Left            =   840
         TabIndex        =   11
         Top             =   480
         Width           =   10215
         Begin VB.TextBox TxtComFaxNo 
            Height          =   285
            Left            =   2160
            MaxLength       =   15
            TabIndex        =   4
            Top             =   2640
            Width           =   1575
         End
         Begin VB.TextBox TxtComPhoneNo 
            Height          =   285
            Left            =   2160
            MaxLength       =   15
            TabIndex        =   3
            Top             =   2160
            Width           =   1575
         End
         Begin VB.TextBox TxtComCurrentYr 
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
            Left            =   2160
            MaxLength       =   4
            TabIndex        =   7
            Top             =   3960
            Width           =   675
         End
         Begin VB.TextBox txtComAddress 
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   855
            Left            =   2160
            MaxLength       =   200
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   2
            Top             =   1080
            Width           =   7035
         End
         Begin VB.TextBox txtComHomepage 
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
            Left            =   2160
            MaxLength       =   150
            TabIndex        =   5
            Top             =   3120
            Width           =   4155
         End
         Begin VB.TextBox txtComEmail 
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
            Left            =   2160
            MaxLength       =   150
            TabIndex        =   6
            Top             =   3480
            Width           =   4155
         End
         Begin VB.TextBox txtComName 
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
            Left            =   2160
            MaxLength       =   150
            TabIndex        =   1
            Top             =   600
            Width           =   4215
         End
         Begin VB.TextBox txtComCode 
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
            Left            =   2160
            MaxLength       =   2
            TabIndex        =   0
            Top             =   240
            Width           =   690
         End
         Begin VB.Label Label2 
            Caption         =   "Company Address"
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
            Left            =   480
            TabIndex        =   19
            Top             =   1080
            Width           =   1575
         End
         Begin VB.Label Label4 
            Caption         =   "Company Name"
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
            Left            =   480
            TabIndex        =   18
            Top             =   600
            Width           =   1335
         End
         Begin VB.Label Label1 
            Caption         =   "Company Code"
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
            Left            =   480
            TabIndex        =   17
            Top             =   240
            Width           =   1335
         End
         Begin VB.Label Label5 
            Caption         =   "Company Current Year"
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
            Left            =   480
            TabIndex        =   16
            Top             =   3960
            Width           =   1695
         End
         Begin VB.Label Label10 
            Caption         =   "Homepage"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   240
            Left            =   480
            TabIndex        =   15
            Top             =   3120
            Width           =   870
         End
         Begin VB.Label Label9 
            Caption         =   "Email"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   240
            Left            =   480
            TabIndex        =   14
            Top             =   3600
            Width           =   1230
         End
         Begin VB.Label Label7 
            Caption         =   "Telephone No."
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
            Left            =   480
            TabIndex        =   13
            Top             =   2160
            Width           =   1335
         End
         Begin VB.Label Label8 
            Caption         =   "Fax No."
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
            Left            =   480
            TabIndex        =   12
            Top             =   2640
            Width           =   1095
         End
      End
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00404040&
      Caption         =   "Company Information"
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
      Height          =   196
      Left            =   0
      TabIndex        =   20
      Top             =   0
      Width           =   11805
   End
End
Attribute VB_Name = "FrmMaintenanceComInfo"
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
'Dim RecordExit
'    RecordExit = MsgBox("Quit this form?", vbYesNo + vbExclamation)
'    If RecordExit = vbYes Then
'        bExit = True
        Unload Me
'    End If
End Sub

Private Sub cmdUpdate_Click()
Dim RecordSaved
Dim COM As New ADODB.Recordset
Dim SaveCOM As New ADODB.Recordset
Dim strsql As String
    'medixmasterconn.Open "File Name=" & BOSPath & "Conn\mediconnect.UDL"

    If txtComCode = "" Then
        MsgBox "Please Enter Company Code", vbOKOnly + vbCritical, DialogBoxTitle
        txtComCode.SetFocus
    ElseIf txtComName = "" Then
        MsgBox "Please Enter Company Name", vbOKOnly + vbCritical, DialogBoxTitle
        txtComName.SetFocus
    ElseIf TxtComCurrentYr = "" Then
        MsgBox "Please Enter Company Current Year", vbOKOnly + vbCritical, DialogBoxTitle
        TxtComCurrentYr.SetFocus
    Else
        strsql = "Select COMCode From Company"
 
        COM.Open strsql, medixmasterconn, adOpenKeyset, adLockOptimistic
        RecordSaved = MsgBox("Do you want to save the record?", vbYesNo + vbExclamation)
        If RecordSaved = vbYes Then
            SaveCOM.Open "Company", medixmasterconn, adOpenKeyset, adLockOptimistic
            If COM.recordCount = 0 Then
                SaveCOM.AddNew
            End If
            
            With SaveCOM
                !COMCode = ChkStr(txtComCode)
                !COMName = ChkStr(txtComName)
                !COMAddress = ChkStr(txtComAddress)
                !COMPhone = ChkStr(TxtComPhoneNo)
                !COMFax = ChkStr(TxtComFaxNo)
                !COMEmail = ChkStr(txtComEmail)
                !COMHomepage = ChkStr(txtComHomepage)
                !COMCurrentYear = TxtComCurrentYr
                !COMLastUpdateUser = Logon
                !COMLastUpdateDate = Date
                .Update
            End With
            MsgBox "Record Saved...", vbOKOnly + vbInformation, DialogBoxTitle
      
            SaveCOM.Close
            Set SaveCOM = Nothing
        End If
        COM.Close
        Set COM = Nothing
       ' medixmasterconn.Close
       ' Set medixmasterconn = Nothing
    End If
End Sub
'******************** End Exit, Update Button **********************************************

'******************** Start Load Company Info **********************************************
Private Sub LoadCOMInfo()
    Dim COM As New ADODB.Recordset
    Dim sql As String
    
    sql = "Select COMCode, COMName, COMAddress, COMPhone, " & _
          "COMFax, COMEmail, COMHomepage, COMCurrentYear " & _
          "from Company "
          
    COM.Open sql, medixmasterconn, adOpenKeyset, adLockOptimistic
    If COM.recordCount = 0 Then
        MsgBox "Record not found", vbOKOnly + vbInformation, DialogBoxTitle
    Else
        'LoadDate = 1
            With COM
                
                txtComCode = Trim(ChkStr(!COMCode))
                txtComName = Trim(ChkStr(!COMName))
                txtComAddress = Trim(ChkStr(!COMAddress))
                TxtComPhoneNo = Trim(ChkStr(!COMPhone))
                TxtComFaxNo = Trim(ChkStr(!COMFax))
                txtComEmail = Trim(ChkStr(!COMEmail))
                txtComHomepage = Trim(ChkStr(!COMHomepage))
                TxtComCurrentYr = Trim(!COMCurrentYear)
            End With
        COM.Close
        Set COM = Nothing
    End If
End Sub
'******************** End Load Company Info **********************************************

'******************** Start Form **********************************************
Private Sub Form_Activate()
    txtComCode.SetFocus
End Sub

Private Sub Form_Load()
    'medixmasterconn.Open "File Name=" & BOSPath & "Conn\mediconnect.UDL"
    LoadCOMInfo
    'medixmasterconn.Close
    'Set medixmasterconn = Nothing

End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    Call EnterToTab(KeyAscii)
End Sub

Private Sub txtComPhoneNo_KeyPress(KeyAscii As Integer)
KeyAscii = KeyNumber(KeyAscii)
End Sub

Private Sub TxtComFaxNo_KeyPress(KeyAscii As Integer)
KeyAscii = KeyNumber(KeyAscii)
End Sub

Private Sub TxtComCurrentYr_KeyPress(KeyAscii As Integer)
KeyAscii = KeyNumber(KeyAscii)
End Sub

