VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmUploadBordxByIC 
   Caption         =   "Upload Bordx By IC"
   ClientHeight    =   7350
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   11685
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   7350
   ScaleWidth      =   11685
   WindowState     =   2  'Maximized
   Begin VB.Frame Frame4 
      Height          =   615
      Left            =   120
      TabIndex        =   0
      Top             =   6720
      Width           =   11535
      Begin VB.CommandButton cmdUpload 
         Caption         =   "&Upload"
         Height          =   315
         Left            =   7440
         TabIndex        =   45
         Top             =   240
         Width           =   960
      End
      Begin VB.CommandButton cmdVerify 
         Caption         =   "&Verify"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   6480
         TabIndex        =   44
         Top             =   240
         Width           =   960
      End
      Begin VB.CommandButton cmdExit 
         Cancel          =   -1  'True
         Caption         =   "&Exit"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   10440
         TabIndex        =   15
         Top             =   240
         Width           =   960
      End
      Begin VB.CommandButton CmdPrint 
         Caption         =   "&Print Err List"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   9360
         TabIndex        =   14
         Top             =   240
         Width           =   1080
      End
      Begin VB.CommandButton cmdStop 
         Caption         =   "Stop"
         Height          =   315
         Left            =   8400
         TabIndex        =   13
         Top             =   240
         Width           =   960
      End
      Begin VB.TextBox txtTempMemberNo 
         Height          =   285
         Left            =   960
         TabIndex        =   12
         Text            =   "txtTempMemberNo"
         Top             =   240
         Visible         =   0   'False
         Width           =   375
      End
      Begin VB.TextBox txtTempERRMessage 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   600
         TabIndex        =   11
         Text            =   "txtTempERRMessage"
         Top             =   240
         Visible         =   0   'False
         Width           =   345
      End
      Begin VB.TextBox txtTempRecordLine 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   240
         TabIndex        =   10
         Text            =   "txtTempRecordLine"
         Top             =   240
         Visible         =   0   'False
         Width           =   345
      End
      Begin VB.TextBox txtTempPLNCode 
         Height          =   285
         Left            =   2040
         TabIndex        =   9
         Text            =   "txtTempPLNCode"
         Top             =   240
         Visible         =   0   'False
         Width           =   375
      End
      Begin VB.TextBox txtTempINSCode 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   1680
         TabIndex        =   8
         Text            =   "txtTempINSCode"
         Top             =   240
         Visible         =   0   'False
         Width           =   345
      End
      Begin VB.TextBox txtImportFile 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   1320
         TabIndex        =   7
         Text            =   "txtImportFile"
         Top             =   240
         Visible         =   0   'False
         Width           =   345
      End
      Begin VB.TextBox txtMBMAdultChild 
         Height          =   285
         Left            =   5400
         TabIndex        =   6
         Top             =   240
         Visible         =   0   'False
         Width           =   615
      End
      Begin VB.TextBox txtAccessFees 
         Height          =   285
         Left            =   4680
         TabIndex        =   5
         Top             =   240
         Visible         =   0   'False
         Width           =   735
      End
      Begin VB.TextBox txtMBMMCOFees 
         Height          =   285
         Left            =   3960
         TabIndex        =   4
         Top             =   240
         Visible         =   0   'False
         Width           =   735
      End
      Begin VB.TextBox txtMBMPremium 
         Height          =   285
         Left            =   3240
         TabIndex        =   3
         Top             =   240
         Visible         =   0   'False
         Width           =   735
      End
      Begin VB.TextBox txtDateJoin 
         Height          =   375
         Left            =   0
         TabIndex        =   2
         Text            =   "Text1"
         Top             =   240
         Visible         =   0   'False
         Width           =   255
      End
      Begin VB.TextBox tempAnnLmt 
         Height          =   285
         Left            =   5400
         TabIndex        =   1
         Text            =   "Text1"
         Top             =   240
         Visible         =   0   'False
         Width           =   975
      End
      Begin MSComDlg.CommonDialog CommonDialog1 
         Left            =   2520
         Top             =   120
         _ExtentX        =   847
         _ExtentY        =   847
         _Version        =   393216
      End
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   6435
      Left            =   120
      TabIndex        =   16
      Top             =   240
      Width           =   11535
      _ExtentX        =   20346
      _ExtentY        =   11351
      _Version        =   393216
      Tabs            =   1
      TabsPerRow      =   1
      TabHeight       =   741
      TabCaption(0)   =   "Data Verification/Uploading"
      TabPicture(0)   =   "frmUploadBordxByIC.frx":0000
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Label8"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "lstERRList"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "Frame2"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "Frame1"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).ControlCount=   4
      Begin VB.Frame Frame1 
         Height          =   615
         Left            =   120
         TabIndex        =   27
         Top             =   360
         Width           =   11295
         Begin VB.ComboBox txtPAYCode 
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
            Left            =   3480
            Sorted          =   -1  'True
            Style           =   2  'Dropdown List
            TabIndex        =   34
            Top             =   195
            Width           =   4455
         End
         Begin VB.TextBox txtImportFilename 
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   8760
            TabIndex        =   33
            Top             =   210
            Width           =   2025
         End
         Begin VB.CommandButton cmdFileLocation 
            Caption         =   " >"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   10800
            TabIndex        =   32
            Top             =   210
            Width           =   360
         End
         Begin VB.ComboBox txtFileName 
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
            Left            =   3120
            TabIndex        =   31
            Text            =   "N - New File"
            Top             =   660
            Visible         =   0   'False
            Width           =   1785
         End
         Begin VB.ComboBox cboHLTCode1 
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
            Left            =   1080
            Sorted          =   -1  'True
            TabIndex        =   30
            Tag             =   "Thanks"
            Text            =   "S - Sihat Malaysia"
            Top             =   200
            Width           =   1785
         End
         Begin VB.ComboBox txtFileVersion 
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
            Left            =   3120
            Sorted          =   -1  'True
            TabIndex        =   35
            Top             =   660
            Visible         =   0   'False
            Width           =   2385
         End
         Begin VB.ComboBox cboSuppAnnualLimit 
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
            Left            =   3840
            Sorted          =   -1  'True
            TabIndex        =   29
            Text            =   "C - Combined"
            Top             =   660
            Visible         =   0   'False
            Width           =   1785
         End
         Begin VB.TextBox txtDiscount 
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   10080
            TabIndex        =   28
            Top             =   675
            Visible         =   0   'False
            Width           =   585
         End
         Begin VB.Label Label4 
            Caption         =   "File Version"
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
            Left            =   4320
            TabIndex        =   42
            Top             =   720
            Visible         =   0   'False
            Width           =   1680
         End
         Begin VB.Label Label55 
            Caption         =   "Payor"
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
            Left            =   3000
            TabIndex        =   41
            Top             =   240
            Width           =   1770
         End
         Begin VB.Label Label1 
            Caption         =   "Filename"
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
            Left            =   8040
            TabIndex        =   40
            Top             =   240
            Width           =   1590
         End
         Begin VB.Label Label5 
            Caption         =   "New/ Endt File"
            Height          =   255
            Left            =   3360
            TabIndex        =   39
            Top             =   720
            Visible         =   0   'False
            Width           =   1095
         End
         Begin VB.Label Label10 
            Caption         =   "Health Type"
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
            TabIndex        =   38
            Top             =   195
            Width           =   1770
         End
         Begin VB.Label Label12 
            Caption         =   "Supp Type"
            Height          =   255
            Left            =   3840
            TabIndex        =   37
            Top             =   720
            Visible         =   0   'False
            Width           =   1215
         End
         Begin VB.Label Label9 
            Caption         =   "Discount"
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
            Left            =   9360
            TabIndex        =   36
            Top             =   675
            Visible         =   0   'False
            Width           =   855
         End
      End
      Begin VB.Frame Frame2 
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
         Height          =   615
         Left            =   120
         TabIndex        =   18
         Top             =   5760
         Width           =   11265
         Begin VB.Label Label11 
            Alignment       =   2  'Center
            Caption         =   "Total Supp"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000FF&
            Height          =   225
            Left            =   6840
            TabIndex        =   26
            Top             =   240
            Width           =   960
         End
         Begin VB.Label txtTotalSupplementary 
            Alignment       =   2  'Center
            BackColor       =   &H00C0FFFF&
            Caption         =   "0"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   270
            Left            =   7800
            TabIndex        =   25
            Top             =   240
            Width           =   960
         End
         Begin VB.Label txtTotalPrincipal 
            Alignment       =   2  'Center
            BackColor       =   &H00C0FFFF&
            Caption         =   "0"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   270
            Left            =   5880
            TabIndex        =   24
            Top             =   240
            Width           =   945
         End
         Begin VB.Label Label7 
            Alignment       =   2  'Center
            Caption         =   "Total Principal"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000FF&
            Height          =   225
            Left            =   4680
            TabIndex        =   23
            Top             =   240
            Width           =   1200
         End
         Begin VB.Label txtSubmissionDate 
            Alignment       =   2  'Center
            BackColor       =   &H00C0FFFF&
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   270
            Left            =   1560
            TabIndex        =   22
            Top             =   240
            Width           =   1200
         End
         Begin VB.Label Label2 
            Alignment       =   2  'Center
            Caption         =   "Submission Date"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000FF&
            Height          =   225
            Left            =   120
            TabIndex        =   21
            Top             =   225
            Width           =   1425
         End
         Begin VB.Label txtTotalError 
            Alignment       =   2  'Center
            BackColor       =   &H00C0FFFF&
            Caption         =   "0"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   270
            Left            =   3720
            TabIndex        =   20
            Top             =   240
            Width           =   960
         End
         Begin VB.Label Label6 
            Alignment       =   2  'Center
            Caption         =   "Total Errors"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000FF&
            Height          =   225
            Left            =   2760
            TabIndex        =   19
            Top             =   240
            Width           =   960
         End
      End
      Begin MSComctlLib.ListView lstERRList 
         Height          =   4455
         Left            =   120
         TabIndex        =   17
         Top             =   1260
         Width           =   11295
         _ExtentX        =   19923
         _ExtentY        =   7858
         View            =   3
         LabelEdit       =   1
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
            Text            =   "Record Line"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Policy No"
            Object.Width           =   2646
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "Name"
            Object.Width           =   3528
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   3
            Text            =   "Error Message"
            Object.Width           =   7056
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   4
            Text            =   "Filename"
            Object.Width           =   1764
         EndProperty
      End
      Begin VB.Label Label8 
         Alignment       =   2  'Center
         Caption         =   "Error Listing"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   225
         Left            =   120
         TabIndex        =   43
         Top             =   960
         Width           =   1155
      End
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      BackColor       =   &H00404040&
      Caption         =   "Verify && Upload Bordx"
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
      TabIndex        =   46
      Top             =   0
      Width           =   11655
   End
End
Attribute VB_Name = "frmUploadBordxByIC"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim MBMCover As New ADODB.Recordset
Dim MBM  As New ADODB.Recordset
Dim MBMOthers As New ADODB.Recordset
Dim COM As New ADODB.Recordset
Dim ERR As New ADODB.Recordset
Dim MBMCov As New ADODB.Recordset
Dim TextLine
Dim Message As String
Dim RECLine As String
Dim recordCount As Integer



Private Sub EndtAddSupplementary()
    Dim strsqlMBMCov As String
    Dim listrecord As Integer
    Dim listloop As Integer
    Dim ListCount As Integer
    Dim strsql2 As String
    Dim MBMCovPersons As New ADODB.Recordset
    Dim LastCoveredID As String
    Dim SelectAge As String
    Dim strdelcov As String
    Dim ForLoop As Integer
    Dim planstring
    Dim BordxDate As String
    Dim PayEffDate As String
    Dim AgeValue As String
    
    strsqlMBMCov = "Select * from MBMCoveredPersons Where MBMCNumber = '" & Trim(txtTempMemberNo) & "'"
    MBMCovPersons.Open strsqlMBMCov, medixmasterconn, adOpenKeyset, adLockOptimistic
    
    'strsqlMBMCov = "SELECT MBMCNumber, MAX(MBMCCoverID) AS CoveredID "
    'strsqlMBMCov = strsqlMBMCov & "From MBMCoveredPersons GROUP BY MBMCNumber "
    'strsqlMBMCov = strsqlMBMCov & "HAVING (MBMCNumber = '" & txtTempMemberNo & "') "
    'strsqlMBMCov = strsqlMBMCov & "(MBMCMemberType = 'C')"
    'MBMCovPersons.Open strsqlMBMCov, medixmasterconn, adOpenKeyset, adLockOptimistic
    'MBMCovPersons.Close
    'Set MBMCovPersons = Nothing
    If MBMCovPersons.recordCount > 0 Then
        LastCoveredID = MBMCovPersons.recordCount
        LastCoveredID = LastCoveredID + 1
    Else
        LastCoveredID = 1
    End If
    
    If LastCoveredID < 10 Then
        LastCoveredID = 0 & LastCoveredID
    End If
    'LastCoveredID = GetLastCoveredID(Trim(Mid(TextLine, 11, 18)))
        
         MBMCovPersons.AddNew
         With MBMCovPersons
             !MBMCNumber = txtTempMemberNo
             !MBMCSalutation = Mid(TextLine, 69, 5)
             SelectAge = GetAge(Mid(TextLine, 401, 2) & "/" & Mid(TextLine, 403, 2) & "/" & Mid(TextLine, 405, 4), Mid(TextLine, 320, 2) & "/" & Mid(TextLine, 322, 2) & "/" & Mid(TextLine, 324, 4))
             AgeValue = Mid(GetAgeBand(Trim(Mid(cboHLTCode1, 1, InStr(1, cboHLTCode1, "-") - 1)), SelectAge), 1, 2)
             !MBMCAgeband = AgeValue
               If AgeValue = "01" Then
                   !MBMCAdultChild = "C"
               Else
                   !MBMCAdultChild = "A"
               End If
             !MBMCStatus = "A"
             !MBMCBordxDate = txtSubmissionDate
             !MBMCName = Trim(Mid(TextLine, 74, 60))
             !MBMCICBCPPNo = Trim(Mid(TextLine, 300, 20))
             !MBMCDOB = Mid(TextLine, 320, 2) & "/" & Mid(TextLine, 322, 2) & "/" & Mid(TextLine, 324, 4)
             !MBMCSex = Mid(TextLine, 328, 1)
             '!MBMAdultChild = lstCovAdd.ListItems(listloop).SubItems(5)
             !MBMCRELCode = Mid(TextLine, 10, 1)
             !MBMCBloodType = Mid(TextLine, 378, 3)
             !MBMCAllergic = Mid(TextLine, 381, 20)
             !MBMCExclusion = Mid(TextLine, 463, 60)
             '!MBMPrinted = "N"
             '!MBMExported = "N"
             !MBMCMemberType = "C"
             !MBMCCoverID = LastCoveredID
        
        End With
        MBMCovPersons.Update
        MBMCovPersons.Close
        Set MBMCovPersons = Nothing
End Sub

Private Sub cboHLTCode1_Change()
Dim PAY As New ADODB.Recordset
Dim strsql As String
    
    strsql = "Select  PAYCode, PAYCompanyName From PAY where HLTCOde ='" & Mid(Trim(cboHLTCode1), 1, 1) & "'"
    PAY.Open strsql, medixmasterconn, adOpenKeyset, adLockOptimistic
    txtPAYCode.Clear
    Do Until PAY.EOF
       txtPAYCode.AddItem ChkStr(PAY!PAYCode) & " - " & ChkStr(PAY!PAYCompanyName)
       PAY.MoveNext
    Loop
    PAY.Close
    Set PAY = Nothing

End Sub

Private Sub cboHLTCode1_Click()
Dim PAY As New ADODB.Recordset
Dim strsql As String
    
    strsql = "Select  PAYCode, PAYCompanyName From PAY where HLTCOde ='" & Mid(Trim(cboHLTCode1), 1, 1) & "'"
    PAY.Open strsql, medixmasterconn, adOpenKeyset, adLockOptimistic
    txtPAYCode.Clear
    Do Until PAY.EOF
       txtPAYCode.AddItem ChkStr(PAY!PAYCode) & " - " & ChkStr(PAY!PAYCompanyName)
       PAY.MoveNext
    Loop
    PAY.Close
    Set PAY = Nothing

End Sub

Private Sub cmdExit_Click()
    Unload Me
End Sub

Private Sub cmdFileLocation_Click()
  CommonDialog1.CancelError = True
  On Error GoTo ErrHandler
  CommonDialog1.InitDir = SRVPath & "Bor_In"
  CommonDialog1.Flags = cdlOFNHideReadOnly
  CommonDialog1.Filter = "All Files (*.*)|*.*|Text Files (*.txt)|*.txt|Dat Files (*.Dat)|*.Dat"
  'CommonDialog1.FilterIndex = 2
  CommonDialog1.ShowOpen
  txtImportFilename = CommonDialog1.FileTitle
  Exit Sub
  
ErrHandler:
  'User pressed the Cancel button
  Exit Sub

End Sub

Private Sub CmdPrint_Click()
    'Call ExportToExcelError(lstERRList)
    Screen.MousePointer = vbHourglass
        
    If lstERRList.listitems.Count <> 0 Then
        Call ExportListViewtoExcel(lstERRList)
    Else
        MsgBox "Report cannot be printed without any record.", vbInformation
    End If
        
    Screen.MousePointer = vbDefault
End Sub

Private Sub cmdUpload_Click()
Dim date1 As Date
Dim show As String
Dim dy As Long
Dim dh As Long
Dim ds As Long
Dim MyDate
Dim RecordSaved
    
    'If txtFileVersion = "" And Mid(txtFileName, 1, 1) = "N" Then
    '    MsgBox "Please Enter File Version", vbOKOnly + vbCritical, DialogBoxTitle
    '    txtFileVersion.SetFocus
    If txtPAYCode = "" Then
        MsgBox "Please Enter Payor Code", vbOKOnly + vbCritical, DialogBoxTitle
        txtPAYCode.SetFocus
    'ElseIf txtFileName = "" Then
    '    MsgBox "Please Enter FileName", vbOKOnly + vbCritical, DialogBoxTitle
    '    txtPAYCode.SetFocus
    ElseIf txtImportFilename = "" Then
        MsgBox "Please Enter Filename", vbOKOnly + vbCritical, DialogBoxTitle
        txtImportFilename.SetFocus
    Else
      RecordSaved = MsgBox("Do you want to upload this file?", vbYesNo + vbExclamation)
        If RecordSaved = vbYes Then
            Screen.MousePointer = vbHourglass
            cmdUpload.Enabled = False
            cmdExit.Enabled = False
            
            Call GenerateENDtMember
            
            SSTab1.Enabled = True
            cmdExit.Enabled = True
            
        End If
        End If
    'End If
'    Exit Sub
'ErrorSave:
'    Screen.MousePointer = vbDefault
'    If startTrans = True Then
'        medixmasterconn.RollbackTrans
'    End If
'    MsgBox ERR.Description
    'COM.Close
    Screen.MousePointer = vbDefault
    cmdExit.Enabled = True

End Sub

Private Sub GenerateENDtMember()
    Dim PayorCode As String
    Dim CoveredIDCount As Integer
    Dim recordCount As Integer
    Dim RecCount As Integer
    Dim strsql As String
    Dim NoofRec As String
    'cmdGenerate.Enabled = False
    'cmdExit.Enabled = False
    NoofRec = 0
    'MBMCov.Open "MBMCoveredPersons", medixmasterconn, adOpenKeyset, adLockOptimistic

    
    Open txtImportFilename For Input As #1   ' Open file.
       Line Input #1, TextLine   ' Read line into variable.
        If Mid(TextLine, 2, 2) <> Mid(txtPAYCode, 1, 2) Then
            MsgBox "Payor Selected Is Not Match The File Payor", vbOKOnly + vbCritical, DialogBoxTitle
            txtPAYCode.SetFocus
        Else
            If Mid(TextLine, 1, 1) = "E" Then
                PayorCode = Mid(TextLine, 2, 2)
                txtSubmissionDate = Mid(TextLine, 4, 2) & "/" & Mid(TextLine, 6, 2) & "/" & Mid(TextLine, 8, 4)
            End If
            recordCount = 1
            'MBM.Open "MBM", medixmasterconn, adOpenKeyset, adLockOptimistic
            'txtBatchNo = CheckBatch(Mid(txtPAYCode, 1, 2), txtSubmissionDate, Mid(txtFileName, 1, 1))
            Do While Not EOF(1)   ' Loop until end of file.
               Line Input #1, TextLine   ' Read line into variable.
                recordCount = recordCount + 1
                    If Mid(TextLine, 10, 1) = "P" Then
                        NoofRec = 0
                        strsql = "Select MBMNumber from MBM Where MBMIcBcPp = '" & Trim(Mid(TextLine, 300, 20)) & "' And MBMPolicyNo = '" & Trim(Mid(TextLine, 29, 25)) & "'"
                        MBM.Open strsql, medixmasterconn, adOpenKeyset, adLockOptimistic
                        txtTempMemberNo = MBM!MBMNumber
                        NoofRec = MBM.recordCount
                    MBM.Close
                    Set MBM = Nothing
                    End If
                        If NoofRec > 0 And Mid(TextLine, 10, 1) <> "P" Then
                            Call EndtAddSupplementary
                            RecCount = RecCount + 1
                        'ElseIf NoofRec = 0 And Mid(TextLine, 10, 1) = "P" Then
                        '    Call EndtPrincipalDataImport
                        '    Call EndtOthersDataImport
                        End If
                DoEvents
                'txtProcessRecord = RecordCount
            Loop
            MsgBox "Record Import Sucessfully...", vbOKOnly + vbInformation, DialogBoxTitle
            'cmdGenerate.Enabled = False
            'MBMCov.Close
            'Set MBMCov = Nothing
        End If
    Close #1   ' Close file.
    'ProcessRecord = RecCount
End Sub

Private Sub ComboListing()
    Dim HLT As New ADODB.Recordset
    Dim PAY As New ADODB.Recordset
    Dim strsql As String
    
    HLT.Open "HLT", medixmasterconn, adOpenKeyset, adLockOptimistic
    'PAY.Open "PAY", medixmasterconn, adOpenKeyset, adLockOptimistic
    
    Do Until HLT.EOF
       cboHLTCode1.AddItem HLT!HLTCode & " - " & HLT!HLTDescription
       'cboHLTCode1.AddItem HLT!HLTCode & " - " & HLT!HLTDescription
       HLT.MoveNext
    Loop
        
    strsql = "Select  PAYCode, PAYCompanyName From PAY where HLTCOde ='" & Mid(Trim(cboHLTCode1), 1, 1) & "'"
    PAY.Open strsql, medixmasterconn, adOpenKeyset, adLockOptimistic
    txtPAYCode.Clear
    Do Until PAY.EOF
       txtPAYCode.AddItem ChkStr(PAY!PAYCode) & " - " & ChkStr(PAY!PAYCompanyName)
       'cboPAY.AddItem ChkStr(PAY!PAYCode) & " - " & ChkStr(PAY!PAYCompanyName)
       PAY.MoveNext
    Loop
    PAY.Close
    Set PAY = Nothing

    'Do Until PAY.EOF
    '    txtPAYCode.AddItem PAY!PAYCode & " - " & PAY!PAYCompanyName
    '    PAY.MoveNext
    'Loop
        
    txtFileName.AddItem "N" & " - " & "New File"
    txtFileName.AddItem "E" & " - " & "Endorsement File"
        
    txtFileVersion.AddItem "01" & " - " & "Version 520"
    txtFileVersion.AddItem "02" & " - " & "Version 790"
    txtFileVersion.AddItem "03" & " - " & "Version 4800"
    
    'cboASCIIFileVer.AddItem "01" & " - " & "Version 185"
    'cboASCIIFileVer.AddItem "02" & " - " & "Version 195"
    
    cboSuppAnnualLimit.AddItem "C" & " - " & "Combined Limit"
    cboSuppAnnualLimit.AddItem "S" & " - " & "Separate Limit"
    
    HLT.Close
    'PAY.Close
    
    Set HLT = Nothing
    'Set PAY = Nothing
End Sub

Private Sub ComboList()
    Dim HLT As New ADODB.Recordset
    Dim PAY As New ADODB.Recordset
    Dim strsql As String
    
    HLT.Open "HLT", medixmasterconn, adOpenKeyset, adLockOptimistic
    'PAY.Open "PAY", medixmasterconn, adOpenKeyset, adLockOptimistic
    
    'Do Until HLT.EOF
        'cboHLTCode.AddItem HLT!HLTCode & " - " & HLT!HLTDescription
        'HLT.MoveNext
    'Loop
    
    'strsql = "Select  PAYCode, PAYCompanyName From PAY where HLTCOde ='" & Mid(Trim(cboHLTCode), 1, 1) & "'"
    'PAY.Open strsql, medixmasterconn, adOpenKeyset, adLockOptimistic
    'cboPAYCode.Clear
    'Do Until PAY.EOF
    '   cboPAYCode.AddItem ChkStr(PAY!PAYCode) & " - " & ChkStr(PAY!PAYCompanyName)
    '   PAY.MoveNext
    'Loop
    'PAY.Close
    'Set PAY = Nothing

   ' Do Until PAY.EOF
   '     cboPAYCode.AddItem PAY!PAYCode & " - " & PAY!PAYCompanyName
   '     cboPAY.AddItem PAY!PAYCode & " - " & PAY!PAYCompanyName
   '     PAY.MoveNext
   ' Loop
        
    'cboExportOption.AddItem "D" & " - " & "Bordx Range"
    'cboExportOption.AddItem "B" & " - " & "Batch Bordx"
    
    'cboBordxOPT.AddItem "D" & " - " & "Bordx Range"
    'cboBordxOPT.AddItem "B" & " - " & "Batch Bordx"
    
    HLT.Close
    Set HLT = Nothing
    'PAY.Close
End Sub

Private Sub cmdVerify_Click()
    txtSubmissionDate = ""
    txtTotalPrincipal = 0
    txtTotalSupplementary = 0
    txtTotalError = 0
    cmdFileLocation.Enabled = False
    cmdUpload.Enabled = False
    'If Mid(txtFileName, 1, 1) = "N" And txtFileVersion = "" Then
    '    MsgBox "Please Enter File Version", vbOKOnly + vbCritical, DialogBoxTitle
    '    txtFileName.SetFocus
    If txtPAYCode = "" Then
        MsgBox "Please Enter Payor Code", vbOKOnly + vbCritical, DialogBoxTitle
        txtPAYCode.SetFocus
    'ElseIf txtFileName = "" Then
    '    MsgBox "Please Enter File Type", vbOKOnly + vbCritical, DialogBoxTitle
    '    txtFileName.SetFocus
    ElseIf txtImportFilename = "" Then
        MsgBox "Please Enter Filename", vbOKOnly + vbCritical, DialogBoxTitle
        'txtImportFilename.SetFocus
    Else
        
        cmdVerify.Enabled = False
        cmdUpload.Enabled = False
        lstERRList.listitems.Clear
                           
        
        'If Mid(txtFileName, 1, 1) = "N" Then
        '    Call VerifyNewMember
        '    If intExit = 0 Then
         '       MsgBox "Verifed Sucessfully...", vbOKOnly + vbInformation, DialogBoxTitle
         '       Screen.MousePointer = vbDefault
         '   Else
         '   Screen.MousePointer = vbDefault
         '   End If
        'If Mid(txtFileName, 1, 1) = "E" Then
            Call VerifyENDtMember
        'If intExit = 0 Then
            MsgBox "Verifed Sucessfully...", vbOKOnly + vbInformation, DialogBoxTitle
            Screen.MousePointer = vbDefault
        'Else
        '    Screen.MousePointer = vbDefault
        'End If
        'End If
    End If
            cmdVerify.Enabled = True
            cmdUpload.Enabled = True
            cmdFileLocation.Enabled = True
            'intExit = 0

End Sub

Private Sub Form_Load()
    Call ComboListing
    Call ComboList
End Sub

'---------------------Start Verification of ENDt File-----------------------

Private Sub VerifyENDtMember()
    Dim PayorCode As String
    Dim CoveredIDCount As Integer
    Dim CheckICValue As String
    Dim ErrorCount As Integer
    Dim TotPrincipal As Integer
    Dim TotSupplementary As Integer
    Dim ListError As ListItem
    Dim msg As VerifyError
    
    lstERRList.listitems.Clear
    
    Screen.MousePointer = vbHourglass
    MBM.Open "MBM", medixmasterconn, adOpenKeyset, adLockOptimistic
    ERR.Open "ERR", medixmasterconn, adOpenKeyset, adLockOptimistic
    Open txtImportFilename For Input As #1   ' Open file.
    Do While Not EOF(1)   ' Loop until end of file.
       Line Input #1, TextLine   ' Read line into variable.
        recordCount = recordCount + 1
            If recordCount = 1 Then
                txtSubmissionDate = Mid(TextLine, 4, 2) & "/" & Mid(TextLine, 6, 2) & "/" & Mid(TextLine, 8, 4)
                If Mid(TextLine, 1, 1) <> "E" Then
                    Call CreateErrorList(recordCount, "", "", "1st Record must be header with Data Type 'F'")
                    ErrorCount = ErrorCount + 1
                    Call ErrorListing
                End If
                If Trim(Mid(TextLine, 2, 2)) = "" Or Trim(Mid(TextLine, 2, 2)) <> Mid(txtPAYCode, 1, 2) Then
                    Call CreateErrorList(recordCount, "", "", "Invalid Insurer's Code " & Trim(Mid(TextLine, 2, 2)))
                    ErrorCount = ErrorCount + 1
                    Call ErrorListing
                End If
                If Mid(TextLine, 4, 11) = "" Then
                    Call CreateErrorList(recordCount, "", "", "Invalid Date" & Trim(Mid(TextLine, 2, 2)))
                    ErrorCount = ErrorCount + 1
                    Call ErrorListing
                End If
            End If
                             
            If recordCount >= 2 Then
                If Mid(TextLine, 10, 1) = "P" Then
                    TotPrincipal = TotPrincipal + 1
                ElseIf Mid(TextLine, 10, 1) <> "P" And Mid(TextLine, 1, 1) <> "T" Then
                    TotSupplementary = TotSupplementary + 1
                End If
            End If
            
            If recordCount >= 2 And Mid(TextLine, 1, 1) <> "T" Then
            
                If Trim(Mid(TextLine, 2, 8)) = "" Then
                    Call CreateErrorList(recordCount, Mid(TextLine, 29, 25), Mid(TextLine, 74, 60), "Endorsement Effective Date is blank!")
                    ErrorCount = ErrorCount + 1
                    Call ErrorListing
                End If
                
                If Trim(Mid(TextLine, 10, 1)) = "" Then
                    Call CreateErrorList(recordCount, Mid(TextLine, 29, 25), Mid(TextLine, 74, 60), "Data Type is blank")
                    ErrorCount = ErrorCount + 1
                    Call ErrorListing
                End If
                
                If Trim(Mid(TextLine, 11, 18)) = "" Then
                            Call CreateErrorList(recordCount, Mid(TextLine, 29, 25), Mid(TextLine, 74, 60), "Membership No Is Blank")
                    ErrorCount = ErrorCount + 1
                    Call ErrorListing
                End If
          
                If Trim(Mid(TextLine, 29, 25)) = "" Then
                    Call CreateErrorList(recordCount, Mid(TextLine, 29, 25), Mid(TextLine, 74, 60), "Policy No Is Blank")
                    ErrorCount = ErrorCount + 1
                    Call ErrorListing
                End If
                
                If Trim(Mid(TextLine, 273, 20)) <> "" Then
                   CheckICValue = CheckIC(Trim(Mid(TextLine, 273, 20)))
                    If CheckICValue <> "" Then
                       Call CreateErrorList(recordCount, Mid(TextLine, 2, 25), Mid(TextLine, 47, 60), "Same IC " & Trim(Mid(TextLine, 273, 20)) & " Existed in Membership No. :" & CheckICValue)
                       ErrorCount = ErrorCount + 1
                       Call ErrorListing
                    End If
                End If
                
                If Trim(Mid(TextLine, 74, 60)) = "" Then
                    Call CreateErrorList(recordCount, Mid(TextLine, 29, 25), Mid(TextLine, 74, 60), "Name Is Blank")
                    ErrorCount = ErrorCount + 1
                    Call ErrorListing
                End If
                                
                If Trim(Mid(TextLine, 134, 30)) = "" Or Trim(Mid(TextLine, 164, 30)) = "" Or Trim(Mid(TextLine, 194, 30)) = "" And Mid(TextLine, 10, 1) = "P" Then
                    Call CreateErrorList(recordCount, Mid(TextLine, 29, 25), Mid(TextLine, 74, 60), "Address Is Blank")
                    ErrorCount = ErrorCount + 1
                    Call ErrorListing
                End If

                If Trim(Mid(TextLine, 224, 20)) = "" And Mid(TextLine, 10, 1) = "P" Then
                    Call CreateErrorList(recordCount, Mid(TextLine, 29, 25), Mid(TextLine, 74, 60), "City Is Blank")
                    ErrorCount = ErrorCount + 1
                    Call ErrorListing
                End If
                
                If Trim(Mid(TextLine, 244, 5)) = "" And Mid(TextLine, 1, 1) = "P" Then
                    Call CreateErrorList(recordCount, Mid(TextLine, 29, 25), Mid(TextLine, 74, 60), "PostCode Is Blank")
                    ErrorCount = ErrorCount + 1
                    Call ErrorListing
                End If
 
                If Trim(Mid(TextLine, 249, 15)) = "" And Mid(TextLine, 1, 1) = "P" Then
                    Call CreateErrorList(recordCount, Mid(TextLine, 29, 25), Mid(TextLine, 74, 60), "State Code Is Blank")
                    ErrorCount = ErrorCount + 1
                    Call ErrorListing
                End If

                If Trim(Mid(TextLine, 300, 20)) = "" Then
                    Call CreateErrorList(recordCount, Mid(TextLine, 29, 25), Mid(TextLine, 74, 60), "IC/BC/PP Is Blank")
                    ErrorCount = ErrorCount + 1
                    Call ErrorListing
                End If
                
                If Trim(Mid(TextLine, 320, 8)) = "" Then
                    Call CreateErrorList(recordCount, Mid(TextLine, 29, 25), Mid(TextLine, 74, 60), "DOB Is Blank")
                    ErrorCount = ErrorCount + 1
                    Call ErrorListing
                End If

                If Trim(Mid(TextLine, 10, 1)) = "P" And Trim(Mid(TextLine, 401, 8)) = "" Then
                    Call CreateErrorList(recordCount, Mid(TextLine, 29, 25), Mid(TextLine, 74, 60), "Effective Date Is Blank")
                    ErrorCount = ErrorCount + 1
                    Call ErrorListing
                End If
                If Trim(Mid(TextLine, 10, 1)) = "P" And Trim(Mid(TextLine, 409, 8)) = "" Then
                    Call CreateErrorList(recordCount, Mid(TextLine, 29, 25), Mid(TextLine, 74, 60), "Expiry Date Is Blank")
                    ErrorCount = ErrorCount + 1
                    Call ErrorListing
                End If
            End If
            
                If Trim(Mid(txtFileVersion, 1, 2)) = "01" Then
                     If Trim(Mid(TextLine, 417, 2)) = "" And Mid(TextLine, 1, 1) = "P" Then
                         Call CreateErrorList(recordCount, Mid(TextLine, 29, 25), Mid(TextLine, 74, 60), "Invalid Plan Code")
                         ErrorCount = ErrorCount + 1
                         Call ErrorListing
                     End If
                
                ElseIf Trim(Mid(txtFileVersion, 1, 2)) = "02" Then
                     If Trim(Mid(TextLine, 417, 4)) = "" And Mid(TextLine, 1, 1) = "P" Then
                         Call CreateErrorList(recordCount, Mid(TextLine, 29, 25), Mid(TextLine, 74, 60), "Invalid Plan Code")
                         ErrorCount = ErrorCount + 1
                         Call ErrorListing
                     End If
                 End If
                 
                If Trim(Mid(TextLine, 421, 1)) = "" And Mid(TextLine, 1, 1) = "P" Then
                    Call CreateErrorList(recordCount, Mid(TextLine, 29, 25), Mid(TextLine, 74, 60), "Invalid Insured Type")
                    ErrorCount = ErrorCount + 1
                    Call ErrorListing
                End If
                                                      
        DoEvents
    Loop
                If Mid(TextLine, 1, 1) <> "T" Then
                    Call CreateErrorList(recordCount, "", "", "1st Record must be trailer with Data Type 'T'")
                    ErrorCount = ErrorCount + 1
                ElseIf Trim(Mid(TextLine, 2, 10)) <> TotPrincipal Then
                     Call CreateErrorList(recordCount, "", "", "Total Principal Record Not Match Trailer Total Principal")
                ElseIf Trim(Mid(TextLine, 12, 10)) <> TotSupplementary Then
                     Call CreateErrorList(recordCount, "", "", "Total Supplementary Record Not Match Trailer Total Supplementary")
                End If
                            
    Close #1   ' Close file.

    MBM.Close
    ERR.Close
    txtTotalError = ErrorCount
    txtTotalPrincipal = TotPrincipal
    txtTotalSupplementary = TotSupplementary
    recordCount = 0
    
End Sub
'**************************End Verification of ENDt File*********************************

Private Sub CreateErrorList(RecordLine, Policyno, Name, msg)
        
        
        ERR.AddNew
            With ERR
                 !ERRRecordLine = RecordLine
                 !ERRPolicyNo = Trim(Policyno)
                 'If Len(Name) Then
                 '   !ERRName = Trim(Name)
                'End If
                 !ERRMessage = msg
                 !ErrFileName = txtImportFilename
            End With
        ERR.Update
        
        Message = msg
        RECLine = RecordLine
        
End Sub

Private Sub ErrorListing()
    txtTempRecordLine = recordCount
    txtTempERRMessage = Message
    'Text1 = RECLine
    Dim ListError As ListItem
            
            
            Set ListError = lstERRList.listitems.Add(, , txtTempRecordLine)
            If Mid(txtFileName, 1, 1) = "N" Then
                ListError.SubItems(1) = Trim(Mid(TextLine, 2, 25))
                ListError.SubItems(2) = Trim(Mid(TextLine, 47, 60))
            ElseIf Mid(txtFileName, 1, 1) = "E" Then
                ListError.SubItems(1) = Trim(Mid(TextLine, 29, 25))
                ListError.SubItems(2) = Trim(Mid(TextLine, 74, 60))
            End If
            
            ListError.SubItems(3) = Trim(txtTempERRMessage)
            ListError.SubItems(4) = Trim(txtImportFilename)
                
        txtTempERRMessage = ""
        Message = ""
        'RECLine = ""
        'lstERRList.ListItems.Clear
        
End Sub

Private Function CheckIC(IC As String)
    Dim MBM As New ADODB.Recordset
    Dim strsql As String
    Dim strsqlOthers As String
    Dim MBMOthers As New ADODB.Recordset
    
    strsql = "Select MBMNumber, MBMICBCPP From MBM Where MBMICBCPP = '" & Trim(IC) & "'"
    MBM.Open strsql, medixmasterconn, adOpenKeyset, adLockOptimistic
    
    If MBM.recordCount > 0 Then
        CheckIC = MBM!MBMNumber
    strsqlOthers = "Select MBMFirstJoinedDate from MBMOthers where MBMNumber = '" & Trim(MBM!MBMNumber) & "'"
    MBMOthers.Open strsqlOthers, medixmasterconn, adOpenKeyset, adLockOptimistic
    
    If MBMOthers.recordCount > 0 Then
        If Len(MBMOthers!MBMFirstJoinedDate) Then
            txtDateJoin = MBMOthers!MBMFirstJoinedDate
        End If
    End If

    MBMOthers.Close
    Set MBMOthers = Nothing

    End If
    
    
    MBM.Close
    Set MBM = Nothing
End Function

Private Sub EndtPrincipalDataImport(Payor As String, CoveredIDCount As Integer)
'On Error Resume Next
'On Error GoTo error1
    Dim PayEffDate As String
    Dim PayExpDate As String
    Dim MemberDOB As String
    Dim FindAge As String
    Dim InsuredCode As String
    Dim PlanCode As String
    Dim AgeValue As String
    Dim AnnualLimitValue As String
    Dim GetProString As String
    Dim tmpPreMemberNo As String
    Dim ListMember As ListItem
    'Dim COM As New ADODB.Recordset
    Dim MBM As New ADODB.Recordset
    Dim MBMOthers As New ADODB.Recordset
    Dim completeMBM, completeMBMOthes, CompleteMBMCov As Boolean
    Dim tmpSex As String
        MBM.Open "MBM", medixmasterconn, adOpenKeyset, adLockOptimistic
        'COM.Open "COM", medixmasterconn, adOpenKeyset, adLockOptimistic
       ' medixmasterconn.BeginTrans
        
        MBM.AddNew
        With MBM
            If Mid(TextLine, 10, 1) = "P" Then
                !MBMLastEndorseStatus = Mid(TextLine, 1, 1)
                !MBMENDtBordxDate = CDate(txtSubmissionDate)
                !MBMENDtEffDate = Mid(TextLine, 2, 2) & "/" & Mid(TextLine, 4, 2) & "/" & Mid(TextLine, 6, 4)
                !MBMMemberType = "P"
                !MBMPolicyNo = Mid(TextLine, 29, 25)
                !MBMSalutation = Mid(TextLine, 67, 5)
                !MBMName = Mid(TextLine, 74, 60)
                !MBMAddress1 = Mid(TextLine, 134, 30)
                !MBMAddress2 = Mid(TextLine, 164, 30)
                !MBMAddress3 = Mid(TextLine, 194, 30)
                !CTYCode = Mid(TextLine, 224, 20)
                !MBMPostCode = Mid(TextLine, 244, 5)
                !STACode = Mid(TextLine, 249, 15)
                !MBMTelnoH = Mid(TextLine, 264, 12)
                !MBMTelNoM = Mid(TextLine, 276, 12)
                !MBMTelNoO = Mid(TextLine, 288, 12)
                !MBMIcBcPp = Mid(TextLine, 300, 20)
                !PAYCode = Mid(txtPAYCode, 1, 2)
                !HLTCode = Mid(cboHLTCode1, 1, 2)
                !MBMDOB = Mid(TextLine, 320, 2) & "/" & Mid(TextLine, 322, 2) & "/" & Mid(TextLine, 324, 4)
                !MBMSex = Mid(TextLine, 328, 1)
                tmpSex = Mid(TextLine, 328, 1)
                !MBMPayorEffDate = Mid(TextLine, 401, 2) & "/" & Mid(TextLine, 403, 2) & "/" & Mid(TextLine, 405, 4)
                !MBMPayorEXPDate = Mid(TextLine, 409, 2) & "/" & Mid(TextLine, 411, 2) & "/" & Mid(TextLine, 413, 4)
                !INSCode = Mid(TextLine, 421, 1)
                !PLNCode = Mid(TextLine, 417, 4)
                !HLTCode = Mid(cboHLTCode1, 1, 1)
                InsuredCode = Trim(!INSCode)
                PlanCode = Trim(!PLNCode)
                !MBMBordxDate = CDate(txtSubmissionDate)
                !MBMDataStatus = "A"
                !MBMCOVID = "00"
                !MBMValueDate = Date
                !SRVCodeList = "EA"
                
        '------------------------------------------------------------------------------------------------------------
                'If Mid(cboHLTCode, 1, 1) = "S" Then
                '    !PLNCode = Mid(TextLine, 417, 2)
                'ElseIf Mid(cboHLTCode, 1, 1) = "N" Then
                '    !PLNCode = Mid(TextLine, 417, 4)
                '    '!MBMRenewal = Mid(TextLine, 435, 1)
                'End If
        '------------------------------------------------------------------------------------------------------------------------
                !MBMStatus = "A"
                'If Len(txtBatchNo) Then
                '    !MBMBatchNo = txtBatchNo
                'End If
                !MBMPolicyYear = 1
        '------------------------------------------------------------------------------------------------------------------------
                PayEffDate = Mid(TextLine, 401, 2) & "/" & Mid(TextLine, 403, 2) & "/" & Mid(TextLine, 405, 4)
                PayExpDate = Mid(TextLine, 409, 2) & "/" & Mid(TextLine, 411, 2) & "/" & Mid(TextLine, 413, 4)
                MemberDOB = Mid(TextLine, 320, 2) & "/" & Mid(TextLine, 322, 2) & "/" & Mid(TextLine, 324, 4)
                FindAge = GetAge(CDate(PayEffDate), CDate(MemberDOB))
                AgeValue = Mid(GetAgeBand(Mid(txtFileVersion, 1, 2), FindAge), 1, 2)
                !AGECode = AgeValue
                If AgeValue = "01" Then
                    !MBMAdultChild = "C"
                Else
                    !MBMAdultChild = "A"
                End If
                
                    txtMBMPremium = Format(GetPremium(InsuredCode, Payor, AgeValue, PlanCode, Mid(cboHLTCode1, 1, 1), CDate(PayEffDate), Trim(tmpSex)), "#,##0.00")
                    
                    txtMBMPremium = Round(InsertPremium(CDate(PayExpDate), CDate(PayEffDate), Format(txtMBMPremium, "#,##0.00")), 0)

                    If AgeValue = "01" Then
                        txtMBMMCOFees = Format(GetMCO(InsuredCode, Mid(cboHLTCode1, 1, 2), "C", CDate(PayEffDate), Trim(PlanCode)), "#,##0.00")
                    Else
                        txtMBMMCOFees = Format(GetMCO(InsuredCode, Mid(cboHLTCode1, 1, 2), "A", CDate(PayEffDate), Trim(PlanCode)), "#,##0.00")
                    End If
                    txtMBMMCOFees = Round(InsertMCOFees(CDate(PayExpDate), CDate(PayEffDate), Format(txtMBMMCOFees, "#,##0.00")), 0)

                    AnnualLimitValue = GetAnnualLimit(InsuredCode, PlanCode, Mid(cboHLTCode1, 1, 1), CDate(PayEffDate))
                    !MBMAvailableLimit = AnnualLimitValue
                                                    
                    If AgeValue = "01" Then
                        GetProString = GetProviderCharges("EA", InsuredCode, Mid(cboHLTCode1, 1, 1), "C")
                    Else
                        GetProString = GetProviderCharges("EA", InsuredCode, Mid(cboHLTCode1, 1, 1), "A")
                    End If
                    If Len(Trim(GetProString)) Then
                        txtAccessFees = Format(GetProString, "#,##0.00")
                    Else
                        txtAccessFees = 0
                    End If
                        txtAccessFees = Round(InsertProvider(CDate(PayExpDate), CDate(PayEffDate), Format(txtAccessFees, "#,##0.00")), 2)
                                                         
                    'AnnualLimitValue = GetAnnualLimit(InsuredCode, PlanCode, Mid(cboHLTCode1, 1, 1), CDate(PayEffDate))
                    '!MBMAvailableLimit = AnnualLimitValue
                    
                    
                    tmpPreMemberNo = Payor & PlanCode & InsuredCode & GenerateMembershipNo(Payor, PlanCode, InsuredCode)
                    !MBMNumber = tmpPreMemberNo
                    txtTempMemberNo = tmpPreMemberNo
            End If
            End With
                MBM.Update
    'COM.Close
    MBM.Close
    Set MBM = Nothing
    
    'MBMOthers.Close
 '   Exit Sub
'error1:
    
'    MsgBox ERR.Description
End Sub

Private Sub EndtOthersDataImport()
'On Error Resume Next
'On Error GoTo error1
    Dim tmpPreMemberNo As String
    Dim MBMOthers As New ADODB.Recordset
        
        MBMOthers.Open "MBMOthers", medixmasterconn, adOpenKeyset, adLockOptimistic
       'medixmasterconn.BeginTrans
        
        MBMOthers.AddNew
        With MBMOthers
            !MBMNumber = txtTempMemberNo
            !MBMAgentCode = Mid(TextLine, 54, 15)
            !MBMWeight = Mid(TextLine, 339, 3)
            !MBMHeight = Mid(TextLine, 342, 3)
            !MBMWorkNature = Mid(TextLine, 377, 1)
            '!MBMOccupation = Mid(TextLine, 347, 30)
            !MBMBlood = Mid(TextLine, 351, 3)
            !MBMAllergic = Mid(TextLine, 381, 20)
            !MBMFirstMEMNo = txtTempMemberNo
            !MBMGroupCompany = Mid(TextLine, 422, 40)
            If Mid(txtFileVersion, 1, 2) = "01" Then
                !MBMExclusion = Mid(TextLine, 463, 60)
                !MBMSmoking = Mid(TextLine, 523, 1)
                !MBMAlcohol = Mid(TextLine, 524, 1)
                !MBMMaritalStatus = Mid(TextLine, 525, 1)
                !MBMPrevPolNo = Mid(TextLine, 541, 25)
                !MBMPrevMemNo = Mid(TextLine, 566, 18)
                !MBMEmployeeNo = Mid(TextLine, 584, 15)
                '!MBMDateJoin = Mid(TextLine, 599, 2) & "/" & Mid(TextLine, 601, 2) & "/" & Mid(TextLine, 603, 4)
                '!MBMClininal = Mid(TextLine, 580, 211)
            ElseIf Mid(txtFileVersion, 1, 2) = "02" Then
                !MBMExclusion = Mid(TextLine, 463, 2500)
                !MBMSmoking = Mid(TextLine, 2965, 1)
                !MBMAlcohol = Mid(TextLine, 2966, 1)
                !MBMMaritalStatus = Mid(TextLine, 2967, 1)
                !MBMBranch = Mid(TextLine, 2968, 15)
                !MBMPrevPolNo = Mid(TextLine, 2983, 25)
                !MBMPrevMemNo = Mid(TextLine, 3008, 18)
                !MBMEmployeeNo = Mid(TextLine, 3026, 15)
                !MBMDepartment = Mid(TextLine, 3250, 30)
                !MBMGRPCompanyAdd1 = Mid(TextLine, 3280, 30)
                !MBMGRPCompanyAdd2 = Mid(TextLine, 3310, 30)
                !MBMGRPCompanyAdd3 = Mid(TextLine, 3340, 30)
                !MBMGRPCompanyAdd4 = Mid(TextLine, 3370, 30)
                !MBMInstallment = Mid(TextLine, 3404, 1)
                !MBMPolSubNo1 = Mid(TextLine, 3405, 10)
                !MBMPOLSubNo2 = Mid(TextLine, 3415, 10)
                !MBMPOLIssuedDate = Mid(TextLine, 3425, 8)
                !MBMRiskNo = Mid(TextLine, 3433, 4)
                !MBMTranNo = Mid(TextLine, 3437, 5)
                !MBMPayNo = Mid(TextLine, 3442, 8)
                !MBMPAYSeqNo = Mid(TextLine, 3450, 2)
                !MBMClientNo = Mid(TextLine, 3452, 8)
                !MBMClientID = Mid(TextLine, 3460, 2)
                
            End If
                '!MBMDateJoin = Mid(TextLine, 599, 2) & "/" & Mid(TextLine, 601, 2) & "/" & Mid(TextLine, 603, 4)
                '!MBMClininal = Mid(TextLine, 580, 211)
            !MBMAccessFee = txtAccessFees
            !MBMPremium = txtMBMPremium
            !MBMMCOFee = txtMBMMCOFees
        End With
        MBMOthers.Update
    MBMOthers.Close
    Set MBMOthers = Nothing
 '   Exit Sub
'error1:
    
'    MsgBox ERR.Description
End Sub
