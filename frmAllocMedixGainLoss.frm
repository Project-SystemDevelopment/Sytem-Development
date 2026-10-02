VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form frmAllocMedixGainLoss 
   Caption         =   "MedixAllocGainLoss"
   ClientHeight    =   8025
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   11835
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   8025
   ScaleWidth      =   11835
   Begin VB.Frame Frame1 
      Height          =   1125
      Left            =   0
      TabIndex        =   39
      Top             =   240
      Width           =   11775
      Begin VB.CheckBox ChkScan 
         Caption         =   "Use Scan"
         Height          =   255
         Left            =   3240
         TabIndex        =   1
         Top             =   180
         Value           =   1  'Checked
         Width           =   1095
      End
      Begin VB.TextBox TxtWSNoTo 
         Height          =   310
         Left            =   3600
         MaxLength       =   15
         TabIndex        =   2
         Top             =   160
         Visible         =   0   'False
         Width           =   1560
      End
      Begin VB.TextBox TxtWSNo 
         Height          =   310
         Left            =   1560
         MaxLength       =   15
         TabIndex        =   0
         Top             =   160
         Width           =   1560
      End
      Begin VB.ComboBox cboHLTCode 
         Height          =   315
         Left            =   1560
         Sorted          =   -1  'True
         TabIndex        =   3
         Top             =   400
         Width           =   3615
      End
      Begin VB.ComboBox CboPayor 
         Height          =   315
         Left            =   1560
         TabIndex        =   4
         Top             =   690
         Width           =   3615
      End
      Begin VB.ComboBox CboAllocStatus 
         Height          =   315
         Left            =   8040
         TabIndex        =   5
         Top             =   180
         Width           =   3015
      End
      Begin VB.TextBox txtPVNofr 
         Height          =   285
         Left            =   8040
         MaxLength       =   15
         TabIndex        =   6
         Top             =   465
         Width           =   1215
      End
      Begin VB.TextBox txtPVNoto 
         Height          =   285
         Left            =   9720
         MaxLength       =   15
         TabIndex        =   7
         Top             =   465
         Width           =   1335
      End
      Begin MSMask.MaskEdBox txtPVDatefr 
         Height          =   315
         Left            =   8040
         TabIndex        =   8
         Top             =   720
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   556
         _Version        =   393216
         MaxLength       =   10
         Mask            =   "##/##/####"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox txtPVDateto 
         Height          =   315
         Left            =   9720
         TabIndex        =   9
         Top             =   720
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   556
         _Version        =   393216
         MaxLength       =   10
         Mask            =   "##/##/####"
         PromptChar      =   "_"
      End
      Begin VB.Label Label35 
         Caption         =   "Health Type"
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
         TabIndex        =   48
         Top             =   480
         Width           =   1215
      End
      Begin VB.Label Label4 
         Caption         =   "Clm No"
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
         TabIndex        =   47
         Top             =   195
         Width           =   1215
      End
      Begin VB.Label Label25 
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
         Left            =   3240
         TabIndex        =   46
         Top             =   160
         Width           =   375
      End
      Begin VB.Label Label19 
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
         Left            =   9360
         TabIndex        =   45
         Top             =   765
         Width           =   375
      End
      Begin VB.Label Label17 
         Caption         =   "PV Date"
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
         Left            =   6840
         TabIndex        =   44
         Top             =   810
         Width           =   855
      End
      Begin VB.Label Label16 
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
         Left            =   9360
         TabIndex        =   43
         Top             =   525
         Width           =   375
      End
      Begin VB.Label Label15 
         Caption         =   "PV No"
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
         Left            =   6840
         TabIndex        =   42
         Top             =   525
         Width           =   855
      End
      Begin VB.Label Label10 
         Caption         =   "Alloc Status"
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
         Left            =   6840
         TabIndex        =   41
         Top             =   225
         Width           =   1215
      End
      Begin VB.Label Label18 
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
         Left            =   120
         TabIndex        =   40
         Top             =   795
         Width           =   1215
      End
   End
   Begin VB.Frame Frame2 
      Height          =   550
      Left            =   0
      TabIndex        =   28
      Top             =   7080
      Width           =   11820
      Begin VB.TextBox lblMedixAmt 
         Height          =   285
         Left            =   5040
         TabIndex        =   56
         Top             =   240
         Width           =   1095
      End
      Begin VB.TextBox LblAmt 
         Height          =   285
         Left            =   2400
         TabIndex        =   55
         Top             =   120
         Width           =   1215
      End
      Begin VB.TextBox lblCurrent 
         Height          =   285
         Left            =   120
         TabIndex        =   54
         Top             =   120
         Width           =   495
      End
      Begin VB.CommandButton CmdSave 
         Caption         =   "Sa&ve"
         Height          =   330
         Left            =   6960
         TabIndex        =   36
         Top             =   180
         Width           =   615
      End
      Begin VB.CommandButton cmdAdd 
         Caption         =   "&Add"
         Height          =   330
         Left            =   6360
         TabIndex        =   35
         Top             =   180
         Width           =   615
      End
      Begin VB.CommandButton CmdPrint 
         Caption         =   "Print to &File"
         Height          =   330
         Left            =   9600
         TabIndex        =   34
         Top             =   180
         Width           =   975
      End
      Begin VB.CommandButton CmdUpdate 
         Caption         =   "&Update"
         Enabled         =   0   'False
         Height          =   330
         Left            =   8160
         TabIndex        =   33
         Top             =   180
         Width           =   735
      End
      Begin VB.CommandButton CmdEdit 
         Caption         =   "&Edit"
         Height          =   330
         Left            =   7560
         TabIndex        =   32
         Top             =   180
         Width           =   615
      End
      Begin VB.CommandButton CmdExit 
         Cancel          =   -1  'True
         Caption         =   "E&xit"
         Height          =   330
         Left            =   11160
         TabIndex        =   31
         Top             =   180
         Width           =   615
      End
      Begin VB.CommandButton CmdClear 
         Caption         =   "&Clear"
         Height          =   330
         Left            =   10560
         TabIndex        =   30
         Top             =   180
         Width           =   615
      End
      Begin VB.CommandButton CmdSearch 
         Caption         =   "&Search"
         Height          =   330
         Left            =   8880
         TabIndex        =   29
         Top             =   180
         Width           =   735
      End
      Begin VB.Label Label11 
         Caption         =   "Records"
         Height          =   255
         Left            =   720
         TabIndex        =   38
         Top             =   240
         Width           =   735
      End
      Begin VB.Label Label29 
         Caption         =   "Checked Amt"
         Height          =   255
         Left            =   1440
         TabIndex        =   37
         Top             =   240
         Width           =   975
      End
      Begin VB.Label Label8 
         Caption         =   "Medix Amt"
         Height          =   255
         Left            =   4080
         TabIndex        =   53
         Top             =   240
         Width           =   975
      End
   End
   Begin VB.PictureBox lstMedixGainLoss 
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   4620
      Left            =   0
      ScaleHeight     =   4560
      ScaleWidth      =   11715
      TabIndex        =   49
      Top             =   1440
      Width           =   11775
   End
   Begin VB.Frame FramePv 
      Height          =   1080
      Left            =   0
      TabIndex        =   19
      Top             =   6000
      Width           =   11820
      Begin VB.TextBox txtAllocNo 
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
         MaxLength       =   15
         TabIndex        =   10
         Top             =   150
         Width           =   1215
      End
      Begin VB.TextBox TxtPVNo 
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
         MaxLength       =   15
         TabIndex        =   11
         Top             =   405
         Width           =   1215
      End
      Begin VB.ComboBox CboACPeriodMth 
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
         ItemData        =   "frmAllocMedixGainLoss.frx":0000
         Left            =   5760
         List            =   "frmAllocMedixGainLoss.frx":0002
         TabIndex        =   16
         Top             =   180
         Width           =   975
      End
      Begin VB.ComboBox CboACPeriodYr 
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
         ItemData        =   "frmAllocMedixGainLoss.frx":0004
         Left            =   7200
         List            =   "frmAllocMedixGainLoss.frx":0006
         TabIndex        =   17
         Top             =   180
         Width           =   975
      End
      Begin VB.TextBox TxtRemark 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   5760
         MaxLength       =   200
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   18
         Top             =   470
         Width           =   2415
      End
      Begin MSMask.MaskEdBox txtPVDate 
         Height          =   285
         Left            =   1080
         TabIndex        =   12
         Top             =   645
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   503
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
      Begin VB.TextBox txtChequeNo 
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
         Left            =   3360
         MaxLength       =   20
         TabIndex        =   13
         Top             =   150
         Width           =   1335
      End
      Begin MSMask.MaskEdBox txtChequeDate 
         Height          =   285
         Left            =   3360
         TabIndex        =   14
         Top             =   405
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   503
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
      Begin VB.TextBox txtChequeAmt 
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
         Left            =   3360
         MaxLength       =   10
         TabIndex        =   15
         Top             =   645
         Width           =   1335
      End
      Begin VB.Label Label1 
         Caption         =   "Alloc. No"
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
         TabIndex        =   52
         Top             =   150
         Width           =   1095
      End
      Begin VB.Label Label7 
         Caption         =   "PV. Remark"
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
         Left            =   4800
         TabIndex        =   27
         Top             =   480
         Width           =   855
      End
      Begin VB.Label Label22 
         Caption         =   "Chq. Date"
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
         Left            =   2520
         TabIndex        =   26
         Top             =   435
         Width           =   975
      End
      Begin VB.Label Label23 
         Caption         =   "A/C Mth"
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
         Left            =   4800
         TabIndex        =   25
         Top             =   225
         Width           =   735
      End
      Begin VB.Label Label24 
         Caption         =   "Year"
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
         Left            =   6840
         TabIndex        =   24
         Top             =   240
         Width           =   375
      End
      Begin VB.Label Label6 
         Caption         =   "Chq Amt"
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
         Left            =   2520
         TabIndex        =   23
         Top             =   675
         Width           =   735
      End
      Begin VB.Label Label13 
         Caption         =   "Chq. No"
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
         Left            =   2520
         TabIndex        =   22
         Top             =   150
         Width           =   975
      End
      Begin VB.Label Label5 
         Caption         =   "MBB PV Date"
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
         Top             =   675
         Width           =   975
      End
      Begin VB.Label Label3 
         Caption         =   "MBB PV No"
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
         TabIndex        =   20
         Top             =   435
         Width           =   1095
      End
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
      TabIndex        =   51
      Top             =   7680
      Width           =   1575
   End
   Begin VB.Label Label28 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00404040&
      Caption         =   "Medix Allocation (Gain/Loss)"
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
      TabIndex        =   50
      Top             =   0
      Width           =   11805
   End
End
Attribute VB_Name = "frmAllocMedixGainLoss"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim intMBMAccount As Integer
Dim intExit As Integer
Public searchCriteria As String
Dim one As Boolean
Dim ScanCode As String
Dim listloop1 As Integer
Dim listloop2 As Integer
Dim Check As Boolean
Dim RecordSelected As Integer
Dim ItemCheck As Boolean
Dim USRRec As New ADODB.Recordset
Dim USRSql As String
Dim Authorize As String
Dim EditFlag As Boolean
Private m_blnDirection(1 To 15) As Boolean

Private Sub ComboListing()
On Error Resume Next
Dim PAY As New ADODB.Recordset
Dim MonthStart
Dim YearStart
Dim HLT As New ADODB.Recordset
Dim conn As New ADODB.Connection
Dim cmd As New ADODB.Command
Dim prm1 As ADODB.Parameter
Dim prm2 As ADODB.Parameter
    medixmasterconn.Open "File Name=" & BOSPath & "Conn\mediconnect.UDL"
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
            CboPayor.AddItem !PAYCode & " - " & Trim(!PAYCompanyName)
        End With
        PAY.MoveNext
    Loop
    PAY.Close
    Set PAY = Nothing
    
    Set cmd.ActiveConnection = medixmasterconn
    cmd.CommandText = "SearchHLTType"
    cmd.CommandType = adCmdStoredProc
    cmd.CommandTimeout = 15
    PAY.CursorLocation = adUseClient
    PAY.CursorType = adOpenForwardOnly
    PAY.CacheSize = 100
    Set HLT = cmd.Execute
    Do Until HLT.EOF
        With HLT
            cboHLTCode.AddItem ChkStr(!HLTCode) & " - " & ChkStr(!HLTDescription)
            HLT.MoveNext
        End With
    Loop
    HLT.Close
    Set HLT = Nothing
    
    With CboAllocStatus
        .AddItem "N - No"
        .AddItem "Y - Yes"
    End With
    
    For MonthStart = 1 To 12
        With CboACPeriodMth
            .AddItem Format(MonthStart, "00")
        End With
    Next MonthStart
    
    For YearStart = 1999 To 3000
        With CboACPeriodYr
            .AddItem YearStart
        End With
    Next YearStart
    
    medixmasterconn.Close
    Set medixmasterconn = Nothing

End Sub

Private Sub CboAllocStatus_KeyPress(KeyAscii As Integer)
Dim NewText As String
Dim ValidCount As Integer
Dim ValidValue As String
Dim i As Integer
    If KeyAscii >= 32 And KeyAscii <> 127 Then
    NewText = LCase(Left(CboAllocStatus, CboAllocStatus.SelStart) + Chr(KeyAscii) + Mid(CboAllocStatus, CboAllocStatus.SelStart + CboAllocStatus.SelLength + 1))
    ValidCount = 0
    ValidValue = ""
    For i = 0 To CboAllocStatus.ListCount - 1
        If NewText = LCase(Left(CboAllocStatus.List(i), Len(NewText))) Then
                ValidCount = ValidCount + 1
                ValidValue = CboAllocStatus.List(i)
        End If
        Next
        If ValidCount <= 1 Then KeyAscii = 0
            If ValidCount = 1 Then
               CboAllocStatus = ValidValue
               CboAllocStatus.SelStart = 0
               CboAllocStatus.SelLength = Len(ValidValue)
            End If
        End If
End Sub

Private Sub CboHLTCode_KeyPress(KeyAscii As Integer)
Dim NewText As String
Dim ValidCount As Integer
Dim ValidValue As String
Dim i As Integer
    If KeyAscii >= 32 And KeyAscii <> 127 Then
    NewText = LCase(Left(cboHLTCode.Text, cboHLTCode.SelStart) + Chr(KeyAscii) + Mid(cboHLTCode.Text, cboHLTCode.SelStart + cboHLTCode.SelLength + 1))
    ValidCount = 0
    ValidValue = ""
    For i = 0 To cboHLTCode.ListCount - 1
        If NewText = LCase(Left(cboHLTCode.List(i), Len(NewText))) Then
                ValidCount = ValidCount + 1
                ValidValue = cboHLTCode.List(i)
        End If
        Next
        If ValidCount <= 1 Then KeyAscii = 0
            If ValidCount = 1 Then
               cboHLTCode.Text = ValidValue
               cboHLTCode.SelStart = 0
               cboHLTCode.SelLength = Len(ValidValue)
            End If
        End If
End Sub

Private Sub CboPayor_KeyPress(KeyAscii As Integer)
Dim NewText As String
Dim ValidCount As Integer
Dim ValidValue As String
Dim i As Integer
    If KeyAscii >= 32 And KeyAscii <> 127 Then
    NewText = LCase(Left(CboPayor, CboPayor.SelStart) + Chr(KeyAscii) + Mid(CboPayor, CboPayor.SelStart + CboPayor.SelLength + 1))
    ValidCount = 0
    ValidValue = ""
    For i = 0 To CboPayor.ListCount - 1
        If NewText = LCase(Left(CboPayor.List(i), Len(NewText))) Then
                ValidCount = ValidCount + 1
                ValidValue = CboPayor.List(i)
        End If
        Next
        If ValidCount <= 1 Then KeyAscii = 0
            If ValidCount = 1 Then
               CboPayor = ValidValue
               CboPayor.SelStart = 0
               CboPayor.SelLength = Len(ValidValue)
            End If
        End If

End Sub

Private Sub ChkScan_Click()
    lstMedixGainLoss.ListItems.Clear
    If ChkScan.Value = 1 Then
        Label25.Visible = False
        txtWSNoTo.Visible = False
        ChkScan.Left = 3240
        TxtWSNo = ""
        TxtWSNo.SetFocus
    Else
        Label25.Visible = True
        txtWSNoTo.Visible = True
        ChkScan.Left = 5280
        TxtWSNo.SetFocus
    End If
End Sub

Private Sub cmdAdd_Click()
Dim zz As Integer
Dim ItemChk As Boolean
    ItemChk = False
    For zz = 1 To lstMedixGainLoss.ListItems.Count
        If lstMedixGainLoss.ListItems(zz).Checked = True Then
            ItemChk = True
        Exit For
        End If
    Next zz
    cmdUpdate.Enabled = False
    Call ClearFrameCase
    If ItemChk Then
        cmdSave.Enabled = True
        FramePv.Enabled = True
        txtPVNo.SetFocus
    Else
        MsgBox "Please checked at least 1 Claim  ! ", vbCritical + vbOKOnly, DialogBoxTitle
    End If
End Sub

Private Sub CmdClear_Click()
    Call ClearForm(Me)
    lstMedixGainLoss.ListItems.Clear
    lblcurrent = 0
    LblAmt = 0
    lblMedixAmt = 0
    one = False
    ChkScan.Value = 1
    EditFlag = False
End Sub

Private Sub cmdEdit_Click()
Dim zz As Integer
Dim ItemChk As Boolean
    EditFlag = True
    ItemChk = False
    For zz = 1 To lstMedixGainLoss.ListItems.Count
        If lstMedixGainLoss.ListItems(zz).Checked = True Then
            ItemChk = True
        Exit For
        End If
    Next zz
    cmdSave.Enabled = False
    cmdUpdate.Enabled = False
    FramePv.Enabled = False
    If ItemChk = False And Len(Trim(txtAllocNo)) Then
        cmdUpdate.Enabled = True
        FramePv.Enabled = True
        txtPVNo.SetFocus
    End If
End Sub

Private Sub cmdExit_Click()
    Unload Me
End Sub

Private Sub cmdPrint_Click()
    Screen.MousePointer = vbHourglass
    If lstMedixGainLoss.ListItems.Count <> 0 Then
        Call ExportListViewtoExcel(lstMedixGainLoss)
    Else
        MsgBox "Report cannot be printed without any record.", vbInformation
    End If
    Screen.MousePointer = vbDefault
End Sub

Private Sub cmdSave_Click()
On Error Resume Next
Dim CASNo As String
Dim CASVer As String
Dim status As String
Dim InvoiceNo As String
Dim Update
Dim ItemChecked As Boolean
Dim i As Integer
Dim MedixPaymentAmt As Double
Dim PVSql As String
Dim AllocNo  As String
    ItemChecked = False
    CASNo = ""
    CASVer = ""
    status = ""
    MedixPaymentAmt = 0
    cmdSave.Enabled = False
    medixmasterconnPayment.Open "File Name=" & BOSPath & "Conn\mediconnectPayment.UDL"
    medixmasterconnMaint.Open "File Name=" & BOSPath & "Conn\MediConnHISMaint.UDL"
    If lstMedixGainLoss.ListItems.Count = 0 Then
        MsgBox "Please select the record", vbOKOnly + vbCritical
    ElseIf txtPVDate = "__/__/____" Then
        MsgBox "Please Enter PV Date!", vbCritical
        txtPVDate.SetFocus
    ElseIf IsDate(txtPVDate) = False Then
        MsgBox "Invalid PV Date!", vbCritical
        txtPVDate.SetFocus
    ElseIf Trim(CboACPeriodMth) = "" Then
        MsgBox "Please Keyin Account Month ! ", vbCritical
        CboACPeriodMth.SetFocus
    ElseIf Trim(CboACPeriodYr) = "" Then
        MsgBox "Please Keyin Account Year ! ", vbCritical
        CboACPeriodYr.SetFocus
    ElseIf Trim(txtChequeAmt) = "" Or IsNumeric(txtChequeAmt) = False Then
        MsgBox "Please Enter Cheque Amt!", vbCritical
        txtChequeAmt.SetFocus
    Else
        Update = MsgBox("Do you want to Save record(s) ?", vbYesNo + vbExclamation)
        If Update = vbYes Then
                AllocNo = GenerateAllocNo("AM")
                txtAllocNo = ""
                txtAllocNo = AllocNo
            For i = 1 To lstMedixGainLoss.ListItems.Count
                If lstMedixGainLoss.ListItems(i).Checked = True Then
                    ItemChecked = True
                Exit For
                End If
            Next i
            If ItemChecked = True Then
                For i = 1 To lstMedixGainLoss.ListItems.Count
                    If lstMedixGainLoss.ListItems(i).Checked = True Then
                        CASNo = Trim(Mid(lstMedixGainLoss.ListItems(i).SubItems(2), 1, IIf(InStr(1, lstMedixGainLoss.ListItems(i).SubItems(2), "-") = 0, 10, InStr(1, lstMedixGainLoss.ListItems(i).SubItems(2), "-") - 1)))
                        CASVer = Trim(Mid(lstMedixGainLoss.ListItems(i).SubItems(2), InStr(1, lstMedixGainLoss.ListItems(i).SubItems(2), "-") + 1, Len(lstMedixGainLoss.ListItems(i).SubItems(2)) - InStr(1, lstMedixGainLoss.ListItems(i).SubItems(2), "-")))
                        MedixPaymentAmt = Format(IIf(IsNumeric(lstMedixGainLoss.ListItems(i).SubItems(6)), lstMedixGainLoss.ListItems(i).SubItems(6), 0), "#,##0.00")
                        Call UpdateMedixGL(CASNo, CInt(CASVer), MedixPaymentAmt)
                        CASNo = ""
                        CASVer = ""
                        status = ""
                        MedixPaymentAmt = 0
                    End If
                Next i
            MsgBox "Record Saved...", vbOKOnly + vbInformation, DialogBoxTitle
            lstMedixGainLoss.ListItems.Clear
            PVSql = " and PVNo = '" & Trim(txtPVNo) & "'"
            Call MedixGLListing(PVSql)
            End If
            EditFlag = False
        End If
    End If
    medixmasterconnPayment.Close
    Set medixmasterconnPayment = Nothing
    medixmasterconnMaint.Close
    Set medixmasterconnMaint = Nothing
    cmdSave.Enabled = True
End Sub

Private Sub cmdSearch_Click()
On Error Resume Next
    Dim listrecord As Integer
    Screen.MousePointer = vbHourglass
    cmdSearch.Enabled = False
    medixmasterconnPayment.Open "File Name=" & BOSPath & "Conn\mediconnectPayment.UDL"
    LblAmt = 0
    lblMedixAmt = 0
        If ChkScan.Value = 1 Then
            If TxtWSNo = "" Then
                MsgBox "Please Enter Worksheet No", vbOKOnly + vbCritical, DialogBoxTitle
                TxtWSNo.SetFocus
            Else
                Screen.MousePointer = vbHourglass
                listrecord = lstMedixGainLoss.ListItems.Count
                If listrecord <> 0 Then
                    For listloop1 = 1 To listrecord
                        Call CaseNoCheck
                    Next listloop1
                End If
                    If Check = True Then
                        MsgBox "Data has been uploaded!", vbOKOnly + vbCritical
                    ElseIf Check = False Then
                        Call SearchRecord
                    End If
                    Check = False
                Screen.MousePointer = Default
                TxtWSNo = ""
                TxtWSNo.SetFocus
            End If
            TxtWSNo.SetFocus
        Else
            Call SearchRecord
        End If
    medixmasterconnPayment.Close
    Set medixmasterconnPayment = Nothing
    cmdSearch.Enabled = True
    Screen.MousePointer = Default
End Sub

Private Sub CmdStop_Click()

End Sub

Private Sub CmdUpdate_Click()
On Error Resume Next
Dim RecordSaved
Dim strsql As String
Dim zz As Integer
Dim ItemChk As Boolean
Dim AmtPaid As Variant
Dim MedixPaymentAmt As Double
Dim CASRemark As String
Dim TempCase As String
Dim TempVer As String
Dim TempPVNo As String
Dim PVSql As String
    cmdUpdate.Enabled = False
    ItemChk = False
    medixmasterconnPayment.Open "File Name=" & BOSPath & "Conn\mediconnectPayment.UDL"
    If lstMedixGainLoss.ListItems.Count = 0 Then
        MsgBox "Please select the record", vbOKOnly + vbCritical
        ElseIf txtPVNo = "" Then
            MsgBox "Please Confirm PV No!", vbCritical
        ElseIf IsDate(txtPVDate) = False Then
            MsgBox "Invalid PV Date!", vbCritical
            txtPVDate.SetFocus
        ElseIf txtPVDate = "__/__/____" Then
            MsgBox "Please Enter PV Date! ", vbCritical
            txtPVDate.SetFocus
        ElseIf Trim(txtChequeAmt) = "" Or IsNumeric(txtChequeAmt) = False Then
            MsgBox "Please Enter Cheque Amt!", vbCritical
            txtChequeAmt.SetFocus
        Else
            RecordSaved = MsgBox("Do you want to update the record?", vbYesNo + vbExclamation)
            If RecordSaved = vbYes Then
                TempCase = Trim(Mid(lstMedixGainLoss.SelectedItem.SubItems(2), 1, IIf(InStr(1, lstMedixGainLoss.SelectedItem.SubItems(2), "-") = 0, 10, InStr(1, lstMedixGainLoss.SelectedItem.SubItems(2), "-") - 1)))
                TempVer = Trim(Mid(lstMedixGainLoss.SelectedItem.SubItems(2), InStr(1, lstMedixGainLoss.SelectedItem.SubItems(2), "-") + 1, Len(lstMedixGainLoss.SelectedItem.SubItems(2)) - InStr(1, lstMedixGainLoss.SelectedItem.SubItems(2), "-")))
                Call UpdateMedixGL(TempCase, CInt(TempVer), "0")
            End If
            MsgBox "Record updated !", vbOKOnly
            EditFlag = False
            cmdUpdate.Enabled = False
            Call ClearFrameCase
            lstMedixGainLoss.ListItems.Clear
    End If
    medixmasterconnPayment.Close
    Set medixmasterconnPayment = Nothing
    cmdUpdate.Enabled = True
End Sub

Private Sub Form_Load()
    Call ComboListing
End Sub

Private Sub ClearFrameCase()
    txtAllocNo = ""
    txtPVNo = ""
    txtPVDate = "__/__/____"
    txtChequeAmt = ""
    txtChequeNo = ""
    txtChequeDate.Text = "__/__/____"
    TxtRemark = ""
    CboACPeriodMth.Text = ""
    CboACPeriodYr.Text = ""
    FramePv.Enabled = False
End Sub

Private Sub lstMedixGainLoss_Click()
Dim rst3 As New ADODB.Recordset
Dim strsql As String
    ItemCheck = False
    cmdSave.Enabled = False
    EditFlag = False
    Call LvDisplayData(, False)
End Sub

Private Sub lstMedixGainLoss_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
    m_blnDirection(ColumnHeader.Index) = Not m_blnDirection(ColumnHeader.Index)
    Select Case ColumnHeader.Index
    Case 1
        SortListView lstMedixGainLoss, ColumnHeader.Index, ldtString, m_blnDirection(1)
    Case 2
        SortListView lstMedixGainLoss, ColumnHeader.Index, ldtString, m_blnDirection(2)
    Case 3
        SortListView lstMedixGainLoss, ColumnHeader.Index, ldtString, m_blnDirection(3)
    Case 4
        SortListView lstMedixGainLoss, ColumnHeader.Index, ldtDateTime, m_blnDirection(4)
    Case 5
        SortListView lstMedixGainLoss, ColumnHeader.Index, ldtNumber, m_blnDirection(5)
    Case 6
        SortListView lstMedixGainLoss, ColumnHeader.Index, ldtNumber, m_blnDirection(6)
    Case 7
        SortListView lstMedixGainLoss, ColumnHeader.Index, ldtString, m_blnDirection(7)
    Case 8
        SortListView lstMedixGainLoss, ColumnHeader.Index, ldtDateTime, m_blnDirection(8)
    Case 9
        SortListView lstMedixGainLoss, ColumnHeader.Index, ldtString, m_blnDirection(9)
    Case 10
        SortListView lstMedixGainLoss, ColumnHeader.Index, ldtString, m_blnDirection(10)
    Case 11
        SortListView lstMedixGainLoss, ColumnHeader.Index, ldtDateTime, m_blnDirection(11)
    Case 12
        SortListView lstMedixGainLoss, ColumnHeader.Index, ldtNumber, m_blnDirection(12)
    Case 13
        SortListView lstMedixGainLoss, ColumnHeader.Index, ldtString, m_blnDirection(13)
    Case 14
        SortListView lstMedixGainLoss, ColumnHeader.Index, ldtString, m_blnDirection(14)
    End Select
End Sub

Private Sub lstMedixGainLoss_ItemCheck(ByVal Item As MSComctlLib.ListItem)
    cmdSave.Enabled = True
    one = False
    ItemCheck = True
    Call ProcessCheckedItem(Item)
End Sub

Private Sub txtPVDate_LostFocus()
    If IsDate(txtPVDate) Then
        CboACPeriodMth.Text = Format(txtPVDate, "mm")
        CboACPeriodYr.Text = Format(txtPVDate, "yyyy")
    Else
        If txtPVDate <> "__/__/____" Then
            MsgBox "Invalid Date Format! ", vbCritical
            txtPVDate.SetFocus
        End If
    End If
End Sub

Private Function CaseNoCheck()
    If Trim(ChkStr(TxtWSNo)) = Trim(Mid(lstMedixGainLoss.ListItems(listloop1).SubItems(2), 1, IIf(InStr(1, lstMedixGainLoss.ListItems(listloop1).SubItems(2), "-") = 0, 10, InStr(1, lstMedixGainLoss.ListItems(listloop1).SubItems(2), "-") - 1))) Then
        Check = True
    End If
End Function

Private Function SearchRecord()
On Error Resume Next
Dim strsql As String
Dim strAppend As String
Dim sql As String
    strAppend = ""
    If Len(Trim(cboHLTCode)) Then
        strAppend = strAppend + " And HLTCode = '" & Mid(cboHLTCode, 1, 1) & "'"
    End If
    If ChkScan.Value = 1 Then
        strAppend = strAppend + " AND CASNumber = '" & Trim(TxtWSNo) & "'"
    Else
        If Len(TxtWSNo) Then
            strAppend = strAppend + " AND CASNumber >= '" & Trim(TxtWSNo) & "'"
        End If
    
        If Len(txtWSNoTo) Then
            strAppend = strAppend + " AND CASNumber <= '" & Trim(txtWSNoTo) & "'"
        End If
    End If
    If Len(Trim(CboPayor)) Then
        strAppend = strAppend + " And PAYCode = '" & Trim(Mid(CboPayor, 1, IIf(InStr(1, CboPayor, "-") = 0, 2, InStr(1, CboPayor, "-") - 1))) & "'"
    End If
    If Len(txtPVNoFr) Then
        strAppend = strAppend + " and PAYPVNo >= '" & Trim(txtPVNoFr) & "'"
    End If
    If Len(txtPVNoTo) Then
        strAppend = strAppend + " and PAYPVNo <= '" & Trim(txtPVNoTo) & "'"
    End If
    
    If (txtPVDateFr) <> "__/__/____" And IsDate(txtPVDateFr) = True _
        And (txtPVDateTo) <> "__/__/____" And IsDate(txtPVDateTo) = True Then
        sql = ""
        sql = " AND PayPVDate >= '" & Format(Trim(txtPVDateFr), "mm/dd/yyyy") & _
              "' And  PayPVDate <= '" & Format(Trim(txtPVDateTo), "mm/dd/yyyy") & "'"
        strAppend = strAppend + sql
    End If
    If Len(Trim(txtPVDateFr)) > 0 And IsDate(txtPVDateFr) = True Then
        strAppend = strAppend + " And PayPVDate >= '" & Format(txtPVDateFr, "mm/dd/yyyy") & "'"
    End If
    If Len(Trim(txtPVDateTo)) > 0 And IsDate(txtPVDateTo) = True Then
        strAppend = strAppend + " And PayPVDate <= '" & Format(txtPVDateTo, "mm/dd/yyyy") & "'"
    End If
    If Len(CboAllocStatus) Then
        If Mid(CboAllocStatus, 1, 1) = "Y" Then
            strAppend = strAppend + " And PAYAllocStatus = '" & Mid(CboAllocStatus, 1, 1) & "'"
        Else
            strAppend = strAppend + " And (PAYAllocStatus = '' or PAYAllocStatus is null) "
        End If
    End If
    searchCriteria = strAppend
    Call MedixGLListing(strAppend)
End Function

Private Sub MedixGLListing(Optional strAppend As String)
On Error Resume Next
Dim rst2 As New ADODB.Recordset
Dim strsqlPay As String
Dim PAY As New ADODB.Recordset
Dim INVAmt As Variant
Dim CasNumber As String
Dim WSVersion As String
Dim Record As ListItem
Dim sql As String
Dim total As String
Dim Current As String
Dim Check As Boolean
Dim RECFull As Boolean
Dim RECPartial As String
Dim MedixPaidMBM As Double
Dim MedixPaidHOS As Double
Dim Paid2Medix  As Double
Dim Paid2Others As Double
Dim TotalPaid As Double
Dim BankCharges As Double
Dim RecMBM As Double
Dim RecMNRB As Double
Dim RecOthers As Double
Dim totalREc As Double
Dim CheckTotal As Double
Dim RecMedix As Double
Dim Contra As Double
Dim CashPos As Double
Dim Pay2Hos As Double
Dim Pay2MBM As Double
Dim TotalApp As Double
Dim Pay2Medix As Double
Dim Pay2Others As Double
Dim GainLoss As Double
    total = 0
    Current = 0
    lstMedixGainLoss.ListItems.Clear
    medixmasterconnPayment.Open "File Name=" & BOSPath & "Conn\mediconnectPayment.UDL"
    sql = " Select CASNumber, ClaimVersion, MBMAmt, MNRBAmt, CLMTotalApproved," & _
            " HOSPaidAmt, MBMPaidAmt, CLMPayableByMedix2H, CLMPayableByMedix2M, " & _
            "ContraCaseAmt, PaymentBankCharge, PVHOSDate,PVDate, MBMReceiptDate, " & _
            "MNRBReceiptDate, HLTCode, PAYAllocStatus, PAYAllocNo, PAYPVNo, PAYPVDate, " & _
            "PAYChqNo, PAYChqDate, PAYChqAmt, PAYRemarks,MedixPAYAmt, " & _
            "PAYAccPeriodMth, PAYAccPeriodYr, PAYLastUpdateUser, PAYLastUpdateDate " & _
            "From VIEW_Medix_GL where 0 = 0 "
    If Len(Trim(strAppend)) Then
        sql = sql + strAppend
    End If
    rst2.Open sql, medixmasterconnPayment, adOpenKeyset, adLockOptimistic
    If rst2.recordCount = 0 Then
        MsgBox "No record found!", vbOKOnly + vbCritical, DialogBoxTitle
    Else
    
        Do Until rst2.EOF
            With rst2
                RecMedix = 0
                RecOthers = 0
                Paid2Others = 0
                If Len(rst2!HOSPaidAmt) Then
                    MedixPaidHOS = Format(rst2!HOSPaidAmt, "#,##0.00;(#,##0.00)")
                Else
                    MedixPaidHOS = 0
                End If
                MedixPaidMBM = IIf(Len(rst2!MBMPaidAmt), rst2!MBMPaidAmt, 0)
                If Len(rst2!MedixPayAmt) Then
                    Paid2Medix = Format(rst2!MedixPayAmt, "#,##0.00;(#,##0.00)")
                Else
                    Paid2Medix = 0
                End If
                Pay2Hos = IIf(Len(rst2!CLMPayableByMedix2H), rst2!CLMPayableByMedix2H, 0)
                Pay2MBM = IIf(Len(rst2!CLMPayableByMedix2M), rst2!CLMPayableByMedix2M, 0)
                TotalApp = IIf(Len(rst2!CLMTotalApproved), rst2!CLMTotalApproved, 0)
                Pay2Medix = Format(CDbl(RecMedix) - CDbl(Paid2Medix), "#,##0.00;(#,##0.00)")
                Pay2Others = Format(CDbl(RecOthers) - CDbl(Paid2Others), "#,##0.00;(#,##0.00)")
                GainLoss = Format(CDbl((Pay2Hos) + CDbl(Pay2MBM) + CDbl(Pay2Medix) + CDbl(Pay2Others)) - CDbl(TotalApp), "#,##0.00;(#,##0.00)")
                CheckTotal = Format(GainLoss, "#,##0.00;(#,##0.00)")
                If GainLoss <> "0" Or Len(rst2!MedixPayAmt) Then
                    Set Record = lstMedixGainLoss.ListItems.Add(, , "")
                    If Trim(Len(!PAYAllocStatus)) Then
                        Record.SubItems(1) = "Allocated"
                    Else
                        Record.SubItems(1) = ""
                    End If
                    If Len(rst2!CasNumber & "-" & rst2!ClaimVersion) Then
                        Record.SubItems(2) = Trim(rst2!CasNumber & "-" & rst2!ClaimVersion)
                    End If
                    If Len(rst2!PayPVNo) Then
                        Record.SubItems(3) = Trim(rst2!PayPVNo)
                    End If
                    If Len(rst2!PayPVDate) Then
                        Record.SubItems(4) = Trim(rst2!PayPVDate)
                    End If
                    If Trim(Len(rst2!PAYAllocStatus)) Then
                        Record.SubItems(5) = 0
                    Else
                        If Len(GainLoss) Then
                            Record.SubItems(5) = Format(GainLoss, "#,##0.00;(#,##0.00)")
                        End If
                    End If
                    If Len(rst2!MedixPayAmt) Then
                        Record.SubItems(6) = Format(rst2!MedixPayAmt, "#,##0.00;(#,##0.00)")
                        lblMedixAmt = CDbl(lblMedixAmt) + CDbl(rst2!MedixPayAmt)
                    End If
                    If Len(rst2!PAYAllocNo) Then
                        Record.SubItems(7) = rst2!PAYAllocNo
                    End If
                    If Len(rst2!PayLastUpdateDate) Then
                        Record.SubItems(8) = rst2!PayLastUpdateDate
                    End If
                    If Len(rst2!PAYLastUpdateUser) Then
                        Record.SubItems(9) = rst2!PAYLastUpdateUser
                    End If
                    
                    If Len(rst2!PAYChqNo) Then
                        Record.SubItems(10) = rst2!PAYChqNo
                    End If
                    
                    If Len(rst2!PAYChqDate) Then
                        Record.SubItems(11) = rst2!PAYChqDate
                    End If
                    
                    If Len(rst2!PAYChqAmt) Then
                        Record.SubItems(12) = rst2!PAYChqAmt
                    End If
                    
                    If Len(rst2!PayAccPeriodMth) Then
                        Record.SubItems(13) = rst2!PayAccPeriodMth & "/" & rst2!PayAccPeriodYr
                    End If
                    
                    If Len(rst2!PayRemarks) Then
                        Record.SubItems(14) = rst2!PayRemarks
                    End If
                    Current = Current + 1
                End If
                lblcurrent = Current
            End With
            rst2.MoveNext
        Loop
    End If
    intExit = 0
    rst2.Close
    Set rst2 = Nothing
    medixmasterconnPayment.Close
    Set medixmasterconnPayment = Nothing
End Sub

Public Function ProcessCheckedItem(Item As ListItem, Optional Partial As Boolean)
    On Error Resume Next
    If Item.Checked = True Then
        If Trim(Item.SubItems(1)) = "Allocated" Then
            MsgBox "Record has been allocated!", vbCritical
            Item.Checked = False
        Else
            Item.SubItems(1) = "Allocated"
            Item.SubItems(6) = Item.SubItems(5)
            LblAmt = LblAmt + CDbl(Format(Item.SubItems(5), "#,##0.00"))
        End If
    Else
        LblAmt = LblAmt - CDbl(Format(Item.SubItems(5), "#,##0.00"))
        Item.SubItems(1) = ""
        Item.SubItems(6) = ""
    End If
End Function

Private Sub UpdateMedixGL(CASNo As String, CASVer As Integer, MedixPay As Double)
On Error Resume Next
Dim Update
Dim strsql As String
Dim PAY As New ADODB.Recordset
Dim strsqlPay As String
Dim CLM As New ADODB.Recordset
Dim strsqlCLM As String
Dim AddNewRec As Boolean
    AddNewRec = False
    strsqlPay = " Select PAYIndex, * from PaymentMedixGL " & _
                " Where PAYClaimNo = '" & Trim(CASNo) & "' AND PAYCLMVersion = '" & Trim(CASVer) & "'"
    PAY.Open strsqlPay, medixmasterconnPayment, adOpenKeyset, adLockOptimistic
        If PAY.recordCount = 0 Then
             PAY.AddNew
        End If
            With PAY
                !PAYClaimNo = Trim(CASNo)
                !PAYCLMVersion = Trim(CASVer)
                !PayPVNo = Trim(txtPVNo)
                !PayPVDate = Format(txtPVDate, "dd/mm/yyyy")
                If txtChequeDate <> "__/__/____" Then
                    !PAYChqDate = Format(txtChequeDate, "dd/mm/yyyy")
                Else
                    !PAYChqDate = ""
                End If
                !PAYChqAmt = Format(IIf(IsNumeric(Trim(txtChequeAmt)), Trim(txtChequeAmt), 0), "#,##0.00")
                !PAYChqNo = IIf(Len(Trim(txtChequeNo)), Trim(txtChequeNo), "")
                !PAYAllocNo = IIf(Len(Trim(txtAllocNo)), Trim(txtAllocNo), "")
                If EditFlag = False Then
                    !PAYAllocStatus = "Y"
                    !MedixPayAmt = MedixPay
                End If
                !PayRemarks = IIf(Len(Trim(TxtRemark)), Trim(TxtRemark), "")
                !PayAccPeriodMth = IIf(Len(Trim(CboACPeriodMth)), Trim(CboACPeriodMth), "")
                !PayAccPeriodYr = IIf(Len(Trim(CboACPeriodYr)), Trim(CboACPeriodYr), "")
                !PayLastUpdateDate = Date
                !PAYLastUpdateUser = Logon
            End With
            PAY.Update
        
        PAY.Close
        Set PAY = Nothing
    LblAmt = 0
End Sub

Private Sub LvDisplayData(Optional Item As ListItem, Optional ItemCheck As Boolean)
Dim lstRec As ListItem
Dim TmpAmtPaid As Double

    Call ClearFrameCase
    If lstMedixGainLoss.ListItems.Count Then
        If ItemCheck = True Then
            Set lstRec = Item
            lstMedixGainLoss.SelectedItem.Selected = False
            lstRec.Selected = True
        Else
            Set lstRec = lstMedixGainLoss.SelectedItem
        End If
            txtPVNo = IIf(Len(Trim(lstRec.SubItems(3))), lstRec.SubItems(3), "")
            If Len(lstRec.SubItems(4)) Then
                txtPVDate = Format(lstRec.SubItems(4), "DD/MM/YYYY")
            End If
            
            If Len(lstRec.SubItems(7)) Then
                txtAllocNo = lstRec.SubItems(7)
            End If
            If Len(lstRec.SubItems(10)) Then
                txtChequeNo = lstRec.SubItems(10)
            End If
            If IsDate(lstRec.SubItems(11)) Then
                txtChequeDate = Format(lstRec.SubItems(11), "DD/MM/YYYY")
            End If
            If IsNumeric(lstRec.SubItems(12)) Then
                txtChequeAmt = lstRec.SubItems(12)
            End If
            
            If Len(lstRec.SubItems(13)) Then
                CboACPeriodMth = Mid(lstRec.SubItems(13), 1, 2)
                CboACPeriodYr = Mid(lstRec.SubItems(13), 4, 4)
            End If
            TxtRemark = IIf(Len(lstRec.SubItems(14)), lstRec.SubItems(14), "")
            cmdSave.Enabled = False
    End If
End Sub
