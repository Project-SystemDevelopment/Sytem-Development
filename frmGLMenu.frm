VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "mscomctl.ocx"
Begin VB.Form frmGLMenu 
   Caption         =   "Generate Guarantee Letter"
   ClientHeight    =   4395
   ClientLeft      =   60
   ClientTop       =   510
   ClientWidth     =   10065
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   ScaleHeight     =   4395
   ScaleWidth      =   10065
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame Frame1 
      Height          =   2775
      Left            =   80
      TabIndex        =   12
      Top             =   240
      Width           =   9930
      Begin VB.ComboBox cboRequestBy 
         Height          =   315
         Left            =   1440
         TabIndex        =   0
         Top             =   225
         Width           =   3855
      End
      Begin VB.TextBox txtRemarks 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2355
         Left            =   6240
         MaxLength       =   500
         ScrollBars      =   2  'Vertical
         TabIndex        =   11
         Top             =   240
         Width           =   3615
      End
      Begin VB.ComboBox cboRequestMode 
         Height          =   315
         Left            =   1440
         TabIndex        =   1
         Top             =   500
         Width           =   3855
      End
      Begin VB.ComboBox cbPurpose 
         Height          =   315
         Left            =   1440
         TabIndex        =   2
         Top             =   780
         Width           =   3855
      End
      Begin VB.ComboBox cboDelayreason 
         Height          =   315
         Left            =   1440
         TabIndex        =   3
         Top             =   1070
         Width           =   3855
      End
      Begin MSMask.MaskEdBox txtRequestDate 
         Height          =   300
         Left            =   1440
         TabIndex        =   4
         Top             =   1350
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   529
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
      Begin MSMask.MaskEdBox txtRequestTime 
         Height          =   300
         Left            =   3960
         TabIndex        =   5
         Top             =   1350
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   529
         _Version        =   393216
         Enabled         =   0   'False
         MaxLength       =   5
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Mask            =   "##:##"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox txtGLDate 
         Height          =   300
         Left            =   1440
         TabIndex        =   6
         Top             =   1605
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   529
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
      Begin MSMask.MaskEdBox txtLGTime 
         Height          =   285
         Left            =   3960
         TabIndex        =   7
         Top             =   1605
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   503
         _Version        =   393216
         Enabled         =   0   'False
         MaxLength       =   5
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Mask            =   "##:##"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox txtFaxDate 
         Height          =   300
         Left            =   4680
         TabIndex        =   8
         Top             =   2715
         Visible         =   0   'False
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   529
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
      Begin MSMask.MaskEdBox txtFaxTime 
         Height          =   285
         Left            =   7200
         TabIndex        =   9
         Top             =   2700
         Visible         =   0   'False
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   503
         _Version        =   393216
         Enabled         =   0   'False
         MaxLength       =   5
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Mask            =   "##:##"
         PromptChar      =   "_"
      End
      Begin VB.TextBox txtCASLGAmt 
         Alignment       =   1  'Right Justify
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
         Left            =   1440
         TabIndex        =   10
         Top             =   1860
         Width           =   1335
      End
      Begin VB.ComboBox cboFollowup 
         Height          =   315
         ItemData        =   "frmGLMenu.frx":0000
         Left            =   3960
         List            =   "frmGLMenu.frx":000A
         TabIndex        =   32
         Text            =   "Y-Yes"
         Top             =   1860
         Width           =   1335
      End
      Begin VB.TextBox txtEmailGL 
         Height          =   285
         Left            =   1440
         TabIndex        =   33
         Top             =   2120
         Width           =   3855
      End
      Begin VB.TextBox txtHPNumber 
         Height          =   285
         Left            =   1440
         TabIndex        =   36
         Top             =   2350
         Width           =   3855
      End
      Begin VB.Label Label65 
         Caption         =   "Mobile Number"
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
         Index           =   11
         Left            =   120
         TabIndex        =   37
         Top             =   2400
         Width           =   1335
      End
      Begin VB.Label Label65 
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
         Height          =   255
         Index           =   10
         Left            =   120
         TabIndex        =   34
         Top             =   2160
         Width           =   735
      End
      Begin VB.Label Label65 
         Caption         =   "Follow-up"
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
         Index           =   9
         Left            =   2880
         TabIndex        =   31
         Top             =   1920
         Width           =   735
      End
      Begin VB.Label Label65 
         Caption         =   "Remarks"
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
         Index           =   8
         Left            =   5520
         TabIndex        =   25
         Top             =   240
         Width           =   735
      End
      Begin VB.Label Label65 
         Caption         =   "Reason for delay"
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
         Index           =   7
         Left            =   120
         TabIndex        =   24
         Top             =   1120
         Width           =   1335
      End
      Begin VB.Label Label65 
         Caption         =   "Faxed Date"
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
         Index           =   6
         Left            =   3360
         TabIndex        =   23
         Top             =   2745
         Visible         =   0   'False
         Width           =   1335
      End
      Begin VB.Label Label65 
         Caption         =   "Faxed Time"
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
         Index           =   5
         Left            =   6120
         TabIndex        =   22
         Top             =   2745
         Visible         =   0   'False
         Width           =   975
      End
      Begin VB.Label Label65 
         Caption         =   "GL Date"
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
         Index           =   4
         Left            =   120
         TabIndex        =   21
         Top             =   1650
         Width           =   1335
      End
      Begin VB.Label Label65 
         Caption         =   "Purpose"
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
         Index           =   3
         Left            =   120
         TabIndex        =   20
         Top             =   830
         Width           =   1095
      End
      Begin VB.Label Label65 
         Caption         =   "Request By"
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
         Index           =   2
         Left            =   120
         TabIndex        =   19
         Top             =   255
         Width           =   1095
      End
      Begin VB.Label Label65 
         Caption         =   "Request Mode"
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
         Index           =   1
         Left            =   120
         TabIndex        =   18
         Top             =   520
         Width           =   1095
      End
      Begin VB.Label Label65 
         Caption         =   "Request Time"
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
         Index           =   0
         Left            =   2880
         TabIndex        =   17
         Top             =   1395
         Width           =   1095
      End
      Begin VB.Label Label65 
         Caption         =   "Request Date"
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
         Index           =   30
         Left            =   120
         TabIndex        =   15
         Top             =   1395
         Width           =   1095
      End
      Begin VB.Label Label65 
         Caption         =   "LG Amt"
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
         Index           =   32
         Left            =   120
         TabIndex        =   14
         Top             =   1920
         Width           =   735
      End
      Begin VB.Label Label65 
         Caption         =   "LG Time"
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
         Index           =   68
         Left            =   2880
         TabIndex        =   13
         Top             =   1650
         Width           =   735
      End
   End
   Begin VB.Frame Frame2 
      Height          =   1335
      Left            =   80
      TabIndex        =   26
      Top             =   3000
      Width           =   9930
      Begin VB.CheckBox chkName 
         ForeColor       =   &H000000C0&
         Height          =   255
         Left            =   120
         TabIndex        =   40
         Top             =   240
         Width           =   5895
      End
      Begin VB.CheckBox chkIC 
         ForeColor       =   &H000000C0&
         Height          =   255
         Left            =   120
         TabIndex        =   39
         Top             =   600
         Width           =   3615
      End
      Begin VB.CheckBox chkExpDate 
         ForeColor       =   &H000000C0&
         Height          =   255
         Left            =   3840
         TabIndex        =   38
         Top             =   600
         Width           =   4575
      End
      Begin VB.CheckBox chkGL 
         Caption         =   "Please ensure the case is coverable. GL will be sent to the patient."
         ForeColor       =   &H000000C0&
         Height          =   255
         Left            =   120
         TabIndex        =   35
         Top             =   960
         Width           =   5175
      End
      Begin VB.CommandButton cmdSave 
         Caption         =   "&Save"
         Enabled         =   0   'False
         Height          =   375
         Left            =   8760
         TabIndex        =   28
         Top             =   220
         Width           =   975
      End
      Begin VB.CommandButton cmdExit 
         Caption         =   "&Exit"
         Height          =   375
         Left            =   8760
         TabIndex        =   27
         Top             =   600
         Width           =   975
      End
      Begin VB.Label lblCASNumber 
         Caption         =   "Label1"
         Height          =   255
         Left            =   6240
         TabIndex        =   30
         Top             =   960
         Width           =   2535
      End
   End
   Begin MSComctlLib.ListView lstLGListing 
      Height          =   480
      Left            =   120
      TabIndex        =   29
      Top             =   3600
      Visible         =   0   'False
      Width           =   9855
      _ExtentX        =   17383
      _ExtentY        =   847
      View            =   3
      LabelWrap       =   -1  'True
      HideSelection   =   0   'False
      AllowReorder    =   -1  'True
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   11
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "LG No"
         Object.Width           =   1940
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   1
         Text            =   "LG Date"
         Object.Width           =   1940
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   2
         Text            =   "LG Amt"
         Object.Width           =   2646
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   3
         Text            =   "Cov ID"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   4
         Text            =   "Patient"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   5
         Text            =   "DOA"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   6
         Text            =   "Hospital"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   7
         Text            =   "Issued By"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(9) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   8
         Text            =   "Remark"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(10) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   9
         Text            =   "LG User Time"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(11) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   10
         Text            =   "LG System Time"
         Object.Width           =   2540
      EndProperty
   End
   Begin VB.Label Label28 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00404040&
      Caption         =   "Generate Guarantee Letter"
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
      TabIndex        =   16
      Top             =   0
      Width           =   10125
   End
End
Attribute VB_Name = "frmGLMenu"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public GLForm As Boolean
Public GLSource As String
Public tmpGLPurpose As String
Public GLFollowUp As Boolean
Public tmpemailDelay As Boolean


Private Sub cboDelayreason_KeyPress(KeyAscii As Integer)
Dim NewText As String
Dim ValidCount As Integer
Dim ValidValue As String
Dim i As Integer

    If KeyAscii >= 32 And KeyAscii <> 127 Then
    NewText = LCase(Left(cboDelayreason.Text, cboDelayreason.SelStart) + Chr(KeyAscii) + Mid(cboDelayreason.Text, cboDelayreason.SelStart + cboDelayreason.SelLength + 1))
 
    ValidCount = 0
    ValidValue = ""
 
    For i = 0 To cboDelayreason.ListCount - 1
 
        If NewText = LCase(Left(cboDelayreason.List(i), Len(NewText))) Then
                ValidCount = ValidCount + 1
                ValidValue = cboDelayreason.List(i)
        End If
    Next
        
        If ValidCount <= 1 Then KeyAscii = 0
            If ValidCount = 1 Then
                cboDelayreason.Text = ValidValue
                cboDelayreason.SelStart = 0
                cboDelayreason.SelLength = Len(ValidValue)
            End If
        End If

End Sub

Private Sub cboRequestBy_KeyPress(KeyAscii As Integer)
Dim NewText As String
Dim ValidCount As Integer
Dim ValidValue As String
Dim i As Integer

    If KeyAscii >= 32 And KeyAscii <> 127 Then
    NewText = LCase(Left(cboRequestBy.Text, cboRequestBy.SelStart) + Chr(KeyAscii) + Mid(cboRequestBy.Text, cboRequestBy.SelStart + cboRequestBy.SelLength + 1))
 
    ValidCount = 0
    ValidValue = ""
 
    For i = 0 To cboRequestBy.ListCount - 1
 
        If NewText = LCase(Left(cboRequestBy.List(i), Len(NewText))) Then
                ValidCount = ValidCount + 1
                ValidValue = cboRequestBy.List(i)
        End If
    Next
        
        If ValidCount <= 1 Then KeyAscii = 0
            If ValidCount = 1 Then
                cboRequestBy.Text = ValidValue
                cboRequestBy.SelStart = 0
                cboRequestBy.SelLength = Len(ValidValue)
            End If
        End If
End Sub

Private Sub cboRequestMode_KeyPress(KeyAscii As Integer)
Dim NewText As String
Dim ValidCount As Integer
Dim ValidValue As String
Dim i As Integer

    If KeyAscii >= 32 And KeyAscii <> 127 Then
    NewText = LCase(Left(cboRequestMode.Text, cboRequestMode.SelStart) + Chr(KeyAscii) + Mid(cboRequestMode.Text, cboRequestMode.SelStart + cboRequestMode.SelLength + 1))
 
    ValidCount = 0
    ValidValue = ""
 
    For i = 0 To cboRequestMode.ListCount - 1
 
        If NewText = LCase(Left(cboRequestMode.List(i), Len(NewText))) Then
                ValidCount = ValidCount + 1
                ValidValue = cboRequestMode.List(i)
        End If
    Next
        
        If ValidCount <= 1 Then KeyAscii = 0
            If ValidCount = 1 Then
                cboRequestMode.Text = ValidValue
                cboRequestMode.SelStart = 0
                cboRequestMode.SelLength = Len(ValidValue)
            End If
        End If
End Sub

Private Sub cbPurpose_Change()
    If cbPurpose = "Admission" Then
        txtCASLGAmt = "2500"
    Else
        txtCASLGAmt = ""
    End If
    'cmdSave.Enabled = True
    If GLSource = "1" Then
        If frmCaseAdjustment2016.CASGLPurposeRcd = True Then
            If frmCaseAdjustment2016.CASGLPurpose = "Outpatient" Then
               ' If cbPurpose = "Admission" Or cbPurpose = "Discharge" Or cbPurpose = "Interim" Or cbPurpose = "Executive Medical Checkup" Or cbPurpose = "Post-Treatment" Then
               '     MsgBox "New GL is not allowed. Please Register A new case.", vbOKOnly + vbCritical, DialogBoxTitle
               '     cmdSave.Enabled = False
               ' End If
            ElseIf frmCaseAdjustment2016.CASGLPurpose = "Executive Medical Checkup" Then
                MsgBox "New GL is not allowed. Please Register A new case.", vbOKOnly + vbCritical, DialogBoxTitle
            ElseIf frmCaseAdjustment2016.CASGLPurpose = "Admission" Then
                If cbPurpose = "Outpatient" Or cbPurpose = "Executive Medical Checkup" Then
                    MsgBox "New GL is not allowed. Please Register A new case.", vbOKOnly + vbCritical, DialogBoxTitle
                    cmdSave.Enabled = False
                    chkGL.Value = 0
                End If
            ElseIf frmCaseAdjustment2016.CASGLPurpose = "Discharge" Then
                If cbPurpose = "Outpatient" Or cbPurpose = "Executive Medical Checkup" Or cbPurpose = "Admission" Then
                    MsgBox "New GL is not allowed. Please Register A new case.", vbOKOnly + vbCritical, DialogBoxTitle
                    cmdSave.Enabled = False
                    chkGL.Value = 0
                End If
    
            End If
        End If
    ElseIf GLSource = "2" Then
        If frmWSAdjGST.CASGLPurposeRcd = True Then
            If frmWSAdjGST.CASGLPurpose = "Outpatient" Then
                'If cbPurpose = "Admission" Or cbPurpose = "Discharge" Or cbPurpose = "Interim" Or cbPurpose = "Executive Medical Checkup" Or cbPurpose = "Post-Treatment" Then
                '    MsgBox "New GL is not allowed. Please Register A new case.", vbOKOnly + vbCritical, DialogBoxTitle
                '    cmdSave.Enabled = False
                'End If
            ElseIf frmWSAdjGST.CASGLPurpose = "Executive Medical Checkup" Then
                MsgBox "New GL is not allowed. Please Register A new case.", vbOKOnly + vbCritical, DialogBoxTitle
                cmdSave.Enabled = False
                chkGL.Value = 0
            ElseIf frmWSAdjGST.CASGLPurpose = "Admission" Then
                If cbPurpose = "Outpatient" Or cbPurpose = "Executive Medical Checkup" Then
                    MsgBox "New GL is not allowed. Please Register A new case.", vbOKOnly + vbCritical, DialogBoxTitle
                    cmdSave.Enabled = False
                    chkGL.Value = 0
                End If
            ElseIf frmWSAdjGST.CASGLPurpose = "Discharge" Then
                If cbPurpose = "Outpatient" Or cbPurpose = "Executive Medical Checkup" Or cbPurpose = "Admission" Then
                    MsgBox "New GL is not allowed. Please Register A new case.", vbOKOnly + vbCritical, DialogBoxTitle
                    cmdSave.Enabled = False
                    chkGL.Value = 0
                End If
    
            End If
        End If
    End If

End Sub

Private Sub cbPurpose_KeyPress(KeyAscii As Integer)
Dim NewText As String
Dim ValidCount As Integer
Dim ValidValue As String
Dim i As Integer

    If KeyAscii >= 32 And KeyAscii <> 127 Then
    NewText = LCase(Left(cbPurpose.Text, cbPurpose.SelStart) + Chr(KeyAscii) + Mid(cbPurpose.Text, cbPurpose.SelStart + cbPurpose.SelLength + 1))
 
    ValidCount = 0
    ValidValue = ""
 
    For i = 0 To cbPurpose.ListCount - 1
 
        If NewText = LCase(Left(cbPurpose.List(i), Len(NewText))) Then
                ValidCount = ValidCount + 1
                ValidValue = cbPurpose.List(i)
        End If
    Next
        
        If ValidCount <= 1 Then KeyAscii = 0
            If ValidCount = 1 Then
                cbPurpose.Text = ValidValue
                cbPurpose.SelStart = 0
                cbPurpose.SelLength = Len(ValidValue)
            End If
        End If
End Sub

Private Sub cbPurpose_LostFocus()
    If cbPurpose = "Admission" Then
        txtCASLGAmt = "2500"
    Else
        txtCASLGAmt = ""
    End If
        
    'cmdSave.Enabled = True
    If GLSource = "1" Then
        If frmCaseAdjustment2016.CASGLPurposeRcd = True Then
            If frmCaseAdjustment2016.CASGLPurpose = "Outpatient" Then
               ' If cbPurpose = "Admission" Or cbPurpose = "Discharge" Or cbPurpose = "Interim" Or cbPurpose = "Executive Medical Checkup" Or cbPurpose = "Post-Treatment" Then
               '     MsgBox "New GL is not allowed. Please Register A new case.", vbOKOnly + vbCritical, DialogBoxTitle
               '     cmdSave.Enabled = False
               ' End If
            ElseIf frmCaseAdjustment2016.CASGLPurpose = "Executive Medical Checkup" Then
                MsgBox "New GL is not allowed. Please Register A new case.", vbOKOnly + vbCritical, DialogBoxTitle
            'ElseIf frmCaseAdjustment2016.CASGLPurpose = "Admission" Then
             '   If cbPurpose = "Outpatient" Or cbPurpose = "Executive Medical Checkup" Then
             '       MsgBox "New GL is not allowed. Please Register A new case.", vbOKOnly + vbCritical, DialogBoxTitle
             '       cmdSave.Enabled = False
             '   End If
            'ElseIf frmCaseAdjustment2016.CASGLPurpose = "Discharge" Then
            '    If cbPurpose = "Outpatient" Or cbPurpose = "Executive Medical Checkup" Or cbPurpose = "Admission" Then
            '        MsgBox "New GL is not allowed. Please Register A new case.", vbOKOnly + vbCritical, DialogBoxTitle
            '        cmdSave.Enabled = False
            '    End If
    
            End If
        End If
    ElseIf GLSource = "2" Then
        If frmWSAdjGST.CASGLPurposeRcd = True Then
            If frmWSAdjGST.CASGLPurpose = "Outpatient" Then
            '    If cbPurpose = "Admission" Or cbPurpose = "Discharge" Or cbPurpose = "Interim" Or cbPurpose = "Executive Medical Checkup" Or cbPurpose = "Post-Treatment" Then
            '        MsgBox "New GL is not allowed. Please Register A new case.", vbOKOnly + vbCritical, DialogBoxTitle
            '        cmdSave.Enabled = False
            '    End If
            ElseIf frmWSAdjGST.CASGLPurpose = "Executive Medical Checkup" Then
                MsgBox "New GL is not allowed. Please Register A new case.", vbOKOnly + vbCritical, DialogBoxTitle
                cmdSave.Enabled = False
                chkGL.Value = 0
            'ElseIf frmWorksheetAdj2013.CASGLPurpose = "Admission" Then
            '    If cbPurpose = "Outpatient" Or cbPurpose = "Executive Medical Checkup" Then
            '        MsgBox "New GL is not allowed. Please Register A new case.", vbOKOnly + vbCritical, DialogBoxTitle
            '        cmdSave.Enabled = False
            '    End If
            'ElseIf frmWorksheetAdj2013.CASGLPurpose = "Discharge" Then
             '   If cbPurpose = "Outpatient" Or cbPurpose = "Executive Medical Checkup" Or cbPurpose = "Admission" Then
             '       MsgBox "New GL is not allowed. Please Register A new case.", vbOKOnly + vbCritical, DialogBoxTitle
             '       cmdSave.Enabled = False
             '   End If
    
            End If
        End If
    End If
End Sub

Private Sub chkExpDate_Click()
    If chkGL.Value = 1 And chkName.Value = 1 And chkIC.Value = 1 And chkExpDate.Value = 1 Then
        cmdSave.Enabled = True
    Else
        cmdSave.Enabled = False
    End If
End Sub

Private Sub chkGL_Click()
    If chkGL.Value = 1 And chkName.Value = 1 And chkIC.Value = 1 And chkExpDate.Value = 1 Then
        cmdSave.Enabled = True
    Else
        cmdSave.Enabled = False
    End If
End Sub

Private Sub chkIC_Click()
    If chkGL.Value = 1 And chkName.Value = 1 And chkIC.Value = 1 And chkExpDate.Value = 1 Then
        cmdSave.Enabled = True
    Else
        cmdSave.Enabled = False
    End If
End Sub

Private Sub chkName_Click()
    If chkGL.Value = 1 And chkName.Value = 1 And chkIC.Value = 1 And chkExpDate.Value = 1 Then
        cmdSave.Enabled = True
    Else
        cmdSave.Enabled = False
    End If
End Sub

Private Sub CmdExit_Click()
    Unload Me
    GLForm = False
End Sub

Private Sub CmdSave_Click()
Dim GLTimetaken As String
Dim tmpGLTimehh As String
Dim tmpReqTimehh As String
Dim tmpGLTimemm As String
Dim tmpReqTimemm As String
        
        'tmpGLTime = Format(txtLGTime, "hh:mm")
        'tmpReqTime = Format(txtRequestTime, "hh:mm")
        
        tmpemailDelay = False
        'tmpGLTimehh = Mid(txtLGTime, 1, 2) - Mid(txtRequestTime, 1, 2)
        'tmpGLTimemm = Mid(txtLGTime, 4, 2) - Mid(txtRequestTime, 4, 2)
        If txtRequestTime <> "__:__" Then
            tmpGLTimemm = DateTime.DateDiff("n", txtRequestTime, txtLGTime)
        End If
        If txtRequestDate = "__/__/____" Then
            MsgBox "Please Enter Request Date", vbOKOnly + vbCritical, DialogBoxTitle
            txtRequestDate.SetFocus
        ElseIf txtRequestTime = "__:__" Then
            MsgBox "Please Enter Request Time", vbOKOnly + vbCritical, DialogBoxTitle
            txtRequestTime.SetFocus
        ElseIf cboRequestMode = "" Then
            MsgBox "Please Enter Request Mode", vbOKOnly + vbCritical, DialogBoxTitle
            cboRequestMode.SetFocus
        ElseIf cboRequestBy = "" Then
            MsgBox "Please Enter Request By", vbOKOnly + vbCritical, DialogBoxTitle
            cboRequestBy.SetFocus
        ElseIf txtGLDate = "__/__/____" Then
            MsgBox "Please Enter GL Date", vbOKOnly + vbCritical, DialogBoxTitle
            txtGLDate.SetFocus
        ElseIf txtLGTime = "__/__/____" Then
            MsgBox "Please Enter GL Time", vbOKOnly + vbCritical, DialogBoxTitle
            txtLGTime.SetFocus
        ElseIf Trim(txtCASLGAmt) = "" And cbPurpose <> "Outpatient" Then
            MsgBox "Please Enter Reserve Amount", vbOKOnly + vbCritical, DialogBoxTitle
            txtCASLGAmt.SetFocus
        ElseIf cbPurpose = "" Then
            MsgBox "Please Enter purpose for GL", vbOKOnly + vbCritical, DialogBoxTitle
            txtCASLGAmt.SetFocus
        ElseIf Format(txtGLDate, "dd/mm/yyyy") > Format(txtRequestDate, "dd/mm/yyyy") And cboDelayreason = "" Then
            MsgBox "Please Enter reason for delay", vbOKOnly + vbCritical, DialogBoxTitle
            cboDelayreason.SetFocus
        ElseIf (Format(txtGLDate, "dd/mm/yyyy") = Format(txtRequestDate, "dd/mm/yyyy")) And tmpGLTimemm > 30 And cboDelayreason = "" And GLSource = "1" Then
            MsgBox "Please Enter reason for delay", vbOKOnly + vbCritical, DialogBoxTitle
            tmpemailDelay = True
            cboDelayreason.SetFocus
        ElseIf (Format(txtGLDate, "dd/mm/yyyy") = Format(txtRequestDate, "dd/mm/yyyy")) And tmpGLTimemm > 45 And cboDelayreason = "" And GLSource = "2" Then
            tmpemailDelay = True
            MsgBox "Please Enter reason for delay", vbOKOnly + vbCritical, DialogBoxTitle
            cboDelayreason.SetFocus
'        ElseIf (Format(txtGLDate, "dd/mm/yyyy") = Format(txtRequestDate, "dd/mm/yyyy")) And DateTime.DateDiff("n", txtRequestTime, txtLGTime) > 30 And cboDelayreason = "" Then
'            MsgBox "Please Enter reason for delay", vbOKOnly + vbCritical, DialogBoxTitle
'            cboDelayreason.SetFocus
        ElseIf txtRemarks = "" Then
            MsgBox "Please Enter remark", vbOKOnly + vbCritical, DialogBoxTitle
            txtRemarks.SetFocus
        ElseIf cboFollowup = "" Then
            MsgBox "Please select follow-up status", vbOKOnly + vbCritical, DialogBoxTitle
            cboFollowup.SetFocus
        ElseIf (Mid(lblCASNumber, 1, 2) = "ET" Or Mid(lblCASNumber, 1, 2) = "EQ" Or Mid(lblCASNumber, 1, 2) = "ES" Or Mid(lblCASNumber, 1, 2) = "EP") And txtHPNumber = "" Then
            'If txtHPNumber = "" Then
                MsgBox "Please enter Patient's Mobile Number", vbOKOnly + vbCritical, DialogBoxTitle
                txtHPNumber.SetFocus
            'End If
            
        Else
            
            If tmpGLTimemm > 45 And GLSource = "2" Then
                tmpemailDelay = True
            ElseIf tmpGLTimemm > 30 And GLSource = "1" Then
                tmpemailDelay = True
            Else
                tmpemailDelay = False
            End If
            If tmpemailDelay = True Then
            Dim ToAdd As String
            Dim Body As String
            
            If Mid(lblCASNumber, 1, 2) = "EM" Or Mid(lblCASNumber, 1, 2) = "ES" Or Mid(lblCASNumber, 1, 2) = "EQ" Or Mid(lblCASNumber, 1, 2) = "EC" Or Mid(lblCASNumber, 1, 2) = "EP" Or Mid(lblCASNumber, 1, 2) = "ET" Or Mid(lblCASNumber, 1, 2) = "WM" Or Mid(lblCASNumber, 1, 2) = "WK" Then
                ToAdd = "pravin@medix.com.my;izham@medix.com.my"
            ElseIf Mid(lblCASNumber, 1, 2) = "TK" Then
                ToAdd = "mjeffri@medix.com.my"
            Else
                ToAdd = "irene@medix.com.my"
            End If
            
            'ToAdd = "bryan@medix.com.my;trinath@medix.com.my;trinath18@gmail.com"
            
            Body = "Dear Sir/Madam," & vbCrLf & vbCrLf & _
                            "DELAY IN ISSUANCE OF GL" & vbCrLf & vbCrLf & _
                            "LG Number : " & lblCASNumber & vbCrLf
                           
            Call MailCDOMsg(ToAdd, Body)
            End If
            frmGLMenu.Visible = False
            
            If Mid(cboFollowup, 1, 1) = "Y" Then
                GLFollowUp = True
            Else
                GLFollowUp = False
            End If
            
            If GLSource = "1" Then
                frmGLMenu.GLForm = True
                frmCaseAdjustment2016.cmdLG_Click
                
            ElseIf GLSource = "2" Then
                frmGLMenu.GLForm = True
                frmWSAdjGST.cmdFinalLG_Click
            End If
            
        
        End If
End Sub

Private Sub Form_Load()
Dim strsqlLG As String
Dim CLMLG As New ADODB.Recordset

    Call ComboListing

End Sub

Private Sub ComboListing()
    cboRequestMode.AddItem "Medical Report"
    cboRequestMode.AddItem "Email"
    cboRequestMode.AddItem "Telephone"

    cboRequestBy.AddItem "Hospital"
    cboRequestBy.AddItem "Payor"
    cboRequestBy.AddItem "Agent"
    cboRequestBy.AddItem "Policyholder"
    
    If GLSource = "1" Then
        cbPurpose.AddItem "Admission"
        cbPurpose.AddItem "Interim"
        cbPurpose.AddItem "Outpatient"
        cbPurpose.AddItem "Executive Medical Checkup"
        
        'cbPurpose.AddItem "Discharge"
        'cbPurpose.AddItem "Follow-up"
    Else
        cbPurpose.AddItem "Admission"
        cbPurpose.AddItem "Interim"
        cbPurpose.AddItem "Outpatient"
        cbPurpose.AddItem "Executive Medical Checkup"
        cbPurpose.AddItem "Discharge"
        cbPurpose.AddItem "Follow-up"
    End If
    cboDelayreason.AddItem "GL Request upon Discharge"
    cboDelayreason.AddItem "Awaiting Renewal Data"
    cboDelayreason.AddItem "Amended GL"
    cboDelayreason.AddItem "Amended Bill"
    cboDelayreason.AddItem "GL Request After Admission"
    cboDelayreason.AddItem "FGL Request After Discharge"
    cboDelayreason.AddItem "Refer to Payor"
    cboDelayreason.AddItem "Pending Policy Status"
    cboDelayreason.AddItem "Query Attending Doctor"
    cboDelayreason.AddItem "Require Investigation Report"
    cboDelayreason.AddItem "Bill Above RM25K"
    cboDelayreason.AddItem "Bill Above RM50K"
    cboDelayreason.AddItem "Pre-Admission"
    cboDelayreason.AddItem "Pending Audit Review"

End Sub

Private Sub txtFaxDate_LostFocus()
   If IsDate(txtFaxDate) Then
    Else
        If txtFaxDate <> "__/__/____" Then
            MsgBox "Invalid Date Format! ", vbCritical
            txtFaxDate.SetFocus
        End If
    End If
End Sub

Private Sub txtFaxTime_LostFocus()
    If isTime(txtFaxTime) Then
    
    Else
        If txtFaxTime <> "__:__" Then
            MsgBox "Invalid Time Format! ", vbCritical
             txtFaxTime.SetFocus
        End If
    End If

End Sub

Private Sub txtGLDate_LostFocus()
    If IsDate(txtGLDate) Then
    Else
        If txtGLDate <> "__/__/____" Then
            MsgBox "Invalid Date Format! ", vbCritical
            txtGLDate.SetFocus
        End If
    End If

End Sub

Private Sub txtLGTime_LostFocus()
    If isTime(txtLGTime) Then
    
    Else
        If txtLGTime <> "__:__" Then
            MsgBox "Invalid Time Format! ", vbCritical
             txtLGTime.SetFocus
        End If
    End If

End Sub

Private Sub txtRequestDate_LostFocus()
    If IsDate(txtRequestDate) Then
    Else
        If txtRequestDate <> "__/__/____" Then
            MsgBox "Invalid Date Format! ", vbCritical
            txtRequestDate.SetFocus
        End If
    End If

End Sub

Private Sub txtRequestTime_LostFocus()
    If isTime(txtRequestTime) Then
    
    Else
        If txtRequestTime <> "__:__" Then
            MsgBox "Invalid Time Format! ", vbCritical
             txtRequestTime.SetFocus
        End If
    End If
End Sub

Public Sub MailCDOMsg(ToAdd As String, Body As String)
Const cdoSendUsingPickup = 1 'Send message using the local SMTP service pickup directory.
Const cdoSendUsingPort = 2 'Send the message using the network (SMTP over the network).
Const cdoAnonymous = 0 'Do not authenticate
Const cdoBasic = 1 'basic (clear-text) authentication
Const cdoNTLM = 2 'NTLM
Dim objMessage As Object
On Error GoTo debugs

Set objMessage = CreateObject("CDO.Message")
objMessage.Subject = "Notification of Delay"
'objMessage.From = "adm@medix.com.my"
objMessage.From = "notifications@medix.com.my"
objMessage.To = ToAdd
'objMessage.BCC = "bryan@medix.com.my; trinath@medix.com.my"
'objMessage.TextBody = "This is some sample message text.." & vbCrLf & "It was sent using SMTP authentication."
objMessage.TextBody = "Dear Sir/Madam," & vbCrLf & vbCrLf & _
                            "NOTIFICATION OF DELAY" & vbCrLf & vbCrLf & _
                            "Reason : " & cboDelayreason & vbCrLf & vbCrLf & _
                            "Request Date : " & txtRequestDate & vbCrLf & vbCrLf & _
                            "Request Time : " & txtRequestTime & vbCrLf & vbCrLf & _
                            "GL Date : " & txtGLDate & vbCrLf & vbCrLf & _
                            "GL Time : " & txtLGTime & vbCrLf & vbCrLf & _
                            "LG Number : " & lblCASNumber & vbCrLf & vbCrLf

objMessage.Configuration.Fields.Item("http://schemas.microsoft.com/cdo/configuration/sendusing") = 2
objMessage.Configuration.Fields.Item("http://schemas.microsoft.com/cdo/configuration/smtpserver") = "mail.medix.com.my"
objMessage.Configuration.Fields.Item("http://schemas.microsoft.com/cdo/configuration/smtpauthenticate") = cdoBasic
'objMessage.Configuration.Fields.Item("http://schemas.microsoft.com/cdo/configuration/sendusername") = "adm@medix.com.my"
objMessage.Configuration.Fields.Item("http://schemas.microsoft.com/cdo/configuration/sendusername") = "notifications@medix.com.my"
objMessage.Configuration.Fields.Item("http://schemas.microsoft.com/cdo/configuration/sendpassword") = "medix123"
objMessage.Configuration.Fields.Item("http://schemas.microsoft.com/cdo/configuration/smtpserverport") = 25
objMessage.Configuration.Fields.Item("http://schemas.microsoft.com/cdo/configuration/smtpusessl") = False
objMessage.Configuration.Fields.Item("http://schemas.microsoft.com/cdo/configuration/smtpconnectiontimeout") = 60
objMessage.Configuration.Fields.Update
objMessage.Send
debugs:
If Err.Description <> "" Then MsgBox Err.Description

End Sub

