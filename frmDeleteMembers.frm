VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmDeleteMembers 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Membership Deletion"
   ClientHeight    =   6810
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   11730
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5399.064
   ScaleMode       =   0  'User
   ScaleWidth      =   11910
   WindowState     =   2  'Maximized
   Begin VB.Frame Frame7 
      Caption         =   "Status"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   645
      Left            =   120
      TabIndex        =   15
      Top             =   6120
      Width           =   11580
      Begin VB.Label txtTotal 
         Alignment       =   2  'Center
         BackColor       =   &H00C0FFFF&
         Caption         =   "0"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Left            =   7800
         TabIndex        =   20
         Top             =   240
         Width           =   1200
      End
      Begin VB.Label Label17 
         Alignment       =   2  'Center
         Caption         =   "Principal"
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
         Left            =   480
         TabIndex        =   19
         Top             =   240
         Width           =   960
      End
      Begin VB.Label Label16 
         Alignment       =   2  'Center
         Caption         =   "Supplementary"
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
         Left            =   3480
         TabIndex        =   18
         Top             =   240
         Width           =   1320
      End
      Begin VB.Label txtTotalPrincipals 
         Alignment       =   2  'Center
         BackColor       =   &H00C0FFFF&
         Caption         =   "0"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Left            =   1440
         TabIndex        =   17
         Top             =   240
         Width           =   960
      End
      Begin VB.Label txtTotalSupplementary1 
         Alignment       =   2  'Center
         BackColor       =   &H00C0FFFF&
         Caption         =   "0"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Left            =   4920
         TabIndex        =   16
         Top             =   240
         Width           =   1200
      End
      Begin VB.Label Label4 
         Alignment       =   2  'Center
         Caption         =   "Total"
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
         Left            =   6960
         TabIndex        =   21
         Top             =   240
         Width           =   960
      End
   End
   Begin MSComctlLib.ListView lstMBMList 
      Height          =   3735
      Left            =   120
      TabIndex        =   13
      Top             =   2280
      Width           =   11535
      _ExtentX        =   20346
      _ExtentY        =   6588
      View            =   3
      Sorted          =   -1  'True
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      FullRowSelect   =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   5
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Bordx. Date"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Bordx. Batch"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "Mem. No"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Text            =   "Covered ID"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   4
         Text            =   "Mem. Name"
         Object.Width           =   2540
      EndProperty
   End
   Begin VB.Frame Frame8 
      Height          =   1695
      Left            =   120
      TabIndex        =   2
      Top             =   240
      Width           =   11535
      Begin VB.Frame Frame4 
         Height          =   855
         Left            =   120
         TabIndex        =   26
         Top             =   120
         Width           =   9615
         Begin VB.ComboBox cboPAYCode 
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
            Left            =   3960
            TabIndex        =   29
            Top             =   165
            Width           =   5595
         End
         Begin VB.ComboBox cboHLTCode 
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
            Left            =   1440
            Sorted          =   -1  'True
            TabIndex        =   27
            Text            =   "S - SIHAT MALAYSIA"
            Top             =   160
            Width           =   1905
         End
         Begin VB.ComboBox cboDeleteOption 
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
            Left            =   1440
            Sorted          =   -1  'True
            TabIndex        =   28
            Top             =   450
            Width           =   8115
         End
         Begin VB.Label Label18 
            Caption         =   "Payor"
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
            TabIndex        =   32
            Top             =   165
            Width           =   915
         End
         Begin VB.Label Label21 
            Caption         =   "Delete Option"
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
            TabIndex        =   31
            Top             =   480
            Width           =   1365
         End
         Begin VB.Label Label19 
            Caption         =   "Health Type"
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
            TabIndex        =   30
            Top             =   165
            Width           =   1365
         End
      End
      Begin VB.TextBox txtMBMType 
         Height          =   285
         Left            =   10320
         TabIndex        =   12
         Text            =   "MBMType"
         Top             =   1800
         Visible         =   0   'False
         Width           =   1095
      End
      Begin VB.TextBox txtMBMNumber 
         Height          =   285
         Left            =   10320
         TabIndex        =   11
         Text            =   "MBMNumber"
         Top             =   1560
         Visible         =   0   'False
         Width           =   1095
      End
      Begin VB.Frame BordxFrame 
         Height          =   600
         Left            =   120
         TabIndex        =   8
         Top             =   960
         Visible         =   0   'False
         Width           =   9615
         Begin MSMask.MaskEdBox txtBordxFrom 
            Height          =   285
            Left            =   1440
            TabIndex        =   0
            Top             =   180
            Width           =   1185
            _ExtentX        =   2090
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
         Begin MSMask.MaskEdBox txtBordxTo 
            Height          =   285
            Left            =   4080
            TabIndex        =   1
            Top             =   180
            Width           =   1185
            _ExtentX        =   2090
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
         Begin VB.Label Label22 
            Caption         =   "Bordx Date From"
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
            Left            =   90
            TabIndex        =   10
            Top             =   225
            Width           =   1275
         End
         Begin VB.Label Label23 
            Caption         =   "Bordx Date To"
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
            Left            =   2880
            TabIndex        =   9
            Top             =   180
            Width           =   1275
         End
      End
      Begin VB.Frame BatchFrame 
         Height          =   600
         Left            =   120
         TabIndex        =   3
         Top             =   960
         Visible         =   0   'False
         Width           =   9615
         Begin VB.TextBox txtBatchNo1 
            Height          =   285
            Left            =   4320
            TabIndex        =   4
            Top             =   180
            Width           =   1185
         End
         Begin MSMask.MaskEdBox txtReceiptDate 
            Height          =   285
            Left            =   1560
            TabIndex        =   5
            Top             =   165
            Width           =   1185
            _ExtentX        =   2090
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
         Begin VB.Label Label26 
            Caption         =   "Receipt Date From"
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
            TabIndex        =   7
            Top             =   210
            Width           =   1515
         End
         Begin VB.Label Label27 
            Caption         =   "Batch No"
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
            Left            =   2940
            TabIndex        =   6
            Top             =   180
            Width           =   1275
         End
      End
      Begin VB.Frame Frame2 
         Height          =   1335
         Left            =   10200
         TabIndex        =   22
         Top             =   120
         Width           =   1215
         Begin VB.CommandButton cmdGenerate 
            Caption         =   "&Delete"
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
            Left            =   120
            TabIndex        =   25
            Top             =   560
            Width           =   960
         End
         Begin VB.CommandButton cmdExit 
            Cancel          =   -1  'True
            Caption         =   "&Exit"
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
            Left            =   120
            TabIndex        =   24
            Top             =   920
            Width           =   960
         End
         Begin VB.CommandButton cmdList 
            Caption         =   "&Go"
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
            Left            =   120
            TabIndex        =   23
            Top             =   200
            Width           =   960
         End
      End
      Begin VB.Frame MBMFrame 
         Height          =   600
         Left            =   120
         TabIndex        =   33
         Top             =   960
         Visible         =   0   'False
         Width           =   9615
         Begin VB.TextBox txtMBMNumberFrom 
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
            TabIndex        =   35
            Top             =   180
            Width           =   1740
         End
         Begin VB.TextBox txtMBMNumberTo 
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
            Left            =   4800
            TabIndex        =   34
            Top             =   180
            Width           =   1740
         End
         Begin VB.Label Label2 
            Caption         =   "MBM'Ship No From"
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
            Left            =   90
            TabIndex        =   37
            Top             =   180
            Width           =   1395
         End
         Begin VB.Label Label5 
            Caption         =   "MBM""Ship No To"
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
            Left            =   3450
            TabIndex        =   36
            Top             =   180
            Width           =   1275
         End
      End
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00404040&
      Caption         =   "Membership Deletion"
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
      Height          =   270
      Left            =   0
      TabIndex        =   38
      Top             =   0
      Width           =   11775
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "Mem Listing"
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
      Left            =   120
      TabIndex        =   14
      Top             =   2040
      Width           =   1320
   End
End
Attribute VB_Name = "frmDeleteMembers"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim Exportsql As String
Dim ExportsqlList As String
Dim TotSupplementary As Integer
Dim TotPrincipal As Integer
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

Private Sub cboDeleteOption_Click()
    If Mid(cboDeleteOption, 1, 1) = "D" Then
        MBMFrame.Visible = False
        BordxFrame.Visible = True
        BatchFrame.Visible = False
        txtBordxFrom.SetFocus
    ElseIf Mid(cboDeleteOption, 1, 1) = "B" Then
        MBMFrame.Visible = False
        BordxFrame.Visible = False
        BatchFrame.Visible = True
        txtReceiptDate.SetFocus
    ElseIf Mid(cboDeleteOption, 1, 1) = "M" Then
        MBMFrame.Visible = True
        BordxFrame.Visible = False
        BatchFrame.Visible = False
        'txtMBMNumberFrom.SetFocus
    End If
End Sub

Private Sub cboDeleteOption_Change()
    If InStr(cboDeleteOption.Text, "null") Then
       cboDeleteOption.Enabled = False
    End If
    
    If Mid(cboDeleteOption, 1, 1) = "D" Then
        BordxFrame.Visible = True
        MBMFrame.Visible = False
        BatchFrame.Visible = False
        txtBordxFrom.SetFocus
    ElseIf Mid(cboDeleteOption, 1, 1) = "B" Then
        BordxFrame.Visible = False
        MBMFrame.Visible = False
        BatchFrame.Visible = True
        txtReceiptDate.SetFocus
    ElseIf Mid(cboDeleteOption, 1, 1) = "M" Then
        MBMFrame.Visible = True
        BordxFrame.Visible = False
        BatchFrame.Visible = False
    End If

End Sub

Private Sub cboDeleteOption_KeyPress(KeyAscii As Integer)
Dim NewText As String
Dim ValidCount As Integer
Dim ValidValue As String
Dim i As Integer
    
    If KeyAscii >= 32 And KeyAscii <> 127 Then
    NewText = LCase(Left(cboDeleteOption.Text, cboDeleteOption.SelStart) + Chr(KeyAscii) + Mid(cboDeleteOption.Text, cboDeleteOption.SelStart + cboDeleteOption.SelLength + 1))

    ValidCount = 0
    ValidValue = ""

    For i = 0 To cboDeleteOption.ListCount - 1
 
        If NewText = LCase(Left(cboDeleteOption.List(i), Len(NewText))) Then
                ValidCount = ValidCount + 1
                ValidValue = cboDeleteOption.List(i)
        End If
        Next
        If ValidCount <= 1 Then KeyAscii = 0
             If ValidCount = 1 Then
               cboDeleteOption.Text = ValidValue
               cboDeleteOption.SelStart = 0
               cboDeleteOption.SelLength = Len(ValidValue)
             End If
        End If
End Sub

Private Sub cboHLTCode_Change()
    If InStr(cboHLTCode.Text, "null") Then
       cboHLTCode.Enabled = False
    End If
End Sub

Private Sub cboHLTCode_KeyPress(KeyAscii As Integer)
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

Private Sub cboPAYCode_Change()
    If InStr(cboPAYCode.Text, "null") Then
       cboPAYCode.Enabled = False
    End If
End Sub

Private Sub cboPAYCode_KeyPress(KeyAscii As Integer)
Dim NewText As String
Dim ValidCount As Integer
Dim ValidValue As String
Dim i As Integer
    
    If KeyAscii >= 32 And KeyAscii <> 127 Then
    NewText = LCase(Left(cboPAYCode.Text, cboPAYCode.SelStart) + Chr(KeyAscii) + Mid(cboPAYCode.Text, cboPAYCode.SelStart + cboPAYCode.SelLength + 1))

    ValidCount = 0
    ValidValue = ""

    For i = 0 To cboPAYCode.ListCount - 1
 
        If NewText = LCase(Left(cboPAYCode.List(i), Len(NewText))) Then
                ValidCount = ValidCount + 1
                ValidValue = cboPAYCode.List(i)
        End If
        Next
        If ValidCount <= 1 Then KeyAscii = 0
             If ValidCount = 1 Then
               cboPAYCode.Text = ValidValue
               cboPAYCode.SelStart = 0
               cboPAYCode.SelLength = Len(ValidValue)
             End If
        End If
End Sub


Private Sub cmdExit_Click()
'    Dim RecordExit
'    RecordExit = MsgBox("Quit this form?", vbYesNo + vbExclamation)
'    If RecordExit = vbYes Then
'        bExit = True
        Unload Me
'    End If
End Sub

Private Sub cmdGenerate_Click()
Dim RecordSaved
     TotPrincipal = 0
     TotSupplementary = 0
     txtTotal = 0
     txtTotalPrincipals = 0
     txtTotalSupplementary1 = 0
     txtMBMNumber = ""
     txtMBMType = ""
     
    If cboHLTCode = "" Then
        MsgBox "Please Enter Health Plan", vbOKOnly + vbCritical, DialogBoxTitle
    ElseIf cboPAYCode = "" Then
        MsgBox "Please Enter Payor Code", vbOKOnly + vbCritical, DialogBoxTitle
    ElseIf cboDeleteOption = "" Then
        MsgBox "Please Enter Export Option", vbOKOnly + vbCritical, DialogBoxTitle
    ElseIf (Mid(cboDeleteOption, 1, 1) = "B") And (txtReceiptDate = "__/__/____" Or txtBatchNo1 = "") Then
        MsgBox "Please Enter Receipt Date or Batch No", vbOKOnly + vbCritical, DialogBoxTitle
    ElseIf (Mid(cboDeleteOption, 1, 1) = "D") And (txtBordxFrom = "__/__/____" Or txtBordxTo = "__/__/____") Then
        MsgBox "Please Enter Bordx Date", vbOKOnly + vbCritical, DialogBoxTitle
    Else
      RecordSaved = MsgBox("Are you sure?", vbYesNo + vbExclamation)
        If RecordSaved = vbYes Then
    
            Screen.MousePointer = vbHourglass

    If Mid(cboDeleteOption, 1, 1) = "B" Then
        Exportsql = "HLTCode = '" & Mid(cboHLTCode, 1, InStr(1, cboHLTCode, "-") - 1) & "' And PAYCode = '" & Mid(cboPAYCode, 1, InStr(1, cboPAYCode, "-") - 1) & "' And MBMBordxDate = CONVERT(SMALLDATETIME, '" & txtReceiptDate & "', 103) And mbmbatchno ='" & txtBatchNo1 & "'"
    ElseIf Mid(cboDeleteOption, 1, 1) = "D" Then
        Exportsql = "HLTCode = '" & Mid(cboHLTCode, 1, InStr(1, cboHLTCode, "-") - 1) & "' And PAYCode = '" & Mid(cboPAYCode, 1, InStr(1, cboPAYCode, "-") - 1) & "' And MBMBordxDate >= CONVERT(SMALLDATETIME, '" & txtBordxFrom & "', 103) And MBMBordxDate <= CONVERT(SMALLDATETIME, '" & txtBordxTo & "', 103)"
    ElseIf Mid(cboDeleteOption, 1, 1) = "M" Then
        Exportsql = "HLTCode = '" & Mid(cboHLTCode, 1, InStr(1, cboHLTCode, "-") - 1) & "' And PAYCode = '" & Mid(cboPAYCode, 1, InStr(1, cboPAYCode, "-") - 1) & "' And MBMNumber >= '" & txtMBMNumberFrom & "' And MBMNumber <= '" & txtMBMNumberTo & "'"
    End If
        Call DeletePrincipals
   End If
   End If
    'txtTotalPrincipals = ""
End Sub

Private Sub ComboListing()
    Dim HLT As New ADODB.Recordset
    Dim PAY As New ADODB.Recordset
    
    HLT.Open "HLT", medixmasterconn, adOpenKeyset, adLockOptimistic
    PAY.Open "PAY", medixmasterconn, adOpenKeyset, adLockOptimistic
    
    Do Until HLT.EOF
       cboHLTCode.AddItem HLT!HLTCode & " - " & HLT!HLTDescription
       HLT.MoveNext
    Loop
    
    Do Until PAY.EOF
        cboPAYCode.AddItem PAY!PAYCode & " - " & PAY!PAYCompanyName
        PAY.MoveNext
    Loop
    
    cboDeleteOption.AddItem "D" & " - " & "Bordx Date"
    cboDeleteOption.AddItem "B" & " - " & "Batch"
    cboDeleteOption.AddItem "M" & " - " & "Membership"
    
    HLT.Close
    PAY.Close

End Sub


Private Sub cmdList_Click()
Dim RecordSaved
Dim DELsqlList As String
Dim MBM As New ADODB.Recordset
Dim strsqlMBM As String
    TotSupplementary = 0
    TotPrincipal = 0
    txtTotalPrincipals = 0
    txtTotalSupplementary1 = 0
    txtTotal = 0
    txtMBMNumber = ""
    txtMBMType = ""
    If cboHLTCode = "" Then
        MsgBox "Please Enter Health Plan", vbOKOnly + vbCritical, DialogBoxTitle
    ElseIf cboPAYCode = "" Then
        MsgBox "Please Enter Payor Code", vbOKOnly + vbCritical, DialogBoxTitle
    ElseIf cboDeleteOption = "" Then
        MsgBox "Please Enter Export Option", vbOKOnly + vbCritical, DialogBoxTitle
    ElseIf (Mid(cboDeleteOption, 1, 1) = "B") And (txtReceiptDate = "__/__/____" Or txtBatchNo1 = "") Then
        MsgBox "Please Enter Receipt Date or Batch No", vbOKOnly + vbCritical, DialogBoxTitle
    ElseIf (Mid(cboDeleteOption, 1, 1) = "D") And (txtBordxFrom = "__/__/____" Or txtBordxTo = "__/__/____") Then
        MsgBox "Please Enter Bordx Date", vbOKOnly + vbCritical, DialogBoxTitle
    ElseIf (Mid(cboDeleteOption, 1, 1) = "M") And (txtMBMNumberFrom = "" Or txtMBMNumberTo = "") Then
        MsgBox "Please Enter Membership No", vbOKOnly + vbCritical, DialogBoxTitle
    Else
      'RecordSaved = MsgBox("Are you sure?", vbYesNo + vbExclamation)
      '  If RecordSaved = vbYes Then
    
            Screen.MousePointer = vbHourglass
        Call DELMBMListing
   End If
   'End If
    'txtTotalPrincipals = ""

End Sub

Private Sub Form_Load()
    Call ComboListing
    Me.Width = 7000
    Me.Height = 3700
    bExit = False
End Sub


Private Sub DeletePrincipals()
Dim strsqlMBM As String
Dim MBM As New ADODB.Recordset
Dim MBMArchieve As New ADODB.Recordset
Dim strdelPrincipal As String
Dim DelPrincipal As New ADODB.Recordset
Dim DelOthers As String
Dim MBMOthers As New ADODB.Recordset

    
    'strsqlMBM = "SELECT HLTCode, PAYCode,MBMMemberType, MBMNumber, "
    'strsqlMBM = strsqlMBM & "MBMPolicyNo, MBMName, MBMIcBcPp, INSCode, "
    'strsqlMBM = strsqlMBM & "MBMBatchNo, MBMBordxDate, MBMStatus, "
    'strsqlMBM = strsqlMBM & "MBMSalutation, MBMAddress1, MBMAddress2, MBMAddress3, "
    'strsqlMBM = strsqlMBM & "CTYCode, MBMPostCode, MBMTelnoH, MBMTelNoM, "
    'strsqlMBM = strsqlMBM & "From MBM Where " & Exportsql
    
    strsqlMBM = "SELECT * From MBM where " & Exportsql
    MBM.Open strsqlMBM, medixmasterconn, adOpenKeyset, adLockOptimistic
    MBMArchieve.Open "MBMArchive", medixmasterconnA, adOpenKeyset, adLockOptimistic
    'MBMCoveredPersonsArchieve.Open "MBMCoveredPersonsArchieve", medixmasterconnA, adOpenKeyset, adLockOptimistic
    If MBM.recordCount = 0 Then
        MsgBox "No Record Found", vbOKOnly + vbInformation, DialogBoxTitle
        Screen.MousePointer = vbDefault
    Else
    Do Until MBM.EOF
        
        MBMArchieve.AddNew
             MBMArchieve!MBMNumber = MBM!MBMNumber
             MBMArchieve!MBMPolicyNo = MBM!MBMPolicyNo
             MBMArchieve!MBMIcBcPp = MBM!MBMIcBcPp
             MBMArchieve!MBMName = MBM!MBMName
             MBMArchieve!MBMSalutation = MBM!MBMSalutation
             MBMArchieve!HLTCode = MBM!HLTCode
             MBMArchieve!PAYCode = MBM!PAYCode
             MBMArchieve!MBMAddress1 = MBM!MBMAddress1
             MBMArchieve!MBMAddress2 = MBM!MBMAddress2
             MBMArchieve!MBMAddress3 = MBM!MBMAddress3
             MBMArchieve!CTYCode = MBM!CTYCode
             MBMArchieve!MBMPostCode = MBM!MBMPostCode
             MBMArchieve!MBMTelnoH = MBM!MBMTelnoH
             MBMArchieve!MBMTelNoM = MBM!MBMTelNoM
             MBMArchieve!MBMTelNoO = MBM!MBMTelNoO
             MBMArchieve!MBMEmail = MBM!MBMEmail
             MBMArchieve!MBMDOB = MBM!MBMDOB
             MBMArchieve!NATCode = MBM!NATCode
             MBMArchieve!RACCode = MBM!RACCode
             MBMArchieve!MBMSex = MBM!MBMSex
             MBMArchieve!MBMPayorEffDate = MBM!MBMPayorEffDate
             MBMArchieve!MBMPayorEXPDate = MBM!MBMPayorEXPDate
             MBMArchieve!AgeCode = MBM!AgeCode
             MBMArchieve!INSCode = MBM!INSCode
             MBMArchieve!PLNCode = MBM!PLNCode
             MBMArchieve!MBMAdultChild = MBM!MBMAdultChild
             MBMArchieve!MBMAvailableLimit = MBM!MBMAvailableLimit
             MBMArchieve!SRVCodeList = MBM!SRVCodeList
             MBMArchieve!STACode = MBM!STACode
             MBMArchieve!LGNCode = MBM!LGNCode
             MBMArchieve!MBMBordxDate = MBM!MBMBordxDate
             MBMArchieve!MBMTakeOver = MBM!MBMTakeOver
             MBMArchieve!MBMRenewFlag = MBM!MBMRenewFlag
             MBMArchieve!BRNCode = MBM!BRNCode
             MBMArchieve!MBMStatus = "D"
             MBMArchieve!MBMValueDate = Date
             MBMArchieve!MBMPolicyYear = MBM!MBMPolicyYear
             MBMArchieve!MBMTakeOver = MBM!MBMTakeOver
             ' Need to check whether the member is from Johor port or not
             MBMArchieve!MBMMemberType = MBM!MBMMemberType
             MBMArchieve!MBMBatchNo = MBM!MBMBatchNo
             MBMArchieve!MBMDataStatus = "D"
        
        MBMArchieve.Update
        
        txtMBMNumber = MBM!MBMNumber
        txtMBMType = MBM!MBMMemberType
        
        Call DeleteMBMOthers
        
        If Trim(MBM!INSCode) = "F" Or Trim(MBM!INSCode) = "H" Then
            Call DeleteSupplementary
        End If

        DoEvents
        
        strdelPrincipal = ""
        
        strdelPrincipal = "Delete from MBM Where MBMNumber =  '" & txtMBMNumber & "'"
        DelPrincipal.Open strdelPrincipal, medixmasterconn, adOpenKeyset, adLockOptimistic
        
        If txtMBMType = "P" Then
            TotPrincipal = TotPrincipal + 1
        End If
    MBM.MoveNext
    Loop
    
    txtTotalPrincipals = TotPrincipal
    'txtTotalSupplementary1 = ""
    txtTotalSupplementary1 = TotSupplementary
    txtTotal = CInt(txtTotalPrincipals) + CInt(txtTotalSupplementary1)
    MBM.Close
    Set MBM = Nothing
    
    MsgBox "File Deleted Sucessfully...", vbOKOnly + vbInformation, DialogBoxTitle
    Screen.MousePointer = vbDefault
   End If
    'DelPrincipal.Close
    'Set DelPrincipal = Nothing
End Sub

Private Sub DeleteSupplementary()
Dim strsqlMBMCov As String
Dim MBMCov As New ADODB.Recordset
Dim MBMCoveredPersonsArchieve As New ADODB.Recordset
Dim strdelcov As String
Dim MBMPrincipalNumber As String
Dim MBMDelCov As New ADODB.Recordset
        

      strsqlMBMCov = "SELECT * From MBMCoveredPersons  Where MBMCNumber = '" & Trim(txtMBMNumber) & "'"
      MBMCov.Open strsqlMBMCov, medixmasterconn, adOpenKeyset, adLockOptimistic
      
      MBMCoveredPersonsArchieve.Open "MBMCoveredPersonsArchive", medixmasterconnA, adOpenKeyset, adLockOptimistic
      TotSupplementary = 0
      
      Do Until MBMCov.EOF
      MBMCoveredPersonsArchieve.AddNew
        MBMCoveredPersonsArchieve!MBMCNumber = MBMCov!MBMCNumber
        MBMCoveredPersonsArchieve!MBMCCoverID = MBMCov!MBMCCoverID
        MBMCoveredPersonsArchieve!MBMCSalutation = MBMCov!MBMCSalutation
        MBMCoveredPersonsArchieve!MBMCName = MBMCov!MBMCName
        MBMCoveredPersonsArchieve!MBMCIcBcPpNo = MBMCov!MBMCIcBcPpNo
        MBMCoveredPersonsArchieve!MBMCDOB = MBMCov!MBMCDOB
        'SelectAge = GetAge(txtMBMPayorEffDate, lstCovAdd.ListItems(listloop).SubItems(3))
        MBMCoveredPersonsArchieve!MBMCAGEBand = MBMCov!MBMCAGEBand
        MBMCoveredPersonsArchieve!MBMCSex = MBMCov!MBMCSex
        MBMCoveredPersonsArchieve!MBMCAdultChild = MBMCov!MBMCAdultChild
        MBMCoveredPersonsArchieve!MBMCRELCode = MBMCov!MBMCRELCode
        MBMCoveredPersonsArchieve!MBMCBloodType = MBMCov!MBMCBloodType
        MBMCoveredPersonsArchieve!MBMCAllergic = MBMCov!MBMCAllergic
        'MBMCoveredPersonsArchieve!MBMCExclusion = MBMCov!MBMCExclusion
        'If MBMtype = "P" Then
        MBMCoveredPersonsArchieve!MBMCMemberType = MBMCov!MBMCMemberType
        
        'Else
           '!MBMCMemberType = "Q"
        'MBM 'CoveredPersonsArchieve!MBMCMBMNumber = Trim(txtPayCode) & Trim(cboPLNCode) & Trim(cboINSCode) & Mid(BusinessYear, 3, 2) & GenerateMembershipNo(Trim(txtPayCode), Trim(cboPLNCode), Trim(cboINSCode))
        'End If
        'If Len(Trim(lstCovAdd.ListItems(listloop).SubItems(11))) Then
        MBMCoveredPersonsArchieve!MBMCAnnualLimit = MBMCov!MBMCAnnualLimit
        'End If
        MBMCoveredPersonsArchieve.Update
      DoEvents
        
        
          If MBMCov!MBMCMemberType <> "P" Then
              TotSupplementary = TotSupplementary + 1
          End If
    MBMCov.MoveNext
    Loop
    
    strdelcov = "Delete from MBMCoveredPersons Where MBMCNumber = '" & txtMBMNumber & "'"
    MBMDelCov.Open strdelcov, medixmasterconn, adOpenKeyset, adLockOptimistic

    MBMCov.Close
    Set MBMCov = Nothing
End Sub

Private Sub DeleteMBMOthers()
'On Error Resume Next
'On Error GoTo error1
    'Dim tmpPreMemberNo As String
Dim strsqlOthers As String
Dim MBMOthers As New ADODB.Recordset
Dim MBMOthersArchieve As New ADODB.Recordset
Dim DelOthers As String
Dim MBMOther As New ADODB.Recordset
        
    strsqlOthers = "SELECT * From MBMOthers where MBMNumber = '" & Trim(txtMBMNumber) & "'"
    MBMOthers.Open strsqlOthers, medixmasterconn, adOpenKeyset, adLockOptimistic
    
    MBMOthersArchieve.Open "MBMOthersArchive", medixmasterconnA, adOpenKeyset, adLockOptimistic
        
        MBMOthersArchieve.AddNew
        'With MBMOthers
            MBMOthersArchieve!MBMNumber = MBMOthers!MBMNumber
            MBMOthersArchieve!MBMAgentCode = MBMOthers!MBMAgentCode
            MBMOthersArchieve!MBMWeight = MBMOthers!MBMWeight
            MBMOthersArchieve!MBMHeight = MBMOthers!MBMHeight
            MBMOthersArchieve!MBMWorkNature = MBMOthers!MBMWorkNature
            MBMOthersArchieve!MBMOccupation = MBMOthers!MBMOccupation
            MBMOthersArchieve!MBMBlood = MBMOthers!MBMBlood
            MBMOthersArchieve!MBMAllergic = MBMOthers!MBMAllergic
            MBMOthersArchieve!MBMFirstMEMNo = MBMOthers!MBMFirstMEMNo
            MBMOthersArchieve!MBMGroupCompany = MBMOthers!MBMGroupCompany
            'MBMOthersArchieve!MBMExclusion = MBMOthers!MBMExclusion
            MBMOthersArchieve!MBMSmoking = MBMOthers!MBMSmoking
            MBMOthersArchieve!MBMAlcohol = MBMOthers!MBMAlcohol
            MBMOthersArchieve!MBMMaritalStatus = MBMOthers!MBMMaritalStatus
            MBMOthersArchieve!MBMPrevPolNo = MBMOthers!MBMPrevPolNo
            MBMOthersArchieve!MBMPrevMemNo = MBMOthers!MBMPrevMemNo
            MBMOthersArchieve!MBMEmployeeNo = MBMOthers!MBMEmployeeNo
            '!MBMDateJoin = Mid(TextLine, 599, 2) & "/" & Mid(TextLine, 601, 2) & "/" & Mid(TextLine, 603, 4)
            '!MBMClininal = Mid(TextLine, 580, 211)

        'End With
        MBMOthersArchieve.Update
        
        DelOthers = "Delete from MBMothers Where MBMNumber =  '" & txtMBMNumber & "'"
        MBMOther.Open DelOthers, medixmasterconn, adOpenKeyset, adLockOptimistic
        
    
    MBMOthers.Close
    'MBMOther.Close
 '   Exit Sub
'error1:
    
'    MsgBox ERR.Description
End Sub

Private Sub PrincipalListing()
Dim strsqlMBM As String
Dim MBM As New ADODB.Recordset
Dim MBMArchieve As New ADODB.Recordset
Dim strdelPrincipal As String
Dim DelPrincipal As New ADODB.Recordset
Dim DelOthers As String
Dim MBMOthers As New ADODB.Recordset
Dim rst2 As New ADODB.Recordset
    
    'strsqlMBM = "SELECT HLTCode, PAYCode,MBMMemberType, MBMNumber, "
    'strsqlMBM = strsqlMBM & "MBMPolicyNo, MBMName, MBMIcBcPp, INSCode, "
    'strsqlMBM = strsqlMBM & "MBMBatchNo, MBMBordxDate, MBMStatus, "
    'strsqlMBM = strsqlMBM & "MBMSalutation, MBMAddress1, MBMAddress2, MBMAddress3, "
    'strsqlMBM = strsqlMBM & "CTYCode, MBMPostCode, MBMTelnoH, MBMTelNoM, "
    'strsqlMBM = strsqlMBM & "From MBM Where " & Exportsql
    
    'strsqlMBM = "SELECT * From MBM where " & ExportsqlList
    'MBM.Open strsqlMBM, medixmasterconn, adOpenKeyset, adLockOptimistic
    'MBMArchieve.Open "MBMArchive", medixmasterconnA, adOpenKeyset, adLockOptimistic
    'MBMCoveredPersonsArchieve.Open "MBMCoveredPersonsArchieve", medixmasterconnA, adOpenKeyset, adLockOptimistic
    
    'TotPrincipal = 0


    'Do Until MBM.EOF
        
    'Do Until rst2.EOF
    '    Set CityMaster = lstCityList.listitems.Add(, , rst2!CTYCode)
    '        CityMaster.SubItems(1) = rst2!CTYDescription
    '    If Mid(ChkStr(rst2!CTYCode), 1, lenStrCity) = ChkStr(strCity) Then
    '        CityMaster.Selected = True
    '    End If
    '    rst2.MoveNext
    'Loop
End Sub

Public Sub DELMBMListing()
Dim MBM As New ADODB.Recordset
Dim DELsqlList As String
Dim MBMList As ListItem
Dim strsqlMBM As String

    If Mid(cboDeleteOption, 1, 1) = "B" Then
        DELsqlList = "HLTCode = '" & Mid(cboHLTCode, 1, InStr(1, cboHLTCode, "-") - 1) & "' And PAYCode = '" & Mid(cboPAYCode, 1, InStr(1, cboPAYCode, "-") - 1) & "' And MBMBordxDate = CONVERT(SMALLDATETIME, '" & txtReceiptDate & "', 103) And mbmbatchno ='" & txtBatchNo1 & "'"
    ElseIf Mid(cboDeleteOption, 1, 1) = "D" Then
        DELsqlList = "HLTCode = '" & Mid(cboHLTCode, 1, InStr(1, cboHLTCode, "-") - 1) & "' And PAYCode = '" & Mid(cboPAYCode, 1, InStr(1, cboPAYCode, "-") - 1) & "' And MBMBordxDate >= CONVERT(SMALLDATETIME, '" & txtBordxFrom & "', 103) And MBMBordxDate <= CONVERT(SMALLDATETIME, '" & txtBordxTo & "', 103)"
    ElseIf Mid(cboDeleteOption, 1, 1) = "M" Then
        DELsqlList = "HLTCode = '" & Mid(cboHLTCode, 1, InStr(1, cboHLTCode, "-") - 1) & "' And PAYCode = '" & Mid(cboPAYCode, 1, InStr(1, cboPAYCode, "-") - 1) & "' And MBMNumber >= '" & txtMBMNumberFrom & "' And MBMNumber <= '" & txtMBMNumberTo & "'"
    End If

    lstMBMList.listitems.Clear
    strsqlMBM = "SELECT * From MBM where " & DELsqlList
    MBM.Open strsqlMBM, medixmasterconn, adOpenKeyset, adLockOptimistic
    If MBM.recordCount = 0 Then
        MsgBox "No Record Found", vbOKOnly + vbInformation, DialogBoxTitle
        Screen.MousePointer = vbDefault
    Else
        Do Until MBM.EOF
        Set MBMList = lstMBMList.listitems.Add(, , MBM!MBMBordxDate)
            MBMList.SubItems(1) = MBM!MBMBatchNo
            MBMList.SubItems(2) = MBM!MBMNumber
            MBMList.SubItems(3) = "00"
            MBMList.SubItems(4) = MBM!MBMName
        txtMBMNumber = MBM!MBMNumber
        If Trim(MBM!INSCode) = "F" Or Trim(MBM!INSCode) = "H" Then
            Call DELSuppListing
        End If
        txtMBMType = MBM!MBMMemberType
        If txtMBMType = "P" Then
            TotPrincipal = TotPrincipal + 1
        End If

        MBM.MoveNext
    Loop
    txtTotalPrincipals = TotPrincipal
    'txtTotalSupplementary1 = ""
    txtTotalSupplementary1 = TotSupplementary
    txtTotal = CInt(txtTotalPrincipals) + CInt(txtTotalSupplementary1)
    MsgBox "File Listed Sucessfully...", vbOKOnly + vbInformation, DialogBoxTitle
    Screen.MousePointer = vbDefault
    End If
End Sub

Public Sub DELSuppListing()
Dim COV As New ADODB.Recordset
Dim strsql As String
Dim COVList As ListItem

    'If Mid(cboDeleteOption, 1, 1) = "B" Then
    '    DELsqlList = "HLTCode = '" & Mid(cboHLTCode, 1, InStr(1, cboHLTCode, "-") - 1) & "' And PAYCode = '" & Mid(cboPAYCode, 1, InStr(1, cboPAYCode, "-") - 1) & "' And MBMBordxDate = CONVERT(SMALLDATETIME, '" & txtReceiptDate & "', 103) And mbmbatchno ='" & txtBatchNo1 & "'"
    'ElseIf Mid(cboDeleteOption, 1, 1) = "D" Then
    '    DELsqlList = "HLTCode = '" & Mid(cboHLTCode, 1, InStr(1, cboHLTCode, "-") - 1) & "' And PAYCode = '" & Mid(cboPAYCode, 1, InStr(1, cboPAYCode, "-") - 1) & "' And MBMBordxDate >= CONVERT(SMALLDATETIME, '" & txtBordxFrom & "', 103) And MBMBordxDate <= CONVERT(SMALLDATETIME, '" & txtBordxTo & "', 103)"
    
    strsql = "SELECT * From MBMCoveredPersons where MBMCnumber = '" & txtMBMNumber & "'"
    COV.Open strsql, medixmasterconn, adOpenKeyset, adLockOptimistic
    
        Do Until COV.EOF
        If Len(COV!MBMCBordxDate) Then
        Set COVList = lstMBMList.listitems.Add(, , COV!MBMCBordxDate)
            If Len(COV!MBMCBatchNo) Then
                COVList.SubItems(1) = COV!MBMCBatchNo
            End If
            COVList.SubItems(2) = COV!MBMCNumber
            COVList.SubItems(3) = COV!MBMCCoverID
            COVList.SubItems(4) = COV!MBMCName
        If COV!MBMCMemberType <> "P" Then
              TotSupplementary = TotSupplementary + 1
        End If
        End If
        'Call DELSuppListing
        COV.MoveNext
    Loop
        
    
End Sub

Private Sub Form_Unload(Cancel As Integer)
'    Dim RecordExit
'    If bExit = False Then
'        RecordExit = MsgBox("Are you sure?", vbYesNo + vbExclamation)
'        If RecordExit = vbYes Then
            Unload Me
'        Else
'            Cancel = True
'        End If
'    End If
End Sub

Private Sub lstMBMList_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
lstMBMList.SortKey = ColumnHeader.Index - 1
lstMBMList.Sorted = True
If lstMBMList.SortOrder = lvwAscending Then
lstMBMList.SortOrder = lvwDescending
Else
lstMBMList.SortOrder = lvwAscending
End If

End Sub
