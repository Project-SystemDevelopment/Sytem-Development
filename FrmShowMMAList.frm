VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form FrmShowMMAList 
   Caption         =   "MMA Schedule"
   ClientHeight    =   4830
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   7860
   KeyPreview      =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4830
   ScaleWidth      =   7860
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame2 
      Height          =   615
      Left            =   0
      TabIndex        =   9
      Top             =   4200
      Width           =   7815
      Begin VB.CommandButton CmdExit 
         Caption         =   "E&xit"
         Height          =   285
         Left            =   6840
         TabIndex        =   10
         Top             =   240
         Width           =   855
      End
   End
   Begin VB.Frame Frame1 
      Height          =   855
      Left            =   0
      TabIndex        =   1
      Top             =   240
      Width           =   7815
      Begin VB.TextBox TxtCaseNo 
         Height          =   285
         Left            =   5400
         TabIndex        =   4
         Top             =   240
         Width           =   1815
      End
      Begin VB.TextBox TxtMBMNo 
         Height          =   285
         Left            =   1440
         TabIndex        =   2
         Top             =   240
         Width           =   2295
      End
      Begin VB.TextBox TxtMBMName 
         Height          =   285
         Left            =   1440
         TabIndex        =   3
         Top             =   480
         Width           =   5775
      End
      Begin VB.Label Label3 
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
         Left            =   4320
         TabIndex        =   7
         Top             =   240
         Width           =   1335
      End
      Begin VB.Label Label2 
         Caption         =   "Member Name"
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
         TabIndex        =   6
         Top             =   480
         Width           =   1335
      End
      Begin VB.Label Label1 
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
         TabIndex        =   5
         Top             =   240
         Width           =   1335
      End
   End
   Begin MSComctlLib.ListView lstMMA 
      Height          =   3000
      Left            =   0
      TabIndex        =   8
      Top             =   1200
      Width           =   7815
      _ExtentX        =   13785
      _ExtentY        =   5292
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
      NumItems        =   10
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Item"
         Object.Width           =   1235
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "CPT"
         Object.Width           =   1764
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "OCPS"
         Object.Width           =   1764
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Text            =   "Body(MMA)"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   4
         Text            =   "CPT Desc"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   5
         Text            =   "MMA Desc"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   6
         Text            =   "Surg's Cat."
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   7
         Text            =   "Anae's Cat."
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(9) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   8
         Text            =   "Surg's Fees"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(10) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   9
         Text            =   "Anae's Fees"
         Object.Width           =   2540
      EndProperty
   End
   Begin VB.Label Label12 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00404040&
      Caption         =   "Investigation"
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
      TabIndex        =   0
      Top             =   0
      Width           =   7845
   End
End
Attribute VB_Name = "FrmShowMMAList"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public Source As Integer

Private Sub CmdExit_Click()
    Unload Me
End Sub

Private Sub Form_Activate()
Dim strsql As String
Dim MMARec As New ADODB.Recordset
Dim LstREc As ListItem
Dim Cnt As String
Dim StrOperator As String
Dim InvCode1 As String
Dim InvCode2 As String
Dim InvCode3 As String
Dim ProcCode1 As String
Dim ProcCode2 As String
Dim ProcCode3 As String
Dim ProcCode4 As String
Dim ProcCode5 As String
Dim StrMBMNo As String
Dim StrMBMName As String
Dim StrCaseNo As String

    medixmasterconnMaint.Open "File Name=" & BOSPath & "Conn\MediConnHISMaint.UDL"

    If Source = 1 Then   'frmCaseAdjustment ~ cmdInvList
        InvCode1 = Trim(frmCaseAdjustment.txtINVCode1)
        InvCode2 = Trim(frmCaseAdjustment.txtINVCode2)
        InvCode3 = Trim(frmCaseAdjustment.txtINVCode3)
        StrMBMNo = Trim(frmCaseAdjustment.txtMBMNumber)
        StrMBMName = Trim(frmCaseAdjustment.TxtMBMName)
        StrCaseNo = Trim(frmCaseAdjustment.TxtCaseNo)
    End If
    If Source = 2 Then  'frmCaseAdjustment ~ cmdCPTList
        ProcCode1 = Trim(frmCaseAdjustment.txtProcCode1)
        ProcCode2 = Trim(frmCaseAdjustment.txtProcCode2)
        ProcCode3 = Trim(frmCaseAdjustment.txtProcCode3)
        ProcCode4 = Trim(frmCaseAdjustment.txtProcCode4)
        ProcCode5 = Trim(frmCaseAdjustment.txtProcCode5)
        StrMBMNo = Trim(frmCaseAdjustment.txtMBMNumber)
        StrMBMName = Trim(frmCaseAdjustment.TxtMBMName)
        StrCaseNo = Trim(frmCaseAdjustment.TxtCaseNo)
    End If
    If Source = 3 Then   'frmCaseInquiry ~ cmdInvList
        InvCode1 = Trim(frmCaseInquiry.txtINVCode1)
        InvCode2 = Trim(frmCaseInquiry.txtINVCode2)
        InvCode3 = Trim(frmCaseInquiry.txtINVCode3)
        StrMBMNo = Trim(frmCaseInquiry.txtMBMNumber)
        StrMBMName = Trim(frmCaseInquiry.TxtMBMName)
        StrCaseNo = Trim(frmCaseInquiry.TxtCaseNo)
    End If
    If Source = 4 Then  'frmCaseInquiry ~ cmdCPTList
        ProcCode1 = Trim(frmCaseInquiry.txtProcCode1)
        ProcCode2 = Trim(frmCaseInquiry.txtProcCode2)
        ProcCode3 = Trim(frmCaseInquiry.txtProcCode3)
        ProcCode4 = Trim(frmCaseInquiry.txtProcCode4)
        ProcCode5 = Trim(frmCaseInquiry.txtProcCode5)
        StrMBMNo = Trim(frmCaseInquiry.txtMBMNumber)
        StrMBMName = Trim(frmCaseInquiry.TxtMBMName)
        StrCaseNo = Trim(frmCaseInquiry.TxtCaseNo)
    End If
    TxtMBMNo = IIf(Len(StrMBMNo), StrMBMNo, "")
    TxtMBMName = IIf(Len(StrMBMName), StrMBMName, "")
    TxtCaseNo = IIf(Len(StrCaseNo), StrCaseNo, "")
    If Source = 1 Or Source = 3 Then
        If Trim(InvCode1) = "" And Trim(InvCode2) = "" And Trim(InvCode3) = "" Then
        Else
            
            StrOperator = " And "
            If Trim(InvCode1) <> "" Then
                strsql = " AND OPCSCode='" & Trim(InvCode1) & "'"
                StrOperator = " OR "
            End If
            If Trim(InvCode2) <> "" Then
                strsql = strsql & StrOperator & " OPCSCode='" & Trim(InvCode2) & "'"
                StrOperator = " OR "
            End If
            If Trim(InvCode3) <> "" Then
                strsql = strsql & StrOperator & " OPCSCode='" & Trim(InvCode3) & "'"
            End If
            
        End If
    ElseIf Source = 2 Or Source = 4 Then
        If Trim(ProcCode1) = "" And Trim(ProcCode2) = "" And Trim(ProcCode3) = "" & _
            Trim(ProcCode4) = "" And Trim(ProcCode5) = "" Then
        Else
        
            StrOperator = " And "
            If Trim(ProcCode1) <> "" Then
                strsql = " AND CPTCode='" & Trim(ProcCode1) & "'"
                StrOperator = " OR "
            End If
            If Trim(ProcCode2) <> "" Then
                strsql = strsql & StrOperator & " CPTCode='" & Trim(ProcCode2) & "'"
                StrOperator = " OR "
            End If
            If Trim(ProcCode3) <> "" Then
                strsql = strsql & StrOperator & " CPTCode='" & Trim(ProcCode3) & "'"
            End If
            If Trim(ProcCode4) <> "" Then
                strsql = strsql & StrOperator & " CPTCode='" & Trim(ProcCode4) & "'"
            End If
            If Trim(ProcCode5) <> "" Then
                strsql = strsql & StrOperator & " CPTCode='" & Trim(ProcCode5) & "'"
            End If
            
        End If
    End If

    If Trim(strsql) <> "" Then
        strsql = "SELECT * FROM VIEW_MMA_Search WHERE 0 =0 " & strsql
        MMARec.Open strsql, medixmasterconnMaint, adOpenKeyset, adLockOptimistic
        Cnt = 0
        
        'If MMARec.recordCount > 0 Then
            Do Until MMARec.EOF
                Cnt = Cnt + 1
                Set LstREc = lstMMA.listitems.Add(, , Cnt)
                With MMARec
                    LstREc.SubItems(1) = IIf(Len(!CPTCode), !CPTCode, "")
                    LstREc.SubItems(2) = IIf(Len(!OPCSCode), !OPCSCode, "")
                    LstREc.SubItems(3) = IIf(Len(!ProcMainBPCode), !ProcMainBPCode, "")
                    LstREc.SubItems(4) = IIf(Len(!CPTDesc), !CPTDesc, "")
                    LstREc.SubItems(5) = IIf(Len(!OPCSNarrative), !OPCSNarrative, "")
                    LstREc.SubItems(6) = IIf(Len(!ProcSurgCat), !ProcSurgCat, "")
                    LstREc.SubItems(7) = IIf(Len(!ProcAnaesCat), !ProcAnaesCat, "")
                    LstREc.SubItems(8) = IIf(Len(!SurgeonsFees), !SurgeonsFees, "")
                    LstREc.SubItems(9) = IIf(Len(!AnaesthetistsFees), !AnaesthetistsFees, "")
                
                End With
                MMARec.MoveNext
            Loop
        'End If
    End If
    medixmasterconnMaint.Close
    Set medixmasterconnMaint = Nothing
    
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

