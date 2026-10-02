VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmClaimHistory 
   Caption         =   "Claims'  History"
   ClientHeight    =   6480
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   8235
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   6480
   ScaleWidth      =   8235
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox txtMBMNumber 
      Height          =   285
      Left            =   6240
      TabIndex        =   8
      Top             =   120
      Width           =   1815
   End
   Begin VB.TextBox txtCasNumber 
      Height          =   285
      Left            =   1680
      TabIndex        =   7
      Top             =   120
      Width           =   2895
   End
   Begin VB.Frame Frame6 
      Caption         =   "Outstanding Claims"
      Height          =   1935
      Left            =   120
      TabIndex        =   4
      Top             =   600
      Width           =   7935
      Begin MSComctlLib.ListView lstClaimsList 
         Height          =   1320
         Left            =   360
         TabIndex        =   5
         Top             =   360
         Width           =   7245
         _ExtentX        =   12779
         _ExtentY        =   2328
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
         NumItems        =   7
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Case No"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Membership No"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "Name"
            Object.Width           =   5292
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   3
            Text            =   "Date Of Loss"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   4
            Text            =   "Status"
            Object.Width           =   1411
         EndProperty
         BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   5
            Text            =   "Patient Cover ID"
            Object.Width           =   1764
         EndProperty
         BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   6
            Text            =   "Remark"
            Object.Width           =   2540
         EndProperty
      End
   End
   Begin VB.Frame Frame7 
      Caption         =   "Closed Claims"
      Height          =   1935
      Left            =   120
      TabIndex        =   2
      Top             =   2520
      Width           =   7935
      Begin MSComctlLib.ListView lstClosedClaims 
         Height          =   1320
         Left            =   360
         TabIndex        =   3
         Top             =   360
         Width           =   7245
         _ExtentX        =   12779
         _ExtentY        =   2328
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
         NumItems        =   7
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Case No"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Membership No"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "Name"
            Object.Width           =   5292
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   3
            Text            =   "Date Of Loss"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   4
            Text            =   "Status"
            Object.Width           =   1411
         EndProperty
         BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   5
            Text            =   "Patient Cover ID"
            Object.Width           =   1764
         EndProperty
         BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   6
            Text            =   "Remark"
            Object.Width           =   2540
         EndProperty
      End
   End
   Begin VB.Frame Frame12 
      Caption         =   "Rejected Claims"
      Height          =   1935
      Left            =   120
      TabIndex        =   0
      Top             =   4560
      Width           =   7935
      Begin MSComctlLib.ListView lstRejectedClaims 
         Height          =   1320
         Left            =   240
         TabIndex        =   1
         Top             =   360
         Width           =   7245
         _ExtentX        =   12779
         _ExtentY        =   2328
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
         NumItems        =   7
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Case No"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Membership No"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "Name"
            Object.Width           =   5292
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   3
            Text            =   "Date Of Loss"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   4
            Text            =   "Status"
            Object.Width           =   1411
         EndProperty
         BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   5
            Text            =   "Patient Cover ID"
            Object.Width           =   1764
         EndProperty
         BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   6
            Text            =   "Remark"
            Object.Width           =   2540
         EndProperty
      End
   End
   Begin VB.Label Label1 
      Caption         =   "Member Number"
      Height          =   255
      Left            =   4800
      TabIndex        =   9
      Top             =   120
      Width           =   1335
   End
   Begin VB.Label Label2 
      Caption         =   "Case Number"
      Height          =   255
      Left            =   120
      TabIndex        =   6
      Top             =   120
      Width           =   1455
   End
End
Attribute VB_Name = "frmClaimHistory"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
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

' ############## Claims History ###############################
Public Sub ListCases(CaseStatus As String, MBMNumber As String, Optional CaseNumber As String)
    Dim CASList As ListItem
    Dim strsql As String
    Dim CAS As New ADODB.Recordset
    
    strsql = "SELECT CASNumber, MBMNumber, " & _
        " MBMPatientCovID, MBMPolicyNo, MBMName, " & _
        " INSCode, PLNCode, CASStatus, " & _
        " CASRemark, CaseDateOfLoss , MBMPayorEffDate , HLTCode"
    strsql = strsql & " from VIEW_Cas_MBM_Claims where CASStatus = '" & Trim(CaseStatus) & "'"
    
    If Trim(MBMNumber) <> "" Then
        strsql = strsql & " And MBMNumber = '" & Trim(MBMNumber) & "'"
    End If
    
    If Trim(CaseNumber) <> "" Then
        strsql = strsql & " And CASNumber <> '" & Trim(CaseNumber) & "'"
    End If
    
    
    CAS.Open strsql, medixmasterconn, adOpenKeyset, adLockOptimistic
    If ChkStr(CaseStatus) = "O" Then
         lstClaimsList.listitems.Clear
    ElseIf ChkStr(CaseStatus) = "C" Then
         lstClosedClaims.listitems.Clear
    End If
    
    If CAS.recordCount <> 0 Then
        Screen.MousePointer = vbHourglass
        Do
         Do Until CAS.EOF
            If ChkStr(CaseStatus) = "O" Then
                Set CASList = lstClaimsList.listitems.Add(, , CAS!CASNumber)
            ElseIf ChkStr(CaseStatus) = "C" Then
                Set CASList = lstClosedClaims.listitems.Add(, , CAS!CASNumber)
            Else
                Set CASList = lstRejectedClaims.listitems.Add(, , CAS!CASNumber)
            End If
                CASList.SubItems(1) = CAS!MBMNumber
                CASList.SubItems(2) = CAS!MBMName
                If Len(Trim(CAS!CaseDateOfLoss)) Then
                    CASList.SubItems(3) = CAS!CaseDateOfLoss
                End If
                CASList.SubItems(4) = CAS!CASStatus
                CASList.SubItems(5) = CAS!MBMPatientCovID
                If Len(Trim(CAS!CASRemark)) Then
                    CASList.SubItems(6) = CAS!CASRemark
                End If
                CAS.MoveNext
                        
            DoEvents
            If intExit = 1 Then
                check = False
                Exit Do
            End If
         Loop
        Loop Until check = False
        Screen.MousePointer = vbDefault
    End If
    intExit = 0
    CAS.Close
    Set CAS = Nothing
End Sub
