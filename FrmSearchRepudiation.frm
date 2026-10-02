VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "mscomctl.ocx"
Begin VB.Form FrmSearchRepudiation 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Search Repudiation"
   ClientHeight    =   4995
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6210
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4995
   ScaleWidth      =   6210
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdPrint 
      Caption         =   "&Print"
      Height          =   315
      Left            =   2280
      TabIndex        =   5
      Top             =   4600
      Width           =   840
   End
   Begin VB.CommandButton CmdClear 
      Caption         =   "&Clear"
      Height          =   315
      Left            =   4200
      TabIndex        =   7
      Top             =   4600
      Width           =   840
   End
   Begin VB.Frame Frame10 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1755
      Left            =   120
      TabIndex        =   4
      Top             =   0
      Width           =   6045
      Begin VB.OptionButton optPremlim 
         Caption         =   "Premlim. Diag"
         Height          =   255
         Left            =   120
         TabIndex        =   15
         Top             =   240
         Width           =   1335
      End
      Begin VB.OptionButton optFinal 
         Caption         =   "Final Diag"
         Height          =   255
         Left            =   1560
         TabIndex        =   14
         Top             =   240
         Width           =   1095
      End
      Begin VB.OptionButton optUnderlying 
         Caption         =   "Underlying Disease"
         Height          =   255
         Left            =   2760
         TabIndex        =   13
         Top             =   240
         Width           =   1695
      End
      Begin VB.OptionButton optRep 
         Caption         =   "Rep. Reasons"
         Height          =   255
         Left            =   4560
         TabIndex        =   12
         Top             =   240
         Width           =   1335
      End
      Begin VB.TextBox txtRepDesc 
         Height          =   285
         Left            =   1560
         MaxLength       =   1000
         TabIndex        =   1
         Top             =   960
         Width           =   3375
      End
      Begin VB.ComboBox cboCriteria 
         Height          =   315
         Left            =   5040
         TabIndex        =   2
         Text            =   "AND"
         Top             =   960
         Width           =   735
      End
      Begin VB.TextBox txtRepDesc2 
         Height          =   285
         Left            =   1560
         MaxLength       =   1000
         TabIndex        =   3
         Top             =   1320
         Width           =   3375
      End
      Begin VB.TextBox TxtRepCode 
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
         Left            =   1560
         MaxLength       =   10
         TabIndex        =   0
         Top             =   600
         Width           =   615
      End
      Begin VB.Label Label33 
         Caption         =   "Rep. Description"
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
         Left            =   240
         TabIndex        =   10
         Top             =   960
         Width           =   1275
      End
      Begin VB.Label Label32 
         Caption         =   "Rep. Code"
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
         Left            =   240
         TabIndex        =   9
         Top             =   600
         Width           =   1335
      End
   End
   Begin VB.CommandButton cmdSearch 
      Caption         =   "&Search"
      Default         =   -1  'True
      Height          =   315
      Left            =   3240
      TabIndex        =   6
      Top             =   4600
      Width           =   840
   End
   Begin VB.CommandButton CmdExit 
      Caption         =   "E&xit"
      Height          =   315
      Left            =   5160
      TabIndex        =   8
      Top             =   4600
      Width           =   840
   End
   Begin MSComctlLib.ListView lstRepudiationList 
      Height          =   2715
      Left            =   120
      TabIndex        =   11
      Top             =   1800
      Width           =   6060
      _ExtentX        =   10689
      _ExtentY        =   4789
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
      NumItems        =   3
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Repudiation Code"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Repudiation Description"
         Object.Width           =   8996
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "Non-Disc Ind"
         Object.Width           =   1764
      EndProperty
   End
End
Attribute VB_Name = "FrmSearchRepudiation"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public Source As Integer
Private m_blnDirection(1 To 2) As Boolean
Private NonDiscStatus As Boolean
Public Paycoce As String

Private Sub cmdPrint_Click()
    If lstRepudiationList.ListItems.Count = 0 Then
        MsgBox "No record found.", vbOKOnly + vbCritical, DialogBoxTitle
        Exit Sub
    Else
        Screen.MousePointer = vbHourglass
        Call ExportListViewtoExcel(lstRepudiationList)
        Screen.MousePointer = Default
    End If
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

Private Sub cmdExit_Click()
    Unload Me
End Sub

Private Sub CmdClear_Click()
TxtRepCode.Text = ""
txtRepDesc.Text = ""
txtRepDesc2.Text = ""
cboCriteria.Text = "AND"
lstRepudiationList.ListItems.Clear
End Sub


Private Sub Form_Load()
    cboCriteria.AddItem ("AND")
    cboCriteria.AddItem ("OR")
    optRep.Value = True
End Sub


Private Sub RepListing(Optional StrAppend As String)
    Dim rst2 As New ADODB.Recordset
    Dim RepMaster As ListItem
    Dim sql As String
     
    lstRepudiationList.ListItems.Clear
        sql = " SELECT RepCode, RepDescription, RepNonDisc " & _
        " from RepCode where 0=0 "
    
    If Len(Trim(StrAppend)) Then
        sql = sql + StrAppend
    End If

    rst2.Open sql, medixmasterconn, adOpenForwardOnly, adLockBatchOptimistic
    If rst2.EOF = True Then
        MsgBox "Record not found", vbInformation + vbOKOnly, DialogBoxTitle
    Else
        Do Until rst2.EOF
            Set RepMaster = lstRepudiationList.ListItems.Add(, , rst2!REPCode)
                If Len(rst2!REPDescription) Then
                    RepMaster.SubItems(1) = rst2!REPDescription
                End If
                If Len(rst2!RepNonDisc) Then
                    RepMaster.SubItems(2) = rst2!RepNonDisc
                End If
                
                rst2.MoveNext
        Loop
    rst2.Close
    Set rst2 = Nothing
    End If
End Sub

Public Sub getDisclouser(strdisc As String, strnondisc As String, strsubdisc As String, strSubDisclouser1 As String)
Dim PAYCode As String
Dim PVRec As New ADODB.Recordset
Dim SqlQry As String
Dim RepMaster As ListItem
PAYCode = Paycoce

lstRepudiationList.ListItems.Clear

'SqlQry = ""
'SqlQry = "EXEC  GetDisclouserSearch '" & Trim(PAYCode) & "','" & strdisc & "'"

PVRec.Open strdisc, medixmasterconnMaint, adOpenForwardOnly, adLockOptimistic
Do While Not PVRec.EOF
     Set RepMaster = lstRepudiationList.ListItems.Add(, , PVRec!RepCatCode)
     RepMaster.SubItems(1) = (PVRec!DisClouser & "||" & PVRec!RepCatCode)
    PVRec.MoveNext
Loop
PVRec.Close
Set PVRec = Nothing


'SqlQry = "EXEC  GetSubDisclouserSearch '" & Trim(PAYCode) & "','" & strsubdisc & "'"

strsubdisc = strsubdisc + " " + strSubDisclouser1
PVRec.Open strsubdisc, medixmasterconnMaint, adOpenForwardOnly, adLockOptimistic
Do While Not PVRec.EOF
    Set RepMaster = lstRepudiationList.ListItems.Add(, , PVRec!RepCatCode)
    RepMaster.SubItems(1) = (PVRec!SubDisclouser & "||" & PVRec!RepCatCode)
    PVRec.MoveNext
Loop
PVRec.Close
Set PVRec = Nothing

'SqlQry = "EXEC  GetNonDisclousersearch '" & Trim(PAYCode) & "','" & strnondisc & "'"
PVRec.Open strnondisc, medixmasterconnMaint, adOpenForwardOnly, adLockOptimistic
Do While Not PVRec.EOF
    Set RepMaster = lstRepudiationList.ListItems.Add(, , PVRec!RepCatCode)
    RepMaster.SubItems(1) = (PVRec!NonDisclouser & "||" & PVRec!RepCatCode)
    PVRec.MoveNext
Loop
PVRec.Close
Set PVRec = Nothing
End Sub

Private Sub CmdSearch_Click()
    'Old code
    'Dim strCriteria As String
    'strCriteria = ""
   '
   ' cmdSearch.Enabled = False'

    'If Len(Trim(TxtRepCode)) Then
    '        strCriteria = " AND RepCode like '" & Trim(TxtRepCode) & "%'"
    'End If
    'If Len(Trim(txtRepDesc)) Then
    '        strCriteria = strCriteria & " AND RepDescription like '%" & Trim(txtRepDesc) & "%' "
    'End If'

    'If Len(Trim(txtRepDesc2)) Then
    '        strCriteria = strCriteria & cboCriteria & " RepDescription like '%" & Trim(txtRepDesc2) & "%'"
    'End If
    'Call RepListing(strCriteria)
    'New Code
    
    Dim strDisclouser As String
    Dim strNonDisclouser As String
    Dim strSubDisclouser As String
    
    Dim strSubDisclouser1 As String
    
    strDisclouser = "Select Tbl_Disclouser.DisclouserID,Tbl_Disclouser.Disclouser,Tbl_Disclouser.RepCatCode from Tbl_PaycodeDisclouser" & _
                    " inner join Tbl_Disclouser on Tbl_Disclouser.DisclouserID = Tbl_PaycodeDisclouser.DisclouserID " & _
                    " where Tbl_PaycodeDisclouser.PayCode='" + Paycoce + "' "
    strNonDisclouser = "Select Tbl_NonDisclouser.DisclouserID,Tbl_NonDisclouser.NonDisclouser,Tbl_NonDisclouser.RepCatCode from Tbl_PaycodeDisclouser" & _
                       " inner join Tbl_NonDisclouser on Tbl_NonDisclouser.DisclouserID = Tbl_PaycodeDisclouser.DisclouserID " & _
                       " where Tbl_PaycodeDisclouser.PayCode='" + Paycoce + "' "
    strSubDisclouser = "Select Tbl_Subdisclouser.DisclouserID,Tbl_Subdisclouser.SubDisclouser,Tbl_Subdisclouser.RepCatCode from Tbl_PaycodeDisclouser" & _
                       " inner join Tbl_Disclouser on Tbl_Disclouser.DisclouserID = Tbl_PaycodeDisclouser.DisclouserID " & _
                       " inner join Tbl_Subdisclouser on  Tbl_Subdisclouser.DID = Tbl_Disclouser.DID and Tbl_Subdisclouser.DisclouserID = Tbl_Disclouser.DisclouserID " & _
                       " Where Tbl_PaycodeDisclouser.PAYCode = '" + Paycoce + "' "
    
    If Len(Trim(TxtRepCode)) Then
            strDisclouser = strDisclouser & " AND Tbl_Disclouser.RepCatCode like '" & Trim(TxtRepCode) & "%'"
            strNonDisclouser = strNonDisclouser & " AND Tbl_NonDisclouser.RepCatCode like '" & Trim(TxtRepCode) & "%'"
            strSubDisclouser1 = strSubDisclouser1 & " AND Tbl_Subdisclouser.RepCatCode like '" & Trim(TxtRepCode) & "%'"
    End If
    If Len(Trim(txtRepDesc)) Then
            strDisclouser = strDisclouser & " AND Tbl_Disclouser.Disclouser like '%" & Trim(txtRepDesc) & "%' "
            strNonDisclouser = strNonDisclouser & " AND Tbl_NonDisclouser.NonDisclouser like '%" & Trim(txtRepDesc) & "%' "
            strSubDisclouser1 = strSubDisclouser1 & " AND Tbl_Subdisclouser.SubDisclouser like '%" & Trim(txtRepDesc) & "%' "
    End If

    If Len(Trim(txtRepDesc2)) Then
            strDisclouser = strDisclouser & cboCriteria & " Tbl_Disclouser.Disclouser like '%" & Trim(txtRepDesc2) & "%'"
            strNonDisclouser = strNonDisclouser & cboCriteria & " Tbl_NonDisclouser.NonDisclouser like '%" & Trim(txtRepDesc2) & "%'"
            strSubDisclouser1 = strSubDisclouser1 & cboCriteria & " Tbl_Subdisclouser.SubDisclouser like '%" & Trim(txtRepDesc2) & "%'"
    End If
    
    Call getDisclouser(strDisclouser, strNonDisclouser, strSubDisclouser, strSubDisclouser1)
    'Call RepListing(strCriteria)
    cmdSearch.Enabled = True
End Sub

'**********************Start Sorted ListView *******************************
Private Sub lstRepudiationList_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
    m_blnDirection(ColumnHeader.Index) = Not m_blnDirection(ColumnHeader.Index)
    Select Case ColumnHeader.Index
    Case 1
        SortListView lstRepudiationList, ColumnHeader.Index, ldtString, m_blnDirection(1)
    Case 2
        SortListView lstRepudiationList, ColumnHeader.Index, ldtString, m_blnDirection(2)
    End Select
End Sub
'**********************End Sorted ListView *******************************

Private Sub lstRepudiationList_DblClick()
    If lstRepudiationList.ListItems.Count <> 0 Then
        If optRep.Value = True Then
            If Source = 1 Then
                If Len(Trim(frmCaseAdjustment2016.cboREPCode1)) = 0 Then
                    frmCaseAdjustment2016.cboREPCode1 = lstRepudiationList.SelectedItem.Text & " - " & Trim(lstRepudiationList.SelectedItem.SubItems(1))
                ElseIf Len(Trim(frmCaseAdjustment2016.cboREPCode2)) = 0 Then
                    frmCaseAdjustment2016.cboREPCode2 = lstRepudiationList.SelectedItem.Text & " - " & Trim(lstRepudiationList.SelectedItem.SubItems(1))
                Else
                    frmCaseAdjustment2016.cboREPCode3 = lstRepudiationList.SelectedItem.Text & " - " & Trim(lstRepudiationList.SelectedItem.SubItems(1))
                End If
                
                If lstRepudiationList.SelectedItem.SubItems(2) = "Y" Then
                    MsgBox "Please Enter Non-Disclosure Information", vbOKOnly + vbInformation, DialogBoxTitle
                    Unload FrmSearchRepudiation
                End If
            ElseIf Source = 2 Then
                If Len(Trim(frmCaseAdmission.cboREPCode1)) = 0 Then
                    frmCaseAdmission.cboREPCode1 = lstRepudiationList.SelectedItem.Text & " - " & Trim(lstRepudiationList.SelectedItem.SubItems(1))
                ElseIf Len(Trim(frmCaseAdmission.cboREPCode2)) = 0 Then
                    frmCaseAdmission.cboREPCode2 = lstRepudiationList.SelectedItem.Text & " - " & Trim(lstRepudiationList.SelectedItem.SubItems(1))
                Else
                    frmCaseAdmission.cboREPCode3 = lstRepudiationList.SelectedItem.Text & " - " & Trim(lstRepudiationList.SelectedItem.SubItems(1))
                End If
                
            ElseIf Source = 3 Then
                If Len(Trim(frmMBMExpiredPolicy.cboREPCode1)) = 0 Then
                    frmMBMExpiredPolicy.cboREPCode1 = lstRepudiationList.SelectedItem.Text & " - " & Trim(lstRepudiationList.SelectedItem.SubItems(1))
                ElseIf Len(Trim(frmMBMExpiredPolicy.cboREPCode2)) = 0 Then
                    frmMBMExpiredPolicy.cboREPCode2 = lstRepudiationList.SelectedItem.Text & " - " & Trim(lstRepudiationList.SelectedItem.SubItems(1))
                Else
                    frmMBMExpiredPolicy.cboREPCode3 = lstRepudiationList.SelectedItem.Text & " - " & Trim(lstRepudiationList.SelectedItem.SubItems(1))
                End If
                
            End If
        End If
    End If
End Sub

Private Sub optFinal_Click()
    Unload Me
    frmSearchICD.Options = 2
    frmSearchICD.Source = Source
    frmSearchICD.Show vbModal
End Sub

Private Sub optPremlim_Click()
    Unload Me
    frmSearchICD.Options = 1
    frmSearchICD.Source = Source
    frmSearchICD.Show vbModal
End Sub

Private Sub optUnderlying_Click()
    Unload Me
    frmSearchICD.Options = 3
    frmSearchICD.Source = Source
    frmSearchICD.Show vbModal
End Sub
