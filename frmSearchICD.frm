VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "mscomctl.ocx"
Begin VB.Form frmSearchICD 
   AutoRedraw      =   -1  'True
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Search ICD"
   ClientHeight    =   5025
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6210
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5025
   ScaleWidth      =   6210
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdPrint 
      Caption         =   "&Print"
      Height          =   315
      Left            =   2400
      TabIndex        =   5
      Top             =   4560
      Width           =   840
   End
   Begin VB.CommandButton CmdClear 
      Caption         =   "&Clear"
      Height          =   315
      Left            =   4320
      TabIndex        =   7
      Top             =   4560
      Width           =   840
   End
   Begin VB.CommandButton cmdExit 
      Caption         =   "Exit"
      Height          =   315
      Left            =   5280
      TabIndex        =   8
      Top             =   4560
      Width           =   840
   End
   Begin VB.CommandButton cmdSearch 
      Caption         =   "&Search"
      Default         =   -1  'True
      Height          =   315
      Left            =   3360
      TabIndex        =   6
      Top             =   4560
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
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   6165
      Begin VB.OptionButton optRep 
         Caption         =   "Rep. Reasons"
         Height          =   255
         Left            =   4680
         TabIndex        =   15
         Top             =   240
         Width           =   1335
      End
      Begin VB.OptionButton optUnderlying 
         Caption         =   "Underlying Disease"
         Height          =   255
         Left            =   2880
         TabIndex        =   14
         Top             =   240
         Width           =   1695
      End
      Begin VB.OptionButton optFinal 
         Caption         =   "Final Diag"
         Height          =   255
         Left            =   1680
         TabIndex        =   13
         Top             =   240
         Width           =   1095
      End
      Begin VB.OptionButton optPremlim 
         Caption         =   "Premlim. Diag"
         Height          =   255
         Left            =   240
         TabIndex        =   12
         Top             =   240
         Width           =   1335
      End
      Begin VB.TextBox TxtICDCode 
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
         MaxLength       =   3
         TabIndex        =   1
         Top             =   600
         Width           =   375
      End
      Begin VB.TextBox txtICDDesc2 
         Height          =   285
         Left            =   1560
         TabIndex        =   4
         Top             =   1320
         Width           =   3375
      End
      Begin VB.ComboBox cboCriteria 
         Height          =   315
         Left            =   5040
         TabIndex        =   3
         Text            =   "AND"
         Top             =   960
         Width           =   735
      End
      Begin VB.TextBox txtICDDesc 
         Height          =   285
         Left            =   1560
         TabIndex        =   2
         Top             =   960
         Width           =   3375
      End
      Begin VB.Label Label32 
         Caption         =   "ICD Code"
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
         Top             =   600
         Width           =   1095
      End
      Begin VB.Label Label33 
         Caption         =   "ICD Description"
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
         Top             =   960
         Width           =   1275
      End
   End
   Begin MSComctlLib.ListView lstICDList 
      Height          =   2715
      Left            =   0
      TabIndex        =   11
      Top             =   1800
      Width           =   6180
      _ExtentX        =   10901
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
      NumItems        =   2
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "ICD Code"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "ICD Description"
         Object.Width           =   8996
      EndProperty
   End
End
Attribute VB_Name = "frmSearchICD"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public Source As Integer
Public Options As Integer
Private m_blnDirection(1 To 2) As Boolean

Private Sub cmdExit_Click()
    Unload Me
End Sub

Private Sub cmdPrint_Click()
    If CheckAccess(158) = True Then 'Invoice Entry
    
        If lstICDList.ListItems.Count = 0 Then
            MsgBox "No record found.", vbOKOnly + vbCritical, DialogBoxTitle
            Exit Sub
        Else
            Screen.MousePointer = vbHourglass
            Call ExportListViewtoExcel(lstICDList)
            Screen.MousePointer = Default
        End If
    Else

       MsgBox "Print Function Denied!", vbOKOnly + vbCritical, DialogBoxTitle
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

Private Sub CmdClear_Click()
TxtICDCode.Text = ""
txtICDDesc.Text = ""
txtICDDesc2.Text = ""
cboCriteria.Text = "AND"
lstICDList.ListItems.Clear

End Sub

Private Sub Form_Activate()
    txtICDDesc.SetFocus
End Sub

Private Sub Form_Load()
    cboCriteria.AddItem ("AND")
    cboCriteria.AddItem ("OR")
    
    If Options = 1 Then
        optPremlim.Value = True
    ElseIf Options = 2 Then
        optFinal.Value = True
    ElseIf Options = 3 Then
        optUnderlying.Value = True
    ElseIf Options = 0 Then
        optPremlim.Value = True
    End If
End Sub

Private Sub ICDListing(Optional strAppend As String)
    Dim rst2 As New ADODB.Recordset
    Dim ICDMaster As ListItem
    Dim sql As String
     
    lstICDList.ListItems.Clear
        sql = " SELECT ICDCode, ICDDescription " & _
        " from ICD where 0=0 "
    
    If Len(Trim(strAppend)) Then
        sql = sql + strAppend
    End If

    rst2.Open sql, medixmasterconn, adOpenKeyset, adLockOptimistic
    If rst2.recordCount = 0 Then
        MsgBox "Record not found", vbInformation + vbOKOnly, DialogBoxTitle
    Else
        Do Until rst2.EOF
            Set ICDMaster = lstICDList.ListItems.Add(, , rst2!ICDCode)
                If Len(rst2!ICDDescription) Then
                    ICDMaster.SubItems(1) = Trim(rst2!ICDDescription)
                End If
                rst2.MoveNext
        Loop
    rst2.Close
    Set rst2 = Nothing
    End If
End Sub

Private Sub cmdSearch_Click()
    Dim strCriteria As String
    
    Screen.MousePointer = vbHourglass
    cmdSearch.Enabled = False
    'medixmasterconn.Open "File Name=" & BOSPath & "Conn\mediconnect.UDL"

    strCriteria = ""
    If Len(Trim(TxtICDCode)) Then
            strCriteria = " AND ICDCOde = '" & Trim(TxtICDCode) & "'"
    End If
    If Len(Trim(txtICDDesc)) Then
            strCriteria = strCriteria & " AND ICDDescription like '%" & Trim(txtICDDesc) & "%' "
    End If

    If Len(Trim(txtICDDesc2)) Then
            strCriteria = strCriteria & cboCriteria & " ICDDescription like '%" & Trim(txtICDDesc2) & "%'"
    End If
    
    Call ICDListing(strCriteria)
    'medixmasterconn.Close
    'Set medixmasterconn = Nothing
    cmdSearch.Enabled = True
    Screen.MousePointer = vbDefault
    
End Sub

'**********************Start Sorted ListView *******************************
Private Sub lstICDList_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
    m_blnDirection(ColumnHeader.Index) = Not m_blnDirection(ColumnHeader.Index)
    Select Case ColumnHeader.Index
    Case 1
        SortListView lstICDList, ColumnHeader.Index, ldtString, m_blnDirection(1)
    Case 2
        SortListView lstICDList, ColumnHeader.Index, ldtString, m_blnDirection(2)
    End Select
End Sub
'**********************End Sorted ListView *******************************

Private Sub lstICDList_DblClick()
    
    If lstICDList.ListItems.Count <> 0 Then
        If Source = 1 Then
            If optPremlim.Value = True Then
                If Len(Trim(frmCaseAdjustment2016.cboICDCode(0))) = 0 Then
                    frmCaseAdjustment2016.cboICDCode(0) = lstICDList.SelectedItem.Text
                ElseIf Len(Trim(frmCaseAdjustment2016.cboICDCode(1))) = 0 Then
                    frmCaseAdjustment2016.cboICDCode(1) = lstICDList.SelectedItem.Text
                Else
                    frmCaseAdjustment2016.cboICDCode(2) = lstICDList.SelectedItem.Text
                End If
            ElseIf optFinal.Value = True Then
                If Len(Trim(frmCaseAdjustment2016.cboFDICDCode1)) = 0 Then
                    frmCaseAdjustment2016.cboFDICDCode1 = lstICDList.SelectedItem.Text
                ElseIf Len(Trim(frmCaseAdjustment2016.cboFDICDCode2)) = 0 Then
                    frmCaseAdjustment2016.cboFDICDCode2 = lstICDList.SelectedItem.Text
                Else
                    frmCaseAdjustment2016.cboFDICDCode3 = lstICDList.SelectedItem.Text
                End If
            ElseIf optUnderlying.Value = True Then
                If Len(Trim(frmCaseAdjustment2016.cboUDICDCode1)) = 0 Then
                    frmCaseAdjustment2016.cboUDICDCode1 = lstICDList.SelectedItem.Text
                ElseIf Len(Trim(frmCaseAdjustment2016.cboUDICDCode2)) = 0 Then
                    frmCaseAdjustment2016.cboUDICDCode2 = lstICDList.SelectedItem.Text
                Else
                    frmCaseAdjustment2016.cboUDICDCode3 = lstICDList.SelectedItem.Text
                End If
            End If
        ElseIf Source = 2 Then
            If optPremlim.Value = True Then
                If Len(Trim(frmCaseAdmission.cboICDCode1)) = 0 Then
                    frmCaseAdmission.cboICDCode1 = lstICDList.SelectedItem.Text
                ElseIf Len(Trim(frmCaseAdmission.cboICDCode2)) = 0 Then
                    frmCaseAdmission.cboICDCode2 = lstICDList.SelectedItem.Text
                Else
                    frmCaseAdmission.cboICDCode3 = lstICDList.SelectedItem.Text
                End If
            ElseIf optFinal.Value = True Then
                If Len(Trim(frmCaseAdmission.cboFDICDCode1)) = 0 Then
                    frmCaseAdmission.cboFDICDCode1 = lstICDList.SelectedItem.Text
                ElseIf Len(Trim(frmCaseAdmission.cboFDICDCode2)) = 0 Then
                    frmCaseAdmission.cboFDICDCode2 = lstICDList.SelectedItem.Text
                Else
                    frmCaseAdmission.cboFDICDCode3 = lstICDList.SelectedItem.Text
                End If
            ElseIf optUnderlying.Value = True Then
                If Len(Trim(frmCaseAdmission.cboUDICDCode1)) = 0 Then
                    frmCaseAdmission.cboUDICDCode1 = lstICDList.SelectedItem.Text
                ElseIf Len(Trim(frmCaseAdmission.cboUDICDCode2)) = 0 Then
                    frmCaseAdmission.cboUDICDCode2 = lstICDList.SelectedItem.Text
                Else
                    frmCaseAdmission.cboUDICDCode3 = lstICDList.SelectedItem.Text
                End If
            End If
        End If
    End If
End Sub

Private Sub optRep_Click()
    Unload Me
    FrmSearchRepudiation.Source = Source
    FrmSearchRepudiation.Show vbModal
End Sub
