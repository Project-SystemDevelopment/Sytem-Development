VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "mscomctl.ocx"
Begin VB.Form frmSearchInv 
   Caption         =   "Search Investigation"
   ClientHeight    =   4920
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   6090
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4920
   ScaleWidth      =   6090
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdPrint 
      Caption         =   "&Print"
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
      Left            =   2160
      TabIndex        =   4
      Top             =   4560
      Width           =   840
   End
   Begin VB.CommandButton CmdInvInfo 
      Caption         =   "Inv &Info"
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
      Left            =   1200
      TabIndex        =   3
      Top             =   4560
      Width           =   840
   End
   Begin VB.CommandButton CmdClear 
      Caption         =   "&Clear"
      Height          =   315
      Left            =   4080
      TabIndex        =   6
      Top             =   4560
      Width           =   840
   End
   Begin VB.CommandButton cmdExit 
      Cancel          =   -1  'True
      Caption         =   "Exit"
      Height          =   315
      Left            =   5040
      TabIndex        =   7
      Top             =   4560
      Width           =   840
   End
   Begin VB.CommandButton cmdSearch 
      Caption         =   "&Search"
      Default         =   -1  'True
      Height          =   315
      Left            =   3120
      TabIndex        =   5
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
      Height          =   1035
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   6045
      Begin VB.TextBox TxtINVCode 
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
         MaxLength       =   10
         TabIndex        =   2
         Top             =   240
         Width           =   1215
      End
      Begin VB.TextBox txtINVDesc 
         Height          =   285
         Left            =   1440
         TabIndex        =   1
         Top             =   600
         Width           =   3375
      End
      Begin VB.Label Label32 
         Caption         =   "Inv Code"
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
         TabIndex        =   9
         Top             =   240
         Width           =   1095
      End
      Begin VB.Label Label33 
         Caption         =   "Inv Description"
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
         TabIndex        =   8
         Top             =   600
         Width           =   1275
      End
   End
   Begin MSComctlLib.ListView lstINVList 
      Height          =   3435
      Left            =   0
      TabIndex        =   10
      Top             =   1080
      Width           =   6060
      _ExtentX        =   10689
      _ExtentY        =   6059
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
         Text            =   "INV Code"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "INV Description"
         Object.Width           =   8996
      EndProperty
   End
End
Attribute VB_Name = "frmSearchInv"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public Source As Integer
Private m_blnDirection(1 To 2) As Boolean

Private Sub CmdExit_Click()
    Unload Me
End Sub

Private Sub CmdInvInfo_Click()
    FrmShowInv.Source = 1
    FrmShowInv.Show vbModal
End Sub

Private Sub cmdPrint_Click()
    If lstINVList.ListItems.Count = 0 Then
        MsgBox "No record found.", vbOKOnly + vbCritical, DialogBoxTitle
        Exit Sub
    Else
        Screen.MousePointer = vbHourglass
        Call ExportListViewtoExcel(lstINVList)
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


Private Sub CmdClear_Click()
TxtINVCode.Text = ""
txtINVDesc.Text = ""
lstINVList.ListItems.Clear

End Sub

Private Sub Form_Activate()
    txtINVDesc.SetFocus
End Sub

Private Sub INVListing(Optional strAppend As String)
    Dim rst2 As New ADODB.Recordset
    Dim CPTMaster As ListItem
    Dim Sql As String
     
    lstINVList.ListItems.Clear
        Sql = " SELECT OPCSCode, OPCSNarrative " & _
        " from MMAProcedures where ProcPart = 'A' "
    
    If Len(Trim(strAppend)) Then
        Sql = Sql + strAppend
    End If

    rst2.Open Sql, medixmasterconnMaint, adOpenKeyset, adLockOptimistic
    If rst2.recordCount = 0 Then
        MsgBox "Record not found", vbInformation + vbOKOnly, DialogBoxTitle
    Else
        Do Until rst2.EOF
            If Len(rst2!OPCSCode) Then
                Set CPTMaster = lstINVList.ListItems.Add(, , rst2!OPCSCode)
                    If Len(rst2!OPCSNarrative) Then
                        CPTMaster.SubItems(1) = Trim(rst2!OPCSNarrative)
                    End If
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
    'medixmasterconnMaint.Open "File Name=" & BOSPath & "Conn\MediConnHISMaint.UDL"

    strCriteria = ""
    If Len(Trim(TxtINVCode)) Then
            strCriteria = " AND OPCSCode like '" & Trim(TxtINVCode) & "%'"
    End If
    If Len(Trim(txtINVDesc)) Then
            strCriteria = strCriteria & " AND OPCSNarrative like '%" & Trim(txtINVDesc) & "%' "
    End If
    
    Call INVListing(strCriteria)
    'medixmasterconnMaint.Close
    'Set medixmasterconnMaint = Nothing
    cmdSearch.Enabled = True
    Screen.MousePointer = vbDefault
    
End Sub

'**********************Start Sorted ListView *******************************
Private Sub lstINVList_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
    m_blnDirection(ColumnHeader.Index) = Not m_blnDirection(ColumnHeader.Index)
    Select Case ColumnHeader.Index
    Case 1
        SortListView lstINVList, ColumnHeader.Index, ldtString, m_blnDirection(1)
    Case 2
        SortListView lstINVList, ColumnHeader.Index, ldtString, m_blnDirection(2)
    End Select
End Sub
'**********************End Sorted ListView *******************************

Private Sub lstINVList_DblClick()
    
    If lstINVList.ListItems.Count <> 0 Then
        If Source = 1 Then
            If Trim(Len(frmCaseAdjustment2016.txtINVCode1)) = 0 Then
                frmCaseAdjustment2016.txtINVCode1 = lstINVList.SelectedItem.Text
            ElseIf Trim(Len(frmCaseAdjustment2016.txtINVCode2)) = 0 Then
                frmCaseAdjustment2016.txtINVCode2 = lstINVList.SelectedItem.Text
            Else
                frmCaseAdjustment2016.txtINVCode3 = lstINVList.SelectedItem.Text
            End If
        ElseIf Source = 2 Then
            If Trim(Len(frmCaseAdmission.txtINVCode1)) = 0 Then
                frmCaseAdmission.txtINVCode1 = lstINVList.SelectedItem.Text
            ElseIf Trim(Len(frmCaseAdmission.txtINVCode2)) = 0 Then
                frmCaseAdmission.txtINVCode2 = lstINVList.SelectedItem.Text
            Else
                frmCaseAdmission.txtINVCode3 = lstINVList.SelectedItem.Text
            End If
        ElseIf Source = 3 Then
            If Trim(Len(frmMMAEntry.txtINVCode1)) = 0 Then
                frmMMAEntry.txtINVCode1 = lstINVList.SelectedItem.Text
            ElseIf Trim(Len(frmMMAEntry.txtINVCode2)) = 0 Then
                frmMMAEntry.txtINVCode2 = lstINVList.SelectedItem.Text
            Else
                frmMMAEntry.txtINVCode3 = lstINVList.SelectedItem.Text
            End If
            
        End If
    End If
End Sub
