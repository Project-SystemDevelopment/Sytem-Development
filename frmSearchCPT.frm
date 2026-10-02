VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Begin VB.Form frmSearchCPT 
   Caption         =   "Search Procedure"
   ClientHeight    =   4935
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   7410
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4935
   ScaleWidth      =   7410
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      Height          =   615
      Left            =   120
      TabIndex        =   10
      Top             =   4320
      Width           =   7215
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
         Left            =   3720
         TabIndex        =   15
         Top             =   240
         Width           =   840
      End
      Begin VB.CommandButton CmdClear 
         Caption         =   "&Clear"
         Height          =   315
         Left            =   5400
         TabIndex        =   14
         Top             =   240
         Width           =   840
      End
      Begin VB.CommandButton cmdExit 
         Cancel          =   -1  'True
         Caption         =   "Exit"
         Height          =   315
         Left            =   6240
         TabIndex        =   13
         Top             =   240
         Width           =   840
      End
      Begin VB.CommandButton cmdSearch 
         Caption         =   "&Search"
         Default         =   -1  'True
         Height          =   315
         Left            =   4560
         TabIndex        =   12
         Top             =   240
         Width           =   840
      End
      Begin VB.CommandButton CmdCPTInfo 
         Caption         =   "CPT &Info"
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
         Left            =   4560
         TabIndex        =   11
         Top             =   240
         Width           =   840
      End
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
      Height          =   675
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   7365
      Begin VB.TextBox txtMMACode 
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
         Left            =   960
         MaxLength       =   10
         TabIndex        =   7
         Top             =   240
         Width           =   1695
      End
      Begin VB.TextBox txtMMADesc 
         Height          =   285
         Left            =   3600
         TabIndex        =   6
         Top             =   240
         Width           =   3615
      End
      Begin VB.TextBox txtCPTDesc 
         Height          =   285
         Left            =   3600
         TabIndex        =   2
         Top             =   240
         Visible         =   0   'False
         Width           =   3615
      End
      Begin VB.TextBox TxtCPTCode 
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
         Left            =   960
         MaxLength       =   10
         TabIndex        =   1
         Top             =   240
         Visible         =   0   'False
         Width           =   1695
      End
      Begin VB.Label Label2 
         Caption         =   "MMA Code"
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
      Begin VB.Label Label1 
         Caption         =   "MMA Desc"
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
         Left            =   2760
         TabIndex        =   8
         Top             =   240
         Width           =   1275
      End
      Begin VB.Label Label33 
         Caption         =   "CPT Desc"
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
         Left            =   2760
         TabIndex        =   4
         Top             =   240
         Visible         =   0   'False
         Width           =   1275
      End
      Begin VB.Label Label32 
         Caption         =   "CPT Code"
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
         TabIndex        =   3
         Top             =   240
         Visible         =   0   'False
         Width           =   1095
      End
   End
   Begin MSComctlLib.ListView lstCPTList 
      Height          =   3555
      Left            =   0
      TabIndex        =   5
      Top             =   720
      Width           =   7380
      _ExtentX        =   13018
      _ExtentY        =   6271
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
         Text            =   "MMA Code"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "MMA Description"
         Object.Width           =   8996
      EndProperty
   End
End
Attribute VB_Name = "frmSearchCPT"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public Source As Integer
Private m_blnDirection(1 To 4) As Boolean

Private Sub CmdCPTInfo_Click()
    FrmShowCPT.Source = 1
    FrmShowCPT.Show vbModal
End Sub

Private Sub CmdExit_Click()
    Unload Me
End Sub

Private Sub cmdPrint_Click()
    Screen.MousePointer = vbHourglass
        Call ExportToExcelAnalysis1(lstCPTList)
    Screen.MousePointer = vbDefault
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

Private Sub cmdClear_Click()
TxtCPTCode.Text = ""
txtCPTDesc.Text = ""
lstCPTList.ListItems.Clear

End Sub

Private Sub CPTListing(Optional strAppend As String)
    Dim rst2 As New ADODB.Recordset
    Dim CPTMaster As ListItem
    Dim sql As String
     
    lstCPTList.ListItems.Clear
        sql = " SELECT OPCSCode,OPCSNarrative " & _
        " from MMAProcedures where 0=0 "
    
    If Len(Trim(strAppend)) Then
        sql = sql + strAppend
    End If

    rst2.Open sql, medixmasterconnMaint, adOpenKeyset, adLockOptimistic
    If rst2.recordCount = 0 Then
        MsgBox "Record not found", vbInformation + vbOKOnly, DialogBoxTitle
    Else
        Do Until rst2.EOF
            With rst2
                Set CPTMaster = lstCPTList.ListItems.Add(, , IIf(Len(!OPCSCode), !OPCSCode, ""))
                If Len(rst2!OPCSNarrative) Then
                    CPTMaster.SubItems(1) = !OPCSNarrative
                End If
                rst2.MoveNext
            End With
        Loop
    rst2.Close
    Set rst2 = Nothing
    End If
End Sub

Private Sub CmdSearch_Click()
    Dim strCriteria As String
    
    Screen.MousePointer = vbHourglass
    cmdSearch.Enabled = False
    'medixmasterconnMaint.Open "File Name=" & BOSPath & "Conn\MediConnHISMaint.UDL"
    strCriteria = ""
    
    
    If Len(Trim(txtMMACode)) Then
            strCriteria = " AND OPCSCode like '" & Trim(txtMMACode) & "%'"
    End If
    If Len(Trim(txtMMADesc)) Then
            strCriteria = strCriteria & " AND OPCSNarrative like '%" & Trim(txtMMADesc) & "%' "
    End If
    
    Call CPTListing(strCriteria)
    'medixmasterconnMaint.Close
    'Set medixmasterconnMaint = Nothing
    cmdSearch.Enabled = True
    Screen.MousePointer = vbDefault
    
End Sub

'**********************Start Sorted ListView *******************************
Private Sub lstCPTList_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
    m_blnDirection(ColumnHeader.Index) = Not m_blnDirection(ColumnHeader.Index)
    Select Case ColumnHeader.Index
    Case 1
        SortListView lstCPTList, ColumnHeader.Index, ldtString, m_blnDirection(1)
    Case 2
        SortListView lstCPTList, ColumnHeader.Index, ldtString, m_blnDirection(2)
    Case 3
        SortListView lstCPTList, ColumnHeader.Index, ldtString, m_blnDirection(3)
    Case 4
        SortListView lstCPTList, ColumnHeader.Index, ldtString, m_blnDirection(4)
    End Select
End Sub
'**********************End Sorted ListView *******************************

Private Sub lstCPTList_DblClick()
    frmMMAEntryGST.CalSource = 0
    If lstCPTList.ListItems.Count <> 0 Then
        If Source = 1 Then
            
            If Trim(Len(frmMMAEntryGST.txtProcCode1)) = 0 Then
                If Len(lstCPTList.SelectedItem.Text) Then
                    frmMMAEntryGST.txtProcCode1 = lstCPTList.SelectedItem.Text
                Else
                    frmMMAEntryGST.txtProcCode1 = "000000"
                End If
            ElseIf Trim(Len(frmMMAEntryGST.txtProcCode2)) = 0 Then
                If Len(lstCPTList.SelectedItem.Text) Then
                    frmMMAEntryGST.txtProcCode2 = lstCPTList.SelectedItem.Text
                Else
                    frmMMAEntryGST.txtProcCode2 = "000000"
                End If
            ElseIf Trim(Len(frmMMAEntryGST.txtProcCode3)) = 0 Then
                If Len(lstCPTList.SelectedItem.Text) Then
                    frmMMAEntryGST.txtProcCode3 = lstCPTList.SelectedItem.Text
                Else
                    frmMMAEntryGST.txtProcCode3 = "000000"
                End If
            ElseIf Trim(Len(frmMMAEntryGST.txtProcCode4)) = 0 Then
                If Len(lstCPTList.SelectedItem.Text) Then
                    frmMMAEntryGST.txtProcCode4 = lstCPTList.SelectedItem.Text
                Else
                    frmMMAEntryGST.txtProcCode4 = "000000"
                End If
            ElseIf Trim(Len(frmMMAEntryGST.txtProcCode5)) = 0 Then
                If Len(lstCPTList.SelectedItem.Text) Then
                    frmMMAEntryGST.txtProcCode5 = lstCPTList.SelectedItem.Text
                Else
                    frmMMAEntryGST.txtProcCode5 = "000000"
                End If
            ElseIf Trim(Len(frmMMAEntryGST.txtProcCode6)) = 0 Then
                If Len(lstCPTList.SelectedItem.Text) Then
                    frmMMAEntryGST.txtProcCode6 = lstCPTList.SelectedItem.Text
                Else
                    frmMMAEntryGST.txtProcCode6 = "000000"
                End If
                
            Else
                If Len(lstCPTList.SelectedItem.Text) Then
                    frmMMAEntryGST.txtProcCode6 = lstCPTList.SelectedItem.Text
                Else
                    frmMMAEntryGST.txtProcCode6 = "000000"
                End If
            End If
        
        ElseIf Source = 0 Then
            If Trim(Len(frmMMAEntryGST.txtNSProcCode1)) = 0 Then
                If Len(lstCPTList.SelectedItem.Text) Then
                    frmMMAEntryGST.txtNSProcCode1 = lstCPTList.SelectedItem.Text
                Else
                    frmMMAEntryGST.txtNSProcCode1 = "000000"
                End If
            ElseIf Trim(Len(frmMMAEntryGST.txtNSProcCode2)) = 0 Then
                If Len(lstCPTList.SelectedItem.Text) Then
                    frmMMAEntryGST.txtNSProcCode2 = lstCPTList.SelectedItem.Text
                Else
                    frmMMAEntryGST.txtNSProcCode2 = "000000"
                End If
            ElseIf Trim(Len(frmMMAEntryGST.txtNSProcCode3)) = 0 Then
                If Len(lstCPTList.SelectedItem.Text) Then
                    frmMMAEntryGST.txtNSProcCode3 = lstCPTList.SelectedItem.Text
                Else
                    frmMMAEntryGST.txtNSProcCode3 = "000000"
                End If
            ElseIf Trim(Len(frmMMAEntryGST.txtNSProcCode4)) = 0 Then
                If Len(lstCPTList.SelectedItem.Text) Then
                    frmMMAEntryGST.txtNSProcCode4 = lstCPTList.SelectedItem.Text
                Else
                    frmMMAEntryGST.txtNSProcCode4 = "000000"
                End If
            ElseIf Trim(Len(frmMMAEntryGST.txtNSProcCode5)) = 0 Then
                If Len(lstCPTList.SelectedItem.Text) Then
                    frmMMAEntryGST.txtNSProcCode5 = lstCPTList.SelectedItem.Text
                Else
                    frmMMAEntryGST.txtNSProcCode5 = "000000"
                End If
            Else
                If Len(lstCPTList.SelectedItem.Text) Then
                    frmMMAEntryGST.txtNSProcCode6 = lstCPTList.SelectedItem.Text
                Else
                    frmMMAEntryGST.txtNSProcCode6 = "000000"
                End If
            End If
        ElseIf Source = 2 Then
            
            If Trim(Len(frmCaseAdjustment2016.txtProcCode1(0))) = 0 Then
                If Len(lstCPTList.SelectedItem.Text) Then
                    frmCaseAdjustment2016.txtProcCode1(0) = lstCPTList.SelectedItem.Text
                Else
                    frmCaseAdjustment2016.txtProcCode1(0) = "000000"
                End If
            ElseIf Trim(Len(frmCaseAdjustment2016.txtProcCode1(1))) = 0 Then
                If Len(lstCPTList.SelectedItem.Text) Then
                    frmCaseAdjustment2016.txtProcCode1(1) = lstCPTList.SelectedItem.Text
                Else
                    frmCaseAdjustment2016.txtProcCode1(1) = "000000"
                End If
            ElseIf Trim(Len(frmCaseAdjustment2016.txtProcCode1(2))) = 0 Then
                If Len(lstCPTList.SelectedItem.Text) Then
                    frmCaseAdjustment2016.txtProcCode1(2) = lstCPTList.SelectedItem.Text
                Else
                    frmCaseAdjustment2016.txtProcCode1(2) = "000000"
                End If
            ElseIf Trim(Len(frmCaseAdjustment2016.txtProcCode1(3))) = 0 Then
                If Len(lstCPTList.SelectedItem.Text) Then
                    frmCaseAdjustment2016.txtProcCode1(3) = lstCPTList.SelectedItem.Text
                Else
                    frmCaseAdjustment2016.txtProcCode1(3) = "000000"
                End If
            ElseIf Trim(Len(frmCaseAdjustment2016.txtProcCode1(4))) = 0 Then
                If Len(lstCPTList.SelectedItem.Text) Then
                    frmCaseAdjustment2016.txtProcCode1(4) = lstCPTList.SelectedItem.Text
                Else
                    frmCaseAdjustment2016.txtProcCode1(4) = "000000"
                End If
            ElseIf Trim(Len(frmCaseAdjustment2016.txtProcCode1(5))) = 0 Then
                If Len(lstCPTList.SelectedItem.Text) Then
                    frmCaseAdjustment2016.txtProcCode1(5) = lstCPTList.SelectedItem.Text
                Else
                    frmCaseAdjustment2016.txtProcCode1(5) = "000000"
                End If
            ElseIf Trim(Len(frmCaseAdjustment2016.txtProcCode1(6))) = 0 Then
                If Len(lstCPTList.SelectedItem.Text) Then
                    frmCaseAdjustment2016.txtProcCode1(6) = lstCPTList.SelectedItem.Text
                Else
                    frmCaseAdjustment2016.txtProcCode1(6) = "000000"
                End If
            Else
                If Len(lstCPTList.SelectedItem.Text) Then
                    frmCaseAdjustment2016.txtProcCode1(7) = lstCPTList.SelectedItem.Text
                Else
                    frmCaseAdjustment2016.txtProcCode1(7) = "000000"
                End If
            End If
            
        ElseIf Source = 3 Then
            If Trim(Len(frmCaseAdjustment2016.txtNSProcCode1(0))) = 0 Then
                If Len(lstCPTList.SelectedItem.Text) Then
                    frmCaseAdjustment2016.txtNSProcCode1(0) = lstCPTList.SelectedItem.Text
                Else
                    frmCaseAdjustment2016.txtNSProcCode1(0) = "000000"
                End If
            ElseIf Trim(Len(frmCaseAdjustment2016.txtNSProcCode1(1))) = 0 Then
                If Len(lstCPTList.SelectedItem.Text) Then
                    frmCaseAdjustment2016.txtNSProcCode1(1) = lstCPTList.SelectedItem.Text
                Else
                    frmCaseAdjustment2016.txtNSProcCode1(1) = "000000"
                End If
            ElseIf Trim(Len(frmCaseAdjustment2016.txtNSProcCode1(2))) = 0 Then
                If Len(lstCPTList.SelectedItem.Text) Then
                    frmCaseAdjustment2016.txtNSProcCode1(2) = lstCPTList.SelectedItem.Text
                Else
                    frmCaseAdjustment2016.txtNSProcCode1(2) = "000000"
                End If
            ElseIf Trim(Len(frmCaseAdjustment2016.txtNSProcCode1(3))) = 0 Then
                If Len(lstCPTList.SelectedItem.Text) Then
                    frmCaseAdjustment2016.txtNSProcCode1(3) = lstCPTList.SelectedItem.Text
                Else
                    frmCaseAdjustment2016.txtNSProcCode1(3) = "000000"
                End If
            ElseIf Trim(Len(frmCaseAdjustment2016.txtNSProcCode1(4))) = 0 Then
                If Len(lstCPTList.SelectedItem.Text) Then
                    frmCaseAdjustment2016.txtNSProcCode1(4) = lstCPTList.SelectedItem.Text
                Else
                    frmCaseAdjustment2016.txtNSProcCode1(4) = "000000"
                End If
            ElseIf Trim(Len(frmCaseAdjustment2016.txtNSProcCode1(5))) = 0 Then
                If Len(lstCPTList.SelectedItem.Text) Then
                    frmCaseAdjustment2016.txtNSProcCode1(5) = lstCPTList.SelectedItem.Text
                Else
                    frmCaseAdjustment2016.txtNSProcCode1(5) = "000000"
                End If
            ElseIf Trim(Len(frmCaseAdjustment2016.txtNSProcCode1(6))) = 0 Then
                If Len(lstCPTList.SelectedItem.Text) Then
                    frmCaseAdjustment2016.txtNSProcCode1(6) = lstCPTList.SelectedItem.Text
                Else
                    frmCaseAdjustment2016.txtNSProcCode1(6) = "000000"
                End If
            Else
                If Len(lstCPTList.SelectedItem.Text) Then
                    frmCaseAdjustment2016.txtNSProcCode1(7) = lstCPTList.SelectedItem.Text
                Else
                    frmCaseAdjustment2016.txtNSProcCode1(7) = "000000"
                End If
            End If
        
        ElseIf Source = 4 Then
            
            If Trim(Len(frmCaseAdmission.txtProcCode1)) = 0 Then
                If Len(lstCPTList.SelectedItem.Text) Then
                    frmCaseAdmission.txtProcCode1 = lstCPTList.SelectedItem.Text
                Else
                    frmCaseAdmission.txtProcCode1 = "000000"
                End If
            ElseIf Trim(Len(frmCaseAdmission.txtProcCode2)) = 0 Then
                If Len(lstCPTList.SelectedItem.Text) Then
                    frmCaseAdmission.txtProcCode2 = lstCPTList.SelectedItem.Text
                Else
                    frmCaseAdmission.txtProcCode2 = "000000"
                End If
            ElseIf Trim(Len(frmCaseAdmission.txtProcCode3)) = 0 Then
                If Len(lstCPTList.SelectedItem.Text) Then
                    frmCaseAdmission.txtProcCode3 = lstCPTList.SelectedItem.Text
                Else
                    frmCaseAdmission.txtProcCode3 = "000000"
                End If
            ElseIf Trim(Len(frmCaseAdmission.txtProcCode4)) = 0 Then
                If Len(lstCPTList.SelectedItem.Text) Then
                    frmCaseAdmission.txtProcCode4 = lstCPTList.SelectedItem.Text
                Else
                    frmCaseAdmission.txtProcCode4 = "000000"
                End If
            ElseIf Trim(Len(frmCaseAdmission.txtProcCode5)) = 0 Then
                If Len(lstCPTList.SelectedItem.Text) Then
                    frmCaseAdmission.txtProcCode5 = lstCPTList.SelectedItem.Text
                Else
                    frmCaseAdmission.txtProcCode5 = "000000"
                End If
            Else
                If Len(lstCPTList.SelectedItem.Text) Then
                    frmCaseAdmission.txtProcCode6 = lstCPTList.SelectedItem.Text
                Else
                    frmCaseAdmission.txtProcCode6 = "000000"
                End If
                
            End If
        
        ElseIf Source = 5 Then
            If Trim(Len(frmCaseAdmission.txtNSProcCode1)) = 0 Then
                If Len(lstCPTList.SelectedItem.Text) Then
                    frmCaseAdmission.txtNSProcCode1 = lstCPTList.SelectedItem.Text
                Else
                    frmCaseAdmission.txtNSProcCode1 = "000000"
                End If
            ElseIf Trim(Len(frmCaseAdmission.txtNSProcCode2)) = 0 Then
                If Len(lstCPTList.SelectedItem.Text) Then
                    frmCaseAdmission.txtNSProcCode2 = lstCPTList.SelectedItem.Text
                Else
                    frmCaseAdmission.txtNSProcCode2 = "000000"
                End If
            ElseIf Trim(Len(frmCaseAdmission.txtNSProcCode3)) = 0 Then
                If Len(lstCPTList.SelectedItem.Text) Then
                    frmCaseAdmission.txtNSProcCode3 = lstCPTList.SelectedItem.Text
                Else
                    frmCaseAdmission.txtNSProcCode3 = "000000"
                End If
            ElseIf Trim(Len(frmCaseAdmission.txtNSProcCode4)) = 0 Then
                If Len(lstCPTList.SelectedItem.Text) Then
                    frmCaseAdmission.txtNSProcCode4 = lstCPTList.SelectedItem.Text
                Else
                    frmCaseAdmission.txtNSProcCode4 = "000000"
                End If
            ElseIf Trim(Len(frmCaseAdmission.txtNSProcCode5)) = 0 Then
                If Len(lstCPTList.SelectedItem.Text) Then
                    frmCaseAdmission.txtNSProcCode5 = lstCPTList.SelectedItem.Text
                Else
                    frmCaseAdmission.txtNSProcCode5 = "000000"
                End If
            Else
                If Len(lstCPTList.SelectedItem.Text) Then
                    frmCaseAdmission.txtNSProcCode6 = lstCPTList.SelectedItem.Text
                Else
                    frmCaseAdmission.txtNSProcCode6 = "000000"
                End If
            End If
        End If
    End If
End Sub

