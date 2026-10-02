VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmPVSunHISFlag 
   Caption         =   "SunHISPVChecking"
   ClientHeight    =   6240
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   9180
   FillColor       =   &H00C0C0C0&
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   6240
   ScaleWidth      =   9180
   WindowState     =   2  'Maximized
   Begin VB.Frame Frame2 
      Height          =   550
      Left            =   120
      TabIndex        =   0
      Top             =   7125
      Width           =   11775
      Begin VB.CommandButton CmdRecon 
         Caption         =   "Recon"
         Height          =   330
         Left            =   9720
         TabIndex        =   5
         Top             =   180
         Width           =   735
      End
      Begin VB.CommandButton CmdExit 
         Cancel          =   -1  'True
         Caption         =   "E&xit"
         Height          =   330
         Left            =   11040
         TabIndex        =   4
         Top             =   180
         Width           =   615
      End
      Begin VB.CommandButton CmdClear 
         Caption         =   "&Clear"
         Height          =   330
         Left            =   10440
         TabIndex        =   3
         Top             =   180
         Width           =   615
      End
      Begin VB.CommandButton cmdPrint 
         Caption         =   "&Print"
         Height          =   330
         Left            =   9120
         TabIndex        =   2
         Top             =   180
         Width           =   615
      End
      Begin VB.CommandButton CmdSearch 
         Caption         =   "&Search"
         Height          =   330
         Left            =   8400
         TabIndex        =   1
         Top             =   180
         Width           =   735
      End
      Begin VB.Label lblTotal 
         Alignment       =   1  'Right Justify
         BackColor       =   &H80000009&
         Height          =   255
         Left            =   1080
         TabIndex        =   9
         Top             =   210
         Width           =   645
      End
      Begin VB.Label Label11 
         Caption         =   "Records"
         Height          =   255
         Left            =   1800
         TabIndex        =   8
         Top             =   240
         Width           =   735
      End
      Begin VB.Label lblCurrent 
         Alignment       =   1  'Right Justify
         BackColor       =   &H80000009&
         Height          =   255
         Left            =   120
         TabIndex        =   7
         Top             =   210
         Width           =   645
      End
      Begin VB.Label Label9 
         Caption         =   "of"
         Height          =   255
         Left            =   840
         TabIndex        =   6
         Top             =   240
         Width           =   375
      End
   End
   Begin MSComctlLib.ListView lstSunPVCheck 
      Height          =   5820
      Left            =   120
      TabIndex        =   22
      Top             =   1320
      Width           =   11775
      _ExtentX        =   20770
      _ExtentY        =   10266
      View            =   3
      LabelEdit       =   1
      Sorted          =   -1  'True
      MultiSelect     =   -1  'True
      LabelWrap       =   0   'False
      HideSelection   =   0   'False
      AllowReorder    =   -1  'True
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   14
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Sun PV No"
         Object.Width           =   1764
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   1
         Text            =   "Sun PV Date"
         Object.Width           =   2117
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "Sun PV Period"
         Object.Width           =   2646
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Text            =   "Sun Amt"
         Object.Width           =   1764
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   4
         Text            =   "Sun Hosp"
         Object.Width           =   1941
      EndProperty
      BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   5
         Text            =   "HIS PV No"
         Object.Width           =   1764
      EndProperty
      BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   6
         Text            =   "HIS PV Date"
         Object.Width           =   2469
      EndProperty
      BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   7
         Text            =   "HIS Amt"
         Object.Width           =   1764
      EndProperty
      BeginProperty ColumnHeader(9) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   8
         Text            =   "HIS Hosp"
         Object.Width           =   1764
      EndProperty
      BeginProperty ColumnHeader(10) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   9
         Text            =   "Amt Diff"
         Object.Width           =   1764
      EndProperty
      BeginProperty ColumnHeader(11) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   10
         Text            =   "Sun PV Period Diff"
         Object.Width           =   3528
      EndProperty
      BeginProperty ColumnHeader(12) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   11
         Text            =   "Hosp Diff"
         Object.Width           =   1764
      EndProperty
      BeginProperty ColumnHeader(13) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   12
         Text            =   "Date Diff"
         Object.Width           =   1764
      EndProperty
      BeginProperty ColumnHeader(14) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   13
         Text            =   "Recon Status"
         Object.Width           =   1764
      EndProperty
   End
   Begin VB.Frame Frame1 
      Height          =   1005
      Left            =   120
      TabIndex        =   10
      Top             =   240
      Width           =   11775
      Begin VB.TextBox txtHISAmtTo 
         Height          =   285
         Left            =   7200
         MaxLength       =   8
         TabIndex        =   26
         Top             =   600
         Width           =   1335
      End
      Begin VB.CheckBox chkPeriodMth 
         Height          =   255
         Left            =   11280
         TabIndex        =   17
         Top             =   1275
         Width           =   375
      End
      Begin VB.CheckBox chkPVDate 
         Height          =   255
         Left            =   11280
         TabIndex        =   16
         Top             =   1020
         Width           =   375
      End
      Begin VB.CheckBox ChkPeriodYr 
         Height          =   255
         Left            =   11280
         TabIndex        =   15
         Top             =   1590
         Width           =   375
      End
      Begin VB.TextBox txtSunAmtTo 
         Height          =   285
         Left            =   2760
         MaxLength       =   8
         TabIndex        =   13
         Top             =   600
         Width           =   1335
      End
      Begin VB.TextBox txtSunPVNo 
         Height          =   285
         Left            =   1080
         MaxLength       =   15
         TabIndex        =   12
         Top             =   360
         Width           =   1215
      End
      Begin VB.TextBox txtHISPVNo 
         Height          =   285
         Left            =   5400
         MaxLength       =   15
         TabIndex        =   11
         Top             =   360
         Width           =   1335
      End
      Begin VB.TextBox txtSunAmtFr 
         Height          =   285
         Left            =   1080
         MaxLength       =   8
         TabIndex        =   14
         Top             =   600
         Width           =   1215
      End
      Begin VB.TextBox txtHISAmtFr 
         Height          =   285
         Left            =   5400
         MaxLength       =   8
         TabIndex        =   25
         Top             =   600
         Width           =   1335
      End
      Begin VB.Label Label4 
         Caption         =   "HIS"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   255
         Left            =   4560
         TabIndex        =   30
         Top             =   120
         Width           =   855
      End
      Begin VB.Label Label3 
         Caption         =   "SUN"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   255
         Left            =   120
         TabIndex        =   29
         Top             =   120
         Width           =   855
      End
      Begin VB.Label Label2 
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
         Left            =   6840
         TabIndex        =   28
         Top             =   600
         Width           =   375
      End
      Begin VB.Label Label1 
         Caption         =   "Inv Amt"
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
         Left            =   4560
         TabIndex        =   27
         Top             =   600
         Width           =   855
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
         Left            =   120
         TabIndex        =   20
         Top             =   360
         Width           =   855
      End
      Begin VB.Label Label8 
         Caption         =   "Inv Amt"
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
         TabIndex        =   18
         Top             =   600
         Width           =   855
      End
      Begin VB.Label Label16 
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
         Left            =   4560
         TabIndex        =   21
         Top             =   360
         Width           =   735
      End
      Begin VB.Label Label14 
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
         Left            =   2400
         TabIndex        =   19
         Top             =   600
         Width           =   375
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
      TabIndex        =   24
      Top             =   7680
      Width           =   1575
   End
   Begin VB.Label Label28 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00404040&
      Caption         =   "SunHIS PV Checking"
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
      TabIndex        =   23
      Top             =   0
      Width           =   11805
   End
End
Attribute VB_Name = "frmPVSunHISFlag"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private m_blnDirection(1 To 15) As Boolean

Private Sub CmdExit_Click()
    Unload Me
End Sub

Private Sub cmdPrint_Click()
    Screen.MousePointer = vbHourglass
        Call ExportToExcelAnalysis1(lstSunPVCheck)
    Screen.MousePointer = vbDefault
End Sub

Private Sub CmdRecon_Click()
Dim ListCount, ForLoop As Integer
Dim strsqlHIS, strsqlSUN As String
Dim HISPV As New ADODB.Recordset
Dim SunPV As New ADODB.Recordset

    ListCount = lstSunPVCheck.listitems.count
    
    For ForLoop = 1 To ListCount
        With lstSunPVCheck.listitems(ForLoop)
            If Trim(.SubItems(9)) = "No" And Trim(.SubItems(10)) = "No" And Trim(.SubItems(11)) = "No" And Trim(.SubItems(12)) = "No" Then
                strsqlSUN = "SELECT PVNo,Flag,Period From PVSUN2000 where PVNo = '" & Trim(lstSunPVCheck.listitems(ForLoop)) & "' And Period = '" & Trim(.SubItems(3)) & "'"
                SunPV.Open strsqlSUN, medixmasterconnPayment, adOpenKeyset, adLockOptimistic
                
                If SunPV.recordCount Then
                    SunPV!Flag = "Y"
                    SunPV.Update
                End If
                
                strsqlHIS = "SELECT PVNo,Flag From PVHISAll where PVNo = '" & Trim(lstSunPVCheck.listitems(ForLoop)) & "'"
                HISPV.Open strsqlHIS, medixmasterconnPayment, adOpenKeyset, adLockOptimistic
                
                If HISPV.recordCount Then
                        HISPV!Flag = "Y"
                        HISPV.Update
                End If
                SunPV.Close
                Set SunPV = Nothing
                HISPV.Close
                Set HISPV = Nothing
            End If
        End With
    Next

End Sub

Private Sub CmdSearch_Click()
    Screen.MousePointer = vbHourglass
        Call SearchRecord
    Screen.MousePointer = Default
End Sub

Private Function SearchRecord()
Dim strsql As String
Dim strAppend As String
Dim Sql As String
    
    strAppend = ""
    If Len(txtSunPVNo) Then
        strAppend = strAppend + " and ClaimVersion = '" & Trim(txtSunPVNo) & "'"
    End If
    
    If Len(txtSunAmtFr) Then
        strAppend = strAppend + " and ClaimVersion >= '" & Trim(txtSunAmtFr) & "'"
    End If
    
    If Len(txtSunAmtTo) Then
        strAppend = strAppend + " and INVAmount <= " & Trim(txtSunAmtTo) & ""
    End If
    
    If Len(txtHISPVNo) Then
        strAppend = strAppend + " and PVHOSNo = '" & Trim(txtHISPVNo) & "'"
    End If
        
    If Len(txtHISAmtFr) Then
        strAppend = strAppend + " and ClaimVersion >= '" & Trim(txtHISAmtFr) & "'"
    End If
    
    If Len(txtHISAmtTo) Then
        strAppend = strAppend + " and INVAmount <= " & Trim(txtHISAmtTo) & ""
    End If
    
    Call PVListing(strAppend)

End Function

Private Sub PVListing(Optional strAppend As String)
Dim rst2 As New ADODB.Recordset
Dim Record As ListItem
Dim Sql As String
Dim total As String
Dim Current As String
Dim TotalAmt As Variant
Dim Check As Boolean
Dim strsql As String
Dim HISPV As New ADODB.Recordset
Dim AmtDiff As Boolean
Dim PVPerDiff As Boolean
Dim HOSDiff As Boolean
Dim DateDiff As Boolean

    total = 0
    Current = 0
    lstSunPVCheck.listitems.Clear
    Sql = "SELECT Period,Date,PVNo,Amount,HospCode From VIEW_PNSunAll where 0 = 0 "
    If Len(Trim(strAppend)) Then
        Sql = Sql + strAppend
    End If
    
    rst2.Open Sql, medixmasterconnPayment, adOpenKeyset, adLockOptimistic
    
    If rst2.recordCount = 0 Then
        MsgBox "No record found!", vbOKOnly + vbCritical, DialogBoxTitle
    Else
    
        Do Until rst2.EOF
        'total = rst2.recordCount
        'lblTotal = total
        Sql = "SELECT PVNo,Amt,Date,HospAccCode From PVHISAll where PVNo = '" & Trim(rst2!PVNo) & "'"
        HISPV.Open Sql, medixmasterconnPayment, adOpenKeyset, adLockOptimistic
            If HISPV.recordCount Then
                If HISPV!Amt <> rst2!Amount Then
                    AmtDiff = False
                Else
                    AmtDiff = True
                End If
                If rst2!Period <> Format(rst2!Date, "mm/yyyy") Then
                    PVPerDiff = False
                Else
                    PVPerDiff = True
                End If
                
                If HISPV!HospAccCode <> rst2!HospCode Then
                    HOSDiff = False
                Else
                    HOSDiff = True
                End If
                
                If HISPV!Date <> rst2!Date Then
                    DateDiff = False
                Else
                    DateDiff = True
                End If
                
                Set Record = lstSunPVCheck.listitems.Add(, , rst2!PVNo)
                
                Record.SubItems(1) = IIf(Len(rst2!Date), rst2!Date, "")
                Record.SubItems(2) = IIf(Len(rst2!Period), rst2!Period, "")
                Record.SubItems(3) = Format(IIf(Len(rst2!Amount), rst2!Amount, ""), "#,##0.00")
                Record.SubItems(4) = IIf(Len(rst2!HospCode), rst2!HospCode, "")
                Record.SubItems(5) = IIf(Len(HISPV!PVNo), HISPV!PVNo, "")
                Record.SubItems(6) = IIf(Len(HISPV!Date), HISPV!Date, "")
                Record.SubItems(7) = Format(IIf(Len(HISPV!Amt), HISPV!Amt, ""), "#,##0.00")
                Record.SubItems(8) = IIf(Len(HISPV!HospAccCode), HISPV!HospAccCode, "")
                Record.SubItems(9) = IIf(AmtDiff = True, "No", "Yes")
                Record.SubItems(10) = IIf(PVPerDiff = True, "No", "Yes")
                Record.SubItems(11) = IIf(HOSDiff = True, "No", "Yes")
                Record.SubItems(12) = IIf(DateDiff = True, "No", "Yes")

                Current = Current + 1
                lblCurrent = Current
            End If
            HISPV.Close
            Set HISPV = Nothing
            
            
            rst2.MoveNext
        Loop
    End If
    rst2.Close
    Set rst2 = Nothing
End Sub

Private Sub lstSunPVCheck_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
    m_blnDirection(ColumnHeader.Index) = Not m_blnDirection(ColumnHeader.Index)
    Select Case ColumnHeader.Index
    Case 1
        SortListView lstSunPVCheck, ColumnHeader.Index, ldtString, m_blnDirection(1)
    Case 2
        SortListView lstSunPVCheck, ColumnHeader.Index, ldtDateTime, m_blnDirection(2)
    Case 3
        SortListView lstSunPVCheck, ColumnHeader.Index, ldtDateTime, m_blnDirection(3)
    Case 4
        SortListView lstSunPVCheck, ColumnHeader.Index, ldtNumber, m_blnDirection(4)
    Case 5
        SortListView lstSunPVCheck, ColumnHeader.Index, ldtString, m_blnDirection(5)
    Case 6
        SortListView lstSunPVCheck, ColumnHeader.Index, ldtString, m_blnDirection(6)
    Case 7
        SortListView lstSunPVCheck, ColumnHeader.Index, ldtDateTime, m_blnDirection(7)
    Case 8
        SortListView lstSunPVCheck, ColumnHeader.Index, ldtNumber, m_blnDirection(8)
    Case 9
        SortListView lstSunPVCheck, ColumnHeader.Index, ldtString, m_blnDirection(9)
    Case 10
        SortListView lstSunPVCheck, ColumnHeader.Index, ldtString, m_blnDirection(10)
    Case 11
        SortListView lstSunPVCheck, ColumnHeader.Index, ldtString, m_blnDirection(11)
    Case 12
        SortListView lstSunPVCheck, ColumnHeader.Index, ldtString, m_blnDirection(12)
    Case 13
        SortListView lstSunPVCheck, ColumnHeader.Index, ldtString, m_blnDirection(13)
    Case 14
        SortListView lstSunPVCheck, ColumnHeader.Index, ldtString, m_blnDirection(14)
    End Select
End Sub
