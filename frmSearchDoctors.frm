VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Begin VB.Form frmSearchDoctors 
   Caption         =   "Search Doctors"
   ClientHeight    =   6090
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   8205
   LinkTopic       =   "Form1"
   ScaleHeight     =   6090
   ScaleWidth      =   8205
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame1 
      Height          =   615
      Left            =   240
      TabIndex        =   4
      Top             =   5160
      Width           =   7815
      Begin VB.CommandButton cmdSearch 
         Caption         =   "&Search"
         Default         =   -1  'True
         Height          =   315
         Left            =   4560
         TabIndex        =   8
         Top             =   240
         Width           =   840
      End
      Begin VB.CommandButton cmdExit 
         Cancel          =   -1  'True
         Caption         =   "Exit"
         Height          =   315
         Left            =   6240
         TabIndex        =   7
         Top             =   240
         Width           =   840
      End
      Begin VB.CommandButton CmdClear 
         Caption         =   "&Clear"
         Height          =   315
         Left            =   5400
         TabIndex        =   6
         Top             =   240
         Width           =   840
      End
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
         TabIndex        =   5
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
      Height          =   1035
      Left            =   240
      TabIndex        =   1
      Top             =   120
      Width           =   7845
      Begin VB.TextBox txtDocs 
         Height          =   285
         Left            =   3120
         TabIndex        =   2
         Top             =   600
         Width           =   3855
      End
      Begin VB.Label lblhospName 
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
         Left            =   3120
         TabIndex        =   10
         Top             =   240
         Width           =   3795
      End
      Begin VB.Label Label1 
         Caption         =   "Hospital Name:"
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
         Top             =   240
         Width           =   2115
      End
      Begin VB.Label lblDoctors 
         Caption         =   "Doctor Name:"
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
         TabIndex        =   3
         Top             =   600
         Width           =   2115
      End
   End
   Begin MSComctlLib.ListView lstDoctors 
      Height          =   3555
      Left            =   240
      TabIndex        =   0
      Top             =   1320
      Width           =   7860
      _ExtentX        =   13864
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
      NumItems        =   1
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Doctor Name"
         Object.Width           =   2540
      EndProperty
   End
End
Attribute VB_Name = "frmSearchDoctors"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public Hospital As String
Private m_blnDirection(1 To 4) As Boolean

Private Sub CmdClear_Click()
    txtDocs.Text = ""
    'lblhospName.Caption = ""
    lstDoctors.ListItems.Clear
End Sub

Private Sub CmdExit_Click()
    Unload Me
End Sub

Private Sub cmdPrint_Click()
    Screen.MousePointer = vbHourglass
        Call ExportToExcelAnalysis1(lstDoctors)
    Screen.MousePointer = vbDefault
End Sub

Private Sub cmdSearch_Click()
    Dim rst2 As New ADODB.Recordset
    Dim CPTMaster As ListItem
    Dim sql As String
    Dim strAppend As String
    
     
    lstDoctors.ListItems.Clear
    
    If Len(Trim(txtDocs)) Then
            strAppend = " AND doctorname like '%" & Trim(txtDocs) & "%' ORDER BY doctorname "
    End If
    
    sql = "Select * from tbl_doctorroles where hospitalname='" + lblhospName + "' "
    
    If Len(Trim(strAppend)) Then
        sql = sql + strAppend
    End If

    rst2.Open sql, medixmasterconn, adOpenKeyset, adLockOptimistic
    If rst2.recordCount = 0 Then
        MsgBox "Record not found", vbInformation + vbOKOnly, DialogBoxTitle
    Else
        Do Until rst2.EOF
            With rst2
                Set CPTMaster = lstDoctors.ListItems.Add(, , IIf(Len(!DoctorName), !DoctorName, ""))
               ' If Len(rst2!OPCSNarrative) Then
                '    CPTMaster.SubItems(1) = !OPCSNarrative
                'End If
                rst2.MoveNext
            End With
        Loop
    rst2.Close
    Set rst2 = Nothing
    End If
End Sub

Private Sub Form_Load()
    lblhospName.Caption = Hospital
End Sub

Private Sub lstDoctors_DblClick()
    If Trim(Len(frmCaseAdjustment2016.txtATTDoc(0))) = 0 Then
        If Len(lstDoctors.SelectedItem.Text) Then
            frmCaseAdjustment2016.txtATTDoc(0) = lstDoctors.SelectedItem.Text
        End If
    ElseIf Trim(Len(frmCaseAdjustment2016.txtATTDoc(1))) = 0 Then
        If Len(lstDoctors.SelectedItem.Text) Then
            frmCaseAdjustment2016.txtATTDoc(1) = lstDoctors.SelectedItem.Text
        End If
    ElseIf Trim(Len(frmCaseAdjustment2016.txtATTDoc(2))) = 0 Then
        If Len(lstDoctors.SelectedItem.Text) Then
            frmCaseAdjustment2016.txtATTDoc(2) = lstDoctors.SelectedItem.Text
        End If
    ElseIf Trim(Len(frmCaseAdjustment2016.txtATTDoc(3))) = 0 Then
        If Len(lstDoctors.SelectedItem.Text) Then
            frmCaseAdjustment2016.txtATTDoc(3) = lstDoctors.SelectedItem.Text
        End If
    ElseIf Trim(Len(frmCaseAdjustment2016.txtATTDoc(4))) = 0 Then
        If Len(lstDoctors.SelectedItem.Text) Then
            frmCaseAdjustment2016.txtATTDoc(4) = lstDoctors.SelectedItem.Text
        End If
    End If
End Sub
