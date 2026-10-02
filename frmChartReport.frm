VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{65E121D4-0C60-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCHRT20.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmChartReport 
   Caption         =   "Chart Report"
   ClientHeight    =   8310
   ClientLeft      =   165
   ClientTop       =   345
   ClientWidth     =   11880
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   MDIChild        =   -1  'True
   ScaleHeight     =   8310
   ScaleWidth      =   11880
   WindowState     =   2  'Maximized
   Begin VB.Frame Frame7 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   405
      Left            =   0
      TabIndex        =   33
      Top             =   0
      Width           =   11895
      Begin VB.Frame Frame8 
         Caption         =   "Frame8"
         Height          =   15
         Left            =   240
         TabIndex        =   34
         Top             =   480
         Width           =   7215
      End
      Begin VB.Label Label3 
         Appearance      =   0  'Flat
         BackColor       =   &H00000000&
         Caption         =   "Chart Report"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000005&
         Height          =   315
         Left            =   180
         TabIndex        =   35
         Top             =   60
         Width           =   6615
      End
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   8055
      Left            =   0
      TabIndex        =   0
      Top             =   480
      Width           =   11895
      _ExtentX        =   20981
      _ExtentY        =   14208
      _Version        =   393216
      Tabs            =   2
      TabsPerRow      =   2
      TabHeight       =   520
      TabCaption(0)   =   "Summary by Payor"
      TabPicture(0)   =   "frmChartReport.frx":0000
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Frame3"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Frame5"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "Frame6"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "Frame1"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "Frame2"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "Frame4"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).ControlCount=   6
      TabCaption(1)   =   "Race Code by Payor"
      TabPicture(1)   =   "frmChartReport.frx":001C
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Frame12"
      Tab(1).Control(1)=   "Frame11"
      Tab(1).Control(2)=   "Frame10"
      Tab(1).Control(3)=   "Frame9"
      Tab(1).ControlCount=   4
      Begin VB.Frame Frame12 
         Height          =   735
         Left            =   -71280
         TabIndex        =   44
         Top             =   6720
         Width           =   8055
         Begin VB.CommandButton CmdShow2 
            Caption         =   "&Show"
            Height          =   375
            Left            =   3000
            TabIndex        =   18
            Top             =   240
            Width           =   855
         End
         Begin VB.CommandButton CmdExit2 
            Cancel          =   -1  'True
            Caption         =   "E&xit"
            Height          =   375
            Left            =   6240
            TabIndex        =   21
            Top             =   220
            Width           =   855
         End
         Begin VB.CommandButton CmdExport2 
            Caption         =   "&Export to Excel"
            Height          =   375
            Left            =   3960
            TabIndex        =   19
            Top             =   220
            Width           =   1215
         End
         Begin VB.CommandButton CmdReset2 
            Caption         =   "&Reset"
            Height          =   375
            Left            =   5280
            TabIndex        =   20
            Top             =   220
            Width           =   855
         End
      End
      Begin VB.Frame Frame11 
         Caption         =   "Race Code - By Payor"
         Height          =   4455
         Left            =   -74880
         TabIndex        =   42
         Top             =   3000
         Width           =   3375
         Begin MSComctlLib.ListView ListView2 
            Height          =   4095
            Left            =   120
            TabIndex        =   43
            Top             =   240
            Width           =   3135
            _ExtentX        =   5530
            _ExtentY        =   7223
            View            =   3
            Sorted          =   -1  'True
            LabelWrap       =   -1  'True
            HideSelection   =   -1  'True
            _Version        =   393217
            ForeColor       =   -2147483640
            BackColor       =   -2147483643
            BorderStyle     =   1
            Appearance      =   1
            NumItems        =   3
            BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Text            =   "Payor"
               Object.Width           =   1199
            EndProperty
            BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Alignment       =   2
               SubItemIndex    =   1
               Text            =   "Race Code"
               Object.Width           =   1235
            EndProperty
            BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Alignment       =   1
               SubItemIndex    =   2
               Text            =   "Total"
               Object.Width           =   1411
            EndProperty
         End
      End
      Begin VB.Frame Frame10 
         Caption         =   "Period Interval"
         Height          =   2415
         Left            =   -74880
         TabIndex        =   38
         Top             =   480
         Width           =   3375
         Begin VB.ComboBox CboPayor 
            Height          =   315
            Left            =   1200
            TabIndex        =   17
            Top             =   1940
            Width           =   1935
         End
         Begin VB.OptionButton Option15 
            Caption         =   "By Bordx Date"
            Height          =   255
            Left            =   240
            TabIndex        =   12
            Top             =   240
            Value           =   -1  'True
            Width           =   1695
         End
         Begin VB.OptionButton Option14 
            Caption         =   "By Effective Date"
            Height          =   255
            Left            =   240
            TabIndex        =   13
            Top             =   525
            Width           =   1575
         End
         Begin VB.OptionButton Option13 
            Caption         =   "By Transaction Date"
            Height          =   255
            Left            =   240
            TabIndex        =   14
            Top             =   825
            Width           =   2295
         End
         Begin MSComCtl2.DTPicker DTPicker3 
            Height          =   360
            Left            =   1200
            TabIndex        =   15
            Top             =   1200
            Width           =   1935
            _ExtentX        =   3413
            _ExtentY        =   635
            _Version        =   393216
            Format          =   25100289
            CurrentDate     =   37023
         End
         Begin MSComCtl2.DTPicker DTPicker4 
            Height          =   360
            Left            =   1200
            TabIndex        =   16
            Top             =   1575
            Width           =   1935
            _ExtentX        =   3413
            _ExtentY        =   635
            _Version        =   393216
            Format          =   25100289
            CurrentDate     =   37023
         End
         Begin VB.Label Label6 
            Caption         =   "Payor"
            Height          =   255
            Left            =   240
            TabIndex        =   41
            Top             =   2040
            Width           =   855
         End
         Begin VB.Label Label5 
            Caption         =   "End Date"
            Height          =   255
            Left            =   240
            TabIndex        =   40
            Top             =   1680
            Width           =   855
         End
         Begin VB.Label Label4 
            Caption         =   "Start Date"
            Height          =   255
            Left            =   240
            TabIndex        =   39
            Top             =   1320
            Width           =   855
         End
      End
      Begin VB.Frame Frame9 
         Height          =   6135
         Left            =   -71280
         TabIndex        =   36
         Top             =   480
         Width           =   8055
         Begin MSChart20Lib.MSChart MSChart3 
            Height          =   5775
            Left            =   360
            OleObjectBlob   =   "frmChartReport.frx":0038
            TabIndex        =   37
            Top             =   240
            Width           =   7575
         End
      End
      Begin VB.Frame Frame4 
         Height          =   6255
         Left            =   3720
         TabIndex        =   29
         Top             =   480
         Width           =   8055
         Begin MSChart20Lib.MSChart MSChart2 
            Height          =   5775
            Left            =   360
            OleObjectBlob   =   "frmChartReport.frx":19DF
            TabIndex        =   30
            Top             =   240
            Width           =   7575
         End
         Begin MSChart20Lib.MSChart MSChart1 
            Height          =   5775
            Left            =   360
            OleObjectBlob   =   "frmChartReport.frx":3375
            TabIndex        =   31
            Top             =   360
            Width           =   7575
         End
      End
      Begin VB.Frame Frame2 
         Caption         =   "Summay - By Payor"
         Height          =   2535
         Left            =   120
         TabIndex        =   27
         Top             =   2520
         Width           =   3375
         Begin MSComctlLib.ListView ListView1 
            Height          =   2175
            Left            =   120
            TabIndex        =   28
            Top             =   240
            Width           =   3135
            _ExtentX        =   5530
            _ExtentY        =   3836
            View            =   3
            Sorted          =   -1  'True
            LabelWrap       =   -1  'True
            HideSelection   =   -1  'True
            _Version        =   393217
            ForeColor       =   -2147483640
            BackColor       =   -2147483643
            BorderStyle     =   1
            Appearance      =   1
            NumItems        =   2
            BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Object.Width           =   1764
            EndProperty
            BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   1
               Object.Width           =   2999
            EndProperty
         End
      End
      Begin VB.Frame Frame1 
         Caption         =   "Period Interval"
         Height          =   2055
         Left            =   120
         TabIndex        =   24
         Top             =   480
         Width           =   3375
         Begin VB.OptionButton Option8 
            Caption         =   "By Transaction Date"
            Height          =   255
            Left            =   240
            TabIndex        =   3
            Top             =   825
            Width           =   2295
         End
         Begin VB.OptionButton Option7 
            Caption         =   "By Effective Date"
            Height          =   255
            Left            =   240
            TabIndex        =   2
            Top             =   525
            Width           =   1575
         End
         Begin VB.OptionButton Option6 
            Caption         =   "By Bordx Date"
            Height          =   255
            Left            =   240
            TabIndex        =   1
            Top             =   240
            Value           =   -1  'True
            Width           =   1695
         End
         Begin MSComCtl2.DTPicker DTPicker1 
            Height          =   360
            Left            =   1200
            TabIndex        =   4
            Top             =   1200
            Width           =   1935
            _ExtentX        =   3413
            _ExtentY        =   635
            _Version        =   393216
            Format          =   24641537
            CurrentDate     =   37023
         End
         Begin MSComCtl2.DTPicker DTPicker2 
            Height          =   360
            Left            =   1200
            TabIndex        =   5
            Top             =   1575
            Width           =   1935
            _ExtentX        =   3413
            _ExtentY        =   635
            _Version        =   393216
            Format          =   24641537
            CurrentDate     =   37023
         End
         Begin VB.Label Label1 
            Caption         =   "Start Date"
            Height          =   255
            Left            =   240
            TabIndex        =   26
            Top             =   1320
            Width           =   855
         End
         Begin VB.Label Label2 
            Caption         =   "End Date"
            Height          =   255
            Left            =   240
            TabIndex        =   25
            Top             =   1680
            Width           =   855
         End
      End
      Begin VB.Frame Frame6 
         Caption         =   "Type of Chart"
         Height          =   735
         Left            =   120
         TabIndex        =   23
         Top             =   6720
         Width           =   3375
         Begin VB.OptionButton Option9 
            Caption         =   "Bar Chart"
            Height          =   255
            Left            =   1200
            TabIndex        =   7
            Top             =   360
            Width           =   975
         End
         Begin VB.OptionButton Option10 
            Caption         =   "Pie Chart"
            Height          =   255
            Left            =   2270
            TabIndex        =   8
            Top             =   360
            Width           =   975
         End
         Begin VB.OptionButton Option11 
            Caption         =   "Line Chart"
            Height          =   255
            Left            =   120
            TabIndex        =   6
            Top             =   360
            Value           =   -1  'True
            Width           =   1095
         End
      End
      Begin VB.Frame Frame5 
         Height          =   735
         Left            =   3720
         TabIndex        =   22
         Top             =   6720
         Width           =   8055
         Begin VB.CommandButton CmdShow 
            Caption         =   "&Show"
            Height          =   375
            Left            =   2640
            TabIndex        =   45
            Top             =   240
            Width           =   814
         End
         Begin VB.CommandButton CmdReset 
            Caption         =   "&Reset"
            Height          =   375
            Left            =   4920
            TabIndex        =   10
            Top             =   220
            Width           =   855
         End
         Begin VB.CommandButton CmdExport 
            Caption         =   "&Export to Excel"
            Height          =   375
            Left            =   3600
            TabIndex        =   9
            Top             =   240
            Width           =   1215
         End
         Begin VB.CommandButton CmdExit 
            Caption         =   "E&xit"
            Height          =   375
            Left            =   5880
            TabIndex        =   11
            Top             =   220
            Width           =   855
         End
      End
      Begin VB.Frame Frame3 
         Height          =   1695
         Left            =   120
         TabIndex        =   32
         Top             =   4980
         Width           =   3375
         Begin VB.OptionButton Option1 
            Caption         =   "Total Membership By Payor"
            Height          =   375
            Left            =   120
            TabIndex        =   50
            Top             =   120
            Value           =   -1  'True
            Width           =   2415
         End
         Begin VB.OptionButton Option2 
            Caption         =   "Total Principals By Payor"
            Height          =   255
            Left            =   120
            TabIndex        =   49
            Top             =   495
            Width           =   2535
         End
         Begin VB.OptionButton Option3 
            Caption         =   "Total Supplementary By Payor"
            Height          =   375
            Left            =   120
            TabIndex        =   48
            Top             =   735
            Width           =   2535
         End
         Begin VB.OptionButton Option4 
            Caption         =   "Total Premium By Payor"
            Height          =   255
            Left            =   120
            TabIndex        =   47
            Top             =   1095
            Width           =   2175
         End
         Begin VB.OptionButton Option5 
            Caption         =   "Total MCO Fees By Payor"
            Height          =   255
            Left            =   120
            TabIndex        =   46
            Top             =   1380
            Width           =   2175
         End
      End
   End
End
Attribute VB_Name = "frmChartReport"
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


Private Sub Form_Load()
    DTPicker1.value = Date
    DTPicker2.value = Date
    DTPicker3.value = Date
    DTPicker4.value = Date
    Call ComboListing
    Tool_Tip
End Sub

'================ ComboListing ============================

Private Sub ComboListing()
Dim PAY As New ADODB.Recordset
Dim Sql As String

PAY.Open "Select PAYCode, PAYCompanyName from PAY", medixmasterconn, adOpenKeyset, adLockOptimistic

   Do Until PAY.EOF
      CboPayor.AddItem PAY!PAYCode & " - " & Trim(PAY!PAYCompanyName)
      PAY.MoveNext
   Loop
   PAY.Close
   Set PAY = Nothing
 
End Sub
'================ End ComboListing ============================

'=================== Command Button =============================
Private Sub cmdExit_Click()
Dim RecordExit
    RecordExit = MsgBox("Quit this form?", vbYesNo + vbExclamation)
    If RecordExit = vbYes Then
        bExit = True
        Unload Me
    End If
End Sub

Private Sub CmdExit2_Click()
Dim RecordExit
    RecordExit = MsgBox("Quit this form?", vbYesNo + vbExclamation)
    If RecordExit = vbYes Then
        bExit = True
        Unload Me
    End If
End Sub

Private Sub CmdReset2_Click()
    DTPicker3.value = Date
    DTPicker4.value = Date
    CboPayor.Text = ""
    ListView2.listitems.Clear
    Call RaceCode
End Sub

Private Sub CmdReset_Click()
    DTPicker1.value = Date
    DTPicker2.value = Date
    ListView1.listitems.Clear
    
    If Option1.value = True Then
        Call Membership
    ElseIf Option2.value = True Then
        Call Principals
    ElseIf Option3.value = True Then
        Call Supplementary
    ElseIf Option4.value = True Then
        Call Premium
    ElseIf Option5.value = True Then
        Call MCOFees
    End If
End Sub

Private Sub CmdShow_Click()
    If Option1.value = True Then
        Call Membership
    ElseIf Option2.value = True Then
        Call Principals
    ElseIf Option3.value = True Then
        Call Supplementary
    ElseIf Option4.value = True Then
        Call Premium
    ElseIf Option5.value = True Then
        Call MCOFees
    End If
End Sub

Private Sub CmdShow2_Click()
    
    If CboPayor = "" Then
        MsgBox "Please Enter Payor Code.", vbCritical
        CboPayor.SetFocus
    Else
        Screen.MousePointer = vbHourglass
            Call RaceCode
        Screen.MousePointer = Default
    End If

End Sub
'=================== End Command Button =============================

'=================== Check Valid Date ==========================

Private Function DateValidDate() As Boolean
    If DTPicker1.value <= DTPicker2.value Then
       DateValidDate = True
    Else
       DateValidDate = False
    End If
    
    If DTPicker3.value <= DTPicker4.value Then
       DateValidDate = True
    Else
       DateValidDate = False
    End If
    
End Function
'=================== End Check Valid Date ==========================

'=================== List Records to List View============================
' 1) Total Membership
Private Sub Membership()
Dim strsql1 As String
Dim rst1 As New ADODB.Recordset
Dim InfoList As ListItem
    
    ListView1.listitems.Clear
    MSChart1.Refresh
       
    Screen.MousePointer = vbHourglass
    
    If Option6.value = True Then
    
        strsql1 = "SELECT PAYCode, COUNT(*) AS PRODUCTION "
        strsql1 = strsql1 & " From MBM "
        strsql1 = strsql1 & "WHERE MBMBordxDate >= CONVERT(DATETIME,  '" & DTPicker1.value & "', 103) "
        strsql1 = strsql1 & "AND MBMBordxDate <= CONVERT(DATETIME,  '" & DTPicker2.value & "', 103) "
        strsql1 = strsql1 & "GROUP BY PAYCode"
        
    ElseIf Option7.value = True Then
        strsql1 = "SELECT PAYCode, COUNT(*) AS PRODUCTION "
        strsql1 = strsql1 & " From MBM "
        strsql1 = strsql1 & "WHERE MBMPayorEffDate >= CONVERT(DATETIME,  '" & DTPicker1.value & "', 103) "
        strsql1 = strsql1 & "AND MBMPayorEffDate <= CONVERT(DATETIME,  '" & DTPicker2.value & "', 103) "
        strsql1 = strsql1 & "GROUP BY PAYCode"
    ElseIf Option8.value = True Then
        strsql1 = "SELECT PAYCode, COUNT(*) AS PRODUCTION "
        strsql1 = strsql1 & " From MBM "
        strsql1 = strsql1 & "WHERE MBMValueDate >= CONVERT(DATETIME,  '" & DTPicker1.value & "', 103) "
        strsql1 = strsql1 & "AND MBMValueDate <= CONVERT(DATETIME,  '" & DTPicker2.value & "', 103) "
        strsql1 = strsql1 & "GROUP BY PAYCode"
    End If
        
    rst1.Open strsql1, medixmasterconn, adOpenKeyset, adLockOptimistic
    
    ListView1.ColumnHeaders(1).Text = "Payor Code"
    ListView1.ColumnHeaders(2).Text = "No Of Membership"
    Do Until rst1.EOF
     Set InfoList = ListView1.listitems.Add(, , rst1!PAYCode)
         InfoList.SubItems(1) = rst1!PRODUCTION
         rst1.MoveNext
    Loop
    MSChart1.ShowLegend = True
    'MSChart1.TitleText = "Total Membership"
    Set MSChart1.DataSource = rst1
    
    MSChart2.ShowLegend = True
    'MSChart2.TitleText = "Total Membership"
    Set MSChart2.DataSource = rst1
    
    rst1.Close
    Screen.MousePointer = vbDefault
End Sub

' 2) Total Principals
Private Sub Principals()
Dim strsql2 As String
Dim rst2 As New ADODB.Recordset
Dim InfoList As ListItem
    
    ListView1.listitems.Clear
    MSChart1.Refresh
       
    Screen.MousePointer = vbHourglass
    
    If Option6.value = True Then
        strsql2 = "SELECT PAYCode, COUNT(*) AS PRODUCTION "
        strsql2 = strsql2 & " From MBM "
        strsql2 = strsql2 & "WHERE MBMBordxDate >= CONVERT(DATETIME,  '" & DTPicker1.value & "', 103) "
        strsql2 = strsql2 & "AND MBMBordxDate <= CONVERT(DATETIME,  '" & DTPicker2.value & "', 103) "
        strsql2 = strsql2 & "GROUP BY PAYCode, MBMMemberType "
        strsql2 = strsql2 & "HAVING (MBMMemberType = 'P')"
    ElseIf Option7.value = True Then
        strsql2 = "SELECT PAYCode, COUNT(*) AS PRODUCTION "
        strsql2 = strsql2 & " From MBM "
        strsql2 = strsql2 & "WHERE MBMPayorEffDate >= CONVERT(DATETIME,  '" & DTPicker1.value & "', 103) "
        strsql2 = strsql2 & "AND MBMPayorEffDate <= CONVERT(DATETIME,  '" & DTPicker2.value & "', 103) "
        strsql2 = strsql2 & "GROUP BY PAYCode, MBMMemberType "
        strsql2 = strsql2 & "HAVING (MBMMemberType = 'P')"
    ElseIf Option8.value = True Then
        strsql2 = "SELECT PAYCode, COUNT(*) AS PRODUCTION "
        strsql2 = strsql2 & " From MBM "
        strsql2 = strsql2 & "WHERE MBMValueDate >= CONVERT(DATETIME,  '" & DTPicker1.value & "', 103) "
        strsql2 = strsql2 & "AND MBMValueDate <= CONVERT(DATETIME,  '" & DTPicker2.value & "', 103) "
        strsql2 = strsql2 & "GROUP BY PAYCode, MBMMemberType "
        strsql2 = strsql2 & "HAVING (MBMMemberType = 'P')"
    End If
    
    rst2.Open strsql2, medixmasterconn, adOpenKeyset, adLockOptimistic
           
    ListView1.ColumnHeaders(1).Text = "Payor Code"
    ListView1.ColumnHeaders(2).Text = "No Of Proicipals"
    Do Until rst2.EOF
     Set InfoList = ListView1.listitems.Add(, , rst2!PAYCode)
         InfoList.SubItems(1) = rst2!PRODUCTION
         rst2.MoveNext
    Loop
    MSChart1.ShowLegend = True
    'MSChart1.TitleText = "Total Principals"
    Set MSChart1.DataSource = rst2
    
    MSChart2.ShowLegend = True
    'MSChart2.TitleText = "Total Principals"
    Set MSChart2.DataSource = rst2
    
    rst2.Close
    Screen.MousePointer = vbDefault
End Sub

' 3) Total Supplementary

Private Sub Supplementary()
Dim strsql3 As String
Dim rst3 As New ADODB.Recordset
Dim InfoList As ListItem
    
    ListView1.listitems.Clear
    MSChart1.Refresh
       
    Screen.MousePointer = vbHourglass
    
    If Option6.value = True Then
        strsql3 = "SELECT PAYCode, COUNT(*) AS PRODUCTION "
        'strsql3 = strsql3 & " From MBM "
        strsql3 = strsql3 & " From View_ChartReport "
        strsql3 = strsql3 & "WHERE MBMCBordxDate >= CONVERT(DATETIME,  '" & DTPicker1.value & "', 103) "
        strsql3 = strsql3 & "AND MBMCBordxDate <= CONVERT(DATETIME,  '" & DTPicker2.value & "', 103) "
        strsql3 = strsql3 & "GROUP BY PAYCode, MBMCMemberType "
        strsql3 = strsql3 & "HAVING (MBMCMemberType = 'C')"
    ElseIf Option7.value = True Then
        strsql3 = "SELECT PAYCode, COUNT(*) AS PRODUCTION "
        'strsql3 = strsql3 & " From MBM "
        strsql3 = strsql3 & " From View_ChartReport "
        strsql3 = strsql3 & "WHERE MBMPayorEffDate >= CONVERT(DATETIME,  '" & DTPicker1.value & "', 103) "
        strsql3 = strsql3 & "AND MBMPayorEffDate <= CONVERT(DATETIME,  '" & DTPicker2.value & "', 103) "
        strsql3 = strsql3 & "GROUP BY PAYCode, MBMCMemberType "
        strsql3 = strsql3 & "HAVING (MBMCMemberType = 'C')"
    ElseIf Option8.value = True Then
        strsql3 = "SELECT PAYCode, COUNT(*) AS PRODUCTION "
        'strsql3 = strsql3 & " From MBM "
        strsql3 = strsql3 & " From View_ChartReport "
        strsql3 = strsql3 & "WHERE MBMValueDate >= CONVERT(DATETIME,  '" & DTPicker1.value & "', 103) "
        strsql3 = strsql3 & "AND MBMValueDate <= CONVERT(DATETIME,  '" & DTPicker2.value & "', 103) "
        strsql3 = strsql3 & "GROUP BY PAYCode, MBMCMemberType "
        strsql3 = strsql3 & "HAVING (MBMCMemberType = 'C')"
    End If
    
    rst3.Open strsql3, medixmasterconn, adOpenKeyset, adLockOptimistic
    
    ListView1.ColumnHeaders(1).Text = "Payor Code"
    ListView1.ColumnHeaders(2).Text = "No Of Supplementary"
    Do Until rst3.EOF
     Set InfoList = ListView1.listitems.Add(, , rst3!PAYCode)
         InfoList.SubItems(1) = rst3!PRODUCTION
         rst3.MoveNext
    Loop
    MSChart1.ShowLegend = True
    'MSChart1.TitleText = "Total Supplementarys"
    Set MSChart1.DataSource = rst3
    
    MSChart2.ShowLegend = True
    'MSChart2.TitleText = "Total Supplementarys"
    Set MSChart2.DataSource = rst3
    
    rst3.Close
    Screen.MousePointer = vbDefault
End Sub

' 4) Total Premium

Private Sub Premium()
Dim strsql4 As String
Dim rst4 As New ADODB.Recordset
Dim InfoList As ListItem
    
    ListView1.listitems.Clear
    MSChart1.Refresh
       
    Screen.MousePointer = vbHourglass
    
    If Option6.value = True Then
        strsql4 = "SELECT PAYCode, SUM(MBMPremium) AS PRODUCTION "
        strsql4 = strsql4 & " From VIEW_MBM_MBMOthers "
        strsql4 = strsql4 & "WHERE MBMBordxDate >= CONVERT(DATETIME,  '" & DTPicker1.value & "', 103) "
        strsql4 = strsql4 & "AND MBMBordxDate <= CONVERT(DATETIME,  '" & DTPicker2.value & "', 103) "
        strsql4 = strsql4 & "GROUP BY PAYCode, MBMMemberType "
        strsql4 = strsql4 & "HAVING (MBMMemberType = 'P')"
    ElseIf Option7.value = True Then
        strsql4 = "SELECT PAYCode, SUM(MBMPremium) AS PRODUCTION "
        strsql4 = strsql4 & " From VIEW_MBM_MBMOthers "
        strsql4 = strsql4 & "WHERE MBMPayorEffDate >= CONVERT(DATETIME,  '" & DTPicker1.value & "', 103) "
        strsql4 = strsql4 & "AND MBMPayorEffDate <= CONVERT(DATETIME,  '" & DTPicker2.value & "', 103) "
        strsql4 = strsql4 & "GROUP BY PAYCode, MBMMemberType "
        strsql4 = strsql4 & "HAVING (MBMMemberType = 'P')"
    ElseIf Option8.value = True Then
        strsql4 = "SELECT PAYCode, SUM(MBMPremium) AS PRODUCTION "
        strsql4 = strsql4 & " From VIEW_MBM_MBMOthers "
        strsql4 = strsql4 & "WHERE MBMValueDate >= CONVERT(DATETIME,  '" & DTPicker1.value & "', 103) "
        strsql4 = strsql4 & "AND MBMValueDate <= CONVERT(DATETIME,  '" & DTPicker2.value & "', 103) "
        strsql4 = strsql4 & "GROUP BY PAYCode, MBMMemberType "
        strsql4 = strsql4 & "HAVING (MBMMemberType = 'P')"
    End If
    
    rst4.Open strsql4, medixmasterconn, adOpenKeyset, adLockOptimistic
    
    ListView1.ColumnHeaders(1).Text = "Payor Code"
    ListView1.ColumnHeaders(2).Text = "Total Premiums"
    Do Until rst4.EOF
     Set InfoList = ListView1.listitems.Add(, , rst4!PAYCode)
         If Len(Trim(rst4!PRODUCTION)) Then
            InfoList.SubItems(1) = rst4!PRODUCTION
         End If
         rst4.MoveNext
    Loop
    
    MSChart1.ShowLegend = True
    'MSChart1.TitleText = "Total Premiums"
    Set MSChart1.DataSource = rst4
    
    MSChart2.ShowLegend = True
    'MSChart2.TitleText = "Total Premiums"
    Set MSChart2.DataSource = rst4

    rst4.Close
    Screen.MousePointer = vbDefault
End Sub

' 5) Total MCO Fees

Private Sub MCOFees()
Dim strsql5 As String
Dim rst5 As New ADODB.Recordset
Dim InfoList As ListItem
    
    ListView1.listitems.Clear
    MSChart1.Refresh
       
    Screen.MousePointer = vbHourglass
    
    If Option6.value = True Then
        strsql5 = "SELECT PAYCode, SUM(MBMMCOFee) AS PRODUCTION "
        strsql5 = strsql5 & " From VIEW_MBM_MBMOthers "
        strsql5 = strsql5 & "WHERE MBMBordxDate >= CONVERT(DATETIME,  '" & DTPicker1.value & "', 103) "
        strsql5 = strsql5 & "AND MBMBordxDate <= CONVERT(DATETIME,  '" & DTPicker2.value & "', 103) "
        strsql5 = strsql5 & "GROUP BY PAYCode, MBMMemberType "
        strsql5 = strsql5 & "HAVING (MBMMemberType = 'P')"
    ElseIf Option7.value = True Then
        strsql5 = "SELECT PAYCode, SUM(MBMMCOFee) AS PRODUCTION "
        strsql5 = strsql5 & " From VIEW_MBM_MBMOthers "
        strsql5 = strsql5 & "WHERE MBMPayorEffDate >= CONVERT(DATETIME,  '" & DTPicker1.value & "', 103) "
        strsql5 = strsql5 & "AND MBMPayorEffDate <= CONVERT(DATETIME,  '" & DTPicker2.value & "', 103) "
        strsql5 = strsql5 & "GROUP BY PAYCode, MBMMemberType "
        strsql5 = strsql5 & "HAVING (MBMMemberType = 'P')"
    ElseIf Option8.value = True Then
        strsql5 = "SELECT PAYCode, SUM(MBMMCOFee) AS PRODUCTION "
        strsql5 = strsql5 & " From VIEW_MBM_MBMOthers "
        strsql5 = strsql5 & "WHERE MBMValueDate >= CONVERT(DATETIME,  '" & DTPicker1.value & "', 103) "
        strsql5 = strsql5 & "AND MBMValueDate <= CONVERT(DATETIME,  '" & DTPicker2.value & "', 103) "
        strsql5 = strsql5 & "GROUP BY PAYCode, MBMMemberType "
        strsql5 = strsql5 & "HAVING (MBMMemberType = 'P')"
    End If
    
    rst5.Open strsql5, medixmasterconn, adOpenKeyset, adLockOptimistic
    
    ListView1.ColumnHeaders(1).Text = "Payor Code"
    ListView1.ColumnHeaders(2).Text = "Total MCO Fees"
    Do Until rst5.EOF
     Set InfoList = ListView1.listitems.Add(, , rst5!PAYCode)
         If Len(Trim(rst5!PRODUCTION)) Then
            InfoList.SubItems(1) = rst5!PRODUCTION
         End If
         rst5.MoveNext
    Loop
    MSChart1.ShowLegend = True
    'MSChart1.TitleText = "Total MCO Fees"
    Set MSChart1.DataSource = rst5
    
    MSChart2.ShowLegend = True
    'MSChart2.TitleText = "Total MCO Fees"
    Set MSChart2.DataSource = rst5
    
    rst5.Close
    Screen.MousePointer = vbDefault
End Sub

' 6) Total Race Code

Private Sub RaceCode()
Dim strsql6 As String
Dim rst6 As New ADODB.Recordset
Dim InfoList As ListItem
Dim Percent As Variant
    
    ListView2.listitems.Clear
    MSChart3.Refresh
       
        
    If Option15.value = True Then
        strsql6 = "Select PAYCode, RACCode, COUNT(RACCode) AS Total"
        strsql6 = strsql6 & " From MBM "
        strsql6 = strsql6 & "WHERE MBMBordxDate >= CONVERT(DATETIME,  '" & DTPicker3.value & "', 103) "
        strsql6 = strsql6 & "AND MBMBordxDate <= CONVERT(DATETIME,  '" & DTPicker4.value & "', 103) "
        strsql6 = strsql6 & "AND PAYCode = '" & Trim(Mid(CboPayor, 1, 2)) & "' "
        strsql6 = strsql6 & "GROUP BY PAYCode, RACCode "
        strsql6 = strsql6 & "HAVING (RACCode IS NOT NULL)"
    ElseIf Option14.value = True Then
        strsql6 = "Select PAYCode, RACCode, COUNT(RACCode) AS Total"
        strsql6 = strsql6 & " From MBM "
        strsql6 = strsql6 & "WHERE MBMPayorEffDate >= CONVERT(DATETIME,  '" & DTPicker3.value & "', 103) "
        strsql6 = strsql6 & "AND MBMPayorEffDate <= CONVERT(DATETIME,  '" & DTPicker4.value & "', 103) "
        strsql6 = strsql6 & "AND PAYCode = '" & Trim(Mid(CboPayor, 1, 2)) & "' "
        strsql6 = strsql6 & "GROUP BY PAYCode, RACCode "
        strsql6 = strsql6 & "HAVING (RACCode IS NOT NULL)"
    ElseIf Option13.value = True Then
        strsql6 = "Select PAYCode, RACCode, COUNT(RACCode) AS Total"
        strsql6 = strsql6 & " From MBM "
        strsql6 = strsql6 & "WHERE MBMValueDate >= CONVERT(DATETIME,  '" & DTPicker3.value & "', 103) "
        strsql6 = strsql6 & "AND MBMValueDate <= CONVERT(DATETIME,  '" & DTPicker4.value & "', 103) "
        strsql6 = strsql6 & "AND PAYCode = '" & Trim(Mid(CboPayor, 1, 2)) & "' "
        strsql6 = strsql6 & "GROUP BY PAYCode, RACCode "
        strsql6 = strsql6 & "HAVING (RACCode IS NOT NULL)"
    End If
    
    rst6.Open strsql6, medixmasterconn, adOpenKeyset, adLockOptimistic
    
       
    Do Until rst6.EOF
     Set InfoList = ListView2.listitems.Add(, , rst6!PAYCode)
         If Len(rst6!RACCode) Then
            InfoList.SubItems(1) = rst6!RACCode
         End If
         If Len(Trim(rst6!Total)) Then
            InfoList.SubItems(2) = rst6!Total
         End If
                         
         rst6.MoveNext
    Loop
    MSChart3.ShowLegend = True
    'MSChart3.FootnoteText = "Total Race Code"
    Set MSChart3.DataSource = rst6
    
    rst6.Close
    Set rst6 = Nothing
    
End Sub
'=================== End List Records to List View============================


'======================== Export to Excel =============================
Private Sub CmdExport_Click()
      
    Dim objExcel As excel.Application
    Dim objWorkbook As excel.Workbook
    Dim objWorksheet As excel.Worksheet
    
    If objExcel Is Nothing Then
        'Start the excel COM and make it visible.
        Set objExcel = GetObject("", "excel.application")
        objExcel.Visible = True
           
        'Start a workbook.
        Set objWorkbook = objExcel.Workbooks.Add
        'Turn off the alerts, otherwise user will have to confirm my actions.
        objExcel.DisplayAlerts = False
    End If
    'Depending on the users excel's settings, there could be many worksheet when starting a workbook.
    'Ensure there is only one worksheet.
    Do While objWorkbook.Worksheets.count > 1
        Set objWorksheet = objWorkbook.Worksheets.Item(objWorkbook.Worksheets.count)
        objWorksheet.Delete
    Loop
    
    'Set objWorksheet to the remaining worksheet.
    Set objWorksheet = objWorkbook.Worksheets.Add
    objWorksheet.Name = "Results"
        
        
    If Option9.value = True Then
        MSChart1.EditCopy
        objWorksheet.PasteSpecial 2
        objWorksheet.PasteSpecial 1
        objWorksheet.Cells(1, "A") = "PAYOR"
    ElseIf Option10.value = True Then
        MSChart1.EditCopy
        objWorksheet.PasteSpecial 2
        objWorksheet.PasteSpecial 1
        objWorksheet.Cells(1, "A") = "PAYOR"
    ElseIf Option11.value = True Then
        MSChart2.EditCopy
        objWorksheet.PasteSpecial 2
        objWorksheet.PasteSpecial 1
        objWorksheet.Cells(1, "A") = "PAYOR"
    End If
End Sub

Private Sub CmdExport2_Click()
Dim objExcel As excel.Application
Dim objWorkbook As excel.Workbook
Dim objWorksheet As excel.Worksheet
    
    If objExcel Is Nothing Then
        'Start the excel COM and make it visible.
        Set objExcel = GetObject("", "excel.application")
        objExcel.Visible = True
           
        'Start a workbook.
        Set objWorkbook = objExcel.Workbooks.Add
        'Turn off the alerts, otherwise user will have to confirm my actions.
        objExcel.DisplayAlerts = False
    End If
    'Depending on the users excel's settings, there could be many worksheet when starting a workbook.
    'Ensure there is only one worksheet.
    Do While objWorkbook.Worksheets.count > 1
        Set objWorksheet = objWorkbook.Worksheets.Item(objWorkbook.Worksheets.count)
        objWorksheet.Delete
    Loop
    
    'Set objWorksheet to the remaining worksheet.
    Set objWorksheet = objWorkbook.Worksheets.Add
    objWorksheet.Name = "Results"
           
    MSChart3.EditCopy
    objWorksheet.PasteSpecial 2
    objWorksheet.PasteSpecial 1
    objWorksheet.Cells(1, "A") = "Payor"
    objWorksheet.Cells(1, "B") = "Race Code"
    
End Sub
'======================== End Export to Excel =============================

'================= Sort List View ===============================
Private Sub ListView1_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
    ListView1.SortKey = ColumnHeader.Index - 1
    ListView1.Sorted = True
    If ListView1.SortOrder = lvwAscending Then
        ListView1.SortOrder = lvwDescending
    Else
        ListView1.SortOrder = lvwAscending
    End If
End Sub

Private Sub ListView2_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
    ListView2.SortKey = ColumnHeader.Index - 1
    ListView2.Sorted = True
    If ListView2.SortOrder = lvwAscending Then
        ListView2.SortOrder = lvwDescending
    Else
        ListView2.SortOrder = lvwAscending
    End If
End Sub
'================= End Sort List View ===============================


'=================== Option Click ============================
Private Sub Option10_Click()
    MSChart2.Visible = False
    MSChart1.Visible = True
    MSChart1.ChartType = VtChChartType2dPie

End Sub

Private Sub Option11_Click()
    MSChart1.Visible = False
    MSChart2.Visible = True
    MSChart2.ChartType = VtChChartType2dLine
End Sub

Private Sub Option9_Click()
    MSChart2.Visible = False
    MSChart1.Visible = True
    MSChart1.ChartType = VtChChartType2dBar
End Sub
'=================== End Option Click ============================


'================= Combobox Change/ KeyPress =================
Private Sub CboPayor_Change()
If CboPayor.Text = "" Then
    If InStr(CboPayor.Text, "null") Then
       CboPayor.Enabled = False
    End If
End If
End Sub

Private Sub CboPayor_KeyPress(KeyAscii As Integer)
Dim NewText As String
Dim ValidCount As Integer
Dim ValidValue As String
Dim i As Integer
    
If CboPayor.Text = "" Then
    If KeyAscii >= 32 And KeyAscii <> 127 Then
    NewText = LCase(Left(CboPayor.Text, CboPayor.SelStart) + Chr(KeyAscii) + Mid(CboPayor.Text, CboPayor.SelStart + CboPayor.SelLength + 1))
 
    ValidCount = 0
    ValidValue = ""

    For i = 0 To CboPayor.ListCount - 1
 
        If NewText = LCase(Left(CboPayor.List(i), Len(NewText))) Then
                ValidCount = ValidCount + 1
                ValidValue = CboPayor.List(i)
        End If
        Next
        If ValidCount <= 1 Then KeyAscii = 0
            If ValidCount = 1 Then
                CboPayor.Text = ValidValue
                CboPayor.SelStart = 0
                CboPayor.SelLength = Len(ValidValue)
            End If
        End If
End If
End Sub
'================= End Combobox Change/ KeyPress =================
