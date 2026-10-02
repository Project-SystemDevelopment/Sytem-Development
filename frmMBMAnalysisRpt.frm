VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmMBMAnalysisRpt 
   Caption         =   "frmMemberReports"
   ClientHeight    =   7980
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   11880
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   7980
   ScaleWidth      =   11880
   WindowState     =   2  'Maximized
   Begin VB.Frame Frame2 
      Height          =   615
      Left            =   0
      TabIndex        =   0
      Top             =   7320
      Width           =   11655
      Begin VB.CommandButton CmdStop 
         Caption         =   "S&top"
         Height          =   375
         Left            =   6240
         TabIndex        =   204
         Top             =   160
         Width           =   855
      End
      Begin VB.CommandButton CmdView 
         Caption         =   "&View"
         Height          =   375
         Left            =   7920
         TabIndex        =   127
         Top             =   160
         Width           =   855
      End
      Begin VB.CommandButton CmdBack 
         Caption         =   "&Back"
         Height          =   375
         Left            =   7080
         TabIndex        =   126
         Top             =   160
         Width           =   855
      End
      Begin VB.TextBox Text1 
         Height          =   285
         Left            =   2160
         TabIndex        =   133
         Top             =   600
         Visible         =   0   'False
         Width           =   495
      End
      Begin VB.TextBox txtDate 
         Height          =   285
         Left            =   2760
         TabIndex        =   132
         Top             =   600
         Visible         =   0   'False
         Width           =   735
      End
      Begin VB.CommandButton CmdPrint 
         Caption         =   "&Print to File"
         Height          =   375
         Left            =   8760
         TabIndex        =   128
         Top             =   160
         Width           =   975
      End
      Begin VB.CommandButton cmdExit 
         Cancel          =   -1  'True
         Caption         =   "&Exit"
         Height          =   375
         Left            =   10560
         TabIndex        =   129
         Top             =   160
         Width           =   855
      End
      Begin VB.CommandButton CmdClear 
         Caption         =   "&Clear"
         Height          =   375
         Left            =   9720
         TabIndex        =   125
         Top             =   160
         Width           =   855
      End
      Begin VB.CommandButton cmdSearch 
         Caption         =   "&Search"
         Default         =   -1  'True
         Height          =   375
         Left            =   5400
         TabIndex        =   124
         Top             =   160
         Width           =   855
      End
      Begin MSComctlLib.ListView ListView3 
         Height          =   255
         Left            =   3600
         TabIndex        =   198
         Top             =   600
         Visible         =   0   'False
         Width           =   1815
         _ExtentX        =   3201
         _ExtentY        =   450
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         FullRowSelect   =   -1  'True
         GridLines       =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         NumItems        =   32
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Object.Width           =   2469
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Object.Width           =   2469
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   2
            Object.Width           =   2469
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   3
            Object.Width           =   1764
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   4
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   5
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   6
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   7
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(9) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   8
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(10) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   9
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(11) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   10
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(12) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   11
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(13) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   12
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(14) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   13
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(15) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   14
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(16) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   15
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(17) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   16
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(18) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   17
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(19) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   18
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(20) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   19
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(21) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   20
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(22) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   21
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(23) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   22
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(24) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   23
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(25) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   24
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(26) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   25
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(27) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   26
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(28) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   27
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(29) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   28
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(30) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   29
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(31) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   30
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(32) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   31
            Object.Width           =   2540
         EndProperty
      End
      Begin VB.Label Label64 
         Caption         =   "Record(s)"
         Height          =   255
         Left            =   2640
         TabIndex        =   208
         Top             =   240
         Width           =   855
      End
      Begin VB.Label Label63 
         Caption         =   "of"
         Height          =   255
         Left            =   1320
         TabIndex        =   207
         Top             =   240
         Width           =   255
      End
      Begin VB.Label lblCurrent 
         Alignment       =   1  'Right Justify
         BackColor       =   &H80000009&
         Height          =   255
         Left            =   240
         TabIndex        =   206
         Top             =   240
         Width           =   975
      End
      Begin VB.Label lblTotalRecord 
         Alignment       =   1  'Right Justify
         BackColor       =   &H80000009&
         Height          =   255
         Left            =   1560
         TabIndex        =   205
         Top             =   240
         Width           =   975
      End
   End
   Begin MSComctlLib.ListView lstRaceList 
      Height          =   1095
      Left            =   -360
      TabIndex        =   193
      Top             =   8640
      Width           =   11775
      _ExtentX        =   20770
      _ExtentY        =   1931
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   8
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Diag Code"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Description"
         Object.Width           =   6174
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   2
         Text            =   "Total of Cases"
         Object.Width           =   2117
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   3
         Text            =   "Malays"
         Object.Width           =   1764
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   4
         Text            =   "Chinese"
         Object.Width           =   1764
      EndProperty
      BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   5
         Text            =   "Indians"
         Object.Width           =   1764
      EndProperty
      BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   6
         Text            =   "Others"
         Object.Width           =   1764
      EndProperty
      BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   7
         Text            =   "Empty"
         Object.Width           =   1764
      EndProperty
   End
   Begin VB.Frame Frame4 
      Caption         =   "Reports Type"
      Height          =   4095
      Left            =   8160
      TabIndex        =   191
      Top             =   360
      Width           =   3615
      Begin VB.OptionButton Option36 
         Caption         =   "Agent No"
         ForeColor       =   &H000000FF&
         Height          =   195
         Left            =   360
         TabIndex        =   98
         Top             =   1080
         Width           =   1335
      End
      Begin VB.OptionButton Option19 
         Caption         =   "Case DOR"
         Height          =   195
         Left            =   360
         TabIndex        =   99
         Top             =   1320
         Width           =   1095
      End
      Begin VB.OptionButton Option14 
         Caption         =   "Pol. Year"
         Height          =   195
         Left            =   1920
         TabIndex        =   116
         Top             =   1800
         Width           =   1335
      End
      Begin VB.OptionButton Option13 
         Caption         =   "Renewal"
         Height          =   195
         Left            =   1920
         TabIndex        =   118
         Top             =   2280
         Width           =   1335
      End
      Begin VB.OptionButton Option12 
         Caption         =   "Doctor"
         Height          =   195
         Left            =   360
         TabIndex        =   104
         Top             =   2520
         Width           =   1095
      End
      Begin VB.OptionButton Option28 
         Caption         =   "Exp Date"
         Height          =   195
         Left            =   360
         TabIndex        =   107
         Top             =   3240
         Width           =   1335
      End
      Begin VB.OptionButton Option27 
         Caption         =   "Eff Date"
         Height          =   195
         Left            =   360
         TabIndex        =   106
         Top             =   3000
         Width           =   975
      End
      Begin VB.OptionButton Option26 
         Caption         =   "Take Over"
         Height          =   195
         Left            =   1920
         TabIndex        =   122
         Top             =   3240
         Width           =   1335
      End
      Begin VB.OptionButton Option25 
         Caption         =   "Cost Details"
         Height          =   195
         Left            =   360
         TabIndex        =   101
         Top             =   1800
         Width           =   1455
      End
      Begin VB.OptionButton Option11 
         Caption         =   "Hosp. Grp"
         Height          =   195
         Left            =   1920
         TabIndex        =   110
         Top             =   360
         Width           =   1215
      End
      Begin VB.OptionButton Option24 
         Caption         =   "Treat. Mode"
         Height          =   195
         Left            =   1920
         TabIndex        =   123
         Top             =   3480
         Width           =   1215
      End
      Begin VB.OptionButton Option23 
         Caption         =   "Acc/Illness"
         Height          =   195
         Left            =   360
         TabIndex        =   95
         Top             =   360
         Width           =   1455
      End
      Begin VB.OptionButton Option22 
         Caption         =   "Claim Type"
         Height          =   195
         Left            =   360
         TabIndex        =   100
         Top             =   1560
         Width           =   1095
      End
      Begin VB.OptionButton Option21 
         Caption         =   "Discharge"
         Height          =   195
         Left            =   360
         TabIndex        =   103
         Top             =   2280
         Width           =   1095
      End
      Begin VB.OptionButton Option20 
         Caption         =   "Admission "
         ForeColor       =   &H000000FF&
         Height          =   195
         Left            =   360
         TabIndex        =   96
         Top             =   600
         Width           =   1095
      End
      Begin VB.OptionButton Option18 
         Caption         =   "Repudiate Code"
         Height          =   195
         Left            =   1920
         TabIndex        =   119
         Top             =   2520
         Width           =   1455
      End
      Begin VB.OptionButton Option17 
         Caption         =   "Repudiate Cat."
         Height          =   195
         Left            =   1920
         TabIndex        =   120
         Top             =   2760
         Width           =   1575
      End
      Begin VB.OptionButton Option16 
         Caption         =   "Duration Stay"
         Height          =   195
         Left            =   360
         TabIndex        =   105
         Top             =   2760
         Width           =   1335
      End
      Begin VB.OptionButton Option15 
         Caption         =   "Diagnosis"
         ForeColor       =   &H000000FF&
         Height          =   195
         Left            =   360
         TabIndex        =   102
         Top             =   2040
         Width           =   1455
      End
      Begin VB.OptionButton Option10 
         Caption         =   "Hospital"
         ForeColor       =   &H000000FF&
         Height          =   195
         Left            =   1920
         TabIndex        =   111
         Top             =   600
         Width           =   975
      End
      Begin VB.OptionButton Option9 
         Caption         =   "Payor"
         ForeColor       =   &H000000FF&
         Height          =   195
         Left            =   1920
         TabIndex        =   114
         Top             =   1320
         Width           =   855
      End
      Begin VB.OptionButton Option8 
         Caption         =   "State"
         ForeColor       =   &H000000FF&
         Height          =   195
         Left            =   1920
         TabIndex        =   121
         Top             =   3000
         Width           =   735
      End
      Begin VB.OptionButton Option7 
         Caption         =   "Race"
         ForeColor       =   &H000000FF&
         Height          =   195
         Left            =   1920
         TabIndex        =   117
         Top             =   2040
         Width           =   855
      End
      Begin VB.OptionButton Option6 
         Caption         =   "Gender"
         ForeColor       =   &H000000FF&
         Height          =   195
         Left            =   360
         TabIndex        =   108
         Top             =   3480
         Width           =   975
      End
      Begin VB.OptionButton Option5 
         Caption         =   "Age Band"
         ForeColor       =   &H000000FF&
         Height          =   195
         Left            =   360
         TabIndex        =   97
         Top             =   840
         Width           =   1095
      End
      Begin VB.OptionButton Option4 
         Caption         =   "Member Type"
         Height          =   195
         Left            =   1920
         TabIndex        =   113
         Top             =   1080
         Width           =   1335
      End
      Begin VB.OptionButton Option3 
         Caption         =   "Plan No."
         ForeColor       =   &H000000FF&
         Height          =   195
         Left            =   1920
         TabIndex        =   115
         Top             =   1560
         Width           =   975
      End
      Begin VB.OptionButton Option1 
         Caption         =   "General Listing"
         ForeColor       =   &H000000FF&
         Height          =   195
         Left            =   360
         TabIndex        =   109
         Top             =   3720
         Value           =   -1  'True
         Width           =   1455
      End
      Begin VB.OptionButton Option2 
         Caption         =   "Insured Type"
         ForeColor       =   &H000000FF&
         Height          =   195
         Left            =   1920
         TabIndex        =   112
         Top             =   840
         Width           =   1335
      End
   End
   Begin VB.Frame Frame8 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   315
      Left            =   0
      TabIndex        =   130
      Top             =   0
      Width           =   11805
      Begin VB.Label Label28 
         Appearance      =   0  'Flat
         BackColor       =   &H00000000&
         Caption         =   "Claims Reports"
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
         Left            =   120
         TabIndex        =   131
         Top             =   0
         Width           =   3285
      End
   End
   Begin VB.Frame Frame22 
      Caption         =   "Selection Criteria"
      Height          =   6045
      Left            =   0
      TabIndex        =   134
      Top             =   360
      Width           =   8145
      Begin VB.CheckBox ChkWksStatus 
         Caption         =   "Check3"
         Height          =   255
         Left            =   2640
         TabIndex        =   26
         Top             =   3500
         Width           =   255
      End
      Begin VB.CheckBox ChkCaseStatus 
         Caption         =   "Check3"
         Height          =   255
         Left            =   2640
         TabIndex        =   24
         Top             =   3210
         Width           =   255
      End
      Begin VB.CheckBox ChkPolStatus 
         Caption         =   "Check3"
         Height          =   255
         Left            =   2640
         TabIndex        =   22
         Top             =   2930
         Width           =   255
      End
      Begin VB.CheckBox ChkOccupation 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   280
         Left            =   2640
         TabIndex        =   28
         Top             =   3770
         Width           =   495
      End
      Begin VB.CheckBox ChkPreMode 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   280
         Left            =   2640
         TabIndex        =   30
         Top             =   4030
         Width           =   495
      End
      Begin VB.CheckBox ChkState 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   280
         Left            =   2640
         TabIndex        =   32
         Top             =   4320
         Width           =   255
      End
      Begin VB.CheckBox ChkSex 
         Caption         =   "Check3"
         Height          =   255
         Left            =   2640
         TabIndex        =   14
         Top             =   1800
         Width           =   255
      End
      Begin VB.CheckBox ChkRace 
         Caption         =   "Check3"
         Height          =   255
         Left            =   2640
         TabIndex        =   12
         Top             =   1560
         Width           =   255
      End
      Begin VB.CheckBox ChkMarital 
         Caption         =   "Check3"
         Height          =   255
         Left            =   2640
         TabIndex        =   10
         Top             =   1280
         Width           =   255
      End
      Begin VB.CheckBox ChkAdultChild 
         Caption         =   "Check3"
         Height          =   255
         Left            =   2640
         TabIndex        =   8
         Top             =   1000
         Width           =   255
      End
      Begin VB.CheckBox ChkMemType 
         Caption         =   "Check3"
         Height          =   255
         Left            =   2640
         TabIndex        =   6
         Top             =   750
         Width           =   255
      End
      Begin VB.CheckBox ChkINS 
         Caption         =   "Check3"
         Height          =   255
         Left            =   2640
         TabIndex        =   4
         Top             =   500
         Width           =   255
      End
      Begin VB.CheckBox ChkHLT 
         Caption         =   "Check3"
         Height          =   255
         Left            =   2640
         TabIndex        =   2
         Top             =   240
         Width           =   255
      End
      Begin VB.CheckBox ChkAgentCode 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   280
         Left            =   2640
         TabIndex        =   38
         Top             =   5160
         Width           =   255
      End
      Begin VB.CheckBox ChkTakeOver 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   280
         Left            =   2640
         TabIndex        =   36
         Top             =   4880
         Width           =   375
      End
      Begin VB.CheckBox ChkRelation 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   280
         Left            =   2640
         TabIndex        =   16
         Top             =   2100
         Width           =   495
      End
      Begin VB.CheckBox ChkDataStatus 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   280
         Left            =   2640
         TabIndex        =   18
         Top             =   2390
         Width           =   495
      End
      Begin VB.CheckBox ChkRenewal 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   280
         Left            =   2640
         TabIndex        =   34
         Top             =   4590
         Width           =   615
      End
      Begin VB.CheckBox ChkClaimability 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   280
         Left            =   2640
         TabIndex        =   20
         Top             =   2640
         Width           =   495
      End
      Begin VB.CheckBox ChkPolYear 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   280
         Left            =   7680
         TabIndex        =   94
         Top             =   5640
         Width           =   255
      End
      Begin VB.CheckBox ChkRepCat 
         Caption         =   "Check3"
         Height          =   255
         Left            =   7680
         TabIndex        =   73
         Top             =   3720
         Width           =   255
      End
      Begin VB.CheckBox ChkRep 
         Caption         =   "Check3"
         Height          =   255
         Left            =   7680
         TabIndex        =   70
         Top             =   3480
         Width           =   255
      End
      Begin VB.CheckBox ChkFinal 
         Caption         =   "Check3"
         Height          =   255
         Left            =   7680
         TabIndex        =   67
         Top             =   3120
         Width           =   255
      End
      Begin VB.CheckBox ChkPlan 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   280
         Left            =   7680
         TabIndex        =   76
         Top             =   3960
         Width           =   375
      End
      Begin VB.CheckBox ChkBordx 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   280
         Left            =   7680
         TabIndex        =   79
         Top             =   4250
         Width           =   375
      End
      Begin VB.CheckBox ChkEffDate 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   280
         Left            =   7680
         TabIndex        =   82
         Top             =   4560
         Width           =   375
      End
      Begin VB.CheckBox ChkFirst 
         Caption         =   "Check3"
         Height          =   255
         Left            =   7680
         TabIndex        =   64
         Top             =   2880
         Width           =   255
      End
      Begin VB.CheckBox ChkDoctor 
         Caption         =   "Check3"
         Height          =   255
         Left            =   7680
         TabIndex        =   51
         Top             =   1560
         Width           =   255
      End
      Begin VB.CheckBox ChkHOSGrp 
         Caption         =   "Check3"
         Height          =   255
         Left            =   7680
         TabIndex        =   49
         Top             =   1320
         Width           =   255
      End
      Begin VB.CheckBox ChkHOS 
         Caption         =   "Check3"
         Height          =   255
         Left            =   7680
         TabIndex        =   47
         Top             =   960
         Width           =   255
      End
      Begin VB.CheckBox ChkPolNo 
         Caption         =   "Check3"
         Height          =   255
         Left            =   2640
         TabIndex        =   40
         Top             =   5430
         Width           =   255
      End
      Begin VB.CheckBox ChkGrpCom 
         Caption         =   "Check3"
         Height          =   255
         Left            =   7680
         TabIndex        =   44
         Top             =   480
         Width           =   255
      End
      Begin VB.CheckBox ChkPayor 
         Caption         =   "Check3"
         Height          =   255
         Left            =   7680
         TabIndex        =   42
         Top             =   240
         Width           =   255
      End
      Begin VB.CheckBox ChkDOB 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   280
         Left            =   7680
         TabIndex        =   91
         Top             =   5400
         Width           =   255
      End
      Begin VB.CheckBox ChkJoinDate 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   280
         Left            =   7680
         TabIndex        =   88
         Top             =   5160
         Width           =   375
      End
      Begin VB.TextBox TxtPlan2 
         Height          =   315
         Left            =   6240
         MaxLength       =   4
         TabIndex        =   75
         Top             =   3950
         Width           =   1200
      End
      Begin VB.TextBox TxtPlan1 
         Height          =   315
         Left            =   4320
         MaxLength       =   4
         TabIndex        =   74
         Top             =   3950
         Width           =   1200
      End
      Begin VB.CheckBox chkDOR 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   280
         Left            =   7680
         TabIndex        =   61
         Top             =   2520
         Width           =   255
      End
      Begin VB.CheckBox chkExpDate 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   280
         Left            =   7680
         TabIndex        =   85
         Top             =   4845
         Width           =   375
      End
      Begin VB.CheckBox ChkDOD 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   280
         Left            =   7680
         TabIndex        =   58
         Top             =   2280
         Width           =   255
      End
      Begin VB.CheckBox ChkDOA 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   280
         Left            =   7680
         TabIndex        =   55
         Top             =   2040
         Width           =   255
      End
      Begin VB.ComboBox cboHLTCode 
         Height          =   315
         Left            =   1200
         Sorted          =   -1  'True
         TabIndex        =   1
         Text            =   "S - Sihat"
         Top             =   200
         Width           =   1200
      End
      Begin VB.ComboBox CboInsType 
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
         Sorted          =   -1  'True
         TabIndex        =   3
         Top             =   480
         Width           =   1200
      End
      Begin VB.ComboBox CboMemType 
         Height          =   315
         Left            =   1200
         Style           =   2  'Dropdown List
         TabIndex        =   5
         Top             =   720
         Width           =   1200
      End
      Begin VB.ComboBox CboAdultChild 
         Height          =   315
         Left            =   1200
         TabIndex        =   7
         Top             =   960
         Width           =   1200
      End
      Begin VB.ComboBox CboMarital 
         Height          =   315
         Left            =   1200
         TabIndex        =   9
         Top             =   1250
         Width           =   1200
      End
      Begin VB.ComboBox CboRace 
         Height          =   315
         Left            =   1200
         TabIndex        =   11
         Top             =   1530
         Width           =   1200
      End
      Begin VB.ComboBox CboSex 
         Height          =   315
         Left            =   1200
         TabIndex        =   13
         Top             =   1800
         Width           =   1200
      End
      Begin VB.ComboBox cboPayorCode 
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
         Left            =   4320
         Sorted          =   -1  'True
         TabIndex        =   41
         Top             =   200
         Width           =   3120
      End
      Begin VB.ComboBox CboGrpCom 
         Height          =   315
         Left            =   4320
         TabIndex        =   43
         Top             =   480
         Width           =   3120
      End
      Begin VB.ComboBox cboSuppType 
         Height          =   315
         Left            =   4320
         Style           =   2  'Dropdown List
         TabIndex        =   45
         Top             =   720
         Width           =   3120
      End
      Begin VB.ComboBox cboHOSName 
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
         Left            =   4320
         Sorted          =   -1  'True
         TabIndex        =   46
         Top             =   960
         Width           =   3120
      End
      Begin VB.ComboBox cboHOSGrp 
         Height          =   315
         Left            =   4320
         TabIndex        =   48
         Top             =   1250
         Width           =   3120
      End
      Begin VB.TextBox txtDOCName 
         Height          =   285
         Left            =   4320
         MaxLength       =   100
         TabIndex        =   50
         Top             =   1530
         Width           =   3120
      End
      Begin VB.TextBox txtRepudCat2 
         Height          =   285
         Left            =   6240
         MaxLength       =   4
         TabIndex        =   69
         Top             =   3400
         Width           =   1200
      End
      Begin VB.TextBox TxtPatientName 
         Height          =   285
         Left            =   4320
         MaxLength       =   60
         TabIndex        =   52
         Top             =   1800
         Width           =   3120
      End
      Begin MSMask.MaskEdBox TxtDOA1 
         Height          =   315
         Left            =   4320
         TabIndex        =   53
         Top             =   2040
         Width           =   1200
         _ExtentX        =   2117
         _ExtentY        =   556
         _Version        =   393216
         MaxLength       =   10
         Mask            =   "##/##/####"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox TxtDOA2 
         Height          =   315
         Left            =   6240
         TabIndex        =   54
         Top             =   2040
         Width           =   1200
         _ExtentX        =   2117
         _ExtentY        =   556
         _Version        =   393216
         MaxLength       =   10
         Mask            =   "##/##/####"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox txtDOD1 
         Height          =   315
         Left            =   4320
         TabIndex        =   56
         Top             =   2310
         Width           =   1200
         _ExtentX        =   2117
         _ExtentY        =   556
         _Version        =   393216
         MaxLength       =   10
         Mask            =   "##/##/####"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox txtDOD2 
         Height          =   315
         Left            =   6240
         TabIndex        =   57
         Top             =   2310
         Width           =   1200
         _ExtentX        =   2117
         _ExtentY        =   556
         _Version        =   393216
         MaxLength       =   10
         Mask            =   "##/##/####"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox TxtDOR1 
         Height          =   315
         Left            =   4320
         TabIndex        =   59
         Top             =   2595
         Width           =   1200
         _ExtentX        =   2117
         _ExtentY        =   556
         _Version        =   393216
         MaxLength       =   10
         Mask            =   "##/##/####"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox TxtDOR2 
         Height          =   315
         Left            =   6240
         TabIndex        =   60
         Top             =   2595
         Width           =   1200
         _ExtentX        =   2117
         _ExtentY        =   556
         _Version        =   393216
         MaxLength       =   10
         Mask            =   "##/##/####"
         PromptChar      =   "_"
      End
      Begin VB.TextBox txtPreDiag1 
         Height          =   285
         Left            =   4320
         MaxLength       =   3
         TabIndex        =   62
         Top             =   2880
         Width           =   1200
      End
      Begin VB.TextBox txtPreDiag2 
         Height          =   285
         Left            =   6240
         MaxLength       =   3
         TabIndex        =   63
         Top             =   2880
         Width           =   1200
      End
      Begin VB.TextBox txtFinalDiag1 
         Height          =   285
         Left            =   4320
         MaxLength       =   3
         TabIndex        =   65
         Top             =   3120
         Width           =   1200
      End
      Begin VB.TextBox txtFinalDiag2 
         Height          =   285
         Left            =   6240
         MaxLength       =   3
         TabIndex        =   66
         Top             =   3135
         Width           =   1200
      End
      Begin VB.ComboBox CboRelation 
         Height          =   315
         Left            =   1200
         TabIndex        =   15
         Top             =   2070
         Width           =   1200
      End
      Begin VB.ComboBox cboDataStatus 
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
         TabIndex        =   17
         Top             =   2360
         Width           =   1200
      End
      Begin VB.ComboBox CboClaimability 
         Height          =   315
         Left            =   1200
         Sorted          =   -1  'True
         TabIndex        =   19
         Top             =   2640
         Width           =   1200
      End
      Begin VB.ComboBox CboStatus 
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
         TabIndex        =   21
         Top             =   2900
         Width           =   1200
      End
      Begin VB.ComboBox CboCaseStatus 
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
         Sorted          =   -1  'True
         TabIndex        =   23
         Top             =   3180
         Width           =   1200
      End
      Begin VB.ComboBox CboWsStatus 
         Height          =   315
         Left            =   1200
         TabIndex        =   25
         Top             =   3480
         Width           =   1200
      End
      Begin VB.ComboBox CboOccupation 
         Height          =   315
         Left            =   1200
         TabIndex        =   27
         Top             =   3740
         Width           =   1200
      End
      Begin VB.ComboBox CboPreMode 
         Height          =   315
         Left            =   1200
         TabIndex        =   29
         Top             =   4000
         Width           =   1200
      End
      Begin VB.ComboBox CboState 
         Height          =   315
         Left            =   1200
         TabIndex        =   31
         Top             =   4290
         Width           =   1200
      End
      Begin VB.ComboBox cboRenewalInd 
         Height          =   315
         Left            =   1200
         TabIndex        =   33
         Top             =   4560
         Width           =   1200
      End
      Begin VB.ComboBox cboTakeOver 
         Height          =   315
         Left            =   1200
         TabIndex        =   35
         Top             =   4850
         Width           =   1200
      End
      Begin VB.TextBox txtAgentCode 
         Height          =   285
         Left            =   1200
         MaxLength       =   15
         TabIndex        =   37
         Top             =   5130
         Width           =   1200
      End
      Begin VB.TextBox txtPolNo 
         Height          =   285
         Left            =   1200
         MaxLength       =   25
         TabIndex        =   39
         Top             =   5400
         Width           =   1200
      End
      Begin VB.TextBox txtRepudCat1 
         Height          =   285
         Left            =   4320
         MaxLength       =   4
         TabIndex        =   68
         Top             =   3400
         Width           =   1200
      End
      Begin VB.TextBox txtRepudCode1 
         Height          =   285
         Left            =   4320
         MaxLength       =   10
         TabIndex        =   71
         Top             =   3680
         Width           =   1200
      End
      Begin VB.TextBox txtRepudCode2 
         Height          =   285
         Left            =   6240
         MaxLength       =   10
         TabIndex        =   72
         Top             =   3680
         Width           =   1200
      End
      Begin VB.TextBox TxtBordxNo1 
         Height          =   315
         Left            =   4320
         MaxLength       =   10
         TabIndex        =   77
         Top             =   4245
         Width           =   1200
      End
      Begin VB.TextBox TxtBordxNo2 
         Height          =   315
         Left            =   6240
         MaxLength       =   10
         TabIndex        =   78
         Top             =   4245
         Width           =   1200
      End
      Begin MSMask.MaskEdBox TxtPayEffDate1 
         Height          =   315
         Left            =   4320
         TabIndex        =   80
         Top             =   4485
         Width           =   1200
         _ExtentX        =   2117
         _ExtentY        =   556
         _Version        =   393216
         MaxLength       =   10
         Mask            =   "##/##/####"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox TxtPayEffDate2 
         Height          =   315
         Left            =   6240
         TabIndex        =   81
         Top             =   4485
         Width           =   1200
         _ExtentX        =   2117
         _ExtentY        =   556
         _Version        =   393216
         MaxLength       =   10
         Mask            =   "##/##/####"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox TxtPayExpDate1 
         Height          =   315
         Left            =   4320
         TabIndex        =   83
         Top             =   4780
         Width           =   1200
         _ExtentX        =   2117
         _ExtentY        =   556
         _Version        =   393216
         MaxLength       =   10
         Mask            =   "##/##/####"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox TxtPayExpDate2 
         Height          =   315
         Left            =   6240
         TabIndex        =   84
         Top             =   4780
         Width           =   1200
         _ExtentX        =   2117
         _ExtentY        =   556
         _Version        =   393216
         MaxLength       =   10
         Mask            =   "##/##/####"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox txtDateJoin1 
         Height          =   315
         Left            =   4320
         TabIndex        =   86
         Top             =   5085
         Width           =   1200
         _ExtentX        =   2117
         _ExtentY        =   556
         _Version        =   393216
         MaxLength       =   10
         Mask            =   "##/##/####"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox txtDateJoin2 
         Height          =   315
         Left            =   6240
         TabIndex        =   87
         Top             =   5085
         Width           =   1200
         _ExtentX        =   2117
         _ExtentY        =   556
         _Version        =   393216
         MaxLength       =   10
         Mask            =   "##/##/####"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox txtDOB1 
         Height          =   315
         Left            =   4320
         TabIndex        =   89
         Top             =   5370
         Width           =   1200
         _ExtentX        =   2117
         _ExtentY        =   556
         _Version        =   393216
         MaxLength       =   10
         Mask            =   "##/##/####"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox txtDOB2 
         Height          =   315
         Left            =   6240
         TabIndex        =   90
         Top             =   5370
         Width           =   1200
         _ExtentX        =   2117
         _ExtentY        =   556
         _Version        =   393216
         MaxLength       =   10
         Mask            =   "##/##/####"
         PromptChar      =   "_"
      End
      Begin VB.TextBox txtPOLYear1 
         Height          =   285
         Left            =   4320
         MaxLength       =   2
         TabIndex        =   92
         Top             =   5640
         Width           =   1200
      End
      Begin VB.TextBox txtPOLYear2 
         Height          =   285
         Left            =   6240
         MaxLength       =   2
         TabIndex        =   93
         Top             =   5640
         Width           =   1200
      End
      Begin VB.TextBox TxtYear 
         Height          =   315
         Left            =   4320
         MaxLength       =   4
         TabIndex        =   203
         Top             =   2040
         Width           =   1200
      End
      Begin VB.Label Label62 
         Caption         =   "Patient Name"
         Height          =   255
         Left            =   3240
         TabIndex        =   197
         Top             =   1830
         Width           =   1095
      End
      Begin VB.Label Label61 
         Caption         =   "To"
         Height          =   255
         Left            =   5880
         TabIndex        =   195
         Top             =   4320
         Width           =   375
      End
      Begin VB.Label Label60 
         Caption         =   "Bordx No"
         Height          =   255
         Left            =   3240
         TabIndex        =   194
         Top             =   4320
         Width           =   855
      End
      Begin VB.Label Label59 
         Caption         =   "*"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   255
         Left            =   2640
         TabIndex        =   190
         Top             =   120
         Width           =   255
      End
      Begin VB.Label Label58 
         Caption         =   "*"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   255
         Left            =   7680
         TabIndex        =   189
         Top             =   120
         Width           =   255
      End
      Begin VB.Label Label57 
         Caption         =   "To"
         Height          =   255
         Left            =   5880
         TabIndex        =   188
         Top             =   5760
         Width           =   495
      End
      Begin VB.Label Label56 
         Caption         =   "Pol. Year"
         Height          =   255
         Left            =   3240
         TabIndex        =   187
         Top             =   5640
         Width           =   1095
      End
      Begin VB.Label Label55 
         Caption         =   "Agent Code"
         Height          =   255
         Left            =   120
         TabIndex        =   186
         Top             =   5160
         Width           =   1095
      End
      Begin VB.Label Label45 
         Caption         =   "Take Over"
         Height          =   255
         Left            =   120
         TabIndex        =   185
         Top             =   4920
         Width           =   1095
      End
      Begin VB.Label Label44 
         Caption         =   "Renewal "
         Height          =   255
         Left            =   120
         TabIndex        =   184
         Top             =   4680
         Width           =   1095
      End
      Begin VB.Label Label43 
         Caption         =   "To"
         Height          =   255
         Left            =   5880
         TabIndex        =   183
         Top             =   5460
         Width           =   495
      End
      Begin VB.Label Label42 
         Caption         =   "DOB"
         Height          =   255
         Left            =   3240
         TabIndex        =   182
         ToolTipText     =   "Payor Expiry Date"
         Top             =   5460
         Width           =   1215
      End
      Begin VB.Label Label41 
         Caption         =   "To"
         Height          =   255
         Left            =   5880
         TabIndex        =   181
         Top             =   5220
         Width           =   495
      End
      Begin VB.Label Label40 
         Caption         =   "Date Join"
         Height          =   255
         Left            =   3240
         TabIndex        =   180
         ToolTipText     =   "Payor Expiry Date"
         Top             =   5220
         Width           =   1215
      End
      Begin VB.Label Label39 
         Caption         =   "Doctor"
         Height          =   255
         Left            =   3240
         TabIndex        =   179
         Top             =   1560
         Width           =   1095
      End
      Begin VB.Label Label37 
         Caption         =   "Hosp. Grp"
         Height          =   255
         Left            =   3240
         TabIndex        =   178
         Top             =   1320
         Width           =   1095
      End
      Begin VB.Label Label36 
         Caption         =   "Hospital"
         Height          =   255
         Left            =   3240
         TabIndex        =   177
         Top             =   1080
         Width           =   975
      End
      Begin VB.Label Label35 
         Caption         =   "To"
         Height          =   255
         Left            =   5880
         TabIndex        =   176
         Top             =   3720
         Width           =   495
      End
      Begin VB.Label Label34 
         Caption         =   "To"
         Height          =   255
         Left            =   5880
         TabIndex        =   175
         Top             =   3480
         Width           =   495
      End
      Begin VB.Label Label33 
         Caption         =   "To"
         Height          =   255
         Left            =   5880
         TabIndex        =   174
         Top             =   3120
         Width           =   495
      End
      Begin VB.Label Label32 
         Caption         =   "To"
         Height          =   255
         Left            =   5880
         TabIndex        =   173
         Top             =   2880
         Width           =   495
      End
      Begin VB.Label Label14 
         Caption         =   "Repud Code"
         Height          =   255
         Left            =   3240
         TabIndex        =   172
         Top             =   3840
         Width           =   975
      End
      Begin VB.Label Label13 
         Caption         =   "Repud Cat"
         Height          =   255
         Left            =   3240
         TabIndex        =   171
         Top             =   3480
         Width           =   975
      End
      Begin VB.Label Label12 
         Caption         =   "Final Diag"
         Height          =   255
         Left            =   3240
         TabIndex        =   170
         Top             =   3240
         Width           =   855
      End
      Begin VB.Label Label11 
         Caption         =   "Pre. Diag"
         Height          =   255
         Left            =   3240
         TabIndex        =   169
         Top             =   2880
         Width           =   855
      End
      Begin VB.Label Label6 
         Caption         =   "Grp Company"
         Height          =   255
         Left            =   3240
         TabIndex        =   168
         Top             =   540
         Width           =   1095
      End
      Begin VB.Label Label54 
         Caption         =   "Case Status"
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
         TabIndex        =   167
         Top             =   3240
         Width           =   945
      End
      Begin VB.Label Label53 
         Caption         =   "Claimability"
         Height          =   255
         Left            =   120
         TabIndex        =   166
         Top             =   2640
         Width           =   975
      End
      Begin VB.Label Label52 
         Caption         =   "Policy Status"
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
         TabIndex        =   165
         Top             =   3000
         Width           =   1020
      End
      Begin VB.Label Label51 
         Caption         =   "Plan Code"
         Height          =   255
         Left            =   3240
         TabIndex        =   164
         Top             =   4080
         Width           =   855
      End
      Begin VB.Label Label46 
         Caption         =   "To"
         Height          =   255
         Left            =   5880
         TabIndex        =   163
         Top             =   4080
         Width           =   375
      End
      Begin VB.Label Label72 
         Caption         =   "Data Status"
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
         TabIndex        =   162
         Top             =   2400
         Width           =   1020
      End
      Begin VB.Label Label38 
         Caption         =   "Payor"
         Height          =   225
         Left            =   3240
         TabIndex        =   161
         Top             =   240
         Width           =   975
      End
      Begin VB.Label Label26 
         Caption         =   "INS Type"
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
         TabIndex        =   160
         Top             =   480
         Width           =   690
      End
      Begin VB.Label Label73 
         Caption         =   "HLT Type"
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
         TabIndex        =   159
         Top             =   225
         Width           =   705
      End
      Begin VB.Label Label7 
         Caption         =   "Eff Date"
         Height          =   255
         Left            =   3240
         TabIndex        =   158
         ToolTipText     =   "Payor Effective Date"
         Top             =   4560
         Width           =   1215
      End
      Begin VB.Label Label23 
         Caption         =   "Exp Date"
         Height          =   255
         Left            =   3240
         TabIndex        =   157
         ToolTipText     =   "Payor Expiry Date"
         Top             =   4905
         Width           =   1215
      End
      Begin VB.Label Label21 
         Caption         =   "DOR"
         Height          =   255
         Left            =   3240
         TabIndex        =   156
         ToolTipText     =   "Member Value Date"
         Top             =   2640
         Width           =   855
      End
      Begin VB.Label Label8 
         Caption         =   "To"
         Height          =   255
         Left            =   5880
         TabIndex        =   155
         Top             =   4560
         Width           =   495
      End
      Begin VB.Label Label9 
         Caption         =   "To"
         Height          =   255
         Left            =   5880
         TabIndex        =   154
         Top             =   4905
         Width           =   495
      End
      Begin VB.Label Label16 
         Caption         =   "To"
         Height          =   255
         Left            =   5880
         TabIndex        =   153
         Top             =   2640
         Width           =   495
      End
      Begin VB.Label Label47 
         Caption         =   "DOA"
         Height          =   255
         Left            =   3240
         TabIndex        =   152
         ToolTipText     =   "Date of Admission"
         Top             =   2160
         Width           =   735
      End
      Begin VB.Label Label48 
         Caption         =   "DOD"
         Height          =   255
         Left            =   3240
         TabIndex        =   151
         ToolTipText     =   "Date of Discharged"
         Top             =   2400
         Width           =   615
      End
      Begin VB.Label Label49 
         Alignment       =   2  'Center
         Caption         =   "To"
         Height          =   255
         Left            =   5760
         TabIndex        =   150
         Top             =   2160
         Width           =   375
      End
      Begin VB.Label Label50 
         Alignment       =   2  'Center
         Caption         =   "To"
         Height          =   255
         Left            =   5760
         TabIndex        =   149
         Top             =   2400
         Width           =   375
      End
      Begin VB.Label Label1 
         Caption         =   "Sex"
         Height          =   255
         Left            =   120
         TabIndex        =   148
         Top             =   1920
         Width           =   855
      End
      Begin VB.Label Label4 
         Caption         =   "Mem Type"
         Height          =   255
         Left            =   120
         TabIndex        =   147
         Top             =   840
         Width           =   855
      End
      Begin VB.Label Label5 
         Caption         =   "Adult / Child"
         Height          =   255
         Left            =   120
         TabIndex        =   146
         Top             =   1080
         Width           =   975
      End
      Begin VB.Label Label15 
         Caption         =   "Marital Status"
         Height          =   255
         Left            =   120
         TabIndex        =   145
         Top             =   1365
         Width           =   1095
      End
      Begin VB.Label Label17 
         Caption         =   "Race"
         Height          =   255
         Left            =   120
         TabIndex        =   144
         Top             =   1665
         Width           =   855
      End
      Begin VB.Label Label18 
         Caption         =   "Relationship"
         Height          =   255
         Left            =   120
         TabIndex        =   143
         Top             =   2160
         Width           =   975
      End
      Begin VB.Label Label19 
         Height          =   255
         Left            =   2520
         TabIndex        =   142
         Top             =   240
         Width           =   975
      End
      Begin VB.Label Label20 
         Caption         =   "Data Status"
         Height          =   255
         Left            =   2520
         TabIndex        =   141
         Top             =   260
         Width           =   975
      End
      Begin VB.Label Label22 
         Caption         =   "WS Status"
         Height          =   255
         Left            =   120
         TabIndex        =   140
         Top             =   3600
         Width           =   975
      End
      Begin VB.Label Label25 
         Caption         =   "Occuptation"
         Height          =   255
         Left            =   120
         TabIndex        =   139
         Top             =   3840
         Width           =   975
      End
      Begin VB.Label Label27 
         Caption         =   "Prem. Mode"
         Height          =   255
         Left            =   120
         TabIndex        =   138
         Top             =   4080
         Width           =   975
      End
      Begin VB.Label Label29 
         Caption         =   "State"
         Height          =   255
         Left            =   120
         TabIndex        =   137
         Top             =   4320
         Width           =   855
      End
      Begin VB.Label Label30 
         Caption         =   "Policy No."
         Height          =   255
         Left            =   120
         TabIndex        =   136
         Top             =   5400
         Width           =   975
      End
      Begin VB.Label Label31 
         Caption         =   "Supp. Type"
         Height          =   255
         Left            =   3240
         TabIndex        =   135
         Top             =   840
         Width           =   1095
      End
   End
   Begin MSComctlLib.ListView ListView2 
      Height          =   855
      Left            =   0
      TabIndex        =   196
      Top             =   6480
      Width           =   11775
      _ExtentX        =   20770
      _ExtentY        =   1508
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   63
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Object.Width           =   1764
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Object.Width           =   1764
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   2
         Object.Width           =   1764
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   3
         Object.Width           =   1764
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   4
         Object.Width           =   1764
      EndProperty
      BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   5
         Object.Width           =   1764
      EndProperty
      BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   6
         Object.Width           =   1764
      EndProperty
      BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   7
         Object.Width           =   1764
      EndProperty
      BeginProperty ColumnHeader(9) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   8
         Object.Width           =   1764
      EndProperty
      BeginProperty ColumnHeader(10) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   9
         Object.Width           =   1764
      EndProperty
      BeginProperty ColumnHeader(11) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   10
         Object.Width           =   1764
      EndProperty
      BeginProperty ColumnHeader(12) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   11
         Object.Width           =   1764
      EndProperty
      BeginProperty ColumnHeader(13) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   12
         Object.Width           =   1764
      EndProperty
      BeginProperty ColumnHeader(14) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   13
         Object.Width           =   1764
      EndProperty
      BeginProperty ColumnHeader(15) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   14
         Object.Width           =   1764
      EndProperty
      BeginProperty ColumnHeader(16) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   15
         Object.Width           =   1764
      EndProperty
      BeginProperty ColumnHeader(17) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   16
         Object.Width           =   1764
      EndProperty
      BeginProperty ColumnHeader(18) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   17
         Object.Width           =   1764
      EndProperty
      BeginProperty ColumnHeader(19) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   18
         Object.Width           =   1764
      EndProperty
      BeginProperty ColumnHeader(20) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   19
         Object.Width           =   1764
      EndProperty
      BeginProperty ColumnHeader(21) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   20
         Object.Width           =   1764
      EndProperty
      BeginProperty ColumnHeader(22) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   21
         Object.Width           =   1764
      EndProperty
      BeginProperty ColumnHeader(23) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   22
         Object.Width           =   1764
      EndProperty
      BeginProperty ColumnHeader(24) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   23
         Object.Width           =   1764
      EndProperty
      BeginProperty ColumnHeader(25) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   24
         Object.Width           =   1764
      EndProperty
      BeginProperty ColumnHeader(26) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   25
         Object.Width           =   1764
      EndProperty
      BeginProperty ColumnHeader(27) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   26
         Object.Width           =   1764
      EndProperty
      BeginProperty ColumnHeader(28) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   27
         Object.Width           =   1764
      EndProperty
      BeginProperty ColumnHeader(29) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   28
         Object.Width           =   1764
      EndProperty
      BeginProperty ColumnHeader(30) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   29
         Object.Width           =   1764
      EndProperty
      BeginProperty ColumnHeader(31) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   30
         Object.Width           =   1764
      EndProperty
      BeginProperty ColumnHeader(32) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   31
         Object.Width           =   1764
      EndProperty
      BeginProperty ColumnHeader(33) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   32
         Object.Width           =   1764
      EndProperty
      BeginProperty ColumnHeader(34) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   33
         Object.Width           =   1764
      EndProperty
      BeginProperty ColumnHeader(35) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   34
         Object.Width           =   1764
      EndProperty
      BeginProperty ColumnHeader(36) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   35
         Object.Width           =   1764
      EndProperty
      BeginProperty ColumnHeader(37) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   36
         Object.Width           =   1764
      EndProperty
      BeginProperty ColumnHeader(38) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   37
         Object.Width           =   1764
      EndProperty
      BeginProperty ColumnHeader(39) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   38
         Object.Width           =   1764
      EndProperty
      BeginProperty ColumnHeader(40) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   39
         Object.Width           =   1764
      EndProperty
      BeginProperty ColumnHeader(41) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   40
         Object.Width           =   1764
      EndProperty
      BeginProperty ColumnHeader(42) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   41
         Object.Width           =   1764
      EndProperty
      BeginProperty ColumnHeader(43) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   42
         Object.Width           =   1764
      EndProperty
      BeginProperty ColumnHeader(44) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   43
         Object.Width           =   1764
      EndProperty
      BeginProperty ColumnHeader(45) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   44
         Object.Width           =   1764
      EndProperty
      BeginProperty ColumnHeader(46) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   45
         Object.Width           =   1764
      EndProperty
      BeginProperty ColumnHeader(47) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   46
         Object.Width           =   1764
      EndProperty
      BeginProperty ColumnHeader(48) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   47
         Object.Width           =   1764
      EndProperty
      BeginProperty ColumnHeader(49) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   48
         Object.Width           =   1764
      EndProperty
      BeginProperty ColumnHeader(50) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   49
         Object.Width           =   1764
      EndProperty
      BeginProperty ColumnHeader(51) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   50
         Object.Width           =   1764
      EndProperty
      BeginProperty ColumnHeader(52) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   51
         Object.Width           =   1764
      EndProperty
      BeginProperty ColumnHeader(53) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   52
         Object.Width           =   1764
      EndProperty
      BeginProperty ColumnHeader(54) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   53
         Object.Width           =   1764
      EndProperty
      BeginProperty ColumnHeader(55) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   54
         Object.Width           =   1764
      EndProperty
      BeginProperty ColumnHeader(56) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   55
         Object.Width           =   1764
      EndProperty
      BeginProperty ColumnHeader(57) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   56
         Object.Width           =   1764
      EndProperty
      BeginProperty ColumnHeader(58) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   57
         Object.Width           =   1764
      EndProperty
      BeginProperty ColumnHeader(59) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   58
         Object.Width           =   1764
      EndProperty
      BeginProperty ColumnHeader(60) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   59
         Object.Width           =   1764
      EndProperty
      BeginProperty ColumnHeader(61) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   60
         Object.Width           =   1764
      EndProperty
      BeginProperty ColumnHeader(62) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   61
         Object.Width           =   1764
      EndProperty
      BeginProperty ColumnHeader(63) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   62
         Object.Width           =   1764
      EndProperty
   End
   Begin VB.Frame Frame6 
      Caption         =   "Reports Selection"
      Height          =   1935
      Left            =   8160
      TabIndex        =   192
      Top             =   4440
      Width           =   3615
      Begin VB.OptionButton Option30 
         Caption         =   "By Nature of Claims"
         Height          =   195
         Left            =   360
         TabIndex        =   210
         Top             =   1560
         Visible         =   0   'False
         Width           =   2055
      End
      Begin VB.OptionButton Option37 
         Caption         =   "By Plan No"
         Height          =   195
         Left            =   360
         TabIndex        =   209
         Top             =   1320
         Visible         =   0   'False
         Width           =   2055
      End
      Begin VB.OptionButton Option35 
         Caption         =   "Detail Listing"
         Height          =   195
         Left            =   360
         TabIndex        =   202
         Top             =   360
         Visible         =   0   'False
         Width           =   1815
      End
      Begin VB.OptionButton Option34 
         Caption         =   "General Summary"
         Height          =   195
         Left            =   360
         TabIndex        =   201
         Top             =   600
         Visible         =   0   'False
         Width           =   1695
      End
      Begin VB.OptionButton Option33 
         Caption         =   "Top 20 by Cases"
         Height          =   195
         Left            =   360
         TabIndex        =   200
         Top             =   840
         Visible         =   0   'False
         Width           =   1935
      End
      Begin VB.OptionButton Option29 
         Caption         =   "Top 20 by Amount"
         Height          =   195
         Left            =   360
         TabIndex        =   199
         Top             =   1080
         Visible         =   0   'False
         Width           =   2055
      End
   End
End
Attribute VB_Name = "frmMBMAnalysisRpt"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private DiagCode, PostCode, InsureCode As String
Private HOSCode, HOSGRPCode As String
Private TotalDay, TotalDay2 As String
Private DurationEffDate, DurationJoinedDate As String
Private Description, StateCode, SexCode, RaceCode, PlanCode, INSCode, Payor, Day As String
Private TotalAgeCode1, TotalAgeCode2 As String
Private TotalAgeCode3, TotalAgeCode4 As String
Private TotalAgeCode5, TotalAgeCode6 As String
Private TotalAgeCode7 As String
Private TotalRace1, TotalRace2, TotalRace3, TotalRace4, TotalRace5 As String
Private TotalSex1, TotalSex2, TotalSex3 As String
Private TotalActual, TotalDuration, TotalApproved, Total, TotalEmpty As String
Private strAppend As String
Private TotalM, TotalF, TotalCases As String
Private TotalM1, TotalM2, TotalM3, TotalM4, TotalM5, TotalM6, TotalM7 As String
Private TotalF1, TotalF2, TotalF3, TotalF4, TotalF5, TotalF6, TotalF7 As String
Private AdmissionDate1, AdmissionDate2, AdmissionDate3 As String
Private AdmissionDate4, AdmissionDate5, AdmissionDate6 As String
Private AdmissionDate7, AdmissionDate8, AdmissionDate9 As String
Private AdmissionDate10, AdmissionDate11, AdmissionDate12 As String
Private AdmissionDate13, AdmissionDate14, AdmissionDate15 As String
Private AdmissionDate16, AdmissionDate17, AdmissionDate18 As String
Private AdmissionDate19, AdmissionDate20, AdmissionDate21 As String
Private AdmissionDate22, AdmissionDate23, AdmissionDate24 As String
Private AdmissionDate25, AdmissionDate26, AdmissionDate27 As String
Private AdmissionDate28, AdmissionDate29, AdmissionDate30 As String
Private AdmissionDate31 As String
Private Day1, Day2, Day3, Day4, Day5, Day6, Day7 As String
Private Day8, Day9, Day10, Day11, Day12, Day13, Day14 As String
Private Day15, Day16, Day17, Day18, Day19, Day20, Day21 As String
Private Day22, Day23, Day24, Day25, Day26, Day27, Day28 As String
Private Day29, Day30, Day31 As String
Private Total1, Total2, Total3, Total4, Total5, Total6, Total7 As String
Private Total8, Total9, Total10, Total11, Total12, Total13, Total14 As String
Private Total15, Total16, Total17, Total18, Total19, Total20, Total21 As String
Private Total22, Total23, Total24, Total25, Total26, Total27, Total28 As String
Private Total29, Total30, Total31 As String
Private intExit As Integer
Private Current As String
Private m_blnDirection(1 To 63) As Boolean


Private Sub cboHLTCode_Change()
    If InStr(cboHLTCode.Text, "null") Then
       cboHLTCode.Enabled = False
    End If
    
    If Len(Trim(cboHLTCode)) Then
        LoadPayor (cboHLTCode)
    End If
End Sub

Private Sub cboHLTCode_Click()
    If Len(Trim(cboHLTCode)) Then
        LoadPayor (cboHLTCode)
    End If
End Sub

Private Sub LoadPayor(Optional HLTCode As String)
    Dim PAY As New ADODB.Recordset
    If Mid(ChkStr(HLTCode), 1, 1) = "N" Then
        sqlPAY = "select PAYCode, PAYCompanyName from PAY "
        sqlPAY = sqlPAY + " Where HLTCODE = 'N' "
    Else
        sqlPAY = "select PAYCode, PAYCompanyName from PAY "
        sqlPAY = sqlPAY + " Where HLTCODE = 'S' "
    End If
    
    cboPayorCode.Clear
    PAY.Open sqlPAY, medixmasterconn, adOpenKeyset, adLockOptimistic
    Do Until PAY.EOF
        cboPayorCode.AddItem ChkStr(PAY!PAYCode) & " - " & ChkStr(PAY!PAYCompanyName)
        PAY.MoveNext
    Loop
    
    PAY.Close
    Set PAY = Nothing
End Sub

Private Sub CmdBack_Click()
    ListView2.Height = 855
    ListView2.Left = 0
    ListView2.Top = 6480
    ListView2.Width = 11775
    Frame22.Visible = True
    Frame4.Visible = True
    Frame6.Visible = True
End Sub

Private Sub CmdStop_Click()
    intExit = 1
    cmdExit.Enabled = True
    CmdClear.Enabled = True
End Sub

Private Sub CmdView_Click()
    ListView2.Height = 6720
    ListView2.Left = 0
    ListView2.Top = 480
    ListView2.Width = 11775
    Frame22.Visible = False
    Frame4.Visible = False
    Frame6.Visible = False
End Sub

Private Sub ListViewHeader()
    ListView2.ColumnHeaders(1).Text = "Claim No"
    ListView2.ColumnHeaders(2).Text = "Member No"
    ListView2.ColumnHeaders(2).Alignment = lvwColumnLeft
    ListView2.ColumnHeaders(3).Text = "Policy No"
    ListView2.ColumnHeaders(3).Alignment = lvwColumnLeft
    ListView2.ColumnHeaders(4).Text = "Covered ID"
    ListView2.ColumnHeaders(4).Alignment = lvwColumnCenter
    ListView2.ColumnHeaders(5).Text = "Prin. Name"
    ListView2.ColumnHeaders(5).Alignment = lvwColumnLeft
    ListView2.ColumnHeaders(6).Text = "Patient Name"
    ListView2.ColumnHeaders(6).Alignment = lvwColumnLeft
    ListView2.ColumnHeaders(7).Text = "Prin. Sex"
    ListView2.ColumnHeaders(7).Alignment = lvwColumnCenter
    ListView2.ColumnHeaders(8).Text = "Prin. Ageband"
    ListView2.ColumnHeaders(8).Alignment = lvwColumnCenter
    ListView2.ColumnHeaders(9).Text = "Patient Sex"
    ListView2.ColumnHeaders(9).Alignment = lvwColumnCenter
    ListView2.ColumnHeaders(10).Text = "Patient Ageband"
    ListView2.ColumnHeaders(10).Alignment = lvwColumnCenter
    ListView2.ColumnHeaders(11).Text = "Insured Type"
    ListView2.ColumnHeaders(11).Alignment = lvwColumnCenter
    ListView2.ColumnHeaders(12).Text = "Plan Code"
    ListView2.ColumnHeaders(12).Alignment = lvwColumnLeft
    ListView2.ColumnHeaders(13).Text = "Hospital"
    ListView2.ColumnHeaders(13).Alignment = lvwColumnLeft
    ListView2.ColumnHeaders(14).Text = "First Diag."
    ListView2.ColumnHeaders(14).Alignment = lvwColumnCenter
    ListView2.ColumnHeaders(15).Text = "Description"
    ListView2.ColumnHeaders(15).Alignment = lvwColumnLeft
    ListView2.ColumnHeaders(16).Text = "Final Diag."
    ListView2.ColumnHeaders(16).Alignment = lvwColumnCenter
    ListView2.ColumnHeaders(17).Text = "Description"
    ListView2.ColumnHeaders(17).Alignment = lvwColumnLeft
    ListView2.ColumnHeaders(18).Text = "Bill Amt"
    ListView2.ColumnHeaders(18).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(19).Text = "Approved Amt"
    ListView2.ColumnHeaders(19).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(20).Text = "Recovery Amt"
    ListView2.ColumnHeaders(20).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(21).Text = "Case Remark"
    ListView2.ColumnHeaders(21).Alignment = lvwColumnLeft
    ListView2.ColumnHeaders(22).Text = "Case Status"
    ListView2.ColumnHeaders(22).Alignment = lvwColumnCenter
    ListView2.ColumnHeaders(23).Text = "Policy Status"
    ListView2.ColumnHeaders(23).Alignment = lvwColumnCenter
    ListView2.ColumnHeaders(24).Text = "DOA"
    ListView2.ColumnHeaders(24).Alignment = lvwColumnCenter
    ListView2.ColumnHeaders(25).Text = "DOD"
    ListView2.ColumnHeaders(25).Alignment = lvwColumnCenter
    ListView2.ColumnHeaders(26).Text = "Duration"
    ListView2.ColumnHeaders(26).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(27).Text = "Payor Eff. Date"
    ListView2.ColumnHeaders(27).Alignment = lvwColumnCenter
    ListView2.ColumnHeaders(28).Text = "Payor Exp. Date"
    ListView2.ColumnHeaders(28).Alignment = lvwColumnCenter
    ListView2.ColumnHeaders(29).Text = ""
    ListView2.ColumnHeaders(29).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(30).Text = ""
    ListView2.ColumnHeaders(30).Alignment = lvwColumnLeft
    ListView2.ColumnHeaders(31).Text = ""
    ListView2.ColumnHeaders(31).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(32).Text = ""
    ListView2.ColumnHeaders(32).Alignment = lvwColumnLeft
    ListView2.ColumnHeaders(33).Text = ""
    ListView2.ColumnHeaders(33).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(34).Text = ""
    ListView2.ColumnHeaders(34).Alignment = lvwColumnLeft
    ListView2.ColumnHeaders(35).Text = ""
    ListView2.ColumnHeaders(35).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(36).Text = ""
    ListView2.ColumnHeaders(36).Alignment = lvwColumnLeft
    ListView2.ColumnHeaders(37).Text = ""
    ListView2.ColumnHeaders(37).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(38).Text = ""
    ListView2.ColumnHeaders(38).Alignment = lvwColumnLeft
    ListView2.ColumnHeaders(39).Text = ""
    ListView2.ColumnHeaders(39).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(40).Text = ""
    ListView2.ColumnHeaders(40).Alignment = lvwColumnLeft
    ListView2.ColumnHeaders(41).Text = ""
    ListView2.ColumnHeaders(41).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(42).Text = ""
    ListView2.ColumnHeaders(42).Alignment = lvwColumnLeft
    ListView2.ColumnHeaders(43).Text = ""
    ListView2.ColumnHeaders(43).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(44).Text = ""
    ListView2.ColumnHeaders(44).Alignment = lvwColumnLeft
    ListView2.ColumnHeaders(45).Text = ""
    ListView2.ColumnHeaders(45).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(46).Text = ""
    ListView2.ColumnHeaders(46).Alignment = lvwColumnLeft
    ListView2.ColumnHeaders(47).Text = ""
    ListView2.ColumnHeaders(47).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(48).Text = ""
    ListView2.ColumnHeaders(48).Alignment = lvwColumnLeft
    ListView2.ColumnHeaders(49).Text = ""
    ListView2.ColumnHeaders(49).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(50).Text = ""
    ListView2.ColumnHeaders(50).Alignment = lvwColumnLeft
    ListView2.ColumnHeaders(51).Text = ""
    ListView2.ColumnHeaders(51).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(52).Text = ""
    ListView2.ColumnHeaders(52).Alignment = lvwColumnLeft
    ListView2.ColumnHeaders(53).Text = ""
    ListView2.ColumnHeaders(53).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(54).Text = ""
    ListView2.ColumnHeaders(54).Alignment = lvwColumnLeft
    ListView2.ColumnHeaders(55).Text = ""
    ListView2.ColumnHeaders(55).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(56).Text = ""
    ListView2.ColumnHeaders(56).Alignment = lvwColumnLeft
    ListView2.ColumnHeaders(57).Text = ""
    ListView2.ColumnHeaders(57).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(58).Text = ""
    ListView2.ColumnHeaders(58).Alignment = lvwColumnLeft
    ListView2.ColumnHeaders(59).Text = ""
    ListView2.ColumnHeaders(59).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(60).Text = ""
    ListView2.ColumnHeaders(60).Alignment = lvwColumnLeft
    ListView2.ColumnHeaders(61).Text = ""
    ListView2.ColumnHeaders(61).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(62).Text = ""
    ListView2.ColumnHeaders(62).Alignment = lvwColumnLeft
    ListView2.ColumnHeaders(63).Text = ""
    ListView2.ColumnHeaders(63).Alignment = lvwColumnRight
End Sub

Private Sub ListViewHeader10()
    ListView2.ColumnHeaders(1).Text = "Hospital"
    ListView2.ColumnHeaders(2).Text = "No of Cases"
    ListView2.ColumnHeaders(2).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(3).Text = "Bill Amt"
    ListView2.ColumnHeaders(3).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(4).Text = "Approved Amt"
    ListView2.ColumnHeaders(4).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(5).Text = "Average Bill"
    ListView2.ColumnHeaders(5).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(6).Text = "Average Approved"
    ListView2.ColumnHeaders(6).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(7).Text = ""
    ListView2.ColumnHeaders(7).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(8).Text = ""
    ListView2.ColumnHeaders(8).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(9).Text = ""
    ListView2.ColumnHeaders(9).Alignment = lvwColumnCenter
    ListView2.ColumnHeaders(10).Text = ""
    ListView2.ColumnHeaders(10).Alignment = lvwColumnCenter
    ListView2.ColumnHeaders(11).Text = ""
    ListView2.ColumnHeaders(11).Alignment = lvwColumnCenter
    ListView2.ColumnHeaders(12).Text = ""
    ListView2.ColumnHeaders(12).Alignment = lvwColumnCenter
    ListView2.ColumnHeaders(13).Text = ""
    ListView2.ColumnHeaders(13).Alignment = lvwColumnCenter
    ListView2.ColumnHeaders(14).Text = ""
    ListView2.ColumnHeaders(14).Alignment = lvwColumnCenter
    ListView2.ColumnHeaders(15).Text = ""
    ListView2.ColumnHeaders(15).Alignment = lvwColumnLeft
    ListView2.ColumnHeaders(16).Text = ""
    ListView2.ColumnHeaders(16).Alignment = lvwColumnCenter
    ListView2.ColumnHeaders(17).Text = ""
    ListView2.ColumnHeaders(17).Alignment = lvwColumnLeft
    ListView2.ColumnHeaders(18).Text = ""
    ListView2.ColumnHeaders(18).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(19).Text = ""
    ListView2.ColumnHeaders(19).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(20).Text = ""
    ListView2.ColumnHeaders(20).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(21).Text = ""
    ListView2.ColumnHeaders(21).Alignment = lvwColumnLeft
    ListView2.ColumnHeaders(22).Text = ""
    ListView2.ColumnHeaders(22).Alignment = lvwColumnCenter
    ListView2.ColumnHeaders(23).Text = ""
    ListView2.ColumnHeaders(23).Alignment = lvwColumnCenter
    ListView2.ColumnHeaders(24).Text = ""
    ListView2.ColumnHeaders(24).Alignment = lvwColumnCenter
    ListView2.ColumnHeaders(25).Text = ""
    ListView2.ColumnHeaders(25).Alignment = lvwColumnCenter
    ListView2.ColumnHeaders(26).Text = ""
    ListView2.ColumnHeaders(26).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(27).Text = ""
    ListView2.ColumnHeaders(27).Alignment = lvwColumnCenter
    ListView2.ColumnHeaders(28).Text = ""
    ListView2.ColumnHeaders(28).Alignment = lvwColumnCenter
    ListView2.ColumnHeaders(29).Text = ""
    ListView2.ColumnHeaders(29).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(30).Text = ""
    ListView2.ColumnHeaders(30).Alignment = lvwColumnLeft
    ListView2.ColumnHeaders(31).Text = ""
    ListView2.ColumnHeaders(31).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(32).Text = ""
    ListView2.ColumnHeaders(32).Alignment = lvwColumnLeft
    ListView2.ColumnHeaders(33).Text = ""
    ListView2.ColumnHeaders(33).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(34).Text = ""
    ListView2.ColumnHeaders(34).Alignment = lvwColumnLeft
    ListView2.ColumnHeaders(35).Text = ""
    ListView2.ColumnHeaders(35).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(36).Text = ""
    ListView2.ColumnHeaders(36).Alignment = lvwColumnLeft
    ListView2.ColumnHeaders(37).Text = ""
    ListView2.ColumnHeaders(37).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(38).Text = ""
    ListView2.ColumnHeaders(38).Alignment = lvwColumnLeft
    ListView2.ColumnHeaders(39).Text = ""
    ListView2.ColumnHeaders(39).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(40).Text = ""
    ListView2.ColumnHeaders(40).Alignment = lvwColumnLeft
    ListView2.ColumnHeaders(41).Text = ""
    ListView2.ColumnHeaders(41).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(42).Text = ""
    ListView2.ColumnHeaders(42).Alignment = lvwColumnLeft
    ListView2.ColumnHeaders(43).Text = ""
    ListView2.ColumnHeaders(43).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(44).Text = ""
    ListView2.ColumnHeaders(44).Alignment = lvwColumnLeft
    ListView2.ColumnHeaders(45).Text = ""
    ListView2.ColumnHeaders(45).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(46).Text = ""
    ListView2.ColumnHeaders(46).Alignment = lvwColumnLeft
    ListView2.ColumnHeaders(47).Text = ""
    ListView2.ColumnHeaders(47).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(48).Text = ""
    ListView2.ColumnHeaders(48).Alignment = lvwColumnLeft
    ListView2.ColumnHeaders(49).Text = ""
    ListView2.ColumnHeaders(49).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(50).Text = ""
    ListView2.ColumnHeaders(50).Alignment = lvwColumnLeft
    ListView2.ColumnHeaders(51).Text = ""
    ListView2.ColumnHeaders(51).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(52).Text = ""
    ListView2.ColumnHeaders(52).Alignment = lvwColumnLeft
    ListView2.ColumnHeaders(53).Text = ""
    ListView2.ColumnHeaders(53).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(54).Text = ""
    ListView2.ColumnHeaders(54).Alignment = lvwColumnLeft
    ListView2.ColumnHeaders(55).Text = ""
    ListView2.ColumnHeaders(55).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(56).Text = ""
    ListView2.ColumnHeaders(56).Alignment = lvwColumnLeft
    ListView2.ColumnHeaders(57).Text = ""
    ListView2.ColumnHeaders(57).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(58).Text = ""
    ListView2.ColumnHeaders(58).Alignment = lvwColumnLeft
    ListView2.ColumnHeaders(59).Text = ""
    ListView2.ColumnHeaders(59).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(60).Text = ""
    ListView2.ColumnHeaders(60).Alignment = lvwColumnLeft
    ListView2.ColumnHeaders(61).Text = ""
    ListView2.ColumnHeaders(61).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(62).Text = ""
    ListView2.ColumnHeaders(62).Alignment = lvwColumnLeft
    ListView2.ColumnHeaders(63).Text = ""
    ListView2.ColumnHeaders(63).Alignment = lvwColumnRight
End Sub

Private Sub ListViewHeader11()
    ListView2.ColumnHeaders(1).Text = "Hospital"
    ListView2.ColumnHeaders(2).Text = "Case No"
    ListView2.ColumnHeaders(2).Alignment = lvwColumnLeft
    ListView2.ColumnHeaders(3).Text = "Member No"
    ListView2.ColumnHeaders(3).Alignment = lvwColumnLeft
    ListView2.ColumnHeaders(4).Text = "Covered ID"
    ListView2.ColumnHeaders(4).Alignment = lvwColumnCenter
    ListView2.ColumnHeaders(5).Text = "Patient Name"
    ListView2.ColumnHeaders(5).Alignment = lvwColumnLeft
    ListView2.ColumnHeaders(6).Text = "Bill Amt"
    ListView2.ColumnHeaders(6).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(7).Text = "Approved Amt"
    ListView2.ColumnHeaders(7).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(8).Text = "Member Amt"
    ListView2.ColumnHeaders(8).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(9).Text = "DOA"
    ListView2.ColumnHeaders(9).Alignment = lvwColumnCenter
    ListView2.ColumnHeaders(10).Text = "DOD"
    ListView2.ColumnHeaders(10).Alignment = lvwColumnCenter
    ListView2.ColumnHeaders(11).Text = "Case Status"
    ListView2.ColumnHeaders(11).Alignment = lvwColumnCenter
    ListView2.ColumnHeaders(12).Text = "Claimability"
    ListView2.ColumnHeaders(12).Alignment = lvwColumnCenter
    ListView2.ColumnHeaders(13).Text = "Final Diag."
    ListView2.ColumnHeaders(13).Alignment = lvwColumnCenter
    ListView2.ColumnHeaders(14).Text = ""
    ListView2.ColumnHeaders(14).Alignment = lvwColumnCenter
    ListView2.ColumnHeaders(15).Text = ""
    ListView2.ColumnHeaders(15).Alignment = lvwColumnLeft
    ListView2.ColumnHeaders(16).Text = ""
    ListView2.ColumnHeaders(16).Alignment = lvwColumnCenter
    ListView2.ColumnHeaders(17).Text = ""
    ListView2.ColumnHeaders(17).Alignment = lvwColumnLeft
    ListView2.ColumnHeaders(18).Text = ""
    ListView2.ColumnHeaders(18).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(19).Text = ""
    ListView2.ColumnHeaders(19).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(20).Text = ""
    ListView2.ColumnHeaders(20).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(21).Text = ""
    ListView2.ColumnHeaders(21).Alignment = lvwColumnLeft
    ListView2.ColumnHeaders(22).Text = ""
    ListView2.ColumnHeaders(22).Alignment = lvwColumnCenter
    ListView2.ColumnHeaders(23).Text = ""
    ListView2.ColumnHeaders(23).Alignment = lvwColumnCenter
    ListView2.ColumnHeaders(24).Text = ""
    ListView2.ColumnHeaders(24).Alignment = lvwColumnCenter
    ListView2.ColumnHeaders(25).Text = ""
    ListView2.ColumnHeaders(25).Alignment = lvwColumnCenter
    ListView2.ColumnHeaders(26).Text = ""
    ListView2.ColumnHeaders(26).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(27).Text = ""
    ListView2.ColumnHeaders(27).Alignment = lvwColumnCenter
    ListView2.ColumnHeaders(28).Text = ""
    ListView2.ColumnHeaders(28).Alignment = lvwColumnCenter
    ListView2.ColumnHeaders(29).Text = ""
    ListView2.ColumnHeaders(29).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(30).Text = ""
    ListView2.ColumnHeaders(30).Alignment = lvwColumnLeft
    ListView2.ColumnHeaders(31).Text = ""
    ListView2.ColumnHeaders(31).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(32).Text = ""
    ListView2.ColumnHeaders(32).Alignment = lvwColumnLeft
    ListView2.ColumnHeaders(33).Text = ""
    ListView2.ColumnHeaders(33).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(34).Text = ""
    ListView2.ColumnHeaders(34).Alignment = lvwColumnLeft
    ListView2.ColumnHeaders(35).Text = ""
    ListView2.ColumnHeaders(35).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(36).Text = ""
    ListView2.ColumnHeaders(36).Alignment = lvwColumnLeft
    ListView2.ColumnHeaders(37).Text = ""
    ListView2.ColumnHeaders(37).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(38).Text = ""
    ListView2.ColumnHeaders(38).Alignment = lvwColumnLeft
    ListView2.ColumnHeaders(39).Text = ""
    ListView2.ColumnHeaders(39).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(40).Text = ""
    ListView2.ColumnHeaders(40).Alignment = lvwColumnLeft
    ListView2.ColumnHeaders(41).Text = ""
    ListView2.ColumnHeaders(41).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(42).Text = ""
    ListView2.ColumnHeaders(42).Alignment = lvwColumnLeft
    ListView2.ColumnHeaders(43).Text = ""
    ListView2.ColumnHeaders(43).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(44).Text = ""
    ListView2.ColumnHeaders(44).Alignment = lvwColumnLeft
    ListView2.ColumnHeaders(45).Text = ""
    ListView2.ColumnHeaders(45).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(46).Text = ""
    ListView2.ColumnHeaders(46).Alignment = lvwColumnLeft
    ListView2.ColumnHeaders(47).Text = ""
    ListView2.ColumnHeaders(47).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(48).Text = ""
    ListView2.ColumnHeaders(48).Alignment = lvwColumnLeft
    ListView2.ColumnHeaders(49).Text = ""
    ListView2.ColumnHeaders(49).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(50).Text = ""
    ListView2.ColumnHeaders(50).Alignment = lvwColumnLeft
    ListView2.ColumnHeaders(51).Text = ""
    ListView2.ColumnHeaders(51).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(52).Text = ""
    ListView2.ColumnHeaders(52).Alignment = lvwColumnLeft
    ListView2.ColumnHeaders(53).Text = ""
    ListView2.ColumnHeaders(53).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(54).Text = ""
    ListView2.ColumnHeaders(54).Alignment = lvwColumnLeft
    ListView2.ColumnHeaders(55).Text = ""
    ListView2.ColumnHeaders(55).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(56).Text = ""
    ListView2.ColumnHeaders(56).Alignment = lvwColumnLeft
    ListView2.ColumnHeaders(57).Text = ""
    ListView2.ColumnHeaders(57).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(58).Text = ""
    ListView2.ColumnHeaders(58).Alignment = lvwColumnLeft
    ListView2.ColumnHeaders(59).Text = ""
    ListView2.ColumnHeaders(59).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(60).Text = ""
    ListView2.ColumnHeaders(60).Alignment = lvwColumnLeft
    ListView2.ColumnHeaders(61).Text = ""
    ListView2.ColumnHeaders(61).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(62).Text = ""
    ListView2.ColumnHeaders(62).Alignment = lvwColumnLeft
    ListView2.ColumnHeaders(63).Text = ""
    ListView2.ColumnHeaders(63).Alignment = lvwColumnRight
End Sub

Private Sub ListViewHeader20()
    ListView2.ColumnHeaders(1).Text = "No"
    ListView2.ColumnHeaders(2).Text = "Case No"
    ListView2.ColumnHeaders(2).Alignment = lvwColumnLeft
    ListView2.ColumnHeaders(3).Text = "Member No"
    ListView2.ColumnHeaders(3).Alignment = lvwColumnLeft
    ListView2.ColumnHeaders(4).Text = "Policy No"
    ListView2.ColumnHeaders(4).Alignment = lvwColumnLeft
    ListView2.ColumnHeaders(5).Text = "Patient Name"
    ListView2.ColumnHeaders(5).Alignment = lvwColumnLeft
    ListView2.ColumnHeaders(6).Text = "DOA"
    ListView2.ColumnHeaders(6).Alignment = lvwColumnCenter
    ListView2.ColumnHeaders(7).Text = "Hospital Name"
    ListView2.ColumnHeaders(7).Alignment = lvwColumnLeft
    ListView2.ColumnHeaders(8).Text = "First Diag."
    ListView2.ColumnHeaders(8).Alignment = lvwColumnCenter
    ListView2.ColumnHeaders(9).Text = ""
    ListView2.ColumnHeaders(9).Alignment = lvwColumnCenter
    ListView2.ColumnHeaders(10).Text = ""
    ListView2.ColumnHeaders(10).Alignment = lvwColumnCenter
    ListView2.ColumnHeaders(11).Text = ""
    ListView2.ColumnHeaders(11).Alignment = lvwColumnCenter
    ListView2.ColumnHeaders(12).Text = ""
    ListView2.ColumnHeaders(12).Alignment = lvwColumnLeft
    ListView2.ColumnHeaders(13).Text = ""
    ListView2.ColumnHeaders(13).Alignment = lvwColumnLeft
    ListView2.ColumnHeaders(14).Text = ""
    ListView2.ColumnHeaders(14).Alignment = lvwColumnCenter
    ListView2.ColumnHeaders(15).Text = ""
    ListView2.ColumnHeaders(15).Alignment = lvwColumnLeft
    ListView2.ColumnHeaders(16).Text = ""
    ListView2.ColumnHeaders(16).Alignment = lvwColumnCenter
    ListView2.ColumnHeaders(17).Text = ""
    ListView2.ColumnHeaders(17).Alignment = lvwColumnLeft
    ListView2.ColumnHeaders(18).Text = ""
    ListView2.ColumnHeaders(18).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(19).Text = ""
    ListView2.ColumnHeaders(19).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(20).Text = ""
    ListView2.ColumnHeaders(20).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(21).Text = ""
    ListView2.ColumnHeaders(21).Alignment = lvwColumnLeft
    ListView2.ColumnHeaders(22).Text = ""
    ListView2.ColumnHeaders(22).Alignment = lvwColumnCenter
    ListView2.ColumnHeaders(23).Text = ""
    ListView2.ColumnHeaders(23).Alignment = lvwColumnCenter
    ListView2.ColumnHeaders(24).Text = ""
    ListView2.ColumnHeaders(24).Alignment = lvwColumnCenter
    ListView2.ColumnHeaders(25).Text = ""
    ListView2.ColumnHeaders(25).Alignment = lvwColumnCenter
    ListView2.ColumnHeaders(26).Text = ""
    ListView2.ColumnHeaders(26).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(27).Text = ""
    ListView2.ColumnHeaders(27).Alignment = lvwColumnCenter
    ListView2.ColumnHeaders(28).Text = ""
    ListView2.ColumnHeaders(28).Alignment = lvwColumnCenter
    ListView2.ColumnHeaders(29).Text = ""
    ListView2.ColumnHeaders(29).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(30).Text = ""
    ListView2.ColumnHeaders(30).Alignment = lvwColumnLeft
    ListView2.ColumnHeaders(31).Text = ""
    ListView2.ColumnHeaders(31).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(32).Text = ""
    ListView2.ColumnHeaders(32).Alignment = lvwColumnLeft
    ListView2.ColumnHeaders(33).Text = ""
    ListView2.ColumnHeaders(33).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(34).Text = ""
    ListView2.ColumnHeaders(34).Alignment = lvwColumnLeft
    ListView2.ColumnHeaders(35).Text = ""
    ListView2.ColumnHeaders(35).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(36).Text = ""
    ListView2.ColumnHeaders(36).Alignment = lvwColumnLeft
    ListView2.ColumnHeaders(37).Text = ""
    ListView2.ColumnHeaders(37).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(38).Text = ""
    ListView2.ColumnHeaders(38).Alignment = lvwColumnLeft
    ListView2.ColumnHeaders(39).Text = ""
    ListView2.ColumnHeaders(39).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(40).Text = ""
    ListView2.ColumnHeaders(40).Alignment = lvwColumnLeft
    ListView2.ColumnHeaders(41).Text = ""
    ListView2.ColumnHeaders(41).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(42).Text = ""
    ListView2.ColumnHeaders(42).Alignment = lvwColumnLeft
    ListView2.ColumnHeaders(43).Text = ""
    ListView2.ColumnHeaders(43).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(44).Text = ""
    ListView2.ColumnHeaders(44).Alignment = lvwColumnLeft
    ListView2.ColumnHeaders(45).Text = ""
    ListView2.ColumnHeaders(45).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(46).Text = ""
    ListView2.ColumnHeaders(46).Alignment = lvwColumnLeft
    ListView2.ColumnHeaders(47).Text = ""
    ListView2.ColumnHeaders(47).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(48).Text = ""
    ListView2.ColumnHeaders(48).Alignment = lvwColumnLeft
    ListView2.ColumnHeaders(49).Text = ""
    ListView2.ColumnHeaders(49).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(50).Text = ""
    ListView2.ColumnHeaders(50).Alignment = lvwColumnLeft
    ListView2.ColumnHeaders(51).Text = ""
    ListView2.ColumnHeaders(51).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(52).Text = ""
    ListView2.ColumnHeaders(52).Alignment = lvwColumnLeft
    ListView2.ColumnHeaders(53).Text = ""
    ListView2.ColumnHeaders(53).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(54).Text = ""
    ListView2.ColumnHeaders(54).Alignment = lvwColumnLeft
    ListView2.ColumnHeaders(55).Text = ""
    ListView2.ColumnHeaders(55).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(56).Text = ""
    ListView2.ColumnHeaders(56).Alignment = lvwColumnLeft
    ListView2.ColumnHeaders(57).Text = ""
    ListView2.ColumnHeaders(57).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(58).Text = ""
    ListView2.ColumnHeaders(58).Alignment = lvwColumnLeft
    ListView2.ColumnHeaders(59).Text = ""
    ListView2.ColumnHeaders(59).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(60).Text = ""
    ListView2.ColumnHeaders(60).Alignment = lvwColumnLeft
    ListView2.ColumnHeaders(61).Text = ""
    ListView2.ColumnHeaders(61).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(62).Text = ""
    ListView2.ColumnHeaders(62).Alignment = lvwColumnLeft
    ListView2.ColumnHeaders(63).Text = ""
    ListView2.ColumnHeaders(63).Alignment = lvwColumnRight
End Sub

Private Sub ListViewHeader21()
    ListView2.ColumnHeaders(1).Text = "Payor"
    ListView2.ColumnHeaders(2).Text = "Date 1"
    ListView2.ColumnHeaders(2).Alignment = lvwColumnLeft
    ListView2.ColumnHeaders(3).Text = "Total"
    ListView2.ColumnHeaders(3).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(4).Text = "Date 2"
    ListView2.ColumnHeaders(4).Alignment = lvwColumnLeft
    ListView2.ColumnHeaders(5).Text = "Total"
    ListView2.ColumnHeaders(5).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(6).Text = "Date 3"
    ListView2.ColumnHeaders(6).Alignment = lvwColumnLeft
    ListView2.ColumnHeaders(7).Text = "Total"
    ListView2.ColumnHeaders(7).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(8).Text = "Date 4"
    ListView2.ColumnHeaders(8).Alignment = lvwColumnLeft
    ListView2.ColumnHeaders(9).Text = "Total"
    ListView2.ColumnHeaders(9).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(10).Text = "Date 5"
    ListView2.ColumnHeaders(10).Alignment = lvwColumnLeft
    ListView2.ColumnHeaders(11).Text = "Total"
    ListView2.ColumnHeaders(11).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(12).Text = "Date 6"
    ListView2.ColumnHeaders(12).Alignment = lvwColumnLeft
    ListView2.ColumnHeaders(13).Text = "Total"
    ListView2.ColumnHeaders(13).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(14).Text = "Date 7"
    ListView2.ColumnHeaders(14).Alignment = lvwColumnLeft
    ListView2.ColumnHeaders(15).Text = "Total"
    ListView2.ColumnHeaders(15).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(16).Text = "Date 8"
    ListView2.ColumnHeaders(16).Alignment = lvwColumnLeft
    ListView2.ColumnHeaders(17).Text = "Total"
    ListView2.ColumnHeaders(17).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(18).Text = "Date 9"
    ListView2.ColumnHeaders(18).Alignment = lvwColumnLeft
    ListView2.ColumnHeaders(19).Text = "Total"
    ListView2.ColumnHeaders(19).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(20).Text = "Date 10"
    ListView2.ColumnHeaders(20).Alignment = lvwColumnLeft
    ListView2.ColumnHeaders(21).Text = "Total"
    ListView2.ColumnHeaders(21).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(22).Text = "Date 11"
    ListView2.ColumnHeaders(22).Alignment = lvwColumnLeft
    ListView2.ColumnHeaders(23).Text = "Total"
    ListView2.ColumnHeaders(23).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(24).Text = "Date 12"
    ListView2.ColumnHeaders(24).Alignment = lvwColumnLeft
    ListView2.ColumnHeaders(25).Text = "Total"
    ListView2.ColumnHeaders(25).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(26).Text = "Date 13"
    ListView2.ColumnHeaders(26).Alignment = lvwColumnLeft
    ListView2.ColumnHeaders(27).Text = "Total"
    ListView2.ColumnHeaders(27).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(28).Text = "Date 14"
    ListView2.ColumnHeaders(28).Alignment = lvwColumnLeft
    ListView2.ColumnHeaders(29).Text = "Total"
    ListView2.ColumnHeaders(29).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(30).Text = "Date 15"
    ListView2.ColumnHeaders(30).Alignment = lvwColumnLeft
    ListView2.ColumnHeaders(31).Text = "Total"
    ListView2.ColumnHeaders(31).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(32).Text = "Date 16"
    ListView2.ColumnHeaders(32).Alignment = lvwColumnLeft
    ListView2.ColumnHeaders(33).Text = "Total"
    ListView2.ColumnHeaders(33).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(34).Text = "Date 17"
    ListView2.ColumnHeaders(34).Alignment = lvwColumnLeft
    ListView2.ColumnHeaders(35).Text = "Total"
    ListView2.ColumnHeaders(35).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(36).Text = "Date 18"
    ListView2.ColumnHeaders(36).Alignment = lvwColumnLeft
    ListView2.ColumnHeaders(37).Text = "Total"
    ListView2.ColumnHeaders(37).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(38).Text = "Date 19"
    ListView2.ColumnHeaders(38).Alignment = lvwColumnLeft
    ListView2.ColumnHeaders(39).Text = "Total"
    ListView2.ColumnHeaders(39).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(40).Text = "Date 20"
    ListView2.ColumnHeaders(40).Alignment = lvwColumnLeft
    ListView2.ColumnHeaders(41).Text = "Total"
    ListView2.ColumnHeaders(41).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(42).Text = "Date 21"
    ListView2.ColumnHeaders(42).Alignment = lvwColumnLeft
    ListView2.ColumnHeaders(43).Text = "Total"
    ListView2.ColumnHeaders(43).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(44).Text = "Date 22"
    ListView2.ColumnHeaders(44).Alignment = lvwColumnLeft
    ListView2.ColumnHeaders(45).Text = "Total"
    ListView2.ColumnHeaders(45).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(46).Text = "Date 23"
    ListView2.ColumnHeaders(46).Alignment = lvwColumnLeft
    ListView2.ColumnHeaders(47).Text = "Total"
    ListView2.ColumnHeaders(47).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(48).Text = "Date 24"
    ListView2.ColumnHeaders(48).Alignment = lvwColumnLeft
    ListView2.ColumnHeaders(49).Text = "Total"
    ListView2.ColumnHeaders(49).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(50).Text = "Date 25"
    ListView2.ColumnHeaders(50).Alignment = lvwColumnLeft
    ListView2.ColumnHeaders(51).Text = "Total"
    ListView2.ColumnHeaders(51).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(52).Text = "Date 26"
    ListView2.ColumnHeaders(52).Alignment = lvwColumnLeft
    ListView2.ColumnHeaders(53).Text = "Total"
    ListView2.ColumnHeaders(53).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(54).Text = "Date 27"
    ListView2.ColumnHeaders(54).Alignment = lvwColumnLeft
    ListView2.ColumnHeaders(55).Text = "Total"
    ListView2.ColumnHeaders(55).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(56).Text = "Date 28"
    ListView2.ColumnHeaders(56).Alignment = lvwColumnLeft
    ListView2.ColumnHeaders(57).Text = "Total"
    ListView2.ColumnHeaders(57).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(58).Text = "Date 29"
    ListView2.ColumnHeaders(58).Alignment = lvwColumnLeft
    ListView2.ColumnHeaders(59).Text = "Total"
    ListView2.ColumnHeaders(59).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(60).Text = "Date 30"
    ListView2.ColumnHeaders(60).Alignment = lvwColumnLeft
    ListView2.ColumnHeaders(61).Text = "Total"
    ListView2.ColumnHeaders(61).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(62).Text = "Date 31"
    ListView2.ColumnHeaders(62).Alignment = lvwColumnLeft
    ListView2.ColumnHeaders(63).Text = "Total"
    ListView2.ColumnHeaders(63).Alignment = lvwColumnRight
End Sub

Private Sub ListViewHeader22()
    ListView2.ColumnHeaders(1).Text = "Payor"
    ListView2.ColumnHeaders(2).Text = "January"
    ListView2.ColumnHeaders(2).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(3).Text = "February"
    ListView2.ColumnHeaders(3).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(4).Text = "March"
    ListView2.ColumnHeaders(4).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(5).Text = "April"
    ListView2.ColumnHeaders(5).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(6).Text = "May"
    ListView2.ColumnHeaders(6).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(7).Text = "June"
    ListView2.ColumnHeaders(7).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(8).Text = "Date July"
    ListView2.ColumnHeaders(8).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(9).Text = "August"
    ListView2.ColumnHeaders(9).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(10).Text = "September"
    ListView2.ColumnHeaders(10).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(11).Text = "October"
    ListView2.ColumnHeaders(11).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(12).Text = "November"
    ListView2.ColumnHeaders(12).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(13).Text = "December"
    ListView2.ColumnHeaders(13).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(14).Text = ""
    ListView2.ColumnHeaders(14).Alignment = lvwColumnLeft
    ListView2.ColumnHeaders(15).Text = ""
    ListView2.ColumnHeaders(15).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(16).Text = ""
    ListView2.ColumnHeaders(16).Alignment = lvwColumnLeft
    ListView2.ColumnHeaders(17).Text = ""
    ListView2.ColumnHeaders(17).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(18).Text = ""
    ListView2.ColumnHeaders(18).Alignment = lvwColumnLeft
    ListView2.ColumnHeaders(19).Text = ""
    ListView2.ColumnHeaders(19).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(20).Text = ""
    ListView2.ColumnHeaders(20).Alignment = lvwColumnLeft
    ListView2.ColumnHeaders(21).Text = ""
    ListView2.ColumnHeaders(21).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(22).Text = ""
    ListView2.ColumnHeaders(22).Alignment = lvwColumnLeft
    ListView2.ColumnHeaders(23).Text = ""
    ListView2.ColumnHeaders(23).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(24).Text = ""
    ListView2.ColumnHeaders(24).Alignment = lvwColumnLeft
    ListView2.ColumnHeaders(25).Text = ""
    ListView2.ColumnHeaders(25).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(26).Text = ""
    ListView2.ColumnHeaders(26).Alignment = lvwColumnLeft
    ListView2.ColumnHeaders(27).Text = ""
    ListView2.ColumnHeaders(27).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(28).Text = ""
    ListView2.ColumnHeaders(28).Alignment = lvwColumnLeft
    ListView2.ColumnHeaders(29).Text = ""
    ListView2.ColumnHeaders(29).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(30).Text = ""
    ListView2.ColumnHeaders(30).Alignment = lvwColumnLeft
    ListView2.ColumnHeaders(31).Text = ""
    ListView2.ColumnHeaders(31).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(32).Text = ""
    ListView2.ColumnHeaders(32).Alignment = lvwColumnLeft
    ListView2.ColumnHeaders(33).Text = ""
    ListView2.ColumnHeaders(33).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(34).Text = ""
    ListView2.ColumnHeaders(34).Alignment = lvwColumnLeft
    ListView2.ColumnHeaders(35).Text = ""
    ListView2.ColumnHeaders(35).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(36).Text = ""
    ListView2.ColumnHeaders(36).Alignment = lvwColumnLeft
    ListView2.ColumnHeaders(37).Text = ""
    ListView2.ColumnHeaders(37).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(38).Text = ""
    ListView2.ColumnHeaders(38).Alignment = lvwColumnLeft
    ListView2.ColumnHeaders(39).Text = ""
    ListView2.ColumnHeaders(39).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(40).Text = ""
    ListView2.ColumnHeaders(40).Alignment = lvwColumnLeft
    ListView2.ColumnHeaders(41).Text = ""
    ListView2.ColumnHeaders(41).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(42).Text = ""
    ListView2.ColumnHeaders(42).Alignment = lvwColumnLeft
    ListView2.ColumnHeaders(43).Text = ""
    ListView2.ColumnHeaders(43).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(44).Text = ""
    ListView2.ColumnHeaders(44).Alignment = lvwColumnLeft
    ListView2.ColumnHeaders(45).Text = ""
    ListView2.ColumnHeaders(45).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(46).Text = ""
    ListView2.ColumnHeaders(46).Alignment = lvwColumnLeft
    ListView2.ColumnHeaders(47).Text = ""
    ListView2.ColumnHeaders(47).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(48).Text = ""
    ListView2.ColumnHeaders(48).Alignment = lvwColumnLeft
    ListView2.ColumnHeaders(49).Text = ""
    ListView2.ColumnHeaders(49).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(50).Text = ""
    ListView2.ColumnHeaders(50).Alignment = lvwColumnLeft
    ListView2.ColumnHeaders(51).Text = ""
    ListView2.ColumnHeaders(51).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(52).Text = ""
    ListView2.ColumnHeaders(52).Alignment = lvwColumnLeft
    ListView2.ColumnHeaders(53).Text = ""
    ListView2.ColumnHeaders(53).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(54).Text = ""
    ListView2.ColumnHeaders(54).Alignment = lvwColumnLeft
    ListView2.ColumnHeaders(55).Text = ""
    ListView2.ColumnHeaders(55).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(56).Text = ""
    ListView2.ColumnHeaders(56).Alignment = lvwColumnLeft
    ListView2.ColumnHeaders(57).Text = ""
    ListView2.ColumnHeaders(57).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(58).Text = ""
    ListView2.ColumnHeaders(58).Alignment = lvwColumnLeft
    ListView2.ColumnHeaders(59).Text = ""
    ListView2.ColumnHeaders(59).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(60).Text = ""
    ListView2.ColumnHeaders(60).Alignment = lvwColumnLeft
    ListView2.ColumnHeaders(61).Text = ""
    ListView2.ColumnHeaders(61).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(62).Text = ""
    ListView2.ColumnHeaders(62).Alignment = lvwColumnLeft
    ListView2.ColumnHeaders(63).Text = ""
    ListView2.ColumnHeaders(63).Alignment = lvwColumnRight
End Sub

Private Sub ListViewHeader23()
    ListView2.ColumnHeaders(1).Text = "Agent No"
    ListView2.ColumnHeaders(2).Text = "Worksheet No"
    ListView2.ColumnHeaders(2).Alignment = lvwColumnLeft
    ListView2.ColumnHeaders(3).Text = "Member No"
    ListView2.ColumnHeaders(3).Alignment = lvwColumnLeft
    ListView2.ColumnHeaders(4).Text = "Covered ID"
    ListView2.ColumnHeaders(4).Alignment = lvwColumnCenter
    ListView2.ColumnHeaders(5).Text = "Patient Name"
    ListView2.ColumnHeaders(5).Alignment = lvwColumnLeft
    ListView2.ColumnHeaders(6).Text = "Doctor"
    ListView2.ColumnHeaders(6).Alignment = lvwColumnLeft
    ListView2.ColumnHeaders(7).Text = "Bill Amt"
    ListView2.ColumnHeaders(7).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(8).Text = "Approved Amt"
    ListView2.ColumnHeaders(8).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(9).Text = "Recovery Amt"
    ListView2.ColumnHeaders(9).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(10).Text = "First Diag."
    ListView2.ColumnHeaders(10).Alignment = lvwColumnCenter
    ListView2.ColumnHeaders(11).Text = "Description"
    ListView2.ColumnHeaders(11).Alignment = lvwColumnLeft
    ListView2.ColumnHeaders(12).Text = "Final Diag."
    ListView2.ColumnHeaders(12).Alignment = lvwColumnCenter
    ListView2.ColumnHeaders(13).Text = "Description"
    ListView2.ColumnHeaders(13).Alignment = lvwColumnLeft
    ListView2.ColumnHeaders(14).Text = "DOA"
    ListView2.ColumnHeaders(14).Alignment = lvwColumnCenter
    ListView2.ColumnHeaders(15).Text = "DOD"
    ListView2.ColumnHeaders(15).Alignment = lvwColumnCenter
    ListView2.ColumnHeaders(16).Text = "Ins Code"
    ListView2.ColumnHeaders(16).Alignment = lvwColumnCenter
    ListView2.ColumnHeaders(17).Text = "Plan Code"
    ListView2.ColumnHeaders(17).Alignment = lvwColumnCenter
    ListView2.ColumnHeaders(18).Text = "Case Status"
    ListView2.ColumnHeaders(18).Alignment = lvwColumnCenter
    ListView2.ColumnHeaders(19).Text = ""
    ListView2.ColumnHeaders(20).Text = ""
    ListView2.ColumnHeaders(21).Text = ""
    ListView2.ColumnHeaders(22).Text = ""
    ListView2.ColumnHeaders(23).Text = ""
    ListView2.ColumnHeaders(24).Text = ""
    ListView2.ColumnHeaders(25).Text = ""
    ListView2.ColumnHeaders(26).Text = ""
    ListView2.ColumnHeaders(27).Text = ""
    ListView2.ColumnHeaders(28).Text = ""
    ListView2.ColumnHeaders(29).Text = ""
    ListView2.ColumnHeaders(30).Text = ""
    ListView2.ColumnHeaders(31).Text = ""
    ListView2.ColumnHeaders(32).Text = ""
    ListView2.ColumnHeaders(33).Text = ""
    ListView2.ColumnHeaders(34).Text = ""
    ListView2.ColumnHeaders(35).Text = ""
    ListView2.ColumnHeaders(36).Text = ""
    ListView2.ColumnHeaders(37).Text = ""
    ListView2.ColumnHeaders(38).Text = ""
    ListView2.ColumnHeaders(39).Text = ""
    ListView2.ColumnHeaders(40).Text = ""
    ListView2.ColumnHeaders(41).Text = ""
    ListView2.ColumnHeaders(42).Text = ""
    ListView2.ColumnHeaders(43).Text = ""
    ListView2.ColumnHeaders(44).Text = ""
    ListView2.ColumnHeaders(45).Text = ""
    ListView2.ColumnHeaders(46).Text = ""
    ListView2.ColumnHeaders(47).Text = ""
    ListView2.ColumnHeaders(48).Text = ""
    ListView2.ColumnHeaders(49).Text = ""
    ListView2.ColumnHeaders(50).Text = ""
    ListView2.ColumnHeaders(51).Text = ""
    ListView2.ColumnHeaders(52).Text = ""
    ListView2.ColumnHeaders(53).Text = ""
    ListView2.ColumnHeaders(54).Text = ""
    ListView2.ColumnHeaders(55).Text = ""
    ListView2.ColumnHeaders(56).Text = ""
    ListView2.ColumnHeaders(57).Text = ""
    ListView2.ColumnHeaders(58).Text = ""
    ListView2.ColumnHeaders(59).Text = ""
    ListView2.ColumnHeaders(60).Text = ""
    ListView2.ColumnHeaders(61).Text = ""
    ListView2.ColumnHeaders(62).Text = ""
    ListView2.ColumnHeaders(63).Text = ""
    
End Sub

Private Sub ListViewHeader24()
    ListView2.ColumnHeaders(1).Text = "Sex"
    ListView2.ColumnHeaders(2).Text = "Ageband 1"
    ListView2.ColumnHeaders(2).Alignment = lvwColumnLeft
    ListView2.ColumnHeaders(3).Text = "Ageband 2"
    ListView2.ColumnHeaders(3).Alignment = lvwColumnLeft
    ListView2.ColumnHeaders(4).Text = "Ageband 3"
    ListView2.ColumnHeaders(4).Alignment = lvwColumnCenter
    ListView2.ColumnHeaders(5).Text = "Ageband 4"
    ListView2.ColumnHeaders(5).Alignment = lvwColumnLeft
    ListView2.ColumnHeaders(6).Text = "Ageband 5"
    ListView2.ColumnHeaders(6).Alignment = lvwColumnLeft
    ListView2.ColumnHeaders(7).Text = "Ageband 6"
    ListView2.ColumnHeaders(7).Alignment = lvwColumnLeft
    ListView2.ColumnHeaders(8).Text = "Total Cases"
    ListView2.ColumnHeaders(8).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(9).Text = "Total Duration"
    ListView2.ColumnHeaders(9).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(10).Text = "Total Actual Amt"
    ListView2.ColumnHeaders(10).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(11).Text = "Total Approved Amt"
    ListView2.ColumnHeaders(11).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(12).Text = ""
    ListView2.ColumnHeaders(13).Text = ""
    ListView2.ColumnHeaders(14).Text = ""
    ListView2.ColumnHeaders(15).Text = ""
    ListView2.ColumnHeaders(16).Text = ""
    ListView2.ColumnHeaders(17).Text = ""
    ListView2.ColumnHeaders(18).Text = ""
    ListView2.ColumnHeaders(19).Text = ""
    ListView2.ColumnHeaders(20).Text = ""
    ListView2.ColumnHeaders(21).Text = ""
    ListView2.ColumnHeaders(22).Text = ""
    ListView2.ColumnHeaders(23).Text = ""
    ListView2.ColumnHeaders(24).Text = ""
    ListView2.ColumnHeaders(25).Text = ""
    ListView2.ColumnHeaders(26).Text = ""
    ListView2.ColumnHeaders(27).Text = ""
    ListView2.ColumnHeaders(28).Text = ""
    ListView2.ColumnHeaders(29).Text = ""
    ListView2.ColumnHeaders(30).Text = ""
    ListView2.ColumnHeaders(31).Text = ""
    ListView2.ColumnHeaders(32).Text = ""
    ListView2.ColumnHeaders(33).Text = ""
    ListView2.ColumnHeaders(34).Text = ""
    ListView2.ColumnHeaders(35).Text = ""
    ListView2.ColumnHeaders(36).Text = ""
    ListView2.ColumnHeaders(37).Text = ""
    ListView2.ColumnHeaders(38).Text = ""
    ListView2.ColumnHeaders(39).Text = ""
    ListView2.ColumnHeaders(40).Text = ""
    ListView2.ColumnHeaders(41).Text = ""
    ListView2.ColumnHeaders(42).Text = ""
    ListView2.ColumnHeaders(43).Text = ""
    ListView2.ColumnHeaders(44).Text = ""
    ListView2.ColumnHeaders(45).Text = ""
    ListView2.ColumnHeaders(46).Text = ""
    ListView2.ColumnHeaders(47).Text = ""
    ListView2.ColumnHeaders(48).Text = ""
    ListView2.ColumnHeaders(49).Text = ""
    ListView2.ColumnHeaders(50).Text = ""
    ListView2.ColumnHeaders(51).Text = ""
    ListView2.ColumnHeaders(52).Text = ""
    ListView2.ColumnHeaders(53).Text = ""
    ListView2.ColumnHeaders(54).Text = ""
    ListView2.ColumnHeaders(55).Text = ""
    ListView2.ColumnHeaders(56).Text = ""
    ListView2.ColumnHeaders(57).Text = ""
    ListView2.ColumnHeaders(58).Text = ""
    ListView2.ColumnHeaders(59).Text = ""
    ListView2.ColumnHeaders(60).Text = ""
    ListView2.ColumnHeaders(61).Text = ""
    ListView2.ColumnHeaders(62).Text = ""
    ListView2.ColumnHeaders(63).Text = ""
    
End Sub

Private Sub ListViewHeader25()
    ListView2.ColumnHeaders(1).Text = "Sex"
    ListView2.ColumnHeaders(2).Text = "Total Cases"
    ListView2.ColumnHeaders(2).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(3).Text = "Total Duration"
    ListView2.ColumnHeaders(3).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(4).Text = "Total Actual Amt"
    ListView2.ColumnHeaders(4).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(5).Text = "Total Approved Amt"
    ListView2.ColumnHeaders(5).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(6).Text = ""
    ListView2.ColumnHeaders(7).Text = ""
    ListView2.ColumnHeaders(8).Text = ""
    ListView2.ColumnHeaders(9).Text = ""
    ListView2.ColumnHeaders(10).Text = ""
    ListView2.ColumnHeaders(11).Text = ""
    ListView2.ColumnHeaders(12).Text = ""
    ListView2.ColumnHeaders(13).Text = ""
    ListView2.ColumnHeaders(14).Text = ""
    ListView2.ColumnHeaders(15).Text = ""
    ListView2.ColumnHeaders(16).Text = ""
    ListView2.ColumnHeaders(17).Text = ""
    ListView2.ColumnHeaders(18).Text = ""
    ListView2.ColumnHeaders(19).Text = ""
    ListView2.ColumnHeaders(20).Text = ""
    ListView2.ColumnHeaders(21).Text = ""
    ListView2.ColumnHeaders(22).Text = ""
    ListView2.ColumnHeaders(23).Text = ""
    ListView2.ColumnHeaders(24).Text = ""
    ListView2.ColumnHeaders(25).Text = ""
    ListView2.ColumnHeaders(26).Text = ""
    ListView2.ColumnHeaders(27).Text = ""
    ListView2.ColumnHeaders(28).Text = ""
    ListView2.ColumnHeaders(29).Text = ""
    ListView2.ColumnHeaders(30).Text = ""
    ListView2.ColumnHeaders(31).Text = ""
    ListView2.ColumnHeaders(32).Text = ""
    ListView2.ColumnHeaders(33).Text = ""
    ListView2.ColumnHeaders(34).Text = ""
    ListView2.ColumnHeaders(35).Text = ""
    ListView2.ColumnHeaders(36).Text = ""
    ListView2.ColumnHeaders(37).Text = ""
    ListView2.ColumnHeaders(38).Text = ""
    ListView2.ColumnHeaders(39).Text = ""
    ListView2.ColumnHeaders(40).Text = ""
    ListView2.ColumnHeaders(41).Text = ""
    ListView2.ColumnHeaders(42).Text = ""
    ListView2.ColumnHeaders(43).Text = ""
    ListView2.ColumnHeaders(44).Text = ""
    ListView2.ColumnHeaders(45).Text = ""
    ListView2.ColumnHeaders(46).Text = ""
    ListView2.ColumnHeaders(47).Text = ""
    ListView2.ColumnHeaders(48).Text = ""
    ListView2.ColumnHeaders(49).Text = ""
    ListView2.ColumnHeaders(50).Text = ""
    ListView2.ColumnHeaders(51).Text = ""
    ListView2.ColumnHeaders(52).Text = ""
    ListView2.ColumnHeaders(53).Text = ""
    ListView2.ColumnHeaders(54).Text = ""
    ListView2.ColumnHeaders(55).Text = ""
    ListView2.ColumnHeaders(56).Text = ""
    ListView2.ColumnHeaders(57).Text = ""
    ListView2.ColumnHeaders(58).Text = ""
    ListView2.ColumnHeaders(59).Text = ""
    ListView2.ColumnHeaders(60).Text = ""
    ListView2.ColumnHeaders(61).Text = ""
    ListView2.ColumnHeaders(62).Text = ""
    ListView2.ColumnHeaders(63).Text = ""
    
End Sub

Private Sub ListViewHeader26()
    ListView2.ColumnHeaders(1).Text = "Race"
    ListView2.ColumnHeaders(2).Text = "Total Cases"
    ListView2.ColumnHeaders(2).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(3).Text = "Total Duration"
    ListView2.ColumnHeaders(3).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(4).Text = "Total Actual Amt"
    ListView2.ColumnHeaders(4).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(5).Text = "Total Approved Amt"
    ListView2.ColumnHeaders(5).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(6).Text = ""
    ListView2.ColumnHeaders(7).Text = ""
    ListView2.ColumnHeaders(8).Text = ""
    ListView2.ColumnHeaders(9).Text = ""
    ListView2.ColumnHeaders(10).Text = ""
    ListView2.ColumnHeaders(11).Text = ""
    ListView2.ColumnHeaders(12).Text = ""
    ListView2.ColumnHeaders(13).Text = ""
    ListView2.ColumnHeaders(14).Text = ""
    ListView2.ColumnHeaders(15).Text = ""
    ListView2.ColumnHeaders(16).Text = ""
    ListView2.ColumnHeaders(17).Text = ""
    ListView2.ColumnHeaders(18).Text = ""
    ListView2.ColumnHeaders(19).Text = ""
    ListView2.ColumnHeaders(20).Text = ""
    ListView2.ColumnHeaders(21).Text = ""
    ListView2.ColumnHeaders(22).Text = ""
    ListView2.ColumnHeaders(23).Text = ""
    ListView2.ColumnHeaders(24).Text = ""
    ListView2.ColumnHeaders(25).Text = ""
    ListView2.ColumnHeaders(26).Text = ""
    ListView2.ColumnHeaders(27).Text = ""
    ListView2.ColumnHeaders(28).Text = ""
    ListView2.ColumnHeaders(29).Text = ""
    ListView2.ColumnHeaders(30).Text = ""
    ListView2.ColumnHeaders(31).Text = ""
    ListView2.ColumnHeaders(32).Text = ""
    ListView2.ColumnHeaders(33).Text = ""
    ListView2.ColumnHeaders(34).Text = ""
    ListView2.ColumnHeaders(35).Text = ""
    ListView2.ColumnHeaders(36).Text = ""
    ListView2.ColumnHeaders(37).Text = ""
    ListView2.ColumnHeaders(38).Text = ""
    ListView2.ColumnHeaders(39).Text = ""
    ListView2.ColumnHeaders(40).Text = ""
    ListView2.ColumnHeaders(41).Text = ""
    ListView2.ColumnHeaders(42).Text = ""
    ListView2.ColumnHeaders(43).Text = ""
    ListView2.ColumnHeaders(44).Text = ""
    ListView2.ColumnHeaders(45).Text = ""
    ListView2.ColumnHeaders(46).Text = ""
    ListView2.ColumnHeaders(47).Text = ""
    ListView2.ColumnHeaders(48).Text = ""
    ListView2.ColumnHeaders(49).Text = ""
    ListView2.ColumnHeaders(50).Text = ""
    ListView2.ColumnHeaders(51).Text = ""
    ListView2.ColumnHeaders(52).Text = ""
    ListView2.ColumnHeaders(53).Text = ""
    ListView2.ColumnHeaders(54).Text = ""
    ListView2.ColumnHeaders(55).Text = ""
    ListView2.ColumnHeaders(56).Text = ""
    ListView2.ColumnHeaders(57).Text = ""
    ListView2.ColumnHeaders(58).Text = ""
    ListView2.ColumnHeaders(59).Text = ""
    ListView2.ColumnHeaders(60).Text = ""
    ListView2.ColumnHeaders(61).Text = ""
    ListView2.ColumnHeaders(62).Text = ""
    ListView2.ColumnHeaders(63).Text = ""
End Sub

Private Sub ListViewHeader27()
    ListView2.ColumnHeaders(1).Text = "Ins Type"
    ListView2.ColumnHeaders(2).Text = "Total Cases"
    ListView2.ColumnHeaders(2).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(3).Text = "Total Duration"
    ListView2.ColumnHeaders(3).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(4).Text = "Total Actual Amt"
    ListView2.ColumnHeaders(4).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(5).Text = "Total Approved Amt"
    ListView2.ColumnHeaders(5).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(6).Text = ""
    ListView2.ColumnHeaders(7).Text = ""
    ListView2.ColumnHeaders(8).Text = ""
    ListView2.ColumnHeaders(9).Text = ""
    ListView2.ColumnHeaders(10).Text = ""
    ListView2.ColumnHeaders(11).Text = ""
    ListView2.ColumnHeaders(12).Text = ""
    ListView2.ColumnHeaders(13).Text = ""
    ListView2.ColumnHeaders(14).Text = ""
    ListView2.ColumnHeaders(15).Text = ""
    ListView2.ColumnHeaders(16).Text = ""
    ListView2.ColumnHeaders(17).Text = ""
    ListView2.ColumnHeaders(18).Text = ""
    ListView2.ColumnHeaders(19).Text = ""
    ListView2.ColumnHeaders(20).Text = ""
    ListView2.ColumnHeaders(21).Text = ""
    ListView2.ColumnHeaders(22).Text = ""
    ListView2.ColumnHeaders(23).Text = ""
    ListView2.ColumnHeaders(24).Text = ""
    ListView2.ColumnHeaders(25).Text = ""
    ListView2.ColumnHeaders(26).Text = ""
    ListView2.ColumnHeaders(27).Text = ""
    ListView2.ColumnHeaders(28).Text = ""
    ListView2.ColumnHeaders(29).Text = ""
    ListView2.ColumnHeaders(30).Text = ""
    ListView2.ColumnHeaders(31).Text = ""
    ListView2.ColumnHeaders(32).Text = ""
    ListView2.ColumnHeaders(33).Text = ""
    ListView2.ColumnHeaders(34).Text = ""
    ListView2.ColumnHeaders(35).Text = ""
    ListView2.ColumnHeaders(36).Text = ""
    ListView2.ColumnHeaders(37).Text = ""
    ListView2.ColumnHeaders(38).Text = ""
    ListView2.ColumnHeaders(39).Text = ""
    ListView2.ColumnHeaders(40).Text = ""
    ListView2.ColumnHeaders(41).Text = ""
    ListView2.ColumnHeaders(42).Text = ""
    ListView2.ColumnHeaders(43).Text = ""
    ListView2.ColumnHeaders(44).Text = ""
    ListView2.ColumnHeaders(45).Text = ""
    ListView2.ColumnHeaders(46).Text = ""
    ListView2.ColumnHeaders(47).Text = ""
    ListView2.ColumnHeaders(48).Text = ""
    ListView2.ColumnHeaders(49).Text = ""
    ListView2.ColumnHeaders(50).Text = ""
    ListView2.ColumnHeaders(51).Text = ""
    ListView2.ColumnHeaders(52).Text = ""
    ListView2.ColumnHeaders(53).Text = ""
    ListView2.ColumnHeaders(54).Text = ""
    ListView2.ColumnHeaders(55).Text = ""
    ListView2.ColumnHeaders(56).Text = ""
    ListView2.ColumnHeaders(57).Text = ""
    ListView2.ColumnHeaders(58).Text = ""
    ListView2.ColumnHeaders(59).Text = ""
    ListView2.ColumnHeaders(60).Text = ""
    ListView2.ColumnHeaders(61).Text = ""
    ListView2.ColumnHeaders(62).Text = ""
    ListView2.ColumnHeaders(63).Text = ""
End Sub

Private Sub ListViewHeader28()
    ListView2.ColumnHeaders(1).Text = "Plan Type"
    ListView2.ColumnHeaders(2).Text = "Total Cases"
    ListView2.ColumnHeaders(2).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(3).Text = "Total Duration"
    ListView2.ColumnHeaders(3).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(4).Text = "Total Actual Amt"
    ListView2.ColumnHeaders(4).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(5).Text = "Total Approved Amt"
    ListView2.ColumnHeaders(5).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(6).Text = ""
    ListView2.ColumnHeaders(7).Text = ""
    ListView2.ColumnHeaders(8).Text = ""
    ListView2.ColumnHeaders(9).Text = ""
    ListView2.ColumnHeaders(10).Text = ""
    ListView2.ColumnHeaders(11).Text = ""
    ListView2.ColumnHeaders(12).Text = ""
    ListView2.ColumnHeaders(13).Text = ""
    ListView2.ColumnHeaders(14).Text = ""
    ListView2.ColumnHeaders(15).Text = ""
    ListView2.ColumnHeaders(16).Text = ""
    ListView2.ColumnHeaders(17).Text = ""
    ListView2.ColumnHeaders(18).Text = ""
    ListView2.ColumnHeaders(19).Text = ""
    ListView2.ColumnHeaders(20).Text = ""
    ListView2.ColumnHeaders(21).Text = ""
    ListView2.ColumnHeaders(22).Text = ""
    ListView2.ColumnHeaders(23).Text = ""
    ListView2.ColumnHeaders(24).Text = ""
    ListView2.ColumnHeaders(25).Text = ""
    ListView2.ColumnHeaders(26).Text = ""
    ListView2.ColumnHeaders(27).Text = ""
    ListView2.ColumnHeaders(28).Text = ""
    ListView2.ColumnHeaders(29).Text = ""
    ListView2.ColumnHeaders(30).Text = ""
    ListView2.ColumnHeaders(31).Text = ""
    ListView2.ColumnHeaders(32).Text = ""
    ListView2.ColumnHeaders(33).Text = ""
    ListView2.ColumnHeaders(34).Text = ""
    ListView2.ColumnHeaders(35).Text = ""
    ListView2.ColumnHeaders(36).Text = ""
    ListView2.ColumnHeaders(37).Text = ""
    ListView2.ColumnHeaders(38).Text = ""
    ListView2.ColumnHeaders(39).Text = ""
    ListView2.ColumnHeaders(40).Text = ""
    ListView2.ColumnHeaders(41).Text = ""
    ListView2.ColumnHeaders(42).Text = ""
    ListView2.ColumnHeaders(43).Text = ""
    ListView2.ColumnHeaders(44).Text = ""
    ListView2.ColumnHeaders(45).Text = ""
    ListView2.ColumnHeaders(46).Text = ""
    ListView2.ColumnHeaders(47).Text = ""
    ListView2.ColumnHeaders(48).Text = ""
    ListView2.ColumnHeaders(49).Text = ""
    ListView2.ColumnHeaders(50).Text = ""
    ListView2.ColumnHeaders(51).Text = ""
    ListView2.ColumnHeaders(52).Text = ""
    ListView2.ColumnHeaders(53).Text = ""
    ListView2.ColumnHeaders(54).Text = ""
    ListView2.ColumnHeaders(55).Text = ""
    ListView2.ColumnHeaders(56).Text = ""
    ListView2.ColumnHeaders(57).Text = ""
    ListView2.ColumnHeaders(58).Text = ""
    ListView2.ColumnHeaders(59).Text = ""
    ListView2.ColumnHeaders(60).Text = ""
    ListView2.ColumnHeaders(61).Text = ""
    ListView2.ColumnHeaders(62).Text = ""
    ListView2.ColumnHeaders(63).Text = ""
End Sub

Private Sub ListViewHeader29()
    ListView2.ColumnHeaders(1).Text = "State"
    ListView2.ColumnHeaders(2).Text = "Total Cases"
    ListView2.ColumnHeaders(2).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(3).Text = "Total Duration"
    ListView2.ColumnHeaders(3).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(4).Text = "Total Actual Amt"
    ListView2.ColumnHeaders(4).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(5).Text = "Total Approved Amt"
    ListView2.ColumnHeaders(5).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(6).Text = ""
    ListView2.ColumnHeaders(7).Text = ""
    ListView2.ColumnHeaders(8).Text = ""
    ListView2.ColumnHeaders(9).Text = ""
    ListView2.ColumnHeaders(10).Text = ""
    ListView2.ColumnHeaders(11).Text = ""
    ListView2.ColumnHeaders(12).Text = ""
    ListView2.ColumnHeaders(13).Text = ""
    ListView2.ColumnHeaders(14).Text = ""
    ListView2.ColumnHeaders(15).Text = ""
    ListView2.ColumnHeaders(16).Text = ""
    ListView2.ColumnHeaders(17).Text = ""
    ListView2.ColumnHeaders(18).Text = ""
    ListView2.ColumnHeaders(19).Text = ""
    ListView2.ColumnHeaders(20).Text = ""
    ListView2.ColumnHeaders(21).Text = ""
    ListView2.ColumnHeaders(22).Text = ""
    ListView2.ColumnHeaders(23).Text = ""
    ListView2.ColumnHeaders(24).Text = ""
    ListView2.ColumnHeaders(25).Text = ""
    ListView2.ColumnHeaders(26).Text = ""
    ListView2.ColumnHeaders(27).Text = ""
    ListView2.ColumnHeaders(28).Text = ""
    ListView2.ColumnHeaders(29).Text = ""
    ListView2.ColumnHeaders(30).Text = ""
    ListView2.ColumnHeaders(31).Text = ""
    ListView2.ColumnHeaders(32).Text = ""
    ListView2.ColumnHeaders(33).Text = ""
    ListView2.ColumnHeaders(34).Text = ""
    ListView2.ColumnHeaders(35).Text = ""
    ListView2.ColumnHeaders(36).Text = ""
    ListView2.ColumnHeaders(37).Text = ""
    ListView2.ColumnHeaders(38).Text = ""
    ListView2.ColumnHeaders(39).Text = ""
    ListView2.ColumnHeaders(40).Text = ""
    ListView2.ColumnHeaders(41).Text = ""
    ListView2.ColumnHeaders(42).Text = ""
    ListView2.ColumnHeaders(43).Text = ""
    ListView2.ColumnHeaders(44).Text = ""
    ListView2.ColumnHeaders(45).Text = ""
    ListView2.ColumnHeaders(46).Text = ""
    ListView2.ColumnHeaders(47).Text = ""
    ListView2.ColumnHeaders(48).Text = ""
    ListView2.ColumnHeaders(49).Text = ""
    ListView2.ColumnHeaders(50).Text = ""
    ListView2.ColumnHeaders(51).Text = ""
    ListView2.ColumnHeaders(52).Text = ""
    ListView2.ColumnHeaders(53).Text = ""
    ListView2.ColumnHeaders(54).Text = ""
    ListView2.ColumnHeaders(55).Text = ""
    ListView2.ColumnHeaders(56).Text = ""
    ListView2.ColumnHeaders(57).Text = ""
    ListView2.ColumnHeaders(58).Text = ""
    ListView2.ColumnHeaders(59).Text = ""
    ListView2.ColumnHeaders(60).Text = ""
    ListView2.ColumnHeaders(61).Text = ""
    ListView2.ColumnHeaders(62).Text = ""
    ListView2.ColumnHeaders(63).Text = ""
End Sub

Private Sub ListViewHeader30()
    ListView2.ColumnHeaders(1).Text = "Payor"
    ListView2.ColumnHeaders(2).Text = "Total Cases"
    ListView2.ColumnHeaders(2).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(3).Text = "Total Duration"
    ListView2.ColumnHeaders(3).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(4).Text = "Total Actual Amt"
    ListView2.ColumnHeaders(4).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(5).Text = "Total Approved Amt"
    ListView2.ColumnHeaders(5).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(6).Text = ""
    ListView2.ColumnHeaders(7).Text = ""
    ListView2.ColumnHeaders(8).Text = ""
    ListView2.ColumnHeaders(9).Text = ""
    ListView2.ColumnHeaders(10).Text = ""
    ListView2.ColumnHeaders(11).Text = ""
    ListView2.ColumnHeaders(12).Text = ""
    ListView2.ColumnHeaders(13).Text = ""
    ListView2.ColumnHeaders(14).Text = ""
    ListView2.ColumnHeaders(15).Text = ""
    ListView2.ColumnHeaders(16).Text = ""
    ListView2.ColumnHeaders(17).Text = ""
    ListView2.ColumnHeaders(18).Text = ""
    ListView2.ColumnHeaders(19).Text = ""
    ListView2.ColumnHeaders(20).Text = ""
    ListView2.ColumnHeaders(21).Text = ""
    ListView2.ColumnHeaders(22).Text = ""
    ListView2.ColumnHeaders(23).Text = ""
    ListView2.ColumnHeaders(24).Text = ""
    ListView2.ColumnHeaders(25).Text = ""
    ListView2.ColumnHeaders(26).Text = ""
    ListView2.ColumnHeaders(27).Text = ""
    ListView2.ColumnHeaders(28).Text = ""
    ListView2.ColumnHeaders(29).Text = ""
    ListView2.ColumnHeaders(30).Text = ""
    ListView2.ColumnHeaders(31).Text = ""
    ListView2.ColumnHeaders(32).Text = ""
    ListView2.ColumnHeaders(33).Text = ""
    ListView2.ColumnHeaders(34).Text = ""
    ListView2.ColumnHeaders(35).Text = ""
    ListView2.ColumnHeaders(36).Text = ""
    ListView2.ColumnHeaders(37).Text = ""
    ListView2.ColumnHeaders(38).Text = ""
    ListView2.ColumnHeaders(39).Text = ""
    ListView2.ColumnHeaders(40).Text = ""
    ListView2.ColumnHeaders(41).Text = ""
    ListView2.ColumnHeaders(42).Text = ""
    ListView2.ColumnHeaders(43).Text = ""
    ListView2.ColumnHeaders(44).Text = ""
    ListView2.ColumnHeaders(45).Text = ""
    ListView2.ColumnHeaders(46).Text = ""
    ListView2.ColumnHeaders(47).Text = ""
    ListView2.ColumnHeaders(48).Text = ""
    ListView2.ColumnHeaders(49).Text = ""
    ListView2.ColumnHeaders(50).Text = ""
    ListView2.ColumnHeaders(51).Text = ""
    ListView2.ColumnHeaders(52).Text = ""
    ListView2.ColumnHeaders(53).Text = ""
    ListView2.ColumnHeaders(54).Text = ""
    ListView2.ColumnHeaders(55).Text = ""
    ListView2.ColumnHeaders(56).Text = ""
    ListView2.ColumnHeaders(57).Text = ""
    ListView2.ColumnHeaders(58).Text = ""
    ListView2.ColumnHeaders(59).Text = ""
    ListView2.ColumnHeaders(60).Text = ""
    ListView2.ColumnHeaders(61).Text = ""
    ListView2.ColumnHeaders(62).Text = ""
    ListView2.ColumnHeaders(63).Text = ""
End Sub


Private Sub Form_Load()
    intExit = 0
    Call ComboListing
    Call ListViewHeader
'    Label17.Visible = True
'    Label4.Visible = False
    ListView2.Visible = True
    'ListView7.Visible = False
    'lstSexList.Visible = False
    lstRaceList.Visible = False
    'lstTopResult.Visible = False
End Sub

' =================== Command Button Click =============================

Private Sub CmdExit_Click()
Dim RecordExit
    RecordExit = MsgBox("Quit this form?", vbYesNo + vbExclamation)
    If RecordExit = vbYes Then
        bExit = True
        Unload Me
    End If
End Sub

Private Sub CmdPrint_Click()
    Screen.MousePointer = vbHourglass
        'Select Case SSTab1.Tab
            'Case 0
                'If Opt1.value = True Then
                    'Call ExportToExcelAnalysis1(lstTopResult)
                'ElseIf Opt2.value = True Then
                    'Call ExportToExcelAnalysis1(lstRaceList)
                'ElseIf Opt3.value = True Then
                    'Call ExportToExcelAnalysis1(lstSexList)
                'ElseIf Opt4.value = True Then
                    'Call ExportToExcelAnalysis1(ListView7)
                'End If
            'Case 1
                'Call ExportToExcelAnalysis1(lstAgeGroupList)
            'Case 2
                'If OptAdmit.value = True Then
                    'Call ExportToExcelAnalysis1(lstDailyList)
                'ElseIf OptDischarge.value = True Then
                    'Call ExportToExcelAnalysis1(ListView9)
                'End If
            'Case 3
                'Call ExportToExcelAnalysis1(lstMonthList)
            'Case 4
                'Call ExportToExcelAnalysis1(lstYearList)
        'End Select
        Call ExportToExcelAnalysis1(ListView2)
    Screen.MousePointer = vbDefault
End Sub

Private Sub cmdSearch_Click()
Dim strAppend As String
    strAppend = ""
    Screen.MousePointer = vbHourglass
    Call SearchRecords
    Screen.MousePointer = vbDefault
End Sub

Public Function ClearForm(frmCall)
On Error Resume Next
Dim ctl As Control

For Each ctl In frmCall.Controls
    If TypeOf ctl Is TextBox Then
        ctl.Text = ""
    ElseIf TypeOf ctl Is ListBox Then
            ctl.Clear
    ElseIf TypeOf ctl Is ComboBox Then
            ctl.Text = ""
    ElseIf TypeOf ctl Is CheckBox Then
            ctl.value = False
    ElseIf TypeOf ctl Is MaskEdBox Then
        ctl.Text = "__/__/____"
    
    End If
Next ctl
End Function

Private Sub CmdClear_Click()
    Call ClearForm(Me)
    lblCurrent = ""
    lblTotalRecord = ""
    ListView2.listitems.Clear
End Sub

Private Sub cmdSearchICD_Click()
    frmSearchICD.Source = 20
    frmSearchICD.show vbModal
End Sub


Private Sub cboSelection_Change()
    If InStr(cboSelection.Text, "null") Then
       cboSelection.Enabled = False
    End If
    
    If Mid(cboSelection, 1, 1) = "0" Or Mid(cboSelection, 1, 1) = "1" Then
        txtFinalICD.Text = ""
        txtFinalICD.Enabled = False
        cmdSearchICD.Enabled = False
        Frame4.Enabled = False
        Frame6.Enabled = False
    Else
        txtFinalICD.Text = ""
        txtFinalICD.Enabled = True
        cmdSearchICD.Enabled = True
        Frame4.Enabled = True
        Frame6.Enabled = True
    End If
End Sub

Private Sub cboSelection_Click()
    If InStr(cboSelection.Text, "null") Then
       cboSelection.Enabled = False
    End If
    
    If Mid(cboSelection, 1, 1) = "0" Or Mid(cboSelection, 1, 1) = "1" Then
        txtFinalICD.Text = ""
        txtFinalICD.Enabled = False
        cmdSearchICD.Enabled = False
    Else
        txtFinalICD.Text = ""
        txtFinalICD.Enabled = True
        cmdSearchICD.Enabled = True
    End If
End Sub
' =================== Command Button Click / Change ==========================


'============== ComboBox Listing =================================
Private Sub ComboListing()
Dim PAY As New ADODB.Recordset
Dim INS As New ADODB.Recordset
Dim REL As New ADODB.Recordset
Dim OCP As New ADODB.Recordset
Dim PolStatus As New ADODB.Recordset
Dim State As New ADODB.Recordset
Dim HOS As New ADODB.Recordset
Dim HOSGrp As New ADODB.Recordset

    INS.Open "Select INSCode, INSDescription from INS", medixmasterconn, adOpenKeyset, adLockOptimistic
    PAY.Open "Select PAYCode, PAYCompanyName from PAY", medixmasterconn, adOpenKeyset, adLockOptimistic
    PolStatus.Open "Select pstatuscode ,pstatusDesc from PolicyStatus", medixmasterconn, adOpenKeyset, adLockOptimistic
    REL.Open "Select RELCode, RELDescription from REL", medixmasterconn, adOpenKeyset, adLockOptimistic
    OCP.Open "Select OCPCode, OCPDescription from OCP", medixmasterconn, adOpenKeyset, adLockOptimistic
    State.Open "Select STADescription from STA", medixmasterconn, adOpenKeyset, adLockOptimistic
    HOS.Open "Select HOSCode, HOSName from HOS", medixmasterconn, adOpenKeyset, adLockOptimistic
    HOSGrp.Open "Select HOSGRPCode, HOSGRPName from HOSGRP", medixmasterconn, adOpenKeyset, adLockOptimistic
    
    
    cboHLTCode.AddItem "S" & " - " & "Sihat"
    cboHLTCode.AddItem "N" & " - " & "Non-Sihat"
    
    Do Until INS.EOF
        CboInsType.AddItem ChkStr(INS!INSCode) & " - " & ChkStr(INS!INSDescription)
        INS.MoveNext
    Loop
    INS.Close
    Set INS = Nothing
    
    CboMemType.AddItem "Both"
    CboMemType.AddItem "P" & " - " & "Principal"
    CboMemType.AddItem "S" & " - " & "Supplementary"
    CboMemType.Text = "Both"
    
    CboAdultChild.AddItem "A - Adult"
    CboAdultChild.AddItem "C - Child"
    
    CboMarital.AddItem "S - Single"
    CboMarital.AddItem "M - Married"
    
    CboRace.AddItem "C - Chinese"
    CboRace.AddItem "I - Indian"
    CboRace.AddItem "M - Malay"
    CboRace.AddItem "O - Other"
            
    CboSex.AddItem "M" & " - " & "Male"
    CboSex.AddItem "F" & " - " & "Female"
    
    Do Until REL.EOF
        CboRelation.AddItem ChkStr(REL!RELCode) & " - " & ChkStr(REL!RELDescription)
        REL.MoveNext
    Loop
    REL.Close
    Set REL = Nothing
    
    cboDataStatus.AddItem "A - Active"
    cboDataStatus.AddItem "D - Delete"
    
    With CboClaimability
        .AddItem "A - Accepted"
        .AddItem "R - Rejected"
    End With
    
    Do Until PolStatus.EOF
        CboStatus.AddItem (PolStatus!pstatuscode) & " - " & (PolStatus!pstatusDesc)
        PolStatus.MoveNext
    Loop
    PolStatus.Close
    Set PolStatus = Nothing
    
    With CboCaseStatus
        .AddItem "O - Open"
        .AddItem "C - Close"
        .AddItem "D - Delete"
        .AddItem "A - All"
    End With
    
    With CboWsStatus
        .AddItem "O - Open"
        .AddItem "C - Close"
    End With
    
    Do Until OCP.EOF
        CboOccupation.AddItem ChkStr(OCP!OCPCode) & " - " & ChkStr(OCP!OCPDescription)
        OCP.MoveNext
    Loop
    OCP.Close
    Set OCP = Nothing
        
    CboPreMode.AddItem "A - Annually"
    CboPreMode.AddItem "Q - Quarterly"
    CboPreMode.AddItem "S - Semi Annual"
    CboPreMode.AddItem "M - Monthly"
        
    Do Until State.EOF
        CboState.AddItem ChkStr(State!STADescription)
        State.MoveNext
    Loop
    State.Close
    Set State = Nothing
        
    cboRenewalInd.AddItem "N - No"
    cboRenewalInd.AddItem "Y - Yes"
    
    cboTakeOver.AddItem "N - No"
    cboTakeOver.AddItem "Y - Yes"
           
    Do Until PAY.EOF
        cboPayorCode.AddItem PAY!PAYCode & " - " & Trim(PAY!PAYCompanyName)
        PAY.MoveNext
    Loop
    PAY.Close
    Set PAY = Nothing
    
    cboSuppType.AddItem "Both"
    cboSuppType.AddItem "C" & " - " & "Combined Limit"
    cboSuppType.AddItem "S" & " - " & "Separated Limit"
    cboSuppType.Text = "C" & " - " & "Combined Limit"
    
    Do Until HOS.EOF
        cboHOSName.AddItem Trim(HOS!HOSName)
        HOS.MoveNext
    Loop
    HOS.Close
    Set HOS = Nothing
    
    Do Until HOSGrp.EOF
        cboHOSGrp.AddItem Trim(HOSGrp!HOSGRPName)
        HOSGrp.MoveNext
    Loop
    HOSGrp.Close
    Set HOSGrp = Nothing
    
    'Combo1.AddItem "0" & " - " & "Detail Listing"
    'Combo1.AddItem "1" & " - " & "General Summary"
    'Combo1.AddItem "2" & " - " & "Top 20 by Cases"
    'Combo1.AddItem "3" & " - " & "Top 20 by Amount"
    'Combo1.Text = "0" & " - " & "Detail Listing"
    
End Sub
'============== ComboBox Listing =================================

' =================== Search Diag. By Ageband Results ======================

Private Function SearchDiagByAgeband()
Dim sql As String
    
    strAppend = ""
    
        If Len(Trim(txtFinalICD)) Then
            strAppend = strAppend + " And FinalDiag = '" & Trim(txtFinalICD) & "'"
        End If
        
        If Option1.value = True Then
            sql = ""
            sql = "AND CASStatus = 'O' "
            strAppend = strAppend + sql
        End If
        
        If Option2.value = True Then
            sql = ""
            sql = "AND CASStatus = 'C'"
            strAppend = strAppend + sql
        End If
        
        If Option3.value = True Then
            sql = ""
            sql = "AND CASStatus = 'D'"
            strAppend = strAppend + sql
        End If
        
        If Option4.value = True Then
            sql = ""
            sql = "AND (CASStatus = 'O' OR CASStatus = 'C')"
            strAppend = strAppend + sql
        End If
        
        If Option5.value = True Then
            sql = ""
            sql = "AND CASClaimability = 'A'"
            strAppend = strAppend + sql
        End If
        
        If Option6.value = True Then
            sql = ""
            sql = "AND CASClaimability = 'R'"
            strAppend = strAppend + sql
        End If
        
        Call DiagByAgebandListing(strAppend)

End Function
' =================== Search Diag. By Ageband Results ======================

' =================== Search Diag By Race Results ======================

Private Function SearchDiagByRace()
Dim sql As String
    
        strAppend = ""
    
        If Len(Trim(txtFinalICD)) Then
            strAppend = strAppend + " And FinalDiag = '" & Trim(txtFinalICD) & "'"
        End If
        
        If Option1.value = True Then
            sql = ""
            sql = "AND CASStatus = 'O' "
            strAppend = strAppend + sql
        End If
        
        If Option2.value = True Then
            sql = ""
            sql = "AND CASStatus = 'C'"
            strAppend = strAppend + sql
        End If
        
        If Option3.value = True Then
            sql = ""
            sql = "AND CASStatus = 'D'"
            strAppend = strAppend + sql
        End If
        
        If Option4.value = True Then
            sql = ""
            sql = "AND (CASStatus = 'O' OR CASStatus = 'C')"
            strAppend = strAppend + sql
        End If
        
        If Option5.value = True Then
            sql = ""
            sql = "AND CASClaimability = 'A'"
            strAppend = strAppend + sql
        End If
        
        If Option6.value = True Then
            sql = ""
            sql = "AND CASClaimability = 'R'"
            strAppend = strAppend + sql
        End If
        
        Call DiagByRaceListing(strAppend)

End Function
' =================== Search Diag By Race Results ======================

' =================== Search Diag. By Sex Results ======================

Private Function SearchDiagBySex()
    
    strAppend = ""
    
        If Len(Trim(txtFinalICD)) Then
            strAppend = strAppend + " And FinalDiag = '" & Trim(txtFinalICD) & "'"
        End If
        
        If Option1.value = True Then
            sql = ""
            sql = "AND CASStatus = 'O' "
            strAppend = strAppend + sql
        End If
        
        If Option2.value = True Then
            sql = ""
            sql = "AND CASStatus = 'C'"
            strAppend = strAppend + sql
        End If
        
        If Option3.value = True Then
            sql = ""
            sql = "AND CASStatus = 'D'"
            strAppend = strAppend + sql
        End If
        
        If Option4.value = True Then
            sql = ""
            sql = "AND (CASStatus = 'O' OR CASStatus = 'C')"
            strAppend = strAppend + sql
        End If
        
        If Option5.value = True Then
            sql = ""
            sql = "AND CASClaimability = 'A'"
            strAppend = strAppend + sql
        End If
        
        If Option6.value = True Then
            sql = ""
            sql = "AND CASClaimability = 'R'"
            strAppend = strAppend + sql
        End If
        
        Call SearchDiagBySexListing(strAppend)

End Function
' =================== Search Diag By Sex Results ======================

' =================== Search Sex with Ageband Selected Results ======================
'Private Function SearchSexWithAgeband()
    
    'strAppend = ""
    
    'If Len(Trim(CboAgeGroup)) Then
        'strAppend = strAppend + " And Sex = '" & Trim(Mid(CboAgeGroup, 1, IIf(InStr(1, CboAgeGroup, "-") = 0, 4, InStr(1, CboAgeGroup, "-") - 1))) & "'"
    'End If
    
    'If Option8.value = True Then
            'sql = ""
            'sql = "AND CASStatus = 'O' "
            'strAppend = strAppend + sql
        'End If
        
        'If Option9.value = True Then
            'sql = ""
            'sql = "AND CASStatus = 'C'"
            'strAppend = strAppend + sql
        'End If
        
        'If Option10.value = True Then
            'sql = ""
            'sql = "AND CASStatus = 'D'"
            'strAppend = strAppend + sql
        'End If
        
        'If Option11.value = True Then
            'sql = ""
            'sql = "AND (CASStatus = 'O' OR CASStatus = 'C')"
            'strAppend = strAppend + sql
        'End If
        
        'If Option12.value = True Then
            'sql = ""
            'sql = "AND CASClaimability = 'A'"
            'strAppend = strAppend + sql
        'End If
        
        'If Option13.value = True Then
            'sql = ""
            'sql = "AND CASClaimability = 'R'"
            'strAppend = strAppend + sql
        'End If
'End Function

' =================== Search Sex with Ageband Selected Results ======================
Private Function SearchDiagSexAge()
    
    strAppend = ""
    
        If Len(Trim(txtFinalICD)) Then
            strAppend = strAppend + " And FinalDiag = '" & Trim(txtFinalICD) & "'"
        End If
        
        If Option1.value = True Then
            sql = ""
            sql = "AND CASStatus = 'O' "
            strAppend = strAppend + sql
        End If
        
        If Option2.value = True Then
            sql = ""
            sql = "AND CASStatus = 'C'"
            strAppend = strAppend + sql
        End If
        
        If Option3.value = True Then
            sql = ""
            sql = "AND CASStatus = 'D'"
            strAppend = strAppend + sql
        End If
        
        If Option4.value = True Then
            sql = ""
            sql = "AND (CASStatus = 'O' OR CASStatus = 'C')"
            strAppend = strAppend + sql
        End If
        
        If Option5.value = True Then
            sql = ""
            sql = "AND CASClaimability = 'A'"
            strAppend = strAppend + sql
        End If
        
        If Option6.value = True Then
            sql = ""
            sql = "AND CASClaimability = 'R'"
            strAppend = strAppend + sql
        End If
        
        Call DiagSexAgeListing(strAppend)

End Function
' =================== Search Sex with Ageband Selected Results ======================

' =================== Search By Final Diag. Results ======================

Private Function SearchFinalDiag()
    
  '  strappend = ""
    
    If Len(Trim(txtFinalICD)) Then
            strAppend = strAppend + " And FinalDiag = '" & Trim(txtFinalICD) & "'"
        End If
        
        If Option1.value = True Then
            sql = ""
            sql = "AND CASStatus = 'O' "
            strAppend = strAppend + sql
        End If
        
        If Option2.value = True Then
            sql = ""
            sql = "AND CASStatus = 'C'"
            strAppend = strAppend + sql
        End If
        
        If Option3.value = True Then
            sql = ""
            sql = "AND CASStatus = 'D'"
            strAppend = strAppend + sql
        End If
        
        If Option4.value = True Then
            sql = ""
            sql = "AND (CASStatus = 'O' OR CASStatus = 'C')"
            strAppend = strAppend + sql
        End If
        
        If Option5.value = True Then
            sql = ""
            sql = "AND CASClaimability = 'A'"
            strAppend = strAppend + sql
        End If
        
        If Option6.value = True Then
            sql = ""
            sql = "AND CASClaimability = 'R'"
            strAppend = strAppend + sql
        End If
    
    Call FinalDiagListing(strAppend)

End Function
' =================== Search By Final Diag. Results ======================


' =================== Search Diag By Post Code Results ======================

Private Function SearchDiagByPCode()
    
    strAppend = ""
    
    If Len(Trim(txtFinalICD)) Then
            strAppend = strAppend + " And FinalDiag = '" & Trim(txtFinalICD) & "'"
        End If
        
        If Option1.value = True Then
            sql = ""
            sql = "AND CASStatus = 'O' "
            strAppend = strAppend + sql
        End If
        
        If Option2.value = True Then
            sql = ""
            sql = "AND CASStatus = 'C'"
            strAppend = strAppend + sql
        End If
        
        If Option3.value = True Then
            sql = ""
            sql = "AND CASStatus = 'D'"
            strAppend = strAppend + sql
        End If
        
        If Option4.value = True Then
            sql = ""
            sql = "AND (CASStatus = 'O' OR CASStatus = 'C')"
            strAppend = strAppend + sql
        End If
        
        If Option5.value = True Then
            sql = ""
            sql = "AND CASClaimability = 'A'"
            strAppend = strAppend + sql
        End If
        
        If Option6.value = True Then
            sql = ""
            sql = "AND CASClaimability = 'R'"
            strAppend = strAppend + sql
        End If
    
    Call DiagByPCodeListing(strAppend)

End Function
' =================== Search Diag By Post Code Results ======================

' ============== List Diag. Code By Ageband ==============================
Private Sub DiagByAgebandListing(Optional strAppend As String)
Dim rst2 As New ADODB.Recordset
Dim Record As ListItem
Dim sql As String
Dim sql2 As String
       
    sql = ""
    lstTopResult.listitems.Clear
    
    If Mid(cboSelection, 1, 1) = "0" Then
        sql = "TOP 10 "
    End If
    If Mid(cboSelection, 1, 1) = "1" Then
        sql = "TOP 20 "
    End If
    
    sql = "SELECT " & sql & " FinalDiag, COUNT(FinalDiag) AS Total" & _
          " From VIEW_AnalysisReportNew2 Where 0 = 0 " & strAppend & " GROUP BY FinalDiag ORDER BY Total DESC"
    
    rst2.Open sql, medixmasterconn, adOpenKeyset, adLockOptimistic
    
        Do Until rst2.EOF
    
            With rst2
                Set Record = lstTopResult.listitems.Add(, , rst2!FinalDiag)
                DiagCode = Trim(rst2!FinalDiag)
                
                Call CountAgebandRecords(strAppend)
                
                If Len(Description) Then
                    Record.SubItems(1) = Trim(Description)
                End If
                
                If Len(rst2!Total) Then
                    Record.SubItems(2) = rst2!Total
                End If
                
                If Len(TotalAgeCode1) Then
                   Record.SubItems(3) = TotalAgeCode1
                End If
                If Len(TotalAgeCode2) Then
                    Record.SubItems(4) = TotalAgeCode2
                End If
                                
                If Len(TotalAgeCode3) Then
                   Record.SubItems(5) = TotalAgeCode3
                End If
                If Len(TotalAgeCode4) Then
                    Record.SubItems(6) = TotalAgeCode4
                End If
                
                If Len(TotalAgeCode5) Then
                   Record.SubItems(7) = TotalAgeCode5
                End If
                If Len(TotalAgeCode6) Then
                    Record.SubItems(8) = TotalAgeCode6
                End If
                If Len(TotalAgeCode7) Then
                    Record.SubItems(9) = TotalAgeCode7
                End If
                
            End With
            rst2.MoveNext
        Loop
    rst2.Close
    Set rst2 = Nothing
End Sub


' =================== Search Diag. BY Ageband Top 10/20 Results ======================
Private Sub DiagByAgebandTop1020()
Dim rst2 As New ADODB.Recordset
Dim Record As ListItem
Dim sql As String
       
    lstTopResult.listitems.Clear
            
    If Mid(cboSelection, 1, 1) = "0" Then
    
    sql = " SELECT TOP 10 FinalDiag, COUNT(FinalDiag) AS Total " & _
          " From VIEW_AnalysisReportNew2 " & _
          " Where (CASStatus = 'O' OR CASStatus = 'C') AND CASClaimability = 'A' " & _
          " GROUP BY FinalDiag ORDER BY Total DESC "
    
    ElseIf Mid(cboSelection, 1, 1) = "1" Then
    
    sql = " SELECT TOP 20 FinalDiag, COUNT(FinalDiag) AS Total " & _
          " From VIEW_AnalysisReportNew2 " & _
          " Where (CASStatus = 'O' OR CASStatus = 'C') AND CASClaimability = 'A' " & _
          " GROUP BY FinalDiag ORDER BY Total DESC "
          
    End If
    
    rst2.Open sql, medixmasterconn, adOpenKeyset, adLockOptimistic
    
        Do Until rst2.EOF
    
            With rst2
                Set Record = lstTopResult.listitems.Add(, , rst2!FinalDiag)
                DiagCode = Trim(rst2!FinalDiag)
                
                Call CountAgebandRecords
                
                If Len(Description) Then
                    Record.SubItems(1) = Trim(Description)
                End If
                
                If Len(rst2!Total) Then
                    Record.SubItems(2) = rst2!Total
                End If
                
                If Len(TotalAgeCode1) Then
                   Record.SubItems(3) = TotalAgeCode1
                End If
                If Len(TotalAgeCode2) Then
                    Record.SubItems(4) = TotalAgeCode2
                End If
                                
                If Len(TotalAgeCode3) Then
                   Record.SubItems(5) = TotalAgeCode3
                End If
                If Len(TotalAgeCode4) Then
                    Record.SubItems(6) = TotalAgeCode4
                End If
                
                If Len(TotalAgeCode5) Then
                   Record.SubItems(7) = TotalAgeCode5
                End If
                If Len(TotalAgeCode6) Then
                    Record.SubItems(8) = TotalAgeCode6
                End If
                If Len(TotalAgeCode7) Then
                    Record.SubItems(9) = TotalAgeCode7
                End If
                
            End With
            rst2.MoveNext
        Loop
    rst2.Close
    Set rst2 = Nothing
End Sub
' =================== Search Diag. By Ageband Top 10/20 Results ======================

'================== Start Counting AgeBand ============================
Private Sub CountAgebandRecords(Optional strAppend As String)
Dim rst2 As New ADODB.Recordset
Dim Record As ListItem
Dim sql As String
Dim sql2 As String
Dim Current1 As String
Dim Current2 As String
Dim Current3 As String
Dim Current4 As String
Dim Current5 As String
Dim Current6 As String
Dim Current7 As String

    Current1 = 0
    Current2 = 0
    Current3 = 0
    Current4 = 0
    Current5 = 0
    Current6 = 0
    Current7 = 0
    
    TotalAgeCode1 = 0
    TotalAgeCode2 = 0
    TotalAgeCode3 = 0
    TotalAgeCode4 = 0
    TotalAgeCode5 = 0
    TotalAgeCode6 = 0
    TotalAgeCode7 = 0
        
    ListView1.listitems.Clear
            
    sql = " Select * From VIEW_AnalysisReportNew2 where FinalDiag = '" & Trim(DiagCode) & "'"
        
    sql2 = " AND (CASStatus = 'O' OR CASStatus = 'C') AND CASClaimability = 'A' "
    
    If Mid(cboSelection, 1, 1) = "2" Then
    
        If Len(Trim(strAppend)) Then
            sql = sql + strAppend
        End If
    
    ElseIf Mid(cboSelection, 1, 1) = "0" Or Mid(cboSelection, 1, 1) = "1" Then
        sql = sql + sql2
    End If
    
    
    rst2.Open sql, medixmasterconn, adOpenKeyset, adLockOptimistic
    
        Do Until rst2.EOF
                    
            With rst2
                Set Record = ListView1.listitems.Add(, , rst2!ICDDescription)
                    Description = rst2!ICDDescription
                
                If rst2!AgeBand = "01" Then
                   Record.SubItems(1) = rst2!AgeBand
                   Current1 = Current1 + 1
                   TotalAgeCode1 = Current1
                ElseIf rst2!AgeBand = "02" Then
                   Record.SubItems(1) = rst2!AgeBand
                   Current2 = Current2 + 1
                   TotalAgeCode2 = Current2
                ElseIf rst2!AgeBand = "03" Then
                   Record.SubItems(1) = rst2!AgeBand
                   Current3 = Current3 + 1
                   TotalAgeCode3 = Current3
                ElseIf rst2!AgeBand = "04" Then
                   Record.SubItems(1) = rst2!AgeBand
                   Current4 = Current4 + 1
                   TotalAgeCode4 = Current4
                ElseIf rst2!AgeBand = "05" Then
                   Record.SubItems(1) = rst2!AgeBand
                   Current5 = Current5 + 1
                   TotalAgeCode5 = Current5
                ElseIf rst2!AgeBand = "06" Then
                   Record.SubItems(1) = rst2!AgeBand
                   Current6 = Current6 + 1
                   TotalAgeCode6 = Current6
                Else
                   Current7 = Current7 + 1
                   TotalAgeCode7 = Current7
                End If
            End With
            rst2.MoveNext
        Loop
    
    rst2.Close
    Set rst2 = Nothing
End Sub


' ============== List All Diag By Race Code ==============================
Private Sub DiagByRaceListing(Optional strAppend As String)
Dim rst2 As New ADODB.Recordset
Dim Record As ListItem
Dim sql As String
Dim sql2 As String
           
    sql = ""
    lstRaceList.listitems.Clear
    
    If Mid(cboSelection, 1, 1) = "0" Then
        sql = "TOP 10 "
    End If
    If Mid(cboSelection, 1, 1) = "1" Then
        sql = "TOP 20 "
    End If
    
    sql = " SELECT " & sql & "FinalDiag, COUNT(FinalDiag) AS Total" & _
          " From VIEW_AnalysisReportNew2 Where 0 = 0 " & _
          " GROUP BY FinalDiag ORDER BY Total DESC "
    
    rst2.Open sql, medixmasterconn, adOpenKeyset, adLockOptimistic
    
        Do Until rst2.EOF
    
            With rst2
                Set Record = lstRaceList.listitems.Add(, , rst2!FinalDiag)
                DiagCode = Trim(rst2!FinalDiag)
                
                Call CountRaceRecords(strAppend)
                
                If Len(Description) Then
                    Record.SubItems(1) = Trim(Description)
                End If
                
                If Len(rst2!Total) Then
                    Record.SubItems(2) = rst2!Total
                End If
                
                If Len(TotalRace1) Then
                   Record.SubItems(3) = TotalRace1
                End If
                If Len(TotalRace2) Then
                    Record.SubItems(4) = TotalRace2
                End If
                                
                If Len(TotalRace3) Then
                   Record.SubItems(5) = TotalRace3
                End If
                If Len(TotalRace4) Then
                    Record.SubItems(6) = TotalRace4
                End If
                
                If Len(TotalRace5) Then
                   Record.SubItems(7) = TotalRace5
                End If
                
            End With
            rst2.MoveNext
        Loop
    
    rst2.Close
    Set rst2 = Nothing
End Sub


' =================== Search Diag By Race Top 10/20 Results ======================
Private Sub DiagByRaceTop1020()
Dim rst2 As New ADODB.Recordset
Dim Record As ListItem
Dim sql As String
       
    lstRaceList.listitems.Clear
            
    If Mid(cboSelection, 1, 1) = "0" Then
    
        sql = " SELECT TOP 10 FinalDiag, COUNT(FinalDiag) AS Total " & _
              " From VIEW_AnalysisReportNew2 " & _
              " Where (CASStatus = 'O' OR CASStatus = 'C') AND CASClaimability = 'A' " & _
              " GROUP BY FinalDiag ORDER BY Total DESC "
         
    ElseIf Mid(cboSelection, 1, 1) = "1" Then
        sql = " SELECT TOP 20 FinalDiag, COUNT(FinalDiag) AS Total " & _
              " From VIEW_AnalysisReportNew2 " & _
              " Where (CASStatus = 'O' OR CASStatus = 'C') AND CASClaimability = 'A' " & _
              " GROUP BY FinalDiag ORDER BY Total DESC "
    End If
    
    rst2.Open sql, medixmasterconn, adOpenKeyset, adLockOptimistic
    
        Do Until rst2.EOF
                    
            With rst2
                Set Record = lstRaceList.listitems.Add(, , rst2!FinalDiag)
                DiagCode = Trim(rst2!FinalDiag)
                
                Call CountRaceRecords
                
                If Len(Description) Then
                    Record.SubItems(1) = Trim(Description)
                End If
                
                If Len(rst2!Total) Then
                    Record.SubItems(2) = rst2!Total
                End If
                
                If Len(TotalRace1) Then
                   Record.SubItems(3) = TotalRace1
                End If
                If Len(TotalRace2) Then
                    Record.SubItems(4) = TotalRace2
                End If
                                
                If Len(TotalRace3) Then
                   Record.SubItems(5) = TotalRace3
                End If
                If Len(TotalRace4) Then
                    Record.SubItems(6) = TotalRace4
                End If
                
                If Len(TotalRace5) Then
                   Record.SubItems(7) = TotalRace5
                End If
            End With
            rst2.MoveNext
        Loop
    
    rst2.Close
    Set rst2 = Nothing
End Sub
' =================== Search Race Top 10/20 Results ======================

' ============== List All Diag By Sex Code ==============================
Private Sub SearchDiagBySexListing(Optional strAppend As String)
Dim rst2 As New ADODB.Recordset
Dim Record As ListItem
Dim sql As String
Dim sql2 As String
       
    lstSexList.listitems.Clear
    
    sql = " SELECT FinalDiag, COUNT(FinalDiag) AS Total" & _
          " From VIEW_AnalysisReportNew2 Where 0 = 0 "
    
    sql2 = " GROUP BY FinalDiag ORDER BY Total DESC "
    
    If Len(Trim(strAppend)) Then
        sql = sql + strAppend + sql2
    Else
        sql = sql + sql2
    End If
    
    rst2.Open sql, medixmasterconn, adOpenKeyset, adLockOptimistic
    
        Do Until rst2.EOF
    
            With rst2
                Set Record = lstSexList.listitems.Add(, , rst2!FinalDiag)
                DiagCode = Trim(rst2!FinalDiag)
                
                Call CountSexRecords(strAppend)
                
                If Len(Description) Then
                    Record.SubItems(1) = Trim(Description)
                End If
                
                If Len(rst2!Total) Then
                    Record.SubItems(2) = rst2!Total
                End If
                
                If Len(TotalSex1) Then
                   Record.SubItems(3) = TotalSex1
                End If
                If Len(TotalSex2) Then
                    Record.SubItems(4) = TotalSex2
                End If
                                
                If Len(TotalSex3) Then
                   Record.SubItems(5) = TotalSex3
                End If
                        End With
            rst2.MoveNext
        Loop
    
    rst2.Close
    Set rst2 = Nothing
End Sub

' =================== Search Diag By Sex Top 10/20 Results ======================
Private Sub DiagBySexTop1020()
Dim rst2 As New ADODB.Recordset
Dim Record As ListItem
Dim sql As String
       
    lstSexList.listitems.Clear
            
    If Mid(cboSelection, 1, 1) = "0" Then
    
        sql = " SELECT TOP 10 FinalDiag, COUNT(FinalDiag) AS Total " & _
              " From VIEW_AnalysisReportNew2 " & _
              " Where (CASStatus = 'O' OR CASStatus = 'C') AND CASClaimability = 'A' " & _
              " GROUP BY FinalDiag ORDER BY Total DESC "
         
    ElseIf Mid(cboSelection, 1, 1) = "1" Then
        sql = " SELECT TOP 20 FinalDiag, COUNT(FinalDiag) AS Total " & _
              " From VIEW_AnalysisReportNew2 " & _
              " Where (CASStatus = 'O' OR CASStatus = 'C') AND CASClaimability = 'A' " & _
              " GROUP BY FinalDiag ORDER BY Total DESC "
    End If
    
    rst2.Open sql, medixmasterconn, adOpenKeyset, adLockOptimistic
    
        Do Until rst2.EOF
                    
            With rst2
                Set Record = lstSexList.listitems.Add(, , rst2!FinalDiag)
                DiagCode = Trim(rst2!FinalDiag)
                
                Call CountSexRecords
                
                If Len(Description) Then
                    Record.SubItems(1) = Trim(Description)
                End If
                
                If Len(rst2!Total) Then
                    Record.SubItems(2) = rst2!Total
                End If
                
                If Len(TotalSex1) Then
                   Record.SubItems(3) = TotalSex1
                End If
                If Len(TotalSex2) Then
                    Record.SubItems(4) = TotalSex2
                End If
                                
                If Len(TotalSex3) Then
                   Record.SubItems(5) = TotalSex3
                End If
                End With
            rst2.MoveNext
        Loop
    
    rst2.Close
    Set rst2 = Nothing
End Sub
' =================== Search Sex Top 10/20 Results ======================


'==================== List Final Diag. Listing Resullts ==========================
Private Sub FinalDiagListing(Optional strAppend As String)
Dim rst2 As New ADODB.Recordset
Dim Record As ListItem
Dim sql As String
       
    sql = ""
    ListView2.listitems.Clear
    
    sql = "SELECT Top 20 COUNT(FinalDiag) AS Total, FinalDiag, Sum(CLMTotalApproved) As TotalApproved, " & _
          " Sum(CLMTotalActual) As TotalActual From VIEW_AnalysisReportNew Where 0 = 0 " & strAppend & _
          " GROUP BY FinalDiag " & _
          " ORDER BY  FinalDiag "
    
    rst2.Open sql, medixmasterconn, adOpenKeyset, adLockOptimistic
       
        Do Until rst2.EOF
    
            With rst2
                Set Record = ListView2.listitems.Add(, , rst2!FinalDiag)
                DiagCode = Trim(rst2!FinalDiag)
                
                Call CountFinalDiag(strAppend)
                
                If Len(Description) Then
                    Record.SubItems(1) = Trim(Description)
                End If
                
                'If Len(rst2!Total) Then
                '    Record.SubItems(2) = rst2!Total
                'End If
                
                If Len(TotalDuration) Then
                   Record.SubItems(3) = TotalDuration
                End If
                
                If Len(rst2!TotalActual) Then
                   Record.SubItems(4) = rst2!TotalActual
                End If
                
                If Len(rst2!TotalApproved) Then
                    Record.SubItems(5) = rst2!TotalApproved
                End If
                                
                Record.SubItems(6) = Format(Round(TotalActual / rst2!Total, 0), "#,##0.00")
                Record.SubItems(7) = Format(Round(TotalApproved / rst2!Total, 0), "#,##0.00")
            
            End With
            rst2.MoveNext
        Loop
    rst2.Close
    Set rst2 = Nothing
End Sub

' =================== Search Final Diag. Top 10/20 Results ======================
Private Sub FinalDiagTop1020()
Dim rst2 As New ADODB.Recordset
Dim Record As ListItem
Dim sql As String
       
    lstFinalDiagList.listitems.Clear
    
    If Mid(cboSelection, 1, 1) = "0" Then
    
        sql = " SELECT TOP 10 FinalDiag, COUNT(FinalDiag) AS Total " & _
              " From VIEW_AnalysisReportNew1 " & _
              " Where (CASStatus = 'O' OR CASStatus = 'C') AND CASClaimability = 'A' " & _
              " GROUP BY FinalDiag ORDER BY Total DESC "
         
    ElseIf Mid(cboSelection, 1, 1) = "1" Then
        sql = " SELECT TOP 20 FinalDiag, COUNT(FinalDiag) AS Total " & _
              " From VIEW_AnalysisReportNew1 " & _
              " Where (CASStatus = 'O' OR CASStatus = 'C') AND CASClaimability = 'A' " & _
              " GROUP BY FinalDiag ORDER BY Total DESC "
    End If
    
    
    rst2.Open sql, medixmasterconn, adOpenKeyset, adLockOptimistic
    
        Do Until rst2.EOF
    
            With rst2
                Set Record = lstFinalDiagList.listitems.Add(, , rst2!FinalDiag)
                DiagCode = Trim(rst2!FinalDiag)
                
                Call CountFinalDiag
                
                If Len(Description) Then
                    Record.SubItems(1) = Trim(Description)
                End If
                
                If Len(rst2!Total) Then
                    Record.SubItems(2) = rst2!Total
                End If
                
                If Len(TotalDuration) Then
                   Record.SubItems(3) = TotalDuration
                End If
                
                If Len(TotalActual) Then
                   Record.SubItems(4) = TotalActual
                End If
                
                If Len(TotalApproved) Then
                    Record.SubItems(5) = TotalApproved
                End If
                                
                Record.SubItems(6) = Format(Round(TotalActual / rst2!Total, 0), "#,##0.00")
                Record.SubItems(7) = Format(Round(TotalApproved / rst2!Total, 0), "#,##0.00")
                
            End With
            rst2.MoveNext
        Loop
    rst2.Close
    Set rst2 = Nothing
End Sub
' =================== Search Final Diag. Top 10/20 Results ======================

' =================== Count Final Diag. Results ======================
Private Sub CountFinalDiag(Optional strAppend As String)
Dim rst2 As New ADODB.Recordset
Dim Record As ListItem
Dim sql As String
        
    TotalActual = 0
    TotalApproved = 0
    TotalDuration = 0
    'ListView2.listitems.Clear
        
    sql = " Select * From VIEW_AnalysisReportNew where FinalDiag = '" & Trim(DiagCode) & "'"
        
    'sql2 = " AND (CASStatus = 'O' OR CASStatus = 'C') AND CASClaimability = 'A' "
    
    'If Mid(cboSelection, 1, 1) = "2" Then
    
        'If Len(Trim(strAppend)) Then
        '    sql = sql + strAppend
        'End If
    
    'ElseIf Mid(cboSelection, 1, 1) = "0" Or Mid(cboSelection, 1, 1) = "1" Then
    '    sql = sql + sql2
    'End If

    'sql = strAppend
    rst2.Open sql, medixmasterconn, adOpenKeyset, adLockOptimistic
    
        Do Until rst2.EOF
        txtNoofDay = ""
            If IsDate(rst2!CASDischargeDate) And Len(rst2!CASDateAdmitted) Then
                txtNoofDay = DateValue(rst2!CASDischargeDate) - DateValue((rst2!CASDateAdmitted))
            End If
            
            If txtNoofDay = "0" Then
                txtNoofDay = txtNoofDay + 1
            End If
            
            With rst2
                'If Len(rst2!CASDischargeDate) Then
                '    Set Record = ListView2.listitems.Add(, , rst2!CASDischargeDate)
                'End If
                
                'If Len(rst2!CASDateAdmitted) Then
                '   Record.SubItems(1) = Format(rst2!CASDateAdmitted, "dd/mm/yyyy")
                'End If
                'If Len(txtNoofDay) Then
                '   Record.SubItems(2) = txtNoofDay
                'End If
                'If Len(rst2!CLMTotalActual) Then
                '    Record.SubItems(3) = Format(rst2!CLMTotalActual, "#,##0.00")
                'End If
                'If Len(rst2!CLMTotalApproved) Then
                '    Record.SubItems(4) = Format(rst2!CLMTotalApproved, "#,##0.00")
                'End If
                
          '      If Len(rst2!ICDDescription) Then
          '          Record.SubItems(5) = rst2!ICDDescription
          '          Description = Trim(rst2!ICDDescription)
          '      End If

                If Len(txtNoofDay) Then
                    TotalDuration = TotalDuration + CDbl(txtNoofDay)
                End If
                                
                'If Len(rst2!CLMTotalActual) Then
                '    TotalActual = TotalActual + rst2!CLMTotalActual
                'End If
                                
                'If Len(rst2!CLMTotalApproved) Then
                '    TotalApproved = TotalApproved + rst2!CLMTotalApproved
                'End If
            End With
            rst2.MoveNext
        Loop
    rst2.Close
    Set rst2 = Nothing
End Sub
' =================== Count Final Diag. Results ======================


'**********************Start Sorted ListView *******************************
Private Sub lstTopResult_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
    'm_blnDirection(ColumnHeader.Index) = Not m_blnDirection(ColumnHeader.Index)
    'Select Case ColumnHeader.Index
   
    'Case 1
    '    SortListView lstTopResult, ColumnHeader.Index, ldtString, m_blnDirection(1)
    'Case 2
    '    SortListView lstTopResult, ColumnHeader.Index, ldtString, m_blnDirection(2)
    'Case 3
    '    SortListView lstTopResult, ColumnHeader.Index, ldtNumber, m_blnDirection(3)
    'Case 4
    '    SortListView lstTopResult, ColumnHeader.Index, ldtNumber, m_blnDirection(4)
    'Case 5
    '    SortListView lstTopResult, ColumnHeader.Index, ldtNumber, m_blnDirection(5)
    'Case 6
    '    SortListView lstTopResult, ColumnHeader.Index, ldtNumber, m_blnDirection(6)
    'Case 7
    '    SortListView lstTopResult, ColumnHeader.Index, ldtNumber, m_blnDirection(7)
    'Case 8
    '    SortListView lstTopResult, ColumnHeader.Index, ldtNumber, m_blnDirection(8)
    'Case 9
    '    SortListView lstTopResult, ColumnHeader.Index, ldtNumber, m_blnDirection(9)
    'Case 10
    '    SortListView lstTopResult, ColumnHeader.Index, ldtNumber, m_blnDirection(10)
    'End Select
End Sub

Private Sub lstRaceList_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
    m_blnDirection(ColumnHeader.Index) = Not m_blnDirection(ColumnHeader.Index)
    Select Case ColumnHeader.Index
    
    Case 1
        SortListView lstRaceList, ColumnHeader.Index, ldtString, m_blnDirection(1)
    Case 2
        SortListView lstRaceList, ColumnHeader.Index, ldtString, m_blnDirection(2)
    Case 3
        SortListView lstRaceList, ColumnHeader.Index, ldtNumber, m_blnDirection(3)
    Case 4
        SortListView lstRaceList, ColumnHeader.Index, ldtNumber, m_blnDirection(4)
    Case 5
        SortListView lstRaceList, ColumnHeader.Index, ldtNumber, m_blnDirection(5)
    Case 6
        SortListView lstRaceList, ColumnHeader.Index, ldtNumber, m_blnDirection(6)
    Case 7
        SortListView lstRaceList, ColumnHeader.Index, ldtNumber, m_blnDirection(7)
    Case 8
        SortListView lstRaceList, ColumnHeader.Index, ldtNumber, m_blnDirection(8)
    End Select
End Sub

Private Sub lstSexList_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
'    m_blnDirection(ColumnHeader.Index) = Not m_blnDirection(ColumnHeader.Index)
'    Select Case ColumnHeader.Index
    
'    Case 1
'        SortListView lstSexList, ColumnHeader.Index, ldtString, m_blnDirection(1)
'    Case 2
'        SortListView lstSexList, ColumnHeader.Index, ldtString, m_blnDirection(2)
'    Case 3
'        SortListView lstSexList, ColumnHeader.Index, ldtNumber, m_blnDirection(3)
'    Case 4
'        SortListView lstSexList, ColumnHeader.Index, ldtNumber, m_blnDirection(4)
'    Case 5
'        SortListView lstSexList, ColumnHeader.Index, ldtNumber, m_blnDirection(5)
'    Case 6
'        SortListView lstSexList, ColumnHeader.Index, ldtNumber, m_blnDirection(6)
'   End Select
End Sub
'**********************End Sorted ListView *******************************

'********************** List Sex with Ageband Results *******************************
Private Sub SexAgeGroupListing(Optional strAppend As String)
Dim rst2 As New ADODB.Recordset
Dim Record As ListItem
Dim sql As String

    Total = 0
    
    ListView2.listitems.Clear
            
    sql = " SELECT DISTINCT Sex From VIEW_AnalysisReportNew Where 0 = 0 "
    
    If Len(Trim(strAppend)) Then
        sql = sql + strAppend
    End If
            
    rst2.Open sql, medixmasterconn, adOpenKeyset, adLockOptimistic
    
        Do Until rst2.EOF
        
            With rst2
                If Len(rst2!sex) Then
                    Set Record = ListView2.listitems.Add(, , rst2!sex)
                    SexCode = Trim(rst2!sex)
                Else
                    Set Record = ListView2.listitems.Add(, , "Empty")
                    SexCode = Trim(rst2!sex)
                End If
                
                    Call CountAgeGroup(strAppend)
                    Call CountAgeGroupRecords(strAppend)
                               
                    If Len(TotalAgeCode1) Then
                       Record.SubItems(1) = TotalAgeCode1
                    End If
                    
                    If Len(TotalAgeCode2) Then
                        Record.SubItems(2) = TotalAgeCode2
                    End If
                                    
                    If Len(TotalAgeCode3) Then
                       Record.SubItems(3) = TotalAgeCode3
                    End If
                    
                    If Len(TotalAgeCode4) Then
                        Record.SubItems(4) = TotalAgeCode4
                    End If
                    
                    If Len(TotalAgeCode5) Then
                       Record.SubItems(5) = TotalAgeCode5
                    End If
                    
                    If Len(TotalAgeCode6) Then
                        Record.SubItems(6) = TotalAgeCode6
                    End If
                                    
                    Total = (CDbl(TotalAgeCode1) + CDbl(TotalAgeCode2) + CDbl(TotalAgeCode3) + CDbl(TotalAgeCode4) + CDbl(TotalAgeCode5) + CDbl(TotalAgeCode6))
                    
                    If Len(Total) Then
                        Record.SubItems(7) = Total
                    End If
                    
                    If Len(TotalDuration) Then
                        Record.SubItems(8) = TotalDuration
                    End If
                    
                    If Len(TotalActual) Then
                        Record.SubItems(9) = Format(TotalActual, "#,##0.00")
                    End If
                    
                    If Len(TotalApproved) Then
                        Record.SubItems(10) = Format(TotalApproved, "#,##0.00")
                    End If
                'End If
            End With
            rst2.MoveNext
       Loop
    rst2.Close
    Set rst2 = Nothing
End Sub

Private Sub SexListing(Optional strAppend As String)
Dim rst2 As New ADODB.Recordset
Dim Record As ListItem
Dim sql As String
Dim sql2 As String

        
    ListView2.listitems.Clear
            
    sql = " SELECT SEX, COUNT(*) AS TotalCases, SUM(CLMTotalActual) " & _
          " AS totalActual, SUM(CLMTotalApproved) AS TotalApproved " & _
          " From VIEW_AnalysisReportNew Where 0 = 0 "
    
    sql2 = " Group By SEX "
    
    If Len(Trim(strAppend)) Then
        sql = sql + strAppend + sql2
    End If
      
    rst2.Open sql, medixmasterconn, adOpenKeyset, adLockOptimistic
    
        Do Until rst2.EOF
        
            With rst2
                If Len(rst2!sex) Then
                    Set Record = ListView2.listitems.Add(, , rst2!sex)
                    SexCode = Trim(rst2!sex)
                Else
                    Set Record = ListView2.listitems.Add(, , "Empty")
                    SexCode = Trim(rst2!sex)
                End If
                
                    Call CountSexRecords(strAppend)
                               
                    If Len(rst2!TotalCases) Then
                        Record.SubItems(1) = rst2!TotalCases
                    End If
                    
                    If Len(TotalDuration) Then
                        Record.SubItems(2) = TotalDuration
                    End If
                    
                    If Len(rst2!TotalActual) Then
                        Record.SubItems(3) = Format(rst2!TotalActual, "#,##0.00")
                    End If
                    
                    If Len(rst2!TotalApproved) Then
                        Record.SubItems(4) = Format(rst2!TotalApproved, "#,##0.00")
                    End If
            End With
            rst2.MoveNext
       Loop
    rst2.Close
    Set rst2 = Nothing
End Sub

Private Sub CountSexRecords(Optional strAppend As String)
Dim rst2 As New ADODB.Recordset
Dim Record As ListItem
Dim sql As String
    
    TotalDuration = 0
    
    ListView3.listitems.Clear
                
    If Len(SexCode) Or SexCode = "" Then
        sql = " Select * From VIEW_AnalysisReportNew Where Sex  = '" & Trim(SexCode) & "'"
    Else
        sql = " Select * From VIEW_AnalysisReportNew Where Sex IS NULL "
    End If
    
        
    If Len(Trim(strAppend)) Then
        sql = sql + strAppend
    End If
            
    rst2.Open sql, medixmasterconn, adOpenKeyset, adLockOptimistic
    
        Do Until rst2.EOF
        TxtDay = ""
            If Len(rst2!CASDischargeDate) And Len(rst2!CASDateAdmitted) Then
                TxtDay = DateValue(rst2!CASDischargeDate) - DateValue((rst2!CASDateAdmitted))
            End If
            
            If TxtDay = "0" Then
                TxtDay = TxtDay + 1
            End If
            
            With rst2
                If Len(rst2!CASDischargeDate) Then
                    Set Record = ListView3.listitems.Add(, , rst2!CASDischargeDate)
                End If
                If Len(rst2!CASDateAdmitted) Then
                   Record.SubItems(1) = Format(rst2!CASDateAdmitted, "dd/mm/yyyy")
                End If
                If Len(TxtDay) Then
                   Record.SubItems(2) = TxtDay
                End If
                
                If Len(TxtDay) Then
                    TotalDuration = TotalDuration + CDbl(TxtDay)
                End If
            
            End With
            rst2.MoveNext
        Loop
    rst2.Close
    Set rst2 = Nothing
End Sub

Private Sub CountAgeGroupRecords(Optional strAppend As String)
Dim rst2 As New ADODB.Recordset
Dim Record As ListItem
Dim sql As String
        
    
    TotalDuration = 0
    TotalActual = 0
    TotalApproved = 0
    
    ListView3.listitems.Clear
    'ListView4.listitems.Clear
            
    If Len(SexCode) Or SexCode = "" Then
        sql = " Select * From VIEW_AnalysisReportNew Where Sex  = '" & Trim(SexCode) & "'"
    Else
        sql = " Select * From VIEW_AnalysisReportNew Where Sex IS NULL "
    End If
    
    'sql = " Select * From VIEW_AnalysisReportNew Where Sex  = '" & Trim(SexCode) & "'"
    
    If Len(Trim(strAppend)) Then
        sql = sql + strAppend
    End If
            
    rst2.Open sql, medixmasterconn, adOpenKeyset, adLockOptimistic
    
        Do Until rst2.EOF
        TxtDay = ""
            If Len(rst2!CASDischargeDate) And Len(rst2!CASDateAdmitted) Then
                TxtDay = DateValue(rst2!CASDischargeDate) - DateValue((rst2!CASDateAdmitted))
            End If
            
            If TxtDay = "0" Then
                TxtDay = TxtDay + 1
            End If
            
            With rst2
                If Len(rst2!CASDischargeDate) Then
                    Set Record = ListView3.listitems.Add(, , rst2!CASDischargeDate)
                End If
                If Len(rst2!CASDateAdmitted) Then
                   Record.SubItems(1) = Format(rst2!CASDateAdmitted, "dd/mm/yyyy")
                End If
                If Len(TxtDay) Then
                   Record.SubItems(2) = TxtDay
                End If
                If Len(rst2!CLMTotalActual) Then
                    Record.SubItems(3) = Format(rst2!CLMTotalActual, "#,##0.00")
                End If
                If Len(rst2!CLMTotalApproved) Then
                    Record.SubItems(4) = Format(rst2!CLMTotalApproved, "#,##0.00")
                End If

                If Len(TxtDay) Then
                    TotalDuration = TotalDuration + CDbl(TxtDay)
                End If
                                
                If Len(rst2!CLMTotalActual) Then
                    TotalActual = TotalActual + rst2!CLMTotalActual
                End If
                                
                If Len(rst2!CLMTotalApproved) Then
                    TotalApproved = TotalApproved + rst2!CLMTotalApproved
                End If
                                
            End With
            rst2.MoveNext
        Loop
    rst2.Close
    Set rst2 = Nothing
End Sub

Private Sub CountAgeGroup(Optional strAppend As String)
Dim rst2 As New ADODB.Recordset
Dim Record As ListItem
Dim sql As String
Dim Current1 As String
Dim Current2 As String
Dim Current3 As String
Dim Current4 As String
Dim Current5 As String
Dim Current6 As String
        
    Current1 = 0
    Current2 = 0
    Current3 = 0
    Current4 = 0
    Current5 = 0
    Current6 = 0
            
    TotalAgeCode1 = 0
    TotalAgeCode2 = 0
    TotalAgeCode3 = 0
    TotalAgeCode4 = 0
    TotalAgeCode5 = 0
    TotalAgeCode6 = 0
    
    
    ListView3.listitems.Clear
            
    If Len(SexCode) Or SexCode = "" Then
        sql = " Select * From VIEW_AnalysisReportNew Where Sex  = '" & Trim(SexCode) & "'"
    Else
        sql = " Select * From VIEW_AnalysisReportNew Where Sex IS NULL "
    End If
        
    If Len(Trim(strAppend)) Then
        sql = sql + strAppend
    End If
        
    rst2.Open sql, medixmasterconn, adOpenKeyset, adLockOptimistic
    
        Do Until rst2.EOF
                    
            With rst2
                If Len(rst2!sex) Then
                    Set Record = ListView3.listitems.Add(, , rst2!sex)
                Else
                    Set Record = ListView3.listitems.Add(, , "Empty")
                End If
                
                If rst2!AgeBand = "01" Then
                   Record.SubItems(1) = rst2!AgeBand
                   Current1 = Current1 + 1
                   TotalAgeCode1 = Current1
                ElseIf rst2!AgeBand = "02" Then
                   Record.SubItems(2) = rst2!AgeBand
                   Current2 = Current2 + 1
                   TotalAgeCode2 = Current2
                ElseIf rst2!AgeBand = "03" Then
                   Record.SubItems(3) = rst2!AgeBand
                   Current3 = Current3 + 1
                   TotalAgeCode3 = Current3
                ElseIf rst2!AgeBand = "04" Then
                   Record.SubItems(4) = rst2!AgeBand
                   Current4 = Current4 + 1
                   TotalAgeCode4 = Current4
                ElseIf rst2!AgeBand = "05" Then
                   Record.SubItems(5) = rst2!AgeBand
                   Current5 = Current5 + 1
                   TotalAgeCode5 = Current5
                ElseIf rst2!AgeBand = "06" Then
                   Record.SubItems(6) = rst2!AgeBand
                   Current6 = Current6 + 1
                   TotalAgeCode6 = Current6
                End If
            End With
            rst2.MoveNext
        Loop
    rst2.Close
    Set rst2 = Nothing
End Sub

' =================== List Sex with Ageband Results ======================
Private Sub DiagSexAgeListing(Optional strAppend As String)
Dim rst2 As New ADODB.Recordset
Dim Record As ListItem
Dim sql As String
    
    TotalM = 0
    TotalF = 0
    TotalCases = 0
    
    sql = ""
    ListView7.listitems.Clear
            
    If Mid(cboSelection, 1, 1) = "0" Then
        sql = "TOP 10 "
    End If
    If Mid(cboSelection, 1, 1) = "1" Then
        sql = "TOP 20 "
    End If
    
    sql = "SELECT " & sql & "FinalDiag, COUNT(FinalDiag) AS Total" & _
          " From VIEW_AnalysisReportNew4 Where 0 = 0 " & strAppend & _
          " GROUP BY FinalDiag ORDER BY Total DESC "
    
    rst2.Open sql, medixmasterconn, adOpenKeyset, adLockOptimistic
    
        Do Until rst2.EOF
    
            With rst2
                Set Record = ListView7.listitems.Add(, , rst2!FinalDiag)
                DiagCode = Trim(rst2!FinalDiag)
                
                Call CountSexAge(strAppend)
                
                If Len(Description) Then
                    Record.SubItems(1) = Trim(Description)
                End If
                
                If Len(TotalM1) Then
                   Record.SubItems(2) = TotalM1
                End If
                If Len(TotalM2) Then
                    Record.SubItems(3) = TotalM2
                End If
                                
                If Len(TotalM3) Then
                   Record.SubItems(4) = TotalM3
                End If
                If Len(TotalM4) Then
                    Record.SubItems(5) = TotalM4
                End If
                
                If Len(TotalM5) Then
                   Record.SubItems(6) = TotalM5
                End If
                If Len(TotalM6) Then
                    Record.SubItems(7) = TotalM6
                End If
                If Len(TotalM7) Then
                    Record.SubItems(8) = TotalM7
                End If
                
                TotalM = (CDbl(TotalM1) + CDbl(TotalM2) + CDbl(TotalM3) + CDbl(TotalM4) + CDbl(TotalM5) + CDbl(TotalM6) + CDbl(TotalM7))
                
                If Len(TotalM) Then
                    Record.SubItems(9) = TotalM
                End If
                
                If Len(TotalF1) Then
                   Record.SubItems(10) = TotalF1
                End If
                If Len(TotalF2) Then
                    Record.SubItems(11) = TotalF2
                End If
                                
                If Len(TotalF3) Then
                   Record.SubItems(12) = TotalF3
                End If
                If Len(TotalF4) Then
                    Record.SubItems(13) = TotalF4
                End If
                
                If Len(TotalF5) Then
                   Record.SubItems(14) = TotalF5
                End If
                If Len(TotalF6) Then
                    Record.SubItems(15) = TotalF6
                End If
                If Len(TotalF7) Then
                    Record.SubItems(16) = TotalF7
                End If
                
                TotalF = (CDbl(TotalF1) + CDbl(TotalF2) + CDbl(TotalF3) + CDbl(TotalF4) + CDbl(TotalF5) + CDbl(TotalF6) + CDbl(TotalF7))
                
                If Len(TotalF) Then
                    Record.SubItems(17) = TotalF
                End If
                
                If Len(TotalEmpty) Then
                    Record.SubItems(18) = TotalEmpty
                End If
                
                TotalCases = (CDbl(TotalM) + CDbl(TotalF) + CDbl(TotalEmpty))
                
                If Len(TotalCases) Then
                    Record.SubItems(19) = TotalCases
                End If
                
            End With
            rst2.MoveNext
        Loop
    rst2.Close
    Set rst2 = Nothing
End Sub
' =================== List Sex with Ageband Results ======================

' =================== Search Diag. Sex Age Top 10/20 Results ======================
Private Sub DiagSexAgeTop1020()
Dim rst2 As New ADODB.Recordset
Dim Record As ListItem
Dim sql As String
       
    TotalM = 0
    TotalF = 0
    TotalCases = 0
    
    ListView7.listitems.Clear
    
    If Mid(cboSelection, 1, 1) = "0" Then
    
        sql = " SELECT TOP 10 FinalDiag, COUNT(FinalDiag) AS Total " & _
              " From VIEW_AnalysisReportNew3 " & _
              " Where (CASStatus = 'O' OR CASStatus = 'C') AND CASClaimability = 'A' " & _
              " GROUP BY FinalDiag ORDER BY Total DESC "
         
    ElseIf Mid(cboSelection, 1, 1) = "1" Then
        sql = " SELECT TOP 20 FinalDiag, COUNT(FinalDiag) AS Total " & _
              " From VIEW_AnalysisReportNew3 " & _
              " Where (CASStatus = 'O' OR CASStatus = 'C') AND CASClaimability = 'A' " & _
              " GROUP BY FinalDiag ORDER BY Total DESC "
    End If
    
    
    rst2.Open sql, medixmasterconn, adOpenKeyset, adLockOptimistic
    
        Do Until rst2.EOF
    
            With rst2
                Set Record = ListView7.listitems.Add(, , rst2!FinalDiag)
                DiagCode = Trim(rst2!FinalDiag)
                
                Call CountSexAge
                
                If Len(Description) Then
                    Record.SubItems(1) = Trim(Description)
                End If
                
                If Len(TotalM1) Then
                   Record.SubItems(2) = TotalM1
                End If
                If Len(TotalM2) Then
                    Record.SubItems(3) = TotalM2
                End If
                                
                If Len(TotalM3) Then
                   Record.SubItems(4) = TotalM3
                End If
                If Len(TotalM4) Then
                    Record.SubItems(5) = TotalM4
                End If
                
                If Len(TotalM5) Then
                   Record.SubItems(6) = TotalM5
                End If
                If Len(TotalM6) Then
                    Record.SubItems(7) = TotalM6
                End If
                If Len(TotalM7) Then
                    Record.SubItems(8) = TotalM7
                End If
                
                TotalM = (CDbl(TotalM1) + CDbl(TotalM2) + CDbl(TotalM3) + CDbl(TotalM4) + CDbl(TotalM5) + CDbl(TotalM6) + CDbl(TotalM7))
                
                If Len(TotalM) Then
                    Record.SubItems(9) = TotalM
                End If
                
                If Len(TotalF1) Then
                   Record.SubItems(10) = TotalF1
                End If
                If Len(TotalF2) Then
                    Record.SubItems(11) = TotalF2
                End If
                                
                If Len(TotalF3) Then
                   Record.SubItems(12) = TotalF3
                End If
                If Len(TotalF4) Then
                    Record.SubItems(13) = TotalF4
                End If
                
                If Len(TotalF5) Then
                   Record.SubItems(14) = TotalF5
                End If
                If Len(TotalF6) Then
                    Record.SubItems(15) = TotalF6
                End If
                If Len(TotalF7) Then
                    Record.SubItems(16) = TotalF7
                End If
                
                TotalF = (CDbl(TotalF1) + CDbl(TotalF2) + CDbl(TotalF3) + CDbl(TotalF4) + CDbl(TotalF5) + CDbl(TotalF6) + CDbl(TotalF7))
                
                If Len(TotalF) Then
                    Record.SubItems(17) = TotalF
                End If
                
                If Len(TotalEmpty) Then
                    Record.SubItems(18) = TotalEmpty
                End If
                
                TotalCases = (CDbl(TotalM) + CDbl(TotalF) + CDbl(TotalEmpty))
                
                If Len(TotalCases) Then
                    Record.SubItems(19) = TotalCases
                End If
                
            End With
            rst2.MoveNext
        Loop
    rst2.Close
    Set rst2 = Nothing
End Sub
' =================== Search Diag. Sex Age Top 10/20 Results ======================


' =================== Search Daily Admission Report Results ======================
Private Function SearchDailyAdmission()
'Dim strAppend As String
'
'    strAppend = ""
    
'    If CboPayor = "" Then
'        MsgBox "Please Enter Payor Code", vbOKOnly + vbCritical, DialogBoxTitle
'        CboPayor.SetFocus
'    ElseIf TxtAdmissionDate = "__/__/____" Then
'        MsgBox "Please Enter Admission Date", vbOKOnly + vbCritical, DialogBoxTitle
'        TxtAdmissionDate.SetFocus
'    Else

    
        If Len(Trim(CboPayor)) Then
            strAppend = strAppend + " And PAYCode = '" & Trim(Mid(CboPayor, 1, IIf(InStr(1, CboPayor, "-") = 0, 4, InStr(1, CboPayor, "-") - 1))) & "'"
        End If
    
        If (TxtAdmissionDate) <> "__/__/____" And IsDate(TxtAdmissionDate) = True Then
            strAppend = strAppend + " AND CasDateAdmitted = '" & Format(Trim(TxtAdmissionDate), "mm/dd/yyyy") & "'"
        End If
    
        If Option15.value = True Then
            sql = ""
            sql = "AND CASStatus = 'O' "
            strAppend = strAppend + sql
        End If
        
        If Option16.value = True Then
            sql = ""
            sql = "AND CASStatus = 'C'"
            strAppend = strAppend + sql
        End If
        
        If Option17.value = True Then
            sql = ""
            sql = "AND CASStatus = 'D'"
            strAppend = strAppend + sql
        End If
        
        If Option18.value = True Then
            sql = ""
            sql = "AND (CASStatus = 'O' OR CASStatus = 'C')"
            strAppend = strAppend + sql
        End If
        
        If Option19.value = True Then
            sql = ""
            sql = "AND CASClaimability = 'A'"
            strAppend = strAppend + sql
        End If
        
        If Option20.value = True Then
            sql = ""
            sql = "AND CASClaimability = 'R'"
            strAppend = strAppend + sql
        End If
    
        Call DailyAdmissionListing(strAppend)
    End If
End Function
' =================== Search Daily Admission Report Results ======================


'================ Daily Admission Listing ===========================
Private Sub DailyAdmissionListing(Optional strAppend As String)
Dim rst2 As New ADODB.Recordset
Dim Record As ListItem
Dim sql As String
Dim intloop As Integer
       
    ListView2.listitems.Clear
                
    sql = " SELECT * From VIEW_AnalysisReportNew Where 0 = 0 "
    
    If Len(Trim(strAppend)) Then
        sql = sql + strAppend
    End If
    
    rst2.Open sql, medixmasterconn, adOpenKeyset, adLockOptimistic
    
    If rst2.recordCount = 0 Then
        MsgBox "No record found!", vbOKOnly + vbCritical, DialogBoxTitle
    Else
    
        Do Until rst2.EOF
                    
            With rst2
            
                intloop = intloop + 1
                
                Set Record = ListView2.listitems.Add(, , intloop)
                                                
                If Len(rst2!CASNumber) Then
                    Record.SubItems(1) = rst2!CASNumber
                'End If
                
                If Len(rst2!MBMNumber) Then
                    Record.SubItems(2) = rst2!MBMNumber
                End If
                
                If Len(rst2!MBMPolicyNo) Then
                    Record.SubItems(3) = rst2!MBMPolicyNo
                End If
                
                If Len(rst2!PatientName) Then
                    Record.SubItems(4) = rst2!PatientName
                End If
                
                If Len(rst2!CASDateAdmitted) Then
                    Record.SubItems(5) = rst2!CASDateAdmitted
                End If
                
                If Len(rst2!HOSName) Then
                    Record.SubItems(6) = rst2!HOSName
                End If
                
                If Len(rst2!FirstDiag) Then
                    Record.SubItems(7) = rst2!FirstDiag
                End If
                
                End If
            End With
            rst2.MoveNext
        Loop
    End If
    rst2.Close
    Set rst2 = Nothing
End Sub
'================ Daily Admission Listing ===========================

' =================== Search Daily Discharge Report Results ======================
Private Function SearchDailyDischarge()
Dim strAppend As String
    
    strAppend = ""
    
    If CboPayor = "" Then
        MsgBox "Please Enter Payor Code", vbOKOnly + vbCritical, DialogBoxTitle
        CboPayor.SetFocus
    ElseIf TxtAdmissionDate = "__/__/____" Then
        MsgBox "Please Enter Discharge Date", vbOKOnly + vbCritical, DialogBoxTitle
        TxtAdmissionDate.SetFocus
    Else

    
    If Len(Trim(CboPayor)) Then
        strAppend = strAppend + " And PAYCode = '" & Trim(Mid(CboPayor, 1, IIf(InStr(1, CboPayor, "-") = 0, 4, InStr(1, CboPayor, "-") - 1))) & "'"
    End If
    
    If (TxtAdmissionDate) <> "__/__/____" And IsDate(TxtAdmissionDate) = True Then
        strAppend = strAppend + " AND CASDischargeDate = '" & Format(Trim(TxtAdmissionDate), "mm/dd/yyyy") & "'"
    End If
    
    If Option15.value = True Then
            sql = ""
            sql = "AND CASStatus = 'O' "
            strAppend = strAppend + sql
        End If
        
        If Option16.value = True Then
            sql = ""
            sql = "AND CASStatus = 'C'"
            strAppend = strAppend + sql
        End If
        
        If Option17.value = True Then
            sql = ""
            sql = "AND CASStatus = 'D'"
            strAppend = strAppend + sql
        End If
        
        If Option18.value = True Then
            sql = ""
            sql = "AND (CASStatus = 'O' OR CASStatus = 'C')"
            strAppend = strAppend + sql
        End If
        
        If Option19.value = True Then
            sql = ""
            sql = "AND CASClaimability = 'A'"
            strAppend = strAppend + sql
        End If
        
        If Option20.value = True Then
            sql = ""
            sql = "AND CASClaimability = 'R'"
            strAppend = strAppend + sql
        End If
    
        Call DailyDischargeListing(strAppend)
    End If
End Function
' =================== Search Daily Discharge Report Results ======================

'================ Daily Discharge Listing ===========================
Private Sub DailyDischargeListing(Optional strAppend As String)
Dim rst2 As New ADODB.Recordset
Dim Record As ListItem
Dim sql As String
Dim intloop As Integer
       
    ListView9.listitems.Clear
    
            
    sql = " SELECT * From VIEW_Daily_Report Where 0 = 0 "
    
    If Len(Trim(strAppend)) Then
        sql = sql + strAppend
    End If
    
    rst2.Open sql, medixmasterconn, adOpenKeyset, adLockOptimistic
    
    If rst2.recordCount = 0 Then
        MsgBox "No record found!", vbOKOnly + vbCritical, DialogBoxTitle
    Else
    
        Do Until rst2.EOF
                    
            With rst2
            
                intloop = intloop + 1
                
                Set Record = ListView9.listitems.Add(, , intloop)
                
                                
                If Len(rst2!CASNumber) Then
                    Record.SubItems(1) = rst2!CASNumber
                End If
                
                If Len(rst2!MBMNumber) Then
                    Record.SubItems(2) = rst2!MBMNumber
                End If
                
                If Len(rst2!MBMPolicyNo) Then
                    Record.SubItems(3) = rst2!MBMPolicyNo
                End If
                
                If Len(rst2!PatientName) Then
                    Record.SubItems(4) = rst2!PatientName
                End If
                
                If Len(rst2!CASDateAdmitted) Then
                    Record.SubItems(5) = rst2!CASDateAdmitted
                End If
                
                If Len(rst2!CASDischargeDate) Then
                    Record.SubItems(6) = rst2!CASDischargeDate
                End If
                
                If Len(rst2!HOSName) Then
                    Record.SubItems(7) = rst2!HOSName
                End If
                
                If Len(rst2!FinalDiag) Then
                    Record.SubItems(8) = rst2!FinalDiag
                End If
                
            End With
            rst2.MoveNext
        Loop
    End If
    rst2.Close
    Set rst2 = Nothing
End Sub
'================ Daily Discharge Listing ===========================



'================== Start Counting AgeBand ============================
Private Sub CountSexAge(Optional strAppend As String)
Dim rst2 As New ADODB.Recordset
Dim Record As ListItem
Dim sql As String
Dim Current1 As String
Dim Current2 As String
Dim Current3 As String
Dim Current4 As String
Dim Current5 As String
Dim Current6 As String
Dim Current7 As String
Dim Current8 As String
Dim Current9 As String
Dim Current10 As String
Dim Current11 As String
Dim Current12 As String
Dim Current13 As String
Dim Current14 As String
Dim Current15 As String

    Current1 = 0
    Current2 = 0
    Current3 = 0
    Current4 = 0
    Current5 = 0
    Current6 = 0
    Current7 = 0
    Current8 = 0
    Current9 = 0
    Current10 = 0
    Current11 = 0
    Current12 = 0
    Current13 = 0
    Current14 = 0
    Current15 = 0
    
    TotalM1 = 0
    TotalM2 = 0
    TotalM3 = 0
    TotalM4 = 0
    TotalM5 = 0
    TotalM6 = 0
    TotalM7 = 0
    
    TotalF1 = 0
    TotalF2 = 0
    TotalF3 = 0
    TotalF4 = 0
    TotalF5 = 0
    TotalF6 = 0
    TotalF7 = 0
    TotalEmpty = 0
        
    ListView1.listitems.Clear
            
    sql = " Select * From VIEW_AnalysisReportNew3 where FinalDiag = '" & Trim(DiagCode) & "'"
        
    sql2 = " AND (CASStatus = 'O' OR CASStatus = 'C') AND CASClaimability = 'A' "
    
    If Mid(cboSelection, 1, 1) = "2" Then
    
        If Len(Trim(strAppend)) Then
            sql = sql + strAppend
        End If
    
    ElseIf Mid(cboSelection, 1, 1) = "0" Or Mid(cboSelection, 1, 1) = "1" Then
        sql = sql + sql2
    End If
    
    rst2.Open sql, medixmasterconn, adOpenKeyset, adLockOptimistic
    
        Do Until rst2.EOF
                    
            With rst2
                Set Record = ListView1.listitems.Add(, , rst2!ICDDescription)
                    Description = rst2!ICDDescription
                
                If rst2!sex = "M" Then
                    If rst2!AgeBand = "01" Then
                       Record.SubItems(1) = rst2!AgeBand
                       Current1 = Current1 + 1
                       TotalM1 = Current1
                    ElseIf rst2!AgeBand = "02" Then
                       Record.SubItems(1) = rst2!AgeBand
                       Current2 = Current2 + 1
                       TotalM2 = Current2
                    ElseIf rst2!AgeBand = "03" Then
                       Record.SubItems(1) = rst2!AgeBand
                       Current3 = Current3 + 1
                       TotalM3 = Current3
                    ElseIf rst2!AgeBand = "04" Then
                       Record.SubItems(1) = rst2!AgeBand
                       Current4 = Current4 + 1
                       TotalM4 = Current4
                    ElseIf rst2!AgeBand = "05" Then
                       Record.SubItems(1) = rst2!AgeBand
                       Current5 = Current5 + 1
                       TotalM5 = Current5
                    ElseIf rst2!AgeBand = "06" Then
                       Record.SubItems(1) = rst2!AgeBand
                       Current6 = Current6 + 1
                       TotalM6 = Current6
                    Else
                       Current7 = Current7 + 1
                       TotalM7 = Current7
                    End If
                ElseIf rst2!sex = "F" Then
                    If rst2!AgeBand = "01" Then
                       Record.SubItems(1) = rst2!AgeBand
                       Current8 = Current8 + 1
                       TotalF1 = Current8
                    ElseIf rst2!AgeBand = "02" Then
                       Record.SubItems(1) = rst2!AgeBand
                       Current9 = Current9 + 1
                       TotalF2 = Current9
                    ElseIf rst2!AgeBand = "03" Then
                       Record.SubItems(1) = rst2!AgeBand
                       Current10 = Current10 + 1
                       TotalF3 = Current10
                    ElseIf rst2!AgeBand = "04" Then
                       Record.SubItems(1) = rst2!AgeBand
                       Current11 = Current11 + 1
                       TotalF4 = Current11
                    ElseIf rst2!AgeBand = "05" Then
                       Record.SubItems(1) = rst2!AgeBand
                       Current12 = Current12 + 1
                       TotalF5 = Current12
                    ElseIf rst2!AgeBand = "06" Then
                       Record.SubItems(1) = rst2!AgeBand
                       Current13 = Current13 + 1
                       TotalF6 = Current13
                    Else
                       Current14 = Current14 + 1
                       TotalF7 = Current14
                    End If
               Else
                    Current15 = Current15 + 1
                    TotalEmpty = Current15
               End If
            End With
            rst2.MoveNext
        Loop
    rst2.Close
    Set rst2 = Nothing
End Sub


'=========== Option Button, Combo Box Change/KeyPress ===================

Private Sub cboPayorCode_Click()
Dim strsql As String
Dim PAYGRP As New ADODB.Recordset

CboGrpCom.Clear
    strsql = "select MBMGroupCompany from View_PAYGrpCom where PAYCode= '" & Mid(cboPayorCode, 1, 2) & "'"
    PAYGRP.Open strsql, medixmasterconn, adOpenKeyset, adLockOptimistic
    
    If PAYGRP.recordCount Then
        Do Until PAYGRP.EOF
           CboGrpCom.AddItem Trim(PAYGRP!MBMGroupCompany)
           PAYGRP.MoveNext
        Loop
    End If
    PAYGRP.Close
    Set PAYGRP = Nothing

End Sub

Private Sub cboPayorCode_Change()

    If cboPayorCode.Text = "" Then
        If InStr(cboPayorCode.Text, "null") Then
           cboPayorCode.Enabled = False
        End If
    End If
End Sub

Private Sub cboPayorCode_KeyPress(KeyAscii As Integer)
Dim NewText As String
Dim ValidCount As Integer
Dim ValidValue As String
Dim i As Integer

    If KeyAscii >= 32 And KeyAscii <> 127 Then
    NewText = LCase(Left(cboPayorCode.Text, cboPayorCode.SelStart) + Chr(KeyAscii) + Mid(cboPayorCode.Text, cboPayorCode.SelStart + cboPayorCode.SelLength + 1))
 
    ValidCount = 0
    ValidValue = ""
 
    For i = 0 To cboPayorCode.ListCount - 1
 
        If NewText = LCase(Left(cboPayorCode.List(i), Len(NewText))) Then
                ValidCount = ValidCount + 1
                ValidValue = cboPayorCode.List(i)
        End If
    Next
        
        If ValidCount <= 1 Then KeyAscii = 0
            If ValidCount = 1 Then
                cboPayorCode.Text = ValidValue
                cboPayorCode.SelStart = 0
                cboPayorCode.SelLength = Len(ValidValue)
            End If
        End If
End Sub

Private Sub Opt1_Click()
    If Opt1.value = True Then
        ListView7.Visible = False
        lstSexList.Visible = False
        lstRaceList.Visible = False
        lstTopResult.Visible = True
    End If
End Sub

Private Sub Opt2_Click()
    If Opt2.value = True Then
        ListView7.Visible = False
        lstSexList.Visible = False
        lstRaceList.Visible = True
        lstTopResult.Visible = False
    End If
End Sub

Private Sub Opt23_Click()
    If Opt23.value = True Then
        lstDiagPCodeList.Visible = True
        lstFinalDiagList.Visible = False
        ListView7.Visible = False
        lstSexList.Visible = False
        lstRaceList.Visible = False
        lstTopResult.Visible = False
    End If
End Sub

Private Sub Opt3_Click()
    If Opt3.value = True Then
        ListView7.Visible = False
        lstSexList.Visible = True
        lstRaceList.Visible = False
        lstTopResult.Visible = False
    End If
End Sub

Private Sub Opt4_Click()
    If Opt4.value = True Then
        ListView7.Visible = True
        lstSexList.Visible = False
        lstRaceList.Visible = False
        lstTopResult.Visible = False
    End If
End Sub

Private Sub OptAdmit_Click()
    If OptAdmit.value = True Then
        lstDailyList.Visible = True
        ListView9.Visible = False
    Else
        ListView9.Visible = True
        lstDailyList.Visible = False
    End If
End Sub

Private Sub OptDischarge_Click()
    If OptDischarge.value = True Then
        ListView9.Visible = True
        lstDailyList.Visible = False
    Else
        lstDailyList.Visible = True
        ListView9.Visible = False
    End If
End Sub
'=========== Combo Box Change/KeyPress ===================

Private Sub Opt22_Click()
    If Opt22.value = True Then
        lstFinalDiagList.Visible = True
        ListView7.Visible = False
        lstSexList.Visible = False
        lstRaceList.Visible = False
        lstTopResult.Visible = False
    End If
End Sub

Private Sub Option1_Click()
    ChkDOA.Enabled = True
    Option29.Visible = False
    Option33.Visible = False
    Option34.Visible = False
    Option35.Visible = False
    Call ListViewHeader
End Sub

Private Sub Option10_Click()
    
    ChkDOA.Enabled = True
    Option35.value = True
    Option29.Visible = True
    Option33.Visible = True
    Option34.Visible = True
    Option35.Visible = True
    
    Option35.value = True
    Option35.Caption = "Detail Listing"
    Option34.Caption = "General Summary"
    Option33.Caption = "Top 20 by Cases"
    Option29.Caption = "Top 20 by Amount"
    Option35.Visible = True
    Option34.Visible = True
    Option33.Visible = True
    Option29.Visible = True
           
    If Option34.value = True Or Option33.value = True Or Option29.value = True Then
        Call ListViewHeader10
    Else
        Call ListViewHeader11
    End If
    
End Sub

Private Sub Option11_Click()
    ChkDOA.Enabled = True
    Option29.Visible = False
    Option33.Visible = False
    Option34.Visible = False
    Option35.Visible = False
End Sub

Private Sub Option12_Click()
    ChkDOA.Enabled = True
    Option29.Visible = False
    Option33.Visible = False
    Option34.Visible = False
    Option35.Visible = False
End Sub

Private Sub Option13_Click()
    ChkDOA.Enabled = True
    Option29.Visible = False
    Option33.Visible = False
    Option34.Visible = False
    Option35.Visible = False
End Sub

Private Sub Option14_Click()
    ChkDOA.Enabled = True
    Option29.Visible = False
    Option33.Visible = False
    Option34.Visible = False
    Option35.Visible = False
End Sub

Private Sub Option15_Click()
    ChkDOA.Enabled = True
    Option35.value = True
    Call ListViewHeader31
    Option35.Caption = "Top 20 by Cases"
    Option34.Caption = "Top 20 by Amount"
    
    Option35.Visible = True
    Option34.Visible = True
    Option33.Visible = False
    Option29.Visible = False
           
  '  If Option35.value = True Or Option34.value = True Then
  '      Call ListViewHeader10
  '  Else
  '      Call ListViewHeader11
  '  End If
End Sub

Private Sub Option16_Click()
    ChkDOA.Enabled = True
    Option29.Visible = False
    Option33.Visible = False
    Option34.Visible = False
    Option35.Visible = False
End Sub

Private Sub Option17_Click()
    ChkDOA.Enabled = True
    Option29.Visible = False
    Option33.Visible = False
    Option34.Visible = False
    Option35.Visible = False
End Sub

Private Sub Option18_Click()
    ChkDOA.Enabled = True
    Option29.Visible = False
    Option33.Visible = False
    Option34.Visible = False
    Option35.Visible = False
End Sub

Private Sub Option19_Click()
    ChkDOA.Enabled = True
    Option29.Visible = False
    Option33.Visible = False
    Option34.Visible = False
    Option35.Visible = False
End Sub

Private Sub Option2_Click()
    ChkDOA.Enabled = True
    Option29.Visible = False
    Option33.Visible = False
    Option34.Visible = False
    Option35.Visible = False
    Call ListViewHeader27
End Sub

Private Sub Option20_Click()
    ChkDOA.Enabled = False
    Option35.value = True
    Call ListViewHeader20
    Option35.Caption = "Daily Report"
    Option34.Caption = "Monthly Report"
    Option33.Caption = "YTD Report"
    Option35.Visible = True
    Option34.Visible = True
    Option33.Visible = True
    Option29.Visible = False
End Sub

Private Sub Option21_Click()
    ChkDOA.Enabled = True
    Option29.Visible = False
    Option33.Visible = False
    Option34.Visible = False
    Option35.Visible = False
End Sub

Private Sub Option22_Click()
    ChkDOA.Enabled = True
    Option29.Visible = False
    Option33.Visible = False
    Option34.Visible = False
    Option35.Visible = False
End Sub

Private Sub Option23_Click()
    ChkDOA.Enabled = True
    Option29.Visible = False
    Option33.Visible = False
    Option34.Visible = False
    Option35.Visible = False
End Sub

Private Sub Option24_Click()
    ChkDOA.Enabled = True
    Option29.Visible = False
    Option33.Visible = False
    Option34.Visible = False
    Option35.Visible = False
End Sub

Private Sub Option25_Click()
    ChkDOA.Enabled = True
    Option29.Visible = False
    Option33.Visible = False
    Option34.Visible = False
    Option35.Visible = False
End Sub

Private Sub Option26_Click()
    ChkDOA.Enabled = True
    Option29.Visible = False
    Option33.Visible = False
    Option34.Visible = False
    Option35.Visible = False
End Sub

Private Sub Option27_Click()
    ChkDOA.Enabled = True
    Option29.Visible = False
    Option33.Visible = False
    Option34.Visible = False
    Option35.Visible = False
End Sub

Private Sub Option28_Click()
    ChkDOA.Enabled = True
    Option29.Visible = False
    Option33.Visible = False
    Option34.Visible = False
    Option35.Visible = False
End Sub

Private Sub Option29_Click()
    cboHOSName.Enabled = False
    cboHOSGrp.Enabled = False
    Call ListViewHeader10
End Sub

Private Sub Option3_Click()
    ChkDOA.Enabled = True
    Option29.Visible = False
    Option33.Visible = False
    Option34.Visible = False
    Option35.Visible = False
    Call ListViewHeader28
End Sub

Private Sub Option33_Click()
    If Option20.value = True Then
        Call ListViewHeader22
        Label47.Caption = "DOA Year"
        TxtDOA1.Visible = False
        TxtDOA2.Visible = False
        Label49.Visible = False
        ChkDOA.Visible = False
    ElseIf Option10.value = True Then
        cboHOSName.Enabled = False
        cboHOSGrp.Enabled = False
        Call ListViewHeader10
    End If
End Sub

Private Sub Option34_Click()
    
    If Option20.value = True Then
        Call ListViewHeader21
        TxtDOA1.Visible = True
        TxtDOA2.Visible = True
        Label47.Caption = "DOA"
        Label49.Visible = True
        ChkDOA.Visible = True
    ElseIf Option10.value = True Then
        cboHOSName.Enabled = True
        cboHOSGrp.Enabled = True
        Call ListViewHeader10
    End If
End Sub

Private Sub Option35_Click()
    
    If Option20.value = True Then
        ChkDOA.Enabled = False
        TxtDOA1.Visible = True
        TxtDOA2.Visible = True
        Label47.Caption = "DOA"
        Label49.Visible = True
        ChkDOA.Visible = True
        Call ListViewHeader20
    ElseIf Option10.value = True Then
        cboHOSName.Enabled = True
        cboHOSGrp.Enabled = True
        Call ListViewHeader11
    End If
End Sub

Private Sub Option36_Click()
    ChkDOA.Enabled = True
    Option29.Visible = False
    Option33.Visible = False
    Option34.Visible = False
    Option35.Visible = False
    Call ListViewHeader23
End Sub

Private Sub Option4_Click()
    ChkDOA.Enabled = True
    Option29.Visible = False
    Option33.Visible = False
    Option34.Visible = False
    Option35.Visible = False
End Sub

Private Sub Option5_Click()
    ChkDOA.Enabled = True
    Option35.value = True
    Option35.Visible = True
    Option35.Caption = "Summary"
    Option34.Visible = True
    Option34.Caption = "By Gender"
    Option33.Visible = True
    Option33.Caption = "By Insured Type"
    Option29.Visible = True
    Option29.Caption = "By Race"
    Option37.Visible = True
    Option37.Caption = "By Plan No"
    Option30.Visible = True
    Option30.Caption = "By Nature of Claims"
    Call ListViewHeader24
End Sub

Private Sub Option6_Click()
    ChkDOA.Enabled = True
    Option29.Visible = False
    Option33.Visible = False
    Option34.Visible = False
    Option35.Visible = False
    Call ListViewHeader25
End Sub

Private Sub Option7_Click()
    ChkDOA.Enabled = True
    Option29.Visible = False
    Option33.Visible = False
    Option34.Visible = False
    Option35.Visible = False
    Call ListViewHeader26
End Sub

Private Sub Option8_Click()
    ChkDOA.Enabled = True
    Option29.Visible = False
    Option33.Visible = False
    Option34.Visible = False
    Option35.Visible = False
    Call ListViewHeader29
End Sub

Private Sub Option9_Click()
    ChkDOA.Enabled = True
    Option29.Visible = False
    Option33.Visible = False
    Option34.Visible = False
    Option35.Visible = False
    Call ListViewHeader30
End Sub

'=========== Check Date Format ===========================
'Private Sub TxtAdmissionDate_LostFocus()
    'If IsDate(txtAdmissionDate) Then
    'Else
        'If txtAdmissionDate <> "__/__/____" Then
            'MsgBox "Invalid Date Format! ", vbCritical
            'txtAdmissionDate.SetFocus
        'End If
    'End If
'End Sub

Private Sub TxtDOA1_LostFocus()
    If IsDate(TxtDOA1) Then
    Else
        If TxtDOA1 <> "__/__/____" Then
            MsgBox "Invalid Date Format! ", vbCritical
            TxtDOA1.SetFocus
        End If
    End If
End Sub

Private Sub TxtDOA2_LostFocus()
    If IsDate(TxtDOA2) Then
    Else
        If TxtDOA2 <> "__/__/____" Then
            MsgBox "Invalid Date Format! ", vbCritical
            TxtDOA2.SetFocus
        End If
    End If
End Sub

Private Sub TxtDOD1_LostFocus()
    If IsDate(txtDOD1) Then
    Else
        If txtDOD1 <> "__/__/____" Then
            MsgBox "Invalid Date Format! ", vbCritical
            txtDOD1.SetFocus
        End If
    End If
End Sub

Private Sub TxtDOD2_LostFocus()
    If IsDate(txtDOD2) Then
    Else
        If txtDOD2 <> "__/__/____" Then
            MsgBox "Invalid Date Format! ", vbCritical
            txtDOD2.SetFocus
        End If
    End If
End Sub

Private Sub TxtDOR1_LostFocus()
    If IsDate(TxtDOR1) Then
    Else
        If TxtDOR1 <> "__/__/____" Then
            MsgBox "Invalid Date Format! ", vbCritical
            TxtDOR1.SetFocus
        End If
    End If
End Sub

Private Sub TxtDOR2_LostFocus()
    If IsDate(TxtDOR2) Then
    Else
        If TxtDOR2 <> "__/__/____" Then
            MsgBox "Invalid Date Format! ", vbCritical
            TxtDOR2.SetFocus
        End If
    End If
End Sub

Private Sub TxtPayEffDate1_LostFocus()
    If IsDate(TxtPayEffDate1) Then
    Else
        If TxtPayEffDate1 <> "__/__/____" Then
            MsgBox "Invalid Date Format! ", vbCritical
            TxtPayEffDate1.SetFocus
        End If
    End If
End Sub

Private Sub TxtPayEffDate2_LostFocus()
    If IsDate(TxtPayEffDate2) Then
    Else
        If TxtPayEffDate2 <> "__/__/____" Then
            MsgBox "Invalid Date Format! ", vbCritical
            TxtPayEffDate2.SetFocus
        End If
    End If
End Sub

Private Sub TxtPayExpDate1_LostFocus()
    If IsDate(TxtPayExpDate1) Then
    Else
        If TxtPayExpDate1 <> "__/__/____" Then
            MsgBox "Invalid Date Format! ", vbCritical
            TxtPayExpDate1.SetFocus
        End If
    End If
End Sub

Private Sub TxtPayExpDate2_LostFocus()
    If IsDate(TxtPayExpDate2) Then
    Else
        If TxtPayExpDate2 <> "__/__/____" Then
            MsgBox "Invalid Date Format! ", vbCritical
            TxtPayExpDate2.SetFocus
        End If
    End If
End Sub

Private Sub TxtDateJoin1_LostFocus()
    If IsDate(txtDateJoin1) Then
    Else
        If txtDateJoin1 <> "__/__/____" Then
            MsgBox "Invalid Date Format! ", vbCritical
            txtDateJoin1.SetFocus
        End If
    End If
End Sub

Private Sub TxtDateJoin2_LostFocus()
    If IsDate(txtDateJoin2) Then
    Else
        If txtDateJoin2 <> "__/__/____" Then
            MsgBox "Invalid Date Format! ", vbCritical
            txtDateJoin2.SetFocus
        End If
    End If
End Sub

Private Sub txtDOB1_LostFocus()
    If IsDate(txtDOB1) Then
    Else
        If txtDOB1 <> "__/__/____" Then
            MsgBox "Invalid Date Format! ", vbCritical
            txtDOB1.SetFocus
        End If
    End If
End Sub

Private Sub txtDOB2_LostFocus()
    If IsDate(txtDOB2) Then
    Else
        If txtDOB2 <> "__/__/____" Then
            MsgBox "Invalid Date Format! ", vbCritical
            txtDOB2.SetFocus
        End If
    End If
End Sub
'=========== Check Date Format ===========================


'================ List Diag. By PostCode Listing ===========================
Private Sub DiagByPCodeListing(Optional strAppend As String)
Dim rst2 As New ADODB.Recordset
Dim Record As ListItem
Dim sql As String
Dim sql2 As String
    
    sql = ""
    lstDiagPCodeList.listitems.Clear
    
    If Mid(cboSelection, 1, 1) = "0" Then
        sql = "TOP 10 "
    End If
    If Mid(cboSelection, 1, 1) = "1" Then
        sql = "TOP 20 "
    End If
    
    sql = " SELECT " & sql & "FinalDiag, MBMPostCode, ICDDescription, COUNT(FinalDiag) AS Total " & _
          " From VIEW_AnalysisReportNew6 Where 0 = 0 " & _
          " GROUP BY FinalDiag, MBMPostCode, ICDDescription ORDER BY Total DESC "
     
    rst2.Open sql, medixmasterconn, adOpenKeyset, adLockOptimistic
        
            Do Until rst2.EOF
                   
            With rst2
                If Len(rst2!FinalDiag) Then
                    Set Record = lstDiagPCodeList.listitems.Add(, , rst2!FinalDiag)
                    DiagCode = Trim(rst2!FinalDiag)
                
                    If Len(rst2!ICDDescription) Then
                        Record.SubItems(1) = Trim(rst2!ICDDescription)
                    End If
                    
                    If Len(rst2!MBMPostCode) Then
                        Record.SubItems(2) = rst2!MBMPostCode
                        PostCode = Trim(rst2!MBMPostCode)
                    End If
                    
                    Call CountDiagPCodeRecords
                    
                    If Len(rst2!Total) Then
                        Record.SubItems(3) = rst2!Total
                    End If
                    
                    If Len(TotalActual) Then
                       Record.SubItems(4) = Format(TotalActual, "#,##0.00")
                    End If
                    If Len(TotalApproved) Then
                        Record.SubItems(5) = Format(TotalApproved, "#,##0.00")
                    End If
            End If
            End With
            rst2.MoveNext
         Loop
    'End If
    rst2.Close
    Set rst2 = Nothing
End Sub
'================ List Diag. By PostCode Listing ===========================

'================ Final Diag. + PostCode Top 10/20 Listing ===========================
Private Sub DiagByPCodeTop1020()
Dim rst2 As New ADODB.Recordset
Dim Record As ListItem
Dim sql As String
       
    lstDiagPCodeList.listitems.Clear
    
    If Mid(cboSelection, 1, 1) = "0" Then
    
        sql = " SELECT TOP 10 FinalDiag, ICDDescription, MBMPostCode, COUNT(FinalDiag) AS Total " & _
              " From VIEW_AnalysisReportNew6 " & _
              " Where (CASStatus = 'O' OR CASStatus = 'C') AND CASClaimability = 'A' " & _
              " GROUP BY FinalDiag, ICDDescription, MBMPostCode ORDER BY Total DESC "
         
    ElseIf Mid(cboSelection, 1, 1) = "1" Then
        sql = " SELECT TOP 20 FinalDiag, ICDDescription, MBMPostCode, COUNT(FinalDiag) AS Total " & _
              " From VIEW_AnalysisReportNew6 " & _
              " Where (CASStatus = 'O' OR CASStatus = 'C') AND CASClaimability = 'A' " & _
              " GROUP BY FinalDiag, ICDDescription, MBMPostCode ORDER BY Total DESC "
    End If
    
    rst2.Open sql, medixmasterconn, adOpenKeyset, adLockOptimistic
    
        Do Until rst2.EOF
                    
            With rst2
                Set Record = lstDiagPCodeList.listitems.Add(, , rst2!FinalDiag)
                DiagCode = Trim(rst2!FinalDiag)
                                               
                If Len(rst2!ICDDescription) Then
                    Record.SubItems(1) = Trim(rst2!ICDDescription)
                End If
                
                If Len(rst2!MBMPostCode) Then
                    Record.SubItems(2) = rst2!MBMPostCode
                    PostCode = Trim(rst2!MBMPostCode)
                End If
                                
                Call CountDiagPCodeRecords(strAppend)
                
                If Len(rst2!Total) Then
                    Record.SubItems(3) = rst2!Total
                End If
                
                If Len(TotalActual) Then
                   Record.SubItems(4) = Format(TotalActual, "#,##0.00")
                End If
                If Len(TotalApproved) Then
                    Record.SubItems(5) = Format(TotalApproved, "#,##0.00")
                End If
            End With
            rst2.MoveNext
        Loop
    rst2.Close
    Set rst2 = Nothing
End Sub
'================ Final Diag. + PostCode Top 10/20 Listing ===========================


'================ Count Final Diag. + PostCode Records ===========================
Private Sub CountDiagPCodeRecords(Optional strAppend As String)
Dim rst2 As New ADODB.Recordset
Dim Record As ListItem
Dim sql As String
       
    TotalActual = 0
    TotalApproved = 0
            
    ListView1.listitems.Clear
            
    sql = " Select * From VIEW_AnalysisReportNew6 Where FinalDiag = '" & Trim(DiagCode) & "' AND MBMPostCode = '" & Trim(PostCode) & "'"
            
    sql2 = " AND (CASStatus = 'O' OR CASStatus = 'C') AND CASClaimability = 'A' "
    
    If Mid(cboSelection, 1, 1) = "2" Then
    
        If Len(Trim(strAppend)) Then
            sql = sql + strAppend
        End If
    
    ElseIf Mid(cboSelection, 1, 1) = "0" Or Mid(cboSelection, 1, 1) = "1" Then
        sql = sql + sql2
    End If
    
    
    rst2.Open sql, medixmasterconn, adOpenKeyset, adLockOptimistic
    
        Do Until rst2.EOF
        
                                
            With rst2
                Set Record = ListView1.listitems.Add(, , rst2!FinalDiag)
                              
                If Len(rst2!MBMPostCode) Then
                    Record.SubItems(1) = rst2!MBMPostCode
                End If
                
                If Len(rst2!CLMTotalActual) Then
                   Record.SubItems(2) = rst2!CLMTotalActual
                   TotalActual = TotalActual + rst2!CLMTotalActual
                End If
                
                If Len(rst2!CLMTotalApproved) Then
                   Record.SubItems(3) = rst2!CLMTotalApproved
                   TotalApproved = TotalApproved + rst2!CLMTotalApproved
                End If
                
            End With
            rst2.MoveNext
        Loop
    rst2.Close
    Set rst2 = Nothing
End Sub
'================ Count Final Diag. + PostCode Records ===========================

' =================== Search Monthy Admission/Discharge Date Results ======================
Private Function SearchMonthyReports()

    If TxtDate1 = "__/__/____" Then
        MsgBox "Please Enter Date", vbOKOnly + vbCritical, DialogBoxTitle
        TxtDate1.SetFocus
    ElseIf TxtDate2 = "__/__/____" Then
        MsgBox "Please Enter Date", vbOKOnly + vbCritical, DialogBoxTitle
        TxtDate2.SetFocus
    ElseIf Mid(TxtDate1, 4, 2) <> Mid(TxtDate2, 4, 2) Or Mid(TxtDate1, 7, 4) <> Mid(TxtDate2, 7, 4) Then
        MsgBox "Invalid date format", vbOKOnly + vbCritical, DialogBoxTitle
    
    Else
        strAppend = ""
    
        If Option22.value = True Then
        
            If (TxtDate1) <> "__/__/____" And IsDate(TxtDate1) = True _
                And (TxtDate2) <> "__/__/____" And IsDate(TxtDate2) = True Then
                sql = ""
                sql = " AND CasDateAdmitted >= '" & Format(Trim(TxtDate1), "mm/dd/yyyy") & _
                      "' And CasDateAdmitted <= '" & Format(Trim(TxtDate2), "mm/dd/yyyy") & "'"
                strAppend = strAppend + sql
            End If
            
        ElseIf Option23.value = True Then
            
            If (TxtDate1) <> "__/__/____" And IsDate(TxtDate1) = True _
                And (TxtDate2) <> "__/__/____" And IsDate(TxtDate2) = True Then
                sql = ""
                sql = " AND CASDischargeDate >= '" & Format(Trim(TxtDate1), "mm/dd/yyyy") & _
                      "' And CASDischargeDate <= '" & Format(Trim(TxtDate2), "mm/dd/yyyy") & "'"
                strAppend = strAppend + sql
            End If
            
        End If
        
        If Option24.value = True Then
            sql = ""
            sql = "AND CASStatus = 'O' "
            strAppend = strAppend + sql
        End If
        
        If Option25.value = True Then
            sql = ""
            sql = "AND CASStatus = 'C'"
            strAppend = strAppend + sql
        End If
        
        If Option26.value = True Then
            sql = ""
            sql = "AND CASStatus = 'D'"
            strAppend = strAppend + sql
        End If
        
        If Option27.value = True Then
            sql = ""
            sql = "AND (CASStatus = 'O' OR CASStatus = 'C')"
            strAppend = strAppend + sql
        End If
        
        If Option28.value = True Then
            sql = ""
            sql = "AND CASClaimability = 'A'"
            strAppend = strAppend + sql
        End If
        
        If Option29.value = True Then
            sql = ""
            sql = "AND CASClaimability = 'R'"
            strAppend = strAppend + sql
        End If
    
        Call MonthyReportListing(strAppend)
    End If
End Function
' =================== Search Monthy Admission/ Discharge Date Results ======================

'================ Monthy Reports Listing ===========================
Private Sub MonthyReportListing(Optional strAppend As String)
Dim rst2 As New ADODB.Recordset
Dim Record As ListItem
Dim sql As String
       
    ListView2.listitems.Clear
    
            
    sql = " SELECT Distinct PAYCode From VIEW_AnalysisReportNew Where 0 = 0 Order by PAYCode"
    
    rst2.Open sql, medixmasterconn, adOpenKeyset, adLockOptimistic
    
        Do Until rst2.EOF
                    
            With rst2
                If Len(rst2!PAYCode) Then
                
                Set Record = ListView2.listitems.Add(, , rst2!PAYCode)
                Payor = rst2!PAYCode
                
                If Option20.value = True And Option34.value = True Then
                    Call SearchAdmissionRecords(strAppend)
                ElseIf Option21.value = True And Option34.value = True Then
                    Call SearchDischargeRecords(strAppend)
                End If
                
                If Len(AdmissionDate1) And Len(Day1) Then
                    Record.SubItems(1) = AdmissionDate1 & "-" & Day1
                End If
                
                If Len(Total1) Then
                    Record.SubItems(2) = Total1
                End If
                
                If Len(AdmissionDate2) And Len(Day2) Then
                    Record.SubItems(3) = AdmissionDate2 & "-" & Day2
                End If
                
                If Len(Total2) Then
                    Record.SubItems(4) = Total2
                End If
                
                If Len(AdmissionDate3) And Len(Day3) Then
                    Record.SubItems(5) = AdmissionDate3 & "-" & Day3
                End If
                
                If Len(Total3) Then
                    Record.SubItems(6) = Total3
                End If
                
                If Len(AdmissionDate4) And Len(Day4) Then
                    Record.SubItems(7) = AdmissionDate4 & "-" & Day4
                End If
                
                If Len(Total4) Then
                    Record.SubItems(8) = Total4
                End If
                
                If Len(AdmissionDate5) And Len(Day5) Then
                    Record.SubItems(9) = AdmissionDate5 & "-" & Day5
                End If
                
                If Len(Total5) Then
                    Record.SubItems(10) = Total5
                End If
                
                If Len(AdmissionDate6) And Len(Day6) Then
                    Record.SubItems(11) = AdmissionDate6 & "-" & Day6
                End If
                
                If Len(Total6) Then
                    Record.SubItems(12) = Total6
                End If
                
                If Len(AdmissionDate7) And Len(Day7) Then
                    Record.SubItems(13) = AdmissionDate7 & "-" & Day7
                End If
                
                If Len(Total7) Then
                    Record.SubItems(14) = Total7
                End If
                
                If Len(AdmissionDate8) And Len(Day8) Then
                    Record.SubItems(15) = AdmissionDate8 & "-" & Day8
                End If
                
                If Len(Total8) Then
                    Record.SubItems(16) = Total8
                End If
                
                If Len(AdmissionDate9) And Len(Day9) Then
                    Record.SubItems(17) = AdmissionDate9 & "-" & Day9
                End If
                
                If Len(Total9) Then
                    Record.SubItems(18) = Total9
                End If
                
                If Len(AdmissionDate10) And Len(Day10) Then
                    Record.SubItems(19) = AdmissionDate10 & "-" & Day10
                End If
                
                If Len(Total10) Then
                    Record.SubItems(20) = Total10
                End If
                
                If Len(AdmissionDate11) And Len(Day11) Then
                    Record.SubItems(21) = AdmissionDate11 & "-" & Day11
                End If
                
                If Len(Total11) Then
                    Record.SubItems(22) = Total11
                End If
                
                If Len(AdmissionDate12) And Len(Day12) Then
                    Record.SubItems(23) = AdmissionDate12 & "-" & Day12
                End If
                
                If Len(Total12) Then
                    Record.SubItems(24) = Total12
                End If
                
                If Len(AdmissionDate13) And Len(Day13) Then
                    Record.SubItems(25) = AdmissionDate13 & "-" & Day13
                End If
                
                If Len(Total13) Then
                    Record.SubItems(26) = Total13
                End If
                
                If Len(AdmissionDate14) And Len(Day14) Then
                    Record.SubItems(27) = AdmissionDate14 & "-" & Day14
                End If
                
                If Len(Total14) Then
                    Record.SubItems(28) = Total14
                End If
                
                If Len(AdmissionDate15) And Len(Day15) Then
                    Record.SubItems(29) = AdmissionDate15 & "-" & Day15
                End If
                
                If Len(Total15) Then
                    Record.SubItems(30) = Total15
                End If
                
                If Len(AdmissionDate16) And Len(Day16) Then
                    Record.SubItems(31) = AdmissionDate16 & "-" & Day16
                End If
                
                If Len(Total16) Then
                    Record.SubItems(32) = Total6
                End If
                
                If Len(AdmissionDate17) And Len(Day17) Then
                    Record.SubItems(33) = AdmissionDate17 & "-" & Day17
                End If
                
                If Len(Total17) Then
                    Record.SubItems(34) = Total17
                End If
                
                If Len(AdmissionDate18) And Len(Day18) Then
                    Record.SubItems(35) = AdmissionDate18 & "-" & Day18
                End If
                
                If Len(Total18) Then
                    Record.SubItems(36) = Total8
                End If
                
                If Len(AdmissionDate19) And Len(Day19) Then
                    Record.SubItems(37) = AdmissionDate19 & "-" & Day19
                End If
                
                If Len(Total19) Then
                    Record.SubItems(38) = Total19
                End If
                
                If Len(AdmissionDate20) And Len(Day20) Then
                    Record.SubItems(39) = AdmissionDate20 & "-" & Day20
                End If
                
                If Len(Total20) Then
                    Record.SubItems(40) = Total20
                End If
                
                If Len(AdmissionDate21) And Len(Day21) Then
                    Record.SubItems(41) = AdmissionDate21 & "-" & Day21
                End If
                
                If Len(Total21) Then
                    Record.SubItems(42) = Total21
                End If
                
                If Len(AdmissionDate22) And Len(Day22) Then
                    Record.SubItems(43) = AdmissionDate22 & "-" & Day22
                End If
                
                If Len(Total22) Then
                    Record.SubItems(44) = Total22
                End If
                
                If Len(AdmissionDate23) And Len(Day23) Then
                    Record.SubItems(45) = AdmissionDate23 & "-" & Day23
                End If
                
                If Len(Total23) Then
                    Record.SubItems(46) = Total23
                End If
                
                If Len(AdmissionDate24) And Len(Day24) Then
                    Record.SubItems(47) = AdmissionDate24 & "-" & Day24
                End If
                
                If Len(Total24) Then
                    Record.SubItems(48) = Total24
                End If
                
                If Len(AdmissionDate25) And Len(Day24) Then
                    Record.SubItems(49) = AdmissionDate25 & "-" & Day25
                End If
                
                If Len(Total25) Then
                    Record.SubItems(50) = Total25
                End If
                
                If Len(AdmissionDate26) And Len(Day26) Then
                    Record.SubItems(51) = AdmissionDate26 & "-" & Day26
                End If
                
                If Len(Total26) Then
                    Record.SubItems(52) = Total26
                End If
                
                If Len(AdmissionDate27) And Len(Day27) Then
                    Record.SubItems(53) = AdmissionDate27 & "-" & Day27
                End If
                
                If Len(Total27) Then
                    Record.SubItems(54) = Total27
                End If
                
                If Len(AdmissionDate28) And Len(Day28) Then
                    Record.SubItems(55) = AdmissionDate28 & "-" & Day28
                End If
                
                If Len(Total28) Then
                    Record.SubItems(56) = Total28
                End If
                
                If Len(AdmissionDate29) And Len(Day29) Then
                    Record.SubItems(57) = AdmissionDate29 & "-" & Day29
                End If
                
                If Len(Total29) Then
                    Record.SubItems(58) = Total29
                End If
                
                If Len(AdmissionDate30) And Len(Day30) Then
                    Record.SubItems(59) = AdmissionDate30 & "-" & Day30
                End If
                
                If Len(Total30) Then
                    Record.SubItems(60) = Total30
                End If
                
                If Len(AdmissionDate31) And Len(Day31) Then
                    Record.SubItems(61) = AdmissionDate31 & "-" & Day31
                End If
                
                If Len(Total31) Then
                    Record.SubItems(62) = Total31
                End If
                End If
            End With
            rst2.MoveNext
        Loop
    rst2.Close
    Set rst2 = Nothing
End Sub
'================ Admission Listing ===========================

'==================== Search One Month Admission Records ===========================
Private Sub SearchAdmissionRecords(Optional strAppend As String)
Dim rst2 As New ADODB.Recordset
Dim Record As ListItem
Dim sql As String

    ListView3.listitems.Clear
    Call ClearBoolean
    
    sql = " Select PAYCode, CASDateAdmitted, Count(PAYCode) AS Total From VIEW_AnalysisReportNew Where PAYCode = '" & Trim(Payor) & "'"
    
    If Len(Trim(strAppend)) Then
        sql = sql + strAppend + (" Group By PAYCode, CASDateAdmitted Order by CASDateAdmitted")
    End If
    
    rst2.Open sql, medixmasterconn, adOpenKeyset, adLockOptimistic
        
                
        Do Until rst2.EOF
        
            With rst2
            
                TxtDay.Text = ""
                txtDate.Text = ""
                
                Set Record = ListView3.listitems.Add(, , rst2!PAYCode)

                
                If one = False Then
                    If Len(rst2!CASDateAdmitted) Then
                        Record.SubItems(1) = rst2!CASDateAdmitted
                        AdmissionDate1 = Record.SubItems(1)
                        txtDate = Record.SubItems(1)
                        Call CheckDay
                        Day1 = Day
                        Record.SubItems(1) = AdmissionDate1 & "-" & Day1
                        Record.SubItems(2) = rst2!Total
                        Total1 = rst2!Total
                        one = True
                    End If
                
                ElseIf two = False Then
                    If Len(rst2!CASDateAdmitted) Then
                        Record.SubItems(1) = rst2!CASDateAdmitted
                        AdmissionDate2 = Record.SubItems(1)
                        txtDate = Record.SubItems(1)
                        Call CheckDay
                        Day2 = Day
                        Record.SubItems(1) = AdmissionDate2 & "-" & Day2
                        Record.SubItems(2) = rst2!Total
                        Total2 = rst2!Total
                        two = True
                    End If
                
                ElseIf three = False Then
                    If Len(rst2!CASDateAdmitted) Then
                        Record.SubItems(1) = rst2!CASDateAdmitted
                        AdmissionDate3 = Record.SubItems(1)
                        txtDate = Record.SubItems(1)
                        Call CheckDay
                        Day3 = Day
                        Record.SubItems(1) = AdmissionDate3 & "-" & Day3
                        Record.SubItems(2) = rst2!Total
                        Total3 = rst2!Total
                        three = True
                    End If
                
                ElseIf four = False Then
                    If Len(rst2!CASDateAdmitted) Then
                        Record.SubItems(1) = rst2!CASDateAdmitted
                        AdmissionDate4 = Record.SubItems(1)
                        txtDate = Record.SubItems(1)
                        Call CheckDay
                        Day4 = Day
                        Record.SubItems(1) = AdmissionDate4 & "-" & Day4
                        Record.SubItems(2) = rst2!Total
                        Total4 = rst2!Total
                        four = True
                    End If
                ElseIf five = False Then
                    If Len(rst2!CASDateAdmitted) Then
                        Record.SubItems(1) = rst2!CASDateAdmitted
                        AdmissionDate5 = Record.SubItems(1)
                        txtDate = Record.SubItems(1)
                        Call CheckDay
                        Day5 = Day
                        Record.SubItems(1) = AdmissionDate5 & "-" & Day5
                        Record.SubItems(2) = rst2!Total
                        Total5 = rst2!Total
                        five = True
                    End If
                ElseIf six = False Then
                    If Len(rst2!CASDateAdmitted) Then
                        Record.SubItems(1) = rst2!CASDateAdmitted
                        AdmissionDate6 = Record.SubItems(1)
                        txtDate = Record.SubItems(1)
                        Call CheckDay
                        Day6 = Day
                        Record.SubItems(1) = AdmissionDate6 & "-" & Day6
                        Record.SubItems(2) = rst2!Total
                        Total6 = rst2!Total
                        six = True
                    End If
                ElseIf seven = False Then
                    If Len(rst2!CASDateAdmitted) Then
                        Record.SubItems(1) = rst2!CASDateAdmitted
                        AdmissionDate7 = Record.SubItems(1)
                        txtDate = Record.SubItems(1)
                        Call CheckDay
                        Day7 = Day
                        Record.SubItems(1) = AdmissionDate7 & "-" & Day7
                        Record.SubItems(2) = rst2!Total
                        Total7 = rst2!Total
                        seven = True
                    End If
                ElseIf eight = False Then
                    If Len(rst2!CASDateAdmitted) Then
                        Record.SubItems(1) = rst2!CASDateAdmitted
                        AdmissionDate8 = Record.SubItems(1)
                        txtDate = Record.SubItems(1)
                        Call CheckDay
                        Day8 = Day
                        Record.SubItems(1) = AdmissionDate8 & "-" & Day8
                        Record.SubItems(2) = rst2!Total
                        Total8 = rst2!Total
                        eight = True
                    End If
                ElseIf nine = False Then
                    If Len(rst2!CASDateAdmitted) Then
                        Record.SubItems(1) = rst2!CASDateAdmitted
                        AdmissionDate9 = Record.SubItems(1)
                        txtDate = Record.SubItems(1)
                        Call CheckDay
                        Day9 = Day
                        Record.SubItems(1) = AdmissionDate9 & "-" & Day9
                        Record.SubItems(2) = rst2!Total
                        Total9 = rst2!Total
                        nine = True
                    End If
                ElseIf ten = False Then
                    If Len(rst2!CASDateAdmitted) Then
                        Record.SubItems(1) = rst2!CASDateAdmitted
                        AdmissionDate10 = Record.SubItems(1)
                        txtDate = Record.SubItems(1)
                        Call CheckDay
                        Day10 = Day
                        Record.SubItems(1) = AdmissionDate10 & "-" & Day10
                        Record.SubItems(2) = rst2!Total
                        Total10 = rst2!Total
                        ten = True
                    End If
                ElseIf eleven = False Then
                    If Len(rst2!CASDateAdmitted) Then
                        Record.SubItems(1) = rst2!CASDateAdmitted
                        AdmissionDate11 = Record.SubItems(1)
                        txtDate = Record.SubItems(1)
                        Call CheckDay
                        Day11 = Day
                        Record.SubItems(1) = AdmissionDate11 & "-" & Day11
                        Record.SubItems(2) = rst2!Total
                        Total11 = rst2!Total
                        eleven = True
                    End If
                ElseIf twelve = False Then
                    If Len(rst2!CASDateAdmitted) Then
                        Record.SubItems(1) = rst2!CASDateAdmitted
                        AdmissionDate12 = Record.SubItems(1)
                        txtDate = Record.SubItems(1)
                        Call CheckDay
                        Day12 = Day
                        Record.SubItems(1) = AdmissionDate12 & "-" & Day12
                        Record.SubItems(2) = rst2!Total
                        Total12 = rst2!Total
                        twelve = True
                    End If
                ElseIf thirteen = False Then
                    If Len(rst2!CASDateAdmitted) Then
                        Record.SubItems(1) = rst2!CASDateAdmitted
                        AdmissionDate13 = Record.SubItems(1)
                        txtDate = Record.SubItems(1)
                        Call CheckDay
                        Day13 = Day
                        Record.SubItems(1) = AdmissionDate13 & "-" & Day13
                        Record.SubItems(2) = rst2!Total
                        Total13 = rst2!Total
                        thirteen = True
                    End If
                ElseIf fourteen = False Then
                    If Len(rst2!CASDateAdmitted) Then
                        Record.SubItems(1) = rst2!CASDateAdmitted
                        AdmissionDate14 = Record.SubItems(1)
                        txtDate = Record.SubItems(1)
                        Call CheckDay
                        Day14 = Day
                        Record.SubItems(1) = AdmissionDate14 & "-" & Day14
                        Record.SubItems(2) = rst2!Total
                        Total14 = rst2!Total
                        fourteen = True
                    End If
                ElseIf fiftteen = False Then
                    If Len(rst2!CASDateAdmitted) Then
                        Record.SubItems(1) = rst2!CASDateAdmitted
                        AdmissionDate15 = Record.SubItems(1)
                        txtDate = Record.SubItems(1)
                        Call CheckDay
                        Day15 = Day
                        Record.SubItems(1) = AdmissionDate15 & "-" & Day15
                        Record.SubItems(2) = rst2!Total
                        Total15 = rst2!Total
                        fiftteen = True
                    End If
                ElseIf sixteen = False Then
                    If Len(rst2!CASDateAdmitted) Then
                        Record.SubItems(1) = rst2!CASDateAdmitted
                        AdmissionDate16 = Record.SubItems(1)
                        txtDate = Record.SubItems(1)
                        Call CheckDay
                        Day16 = Day
                        Record.SubItems(1) = AdmissionDate16 & "-" & Day16
                        Record.SubItems(2) = rst2!Total
                        Total16 = rst2!Total
                        sixteen = True
                    End If
                ElseIf seventeen = False Then
                    If Len(rst2!CASDateAdmitted) Then
                        Record.SubItems(1) = rst2!CASDateAdmitted
                        AdmissionDate17 = Record.SubItems(1)
                        txtDate = Record.SubItems(1)
                        Call CheckDay
                        Day17 = Day
                        Record.SubItems(1) = AdmissionDate17 & "-" & Day17
                        Record.SubItems(2) = rst2!Total
                        Total17 = rst2!Total
                        seventeen = True
                    End If
                ElseIf eightteen = False Then
                    If Len(rst2!CASDateAdmitted) Then
                        Record.SubItems(1) = rst2!CASDateAdmitted
                        AdmissionDate18 = Record.SubItems(1)
                        txtDate = Record.SubItems(1)
                        Call CheckDay
                        Day18 = Day
                        Record.SubItems(1) = AdmissionDate18 & "-" & Day18
                        Record.SubItems(2) = rst2!Total
                        Total18 = rst2!Total
                        eightteen = True
                    End If
                ElseIf nineteen = False Then
                    If Len(rst2!CASDateAdmitted) Then
                        Record.SubItems(1) = rst2!CASDateAdmitted
                        AdmissionDate19 = Record.SubItems(1)
                        txtDate = Record.SubItems(1)
                        Call CheckDay
                        Day19 = Day
                        Record.SubItems(1) = AdmissionDate19 & "-" & Day19
                        Record.SubItems(2) = rst2!Total
                        Total19 = rst2!Total
                        nineteen = True
                    End If
                ElseIf twenty = False Then
                    If Len(rst2!CASDateAdmitted) Then
                        Record.SubItems(1) = rst2!CASDateAdmitted
                        AdmissionDate20 = Record.SubItems(1)
                        txtDate = Record.SubItems(1)
                        Call CheckDay
                        Day20 = Day
                        Record.SubItems(1) = AdmissionDate20 & "-" & Day20
                        Record.SubItems(2) = rst2!Total
                        Total20 = rst2!Total
                        twenty = True
                    End If
                ElseIf twentyone = False Then
                    If Len(rst2!CASDateAdmitted) Then
                        Record.SubItems(1) = rst2!CASDateAdmitted
                        AdmissionDate21 = Record.SubItems(1)
                        txtDate = Record.SubItems(1)
                        Call CheckDay
                        Day21 = Day
                        Record.SubItems(1) = AdmissionDate21 & "-" & Day21
                        Record.SubItems(2) = rst2!Total
                        Total21 = rst2!Total
                        twentyone = True
                    End If
                ElseIf twentytwo = False Then
                    If Len(rst2!CASDateAdmitted) Then
                        Record.SubItems(1) = rst2!CASDateAdmitted
                        AdmissionDate22 = Record.SubItems(1)
                        txtDate = Record.SubItems(1)
                        Call CheckDay
                        Day22 = Day
                        Record.SubItems(1) = AdmissionDate22 & "-" & Day22
                        Record.SubItems(2) = rst2!Total
                        Total22 = rst2!Total
                        twentytwo = True
                    End If
                ElseIf twentythree = False Then
                    If Len(rst2!CASDateAdmitted) Then
                        Record.SubItems(1) = rst2!CASDateAdmitted
                        AdmissionDate23 = Record.SubItems(1)
                        txtDate = Record.SubItems(1)
                        Call CheckDay
                        Day20 = Day
                        Record.SubItems(1) = AdmissionDate23 & "-" & Day23
                        Record.SubItems(2) = rst2!Total
                        Total23 = rst2!Total
                        twentythree = True
                    End If
                ElseIf twentyfour = False Then
                    If Len(rst2!CASDateAdmitted) Then
                        Record.SubItems(1) = rst2!CASDateAdmitted
                        AdmissionDate24 = Record.SubItems(1)
                        txtDate = Record.SubItems(1)
                        Call CheckDay
                        Day24 = Day
                        Record.SubItems(1) = AdmissionDate24 & "-" & Day24
                        Record.SubItems(2) = rst2!Total
                        Total24 = rst2!Total
                        twentyfour = True
                    End If
                ElseIf twentyfive = False Then
                    If Len(rst2!CASDateAdmitted) Then
                        Record.SubItems(1) = rst2!CASDateAdmitted
                        AdmissionDate25 = Record.SubItems(1)
                        txtDate = Record.SubItems(1)
                        Call CheckDay
                        Day25 = Day
                        Record.SubItems(1) = AdmissionDate25 & "-" & Day25
                        Record.SubItems(2) = rst2!Total
                        Total25 = rst2!Total
                        twentyfive = True
                    End If
                ElseIf twentysix = False Then
                    If Len(rst2!CASDateAdmitted) Then
                        Record.SubItems(1) = rst2!CASDateAdmitted
                        AdmissionDate26 = Record.SubItems(1)
                        txtDate = Record.SubItems(1)
                        Call CheckDay
                        Day26 = Day
                        Record.SubItems(1) = AdmissionDate26 & "-" & Day26
                        Record.SubItems(2) = rst2!Total
                        Total26 = rst2!Total
                        twentysix = True
                    End If
                ElseIf twentyseven = False Then
                    If Len(rst2!CASDateAdmitted) Then
                        Record.SubItems(1) = rst2!CASDateAdmitted
                        AdmissionDate27 = Record.SubItems(1)
                        txtDate = Record.SubItems(1)
                        Call CheckDay
                        Day27 = Day
                        Record.SubItems(1) = AdmissionDate27 & "-" & Day27
                        Record.SubItems(2) = rst2!Total
                        Total27 = rst2!Total
                        twentyseven = True
                    End If
                ElseIf twentyeight = False Then
                    If Len(rst2!CASDateAdmitted) Then
                        Record.SubItems(1) = rst2!CASDateAdmitted
                        AdmissionDate28 = Record.SubItems(1)
                        txtDate = Record.SubItems(1)
                        Call CheckDay
                        Day28 = Day
                        Record.SubItems(1) = AdmissionDate28 & "-" & Day28
                        Record.SubItems(2) = rst2!Total
                        Total28 = rst2!Total
                        twentyeight = True
                    End If
                ElseIf twentynine = False Then
                    If Len(rst2!CASDateAdmitted) Then
                        Record.SubItems(1) = rst2!CASDateAdmitted
                        AdmissionDate29 = Record.SubItems(1)
                        txtDate = Record.SubItems(1)
                        Call CheckDay
                        Day29 = Day
                        Record.SubItems(1) = AdmissionDate29 & "-" & Day29
                        Record.SubItems(2) = rst2!Total
                        Total29 = rst2!Total
                        twentynine = True
                    End If
                ElseIf thirty = False Then
                    If Len(rst2!CASDateAdmitted) Then
                        Record.SubItems(1) = rst2!CASDateAdmitted
                        AdmissionDate30 = Record.SubItems(1)
                        txtDate = Record.SubItems(1)
                        Call CheckDay
                        Day30 = Day
                        Record.SubItems(1) = AdmissionDate30 & "-" & Day30
                        Record.SubItems(2) = rst2!Total
                        Total30 = rst2!Total
                        thirty = True
                    End If
                ElseIf thirtyone = False Then
                    If Len(rst2!CASDateAdmitted) Then
                        Record.SubItems(1) = rst2!CASDateAdmitted
                        AdmissionDate31 = Record.SubItems(1)
                        txtDate = Record.SubItems(1)
                        Call CheckDay
                        Day31 = Day
                        Record.SubItems(1) = AdmissionDate31 & "-" & Day31
                        Record.SubItems(2) = rst2!Total
                        Total31 = rst2!Total
                        thirtyone = True
                    End If
                End If
            End With
            rst2.MoveNext
        Loop
    rst2.Close
    Set rst2 = Nothing
End Sub
'==================== Search One Month Admission Records ===========================

'==================== Search One Month Discharge Records ===========================
Private Sub SearchDischargeRecords(Optional strAppend3 As String)
Dim rst2 As New ADODB.Recordset
Dim Record As ListItem
Dim sql As String

    ListView1.listitems.Clear
    Call ClearBoolean
    
    sql = " Select PAYCode, CASDischargeDate, Count(PAYCode) AS Total From VIEW_Admit_Discharge_Report Where PAYCode = '" & Trim(Payor) & "'"

    
    If Len(Trim(strAppend3)) Then
        sql = sql + strAppend3 + (" Group By PAYCode, CASDischargeDate ")
    End If
    
    rst2.Open sql, medixmasterconn, adOpenKeyset, adLockOptimistic
        
        Do Until rst2.EOF
        
            With rst2
            
                TxtDay.Text = ""
                txtDate.Text = ""
                
                Set Record = ListView1.listitems.Add(, , rst2!PAYCode)

                
                If one = False Then
                    If Len(rst2!CASDischargeDate) Then
                        Record.SubItems(1) = rst2!CASDischargeDate
                        AdmissionDate1 = Record.SubItems(1)
                        txtDate = Record.SubItems(1)
                        Call CheckDay
                        Day1 = Day
                        Record.SubItems(1) = AdmissionDate1 & "-" & Day1
                        Record.SubItems(2) = rst2!Total
                        Total1 = rst2!Total
                        one = True
                    End If
                
                ElseIf two = False Then
                    If Len(rst2!CASDischargeDate) Then
                        Record.SubItems(1) = rst2!CASDischargeDate
                        AdmissionDate2 = Record.SubItems(1)
                        txtDate = Record.SubItems(1)
                        Call CheckDay
                        Day2 = Day
                        Record.SubItems(1) = AdmissionDate2 & "-" & Day2
                        Record.SubItems(2) = rst2!Total
                        Total2 = rst2!Total
                        two = True
                    End If
                
                ElseIf three = False Then
                    If Len(rst2!CASDischargeDate) Then
                        Record.SubItems(1) = rst2!CASDischargeDate
                        AdmissionDate3 = Record.SubItems(1)
                        txtDate = Record.SubItems(1)
                        Call CheckDay
                        Day3 = Day
                        Record.SubItems(1) = AdmissionDate3 & "-" & Day3
                        Record.SubItems(2) = rst2!Total
                        Total3 = rst2!Total
                        three = True
                    End If
                
                ElseIf four = False Then
                    If Len(rst2!CASDischargeDate) Then
                        Record.SubItems(1) = rst2!CASDischargeDate
                        AdmissionDate4 = Record.SubItems(1)
                        txtDate = Record.SubItems(1)
                        Call CheckDay
                        Day4 = Day
                        Record.SubItems(1) = AdmissionDate4 & "-" & Day4
                        Record.SubItems(2) = rst2!Total
                        Total4 = rst2!Total
                        four = True
                    End If
                ElseIf five = False Then
                    If Len(rst2!CASDischargeDate) Then
                        Record.SubItems(1) = rst2!CASDischargeDate
                        AdmissionDate5 = Record.SubItems(1)
                        txtDate = Record.SubItems(1)
                        Call CheckDay
                        Day5 = Day
                        Record.SubItems(1) = AdmissionDate5 & "-" & Day5
                        Record.SubItems(2) = rst2!Total
                        Total5 = rst2!Total
                        five = True
                    End If
                ElseIf six = False Then
                    If Len(rst2!CASDischargeDate) Then
                        Record.SubItems(1) = rst2!CASDischargeDate
                        AdmissionDate6 = Record.SubItems(1)
                        txtDate = Record.SubItems(1)
                        Call CheckDay
                        Day6 = Day
                        Record.SubItems(1) = AdmissionDate6 & "-" & Day6
                        Record.SubItems(2) = rst2!Total
                        Total6 = rst2!Total
                        six = True
                    End If
                ElseIf seven = False Then
                    If Len(rst2!CASDischargeDate) Then
                        Record.SubItems(1) = rst2!CASDischargeDate
                        AdmissionDate7 = Record.SubItems(1)
                        txtDate = Record.SubItems(1)
                        Call CheckDay
                        Day7 = Day
                        Record.SubItems(1) = AdmissionDate7 & "-" & Day7
                        Record.SubItems(2) = rst2!Total
                        Total7 = rst2!Total
                        seven = True
                    End If
                ElseIf eight = False Then
                    If Len(rst2!CASDischargeDate) Then
                        Record.SubItems(1) = rst2!CASDischargeDate
                        AdmissionDate8 = Record.SubItems(1)
                        txtDate = Record.SubItems(1)
                        Call CheckDay
                        Day8 = Day
                        Record.SubItems(1) = AdmissionDate8 & "-" & Day8
                        Record.SubItems(2) = rst2!Total
                        Total8 = rst2!Total
                        eight = True
                    End If
                ElseIf nine = False Then
                    If Len(rst2!CASDischargeDate) Then
                        Record.SubItems(1) = rst2!CASDischargeDate
                        AdmissionDate9 = Record.SubItems(1)
                        txtDate = Record.SubItems(1)
                        Call CheckDay
                        Day9 = Day
                        Record.SubItems(1) = AdmissionDate9 & "-" & Day9
                        Record.SubItems(2) = rst2!Total
                        Total9 = rst2!Total
                        nine = True
                    End If
                ElseIf ten = False Then
                    If Len(rst2!CASDischargeDate) Then
                        Record.SubItems(1) = rst2!CASDischargeDate
                        AdmissionDate10 = Record.SubItems(1)
                        txtDate = Record.SubItems(1)
                        Call CheckDay
                        Day10 = Day
                        Record.SubItems(1) = AdmissionDate10 & "-" & Day10
                        Record.SubItems(2) = rst2!Total
                        Total10 = rst2!Total
                        ten = True
                    End If
                ElseIf eleven = False Then
                    If Len(rst2!CASDischargeDate) Then
                        Record.SubItems(1) = rst2!CASDischargeDate
                        AdmissionDate11 = Record.SubItems(1)
                        txtDate = Record.SubItems(1)
                        Call CheckDay
                        Day11 = Day
                        Record.SubItems(1) = AdmissionDate11 & "-" & Day11
                        Record.SubItems(2) = rst2!Total
                        Total11 = rst2!Total
                        eleven = True
                    End If
                ElseIf twelve = False Then
                    If Len(rst2!CASDischargeDate) Then
                        Record.SubItems(1) = rst2!CASDischargeDate
                        AdmissionDate12 = Record.SubItems(1)
                        txtDate = Record.SubItems(1)
                        Call CheckDay
                        Day12 = Day
                        Record.SubItems(1) = AdmissionDate12 & "-" & Day12
                        Record.SubItems(2) = rst2!Total
                        Total12 = rst2!Total
                        twelve = True
                    End If
                ElseIf thirteen = False Then
                    If Len(rst2!CASDischargeDate) Then
                        Record.SubItems(1) = rst2!CASDischargeDate
                        AdmissionDate13 = Record.SubItems(1)
                        txtDate = Record.SubItems(1)
                        Call CheckDay
                        Day13 = Day
                        Record.SubItems(1) = AdmissionDate13 & "-" & Day13
                        Record.SubItems(2) = rst2!Total
                        Total13 = rst2!Total
                        thirteen = True
                    End If
                ElseIf fourteen = False Then
                    If Len(rst2!CASDischargeDate) Then
                        Record.SubItems(1) = rst2!CASDischargeDate
                        AdmissionDate14 = Record.SubItems(1)
                        txtDate = Record.SubItems(1)
                        Call CheckDay
                        Day14 = Day
                        Record.SubItems(1) = AdmissionDate14 & "-" & Day14
                        Record.SubItems(2) = rst2!Total
                        Total14 = rst2!Total
                        fourteen = True
                    End If
                ElseIf fiftteen = False Then
                    If Len(rst2!CASDischargeDate) Then
                        Record.SubItems(1) = rst2!CASDischargeDate
                        AdmissionDate15 = Record.SubItems(1)
                        txtDate = Record.SubItems(1)
                        Call CheckDay
                        Day15 = Day
                        Record.SubItems(1) = AdmissionDate15 & "-" & Day15
                        Record.SubItems(2) = rst2!Total
                        Total15 = rst2!Total
                        fiftteen = True
                    End If
                ElseIf sixteen = False Then
                    If Len(rst2!CASDischargeDate) Then
                        Record.SubItems(1) = rst2!CASDischargeDate
                        AdmissionDate16 = Record.SubItems(1)
                        txtDate = Record.SubItems(1)
                        Call CheckDay
                        Day16 = Day
                        Record.SubItems(1) = AdmissionDate16 & "-" & Day16
                        Record.SubItems(2) = rst2!Total
                        Total16 = rst2!Total
                        sixteen = True
                    End If
                ElseIf seventeen = False Then
                    If Len(rst2!CASDischargeDate) Then
                        Record.SubItems(1) = rst2!CASDischargeDate
                        AdmissionDate17 = Record.SubItems(1)
                        txtDate = Record.SubItems(1)
                        Call CheckDay
                        Day17 = Day
                        Record.SubItems(1) = AdmissionDate17 & "-" & Day17
                        Record.SubItems(2) = rst2!Total
                        Total17 = rst2!Total
                        seventeen = True
                    End If
                ElseIf eightteen = False Then
                    If Len(rst2!CASDischargeDate) Then
                        Record.SubItems(1) = rst2!CASDischargeDate
                        AdmissionDate18 = Record.SubItems(1)
                        txtDate = Record.SubItems(1)
                        Call CheckDay
                        Day18 = Day
                        Record.SubItems(1) = AdmissionDate18 & "-" & Day18
                        Record.SubItems(2) = rst2!Total
                        Total18 = rst2!Total
                        eightteen = True
                    End If
                ElseIf nineteen = False Then
                    If Len(rst2!CASDischargeDate) Then
                        Record.SubItems(1) = rst2!CASDischargeDate
                        AdmissionDate19 = Record.SubItems(1)
                        txtDate = Record.SubItems(1)
                        Call CheckDay
                        Day19 = Day
                        Record.SubItems(1) = AdmissionDate19 & "-" & Day19
                        Record.SubItems(2) = rst2!Total
                        Total19 = rst2!Total
                        nineteen = True
                    End If
                ElseIf twenty = False Then
                    If Len(rst2!CASDischargeDate) Then
                        Record.SubItems(1) = rst2!CASDischargeDate
                        AdmissionDate20 = Record.SubItems(1)
                        txtDate = Record.SubItems(1)
                        Call CheckDay
                        Day20 = Day
                        Record.SubItems(1) = AdmissionDate20 & "-" & Day20
                        Record.SubItems(2) = rst2!Total
                        Total20 = rst2!Total
                        twenty = True
                    End If
                ElseIf twentyone = False Then
                    If Len(rst2!CASDischargeDate) Then
                        Record.SubItems(1) = rst2!CASDischargeDate
                        AdmissionDate21 = Record.SubItems(1)
                        txtDate = Record.SubItems(1)
                        Call CheckDay
                        Day21 = Day
                        Record.SubItems(1) = AdmissionDate21 & "-" & Day21
                        Record.SubItems(2) = rst2!Total
                        Total21 = rst2!Total
                        twentyone = True
                    End If
                ElseIf twentytwo = False Then
                    If Len(rst2!CASDischargeDate) Then
                        Record.SubItems(1) = rst2!CASDischargeDate
                        AdmissionDate22 = Record.SubItems(1)
                        txtDate = Record.SubItems(1)
                        Call CheckDay
                        Day22 = Day
                        Record.SubItems(1) = AdmissionDate22 & "-" & Day22
                        Record.SubItems(2) = rst2!Total
                        Total22 = rst2!Total
                        twentytwo = True
                    End If
                ElseIf twentythree = False Then
                    If Len(rst2!CASDischargeDate) Then
                        Record.SubItems(1) = rst2!CASDischargeDate
                        AdmissionDate23 = Record.SubItems(1)
                        txtDate = Record.SubItems(1)
                        Call CheckDay
                        Day20 = Day
                        Record.SubItems(1) = AdmissionDate23 & "-" & Day23
                        Record.SubItems(2) = rst2!Total
                        Total23 = rst2!Total
                        twentythree = True
                    End If
                ElseIf twentyfour = False Then
                    If Len(rst2!CASDischargeDate) Then
                        Record.SubItems(1) = rst2!CASDischargeDate
                        AdmissionDate24 = Record.SubItems(1)
                        txtDate = Record.SubItems(1)
                        Call CheckDay
                        Day24 = Day
                        Record.SubItems(1) = AdmissionDate24 & "-" & Day24
                        Record.SubItems(2) = rst2!Total
                        Total24 = rst2!Total
                        twentyfour = True
                    End If
                ElseIf twentyfive = False Then
                    If Len(rst2!CASDischargeDate) Then
                        Record.SubItems(1) = rst2!CASDischargeDate
                        AdmissionDate25 = Record.SubItems(1)
                        txtDate = Record.SubItems(1)
                        Call CheckDay
                        Day25 = Day
                        Record.SubItems(1) = AdmissionDate25 & "-" & Day25
                        Record.SubItems(2) = rst2!Total
                        Total25 = rst2!Total
                        twentyfive = True
                    End If
                ElseIf twentysix = False Then
                    If Len(rst2!CASDischargeDate) Then
                        Record.SubItems(1) = rst2!CASDischargeDate
                        AdmissionDate26 = Record.SubItems(1)
                        txtDate = Record.SubItems(1)
                        Call CheckDay
                        Day26 = Day
                        Record.SubItems(1) = AdmissionDate26 & "-" & Day26
                        Record.SubItems(2) = rst2!Total
                        Total26 = rst2!Total
                        twentysix = True
                    End If
                ElseIf twentyseven = False Then
                    If Len(rst2!CASDischargeDate) Then
                        Record.SubItems(1) = rst2!CASDischargeDate
                        AdmissionDate27 = Record.SubItems(1)
                        txtDate = Record.SubItems(1)
                        Call CheckDay
                        Day27 = Day
                        Record.SubItems(1) = AdmissionDate27 & "-" & Day27
                        Record.SubItems(2) = rst2!Total
                        Total27 = rst2!Total
                        twentyseven = True
                    End If
                ElseIf twentyeight = False Then
                    If Len(rst2!CASDischargeDate) Then
                        Record.SubItems(1) = rst2!CASDischargeDate
                        AdmissionDate28 = Record.SubItems(1)
                        txtDate = Record.SubItems(1)
                        Call CheckDay
                        Day28 = Day
                        Record.SubItems(1) = AdmissionDate28 & "-" & Day28
                        Record.SubItems(2) = rst2!Total
                        Total28 = rst2!Total
                        twentyeight = True
                    End If
                ElseIf twentynine = False Then
                    If Len(rst2!CASDischargeDate) Then
                        Record.SubItems(1) = rst2!CASDischargeDate
                        AdmissionDate29 = Record.SubItems(1)
                        txtDate = Record.SubItems(1)
                        Call CheckDay
                        Day29 = Day
                        Record.SubItems(1) = AdmissionDate29 & "-" & Day29
                        Record.SubItems(2) = rst2!Total
                        Total29 = rst2!Total
                        twentynine = True
                    End If
                ElseIf thirty = False Then
                    If Len(rst2!CASDischargeDate) Then
                        Record.SubItems(1) = rst2!CASDischargeDate
                        AdmissionDate30 = Record.SubItems(1)
                        txtDate = Record.SubItems(1)
                        Call CheckDay
                        Day30 = Day
                        Record.SubItems(1) = AdmissionDate30 & "-" & Day30
                        Record.SubItems(2) = rst2!Total
                        Total30 = rst2!Total
                        thirty = True
                    End If
                ElseIf thirtyone = False Then
                    If Len(rst2!CASDischargeDate) Then
                        Record.SubItems(1) = rst2!CASDischargeDate
                        AdmissionDate31 = Record.SubItems(1)
                        txtDate = Record.SubItems(1)
                        Call CheckDay
                        Day31 = Day
                        Record.SubItems(1) = AdmissionDate31 & "-" & Day31
                        Record.SubItems(2) = rst2!Total
                        Total31 = rst2!Total
                        thirtyone = True
                    End If
                End If
            End With
            rst2.MoveNext
        Loop
    rst2.Close
    Set rst2 = Nothing
End Sub
'==================== Search One Month Discharge Records ===========================


'==================== Check Day ===========================
Private Sub CheckDay()
Dim intYears As Integer 'Used as a counter
Dim intNumDaysToCheck As Long 'Used to limit search
Dim tmpDate As Date 'Temporary date variable
Dim tmpMonth As Date 'Temporary month variable
Dim intPaydays As Integer 'Counts number of paydays Each month
Dim intPayDayOfWeek As Integer 'Pay Day of Week
Dim intX As Long 'Counter
    'computes number of days between pay date and 10 years from paydate
    intNumDaysToCheck = DateDiff("d", CDate(txtDate), DateAdd("yyyy", 100, CDate(txtDate)))
    intPayDayOfWeek = Weekday(CDate(txtDate))
    intPaydays = 0
    tmpMonth = DatePart("m", CDate(txtDate))
    'get paid every two weeks
    'if you get paid every friday and want to check for
    '5 paycheck months, change this to step 7 and change
    'commented value below


    For intX = 0 To intNumDaysToCheck Step 14
        'set the day we are working with
        tmpDate = CDate(txtDate) + intX
        'make sure the month hasn't changed


        If tmpMonth = DatePart("m", tmpDate) Then
            'if it is a payday


            If Weekday(tmpDate) <> intPayDayOfWeek Then
                'there is a problem
                MsgBox "An Error was encountered. Calculation cannot continue"
            Else
                'increment the paydays
                intPaydays = intPaydays + 1
                'if you have arrived at a 3 payday month, add it to the list
                'if you get weekly paychecks, change this value from 3 to 5
            End If
        Else
            'change the month
            tmpMonth = DatePart("m", tmpDate)
            'set the number of paydays to 1
            intPaydays = 1
        End If
    Next
    
    Call DateDay
End Sub
'==================== Check Day ===========================

' ================= Call Day ==============================
Private Sub DateDay()

    Select Case Weekday(CDate(txtDate))
        Case vbMonday
        TxtDay = "Monday"
        Case vbTuesday
        TxtDay = "Tuesday"
        Case vbWednesday
        TxtDay = "Wednesday"
        Case vbThursday
        TxtDay = "Thursday"
        Case vbFriday
        TxtDay = "Friday"
        Case vbSaturday
        TxtDay = "Saturday"
        Case vbSunday
        TxtDay = "Sunday"
        Case Else
        TxtDay = "Error With date"
    End Select
    
    Day = TxtDay
End Sub
' ================= Call Day ==============================

'=================== Clear String Records ========================
Public Function ClearBoolean()
    one = False
    two = False
    three = False
    four = False
    five = False
    six = False
    seven = False
    eight = False
    nine = False
    ten = False
    eleven = False
    twelve = False
    thirteen = False
    fourteen = False
    fifteen = False
    sixteen = False
    seventeen = False
    eightteen = False
    nineteen = False
    twenty = False
    twentyone = False
    twentytwo = False
    twentythree = False
    twentyfour = False
    twentyfive = False
    twentysix = False
    twentyseven = False
    twentyeight = False
    twentynine = False
    thirty = False
    thirtyone = False
    
    AdmissionDate1 = ""
    Day1 = ""
    Total1 = ""
    
    AdmissionDate2 = ""
    Day2 = ""
    Total2 = ""
    
        
    AdmissionDate3 = ""
    Day3 = ""
    Total3 = ""
    
    AdmissionDate4 = ""
    Day4 = ""
    Total4 = ""
    
    AdmissionDate5 = ""
    Day5 = ""
    Total5 = ""
    
    AdmissionDate6 = ""
    Day6 = ""
    Total6 = ""
    
    AdmissionDate7 = ""
    Day7 = ""
    Total7 = ""
    
    AdmissionDate8 = ""
    Day8 = ""
    Total8 = ""
    
    AdmissionDate9 = ""
    Day9 = ""
    Total9 = ""
    
    AdmissionDate10 = ""
    Day10 = ""
    Total10 = ""
    
    AdmissionDate11 = ""
    Day11 = ""
    Total11 = ""
    
    AdmissionDate12 = ""
    Day12 = ""
    Total12 = ""
    
    AdmissionDate13 = ""
    Day13 = ""
    Total13 = ""
    
    AdmissionDate14 = ""
    Day14 = ""
    Total14 = ""
    
    AdmissionDate15 = ""
    Day15 = ""
    Total15 = ""
    
    AdmissionDate16 = ""
    Day16 = ""
    Total16 = ""
    
    AdmissionDate17 = ""
    Day17 = ""
    Total17 = ""
    
    AdmissionDate18 = ""
    Day18 = ""
    Total18 = ""
    
    AdmissionDate19 = ""
    Day19 = ""
    Total19 = ""
    
    AdmissionDate20 = ""
    Day20 = ""
    Total20 = ""
    
    AdmissionDate21 = ""
    Day21 = ""
    Total21 = ""
    
    AdmissionDate22 = ""
    Day22 = ""
    Total22 = ""
    
    AdmissionDate23 = ""
    Day23 = ""
    Total23 = ""
    
    AdmissionDate24 = ""
    Day24 = ""
    Total24 = ""
    
    AdmissionDate25 = ""
    Day25 = ""
    Total25 = ""
    
    AdmissionDate26 = ""
    Day26 = ""
    Total26 = ""
    
    AdmissionDate27 = ""
    Day27 = ""
    Total27 = ""
    
    AdmissionDate28 = ""
    Day28 = ""
    Total28 = ""
    
    AdmissionDate29 = ""
    Day29 = ""
    Total29 = ""
    
    AdmissionDate30 = ""
    Day30 = ""
    Total30 = ""
    
    AdmissionDate31 = ""
    Day31 = ""
    Total31 = ""
    
End Function
'=================== Clear String Records ========================

' =================== Search Year Admission / Discharge Results ======================
Private Function SearchYearReports()

    If TxtYear = "" Then
        MsgBox "Please Enter Year", vbOKOnly + vbCritical, DialogBoxTitle
        TxtYear.SetFocus
    Else
        strAppend = ""
    
        If Len(Trim(TxtYear)) Then
            strAppend = strAppend + " AND Year = '" & Trim(TxtYear) & "'"
        End If
        
        If Option33.value = True Then
            sql = ""
            sql = "AND CASStatus = 'O' "
            strAppend = strAppend + sql
        End If
        
        If Option34.value = True Then
            sql = ""
            sql = "AND CASStatus = 'C'"
            strAppend = strAppend + sql
        End If
        
        If Option35.value = True Then
            sql = ""
            sql = "AND CASStatus = 'D'"
            strAppend = strAppend + sql
        End If
        
        If Option36.value = True Then
            sql = ""
            sql = "AND (CASStatus = 'O' OR CASStatus = 'C')"
            strAppend = strAppend + sql
        End If
        
        If Option37.value = True Then
            sql = ""
            sql = "AND CASClaimability = 'A'"
            strAppend = strAppend + sql
        End If
        
        If Option38.value = True Then
            sql = ""
            sql = "AND CASClaimability = 'R'"
            strAppend = strAppend + sql
        End If
    
        Call YearListing(strAppend)
    End If
End Function
' =================== Search Year Admission / Discharge Results ======================

'================ Admission / Discharge Year Listing ===========================
Private Sub YearListing(Optional strAppend As String)
Dim rst2 As New ADODB.Recordset
Dim Record As ListItem
Dim sql As String
       
    ListView2.listitems.Clear
    
    'If Option20.value = True And Option33.value = True Then
            
        sql = " SELECT Distinct PAYCode From VIEW_DOAYearReport Where 0 = 0 Order by PAYCode"
        
    'ElseIf Option21.value = True And Option33.value = True Then
        'sql = " SELECT Distinct PAYCode From VIEW_Discharge_Year_Report Where 0 = 0 Order by PAYCode"
    'End If
    
       
    rst2.Open sql, medixmasterconn, adOpenKeyset, adLockOptimistic
    
        Do Until rst2.EOF
                    
            With rst2
                If Len(rst2!PAYCode) Then
                
                Set Record = ListView2.listitems.Add(, , rst2!PAYCode)
                Payor = rst2!PAYCode
                
                If Option20.value = True And Option33.value = True Then
                    Call SearchAdmissionYearRecords(strAppend)
                ElseIf Option21.value = True And Option33.value = True Then
                    Call SearchDischargeYearRecords(strAppend)
                End If
                
                If Len(Total1) Then
                    Record.SubItems(1) = Total1
                End If
                                
                If Len(Total2) Then
                    Record.SubItems(2) = Total2
                End If
                
                If Len(Total3) Then
                    Record.SubItems(3) = Total3
                End If
                
                If Len(Total4) Then
                    Record.SubItems(4) = Total4
                End If
                                
                If Len(Total5) Then
                    Record.SubItems(5) = Total5
                End If
                
                If Len(Total6) Then
                    Record.SubItems(6) = Total6
                End If
                
                If Len(Total7) Then
                    Record.SubItems(7) = Total7
                End If
                
                If Len(Total8) Then
                    Record.SubItems(8) = Total8
                End If
                
                If Len(Total9) Then
                    Record.SubItems(9) = Total9
                End If
                
                If Len(Total10) Then
                    Record.SubItems(10) = Total10
                End If
                
                If Len(Total11) Then
                    Record.SubItems(11) = Total11
                End If
                
                If Len(Total12) Then
                    Record.SubItems(12) = Total12
                End If
                End If
            End With
            rst2.MoveNext
        Loop
    rst2.Close
    Set rst2 = Nothing
End Sub
'================ Admission / Discharge Year Listing ===========================

'==================== Search One Year Admission Records ===========================
Private Sub SearchAdmissionYearRecords(Optional strAppend As String)
Dim rst2 As New ADODB.Recordset
Dim Record As ListItem
Dim sql As String

    ListView3.listitems.Clear
    Call ClearBoolean
    
    'sql = " SELECT PAYCode, { fn MONTH(CASDateAdmitted) } AS Month, " & _
          " { fn YEAR(CASDateAdmitted) } AS Year, COUNT(PAYCode) AS Total " & _
          " FROM VIEW_AnalysisReportNew Where PAYCode = '" & Trim(Payor) & "'"
          
    sql = " Select * From VIEW_DOAYearReport Where PAYCode = '" & Trim(Payor) & "'"
    
    'sql2 = " GROUP BY { fn MONTH(CASDateAdmitted) }, { fn YEAR(CASDateAdmitted) }, PAYCode "
    
    If Len(Trim(strAppend)) Then
        sql = sql + strAppend '+ sql2
    End If
            
    rst2.Open sql, medixmasterconn, adOpenKeyset, adLockOptimistic
        
        Do Until rst2.EOF
        
            With rst2
                               
                Set Record = ListView3.listitems.Add(, , rst2!PAYCode)
    
                    If rst2!Month = "1" And one = False Then
                        If Len(rst2!Month) Then
                            Record.SubItems(1) = rst2!Total
                            Total1 = rst2!Total
                            one = True
                        End If
                
                    ElseIf rst2!Month = "2" And two = False Then
                        If Len(rst2!Month) Then
                            Record.SubItems(1) = rst2!Total
                            Total2 = rst2!Total
                            two = True
                        End If
                
                    ElseIf rst2!Month = "3" And three = False Then
                        If Len(rst2!Month) Then
                            Record.SubItems(1) = rst2!Total
                            Total3 = rst2!Total
                            three = True
                        End If
                
                    ElseIf rst2!Month = "4" And four = False Then
                        If Len(rst2!Month) Then
                            Record.SubItems(1) = rst2!Total
                            Total4 = rst2!Total
                            four = True
                        End If
                
                    ElseIf rst2!Month = "5" And five = False Then
                        If Len(rst2!Month) Then
                            Record.SubItems(1) = rst2!Total
                            Total5 = rst2!Total
                            five = True
                        End If
                
                    ElseIf rst2!Month = "6" And six = False Then
                        If Len(rst2!Month) Then
                            Record.SubItems(1) = rst2!Total
                            Total6 = rst2!Total
                            six = True
                        End If
                
                    ElseIf rst2!Month = "7" And seven = False Then
                        If Len(rst2!Month) Then
                            Record.SubItems(1) = rst2!Total
                            Total7 = rst2!Total
                            seven = True
                        End If
                
                    ElseIf rst2!Month = "8" And eight = False Then
                        If Len(rst2!Month) Then
                            Record.SubItems(1) = rst2!Total
                            Total8 = rst2!Total
                            eight = True
                        End If
                
                    ElseIf rst2!Month = "9" And nine = False Then
                        If Len(rst2!Month) Then
                            Record.SubItems(1) = rst2!Total
                            Total9 = rst2!Total
                            nine = True
                        End If
                
                    ElseIf rst2!Month = "10" And ten = False Then
                        If Len(rst2!Month) Then
                            Record.SubItems(1) = rst2!Total
                            Total10 = rst2!Total
                            ten = True
                        End If
                
                    ElseIf rst2!Month = "11" And eleven = False Then
                        If Len(rst2!Month) Then
                            Record.SubItems(1) = rst2!Total
                            Total11 = rst2!Total
                            eleven = True
                        End If
                
                    ElseIf rst2!Month = "12" And twelve = False Then
                        If Len(rst2!Month) Then
                            Record.SubItems(1) = rst2!Total
                            Total12 = rst2!Total
                            twelve = True
                        End If
                    End If
                
            End With
            rst2.MoveNext
        Loop
    rst2.Close
    Set rst2 = Nothing
End Sub
'==================== Search One Year Admission Records ===========================

'==================== Search One Year Discharge Records ===========================
Private Sub SearchDischargeYearRecords(Optional strAppend As String)
Dim rst2 As New ADODB.Recordset
Dim Record As ListItem
Dim sql As String

    ListView1.listitems.Clear
    Call ClearBoolean
    
    sql = " Select * From VIEW_AnalysisReportNew Where PAYCode = '" & Trim(Payor) & "'"

    
    If Len(Trim(strAppend)) Then
        sql = sql + strAppend
    End If
    
    rst2.Open sql, medixmasterconn, adOpenKeyset, adLockOptimistic
        
        Do Until rst2.EOF
        
            With rst2
                               
                Set Record = ListView3.listitems.Add(, , rst2!PAYCode)
    
                    If rst2!Month = "1" And one = False Then
                        If Len(rst2!Month) Then
                            Record.SubItems(1) = rst2!Total
                            Total1 = rst2!Total
                            one = True
                        End If
                
                    ElseIf rst2!Month = "2" And two = False Then
                        If Len(rst2!Month) Then
                            Record.SubItems(1) = rst2!Total
                            Total2 = rst2!Total
                            two = True
                        End If
                
                    ElseIf rst2!Month = "3" And three = False Then
                        If Len(rst2!Month) Then
                            Record.SubItems(1) = rst2!Total
                            Total3 = rst2!Total
                            three = True
                        End If
                
                    ElseIf rst2!Month = "4" And four = False Then
                        If Len(rst2!Month) Then
                            Record.SubItems(1) = rst2!Total
                            Total4 = rst2!Total
                            four = True
                        End If
                
                    ElseIf rst2!Month = "5" And five = False Then
                        If Len(rst2!Month) Then
                            Record.SubItems(1) = rst2!Total
                            Total5 = rst2!Total
                            five = True
                        End If
                
                    ElseIf rst2!Month = "6" And six = False Then
                        If Len(rst2!Month) Then
                            Record.SubItems(1) = rst2!Total
                            Total6 = rst2!Total
                            six = True
                        End If
                
                    ElseIf rst2!Month = "7" And seven = False Then
                        If Len(rst2!Month) Then
                            Record.SubItems(1) = rst2!Total
                            Total7 = rst2!Total
                            seven = True
                        End If
                
                    ElseIf rst2!Month = "8" And eight = False Then
                        If Len(rst2!Month) Then
                            Record.SubItems(1) = rst2!Total
                            Total8 = rst2!Total
                            eight = True
                        End If
                
                    ElseIf rst2!Month = "9" And nine = False Then
                        If Len(rst2!Month) Then
                            Record.SubItems(1) = rst2!Total
                            Total9 = rst2!Total
                            nine = True
                        End If
                
                    ElseIf rst2!Month = "10" And ten = False Then
                        If Len(rst2!Month) Then
                            Record.SubItems(1) = rst2!Total
                            Total10 = rst2!Total
                            ten = True
                        End If
                
                    ElseIf rst2!Month = "11" And eleven = False Then
                        If Len(rst2!Month) Then
                            Record.SubItems(1) = rst2!Total
                            Total11 = rst2!Total
                            eleven = True
                        End If
                
                    ElseIf rst2!Month = "12" And twelve = False Then
                        If Len(rst2!Month) Then
                            Record.SubItems(1) = rst2!Total
                            Total12 = rst2!Total
                            twelve = True
                        End If
                    End If
                
            End With
            rst2.MoveNext
        Loop
    rst2.Close
    Set rst2 = Nothing
End Sub
'==================== Search One Year Discharge Records ===========================

'Private Function SearchCriteria(Optional strAppend As String)
Private Function SearchRecords()
Dim sql As String
Dim strAppend As String

    
    If Len(Mid(cboHLTCode, 1, 1)) Then
        strAppend = strAppend + " AND HLTCode ='" & Trim(Mid(cboHLTCode, 1, 1)) & "'"
    End If

    If Len(Mid(CboInsType, 1, 1)) Then
        strAppend = strAppend & " AND INSCode = '" & Trim(Mid(CboInsType, 1, 1)) & "'"
    End If
       
     If Trim(Mid(CboMemType, 1, 1)) = "P" And Trim(cboSuppType) = "Both" Then
         strAppend = strAppend & " AND (MBMMemberType = 'P' Or MBMMemberType = 'Q') And MBMPatientCovID = '00' "
      ElseIf Trim(Mid(CboMemType, 1, 1)) = "P" And Trim(Mid(cboSuppType, 1, 1)) = "C" Then
         strAppend = strAppend & " AND MBMMemberType = 'P'  And MBMPatientCovID = '00'"
      ElseIf Trim(Mid(CboMemType, 1, 1)) = "P" And Trim(Mid(cboSuppType, 1, 1)) = "S" Then
         strAppend = strAppend & " AND MBMMemberType = 'Q' And MBMPatientCovID = '00'"
      ElseIf Trim(Mid(CboMemType, 1, 1)) = "S" And Trim(cboSuppType) = "Both" Then
         strAppend = strAppend & " AND (MBMCMemberType = 'S' OR MBMCMemberType = 'C')"
      ElseIf Trim(Mid(CboMemType, 1, 1)) = "S" And Trim(Mid(cboSuppType, 1, 1)) = "C" Then
         strAppend = strAppend & " AND MBMCMemberType = 'C'"
      ElseIf Trim(Mid(CboMemType, 1, 1)) = "S" And Trim(Mid(cboSuppType, 1, 1)) = "S" Then
         strAppend = strAppend & " AND MBMCMemberType = 'S'"
      ElseIf Trim(CboMemType) = "Both" And Trim(Mid(cboSuppType, 1, 1)) = "C" Then
         strAppend = strAppend & " AND (MBMMemberType = 'P' OR MBMCMemberType = 'C')"
      ElseIf Trim(CboMemType) = "Both" And Trim(Mid(cboSuppType, 1, 1)) = "S" Then
         strAppend = strAppend & " AND (MBMMemberType = 'Q' OR MBMCMemberType = 'S')"
      End If
    
    If Len(Mid(CboAdultChild, 1, 1)) Then
        strAppend = strAppend & " AND AdultChild = '" & Trim(Mid(CboAdultChild, 1, 1)) & "'"
    End If
    
    If Len(Mid(CboMarital, 1, 1)) Then
        strAppend = strAppend & " AND MBMMaritalStatus = '" & Trim(Mid(CboMarital, 1, 1)) & "'"
    End If
    
    If Len(Mid(CboRace, 1, 1)) Then
        strAppend = strAppend & " AND RACCode = '" & Trim(Mid(CboRace, 1, 1)) & "'"
    End If
    
    If Len(Mid(CboSex, 1, 1)) Then
        strAppend = strAppend & " AND Sex = '" & Trim(Mid(CboSex, 1, 1)) & "'"
    End If
    
    If Len(Mid(CboRelation, 1, 1)) Then
        strAppend = strAppend & " AND Relation = '" & Trim(Mid(CboRelation, 1, 1)) & "'"
    End If
    
    If Len(Mid(cboDataStatus, 1, 1)) Then
        strAppend = strAppend & " AND MBMDataStatus = '" & Trim(Mid(cboDataStatus, 1, 1)) & "'"
    End If
    
    If Len(Mid(CboClaimability, 1, 1)) Then
        strAppend = strAppend & " AND CASClaimability = '" & Trim(Mid(CboClaimability, 1, 1)) & "'"
    End If

    If Len(Mid(CboStatus, 1, 1)) Then
        strAppend = strAppend & " AND MBMStatus = '" & Trim(Mid(CboStatus, 1, 1)) & "'"
    End If

    If Len(Mid(CboCaseStatus, 1, 1)) Then
        strAppend = strAppend & " AND CASStatus = '" & Trim(Mid(CboCaseStatus, 1, 1)) & "'"
    End If
    
    If Len(Mid(CboWsStatus, 1, 1)) Then
        strAppend = strAppend & " AND ClaimStatus = '" & Trim(Mid(CboWsStatus, 1, 1)) & "'"
    End If
    
    If Len(Trim(Mid(CboOccupation, 1, 2))) Then
          strAppend = strAppend & "AND Occupation = '" & Trim(Mid(CboOccupation, 1, 2)) & "' "
    End If
    
    If Len(Mid(CboPreMode, 1, 1)) Then
        strAppend = strAppend & " AND MBMInstallment = '" & Trim(Mid(CboPreMode, 1, 1)) & "'"
    End If
    
    If Len(Trim(CboState)) Then
          strCriteria = strCriteria & "AND STACode= '" & Trim(CboState) & "' "
    End If
    
    If Len(Mid(cboRenewalInd, 1, 1)) Then
        strAppend = strAppend & " AND MBMRenewFlag = '" & Trim(Mid(cboRenewalInd, 1, 1)) & "'"
    End If
    
    If Len(Mid(cboTakeOver, 1, 1)) Then
        strAppend = strAppend & " AND MBMTakeOver = '" & Trim(Mid(cboTakeOver, 1, 1)) & "'"
    End If
    
    If Len(Trim(txtAgentCode)) Then
       strAppend = strAppend + " And  MBMAgentCode like '%" & ChkStr(txtAgentCode) & "%' "
    End If
    
    If Len(Trim(txtPolNo)) Then
       strAppend = strAppend + " And  MBMPolicyNo like '%" & ChkStr(txtPolNo) & "%' "
    End If
    
    If Len(Mid(cboPayorCode, 1, 2)) Then
        strAppend = strAppend & " AND PAYCode = '" & Trim(Mid(cboPayorCode, 1, 2)) & "'"
    End If
           
    If Len(Trim(CboGrpCom)) Then
        strAppend = strAppend & " AND MBMGroupCompany = '" & Trim(CboGrpCom) & "'"
    End If
    
    If Len(Trim(cboHOSName)) Then
        strAppend = strAppend + " And HOSName = '" & Trim(cboHOSName) & "'"
    End If
    
    If Len(Trim(cboHOSGrp)) Then
        strAppend = strAppend + " And HOSGRPName = '" & Trim(cboHOSGrp) & "'"
    End If
    
    If Len(Trim(txtDOCName)) Then
       strAppend = strAppend + " And  AttendDoctor like '%" & ChkStr(txtDOCName) & "%' "
    End If
    
    If Len(Trim(TxtPatientName)) Then
       strAppend = strAppend + " And  PatientName like '%" & ChkStr(TxtPatientName) & "%' "
    End If
        
    If (TxtDOA1) <> "__/__/____" And IsDate(TxtDOA1) = True _
        And (TxtDOA2) <> "__/__/____" And IsDate(TxtDOA2) = True Then
        sql = ""
        sql = " AND CasDateAdmitted >= '" & Format(Trim(TxtDOA1), "mm/dd/yyyy") & _
              "' And CasDateAdmitted <= '" & Format(Trim(TxtDOA2), "mm/dd/yyyy") & "'"
        strAppend = strAppend + sql
    End If
    
    If Len(Trim(TxtDOA1)) > 0 And IsDate(TxtDOA1) = True Then
        strAppend = strAppend + " And CasDateAdmitted >= '" & Format(TxtDOA1, "mm/dd/yyyy") & "'"
    End If
    
    If Len(Trim(TxtDOA2)) > 0 And IsDate(TxtDOA2) = True Then
        strAppend = strAppend + " And CasDateAdmitted <= '" & Format(TxtDOA2, "mm/dd/yyyy") & "'"
    End If
       
    If (txtDOD1) <> "__/__/____" And IsDate(txtDOD1) = True _
        And (txtDOD2) <> "__/__/____" And IsDate(txtDOD2) = True Then
        sql = ""
        sql = " AND CasDischargeDate >= '" & Format(Trim(txtDOD1), "mm/dd/yyyy") & _
              "' And CasDischargeDate <= '" & Format(Trim(txtDOD2), "mm/dd/yyyy") & "'"
        strAppend = strAppend + sql
    End If
    
    If Len(Trim(txtDOD1)) > 0 And IsDate(txtDOD1) = True Then
        strAppend = strAppend + " And CasDischargeDate >= '" & Format(txtDOD1, "mm/dd/yyyy") & "'"
    End If
    
    If Len(Trim(txtDOD2)) > 0 And IsDate(txtDOD2) = True Then
        strAppend = strAppend + " And CasDischargeDate <= '" & Format(txtDOD2, "mm/dd/yyyy") & "'"
    End If
    
    If (TxtDOR1) <> "__/__/____" And IsDate(TxtDOR1) = True _
        And (TxtDOR2) <> "__/__/____" And IsDate(TxtDOR2) = True Then
        sql = ""
        sql = " AND CaseRegDate >= '" & Format(Trim(TxtDOR1), "mm/dd/yyyy") & _
              "' And CaseRegDate <= '" & Format(Trim(TxtDOR2), "mm/dd/yyyy") & "'"
        strAppend = strAppend + sql
    End If
    
    If Len(Trim(TxtDOR1)) > 0 And IsDate(TxtDOR1) = True Then
        strAppend = strAppend + " And CaseRegDate >= '" & Format(TxtDOR1, "mm/dd/yyyy") & "'"
    End If
    
    If Len(Trim(TxtDOR2)) > 0 And IsDate(TxtDOR2) = True Then
        strAppend = strAppend + " And CaseRegDate <= '" & Format(TxtDOR2, "mm/dd/yyyy") & "'"
    End If
    
    If Len(Trim(txtPreDiag1)) Then
        strAppend = strAppend + " And FirstDiag >= '" & Trim(txtPreDiag1) & "'"
    End If
    
    If Len(Trim(txtPreDiag2)) Then
        strAppend = strAppend + " And FirstDiag <= '" & Trim(txtPreDiag2) & "'"
    End If
    
    If Len(Trim(txtFinalDiag1)) Then
        strAppend = strAppend + " And FinalDiag >= '" & Trim(txtFinalDiag1) & "'"
    End If
    
    If Len(Trim(txtFinalDiag2)) Then
        strAppend = strAppend + " And FinalDiag <= '" & Trim(txtFinalDiag2) & "'"
    End If
    
    If Len(Trim(txtRepudCat1)) Then
        strAppend = strAppend + " And CASREPCode1 >= '" & Trim(txtRepudCat1) & "'"
    End If
    
    If Len(Trim(txtRepudCat2)) Then
        strAppend = strAppend + " And CASREPCode1 <= '" & Trim(txtRepudCat2) & "'"
    End If
    
    If Len(Trim(txtRepudCode1)) Then
        strAppend = strAppend + " And RepCatCode >= '" & Trim(txtRepudCode1) & "'"
    End If
    
    If Len(Trim(txtRepudCode2)) Then
        strAppend = strAppend + " And RepCatCode <= '" & Trim(txtRepudCode2) & "'"
    End If
    
    If Len(Trim(TxtPlan1)) Then
        strAppend = strAppend & " AND PLNCode >= '" & Trim(TxtPlan1) & "'"
    End If
    
    If Len(Trim(TxtPlan2)) Then
        strAppend = strAppend & " AND PLNCode <= '" & Trim(TxtPlan2) & "'"
    End If
    
    If Len(Trim(TxtBordxNo1)) Then
        strAppend = strAppend & " AND BordxRefNo >= '" & Trim(TxtBordxNo1) & "'"
    End If
    
    If Len(Trim(TxtBordxNo2)) Then
        strAppend = strAppend & " AND BordxRefNo <= '" & Trim(TxtBordxNo2) & "'"
    End If
    
    If (TxtPayEffDate1) <> "__/__/____" And IsDate(TxtPayEffDate1) = True _
        And (TxtPayEffDate2) <> "__/__/____" And IsDate(TxtPayEffDate2) = True Then
        sql = ""
        sql = " AND MBMPayorEffDate >= '" & Format(Trim(TxtPayEffDate1), "mm/dd/yyyy") & _
              "' And MBMPayorEffDate <= '" & Format(Trim(TxtPayEffDate2), "mm/dd/yyyy") & "'"
        strAppend = strAppend + sql
    End If
    
    If Len(Trim(TxtPayEffDate1)) And IsDate(TxtPayEffDate1) = True Then
        strAppend = strAppend & " AND MBMPayorEffDate >= '" & Format(TxtPayEffDate1, "m/dd/yyy") & "'"
    End If
    
    If Len(Trim(TxtPayEffDate2)) And IsDate(TxtPayEffDate2) = True Then
        strAppend = strAppend & " AND MBMPayorEffDate <= '" & Format(TxtPayEffDate2, "m/dd/yyy") & "'"
    End If
    
    If (TxtPayExpDate1) <> "__/__/____" And IsDate(TxtPayExpDate1) = True _
        And (TxtPayExpDate2) <> "__/__/____" And IsDate(TxtPayExpDate2) = True Then
        sql = ""
        sql = " AND MBMPayorExpDate >= '" & Format(Trim(TxtPayExpDate1), "mm/dd/yyyy") & _
              "' And MBMPayorExpDate <= '" & Format(Trim(TxtPayExpDate2), "mm/dd/yyyy") & "'"
        strAppend = strAppend + sql
    End If
    
    If Len(Trim(TxtPayExpDate1)) And IsDate(TxtPayExpDate1) = True Then
        strAppend = strAppend & " AND MBMPayorExpDate >= '" & Format(TxtPayExpDate1, "m/dd/yyy") & "'"
    End If
    
    If Len(Trim(TxtPayExpDate2)) And IsDate(TxtPayExpDate2) = True Then
        strAppend = strAppend & " AND MBMPayorExpDate <= '" & Format(TxtPayExpDate2, "m/dd/yyy") & "'"
    End If
    
    If (txtDateJoin1) <> "__/__/____" And IsDate(txtDateJoin1) = True _
        And (txtDateJoin2) <> "__/__/____" And IsDate(txtDateJoin2) = True Then
        sql = ""
        sql = " AND MBMFirstJoinedDate >= '" & Format(Trim(TxtExpDate1), "mm/dd/yyyy") & _
              "' And MBMFirstJoinedDate <= '" & Format(Trim(TxtExpDate2), "mm/dd/yyyy") & "'"
        strAppend = strAppend + sql
    End If
    
    If Len(Trim(txtDateJoin1)) And IsDate(txtDateJoin1) = True Then
        strAppend = strAppend & " AND MBMFirstJoinedDate >= '" & Format(txtDateJoin1, "m/dd/yyy") & "'"
    End If
    
    If Len(Trim(txtDateJoin2)) And IsDate(txtDateJoin2) = True Then
        strAppend = strAppend & " AND MBMFirstJoinedDate <= '" & Format(txtDateJoin2, "m/dd/yyy") & "'"
    End If
    
    If (txtDOB1) <> "__/__/____" And IsDate(txtDOB1) = True _
        And (txtDOB2) <> "__/__/____" And IsDate(txtDOB2) = True Then
        sql = ""
        sql = " AND DOB >= '" & Format(Trim(txtDOB1), "mm/dd/yyyy") & _
              "' And DOB <= '" & Format(Trim(txtDOB2), "mm/dd/yyyy") & "'"
        strAppend = strAppend + sql
    End If
    
    If Len(Trim(txtDOB1)) And IsDate(txtDOB1) = True Then
        strAppend = strAppend & " AND DOB >= '" & Format(txtDOB1, "m/dd/yyy") & "'"
    End If
    
    If Len(Trim(txtDOB2)) And IsDate(txtDOB2) = True Then
        strAppend = strAppend & " AND DOB <= '" & Format(txtDOB2, "m/dd/yyy") & "'"
    End If
    
    If Len(Trim(txtPOLYear1)) Then
        strAppend = strAppend & " AND MBMPolicyYear >= '" & Trim(txtPOLYear1) & "'"
    End If
    
    If Len(Trim(txtPOLYear2)) Then
        strAppend = strAppend & " AND MBMPolicyYear <= '" & Trim(txtPOLYear2) & "'"
    End If
    
    If ChkHLT.value = 1 Then
        strAppend = strAppend & " AND HLTCode Is null "
    End If
    
    If ChkINS.value = 1 Then
        strAppend = strAppend & " AND INSCode Is null "
    End If
    
    If ChkMemType.value = 1 Then
        strAppend = strAppend & " AND MemberType Is null "
    End If
    
    If ChkAdultChild.value = 1 Then
        strAppend = strAppend & " AND AdultChild Is null "
    End If
    
    If ChkMarital.value = 1 Then
        strAppend = strAppend & " AND MBMMaritalStatus Is null "
    End If
    
    If ChkRace.value = 1 Then
        strAppend = strAppend & " AND RACCode Is null "
    End If
    
    If ChkSex.value = 1 Then
        strAppend = strAppend & " AND Sex Is null "
    End If
    
    If ChkRelation.value = 1 Then
        strAppend = strAppend & " AND Relation Is null "
    End If
    
    If ChkDataStatus.value = 1 Then
        strAppend = strAppend & " AND MBMDataStatus Is null "
    End If
    
    If ChkClaimability.value = 1 Then
        strAppend = strAppend & " AND CASClaimability Is null "
    End If
    
    If ChkPolStatus.value = 1 Then
        strAppend = strAppend & " AND MBMStatus Is null "
    End If
    
    If ChkCaseStatus.value = 1 Then
        strAppend = strAppend & " AND CASStatus Is null "
    End If
    
    If ChkWksStatus.value = 1 Then
        strAppend = strAppend & " AND ClaimStatus Is null "
    End If
    
    If ChkOccupation.value = 1 Then
        strAppend = strAppend & " AND Occupation Is null "
    End If
    
    If ChkPreMode.value = 1 Then
        strAppend = strAppend & " AND MBMInstallment Is null "
    End If
    
    If ChkState.value = 1 Then
        strAppend = strAppend & " AND STACode Is null "
    End If
    
    If ChkRenewal.value = 1 Then
        strAppend = strAppend & " AND MBMRenewFlag Is null "
    End If
    
    If ChkTakeOver.value = 1 Then
        strAppend = strAppend & " AND MBMTakeOver Is null "
    End If
    
    If ChkAgentCode.value = 1 Then
        strAppend = strAppend & " AND MBMAgentCode Is null "
    End If
    
    If ChkPolNo.value = 1 Then
        strAppend = strAppend & " AND MBMPolicyNo Is null "
    End If
    
    If ChkPayor.value = 1 Then
        strAppend = strAppend & " AND PAYCode Is null "
    End If
    
    If ChkGrpCom.value = 1 Then
        strAppend = strAppend & " AND MBMGroupCompany Is null "
    End If
    
    If ChkHOS.value = 1 Then
        strAppend = strAppend & " AND HOSName Is null "
    End If
    
    If ChkHOSGrp.value = 1 Then
        strAppend = strAppend & " AND HOSGRPName Is null "
    End If
    
    If ChkDoctor.value = 1 Then
        strAppend = strAppend & " AND AttendDoctor Is null "
    End If
            
    If ChkDOA.value = 1 Then
        strAppend = strAppend & " AND CASDateAdmitted Is null "
    End If
    
    If ChkDOD.value = 1 Then
        strAppend = strAppend & " AND CASDischargeDate Is null "
    End If
    
    If chkDOR.value = 1 Then
        strAppend = strAppend & " AND CaseRegDate Is null "
    End If
    
    If ChkFirst.value = 1 Then
        strAppend = strAppend & " AND FirstDiag Is null "
    End If
    
    If ChkFinal.value = 1 Then
        strAppend = strAppend & " AND FinalDiag Is null "
    End If
    
    If ChkRep.value = 1 Then
        strAppend = strAppend & " AND CASRepCode1 Is null "
    End If
    
    If ChkRepCat.value = 1 Then
        strAppend = strAppend & " AND RepCatCode Is null "
    End If
    
    If ChkPlan.value = 1 Then
        strAppend = strAppend & " AND PLNCode Is null "
    End If
    
    If ChkBordx.value = 1 Then
        strAppend = strAppend & " AND BordxRefNo Is null "
    End If
    
    If ChkEffDate.value = 1 Then
        strAppend = strAppend & " AND MBMPayorEffDate Is null "
    End If
    
    If chkExpDate.value = 1 Then
        strAppend = strAppend & " AND MBMPayorExpDate Is null "
    End If
    
    If ChkJoinDate.value = 1 Then
        strAppend = strAppend & " AND MBMFirstJoinedDate Is null "
    End If
    
    If ChkDOB.value = 1 Then
        strAppend = strAppend & " AND DOB Is null "
    End If
    
    If ChkPolYear.value = 1 Then
        strAppend = strAppend & " AND MBMPolicyYear Is null "
    End If
    
    If Len(Trim(TxtYear)) Then
        strAppend = strAppend + " AND Year = '" & Trim(TxtYear) & "'"
    End If
                
    ListView2.Height = 6720
    ListView2.Left = 0
    ListView2.Top = 480
    ListView2.Width = 11775
    Frame22.Visible = False
    Frame4.Visible = False
    Frame6.Visible = False
    
        
    If Option1.value = True Then
        Call ResultListing(strAppend)
    ElseIf Option2.value = True Then
        Call InsListing(strAppend)
    ElseIf Option3.value = True Then
        Call PlanListing(strAppend)
    ElseIf Option5.value = True Then
        Call SexAgeGroupListing(strAppend)
    ElseIf Option6.value = True Then
        Call SexListing(strAppend)
    ElseIf Option7.value = True Then
        Call RaceListing(strAppend)
    ElseIf Option8.value = True Then
        Call StateListing(strAppend)
    ElseIf Option9.value = True Then
        Call PayorListing(strAppend)
    ElseIf Option10.value = True And Option33.value = True Or Option10.value = True And Option29.value = True Then
        Call HOSTop20(strAppend)
    ElseIf Option10.value = True And Option34.value = True Then
        Call HOSListing(strAppend)
    ElseIf Option10.value = True And Option35.value = True Then
        Call DetailListing(strAppend)
    ElseIf Option15.value = True Then
        Call FinalDiagListing(strAppend)
    ElseIf Option20.value = True And Option35.value = True Then
        If TxtDOA2 <> TxtDOA1 Then
            MsgBox "Must key in same Date of Admitted", vbOKOnly + vbCritical, DialogBoxTitle
        ElseIf TxtDOA1 = "__/__/____" Or TxtDOA2 = "__/__/____" Then
            MsgBox "Date of Admitted can't be empty", vbOKOnly + vbCritical, DialogBoxTitle
        Else
            Call DailyAdmissionListing(strAppend)
        End If
   ElseIf Option20.value = True And Option34.value = True Then
        If Mid(TxtDOA1, 4, 2) <> Mid(TxtDOA2, 4, 2) Or Mid(TxtDOA1, 7, 4) <> Mid(TxtDOA2, 7, 4) Then
            MsgBox "Must key in same Month", vbOKOnly + vbCritical, DialogBoxTitle
        ElseIf TxtDOA1 = "__/__/____" Or TxtDOA2 = "__/__/____" Then
            MsgBox "Date of Admitted can't be empty", vbOKOnly + vbCritical, DialogBoxTitle
        Else
            Call MonthyReportListing(strAppend)
        End If
   ElseIf Option20.value = True And Option33.value = True Then
        If TxtYear = "" Then
            MsgBox "Please enter Year", vbOKOnly + vbCritical, DialogBoxTitle
        Else
            Call YearListing(strAppend)
        End If
   ElseIf Option36.value = True Then
        Call AGENTListing(strAppend)
   End If
End Function

Private Sub HOSTop20(Optional strAppend As String)
Dim rst2 As New ADODB.Recordset
Dim Record As ListItem
Dim sql As String
Dim sql2 As String
       
    ListView2.listitems.Clear
        
    sql = " SELECT TOP 20 HOSName, COUNT(HOSName) AS Total, " & _
          " SUM(CLMTotalActual) AS TotalActual, SUM(CLMTotalApproved) AS TotalApproved" & _
          " From VIEW_AnalysisReportNew Where 0 = 0 "
    
    If Option33.value = True Then
        sql2 = " GROUP BY HOSName ORDER BY Total DESC "
    ElseIf Option29.value = True Then
        sql2 = " GROUP BY HOSName ORDER BY TotalActual DESC "
    End If
    
    sql = sql + strAppend + sql2
    
    rst2.Open sql, medixmasterconn, adOpenKeyset, adLockOptimistic
    
        Do Until rst2.EOF
    
            With rst2
                Set Record = ListView2.listitems.Add(, , rst2!HOSName)
                'HOSCode = Trim(rst2!HOSName)
                
                'Call SearchHOSRecords(strAppend)
                
                If Len(rst2!Total) Then
                    Record.SubItems(1) = rst2!Total
                End If
                
                If Len(rst2!TotalActual) Then
                   Record.SubItems(2) = Format(rst2!TotalActual, "#,##0.00")
                End If
                
                If Len(rst2!TotalApproved) Then
                    Record.SubItems(3) = Format(rst2!TotalApproved, "#,##0.00")
                End If
                                
                Record.SubItems(4) = Format(Round(rst2!TotalActual / rst2!Total, 0), "#,##0.00")
                Record.SubItems(5) = Format(Round(rst2!TotalApproved / rst2!Total, 0), "#,##0.00")
                
            End With
            rst2.MoveNext
        Loop
    rst2.Close
    Set rst2 = Nothing
End Sub

Private Sub HOSListing(Optional strAppend As String)
Dim rst2 As New ADODB.Recordset
Dim Record As ListItem
Dim sql As String
Dim sql2 As String

    
    ListView2.listitems.Clear
            
    'sql = " SELECT HOSName, COUNT(HOSName) AS Total" & _
          " From VIEW_AnalysisReportNew Where 0 = 0 "
    
    sql = " SELECT HOSName, COUNT(HOSName) AS Total, " & _
          " SUM(CLMTotalActual) AS TotalActual, SUM(CLMTotalApproved) AS TotalApproved " & _
          " From VIEW_AnalysisReportNew Where 0 = 0 "
    
    sql2 = " GROUP BY HOSName Order by HOSName "
    
    If Len(Trim(strAppend)) Then
        sql = sql + strAppend + sql2
    End If
            
    rst2.Open sql, medixmasterconn, adOpenKeyset, adLockOptimistic
    
    
            Do Until rst2.EOF
                
            With rst2
                Set Record = ListView2.listitems.Add(, , rst2!HOSName)
                                
                If Len(rst2!Total) Then
                    Record.SubItems(1) = rst2!Total
                End If
                
                If Len(rst2!TotalActual) Then
                   Record.SubItems(2) = Format(rst2!TotalActual, "#,##0.00")
                End If
                
                If Len(rst2!TotalApproved) Then
                    Record.SubItems(3) = Format(rst2!TotalApproved, "#,##0.00")
                End If
                                
                Record.SubItems(4) = Format(Round(rst2!TotalActual / rst2!Total, 0), "#,##0.00")
                Record.SubItems(5) = Format(Round(rst2!TotalApproved / rst2!Total, 0), "#,##0.00")
                           
            End With
            rst2.MoveNext
        Loop
    rst2.Close
    Set rst2 = Nothing
End Sub

Private Sub DetailListing(Optional strAppend As String)
Dim rst2 As New ADODB.Recordset
Dim Record As ListItem
Dim sql As String
    
    ListView2.listitems.Clear
    Current = 0
    Total = 0
            
    sql = " SELECT * From VIEW_AnalysisReportNew Where 0 = 0 "
    
    If Len(Trim(strAppend)) Then
        sql = sql + strAppend
    End If
            
    rst2.Open sql, medixmasterconn, adOpenKeyset, adLockOptimistic
    
    If rst2.recordCount = 0 Then
        MsgBox "No record found!", vbOKOnly + vbCritical, DialogBoxTitle
    Else
    
    Do
    
            Do Until rst2.EOF
            
            Total = rst2.recordCount
            lblTotalRecord = Total
                
            With rst2
                Set Record = ListView2.listitems.Add(, , rst2!HOSName)
                                
                If Len(rst2!CASNumber) Then
                    Record.SubItems(1) = rst2!CASNumber
                End If
                                
                If Len(rst2!MBMNumber) Then
                    Record.SubItems(2) = rst2!MBMNumber
                End If
                
                If Len(rst2!MBMPatientCovID) Then
                    Record.SubItems(3) = rst2!MBMPatientCovID
                End If
                
                If Len(rst2!PatientName) Then
                    Record.SubItems(4) = rst2!PatientName
                End If
                
                If Len(rst2!CLMTotalActual) Then
                    Record.SubItems(5) = Format(rst2!CLMTotalActual, "#,##0.00")
                End If
                
                If Len(rst2!CLMTotalApproved) Then
                    Record.SubItems(6) = Format(rst2!CLMTotalApproved, "#,##0.00")
                End If
                
                If Len(rst2!CLMRecoveryPaidbyM) Then
                    Record.SubItems(7) = Format(rst2!CLMRecoveryPaidbyM, "#,##0.00")
                End If
                
                If Len(rst2!CASDateAdmitted) Then
                    Record.SubItems(8) = Format(rst2!CASDateAdmitted, "dd/mm/yyyy")
                End If
                
                If Len(rst2!CASDischargeDate) Then
                    Record.SubItems(9) = Format(rst2!CASDischargeDate, "dd/mm/yyyy")
                End If
                
                If Len(rst2!CASStatus) Then
                   Record.SubItems(10) = rst2!CASStatus
                End If
                
                If Len(rst2!CASClaimability) Then
                    Record.SubItems(11) = rst2!CASClaimability
                End If
                                
                If Len(rst2!FinalDiag) Then
                    Record.SubItems(12) = rst2!FinalDiag
                End If
                
                Current = Current + 1
                lblCurrent = Current
                           
            End With
            rst2.MoveNext
            DoEvents
        
            If intExit = 1 Then
                check = False
                Exit Do
            End If
            Loop
        Loop Until check = False
    End If
    intExit = 0
    rst2.Close
    Set rst2 = Nothing
End Sub

Private Sub SearchHOSRecords(Optional strAppend As String)
Dim rst2 As New ADODB.Recordset
Dim Record As ListItem
Dim sql As String
        
    Total = 0
    TotalActual = 0
    TotalApproved = 0
    TotalDuration = 0
    
    ListView3.listitems.Clear
            
    sql = " Select * From VIEW_AnalysisReportNew where HOSName = '" & Trim(HOSCode) & "'"
    
    If Len(Trim(strAppend)) Then
        sql = sql + strAppend
    End If
    
    
    rst2.Open sql, medixmasterconn, adOpenKeyset, adLockOptimistic
    
        
        Do Until rst2.EOF
                    
            With rst2
                
                Set Record = ListView3.listitems.Add(, , rst2!HOSName)
                
                If Len(rst2!CLMTotalActual) Then
                    Record.SubItems(1) = Format(rst2!CLMTotalActual, "#,##0.00")
                End If
                
                If Len(rst2!CLMTotalApproved) Then
                    Record.SubItems(2) = Format(rst2!CLMTotalApproved, "#,##0.00")
                End If
                                                               
                If Len(rst2!CLMTotalActual) Then
                    TotalActual = TotalActual + rst2!CLMTotalActual
                End If
                                
                If Len(rst2!CLMTotalApproved) Then
                    TotalApproved = TotalApproved + rst2!CLMTotalApproved
                End If
                
                Total = Total + 1
                
            End With
            rst2.MoveNext
        Loop
    
    rst2.Close
    Set rst2 = Nothing
End Sub

Private Sub ResultListing(Optional strAppend As String)
Dim rst2 As New ADODB.Recordset
Dim rst3 As New ADODB.Recordset
Dim Record As ListItem
Dim sql As String
Dim sql2 As String
Dim sql3 As String
Dim FirstICDCode As String
Dim FinalICDCode As String
Dim FirstDescription As String
Dim FinalDescription As String

        
    ListView2.listitems.Clear
    Current = 0
    Total = 0
            
    sql = " SELECT * From VIEW_AnalysisReportNew Where 0 = 0 "
    
    sql2 = " Order By CASNumber, ClaimVersion "
    
    If Len(Trim(strAppend)) Then
        sql = sql + strAppend + sql2
    End If
            
    rst2.Open sql, medixmasterconn, adOpenKeyset, adLockOptimistic
    
    If rst2.recordCount = 0 Then
        MsgBox "No record found!", vbOKOnly + vbCritical, DialogBoxTitle
    Else
    
    Do
              
       Do Until rst2.EOF
       
       Total = rst2.recordCount
       lblTotalRecord = Total
        
        FirstICDCode = ""
        FinalICDCode = ""
        FirstDescription = ""
        FinalDescription = ""
        Text1 = ""
        
        If Len(rst2!CASDischargeDate) And Len(rst2!CASDateAdmitted) Then
            Text1 = DateValue(rst2!CASDischargeDate) - DateValue((rst2!CASDateAdmitted))
        End If
        
        If Text1 = "0" Then
            Text1 = Text1 + 1
        End If
        
            With rst2
                    'If Len(rst2!CASNumber & "-" & rst2!claimversion) Then
                        Set Record = ListView2.listitems.Add(, , rst2!CASNumber & "-" & rst2!claimversion)
                                                  
                    If Len(rst2!MBMNumber) Then
                       Record.SubItems(1) = rst2!MBMNumber
                    End If
                    
                    If Len(rst2!MBMPolicyNo) Then
                       Record.SubItems(2) = rst2!MBMPolicyNo
                    End If
                                    
                    If Len(rst2!MBMPatientCovID) Then
                       Record.SubItems(3) = rst2!MBMPatientCovID
                    End If
                    
                    If Len(rst2!MBMName) Then
                       Record.SubItems(4) = rst2!MBMName
                    End If
                    
                    If Len(rst2!PatientName) Then
                       Record.SubItems(5) = rst2!PatientName
                    End If
                    
                    If Len(rst2!MBMSex) Then
                       Record.SubItems(6) = rst2!MBMSex
                    End If
                    
                    If Len(rst2!AgeCode) Then
                       Record.SubItems(7) = rst2!AgeCode
                    End If
                    
                    If Len(rst2!sex) Then
                       Record.SubItems(8) = rst2!sex
                    End If
                    
                    If Len(rst2!AgeBand) Then
                       Record.SubItems(9) = rst2!AgeBand
                    End If
                    
                    If Len(rst2!INSCode) Then
                       Record.SubItems(10) = rst2!INSCode
                    End If
                    
                    If Len(rst2!PLNCode) Then
                       Record.SubItems(11) = rst2!PLNCode
                    End If
                    
                    If Len(rst2!HOSName) Then
                       Record.SubItems(12) = rst2!HOSName
                    End If
                    
                    If Len(rst2!FirstDiag) Then
                       Record.SubItems(13) = rst2!FirstDiag
                       FirstICDCode = Record.SubItems(13)
                       sql3 = " SELECT ICDCode, ICDDescription From ICD Where ICDCode = '" & Trim(FirstICDCode) & "'"
                       rst3.Open sql3, medixmasterconn, adOpenKeyset, adLockOptimistic
                       
                       If rst3.recordCount = 0 Then
                            rst3.Close
                            Set rst3 = Nothing
                       Else
                            FirstDescription = rst3!ICDDescription
                            rst3.Close
                            Set rst3 = Nothing
                       End If
                    End If
                    
                    If Len(FirstDescription) Then
                       Record.SubItems(14) = Trim(FirstDescription)
                    End If
                                                            
                    If Len(rst2!FinalDiag) Then
                       Record.SubItems(15) = rst2!FinalDiag
                       FinalICDCode = Record.SubItems(15)
                       sql3 = " SELECT ICDCode, ICDDescription From ICD Where ICDCode = '" & Trim(FinalICDCode) & "'"
                       rst3.Open sql3, medixmasterconn, adOpenKeyset, adLockOptimistic
                       
                       If rst3.recordCount = 0 Then
                            rst3.Close
                            Set rst3 = Nothing
                       Else
                            FinalDescription = rst3!ICDDescription
                            rst3.Close
                            Set rst3 = Nothing
                       End If
                    End If
                    
                    If Len(FinalDescription) Then
                       Record.SubItems(16) = Trim(FinalDescription)
                    End If
                    
                    If Len(rst2!CLMTotalActual) Then
                       Record.SubItems(17) = Format(rst2!CLMTotalActual, "#,##0.00;(#,##0.00)")
                    End If
                    
                    If Len(rst2!CLMTotalApproved) Then
                       Record.SubItems(18) = Format(rst2!CLMTotalApproved, "#,##0.00;(#,##0.00)")
                    End If
                    
                    If Len(rst2!CLMRecoveryPaidbyM) Then
                       Record.SubItems(19) = Format(rst2!CLMRecoveryPaidbyM, "#,##0.00;(#,##0.00)")
                    End If
                    
                    If Len(rst2!CASRemark) Then
                       Record.SubItems(20) = Trim(rst2!CASRemark)
                    End If
                    
                    If Len(rst2!CASStatus) Then
                       Record.SubItems(21) = rst2!CASStatus
                    End If
                    
                    If Len(rst2!MBMStatus) Then
                       Record.SubItems(22) = rst2!MBMStatus
                    End If
                    
                    If Len(rst2!CASDateAdmitted) Then
                        Record.SubItems(23) = Format(rst2!CASDateAdmitted, "dd/mm/yyyy")
                    End If
                    
                    If Len(rst2!CASDischargeDate) Then
                        Record.SubItems(24) = Format(rst2!CASDischargeDate, "dd/mm/yyyy")
                    End If
                    
                    If Len(Text1) Then
                        Record.SubItems(25) = Text1
                    End If
                    
                    If Len(rst2!MBMPayorEffDate) Then
                        Record.SubItems(26) = Format(rst2!MBMPayorEffDate, "dd/mm/yyyy")
                    End If
                    
                    If Len(rst2!MBMPayorEXPDate) Then
                        Record.SubItems(27) = Format(rst2!MBMPayorEXPDate, "dd/mm/yyyy")
                    End If
                    
                    Current = Current + 1
                    lblCurrent = Current
            End With
            rst2.MoveNext
            DoEvents
        
            If intExit = 1 Then
                check = False
                Exit Do
            End If
            Loop
        Loop Until check = False
    intExit = 0
    End If
    rst2.Close
    Set rst2 = Nothing
    
End Sub


Private Sub DiagBySexListing(Optional strAppend As String)
Dim rst2 As New ADODB.Recordset
Dim Record As ListItem
Dim sql As String
       
    sql = ""
    lstSexList.listitems.Clear
            
    If Mid(cboSelection, 1, 1) = "0" Then
        sql = "TOP 10 "
    End If
    If Mid(cboSelection, 1, 1) = "1" Then
        sql = "TOP 20 "
    End If
    
    sql = "SELECT " & sql & "FinalDiag, COUNT(FinalDiag) AS Total " & _
          "From VIEW_AnalysisReportNew4 WHERE 0 = 0 " & strAppend & _
          " GROUP BY FinalDiag ORDER BY Total DESC "
    
    rst2.Open sql, medixmasterconn, adOpenKeyset, adLockOptimistic
    
        Do Until rst2.EOF
                    
            With rst2
                Set Record = lstSexList.listitems.Add(, , rst2!FinalDiag)
                DiagCode = Trim(rst2!FinalDiag)
                
                Call CountSexRecords
                
                If Len(Description) Then
                    Record.SubItems(1) = Trim(Description)
                End If
                
                If Len(rst2!Total) Then
                    Record.SubItems(2) = rst2!Total
                End If
                
                If Len(TotalSex1) Then
                   Record.SubItems(3) = TotalSex1
                End If
                If Len(TotalSex2) Then
                    Record.SubItems(4) = TotalSex2
                End If
                 If Len(TotalSex3) Then
                   Record.SubItems(5) = TotalSex3
                End If
                
                End With
            rst2.MoveNext
        Loop
    
    rst2.Close
    Set rst2 = Nothing
End Sub

Private Sub AGENTListing(Optional strAppend As String)
Dim rst2 As New ADODB.Recordset
Dim rst3 As New ADODB.Recordset
Dim Record As ListItem
Dim sql As String
Dim sql2 As String
Dim sql3 As String
Dim FirstICDCode As String
Dim FinalICDCode As String
Dim FirstDescription As String
Dim FinalDescription As String

        
    ListView2.listitems.Clear
    Current = 0
    Total = 0
            
    sql = " SELECT * From VIEW_AnalysisReportNew Where 0 = 0 "
    
    sql2 = " Order By MBMAgentCode "
    
    If Len(Trim(strAppend)) Then
        sql = sql + strAppend + sql2
    End If
            
    rst2.Open sql, medixmasterconn, adOpenKeyset, adLockOptimistic
    
    If rst2.recordCount = 0 Then
        MsgBox "No record found!", vbOKOnly + vbCritical, DialogBoxTitle
    Else
    
    Do
              
       Do Until rst2.EOF
       
       Total = rst2.recordCount
       lblTotalRecord = Total
                      
        FirstICDCode = ""
        FinalICDCode = ""
        FirstDescription = ""
        FinalDescription = ""
        Text1 = ""
        
        If Len(rst2!CASDischargeDate) And Len(rst2!CASDateAdmitted) Then
            Text1 = DateValue(rst2!CASDischargeDate) - DateValue((rst2!CASDateAdmitted))
        End If
        
        If Text1 = "0" Then
            Text1 = Text1 + 1
        End If
        
            With rst2
                    If Len(rst2!MBMAgentCode) Then
                        Set Record = ListView2.listitems.Add(, , rst2!MBMAgentCode)
                    Else
                        Set Record = ListView2.listitems.Add(, , "Empty")
                    End If
                    
                    If Len(rst2!CASNumber & "-" & rst2!claimversion) Then
                       Record.SubItems(1) = rst2!CASNumber & "-" & rst2!claimversion
                    End If
                    
                    If Len(rst2!MBMNumber) Then
                       Record.SubItems(2) = rst2!MBMNumber
                    End If
                                    
                    If Len(rst2!MBMPatientCovID) Then
                       Record.SubItems(3) = rst2!MBMPatientCovID
                    End If
                    
                    If Len(rst2!PatientName) Then
                       Record.SubItems(4) = rst2!PatientName
                    End If
                    
                    If Len(rst2!AttendDoctor) Then
                       Record.SubItems(5) = rst2!AttendDoctor
                    End If
                    
                    If Len(rst2!CLMTotalActual) Then
                       Record.SubItems(6) = Format(rst2!CLMTotalActual, "#,##0.00;(#,##0.00)")
                    End If
                    
                    If Len(rst2!CLMTotalApproved) Then
                       Record.SubItems(7) = Format(rst2!CLMTotalApproved, "#,##0.00;(#,##0.00)")
                    End If
                    
                    If Len(rst2!CLMRecoveryPaidbyM) Then
                       Record.SubItems(8) = Format(rst2!CLMRecoveryPaidbyM, "#,##0.00;(#,##0.00)")
                    End If
                    
                    If Len(rst2!FirstDiag) Then
                       Record.SubItems(9) = rst2!FirstDiag
                       FirstICDCode = Record.SubItems(9)
                       sql3 = " SELECT ICDCode, ICDDescription From ICD Where ICDCode = '" & Trim(FirstICDCode) & "'"
                       rst3.Open sql3, medixmasterconn, adOpenKeyset, adLockOptimistic
                       
                       If rst3.recordCount = 0 Then
                            rst3.Close
                            Set rst3 = Nothing
                       Else
                            FirstDescription = rst3!ICDDescription
                            rst3.Close
                            Set rst3 = Nothing
                       End If
                    End If
                    
                    If Len(FirstDescription) Then
                       Record.SubItems(10) = Trim(FirstDescription)
                    End If
                                                            
                    If Len(rst2!FinalDiag) Then
                       Record.SubItems(11) = rst2!FinalDiag
                       FinalICDCode = Record.SubItems(11)
                       sql3 = " SELECT ICDCode, ICDDescription From ICD Where ICDCode = '" & Trim(FinalICDCode) & "'"
                       rst3.Open sql3, medixmasterconn, adOpenKeyset, adLockOptimistic
                       
                       If rst3.recordCount = 0 Then
                            rst3.Close
                            Set rst3 = Nothing
                       Else
                            FinalDescription = rst3!ICDDescription
                            rst3.Close
                            Set rst3 = Nothing
                       End If
                    End If
                    
                    If Len(FinalDescription) Then
                       Record.SubItems(12) = Trim(FinalDescription)
                    End If
                                        
                    If Len(rst2!CASDateAdmitted) Then
                        Record.SubItems(13) = Format(rst2!CASDateAdmitted, "dd/mm/yyyy")
                    End If
                    
                    If Len(rst2!CASDischargeDate) Then
                        Record.SubItems(14) = Format(rst2!CASDischargeDate, "dd/mm/yyyy")
                    End If
                    
                    If Len(rst2!INSCode) Then
                       Record.SubItems(15) = rst2!INSCode
                    End If
                    
                    If Len(rst2!PLNCode) Then
                       Record.SubItems(16) = rst2!PLNCode
                    End If
                                        
                    If Len(rst2!CASStatus) Then
                       Record.SubItems(17) = rst2!CASStatus
                    End If
                    
                    Current = Current + 1
                    lblCurrent = Current
                    
            End With
            rst2.MoveNext
            DoEvents
        
            If intExit = 1 Then
                check = False
                Exit Do
            End If
            Loop
        Loop Until check = False
    intExit = 0
    End If
    rst2.Close
    Set rst2 = Nothing
End Sub

Private Sub RaceListing(Optional strAppend As String)
Dim rst2 As New ADODB.Recordset
Dim Record As ListItem
Dim sql As String
Dim sql2 As String

        
    ListView2.listitems.Clear
            
    sql = " SELECT RACCode, COUNT(*) as TotalCases, SUM(clmtotalactual) AS totalActual, " & _
          " SUM(clmtotalapproved) As TotalApproved From VIEW_AnalysisReportNew Where 0 = 0 "
    
    sql2 = " Group By RACCode "
    
    If Len(Trim(strAppend)) Then
        sql = sql + strAppend + sql2
    End If
    
               
    rst2.Open sql, medixmasterconn, adOpenKeyset, adLockOptimistic
    
        Do Until rst2.EOF
        
            With rst2
                If Len(rst2!RACCode) Then
                    Set Record = ListView2.listitems.Add(, , rst2!RACCode)
                    RaceCode = Trim(rst2!RACCode)
                Else
                    Set Record = ListView2.listitems.Add(, , "Empty")
                    RaceCode = Trim(rst2!RACCode)
                End If
                
                    Call CountRaceRecords(strAppend)
                               
                    If Len(rst2!TotalCases) Then
                        Record.SubItems(1) = rst2!TotalCases
                    End If
                    
                    If Len(TotalDuration) Then
                        Record.SubItems(2) = TotalDuration
                    End If
                    
                    If Len(rst2!TotalActual) Then
                        Record.SubItems(3) = Format(rst2!TotalActual, "#,##0.00")
                    End If
                    
                    If Len(rst2!TotalApproved) Then
                        Record.SubItems(4) = Format(rst2!TotalApproved, "#,##0.00")
                    End If
            End With
            rst2.MoveNext
       Loop
    rst2.Close
    Set rst2 = Nothing
End Sub


Private Sub CountRaceRecords(Optional strAppend As String)
Dim rst2 As New ADODB.Recordset
Dim Record As ListItem
Dim sql As String
        
    
    TotalDuration = 0
    
    ListView3.listitems.Clear
                
    If Len(RaceCode) Or RaceCode = "" Then
        sql = " Select * From VIEW_AnalysisReportNew Where RACCode  = '" & Trim(RaceCode) & "'"
    Else
        sql = " Select * From VIEW_AnalysisReportNew Where RACCode IS NULL "
    End If
    
        
    If Len(Trim(strAppend)) Then
        sql = sql + strAppend
    End If
            
    rst2.Open sql, medixmasterconn, adOpenKeyset, adLockOptimistic
    
        Do Until rst2.EOF
        TxtDay = ""
            If Len(rst2!CASDischargeDate) And Len(rst2!CASDateAdmitted) Then
                TxtDay = DateValue(rst2!CASDischargeDate) - DateValue((rst2!CASDateAdmitted))
            End If
            
            If TxtDay = "0" Then
                TxtDay = TxtDay + 1
            End If
            
            With rst2
                If Len(rst2!CASDischargeDate) Then
                    Set Record = ListView3.listitems.Add(, , rst2!CASDischargeDate)
                End If
                If Len(rst2!CASDateAdmitted) Then
                   Record.SubItems(1) = Format(rst2!CASDateAdmitted, "dd/mm/yyyy")
                End If
                If Len(TxtDay) Then
                   Record.SubItems(2) = TxtDay
                End If
                
                If Len(TxtDay) Then
                    TotalDuration = TotalDuration + CDbl(TxtDay)
                End If
            End With
            rst2.MoveNext
        Loop
    rst2.Close
    Set rst2 = Nothing
End Sub

Private Sub InsListing(Optional strAppend As String)
Dim rst2 As New ADODB.Recordset
Dim Record As ListItem
Dim sql As String
Dim sql2 As String

       
    ListView2.listitems.Clear
        
    sql = " SELECT inscode, COUNT(*) as TotalCases, SUM(clmtotalactual) AS totalActual, " & _
          " SUM(clmtotalapproved) As TotalApproved From VIEW_AnalysisReportNew Where 0 = 0 "
    
    sql2 = " Group By INSCode "
    
    If Len(Trim(strAppend)) Then
        sql = sql + strAppend + sql2
    End If
            
    rst2.Open sql, medixmasterconn, adOpenKeyset, adLockOptimistic
    
        Do Until rst2.EOF
        
            With rst2
                If Len(rst2!INSCode) Then
                    Set Record = ListView2.listitems.Add(, , rst2!INSCode)
                    INSCode = Trim(rst2!INSCode)
                Else
                    Set Record = ListView2.listitems.Add(, , "Empty")
                    INSCode = Trim(rst2!INSCode)
                End If
                
                    Call CountINSRecords(strAppend)
                               
                    If Len(rst2!TotalCases) Then
                        Record.SubItems(1) = rst2!TotalCases
                    End If
                    
                    If Len(TotalDuration) Then
                        Record.SubItems(2) = TotalDuration
                    End If
                    
                    If Len(rst2!TotalActual) Then
                        Record.SubItems(3) = Format(rst2!TotalActual, "#,##0.00")
                    End If
                    
                    If Len(rst2!TotalApproved) Then
                        Record.SubItems(4) = Format(rst2!TotalApproved, "#,##0.00")
                    End If
            End With
            rst2.MoveNext
       Loop
    rst2.Close
    Set rst2 = Nothing
End Sub

Private Sub CountINSRecords(Optional strAppend As String)
Dim rst2 As New ADODB.Recordset
Dim Record As ListItem
Dim sql As String
        
    
    TotalDuration = 0
        
    ListView3.listitems.Clear
                
    If Len(INSCode) Or INSCode = "" Then
        sql = " Select * From VIEW_AnalysisReportNew Where INSCode  = '" & Trim(INSCode) & "'"
    Else
        sql = " Select * From VIEW_AnalysisReportNew Where INSCode IS NULL "
    End If
    
        
    If Len(Trim(strAppend)) Then
        sql = sql + strAppend
    End If
            
    rst2.Open sql, medixmasterconn, adOpenKeyset, adLockOptimistic
    
        Do Until rst2.EOF
        TxtDay = ""
            If Len(rst2!CASDischargeDate) And Len(rst2!CASDateAdmitted) Then
                TxtDay = DateValue(rst2!CASDischargeDate) - DateValue((rst2!CASDateAdmitted))
            End If
            
            If TxtDay = "0" Then
                TxtDay = TxtDay + 1
            End If
            
            With rst2
                If Len(rst2!CASDischargeDate) Then
                    Set Record = ListView3.listitems.Add(, , rst2!CASDischargeDate)
                End If
                If Len(rst2!CASDateAdmitted) Then
                   Record.SubItems(1) = Format(rst2!CASDateAdmitted, "dd/mm/yyyy")
                End If
                If Len(TxtDay) Then
                   Record.SubItems(2) = TxtDay
                End If
                
                If Len(TxtDay) Then
                    TotalDuration = TotalDuration + CDbl(TxtDay)
                End If
                
            End With
            rst2.MoveNext
        Loop
    rst2.Close
    Set rst2 = Nothing
End Sub

Private Sub PlanListing(Optional strAppend As String)
Dim rst2 As New ADODB.Recordset
Dim Record As ListItem
Dim sql As String
Dim sql2 As String

       
    ListView2.listitems.Clear
        
    sql = " SELECT PLNCode, COUNT(*) as TotalCases, SUM(clmtotalactual) AS totalActual, " & _
          " SUM(clmtotalapproved) As TotalApproved From VIEW_AnalysisReportNew Where 0 = 0 "
    
    sql2 = " Group By PLNCode "
    
    If Len(Trim(strAppend)) Then
        sql = sql + strAppend + sql2
    End If
            
    rst2.Open sql, medixmasterconn, adOpenKeyset, adLockOptimistic
    
        Do Until rst2.EOF
        
            With rst2
                If Len(rst2!PLNCode) Then
                    Set Record = ListView2.listitems.Add(, , rst2!PLNCode)
                    PlanCode = Trim(rst2!PLNCode)
                Else
                    Set Record = ListView2.listitems.Add(, , "Empty")
                    PlanCode = Trim(rst2!PLNCode)
                End If
                
                    Call CountPlanRecords(strAppend)
                               
                    If Len(rst2!TotalCases) Then
                        Record.SubItems(1) = rst2!TotalCases
                    End If
                    
                    If Len(TotalDuration) Then
                        Record.SubItems(2) = TotalDuration
                    End If
                    
                    If Len(rst2!TotalActual) Then
                        Record.SubItems(3) = Format(rst2!TotalActual, "#,##0.00")
                    End If
                    
                    If Len(rst2!TotalApproved) Then
                        Record.SubItems(4) = Format(rst2!TotalApproved, "#,##0.00")
                    End If
            End With
            rst2.MoveNext
       Loop
    rst2.Close
    Set rst2 = Nothing
End Sub

Private Sub CountPlanRecords(Optional strAppend As String)
Dim rst2 As New ADODB.Recordset
Dim Record As ListItem
Dim sql As String
        
    
    TotalDuration = 0
        
    ListView3.listitems.Clear
                
    If Len(PlanCode) Or PlanCode = "" Then
        sql = " Select * From VIEW_AnalysisReportNew Where PLNCode  = '" & Trim(PlanCode) & "'"
    Else
        sql = " Select * From VIEW_AnalysisReportNew Where PLNCode IS NULL "
    End If
    
        
    If Len(Trim(strAppend)) Then
        sql = sql + strAppend
    End If
            
    rst2.Open sql, medixmasterconn, adOpenKeyset, adLockOptimistic
    
        Do Until rst2.EOF
        TxtDay = ""
            If Len(rst2!CASDischargeDate) And Len(rst2!CASDateAdmitted) Then
                TxtDay = DateValue(rst2!CASDischargeDate) - DateValue((rst2!CASDateAdmitted))
            End If
            
            If TxtDay = "0" Then
                TxtDay = TxtDay + 1
            End If
            
            With rst2
                If Len(rst2!CASDischargeDate) Then
                    Set Record = ListView3.listitems.Add(, , rst2!CASDischargeDate)
                End If
                If Len(rst2!CASDateAdmitted) Then
                   Record.SubItems(1) = Format(rst2!CASDateAdmitted, "dd/mm/yyyy")
                End If
                If Len(TxtDay) Then
                   Record.SubItems(2) = TxtDay
                End If
                
                If Len(TxtDay) Then
                    TotalDuration = TotalDuration + CDbl(TxtDay)
                End If
                
            End With
            rst2.MoveNext
        Loop
    rst2.Close
    Set rst2 = Nothing
End Sub

Private Sub StateListing(Optional strAppend As String)
Dim rst2 As New ADODB.Recordset
Dim Record As ListItem
Dim sql As String
Dim sql2 As String

    ListView2.listitems.Clear
        
    sql = " SELECT STACode, COUNT(*) as TotalCases, SUM(clmtotalactual) AS totalActual, " & _
          " SUM(clmtotalapproved) As TotalApproved From VIEW_AnalysisReportNew Where 0 = 0 "
    
    sql2 = " Group By STACode Order By STACode"
    
    If Len(Trim(strAppend)) Then
        sql = sql + strAppend + sql2
    End If
            
    rst2.Open sql, medixmasterconn, adOpenKeyset, adLockOptimistic
    
        Do Until rst2.EOF
        
            With rst2
                If Len(rst2!STACode) Then
                    Set Record = ListView2.listitems.Add(, , rst2!STACode)
                    StateCode = Trim(rst2!STACode)
                Else
                    Set Record = ListView2.listitems.Add(, , "Empty")
                    StateCode = Trim(rst2!STACode)
                End If
                
                    Call CountStateRecords(strAppend)
                               
                    If Len(rst2!TotalCases) Then
                        Record.SubItems(1) = rst2!TotalCases
                    End If
                    
                    If Len(TotalDuration) Then
                        Record.SubItems(2) = TotalDuration
                    End If
                    
                    If Len(rst2!TotalActual) Then
                        Record.SubItems(3) = Format(rst2!TotalActual, "#,##0.00")
                    End If
                    
                    If Len(rst2!TotalApproved) Then
                        Record.SubItems(4) = Format(rst2!TotalApproved, "#,##0.00")
                    End If
            End With
            rst2.MoveNext
       Loop
    rst2.Close
    Set rst2 = Nothing
End Sub

Private Sub CountStateRecords(Optional strAppend As String)
Dim rst2 As New ADODB.Recordset
Dim Record As ListItem
Dim sql As String
        
    
    TotalDuration = 0
        
    ListView3.listitems.Clear
                
    If Len(StateCode) Or StateCode = "" Then
        sql = " Select * From VIEW_AnalysisReportNew Where STACode  = '" & Trim(StateCode) & "'"
    Else
        sql = " Select * From VIEW_AnalysisReportNew Where STACode IS NULL "
    End If
    
        
    If Len(Trim(strAppend)) Then
        sql = sql + strAppend
    End If
            
    rst2.Open sql, medixmasterconn, adOpenKeyset, adLockOptimistic
    
        Do Until rst2.EOF
        TxtDay = ""
            If Len(rst2!CASDischargeDate) And Len(rst2!CASDateAdmitted) Then
                TxtDay = DateValue(rst2!CASDischargeDate) - DateValue((rst2!CASDateAdmitted))
            End If
            
            If TxtDay = "0" Then
                TxtDay = TxtDay + 1
            End If
            
            With rst2
                If Len(rst2!CASDischargeDate) Then
                    Set Record = ListView3.listitems.Add(, , rst2!CASDischargeDate)
                End If
                If Len(rst2!CASDateAdmitted) Then
                   Record.SubItems(1) = Format(rst2!CASDateAdmitted, "dd/mm/yyyy")
                End If
                If Len(TxtDay) Then
                   Record.SubItems(2) = TxtDay
                End If
                
                If Len(TxtDay) Then
                    TotalDuration = TotalDuration + CDbl(TxtDay)
                End If
                
            End With
            rst2.MoveNext
        Loop
    rst2.Close
    Set rst2 = Nothing
End Sub

Private Sub PayorListing(Optional strAppend As String)
Dim rst2 As New ADODB.Recordset
Dim Record As ListItem
Dim sql As String
Dim sql2 As String

    ListView2.listitems.Clear
        
    sql = " SELECT PAYCode, COUNT(*) as TotalCases, SUM(clmtotalactual) AS totalActual, " & _
          " SUM(clmtotalapproved) As TotalApproved From VIEW_AnalysisReportNew Where 0 = 0 "
    
    sql2 = " Group By PAYCode Order By PAYCode"
    
    If Len(Trim(strAppend)) Then
        sql = sql + strAppend + sql2
    End If
            
    rst2.Open sql, medixmasterconn, adOpenKeyset, adLockOptimistic
    
        Do Until rst2.EOF
        
            With rst2
                If Len(rst2!PAYCode) Then
                    Set Record = ListView2.listitems.Add(, , rst2!PAYCode)
                    Payor = Trim(rst2!PAYCode)
                Else
                    Set Record = ListView2.listitems.Add(, , "Empty")
                    Payor = Trim(rst2!PAYCode)
                End If
                
                    Call CountPayorRecords(strAppend)
                               
                    If Len(rst2!TotalCases) Then
                        Record.SubItems(1) = rst2!TotalCases
                    End If
                    
                    If Len(TotalDuration) Then
                        Record.SubItems(2) = TotalDuration
                    End If
                    
                    If Len(rst2!TotalActual) Then
                        Record.SubItems(3) = Format(rst2!TotalActual, "#,##0.00")
                    End If
                    
                    If Len(rst2!TotalApproved) Then
                        Record.SubItems(4) = Format(rst2!TotalApproved, "#,##0.00")
                    End If
            End With
            rst2.MoveNext
       Loop
    rst2.Close
    Set rst2 = Nothing
End Sub

Private Sub CountPayorRecords(Optional strAppend As String)
Dim rst2 As New ADODB.Recordset
Dim Record As ListItem
Dim sql As String
        
    
    TotalDuration = 0
        
    ListView3.listitems.Clear
                
    If Len(Payor) Or Payor = "" Then
        sql = " Select * From VIEW_AnalysisReportNew Where PAYCode  = '" & Trim(Payor) & "'"
    Else
        sql = " Select * From VIEW_AnalysisReportNew Where PAYCode IS NULL "
    End If
    
        
    If Len(Trim(strAppend)) Then
        sql = sql + strAppend
    End If
            
    rst2.Open sql, medixmasterconn, adOpenKeyset, adLockOptimistic
    
        Do Until rst2.EOF
        TxtDay = ""
            If Len(rst2!CASDischargeDate) And Len(rst2!CASDateAdmitted) Then
                TxtDay = DateValue(rst2!CASDischargeDate) - DateValue((rst2!CASDateAdmitted))
            End If
            
            If TxtDay = "0" Then
                TxtDay = TxtDay + 1
            End If
            
            With rst2
                If Len(rst2!CASDischargeDate) Then
                    Set Record = ListView3.listitems.Add(, , rst2!CASDischargeDate)
                End If
                If Len(rst2!CASDateAdmitted) Then
                   Record.SubItems(1) = Format(rst2!CASDateAdmitted, "dd/mm/yyyy")
                End If
                If Len(TxtDay) Then
                   Record.SubItems(2) = TxtDay
                End If
                
                If Len(TxtDay) Then
                    TotalDuration = TotalDuration + CDbl(TxtDay)
                End If
                
            End With
            rst2.MoveNext
        Loop
    rst2.Close
    Set rst2 = Nothing
End Sub


Private Sub ListViewHeader31()
    Dim AA As Integer
    ListView2.ColumnHeaders(1).Text = "Diag. Code"
    ListView2.ColumnHeaders(2).Text = "Diag. Desc"
    ListView2.ColumnHeaders(3).Text = "Total Cases"
    ListView2.ColumnHeaders(4).Text = "Total Duration"
    ListView2.ColumnHeaders(5).Text = "Total Actual Amt"
    ListView2.ColumnHeaders(6).Text = "Total Approved Amt"
    For AA = 2 To 6
        ListView2.ColumnHeaders(AA).Alignment = lvwColumnRight
    Next AA
    For AA = 7 To 63
        ListView2.ColumnHeaders(AA).Text = ""
    Next AA
End Sub
