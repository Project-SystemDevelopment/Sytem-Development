VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form FrmDCUReport1 
   Caption         =   "DCU Reports"
   ClientHeight    =   6795
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   10680
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   6795
   ScaleWidth      =   10680
   WindowState     =   2  'Maximized
   Begin VB.Frame Frame1 
      Caption         =   "Claim / Case"
      Height          =   855
      Left            =   0
      TabIndex        =   0
      Top             =   240
      Width           =   11895
      Begin VB.ComboBox cboReportType 
         Height          =   315
         Left            =   3720
         TabIndex        =   2
         Top             =   360
         Width           =   3000
      End
      Begin VB.ComboBox cboClaimCase 
         Height          =   315
         Left            =   600
         TabIndex        =   1
         Top             =   360
         Width           =   1500
      End
      Begin VB.ComboBox cboReportSelection 
         Height          =   315
         Left            =   8160
         TabIndex        =   3
         Top             =   360
         Width           =   3000
      End
      Begin VB.Label Label10 
         Caption         =   "Report Selection"
         Height          =   255
         Left            =   7680
         TabIndex        =   121
         Top             =   0
         Width           =   1335
      End
      Begin VB.Label Label5 
         Caption         =   "Report Type"
         Height          =   255
         Left            =   3480
         TabIndex        =   120
         Top             =   0
         Width           =   975
      End
   End
   Begin VB.Frame Frame22 
      Caption         =   "Selection Criteria"
      Height          =   4515
      Left            =   0
      TabIndex        =   76
      Top             =   1080
      Width           =   11865
      Begin VB.CheckBox ChkGrpCom2 
         Caption         =   "Check1"
         Height          =   255
         Left            =   10080
         TabIndex        =   122
         Top             =   550
         Width           =   255
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
         Left            =   6360
         Sorted          =   -1  'True
         TabIndex        =   19
         Top             =   270
         Width           =   3360
      End
      Begin VB.CheckBox ChkCaseStatus 
         Caption         =   "Check3"
         Height          =   255
         Left            =   4065
         TabIndex        =   56
         Top             =   1905
         Width           =   255
      End
      Begin VB.CheckBox ChkPolStatus 
         Caption         =   "Check3"
         Height          =   255
         Left            =   4065
         TabIndex        =   55
         Top             =   1620
         Width           =   255
      End
      Begin VB.CheckBox ChkINS 
         Caption         =   "Check3"
         Height          =   255
         Left            =   4065
         TabIndex        =   53
         Top             =   510
         Width           =   255
      End
      Begin VB.CheckBox ChkHLT 
         Caption         =   "Check3"
         Height          =   255
         Left            =   4065
         TabIndex        =   52
         Top             =   255
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
         Left            =   4065
         TabIndex        =   58
         Top             =   2415
         Width           =   255
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
         Left            =   4065
         TabIndex        =   54
         Top             =   1320
         Width           =   255
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
         Left            =   4065
         TabIndex        =   57
         Top             =   2175
         Width           =   255
      End
      Begin VB.CheckBox ChkFinal 
         Caption         =   "Check3"
         Height          =   255
         Left            =   9800
         TabIndex        =   65
         Top             =   1710
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
         Left            =   9800
         TabIndex        =   66
         Top             =   1920
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
         Height          =   255
         Left            =   9800
         TabIndex        =   70
         Top             =   3480
         Width           =   255
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
         Height          =   315
         Left            =   9800
         TabIndex        =   67
         Top             =   2160
         Width           =   255
      End
      Begin VB.CheckBox ChkPolNo 
         Caption         =   "Check3"
         Height          =   255
         Left            =   4065
         TabIndex        =   59
         Top             =   2655
         Width           =   255
      End
      Begin VB.CheckBox ChkGrpCom 
         Caption         =   "Check3"
         Height          =   255
         Left            =   9800
         TabIndex        =   61
         Top             =   550
         Width           =   255
      End
      Begin VB.CheckBox ChkPayor 
         Caption         =   "Check3"
         Height          =   255
         Left            =   9800
         TabIndex        =   60
         Top             =   300
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
         Height          =   315
         Left            =   9800
         TabIndex        =   69
         Top             =   2640
         Width           =   255
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
         Left            =   9800
         TabIndex        =   64
         Top             =   1380
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
         Height          =   315
         Left            =   9800
         TabIndex        =   68
         Top             =   2400
         Width           =   255
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
         Left            =   9800
         TabIndex        =   63
         Top             =   1095
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
         Left            =   9800
         TabIndex        =   62
         Top             =   855
         Width           =   255
      End
      Begin VB.ComboBox cboHLTCode 
         Height          =   315
         Left            =   1965
         Sorted          =   -1  'True
         TabIndex        =   4
         Text            =   "S - Sihat"
         Top             =   210
         Width           =   2000
      End
      Begin VB.CheckBox ChkBordxDate 
         Caption         =   "Check1"
         Height          =   255
         Left            =   9800
         TabIndex        =   71
         Top             =   3720
         Width           =   255
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
         Left            =   1965
         Sorted          =   -1  'True
         TabIndex        =   5
         Top             =   495
         Width           =   2000
      End
      Begin VB.ComboBox CboMemType 
         Height          =   315
         Left            =   1965
         Style           =   2  'Dropdown List
         TabIndex        =   6
         Top             =   735
         Width           =   2000
      End
      Begin VB.ComboBox cboSuppType 
         Height          =   315
         Left            =   1965
         Style           =   2  'Dropdown List
         TabIndex        =   7
         Top             =   1020
         Width           =   2000
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
         Left            =   1965
         TabIndex        =   8
         Top             =   1290
         Width           =   2000
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
         Left            =   1965
         TabIndex        =   9
         Top             =   1575
         Width           =   2000
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
         Left            =   1965
         Sorted          =   -1  'True
         TabIndex        =   10
         Top             =   1860
         Width           =   2000
      End
      Begin VB.ComboBox CboClaimability 
         Height          =   315
         Left            =   1965
         Sorted          =   -1  'True
         TabIndex        =   11
         Top             =   2145
         Width           =   2000
      End
      Begin VB.TextBox txtAgentCode 
         Height          =   285
         Left            =   1965
         MaxLength       =   15
         TabIndex        =   12
         Top             =   2415
         Width           =   2000
      End
      Begin VB.TextBox txtPolNo 
         Height          =   285
         Left            =   1965
         MaxLength       =   25
         TabIndex        =   13
         Top             =   2655
         Width           =   2000
      End
      Begin VB.ComboBox CboCasePending 
         Height          =   315
         Left            =   1965
         TabIndex        =   14
         Top             =   2895
         Width           =   2000
      End
      Begin VB.ComboBox CboClmPending 
         Height          =   315
         Left            =   1965
         TabIndex        =   15
         Top             =   3170
         Width           =   2000
      End
      Begin VB.ComboBox CboGrpCom 
         Height          =   315
         Left            =   6360
         TabIndex        =   20
         Top             =   555
         Width           =   3360
      End
      Begin MSMask.MaskEdBox TxtDOA1 
         Height          =   315
         Left            =   6360
         TabIndex        =   21
         Top             =   825
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
         Left            =   8520
         TabIndex        =   22
         Top             =   825
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
         Left            =   6360
         TabIndex        =   23
         Top             =   1095
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
         Left            =   8520
         TabIndex        =   24
         Top             =   1095
         Width           =   1200
         _ExtentX        =   2117
         _ExtentY        =   556
         _Version        =   393216
         MaxLength       =   10
         Mask            =   "##/##/####"
         PromptChar      =   "_"
      End
      Begin VB.ComboBox CboStayType 
         Height          =   315
         Left            =   1965
         Sorted          =   -1  'True
         TabIndex        =   16
         Top             =   3450
         Width           =   2000
      End
      Begin VB.ComboBox CboCashType 
         Height          =   315
         Left            =   1965
         Sorted          =   -1  'True
         TabIndex        =   17
         Top             =   3720
         Width           =   2000
      End
      Begin VB.ComboBox CboModality 
         Height          =   315
         Left            =   1965
         Sorted          =   -1  'True
         TabIndex        =   18
         Top             =   4000
         Width           =   2000
      End
      Begin MSMask.MaskEdBox TxtDOR2 
         Height          =   315
         Left            =   8520
         TabIndex        =   26
         Top             =   1380
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
         Left            =   6360
         TabIndex        =   25
         Top             =   1380
         Width           =   1200
         _ExtentX        =   2117
         _ExtentY        =   556
         _Version        =   393216
         MaxLength       =   10
         Mask            =   "##/##/####"
         PromptChar      =   "_"
      End
      Begin VB.TextBox txtFinalDiag1 
         Height          =   285
         Left            =   6360
         MaxLength       =   3
         TabIndex        =   27
         Top             =   1670
         Width           =   1200
      End
      Begin VB.TextBox txtFinalDiag2 
         Height          =   285
         Left            =   8520
         MaxLength       =   3
         TabIndex        =   28
         Top             =   1670
         Width           =   1200
      End
      Begin VB.TextBox TxtPlan2 
         Height          =   315
         Left            =   8520
         MaxLength       =   4
         TabIndex        =   30
         Top             =   1920
         Width           =   1200
      End
      Begin VB.TextBox TxtPlan1 
         Height          =   315
         Left            =   6360
         MaxLength       =   4
         TabIndex        =   29
         Top             =   1920
         Width           =   1200
      End
      Begin MSMask.MaskEdBox TxtPayEffDate2 
         Height          =   315
         Left            =   8520
         TabIndex        =   32
         Top             =   2160
         Width           =   1200
         _ExtentX        =   2117
         _ExtentY        =   556
         _Version        =   393216
         MaxLength       =   10
         Mask            =   "##/##/####"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox TxtPayEffDate1 
         Height          =   315
         Left            =   6360
         TabIndex        =   31
         Top             =   2160
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
         Left            =   8520
         TabIndex        =   34
         Top             =   2445
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
         Left            =   6360
         TabIndex        =   33
         Top             =   2445
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
         Left            =   8520
         TabIndex        =   36
         Top             =   2730
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
         Left            =   6360
         TabIndex        =   35
         Top             =   2730
         Width           =   1200
         _ExtentX        =   2117
         _ExtentY        =   556
         _Version        =   393216
         MaxLength       =   10
         Mask            =   "##/##/####"
         PromptChar      =   "_"
      End
      Begin VB.TextBox TxtCaseNo2 
         Height          =   285
         Left            =   8520
         MaxLength       =   15
         TabIndex        =   38
         Top             =   3000
         Width           =   1200
      End
      Begin VB.TextBox TxtCaseNo1 
         Height          =   285
         Left            =   6360
         MaxLength       =   15
         TabIndex        =   37
         Top             =   3000
         Width           =   1200
      End
      Begin VB.TextBox TxtVersion2 
         Height          =   285
         Left            =   8520
         MaxLength       =   2
         TabIndex        =   40
         Top             =   3240
         Width           =   1200
      End
      Begin VB.TextBox TxtVersion1 
         Height          =   285
         Left            =   6360
         MaxLength       =   2
         TabIndex        =   39
         Top             =   3240
         Width           =   1200
      End
      Begin VB.TextBox TxtBordxNo2 
         Height          =   285
         Left            =   8520
         MaxLength       =   10
         TabIndex        =   42
         Top             =   3480
         Width           =   1200
      End
      Begin VB.TextBox TxtBordxNo1 
         Height          =   285
         Left            =   6360
         MaxLength       =   10
         TabIndex        =   41
         Top             =   3480
         Width           =   1200
      End
      Begin MSMask.MaskEdBox TxtBordxDate2 
         Height          =   315
         Left            =   8520
         TabIndex        =   44
         Top             =   3720
         Width           =   1200
         _ExtentX        =   2117
         _ExtentY        =   556
         _Version        =   393216
         MaxLength       =   10
         Mask            =   "##/##/####"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox TxtBordxDate1 
         Height          =   315
         Left            =   6360
         TabIndex        =   43
         Top             =   3720
         Width           =   1200
         _ExtentX        =   2117
         _ExtentY        =   556
         _Version        =   393216
         MaxLength       =   10
         Mask            =   "##/##/####"
         PromptChar      =   "_"
      End
      Begin VB.Label Label15 
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
         ForeColor       =   &H00FF0000&
         Height          =   255
         Left            =   10130
         TabIndex        =   125
         Top             =   120
         Width           =   255
      End
      Begin VB.Label Label14 
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
         Left            =   9850
         TabIndex        =   124
         Top             =   120
         Width           =   255
      End
      Begin VB.Label Label11 
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
         Left            =   4080
         TabIndex        =   123
         Top             =   120
         Width           =   255
      End
      Begin VB.Label Label1 
         Caption         =   "Modality Code"
         Height          =   240
         Left            =   840
         TabIndex        =   118
         Top             =   4050
         Width           =   1215
      End
      Begin VB.Label Label2 
         Caption         =   "Stay Type"
         Height          =   255
         Left            =   840
         TabIndex        =   117
         Top             =   3495
         Width           =   1095
      End
      Begin VB.Label Label3 
         Caption         =   "Cash Type"
         Height          =   240
         Left            =   840
         TabIndex        =   116
         Top             =   3780
         Width           =   975
      End
      Begin VB.Label Label38 
         Caption         =   "Payor"
         Height          =   225
         Left            =   5160
         TabIndex        =   77
         Top             =   345
         Width           =   975
      End
      Begin VB.Label Label61 
         Alignment       =   2  'Center
         Caption         =   "To"
         Height          =   255
         Left            =   7920
         TabIndex        =   115
         Top             =   3480
         Width           =   375
      End
      Begin VB.Label Label60 
         Caption         =   "Bordx No"
         Height          =   255
         Left            =   5160
         TabIndex        =   114
         Top             =   3525
         Width           =   855
      End
      Begin VB.Label Label55 
         Caption         =   "Agent Code"
         Height          =   255
         Left            =   840
         TabIndex        =   113
         Top             =   2460
         Width           =   1095
      End
      Begin VB.Label Label41 
         Alignment       =   2  'Center
         Caption         =   "To"
         Height          =   255
         Left            =   7920
         TabIndex        =   112
         Top             =   2760
         Width           =   375
      End
      Begin VB.Label Label40 
         Caption         =   "Date Join"
         Height          =   255
         Left            =   5160
         TabIndex        =   111
         Top             =   2760
         Width           =   855
      End
      Begin VB.Label Label34 
         Caption         =   "Clm Pending"
         Height          =   255
         Left            =   840
         TabIndex        =   110
         Top             =   3225
         Width           =   975
      End
      Begin VB.Label Label33 
         Alignment       =   2  'Center
         Caption         =   "To"
         Height          =   255
         Left            =   7920
         TabIndex        =   109
         Top             =   1770
         Width           =   375
      End
      Begin VB.Label Label13 
         Caption         =   "Case Pending"
         Height          =   255
         Left            =   840
         TabIndex        =   108
         Top             =   2970
         Width           =   1095
      End
      Begin VB.Label Label12 
         Caption         =   "Final Diag"
         Height          =   255
         Left            =   5160
         TabIndex        =   107
         Top             =   1695
         Width           =   855
      End
      Begin VB.Label Label6 
         Caption         =   "Grp Company"
         Height          =   255
         Left            =   5160
         TabIndex        =   106
         Top             =   615
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
         Left            =   840
         TabIndex        =   105
         Top             =   1905
         Width           =   945
      End
      Begin VB.Label Label53 
         Caption         =   "Claimability"
         Height          =   255
         Left            =   840
         TabIndex        =   104
         Top             =   2190
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
         Left            =   840
         TabIndex        =   103
         Top             =   1635
         Width           =   1020
      End
      Begin VB.Label Label51 
         Caption         =   "Plan Code"
         Height          =   255
         Left            =   5160
         TabIndex        =   102
         Top             =   1995
         Width           =   855
      End
      Begin VB.Label Label46 
         Alignment       =   2  'Center
         Caption         =   "To"
         Height          =   255
         Left            =   7920
         TabIndex        =   101
         Top             =   2040
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
         Left            =   840
         TabIndex        =   100
         Top             =   1335
         Width           =   1020
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
         Left            =   840
         TabIndex        =   99
         Top             =   525
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
         Left            =   840
         TabIndex        =   98
         Top             =   225
         Width           =   705
      End
      Begin VB.Label Label7 
         Caption         =   "Eff Date"
         Height          =   255
         Left            =   5160
         TabIndex        =   97
         ToolTipText     =   "Payor Effective Date"
         Top             =   2250
         Width           =   735
      End
      Begin VB.Label Label23 
         Caption         =   "Exp Date"
         Height          =   255
         Left            =   5160
         TabIndex        =   96
         ToolTipText     =   "Payor Expiry Date"
         Top             =   2505
         Width           =   855
      End
      Begin VB.Label Label21 
         Caption         =   "DORC"
         Height          =   255
         Left            =   5160
         TabIndex        =   95
         ToolTipText     =   "Date of Register"
         Top             =   1425
         Width           =   855
      End
      Begin VB.Label Label8 
         Alignment       =   2  'Center
         Caption         =   "To"
         Height          =   255
         Left            =   7920
         TabIndex        =   94
         Top             =   2280
         Width           =   375
      End
      Begin VB.Label Label9 
         Alignment       =   2  'Center
         Caption         =   "To"
         Height          =   255
         Left            =   7920
         TabIndex        =   93
         Top             =   2520
         Width           =   375
      End
      Begin VB.Label Label16 
         Alignment       =   2  'Center
         Caption         =   "To"
         Height          =   255
         Left            =   7920
         TabIndex        =   92
         Top             =   1440
         Width           =   375
      End
      Begin VB.Label Label47 
         Caption         =   "DOA"
         Height          =   195
         Left            =   5160
         TabIndex        =   91
         ToolTipText     =   "Date of Admission"
         Top             =   900
         Width           =   975
      End
      Begin VB.Label Label48 
         Caption         =   "DOD"
         Height          =   255
         Left            =   5160
         TabIndex        =   90
         ToolTipText     =   "Date of Discharged"
         Top             =   1155
         Width           =   975
      End
      Begin VB.Label Label49 
         Alignment       =   2  'Center
         Caption         =   "To"
         Height          =   255
         Left            =   7920
         TabIndex        =   89
         Top             =   960
         Width           =   375
      End
      Begin VB.Label Label50 
         Alignment       =   2  'Center
         Caption         =   "To"
         Height          =   255
         Left            =   7920
         TabIndex        =   88
         Top             =   1200
         Width           =   375
      End
      Begin VB.Label Label4 
         Caption         =   "Mem Type"
         Height          =   255
         Left            =   840
         TabIndex        =   87
         Top             =   795
         Width           =   855
      End
      Begin VB.Label Label19 
         Height          =   255
         Left            =   2520
         TabIndex        =   86
         Top             =   240
         Width           =   615
      End
      Begin VB.Label Label30 
         Caption         =   "Policy No."
         Height          =   255
         Left            =   840
         TabIndex        =   85
         Top             =   2730
         Width           =   975
      End
      Begin VB.Label Label31 
         Caption         =   "Supp. Type"
         Height          =   255
         Left            =   840
         TabIndex        =   84
         Top             =   1050
         Width           =   1095
      End
      Begin VB.Label Label66 
         Caption         =   "Claim No"
         Height          =   255
         Left            =   5160
         TabIndex        =   83
         Top             =   3030
         Width           =   1095
      End
      Begin VB.Label Label67 
         Alignment       =   2  'Center
         Caption         =   "To"
         Height          =   255
         Left            =   7920
         TabIndex        =   82
         Top             =   3000
         Width           =   375
      End
      Begin VB.Label Label70 
         Caption         =   "Bordx Date"
         Height          =   255
         Left            =   5160
         TabIndex        =   81
         Top             =   3795
         Width           =   975
      End
      Begin VB.Label Label71 
         Alignment       =   2  'Center
         Caption         =   "To"
         Height          =   255
         Left            =   7920
         TabIndex        =   80
         Top             =   3765
         Width           =   375
      End
      Begin VB.Label Label63 
         Caption         =   "Version"
         Height          =   255
         Left            =   5160
         TabIndex        =   79
         Top             =   3270
         Width           =   1095
      End
      Begin VB.Label Label81 
         Alignment       =   2  'Center
         Caption         =   "To"
         Height          =   255
         Left            =   7920
         TabIndex        =   78
         Top             =   3240
         Width           =   375
      End
   End
   Begin VB.Frame Frame2 
      Height          =   615
      Left            =   0
      TabIndex        =   72
      Top             =   7320
      Width           =   11775
      Begin VB.CommandButton CmdStop 
         Caption         =   "S&top"
         Height          =   375
         Left            =   6240
         TabIndex        =   46
         Top             =   160
         Width           =   855
      End
      Begin VB.CommandButton CmdView 
         Caption         =   "&View"
         Height          =   375
         Left            =   7920
         TabIndex        =   48
         Top             =   160
         Width           =   855
      End
      Begin VB.CommandButton CmdBack 
         Caption         =   "&Back"
         Height          =   375
         Left            =   7080
         TabIndex        =   47
         Top             =   160
         Width           =   855
      End
      Begin VB.CommandButton CmdPrint 
         Caption         =   "&Print to File"
         Height          =   375
         Left            =   8760
         TabIndex        =   49
         Top             =   160
         Width           =   975
      End
      Begin VB.CommandButton cmdExit 
         Cancel          =   -1  'True
         Caption         =   "&Exit"
         Height          =   375
         Left            =   10560
         TabIndex        =   51
         Top             =   160
         Width           =   855
      End
      Begin VB.CommandButton CmdClear 
         Caption         =   "&Clear"
         Height          =   375
         Left            =   9720
         TabIndex        =   50
         Top             =   160
         Width           =   855
      End
      Begin VB.CommandButton cmdSearch 
         Caption         =   "&Search"
         Default         =   -1  'True
         Height          =   375
         Left            =   5400
         TabIndex        =   45
         Top             =   160
         Width           =   855
      End
      Begin VB.Label Label64 
         Caption         =   "Record(s)"
         Height          =   255
         Left            =   1560
         TabIndex        =   74
         Top             =   240
         Width           =   855
      End
      Begin VB.Label lblCurrent 
         Alignment       =   1  'Right Justify
         BackColor       =   &H80000009&
         Height          =   255
         Left            =   240
         TabIndex        =   73
         Top             =   240
         Width           =   1215
      End
   End
   Begin MSComctlLib.ListView ListView2 
      CausesValidation=   0   'False
      Height          =   1695
      Left            =   0
      TabIndex        =   75
      Top             =   5640
      Width           =   11895
      _ExtentX        =   20981
      _ExtentY        =   2990
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   0   'False
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
         Object.Width           =   2469
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
   Begin VB.Label Label28 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00404040&
      Caption         =   "DCU Reports"
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
      TabIndex        =   119
      Top             =   0
      Width           =   11805
   End
End
Attribute VB_Name = "FrmDCUReport1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private DiagCode, plancode, ClaimNatureCode As String
Private AA As Integer
Private AgeCode, INSCode, Hospital, HOspitals As String
Private Description, Month, Month2, Year As String
Private SexCode, RaceCode, Payor As String
Private Total As String
Private StrAppend As String
Private intExit As Integer
Private Current As String
Private GrpCompany As String
Dim NoInv, NoInv2, NoInv3, NoInv4, NoInv5, TotalInv As Variant
Private m_blnDirection(1 To 63) As Boolean
Private opt35 As Boolean
Private opt34 As Boolean
Private opt33 As Boolean
Private opt29 As Boolean
Private opt37 As Boolean
Private opt30 As Boolean
Private opt31 As Boolean
Private Opt8 As Boolean
Private Opt11 As Boolean
Private opt12 As Boolean
Private opt13 As Boolean
Private Opt7 As Boolean
Private opt16 As Boolean

Private Sub cboReportSelection_KeyPress(KeyAscii As Integer)
Dim NewText As String
Dim ValidCount As Integer
Dim ValidValue As String
Dim i As Integer
    If KeyAscii >= 32 And KeyAscii <> 127 Then
    NewText = LCase(Left(cboReportSelection.Text, cboReportSelection.SelStart) + Chr(KeyAscii) + Mid(cboReportSelection.Text, cboReportSelection.SelStart + cboReportSelection.SelLength + 1))
 
    ValidCount = 0
    ValidValue = ""
 
    For i = 0 To cboReportSelection.ListCount - 1
 
        If NewText = LCase(Left(cboReportSelection.List(i), Len(NewText))) Then
                ValidCount = ValidCount + 1
                ValidValue = cboReportSelection.List(i)
        End If
        Next
        If ValidCount <= 1 Then KeyAscii = 0
            If ValidCount = 1 Then
                cboReportSelection.Text = ValidValue
                cboReportSelection.SelStart = 0
                cboReportSelection.SelLength = Len(ValidValue)
            End If
        End If
End Sub

Private Sub CboReportType_KeyPress(KeyAscii As Integer)
Dim NewText As String
Dim ValidCount As Integer
Dim ValidValue As String
Dim i As Integer
    If KeyAscii >= 32 And KeyAscii <> 127 Then
    NewText = LCase(Left(cboReportType.Text, cboReportType.SelStart) + Chr(KeyAscii) + Mid(cboReportType.Text, cboReportType.SelStart + cboReportType.SelLength + 1))
 
    ValidCount = 0
    ValidValue = ""
 
    For i = 0 To cboReportType.ListCount - 1
 
        If NewText = LCase(Left(cboReportType.List(i), Len(NewText))) Then
                ValidCount = ValidCount + 1
                ValidValue = cboReportType.List(i)
        End If
        Next
        If ValidCount <= 1 Then KeyAscii = 0
            If ValidCount = 1 Then
                cboReportType.Text = ValidValue
                cboReportType.SelStart = 0
                cboReportType.SelLength = Len(ValidValue)
            End If
        End If
End Sub

Private Sub cboClaimCase_KeyPress(KeyAscii As Integer)
Dim NewText As String
Dim ValidCount As Integer
Dim ValidValue As String
Dim i As Integer
    If KeyAscii >= 32 And KeyAscii <> 127 Then
    NewText = LCase(Left(cboClaimCase.Text, cboClaimCase.SelStart) + Chr(KeyAscii) + Mid(cboClaimCase.Text, cboClaimCase.SelStart + cboClaimCase.SelLength + 1))
 
    ValidCount = 0
    ValidValue = ""
 
    For i = 0 To cboReportType.ListCount - 1
 
        If NewText = LCase(Left(cboClaimCase.List(i), Len(NewText))) Then
                ValidCount = ValidCount + 1
                ValidValue = cboClaimCase.List(i)
        End If
        Next
        If ValidCount <= 1 Then KeyAscii = 0
            If ValidCount = 1 Then
                cboClaimCase.Text = ValidValue
                cboClaimCase.SelStart = 0
                cboClaimCase.SelLength = Len(ValidValue)
            End If
        End If
End Sub




Private Sub ChkGrpCom2_Click()
Dim StrSql As String
Dim PAYGRP As New ADODB.Recordset
Dim cmd As New ADODB.Command
Dim prm1 As New ADODB.Parameter
Dim StrsqlAppend As String

    If Len(Mid(cboPayorCode, 1, 2)) Then
        If ChkGrpCom2.Value = 1 Then
            CboGrpCom.Clear
          '  medixmasterconn.Open "File Name=" & BOSPath & "Conn\mediconnect.UDL"
           ' StrsqlAppend = StrsqlAppend & "And PAYCode = '" & Mid(cboPayorCode, 1, 2) & "' GROUP BY MBMGroupCompany, PAYCode"
            
            Set cmd.ActiveConnection = medixmasterconn
            cmd.CommandText = "SearchGrpCompanyPay"
            cmd.CommandType = adCmdStoredProc
            cmd.CommandTimeout = 15
            Set prm1 = cmd.CreateParameter("StrsqlAppend", adVarChar, adParamInput, 255, StrsqlAppend)
            cmd.Parameters.Append prm1
            PAYGRP.CursorLocation = adUseClient
            PAYGRP.CursorType = adOpenForwardOnly
            PAYGRP.CacheSize = 100
            Set PAYGRP = cmd.Execute
            
            Do Until PAYGRP.EOF
                If Len(PAYGRP!MBMGroupCompany) Then
                    CboGrpCom.AddItem Trim(PAYGRP!MBMGroupCompany)
                End If
                PAYGRP.MoveNext
            Loop
            PAYGRP.Close
            Set PAYGRP = Nothing
          '  medixmasterconn.Close
           ' Set medixmasterconn = Nothing
        Else
            CboGrpCom.Clear
        End If
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

Private Sub cboHLTCode_Change()
    If InStr(cboHLTCode.Text, "null") Then
       cboHLTCode.Enabled = False
    End If
    
    If Len(Trim(cboHLTCode)) Then
        'medixmasterconn.Open "File Name=" & BOSPath & "Conn\mediconnect.UDL"
        LoadPayor (cboHLTCode)
       ' medixmasterconn.Close
       ' Set medixmasterconn = Nothing
    End If
End Sub

Private Sub cboHLTCode_Click()
    If Len(Trim(cboHLTCode)) Then
     '   medixmasterconn.Open "File Name=" & BOSPath & "Conn\mediconnect.UDL"
        LoadPayor (cboHLTCode)
       ' medixmasterconn.Close
      '  Set medixmasterconn = Nothing
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
    PAY.Open sqlPAY, medixmasterconn, adOpenForwardOnly, adLockBatchOptimistic
    Do Until PAY.EOF
        cboPayorCode.AddItem ChkStr(PAY!PAYCode) & " - " & ChkStr(PAY!PAYCompanyName)
        PAY.MoveNext
    Loop
    
    PAY.Close
    Set PAY = Nothing
End Sub

Private Sub CmdBack_Click()
    ListView2.Height = 2655
    ListView2.Left = 0
    ListView2.Top = 5640
    ListView2.Width = 11775
    Frame22.Visible = True
    Frame1.Visible = True
End Sub

Private Sub CmdStop_Click()
    intExit = 1
    cmdExit.Enabled = True
    CmdClear.Enabled = True
    CmdPrint.Enabled = True
End Sub

Private Sub CmdView_Click()
    ListView2.Height = 6720
    ListView2.Left = 0
    ListView2.Top = 480
    ListView2.Width = 11775
    Frame22.Visible = False
    Frame1.Visible = False
End Sub

Private Sub ListViewHeader1()
    
If cboClaimCase.ListIndex = 0 Then
    ListView2.ColumnHeaders(1).Text = "Claim Nature"
    ListView2.ColumnHeaders(2).Text = "No of Claims"
    ListView2.ColumnHeaders(2).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(3).Text = "Total Duration"
    ListView2.ColumnHeaders(3).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(4).Text = "Total Actual Amt"
    ListView2.ColumnHeaders(4).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(5).Text = "Total Approved Amt"
    ListView2.ColumnHeaders(5).Alignment = lvwColumnRight
    
    For AA = 6 To 63
        ListView2.ColumnHeaders(AA).Text = ""
    Next AA
ElseIf cboClaimCase.ListIndex = 1 Then
    ListView2.ColumnHeaders(1).Text = "Claim Nature"
    ListView2.ColumnHeaders(2).Text = "No of Cases"
    ListView2.ColumnHeaders(2).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(3).Text = "Total Duration"
    ListView2.ColumnHeaders(3).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(4).Text = "Total Actual Amt"
    ListView2.ColumnHeaders(4).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(5).Text = "Total Approved Amt"
    ListView2.ColumnHeaders(5).Alignment = lvwColumnRight
    
    For AA = 6 To 63
        ListView2.ColumnHeaders(AA).Text = ""
    Next AA
End If
End Sub

Private Sub ListViewHeader15()

    If opt35 = True And cboClaimCase.ListIndex = 0 Then
        ListView2.ColumnHeaders(1).Text = "Payor"
        ListView2.ColumnHeaders(2).Text = "Claim No"
        ListView2.ColumnHeaders(2).Alignment = lvwColumnLeft
        ListView2.ColumnHeaders(3).Text = "BordxNo"
        ListView2.ColumnHeaders(3).Alignment = lvwColumnLeft
        ListView2.ColumnHeaders(4).Text = "Bordx Date"
        ListView2.ColumnHeaders(4).Alignment = lvwColumnCenter
        ListView2.ColumnHeaders(5).Text = "Member No"
        ListView2.ColumnHeaders(5).Alignment = lvwColumnLeft
        ListView2.ColumnHeaders(6).Text = "Policy No"
        ListView2.ColumnHeaders(6).Alignment = lvwColumnLeft
        ListView2.ColumnHeaders(7).Text = "Covered ID"
        ListView2.ColumnHeaders(7).Alignment = lvwColumnCenter
        ListView2.ColumnHeaders(8).Text = "Prin. Name"
        ListView2.ColumnHeaders(8).Alignment = lvwColumnLeft
        ListView2.ColumnHeaders(9).Text = "Patient Name"
        ListView2.ColumnHeaders(9).Alignment = lvwColumnLeft
        ListView2.ColumnHeaders(10).Text = "Prin. IC No"
        ListView2.ColumnHeaders(10).Alignment = lvwColumnLeft
        ListView2.ColumnHeaders(11).Text = "Prin. Sex"
        ListView2.ColumnHeaders(11).Alignment = lvwColumnCenter
        ListView2.ColumnHeaders(12).Text = "Prin. Ageband"
        ListView2.ColumnHeaders(12).Alignment = lvwColumnCenter
        ListView2.ColumnHeaders(13).Text = "Prin. DOB"
        ListView2.ColumnHeaders(13).Alignment = lvwColumnCenter
        ListView2.ColumnHeaders(14).Text = "Patient IC No"
        ListView2.ColumnHeaders(14).Alignment = lvwColumnLeft
        ListView2.ColumnHeaders(15).Text = "Patient Sex"
        ListView2.ColumnHeaders(15).Alignment = lvwColumnCenter
        ListView2.ColumnHeaders(16).Text = "Patient Ageband"
        ListView2.ColumnHeaders(16).Alignment = lvwColumnCenter
        ListView2.ColumnHeaders(17).Text = "Patient DOB"
        ListView2.ColumnHeaders(17).Alignment = lvwColumnCenter
        ListView2.ColumnHeaders(18).Text = "Insured Type"
        ListView2.ColumnHeaders(18).Alignment = lvwColumnCenter
        ListView2.ColumnHeaders(19).Text = "Plan Code"
        ListView2.ColumnHeaders(19).Alignment = lvwColumnCenter
        ListView2.ColumnHeaders(20).Text = "Hospital"
        ListView2.ColumnHeaders(20).Alignment = lvwColumnLeft
        ListView2.ColumnHeaders(21).Text = "First Diag."
        ListView2.ColumnHeaders(21).Alignment = lvwColumnCenter
        ListView2.ColumnHeaders(22).Text = "Description"
        ListView2.ColumnHeaders(22).Alignment = lvwColumnLeft
        ListView2.ColumnHeaders(23).Text = "Final Diag."
        ListView2.ColumnHeaders(23).Alignment = lvwColumnCenter
        ListView2.ColumnHeaders(24).Text = "Description"
        ListView2.ColumnHeaders(24).Alignment = lvwColumnLeft
        ListView2.ColumnHeaders(25).Text = "Bill Amt"
        ListView2.ColumnHeaders(25).Alignment = lvwColumnRight
        ListView2.ColumnHeaders(26).Text = "Approved Amt"
        ListView2.ColumnHeaders(26).Alignment = lvwColumnRight
        ListView2.ColumnHeaders(27).Text = "Reserve Amt"
        ListView2.ColumnHeaders(27).Alignment = lvwColumnRight
        ListView2.ColumnHeaders(28).Text = "Recovery Amt"
        ListView2.ColumnHeaders(28).Alignment = lvwColumnRight
        ListView2.ColumnHeaders(29).Text = "Premium Amt"
        ListView2.ColumnHeaders(29).Alignment = lvwColumnRight
        ListView2.ColumnHeaders(30).Text = "Case Status"
        ListView2.ColumnHeaders(30).Alignment = lvwColumnCenter
        ListView2.ColumnHeaders(31).Text = "Claimability"
        ListView2.ColumnHeaders(31).Alignment = lvwColumnCenter
        ListView2.ColumnHeaders(32).Text = "Policy Status"
        ListView2.ColumnHeaders(32).Alignment = lvwColumnCenter
        ListView2.ColumnHeaders(33).Text = "DOA"
        ListView2.ColumnHeaders(33).Alignment = lvwColumnCenter
        ListView2.ColumnHeaders(34).Text = "DOD"
        ListView2.ColumnHeaders(34).Alignment = lvwColumnCenter
        ListView2.ColumnHeaders(35).Text = "DOR - C"
        ListView2.ColumnHeaders(35).Alignment = lvwColumnCenter
        ListView2.ColumnHeaders(36).Text = "Duration"
        ListView2.ColumnHeaders(36).Alignment = lvwColumnRight
        ListView2.ColumnHeaders(37).Text = "DOR - W"
        ListView2.ColumnHeaders(37).Alignment = lvwColumnCenter
        ListView2.ColumnHeaders(38).Text = "Payor Eff. Date"
        ListView2.ColumnHeaders(38).Alignment = lvwColumnCenter
        ListView2.ColumnHeaders(39).Text = "Payor Exp. Date"
        ListView2.ColumnHeaders(39).Alignment = lvwColumnCenter
        ListView2.ColumnHeaders(40).Text = "CPT Code 1"
        ListView2.ColumnHeaders(40).Alignment = lvwColumnCenter
        ListView2.ColumnHeaders(41).Text = "Procedures Desc."
        ListView2.ColumnHeaders(41).Alignment = lvwColumnLeft
        ListView2.ColumnHeaders(42).Text = "CPT Code 2"
        ListView2.ColumnHeaders(42).Alignment = lvwColumnCenter
        ListView2.ColumnHeaders(43).Text = "Procedures Desc."
        ListView2.ColumnHeaders(43).Alignment = lvwColumnLeft
        ListView2.ColumnHeaders(44).Text = "CPT Code 3"
        ListView2.ColumnHeaders(44).Alignment = lvwColumnCenter
        ListView2.ColumnHeaders(45).Text = "Procedures Desc."
        ListView2.ColumnHeaders(45).Alignment = lvwColumnLeft
        ListView2.ColumnHeaders(46).Text = "CPT Code 4"
        ListView2.ColumnHeaders(46).Alignment = lvwColumnCenter
        ListView2.ColumnHeaders(47).Text = "Procedures Desc."
        ListView2.ColumnHeaders(47).Alignment = lvwColumnLeft
        ListView2.ColumnHeaders(48).Text = "CPT Code 5"
        ListView2.ColumnHeaders(48).Alignment = lvwColumnCenter
        ListView2.ColumnHeaders(49).Text = "Procedures Desc."
        ListView2.ColumnHeaders(49).Alignment = lvwColumnLeft
        
        For AA = 50 To 63
            ListView2.ColumnHeaders(AA).Text = ""
        Next AA
        
    ElseIf opt34 = True And cboClaimCase.ListIndex = 0 Then
        ListView2.ColumnHeaders(1).Text = "Payor"
        ListView2.ColumnHeaders(2).Text = "Claim No"
        ListView2.ColumnHeaders(2).Alignment = lvwColumnLeft
        ListView2.ColumnHeaders(3).Text = "BordxNo"
        ListView2.ColumnHeaders(3).Alignment = lvwColumnLeft
        ListView2.ColumnHeaders(4).Text = "Bordx Date"
        ListView2.ColumnHeaders(4).Alignment = lvwColumnCenter
        ListView2.ColumnHeaders(5).Text = "Member No"
        ListView2.ColumnHeaders(5).Alignment = lvwColumnLeft
        ListView2.ColumnHeaders(6).Text = "Policy No"
        ListView2.ColumnHeaders(6).Alignment = lvwColumnLeft
        ListView2.ColumnHeaders(7).Text = "Covered ID"
        ListView2.ColumnHeaders(7).Alignment = lvwColumnCenter
        ListView2.ColumnHeaders(8).Text = "Prin. Name"
        ListView2.ColumnHeaders(8).Alignment = lvwColumnLeft
        ListView2.ColumnHeaders(9).Text = "Patient Name"
        ListView2.ColumnHeaders(9).Alignment = lvwColumnLeft
        ListView2.ColumnHeaders(10).Text = "Prin. IC No"
        ListView2.ColumnHeaders(10).Alignment = lvwColumnLeft
        ListView2.ColumnHeaders(11).Text = "Prin. Sex"
        ListView2.ColumnHeaders(11).Alignment = lvwColumnCenter
        ListView2.ColumnHeaders(12).Text = "Prin. Ageband"
        ListView2.ColumnHeaders(12).Alignment = lvwColumnCenter
        ListView2.ColumnHeaders(13).Text = "Prin. DOB"
        ListView2.ColumnHeaders(13).Alignment = lvwColumnCenter
        ListView2.ColumnHeaders(14).Text = "Patient IC No"
        ListView2.ColumnHeaders(14).Alignment = lvwColumnLeft
        ListView2.ColumnHeaders(15).Text = "Patient Sex"
        ListView2.ColumnHeaders(15).Alignment = lvwColumnCenter
        ListView2.ColumnHeaders(16).Text = "Patient Ageband"
        ListView2.ColumnHeaders(16).Alignment = lvwColumnCenter
        ListView2.ColumnHeaders(17).Text = "Patient DOB"
        ListView2.ColumnHeaders(17).Alignment = lvwColumnCenter
        ListView2.ColumnHeaders(18).Text = "Insured Type"
        ListView2.ColumnHeaders(18).Alignment = lvwColumnCenter
        ListView2.ColumnHeaders(19).Text = "Plan Code"
        ListView2.ColumnHeaders(19).Alignment = lvwColumnCenter
        ListView2.ColumnHeaders(20).Text = "Hospital"
        ListView2.ColumnHeaders(20).Alignment = lvwColumnLeft
        ListView2.ColumnHeaders(21).Text = "First Diag."
        ListView2.ColumnHeaders(21).Alignment = lvwColumnCenter
        ListView2.ColumnHeaders(22).Text = "Description"
        ListView2.ColumnHeaders(22).Alignment = lvwColumnLeft
        ListView2.ColumnHeaders(23).Text = "Final Diag."
        ListView2.ColumnHeaders(23).Alignment = lvwColumnCenter
        ListView2.ColumnHeaders(24).Text = "Description"
        ListView2.ColumnHeaders(24).Alignment = lvwColumnLeft
        ListView2.ColumnHeaders(25).Text = "Bill Amt"
        ListView2.ColumnHeaders(25).Alignment = lvwColumnRight
        ListView2.ColumnHeaders(26).Text = "Approved Amt"
        ListView2.ColumnHeaders(26).Alignment = lvwColumnRight
        ListView2.ColumnHeaders(27).Text = "Reserve Amt"
        ListView2.ColumnHeaders(27).Alignment = lvwColumnRight
        ListView2.ColumnHeaders(28).Text = "Recovery Amt"
        ListView2.ColumnHeaders(28).Alignment = lvwColumnRight
        ListView2.ColumnHeaders(29).Text = "Premium Amt"
        ListView2.ColumnHeaders(29).Alignment = lvwColumnRight
        ListView2.ColumnHeaders(30).Text = "Case Status"
        ListView2.ColumnHeaders(30).Alignment = lvwColumnCenter
        ListView2.ColumnHeaders(31).Text = "Claimability"
        ListView2.ColumnHeaders(31).Alignment = lvwColumnCenter
        ListView2.ColumnHeaders(32).Text = "Policy Status"
        ListView2.ColumnHeaders(32).Alignment = lvwColumnCenter
        ListView2.ColumnHeaders(33).Text = "DOA"
        ListView2.ColumnHeaders(33).Alignment = lvwColumnCenter
        ListView2.ColumnHeaders(34).Text = "DOD"
        ListView2.ColumnHeaders(34).Alignment = lvwColumnCenter
        ListView2.ColumnHeaders(35).Text = "DOR - C"
        ListView2.ColumnHeaders(35).Alignment = lvwColumnCenter
        ListView2.ColumnHeaders(36).Text = "Duration"
        ListView2.ColumnHeaders(36).Alignment = lvwColumnRight
        ListView2.ColumnHeaders(37).Text = "DOR - W"
        ListView2.ColumnHeaders(37).Alignment = lvwColumnCenter
        ListView2.ColumnHeaders(38).Text = "Payor Eff. Date"
        ListView2.ColumnHeaders(38).Alignment = lvwColumnCenter
        ListView2.ColumnHeaders(39).Text = "Payor Exp. Date"
        ListView2.ColumnHeaders(39).Alignment = lvwColumnCenter
        ListView2.ColumnHeaders(40).Text = "Case Remarks"
        ListView2.ColumnHeaders(40).Alignment = lvwColumnLeft
        ListView2.ColumnHeaders(41).Text = "CPT Code 1"
        ListView2.ColumnHeaders(41).Alignment = lvwColumnCenter
        ListView2.ColumnHeaders(42).Text = "Procedures Desc."
        ListView2.ColumnHeaders(42).Alignment = lvwColumnLeft
        ListView2.ColumnHeaders(43).Text = "CPT Code 2"
        ListView2.ColumnHeaders(43).Alignment = lvwColumnCenter
        ListView2.ColumnHeaders(44).Text = "Procedures Desc."
        ListView2.ColumnHeaders(44).Alignment = lvwColumnLeft
        ListView2.ColumnHeaders(45).Text = "CPT Code 3"
        ListView2.ColumnHeaders(45).Alignment = lvwColumnCenter
        ListView2.ColumnHeaders(46).Text = "Procedures Desc."
        ListView2.ColumnHeaders(46).Alignment = lvwColumnLeft
        ListView2.ColumnHeaders(47).Text = "CPT Code 4"
        ListView2.ColumnHeaders(47).Alignment = lvwColumnCenter
        ListView2.ColumnHeaders(48).Text = "Procedures Desc."
        ListView2.ColumnHeaders(48).Alignment = lvwColumnLeft
        ListView2.ColumnHeaders(49).Text = "CPT Code 5"
        ListView2.ColumnHeaders(49).Alignment = lvwColumnCenter
        ListView2.ColumnHeaders(50).Text = "Procedures Desc."
        ListView2.ColumnHeaders(50).Alignment = lvwColumnLeft
        
        For AA = 51 To 63
            ListView2.ColumnHeaders(AA).Text = ""
        Next AA
    
    ElseIf opt35 = True And cboClaimCase.ListIndex = 1 Then
        ListView2.ColumnHeaders(1).Text = "Payor"
        ListView2.ColumnHeaders(2).Text = "Case No"
        ListView2.ColumnHeaders(2).Alignment = lvwColumnLeft
        ListView2.ColumnHeaders(3).Text = "Member No"
        ListView2.ColumnHeaders(3).Alignment = lvwColumnLeft
        ListView2.ColumnHeaders(4).Text = "Policy No"
        ListView2.ColumnHeaders(4).Alignment = lvwColumnLeft
        ListView2.ColumnHeaders(5).Text = "Covered ID"
        ListView2.ColumnHeaders(5).Alignment = lvwColumnCenter
        ListView2.ColumnHeaders(6).Text = "Prin. Name"
        ListView2.ColumnHeaders(6).Alignment = lvwColumnLeft
        ListView2.ColumnHeaders(7).Text = "Patient Name"
        ListView2.ColumnHeaders(7).Alignment = lvwColumnLeft
        ListView2.ColumnHeaders(8).Text = "Prin. IC No"
        ListView2.ColumnHeaders(8).Alignment = lvwColumnLeft
        ListView2.ColumnHeaders(9).Text = "Prin. Sex"
        ListView2.ColumnHeaders(9).Alignment = lvwColumnCenter
        ListView2.ColumnHeaders(10).Text = "Prin. Ageband"
        ListView2.ColumnHeaders(10).Alignment = lvwColumnCenter
        ListView2.ColumnHeaders(11).Text = "Prin. DOB"
        ListView2.ColumnHeaders(11).Alignment = lvwColumnCenter
        ListView2.ColumnHeaders(12).Text = "Patient IC No"
        ListView2.ColumnHeaders(12).Alignment = lvwColumnLeft
        ListView2.ColumnHeaders(13).Text = "Patient Sex"
        ListView2.ColumnHeaders(13).Alignment = lvwColumnCenter
        ListView2.ColumnHeaders(14).Text = "Patient Ageband"
        ListView2.ColumnHeaders(14).Alignment = lvwColumnCenter
        ListView2.ColumnHeaders(15).Text = "Patient DOB"
        ListView2.ColumnHeaders(15).Alignment = lvwColumnCenter
        ListView2.ColumnHeaders(16).Text = "Insured Type"
        ListView2.ColumnHeaders(16).Alignment = lvwColumnCenter
        ListView2.ColumnHeaders(17).Text = "Plan Code"
        ListView2.ColumnHeaders(17).Alignment = lvwColumnCenter
        ListView2.ColumnHeaders(18).Text = "Hospital"
        ListView2.ColumnHeaders(18).Alignment = lvwColumnLeft
        ListView2.ColumnHeaders(19).Text = "First Diag."
        ListView2.ColumnHeaders(19).Alignment = lvwColumnCenter
        ListView2.ColumnHeaders(20).Text = "Description"
        ListView2.ColumnHeaders(20).Alignment = lvwColumnLeft
        ListView2.ColumnHeaders(21).Text = "Final Diag."
        ListView2.ColumnHeaders(21).Alignment = lvwColumnCenter
        ListView2.ColumnHeaders(22).Text = "Description"
        ListView2.ColumnHeaders(22).Alignment = lvwColumnLeft
        ListView2.ColumnHeaders(23).Text = "Bill Amt"
        ListView2.ColumnHeaders(23).Alignment = lvwColumnRight
        ListView2.ColumnHeaders(24).Text = "Approved Amt"
        ListView2.ColumnHeaders(24).Alignment = lvwColumnRight
        ListView2.ColumnHeaders(25).Text = "Not Yet Bordx"
        ListView2.ColumnHeaders(25).Alignment = lvwColumnRight
        ListView2.ColumnHeaders(26).Text = "Bordx"
        ListView2.ColumnHeaders(26).Alignment = lvwColumnRight
        ListView2.ColumnHeaders(27).Text = "Reserve Amt"
        ListView2.ColumnHeaders(27).Alignment = lvwColumnRight
        ListView2.ColumnHeaders(28).Text = "Recovery Amt"
        ListView2.ColumnHeaders(28).Alignment = lvwColumnRight
        ListView2.ColumnHeaders(29).Text = "Premium Amt"
        ListView2.ColumnHeaders(29).Alignment = lvwColumnRight
        ListView2.ColumnHeaders(30).Text = "Case Status"
        ListView2.ColumnHeaders(30).Alignment = lvwColumnCenter
        ListView2.ColumnHeaders(31).Text = "Claimability"
        ListView2.ColumnHeaders(31).Alignment = lvwColumnCenter
        ListView2.ColumnHeaders(32).Text = "Policy Status"
        ListView2.ColumnHeaders(32).Alignment = lvwColumnCenter
        ListView2.ColumnHeaders(33).Text = "DOA"
        ListView2.ColumnHeaders(33).Alignment = lvwColumnCenter
        ListView2.ColumnHeaders(34).Text = "DOD"
        ListView2.ColumnHeaders(34).Alignment = lvwColumnCenter
        ListView2.ColumnHeaders(35).Text = "DOR - C"
        ListView2.ColumnHeaders(35).Alignment = lvwColumnCenter
        ListView2.ColumnHeaders(36).Text = "Duration"
        ListView2.ColumnHeaders(36).Alignment = lvwColumnRight
        ListView2.ColumnHeaders(37).Text = "Payor Eff. Date"
        ListView2.ColumnHeaders(37).Alignment = lvwColumnCenter
        ListView2.ColumnHeaders(38).Text = "Payor Exp. Date"
        ListView2.ColumnHeaders(38).Alignment = lvwColumnCenter
        ListView2.ColumnHeaders(39).Text = "CPT Code 1"
        ListView2.ColumnHeaders(39).Alignment = lvwColumnCenter
        ListView2.ColumnHeaders(40).Text = "Procedures Desc."
        ListView2.ColumnHeaders(40).Alignment = lvwColumnLeft
        ListView2.ColumnHeaders(41).Text = "CPT Code 2"
        ListView2.ColumnHeaders(41).Alignment = lvwColumnCenter
        ListView2.ColumnHeaders(42).Text = "Procedures Desc."
        ListView2.ColumnHeaders(42).Alignment = lvwColumnLeft
        ListView2.ColumnHeaders(43).Text = "CPT Code 3"
        ListView2.ColumnHeaders(43).Alignment = lvwColumnCenter
        ListView2.ColumnHeaders(44).Text = "Procedures Desc."
        ListView2.ColumnHeaders(44).Alignment = lvwColumnLeft
        ListView2.ColumnHeaders(45).Text = "CPT Code 4"
        ListView2.ColumnHeaders(45).Alignment = lvwColumnCenter
        ListView2.ColumnHeaders(46).Text = "Procedures Desc."
        ListView2.ColumnHeaders(46).Alignment = lvwColumnLeft
        ListView2.ColumnHeaders(47).Text = "CPT Code 5"
        ListView2.ColumnHeaders(47).Alignment = lvwColumnCenter
        ListView2.ColumnHeaders(48).Text = "Procedures Desc."
        ListView2.ColumnHeaders(48).Alignment = lvwColumnLeft
        
        For AA = 49 To 63
            ListView2.ColumnHeaders(AA).Text = ""
        Next AA
        
    ElseIf opt34 = True And cboClaimCase.ListIndex = 1 Then
        ListView2.ColumnHeaders(1).Text = "Payor"
        ListView2.ColumnHeaders(2).Text = "Case No"
        ListView2.ColumnHeaders(2).Alignment = lvwColumnLeft
        ListView2.ColumnHeaders(3).Text = "Member No"
        ListView2.ColumnHeaders(3).Alignment = lvwColumnLeft
        ListView2.ColumnHeaders(4).Text = "Policy No"
        ListView2.ColumnHeaders(4).Alignment = lvwColumnLeft
        ListView2.ColumnHeaders(5).Text = "Covered ID"
        ListView2.ColumnHeaders(5).Alignment = lvwColumnCenter
        ListView2.ColumnHeaders(6).Text = "Prin. Name"
        ListView2.ColumnHeaders(6).Alignment = lvwColumnLeft
        ListView2.ColumnHeaders(7).Text = "Patient Name"
        ListView2.ColumnHeaders(7).Alignment = lvwColumnLeft
        ListView2.ColumnHeaders(8).Text = "Prin. IC No"
        ListView2.ColumnHeaders(8).Alignment = lvwColumnLeft
        ListView2.ColumnHeaders(9).Text = "Prin. Sex"
        ListView2.ColumnHeaders(9).Alignment = lvwColumnCenter
        ListView2.ColumnHeaders(10).Text = "Prin. Ageband"
        ListView2.ColumnHeaders(10).Alignment = lvwColumnCenter
        ListView2.ColumnHeaders(11).Text = "Prin. DOB"
        ListView2.ColumnHeaders(11).Alignment = lvwColumnCenter
        ListView2.ColumnHeaders(12).Text = "Patient IC No"
        ListView2.ColumnHeaders(12).Alignment = lvwColumnLeft
        ListView2.ColumnHeaders(13).Text = "Patient Sex"
        ListView2.ColumnHeaders(13).Alignment = lvwColumnCenter
        ListView2.ColumnHeaders(14).Text = "Patient Ageband"
        ListView2.ColumnHeaders(14).Alignment = lvwColumnCenter
        ListView2.ColumnHeaders(15).Text = "Patient DOB"
        ListView2.ColumnHeaders(15).Alignment = lvwColumnCenter
        ListView2.ColumnHeaders(16).Text = "Insured Type"
        ListView2.ColumnHeaders(16).Alignment = lvwColumnCenter
        ListView2.ColumnHeaders(17).Text = "Plan Code"
        ListView2.ColumnHeaders(17).Alignment = lvwColumnCenter
        ListView2.ColumnHeaders(18).Text = "Hospital"
        ListView2.ColumnHeaders(18).Alignment = lvwColumnLeft
        ListView2.ColumnHeaders(19).Text = "First Diag."
        ListView2.ColumnHeaders(19).Alignment = lvwColumnCenter
        ListView2.ColumnHeaders(20).Text = "Description"
        ListView2.ColumnHeaders(20).Alignment = lvwColumnLeft
        ListView2.ColumnHeaders(21).Text = "Final Diag."
        ListView2.ColumnHeaders(21).Alignment = lvwColumnCenter
        ListView2.ColumnHeaders(22).Text = "Description"
        ListView2.ColumnHeaders(22).Alignment = lvwColumnLeft
        ListView2.ColumnHeaders(23).Text = "Bill Amt"
        ListView2.ColumnHeaders(23).Alignment = lvwColumnRight
        ListView2.ColumnHeaders(24).Text = "Approved Amt"
        ListView2.ColumnHeaders(24).Alignment = lvwColumnRight
        ListView2.ColumnHeaders(25).Text = "Not Yet Bordx"
        ListView2.ColumnHeaders(25).Alignment = lvwColumnRight
        ListView2.ColumnHeaders(26).Text = "Bordx"
        ListView2.ColumnHeaders(26).Alignment = lvwColumnRight
        ListView2.ColumnHeaders(27).Text = "Reserve Amt"
        ListView2.ColumnHeaders(27).Alignment = lvwColumnRight
        ListView2.ColumnHeaders(28).Text = "Recovery Amt"
        ListView2.ColumnHeaders(28).Alignment = lvwColumnRight
        ListView2.ColumnHeaders(29).Text = "Premium Amt"
        ListView2.ColumnHeaders(29).Alignment = lvwColumnRight
        ListView2.ColumnHeaders(30).Text = "Case Status"
        ListView2.ColumnHeaders(30).Alignment = lvwColumnCenter
        ListView2.ColumnHeaders(31).Text = "Claimability"
        ListView2.ColumnHeaders(31).Alignment = lvwColumnCenter
        ListView2.ColumnHeaders(32).Text = "Policy Status"
        ListView2.ColumnHeaders(32).Alignment = lvwColumnCenter
        ListView2.ColumnHeaders(33).Text = "DOA"
        ListView2.ColumnHeaders(33).Alignment = lvwColumnCenter
        ListView2.ColumnHeaders(34).Text = "DOD"
        ListView2.ColumnHeaders(34).Alignment = lvwColumnCenter
        ListView2.ColumnHeaders(35).Text = "DOR - C"
        ListView2.ColumnHeaders(35).Alignment = lvwColumnCenter
        ListView2.ColumnHeaders(36).Text = "Duration"
        ListView2.ColumnHeaders(36).Alignment = lvwColumnRight
        ListView2.ColumnHeaders(37).Text = "Payor Eff. Date"
        ListView2.ColumnHeaders(37).Alignment = lvwColumnCenter
        ListView2.ColumnHeaders(38).Text = "Payor Exp. Date"
        ListView2.ColumnHeaders(38).Alignment = lvwColumnCenter
        ListView2.ColumnHeaders(39).Text = "Case Remarks"
        ListView2.ColumnHeaders(39).Alignment = lvwColumnLeft
        ListView2.ColumnHeaders(40).Text = "CPT Code 1"
        ListView2.ColumnHeaders(40).Alignment = lvwColumnCenter
        ListView2.ColumnHeaders(41).Text = "Procedures Desc."
        ListView2.ColumnHeaders(41).Alignment = lvwColumnLeft
        ListView2.ColumnHeaders(42).Text = "CPT Code 2"
        ListView2.ColumnHeaders(42).Alignment = lvwColumnCenter
        ListView2.ColumnHeaders(43).Text = "Procedures Desc."
        ListView2.ColumnHeaders(43).Alignment = lvwColumnLeft
        ListView2.ColumnHeaders(44).Text = "CPT Code 3"
        ListView2.ColumnHeaders(44).Alignment = lvwColumnCenter
        ListView2.ColumnHeaders(45).Text = "Procedures Desc."
        ListView2.ColumnHeaders(45).Alignment = lvwColumnLeft
        ListView2.ColumnHeaders(46).Text = "CPT Code 4"
        ListView2.ColumnHeaders(46).Alignment = lvwColumnCenter
        ListView2.ColumnHeaders(47).Text = "Procedures Desc."
        ListView2.ColumnHeaders(47).Alignment = lvwColumnLeft
        ListView2.ColumnHeaders(48).Text = "CPT Code 5"
        ListView2.ColumnHeaders(48).Alignment = lvwColumnCenter
        ListView2.ColumnHeaders(49).Text = "Procedures Desc."
        ListView2.ColumnHeaders(49).Alignment = lvwColumnLeft
        
        For AA = 50 To 63
            ListView2.ColumnHeaders(AA).Text = ""
        Next AA
    End If
End Sub

Private Sub ListViewHeader16()
If cboClaimCase.ListIndex = 0 Then
    ListView2.ColumnHeaders(1).Text = "Hospital Group"
    ListView2.ColumnHeaders(2).Text = "No of Claims"
    ListView2.ColumnHeaders(2).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(3).Text = "Total Duration"
    ListView2.ColumnHeaders(3).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(4).Text = "Total Bill Amt"
    ListView2.ColumnHeaders(4).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(5).Text = "Total Approved Amt"
    ListView2.ColumnHeaders(5).Alignment = lvwColumnRight
    
    For AA = 6 To 63
        ListView2.ColumnHeaders(AA).Text = ""
    Next AA
ElseIf cboClaimCase.ListIndex = 1 Then
    ListView2.ColumnHeaders(1).Text = "Hospital Group"
    ListView2.ColumnHeaders(2).Text = "No of Cases"
    ListView2.ColumnHeaders(2).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(3).Text = "Total Duration"
    ListView2.ColumnHeaders(3).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(4).Text = "Total Bill Amt"
    ListView2.ColumnHeaders(4).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(5).Text = "Total Approved Amt"
    ListView2.ColumnHeaders(5).Alignment = lvwColumnRight
    
    For AA = 6 To 63
        ListView2.ColumnHeaders(AA).Text = ""
    Next AA
End If
    
End Sub

Private Sub ListViewHeader17()

If cboClaimCase.ListIndex = 0 And opt34 = True Then
    ListView2.ColumnHeaders(1).Text = "Payor"
    ListView2.ColumnHeaders(2).Text = "Hospital"
    ListView2.ColumnHeaders(2).Alignment = lvwColumnLeft
    ListView2.ColumnHeaders(3).Text = "No of Claims"
    ListView2.ColumnHeaders(3).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(4).Text = "Total Duration"
    ListView2.ColumnHeaders(4).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(5).Text = "Bill Amt"
    ListView2.ColumnHeaders(5).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(6).Text = "Approved Amt"
    ListView2.ColumnHeaders(6).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(7).Text = "Average Bill"
    ListView2.ColumnHeaders(7).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(8).Text = "Average Approved"
    ListView2.ColumnHeaders(8).Alignment = lvwColumnRight
    For AA = 9 To 63
        ListView2.ColumnHeaders(AA).Text = ""
    Next AA

ElseIf cboClaimCase.ListIndex = 0 And opt37 = True Then
    ListView2.ColumnHeaders(1).Text = "Payor"
    ListView2.ColumnHeaders(2).Text = "Hospital"
    ListView2.ColumnHeaders(2).Alignment = lvwColumnLeft
    ListView2.ColumnHeaders(3).Text = "No of Claims"
    ListView2.ColumnHeaders(3).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(4).Text = "Total Duration"
    ListView2.ColumnHeaders(4).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(5).Text = "Bill Amt"
    ListView2.ColumnHeaders(5).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(6).Text = "Approved Amt"
    ListView2.ColumnHeaders(6).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(7).Text = "Average Bill"
    ListView2.ColumnHeaders(7).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(8).Text = "Average Approved"
    ListView2.ColumnHeaders(8).Alignment = lvwColumnRight
    For AA = 9 To 63
        ListView2.ColumnHeaders(AA).Text = ""
    Next AA

ElseIf cboClaimCase.ListIndex = 0 And opt29 = True Then
    ListView2.ColumnHeaders(1).Text = "Payor"
    ListView2.ColumnHeaders(2).Text = "Hospital"
    ListView2.ColumnHeaders(2).Alignment = lvwColumnLeft
    ListView2.ColumnHeaders(3).Text = "No of Claims"
    ListView2.ColumnHeaders(3).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(4).Text = "Total Duration"
    ListView2.ColumnHeaders(4).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(5).Text = "Bill Amt"
    ListView2.ColumnHeaders(5).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(6).Text = "Approved Amt"
    ListView2.ColumnHeaders(6).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(7).Text = "Average Bill"
    ListView2.ColumnHeaders(7).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(8).Text = "Average Approved"
    ListView2.ColumnHeaders(8).Alignment = lvwColumnRight
    For AA = 9 To 63
        ListView2.ColumnHeaders(AA).Text = ""
    Next AA

ElseIf cboClaimCase.ListIndex = 0 And opt31 = True Then
    ListView2.ColumnHeaders(1).Text = "Hospital"
    ListView2.ColumnHeaders(2).Text = "No of Claims"
    ListView2.ColumnHeaders(2).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(3).Text = "Total Duration"
    ListView2.ColumnHeaders(3).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(4).Text = "Bill Amt"
    ListView2.ColumnHeaders(4).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(5).Text = "Approved Amt"
    ListView2.ColumnHeaders(5).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(6).Text = "Average Bill"
    ListView2.ColumnHeaders(6).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(7).Text = "Average Approved"
    ListView2.ColumnHeaders(7).Alignment = lvwColumnRight
    For AA = 8 To 63
        ListView2.ColumnHeaders(AA).Text = ""
    Next AA

ElseIf cboClaimCase.ListIndex = 1 And opt31 = True Then
    ListView2.ColumnHeaders(1).Text = "Hospital"
    ListView2.ColumnHeaders(2).Text = "No of Cases"
    ListView2.ColumnHeaders(2).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(3).Text = "Total Duration"
    ListView2.ColumnHeaders(3).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(4).Text = "Bill Amt"
    ListView2.ColumnHeaders(4).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(5).Text = "Approved Amt"
    ListView2.ColumnHeaders(5).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(6).Text = "Average Bill"
    ListView2.ColumnHeaders(6).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(7).Text = "Average Approved"
    ListView2.ColumnHeaders(7).Alignment = lvwColumnRight
    For AA = 8 To 63
        ListView2.ColumnHeaders(AA).Text = ""
    Next AA


ElseIf cboClaimCase.ListIndex = 0 And opt33 = True Then
    ListView2.ColumnHeaders(1).Text = "Hospital"
    ListView2.ColumnHeaders(2).Text = "No of Claims"
    ListView2.ColumnHeaders(2).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(3).Text = "Total Duration"
    ListView2.ColumnHeaders(3).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(4).Text = "Bill Amt"
    ListView2.ColumnHeaders(4).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(5).Text = "Approved Amt"
    ListView2.ColumnHeaders(5).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(6).Text = "Average Bill"
    ListView2.ColumnHeaders(6).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(7).Text = "Average Approved"
    ListView2.ColumnHeaders(7).Alignment = lvwColumnRight
    For AA = 8 To 63
        ListView2.ColumnHeaders(AA).Text = ""
    Next AA

ElseIf cboClaimCase.ListIndex = 0 And opt30 = True Then
    ListView2.ColumnHeaders(1).Text = "Hospital"
    ListView2.ColumnHeaders(2).Text = "Final Diag."
    ListView2.ColumnHeaders(2).Alignment = lvwColumnCenter
    ListView2.ColumnHeaders(3).Text = "Description"
    ListView2.ColumnHeaders(3).Alignment = lvwColumnLeft
    ListView2.ColumnHeaders(4).Text = "No of Claims"
    ListView2.ColumnHeaders(4).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(5).Text = "Total Duration"
    ListView2.ColumnHeaders(5).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(6).Text = "Bill Amt"
    ListView2.ColumnHeaders(6).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(7).Text = "Approved Amt"
    ListView2.ColumnHeaders(7).Alignment = lvwColumnRight
    For AA = 8 To 63
        ListView2.ColumnHeaders(AA).Text = ""
    Next AA

ElseIf cboClaimCase.ListIndex = 1 And opt34 = True Then
    ListView2.ColumnHeaders(1).Text = "Payor"
    ListView2.ColumnHeaders(2).Text = "Hospital"
    ListView2.ColumnHeaders(2).Alignment = lvwColumnLeft
    ListView2.ColumnHeaders(3).Text = "No of Cases"
    ListView2.ColumnHeaders(3).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(4).Text = "Total Duration"
    ListView2.ColumnHeaders(4).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(5).Text = "Bill Amt"
    ListView2.ColumnHeaders(5).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(6).Text = "Approved Amt"
    ListView2.ColumnHeaders(6).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(7).Text = "Average Bill"
    ListView2.ColumnHeaders(7).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(8).Text = "Average Approved"
    ListView2.ColumnHeaders(8).Alignment = lvwColumnRight
    For AA = 9 To 63
        ListView2.ColumnHeaders(AA).Text = ""
    Next AA

ElseIf cboClaimCase.ListIndex = 0 And opt37 = True Then
    ListView2.ColumnHeaders(1).Text = "Payor"
    ListView2.ColumnHeaders(2).Text = "Hospital"
    ListView2.ColumnHeaders(2).Alignment = lvwColumnLeft
    ListView2.ColumnHeaders(3).Text = "No of Claims"
    ListView2.ColumnHeaders(3).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(4).Text = "Total Duration"
    ListView2.ColumnHeaders(4).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(5).Text = "Bill Amt"
    ListView2.ColumnHeaders(5).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(6).Text = "Approved Amt"
    ListView2.ColumnHeaders(6).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(7).Text = "Average Bill"
    ListView2.ColumnHeaders(7).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(8).Text = "Average Approved"
    ListView2.ColumnHeaders(8).Alignment = lvwColumnRight
    For AA = 9 To 63
        ListView2.ColumnHeaders(AA).Text = ""
    Next AA

ElseIf cboClaimCase.ListIndex = 1 And opt29 = True Then
    ListView2.ColumnHeaders(1).Text = "Payor"
    ListView2.ColumnHeaders(2).Text = "Hospital"
    ListView2.ColumnHeaders(2).Alignment = lvwColumnLeft
    ListView2.ColumnHeaders(3).Text = "No of Cases"
    ListView2.ColumnHeaders(3).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(4).Text = "Total Duration"
    ListView2.ColumnHeaders(4).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(5).Text = "Bill Amt"
    ListView2.ColumnHeaders(5).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(6).Text = "Approved Amt"
    ListView2.ColumnHeaders(6).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(7).Text = "Average Bill"
    ListView2.ColumnHeaders(7).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(8).Text = "Average Approved"
    ListView2.ColumnHeaders(8).Alignment = lvwColumnRight
    For AA = 9 To 63
        ListView2.ColumnHeaders(AA).Text = ""
    Next AA


ElseIf cboClaimCase.ListIndex = 1 And opt33 = True Then
    ListView2.ColumnHeaders(1).Text = "Hospital"
    ListView2.ColumnHeaders(2).Text = "No of Cases"
    ListView2.ColumnHeaders(2).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(3).Text = "Total Duration"
    ListView2.ColumnHeaders(3).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(4).Text = "Bill Amt"
    ListView2.ColumnHeaders(4).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(5).Text = "Approved Amt"
    ListView2.ColumnHeaders(5).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(6).Text = "Average Bill"
    ListView2.ColumnHeaders(6).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(7).Text = "Average Approved"
    ListView2.ColumnHeaders(7).Alignment = lvwColumnRight
    For AA = 8 To 63
        ListView2.ColumnHeaders(AA).Text = ""
    Next AA

ElseIf cboClaimCase.ListIndex = 1 And opt30 = True Then
    ListView2.ColumnHeaders(1).Text = "Hospital"
    ListView2.ColumnHeaders(2).Text = "Final Diag."
    ListView2.ColumnHeaders(2).Alignment = lvwColumnCenter
    ListView2.ColumnHeaders(3).Text = "Description"
    ListView2.ColumnHeaders(3).Alignment = lvwColumnLeft
    ListView2.ColumnHeaders(4).Text = "No of Cases"
    ListView2.ColumnHeaders(4).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(5).Text = "Total Duration"
    ListView2.ColumnHeaders(5).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(6).Text = "Bill Amt"
    ListView2.ColumnHeaders(6).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(7).Text = "Approved Amt"
    ListView2.ColumnHeaders(7).Alignment = lvwColumnRight
    For AA = 8 To 63
        ListView2.ColumnHeaders(AA).Text = ""
    Next AA
    
ElseIf cboClaimCase.ListIndex = 0 And opt35 = True Then

    ListView2.ColumnHeaders(1).Text = "Payor"
    ListView2.ColumnHeaders(2).Text = "Hospital"
    ListView2.ColumnHeaders(2).Alignment = lvwColumnLeft
    ListView2.ColumnHeaders(3).Text = "Claim No"
    ListView2.ColumnHeaders(3).Alignment = lvwColumnLeft
    ListView2.ColumnHeaders(4).Text = "Member No"
    ListView2.ColumnHeaders(4).Alignment = lvwColumnLeft
    ListView2.ColumnHeaders(5).Text = "Covered ID"
    ListView2.ColumnHeaders(5).Alignment = lvwColumnCenter
    ListView2.ColumnHeaders(6).Text = "Patient Name"
    ListView2.ColumnHeaders(6).Alignment = lvwColumnLeft
    ListView2.ColumnHeaders(7).Text = "Bill Amt"
    ListView2.ColumnHeaders(7).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(8).Text = "Approved Amt"
    ListView2.ColumnHeaders(8).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(9).Text = "Member Amt"
    ListView2.ColumnHeaders(9).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(10).Text = "DOA"
    ListView2.ColumnHeaders(10).Alignment = lvwColumnCenter
    ListView2.ColumnHeaders(11).Text = "DOD"
    ListView2.ColumnHeaders(11).Alignment = lvwColumnCenter
    ListView2.ColumnHeaders(12).Text = "Case Status"
    ListView2.ColumnHeaders(12).Alignment = lvwColumnCenter
    ListView2.ColumnHeaders(13).Text = "Claimability"
    ListView2.ColumnHeaders(13).Alignment = lvwColumnCenter
    ListView2.ColumnHeaders(14).Text = "Final Diag."
    ListView2.ColumnHeaders(14).Alignment = lvwColumnCenter
    
    For AA = 15 To 63
        ListView2.ColumnHeaders(AA).Text = ""
    Next AA
    
ElseIf cboClaimCase.ListIndex = 1 And opt35 = True Then

    ListView2.ColumnHeaders(1).Text = "Payor"
    ListView2.ColumnHeaders(2).Text = "Hospital"
    ListView2.ColumnHeaders(2).Alignment = lvwColumnLeft
    ListView2.ColumnHeaders(3).Text = "Case No"
    ListView2.ColumnHeaders(3).Alignment = lvwColumnLeft
    ListView2.ColumnHeaders(4).Text = "Member No"
    ListView2.ColumnHeaders(4).Alignment = lvwColumnLeft
    ListView2.ColumnHeaders(5).Text = "Covered ID"
    ListView2.ColumnHeaders(5).Alignment = lvwColumnCenter
    ListView2.ColumnHeaders(6).Text = "Patient Name"
    ListView2.ColumnHeaders(6).Alignment = lvwColumnLeft
    ListView2.ColumnHeaders(7).Text = "Bill Amt"
    ListView2.ColumnHeaders(7).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(8).Text = "Approved Amt"
    ListView2.ColumnHeaders(8).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(9).Text = "Member Amt"
    ListView2.ColumnHeaders(9).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(10).Text = "DOA"
    ListView2.ColumnHeaders(10).Alignment = lvwColumnCenter
    ListView2.ColumnHeaders(11).Text = "DOD"
    ListView2.ColumnHeaders(11).Alignment = lvwColumnCenter
    ListView2.ColumnHeaders(12).Text = "Case Status"
    ListView2.ColumnHeaders(12).Alignment = lvwColumnCenter
    ListView2.ColumnHeaders(13).Text = "Claimability"
    ListView2.ColumnHeaders(13).Alignment = lvwColumnCenter
    ListView2.ColumnHeaders(14).Text = "Final Diag."
    ListView2.ColumnHeaders(14).Alignment = lvwColumnCenter
    
    For AA = 15 To 63
        ListView2.ColumnHeaders(AA).Text = ""
    Next AA
    
End If

End Sub

Private Sub ListViewHeader2()
If opt35 = True Then
    ListView2.ColumnHeaders(1).Text = "No"
    ListView2.ColumnHeaders(2).Text = "Claim No"
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
    For AA = 9 To 63
        ListView2.ColumnHeaders(AA).Text = ""
    Next AA
    
ElseIf opt34 = True Then
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

ElseIf opt33 = True Then
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
    ListView2.ColumnHeaders(8).Text = "July"
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
    
    For AA = 14 To 63
        ListView2.ColumnHeaders(AA).Text = ""
    Next AA
    
End If

End Sub

Private Sub ListViewHeader9()
If opt35 = True Then
    ListView2.ColumnHeaders(1).Text = "No"
    ListView2.ColumnHeaders(2).Text = "Claim No"
    ListView2.ColumnHeaders(2).Alignment = lvwColumnLeft
    ListView2.ColumnHeaders(3).Text = "Member No"
    ListView2.ColumnHeaders(3).Alignment = lvwColumnLeft
    ListView2.ColumnHeaders(4).Text = "Policy No"
    ListView2.ColumnHeaders(4).Alignment = lvwColumnLeft
    ListView2.ColumnHeaders(5).Text = "Patient Name"
    ListView2.ColumnHeaders(5).Alignment = lvwColumnLeft
    ListView2.ColumnHeaders(6).Text = "DOA"
    ListView2.ColumnHeaders(6).Alignment = lvwColumnCenter
    ListView2.ColumnHeaders(7).Text = "DOD"
    ListView2.ColumnHeaders(7).Alignment = lvwColumnCenter
    ListView2.ColumnHeaders(8).Text = "Hospital Name"
    ListView2.ColumnHeaders(8).Alignment = lvwColumnLeft
    ListView2.ColumnHeaders(9).Text = "First Diag."
    ListView2.ColumnHeaders(9).Alignment = lvwColumnCenter
    For AA = 10 To 63
        ListView2.ColumnHeaders(AA).Text = ""
    Next AA
    
ElseIf opt34 = True Then
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

ElseIf opt33 = True Then
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
    ListView2.ColumnHeaders(8).Text = "July"
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
    
    For AA = 14 To 63
        ListView2.ColumnHeaders(AA).Text = ""
    Next AA
    
End If

End Sub

    


Private Sub ListViewHeader4()
    
If cboClaimCase.ListIndex = 0 And opt35 = True Then
    ListView2.ColumnHeaders(1).Text = "Payor"
    ListView2.ColumnHeaders(2).Text = "Agent No"
    ListView2.ColumnHeaders(2).Alignment = lvwColumnLeft
    ListView2.ColumnHeaders(3).Text = "Claim No"
    ListView2.ColumnHeaders(3).Alignment = lvwColumnLeft
    ListView2.ColumnHeaders(4).Text = "Member No"
    ListView2.ColumnHeaders(4).Alignment = lvwColumnLeft
    ListView2.ColumnHeaders(5).Text = "Covered ID"
    ListView2.ColumnHeaders(5).Alignment = lvwColumnCenter
    ListView2.ColumnHeaders(6).Text = "Patient Name"
    ListView2.ColumnHeaders(6).Alignment = lvwColumnLeft
    ListView2.ColumnHeaders(7).Text = "Doctor"
    ListView2.ColumnHeaders(7).Alignment = lvwColumnLeft
    ListView2.ColumnHeaders(8).Text = "Bill Amt"
    ListView2.ColumnHeaders(8).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(9).Text = "Approved Amt"
    ListView2.ColumnHeaders(9).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(10).Text = "Recovery Amt"
    ListView2.ColumnHeaders(10).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(11).Text = "First Diag."
    ListView2.ColumnHeaders(11).Alignment = lvwColumnCenter
    ListView2.ColumnHeaders(12).Text = "Description"
    ListView2.ColumnHeaders(12).Alignment = lvwColumnLeft
    ListView2.ColumnHeaders(13).Text = "Final Diag."
    ListView2.ColumnHeaders(13).Alignment = lvwColumnCenter
    ListView2.ColumnHeaders(14).Text = "Description"
    ListView2.ColumnHeaders(14).Alignment = lvwColumnLeft
    ListView2.ColumnHeaders(15).Text = "DOA"
    ListView2.ColumnHeaders(15).Alignment = lvwColumnCenter
    ListView2.ColumnHeaders(16).Text = "DOD"
    ListView2.ColumnHeaders(16).Alignment = lvwColumnCenter
    ListView2.ColumnHeaders(17).Text = "Ins Code"
    ListView2.ColumnHeaders(17).Alignment = lvwColumnCenter
    ListView2.ColumnHeaders(18).Text = "Plan Code"
    ListView2.ColumnHeaders(18).Alignment = lvwColumnCenter
    ListView2.ColumnHeaders(19).Text = "Case Status"
    ListView2.ColumnHeaders(19).Alignment = lvwColumnCenter
    For AA = 20 To 63
        ListView2.ColumnHeaders(AA).Text = ""
    Next AA
    
ElseIf cboClaimCase.ListIndex = 0 And opt34 = True Then
    ListView2.ColumnHeaders(1).Text = "Payor"
    ListView2.ColumnHeaders(2).Text = "Agent No"
    ListView2.ColumnHeaders(2).Alignment = lvwColumnLeft
    ListView2.ColumnHeaders(3).Text = "No of Claims"
    ListView2.ColumnHeaders(3).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(4).Text = "Total Duration"
    ListView2.ColumnHeaders(4).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(5).Text = "Total Actual Amt"
    ListView2.ColumnHeaders(5).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(6).Text = "Total Approved Amt"
    ListView2.ColumnHeaders(6).Alignment = lvwColumnRight
    
    For AA = 7 To 63
        ListView2.ColumnHeaders(AA).Text = ""
    Next AA

ElseIf cboClaimCase.ListIndex = 0 And opt33 = True Then
    ListView2.ColumnHeaders(1).Text = "Payor"
    ListView2.ColumnHeaders(2).Text = "Agent No"
    ListView2.ColumnHeaders(2).Alignment = lvwColumnLeft
    ListView2.ColumnHeaders(3).Text = "No of Claims"
    ListView2.ColumnHeaders(3).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(4).Text = "Total Duration"
    ListView2.ColumnHeaders(4).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(5).Text = "Total Actual Amt"
    ListView2.ColumnHeaders(5).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(6).Text = "Total Approved Amt"
    ListView2.ColumnHeaders(6).Alignment = lvwColumnRight
    
    For AA = 7 To 63
        ListView2.ColumnHeaders(AA).Text = ""
    Next AA

ElseIf cboClaimCase.ListIndex = 1 And opt34 = True Then
    ListView2.ColumnHeaders(1).Text = "Payor"
    ListView2.ColumnHeaders(2).Text = "Agent No"
    ListView2.ColumnHeaders(2).Alignment = lvwColumnLeft
    ListView2.ColumnHeaders(3).Text = "No of Cases"
    ListView2.ColumnHeaders(3).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(4).Text = "Total Duration"
    ListView2.ColumnHeaders(4).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(5).Text = "Total Actual Amt"
    ListView2.ColumnHeaders(5).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(6).Text = "Total Approved Amt"
    ListView2.ColumnHeaders(6).Alignment = lvwColumnRight
    
    For AA = 7 To 63
        ListView2.ColumnHeaders(AA).Text = ""
    Next AA

ElseIf cboClaimCase.ListIndex = 1 And opt33 = True Then
    ListView2.ColumnHeaders(1).Text = "Payor"
    ListView2.ColumnHeaders(2).Text = "Agent No"
    ListView2.ColumnHeaders(2).Alignment = lvwColumnLeft
    ListView2.ColumnHeaders(3).Text = "No of Cases"
    ListView2.ColumnHeaders(3).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(4).Text = "Total Duration"
    ListView2.ColumnHeaders(4).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(5).Text = "Total Actual Amt"
    ListView2.ColumnHeaders(5).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(6).Text = "Total Approved Amt"
    ListView2.ColumnHeaders(6).Alignment = lvwColumnRight
    
    For AA = 7 To 63
        ListView2.ColumnHeaders(AA).Text = ""
    Next AA

ElseIf cboClaimCase.ListIndex = 0 And opt29 = True Then
    ListView2.ColumnHeaders(1).Text = "Payor"
    ListView2.ColumnHeaders(2).Text = "Agent No"
    ListView2.ColumnHeaders(2).Alignment = lvwColumnLeft
    ListView2.ColumnHeaders(3).Text = "Doctor Name"
    ListView2.ColumnHeaders(3).Alignment = lvwColumnLeft
    ListView2.ColumnHeaders(4).Text = "No of Claims"
    ListView2.ColumnHeaders(4).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(5).Text = "Total Duration"
    ListView2.ColumnHeaders(5).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(6).Text = "Total Actual Amt"
    ListView2.ColumnHeaders(6).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(7).Text = "Total Approved Amt"
    ListView2.ColumnHeaders(7).Alignment = lvwColumnRight
    
    For AA = 8 To 63
        ListView2.ColumnHeaders(AA).Text = ""
    Next AA

ElseIf cboClaimCase.ListIndex = 1 And opt29 = True Then
    ListView2.ColumnHeaders(1).Text = "Payor"
    ListView2.ColumnHeaders(2).Text = "Agent No"
    ListView2.ColumnHeaders(2).Alignment = lvwColumnLeft
    ListView2.ColumnHeaders(3).Text = "Doctor Name"
    ListView2.ColumnHeaders(3).Alignment = lvwColumnLeft
    ListView2.ColumnHeaders(4).Text = "No of Cases"
    ListView2.ColumnHeaders(4).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(5).Text = "Total Duration"
    ListView2.ColumnHeaders(5).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(6).Text = "Total Actual Amt"
    ListView2.ColumnHeaders(6).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(7).Text = "Total Approved Amt"
    ListView2.ColumnHeaders(7).Alignment = lvwColumnRight
    
    For AA = 8 To 63
        ListView2.ColumnHeaders(AA).Text = ""
    Next AA
    
ElseIf cboClaimCase.ListIndex = 1 And opt35 = True Then
    ListView2.ColumnHeaders(1).Text = "Payor"
    ListView2.ColumnHeaders(2).Text = "Agent No"
    ListView2.ColumnHeaders(2).Alignment = lvwColumnLeft
    ListView2.ColumnHeaders(3).Text = "Case No"
    ListView2.ColumnHeaders(3).Alignment = lvwColumnLeft
    ListView2.ColumnHeaders(4).Text = "Member No"
    ListView2.ColumnHeaders(4).Alignment = lvwColumnLeft
    ListView2.ColumnHeaders(5).Text = "Covered ID"
    ListView2.ColumnHeaders(5).Alignment = lvwColumnCenter
    ListView2.ColumnHeaders(6).Text = "Patient Name"
    ListView2.ColumnHeaders(6).Alignment = lvwColumnLeft
    ListView2.ColumnHeaders(7).Text = "Doctor"
    ListView2.ColumnHeaders(7).Alignment = lvwColumnLeft
    ListView2.ColumnHeaders(8).Text = "Bill Amt"
    ListView2.ColumnHeaders(8).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(9).Text = "Approved Amt"
    ListView2.ColumnHeaders(9).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(10).Text = "Recovery Amt"
    ListView2.ColumnHeaders(10).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(11).Text = "First Diag."
    ListView2.ColumnHeaders(11).Alignment = lvwColumnCenter
    ListView2.ColumnHeaders(12).Text = "Description"
    ListView2.ColumnHeaders(12).Alignment = lvwColumnLeft
    ListView2.ColumnHeaders(13).Text = "Final Diag."
    ListView2.ColumnHeaders(13).Alignment = lvwColumnCenter
    ListView2.ColumnHeaders(14).Text = "Description"
    ListView2.ColumnHeaders(14).Alignment = lvwColumnLeft
    ListView2.ColumnHeaders(15).Text = "DOA"
    ListView2.ColumnHeaders(15).Alignment = lvwColumnCenter
    ListView2.ColumnHeaders(16).Text = "DOD"
    ListView2.ColumnHeaders(16).Alignment = lvwColumnCenter
    ListView2.ColumnHeaders(17).Text = "Ins Code"
    ListView2.ColumnHeaders(17).Alignment = lvwColumnCenter
    ListView2.ColumnHeaders(18).Text = "Plan Code"
    ListView2.ColumnHeaders(18).Alignment = lvwColumnCenter
    ListView2.ColumnHeaders(19).Text = "Case Status"
    ListView2.ColumnHeaders(19).Alignment = lvwColumnCenter
    For AA = 20 To 63
        ListView2.ColumnHeaders(AA).Text = ""
    Next AA
End If
End Sub

Private Sub ListViewHeader3()
If cboClaimCase.ListIndex = 0 And opt35 = True Then
    ListView2.ColumnHeaders(1).Text = "Ageband"
    ListView2.ColumnHeaders(2).Text = "Total Claims"
    ListView2.ColumnHeaders(2).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(3).Text = "Total Duration"
    ListView2.ColumnHeaders(3).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(4).Text = "Total Actual Amt"
    ListView2.ColumnHeaders(4).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(5).Text = "Total Approved Amt"
    ListView2.ColumnHeaders(5).Alignment = lvwColumnRight
    
    For AA = 6 To 63
        ListView2.ColumnHeaders(AA).Text = ""
    Next AA
ElseIf cboClaimCase.ListIndex = 0 And opt34 = True Then
    ListView2.ColumnHeaders(1).Text = "Gender"
    ListView2.ColumnHeaders(2).Text = "Ageband"
    ListView2.ColumnHeaders(2).Alignment = lvwColumnCenter
    ListView2.ColumnHeaders(3).Text = "Total Claims"
    ListView2.ColumnHeaders(3).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(4).Text = "Total Duration"
    ListView2.ColumnHeaders(4).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(5).Text = "Total Actual Amt"
    ListView2.ColumnHeaders(5).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(6).Text = "Total Approved Amt"
    ListView2.ColumnHeaders(6).Alignment = lvwColumnRight
        
    For AA = 7 To 63
        ListView2.ColumnHeaders(AA).Text = ""
    Next AA
ElseIf cboClaimCase.ListIndex = 0 And opt33 = True Then
    ListView2.ColumnHeaders(1).Text = "Payor"
    ListView2.ColumnHeaders(2).Text = "Insured Type"
    ListView2.ColumnHeaders(2).Alignment = lvwColumnCenter
    ListView2.ColumnHeaders(3).Text = "Ageband"
    ListView2.ColumnHeaders(3).Alignment = lvwColumnCenter
    ListView2.ColumnHeaders(4).Text = "Total Claims"
    ListView2.ColumnHeaders(4).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(5).Text = "Total Duration"
    ListView2.ColumnHeaders(5).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(6).Text = "Total Actual Amt"
    ListView2.ColumnHeaders(6).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(7).Text = "Total Approved Amt"
    ListView2.ColumnHeaders(7).Alignment = lvwColumnRight
        
    For AA = 8 To 63
        ListView2.ColumnHeaders(AA).Text = ""
    Next AA
ElseIf cboClaimCase.ListIndex = 0 And opt29 = True Then
    ListView2.ColumnHeaders(1).Text = "Race Code"
    ListView2.ColumnHeaders(2).Text = "Ageband"
    ListView2.ColumnHeaders(2).Alignment = lvwColumnCenter
    ListView2.ColumnHeaders(3).Text = "Total Claims"
    ListView2.ColumnHeaders(3).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(4).Text = "Total Duration"
    ListView2.ColumnHeaders(4).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(5).Text = "Total Actual Amt"
    ListView2.ColumnHeaders(5).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(6).Text = "Total Approved Amt"
    ListView2.ColumnHeaders(6).Alignment = lvwColumnRight
        
    For AA = 7 To 63
        ListView2.ColumnHeaders(AA).Text = ""
    Next AA
ElseIf cboClaimCase.ListIndex = 0 And opt37 = True Then
    ListView2.ColumnHeaders(1).Text = "Insured Type"
    ListView2.ColumnHeaders(2).Text = "Plan Code"
    ListView2.ColumnHeaders(2).Alignment = lvwColumnCenter
    ListView2.ColumnHeaders(3).Text = "Age Band"
    ListView2.ColumnHeaders(3).Alignment = lvwColumnCenter
    ListView2.ColumnHeaders(4).Text = "Total Claims"
    ListView2.ColumnHeaders(4).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(5).Text = "Total Duration"
    ListView2.ColumnHeaders(5).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(6).Text = "Total Actual Amt"
    ListView2.ColumnHeaders(6).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(7).Text = "Total Approved Amt"
    ListView2.ColumnHeaders(7).Alignment = lvwColumnRight
        
    For AA = 8 To 63
        ListView2.ColumnHeaders(AA).Text = ""
    Next AA
ElseIf cboClaimCase.ListIndex = 0 And opt30 = True Then
    ListView2.ColumnHeaders(1).Text = "Nature of Claim"
    ListView2.ColumnHeaders(2).Text = "Age Band"
    ListView2.ColumnHeaders(2).Alignment = lvwColumnCenter
    ListView2.ColumnHeaders(3).Text = "Total Claims"
    ListView2.ColumnHeaders(3).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(4).Text = "Total Duration"
    ListView2.ColumnHeaders(4).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(5).Text = "Total Actual Amt"
    ListView2.ColumnHeaders(5).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(6).Text = "Total Approved Amt"
    ListView2.ColumnHeaders(6).Alignment = lvwColumnRight
        
    For AA = 7 To 63
        ListView2.ColumnHeaders(AA).Text = ""
    Next AA
    
ElseIf cboClaimCase.ListIndex = 1 And opt35 = True Then
    ListView2.ColumnHeaders(1).Text = "Ageband"
    ListView2.ColumnHeaders(2).Text = "Total Cases"
    ListView2.ColumnHeaders(2).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(3).Text = "Total Duration"
    ListView2.ColumnHeaders(3).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(4).Text = "Total Actual Amt"
    ListView2.ColumnHeaders(4).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(5).Text = "Total Approved Amt"
    ListView2.ColumnHeaders(5).Alignment = lvwColumnRight
    
    For AA = 6 To 63
        ListView2.ColumnHeaders(AA).Text = ""
    Next AA
ElseIf cboClaimCase.ListIndex = 1 And opt34 = True Then
    ListView2.ColumnHeaders(1).Text = "Gender"
    ListView2.ColumnHeaders(2).Text = "Ageband"
    ListView2.ColumnHeaders(2).Alignment = lvwColumnCenter
    ListView2.ColumnHeaders(3).Text = "Total Cases"
    ListView2.ColumnHeaders(3).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(4).Text = "Total Duration"
    ListView2.ColumnHeaders(4).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(5).Text = "Total Actual Amt"
    ListView2.ColumnHeaders(5).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(6).Text = "Total Approved Amt"
    ListView2.ColumnHeaders(6).Alignment = lvwColumnRight
        
    For AA = 7 To 63
        ListView2.ColumnHeaders(AA).Text = ""
    Next AA
ElseIf cboClaimCase.ListIndex = 1 And opt33 = True Then
    ListView2.ColumnHeaders(1).Text = "Payor"
    ListView2.ColumnHeaders(2).Text = "Insured Type"
    ListView2.ColumnHeaders(2).Alignment = lvwColumnCenter
    ListView2.ColumnHeaders(3).Text = "Ageband"
    ListView2.ColumnHeaders(3).Alignment = lvwColumnCenter
    ListView2.ColumnHeaders(4).Text = "Total Cases"
    ListView2.ColumnHeaders(4).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(5).Text = "Total Duration"
    ListView2.ColumnHeaders(5).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(6).Text = "Total Actual Amt"
    ListView2.ColumnHeaders(6).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(7).Text = "Total Approved Amt"
    ListView2.ColumnHeaders(7).Alignment = lvwColumnRight
        
    For AA = 8 To 63
        ListView2.ColumnHeaders(AA).Text = ""
    Next AA
ElseIf cboClaimCase.ListIndex = 1 And opt29 = True Then
    ListView2.ColumnHeaders(1).Text = "Race Code"
    ListView2.ColumnHeaders(2).Text = "Ageband"
    ListView2.ColumnHeaders(2).Alignment = lvwColumnCenter
    ListView2.ColumnHeaders(3).Text = "Total Cases"
    ListView2.ColumnHeaders(3).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(4).Text = "Total Duration"
    ListView2.ColumnHeaders(4).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(5).Text = "Total Actual Amt"
    ListView2.ColumnHeaders(5).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(6).Text = "Total Approved Amt"
    ListView2.ColumnHeaders(6).Alignment = lvwColumnRight
        
    For AA = 7 To 63
        ListView2.ColumnHeaders(AA).Text = ""
    Next AA
ElseIf cboClaimCase.ListIndex = 1 And opt37 = True Then
    ListView2.ColumnHeaders(1).Text = "Insured Type"
    ListView2.ColumnHeaders(2).Text = "Plan Code"
    ListView2.ColumnHeaders(2).Alignment = lvwColumnCenter
    ListView2.ColumnHeaders(3).Text = "Age Band"
    ListView2.ColumnHeaders(3).Alignment = lvwColumnCenter
    ListView2.ColumnHeaders(4).Text = "Total Cases"
    ListView2.ColumnHeaders(4).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(5).Text = "Total Duration"
    ListView2.ColumnHeaders(5).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(6).Text = "Total Actual Amt"
    ListView2.ColumnHeaders(6).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(7).Text = "Total Approved Amt"
    ListView2.ColumnHeaders(7).Alignment = lvwColumnRight
        
    For AA = 8 To 63
        ListView2.ColumnHeaders(AA).Text = ""
    Next AA
ElseIf cboClaimCase.ListIndex = 1 And opt30 = True Then
    ListView2.ColumnHeaders(1).Text = "Nature of Claim"
    ListView2.ColumnHeaders(2).Text = "Age Band"
    ListView2.ColumnHeaders(2).Alignment = lvwColumnCenter
    ListView2.ColumnHeaders(3).Text = "Total Cases"
    ListView2.ColumnHeaders(3).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(4).Text = "Total Duration"
    ListView2.ColumnHeaders(4).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(5).Text = "Total Actual Amt"
    ListView2.ColumnHeaders(5).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(6).Text = "Total Approved Amt"
    ListView2.ColumnHeaders(6).Alignment = lvwColumnRight
        
    For AA = 7 To 63
        ListView2.ColumnHeaders(AA).Text = ""
    Next AA
End If
End Sub

Private Sub ListViewHeader14()
If opt35 = True And cboClaimCase.ListIndex = 0 Then
    ListView2.ColumnHeaders(1).Text = "Sex"
    ListView2.ColumnHeaders(2).Text = "Total Claims"
    ListView2.ColumnHeaders(2).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(3).Text = "Total Duration"
    ListView2.ColumnHeaders(3).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(4).Text = "Total Actual Amt"
    ListView2.ColumnHeaders(4).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(5).Text = "Total Approved Amt"
    ListView2.ColumnHeaders(5).Alignment = lvwColumnRight
    For AA = 6 To 63
        ListView2.ColumnHeaders(AA).Text = ""
    Next AA
ElseIf opt34 = True And cboClaimCase.ListIndex = 0 Then
    ListView2.ColumnHeaders(1).Text = "Gender"
    ListView2.ColumnHeaders(2).Text = "Age Band"
    ListView2.ColumnHeaders(2).Alignment = lvwColumnCenter
    ListView2.ColumnHeaders(3).Text = "Total Claims"
    ListView2.ColumnHeaders(3).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(4).Text = "Total Duration"
    ListView2.ColumnHeaders(4).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(5).Text = "Total Actual Amt"
    ListView2.ColumnHeaders(5).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(6).Text = "Total Approved Amt"
    ListView2.ColumnHeaders(6).Alignment = lvwColumnRight
        
    For AA = 7 To 63
        ListView2.ColumnHeaders(AA).Text = ""
    Next AA
ElseIf opt33 = True And cboClaimCase.ListIndex = 0 Then
    ListView2.ColumnHeaders(1).Text = "Payor"
    ListView2.ColumnHeaders(2).Text = "Insured Type"
    ListView2.ColumnHeaders(2).Alignment = lvwColumnCenter
    ListView2.ColumnHeaders(3).Text = "Gender"
    ListView2.ColumnHeaders(3).Alignment = lvwColumnCenter
    ListView2.ColumnHeaders(4).Text = "Total Claims"
    ListView2.ColumnHeaders(4).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(5).Text = "Total Duration"
    ListView2.ColumnHeaders(5).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(6).Text = "Total Actual Amt"
    ListView2.ColumnHeaders(6).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(7).Text = "Total Approved Amt"
    ListView2.ColumnHeaders(7).Alignment = lvwColumnRight
        
    For AA = 8 To 63
        ListView2.ColumnHeaders(AA).Text = ""
    Next AA
ElseIf opt29 = True And cboClaimCase.ListIndex = 0 Then
    ListView2.ColumnHeaders(1).Text = "Race Code"
    ListView2.ColumnHeaders(2).Text = "Gender"
    ListView2.ColumnHeaders(2).Alignment = lvwColumnCenter
    ListView2.ColumnHeaders(3).Text = "Total Claims"
    ListView2.ColumnHeaders(3).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(4).Text = "Total Duration"
    ListView2.ColumnHeaders(4).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(5).Text = "Total Actual Amt"
    ListView2.ColumnHeaders(5).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(6).Text = "Total Approved Amt"
    ListView2.ColumnHeaders(6).Alignment = lvwColumnRight
        
    For AA = 7 To 63
        ListView2.ColumnHeaders(AA).Text = ""
    Next AA
ElseIf opt37 = True And cboClaimCase.ListIndex = 0 Then
    ListView2.ColumnHeaders(1).Text = "Insured Type"
    ListView2.ColumnHeaders(2).Text = "Plan Code"
    ListView2.ColumnHeaders(2).Alignment = lvwColumnCenter
    ListView2.ColumnHeaders(3).Text = "Gender"
    ListView2.ColumnHeaders(3).Alignment = lvwColumnCenter
    ListView2.ColumnHeaders(4).Text = "Total Claims"
    ListView2.ColumnHeaders(4).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(5).Text = "Total Duration"
    ListView2.ColumnHeaders(5).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(6).Text = "Total Actual Amt"
    ListView2.ColumnHeaders(6).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(7).Text = "Total Approved Amt"
    ListView2.ColumnHeaders(7).Alignment = lvwColumnRight
        
    For AA = 8 To 63
        ListView2.ColumnHeaders(AA).Text = ""
    Next AA
ElseIf opt30 = True And cboClaimCase.ListIndex = 0 Then
    ListView2.ColumnHeaders(1).Text = "Nature of Claim"
    ListView2.ColumnHeaders(2).Text = "Gender"
    ListView2.ColumnHeaders(2).Alignment = lvwColumnCenter
    ListView2.ColumnHeaders(3).Text = "Total Claims"
    ListView2.ColumnHeaders(3).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(4).Text = "Total Duration"
    ListView2.ColumnHeaders(4).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(5).Text = "Total Actual Amt"
    ListView2.ColumnHeaders(5).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(6).Text = "Total Approved Amt"
    ListView2.ColumnHeaders(6).Alignment = lvwColumnRight
        
    For AA = 7 To 63
        ListView2.ColumnHeaders(AA).Text = ""
    Next AA
ElseIf opt35 = True And cboClaimCase.ListIndex = 1 Then
    ListView2.ColumnHeaders(1).Text = "Sex"
    ListView2.ColumnHeaders(2).Text = "Total Cases"
    ListView2.ColumnHeaders(2).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(3).Text = "Total Duration"
    ListView2.ColumnHeaders(3).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(4).Text = "Total Actual Amt"
    ListView2.ColumnHeaders(4).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(5).Text = "Total Approved Amt"
    ListView2.ColumnHeaders(5).Alignment = lvwColumnRight
    For AA = 6 To 63
        ListView2.ColumnHeaders(AA).Text = ""
    Next AA
ElseIf opt34 = True And cboClaimCase.ListIndex = 1 Then
    ListView2.ColumnHeaders(1).Text = "Gender"
    ListView2.ColumnHeaders(2).Text = "Age Band"
    ListView2.ColumnHeaders(2).Alignment = lvwColumnCenter
    ListView2.ColumnHeaders(3).Text = "Total Cases"
    ListView2.ColumnHeaders(3).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(4).Text = "Total Duration"
    ListView2.ColumnHeaders(4).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(5).Text = "Total Actual Amt"
    ListView2.ColumnHeaders(5).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(6).Text = "Total Approved Amt"
    ListView2.ColumnHeaders(6).Alignment = lvwColumnRight
        
    For AA = 7 To 63
        ListView2.ColumnHeaders(AA).Text = ""
    Next AA
ElseIf opt33 = True And cboClaimCase.ListIndex = 1 Then
    ListView2.ColumnHeaders(1).Text = "Payor"
    ListView2.ColumnHeaders(2).Text = "Insured Type"
    ListView2.ColumnHeaders(2).Alignment = lvwColumnCenter
    ListView2.ColumnHeaders(3).Text = "Gender"
    ListView2.ColumnHeaders(3).Alignment = lvwColumnCenter
    ListView2.ColumnHeaders(4).Text = "Total Cases"
    ListView2.ColumnHeaders(4).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(5).Text = "Total Duration"
    ListView2.ColumnHeaders(5).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(6).Text = "Total Actual Amt"
    ListView2.ColumnHeaders(6).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(7).Text = "Total Approved Amt"
    ListView2.ColumnHeaders(7).Alignment = lvwColumnRight
        
    For AA = 8 To 63
        ListView2.ColumnHeaders(AA).Text = ""
    Next AA
ElseIf opt29 = True And cboClaimCase.ListIndex = 1 Then
    ListView2.ColumnHeaders(1).Text = "Race Code"
    ListView2.ColumnHeaders(2).Text = "Gender"
    ListView2.ColumnHeaders(2).Alignment = lvwColumnCenter
    ListView2.ColumnHeaders(3).Text = "Total Cases"
    ListView2.ColumnHeaders(3).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(4).Text = "Total Duration"
    ListView2.ColumnHeaders(4).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(5).Text = "Total Actual Amt"
    ListView2.ColumnHeaders(5).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(6).Text = "Total Approved Amt"
    ListView2.ColumnHeaders(6).Alignment = lvwColumnRight
        
    For AA = 7 To 63
        ListView2.ColumnHeaders(AA).Text = ""
    Next AA
ElseIf opt37 = True And cboClaimCase.ListIndex = 1 Then
    ListView2.ColumnHeaders(1).Text = "Insured Type"
    ListView2.ColumnHeaders(2).Text = "Plan Code"
    ListView2.ColumnHeaders(2).Alignment = lvwColumnCenter
    ListView2.ColumnHeaders(3).Text = "Gender"
    ListView2.ColumnHeaders(3).Alignment = lvwColumnCenter
    ListView2.ColumnHeaders(4).Text = "Total Cases"
    ListView2.ColumnHeaders(4).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(5).Text = "Total Duration"
    ListView2.ColumnHeaders(5).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(6).Text = "Total Actual Amt"
    ListView2.ColumnHeaders(6).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(7).Text = "Total Approved Amt"
    ListView2.ColumnHeaders(7).Alignment = lvwColumnRight
        
    For AA = 8 To 63
        ListView2.ColumnHeaders(AA).Text = ""
    Next AA
ElseIf opt30 = True And cboClaimCase.ListIndex = 1 Then
    ListView2.ColumnHeaders(1).Text = "Nature of Claim"
    ListView2.ColumnHeaders(2).Text = "Gender"
    ListView2.ColumnHeaders(2).Alignment = lvwColumnCenter
    ListView2.ColumnHeaders(3).Text = "Total Cases"
    ListView2.ColumnHeaders(3).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(4).Text = "Total Duration"
    ListView2.ColumnHeaders(4).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(5).Text = "Total Actual Amt"
    ListView2.ColumnHeaders(5).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(6).Text = "Total Approved Amt"
    ListView2.ColumnHeaders(6).Alignment = lvwColumnRight
        
    For AA = 7 To 63
        ListView2.ColumnHeaders(AA).Text = ""
    Next AA
End If
End Sub

Private Sub ListViewHeader23()
If cboClaimCase.ListIndex = 0 And opt35 = True Then
    ListView2.ColumnHeaders(1).Text = "Race Code"
    ListView2.ColumnHeaders(2).Text = "Total Claims"
    ListView2.ColumnHeaders(2).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(3).Text = "Total Duration"
    ListView2.ColumnHeaders(3).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(4).Text = "Total Actual Amt"
    ListView2.ColumnHeaders(4).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(5).Text = "Total Approved Amt"
    ListView2.ColumnHeaders(5).Alignment = lvwColumnRight
        
    For AA = 6 To 63
        ListView2.ColumnHeaders(AA).Text = ""
    Next AA

ElseIf cboClaimCase.ListIndex = 0 And opt34 = True Then
    ListView2.ColumnHeaders(1).Text = "Race Code"
    ListView2.ColumnHeaders(2).Text = "Age Band"
    ListView2.ColumnHeaders(2).Alignment = lvwColumnCenter
    ListView2.ColumnHeaders(3).Text = "Total Claims"
    ListView2.ColumnHeaders(3).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(4).Text = "Total Duration"
    ListView2.ColumnHeaders(4).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(5).Text = "Total Actual Amt"
    ListView2.ColumnHeaders(5).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(6).Text = "Total Approved Amt"
    ListView2.ColumnHeaders(6).Alignment = lvwColumnRight
        
    For AA = 7 To 63
        ListView2.ColumnHeaders(AA).Text = ""
    Next AA
ElseIf cboClaimCase.ListIndex = 0 And opt33 = True Then
    ListView2.ColumnHeaders(1).Text = "Payor"
    ListView2.ColumnHeaders(2).Text = "Insured Type"
    ListView2.ColumnHeaders(2).Alignment = lvwColumnCenter
    ListView2.ColumnHeaders(3).Text = "Race Code"
    ListView2.ColumnHeaders(3).Alignment = lvwColumnCenter
    ListView2.ColumnHeaders(4).Text = "Total Claims"
    ListView2.ColumnHeaders(4).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(5).Text = "Total Duration"
    ListView2.ColumnHeaders(5).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(6).Text = "Total Actual Amt"
    ListView2.ColumnHeaders(6).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(7).Text = "Total Approved Amt"
    ListView2.ColumnHeaders(7).Alignment = lvwColumnRight
        
    For AA = 8 To 63
        ListView2.ColumnHeaders(AA).Text = ""
    Next AA
    
ElseIf cboClaimCase.ListIndex = 0 And opt29 = True Then
    ListView2.ColumnHeaders(1).Text = "Race Code"
    ListView2.ColumnHeaders(2).Text = "Sex"
    ListView2.ColumnHeaders(2).Alignment = lvwColumnCenter
    ListView2.ColumnHeaders(3).Text = "Total Claims"
    ListView2.ColumnHeaders(3).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(4).Text = "Total Duration"
    ListView2.ColumnHeaders(4).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(5).Text = "Total Actual Amt"
    ListView2.ColumnHeaders(5).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(6).Text = "Total Approved Amt"
    ListView2.ColumnHeaders(6).Alignment = lvwColumnRight
        
    For AA = 7 To 63
        ListView2.ColumnHeaders(AA).Text = ""
    Next AA
    
ElseIf cboClaimCase.ListIndex = 0 And opt37 = True Then
    ListView2.ColumnHeaders(1).Text = "Insured Type"
    ListView2.ColumnHeaders(2).Text = "Plan No"
    ListView2.ColumnHeaders(2).Alignment = lvwColumnCenter
    ListView2.ColumnHeaders(3).Text = "Race Code"
    ListView2.ColumnHeaders(3).Alignment = lvwColumnCenter
    ListView2.ColumnHeaders(4).Text = "Total Claims"
    ListView2.ColumnHeaders(4).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(5).Text = "Total Duration"
    ListView2.ColumnHeaders(5).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(6).Text = "Total Actual Amt"
    ListView2.ColumnHeaders(6).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(7).Text = "Total Approved Amt"
    ListView2.ColumnHeaders(7).Alignment = lvwColumnRight
        
    For AA = 8 To 63
        ListView2.ColumnHeaders(AA).Text = ""
    Next AA
    
ElseIf cboClaimCase.ListIndex = 0 And opt30 = True Then
    ListView2.ColumnHeaders(1).Text = "Nature of Claim"
    ListView2.ColumnHeaders(2).Text = "Race Code"
    ListView2.ColumnHeaders(2).Alignment = lvwColumnCenter
    ListView2.ColumnHeaders(3).Text = "Total Claims"
    ListView2.ColumnHeaders(3).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(4).Text = "Total Duration"
    ListView2.ColumnHeaders(4).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(5).Text = "Total Actual Amt"
    ListView2.ColumnHeaders(5).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(6).Text = "Total Approved Amt"
    ListView2.ColumnHeaders(6).Alignment = lvwColumnRight
        
    For AA = 7 To 63
        ListView2.ColumnHeaders(AA).Text = ""
    Next AA

ElseIf cboClaimCase.ListIndex = 1 And opt35 = True Then
    ListView2.ColumnHeaders(1).Text = "Race Code"
    ListView2.ColumnHeaders(2).Text = "Total Cases"
    ListView2.ColumnHeaders(2).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(3).Text = "Total Duration"
    ListView2.ColumnHeaders(3).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(4).Text = "Total Actual Amt"
    ListView2.ColumnHeaders(4).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(5).Text = "Total Approved Amt"
    ListView2.ColumnHeaders(5).Alignment = lvwColumnRight
        
    For AA = 6 To 63
        ListView2.ColumnHeaders(AA).Text = ""
    Next AA

ElseIf cboClaimCase.ListIndex = 1 And opt34 = True Then
    ListView2.ColumnHeaders(1).Text = "Race Code"
    ListView2.ColumnHeaders(2).Text = "Age Band"
    ListView2.ColumnHeaders(2).Alignment = lvwColumnCenter
    ListView2.ColumnHeaders(3).Text = "Total Cases"
    ListView2.ColumnHeaders(3).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(4).Text = "Total Duration"
    ListView2.ColumnHeaders(4).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(5).Text = "Total Actual Amt"
    ListView2.ColumnHeaders(5).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(6).Text = "Total Approved Amt"
    ListView2.ColumnHeaders(6).Alignment = lvwColumnRight
        
    For AA = 7 To 63
        ListView2.ColumnHeaders(AA).Text = ""
    Next AA
ElseIf cboClaimCase.ListIndex = 1 And opt33 = True Then
    ListView2.ColumnHeaders(1).Text = "Payor"
    ListView2.ColumnHeaders(2).Text = "Insured Type"
    ListView2.ColumnHeaders(2).Alignment = lvwColumnCenter
    ListView2.ColumnHeaders(3).Text = "Race Code"
    ListView2.ColumnHeaders(3).Alignment = lvwColumnCenter
    ListView2.ColumnHeaders(4).Text = "Total Cases"
    ListView2.ColumnHeaders(4).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(5).Text = "Total Duration"
    ListView2.ColumnHeaders(5).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(6).Text = "Total Actual Amt"
    ListView2.ColumnHeaders(6).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(7).Text = "Total Approved Amt"
    ListView2.ColumnHeaders(7).Alignment = lvwColumnRight
        
    For AA = 8 To 63
        ListView2.ColumnHeaders(AA).Text = ""
    Next AA
    
ElseIf cboClaimCase.ListIndex = 1 And opt29 = True Then
    ListView2.ColumnHeaders(1).Text = "Race Code"
    ListView2.ColumnHeaders(2).Text = "Sex"
    ListView2.ColumnHeaders(2).Alignment = lvwColumnCenter
    ListView2.ColumnHeaders(3).Text = "Total Cases"
    ListView2.ColumnHeaders(3).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(4).Text = "Total Duration"
    ListView2.ColumnHeaders(4).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(5).Text = "Total Actual Amt"
    ListView2.ColumnHeaders(5).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(6).Text = "Total Approved Amt"
    ListView2.ColumnHeaders(6).Alignment = lvwColumnRight
        
    For AA = 7 To 63
        ListView2.ColumnHeaders(AA).Text = ""
    Next AA
    
ElseIf cboClaimCase.ListIndex = 1 And opt37 = True Then
    ListView2.ColumnHeaders(1).Text = "Insured Type"
    ListView2.ColumnHeaders(2).Text = "Plan No"
    ListView2.ColumnHeaders(2).Alignment = lvwColumnCenter
    ListView2.ColumnHeaders(3).Text = "Race Code"
    ListView2.ColumnHeaders(3).Alignment = lvwColumnCenter
    ListView2.ColumnHeaders(4).Text = "Total Cases"
    ListView2.ColumnHeaders(4).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(5).Text = "Total Duration"
    ListView2.ColumnHeaders(5).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(6).Text = "Total Actual Amt"
    ListView2.ColumnHeaders(6).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(7).Text = "Total Approved Amt"
    ListView2.ColumnHeaders(7).Alignment = lvwColumnRight
        
    For AA = 8 To 63
        ListView2.ColumnHeaders(AA).Text = ""
    Next AA
    
ElseIf cboClaimCase.ListIndex = 1 And opt30 = True Then
    ListView2.ColumnHeaders(1).Text = "Nature of Claim"
    ListView2.ColumnHeaders(2).Text = "Race Code"
    ListView2.ColumnHeaders(2).Alignment = lvwColumnCenter
    ListView2.ColumnHeaders(3).Text = "Total Cases"
    ListView2.ColumnHeaders(3).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(4).Text = "Total Duration"
    ListView2.ColumnHeaders(4).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(5).Text = "Total Actual Amt"
    ListView2.ColumnHeaders(5).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(6).Text = "Total Approved Amt"
    ListView2.ColumnHeaders(6).Alignment = lvwColumnRight
        
    For AA = 7 To 63
        ListView2.ColumnHeaders(AA).Text = ""
    Next AA
End If
End Sub

Private Sub ListViewHeader18()
If cboClaimCase.ListIndex = 0 And opt35 = True Then
    
    ListView2.ColumnHeaders(1).Text = "Payor"
    ListView2.ColumnHeaders(2).Text = "Ins Type"
    ListView2.ColumnHeaders(2).Alignment = lvwColumnCenter
    ListView2.ColumnHeaders(3).Text = "Total Claims"
    ListView2.ColumnHeaders(3).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(4).Text = "Total Duration"
    ListView2.ColumnHeaders(4).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(5).Text = "Total Actual Amt"
    ListView2.ColumnHeaders(5).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(6).Text = "Total Approved Amt"
    ListView2.ColumnHeaders(6).Alignment = lvwColumnRight
    For AA = 7 To 63
        ListView2.ColumnHeaders(AA).Text = ""
    Next AA

ElseIf cboClaimCase.ListIndex = 0 And opt34 = True Then
    ListView2.ColumnHeaders(1).Text = "Payor"
    ListView2.ColumnHeaders(2).Text = "Insured Type"
    ListView2.ColumnHeaders(2).Alignment = lvwColumnCenter
    ListView2.ColumnHeaders(3).Text = "Plan Type"
    ListView2.ColumnHeaders(3).Alignment = lvwColumnCenter
    ListView2.ColumnHeaders(4).Text = "No of Claims"
    ListView2.ColumnHeaders(4).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(5).Text = "Total Duration "
    ListView2.ColumnHeaders(5).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(6).Text = "Total Actual Amt"
    ListView2.ColumnHeaders(6).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(7).Text = "Total Approved Amt"
    ListView2.ColumnHeaders(7).Alignment = lvwColumnRight
    
    For AA = 8 To 63
        ListView2.ColumnHeaders(AA).Text = ""
    Next AA

ElseIf cboClaimCase.ListIndex = 0 And opt33 = True Then
    ListView2.ColumnHeaders(1).Text = "Payor"
    ListView2.ColumnHeaders(2).Text = "Insured Type"
    ListView2.ColumnHeaders(2).Alignment = lvwColumnCenter
    ListView2.ColumnHeaders(3).Text = "Ageband"
    ListView2.ColumnHeaders(3).Alignment = lvwColumnCenter
    ListView2.ColumnHeaders(4).Text = "Total Claims"
    ListView2.ColumnHeaders(4).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(5).Text = "Total Duration"
    ListView2.ColumnHeaders(5).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(6).Text = "Total Actual Amt"
    ListView2.ColumnHeaders(6).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(7).Text = "Total Approved Amt"
    ListView2.ColumnHeaders(7).Alignment = lvwColumnRight
        
    For AA = 8 To 63
        ListView2.ColumnHeaders(AA).Text = ""
    Next AA
ElseIf cboClaimCase.ListIndex = 0 And opt29 = True Then
    ListView2.ColumnHeaders(1).Text = "Payor"
    ListView2.ColumnHeaders(2).Text = "Insured Type"
    ListView2.ColumnHeaders(2).Alignment = lvwColumnCenter
    ListView2.ColumnHeaders(3).Text = "Gender"
    ListView2.ColumnHeaders(3).Alignment = lvwColumnCenter
    ListView2.ColumnHeaders(4).Text = "Total Claims"
    ListView2.ColumnHeaders(4).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(5).Text = "Total Duration"
    ListView2.ColumnHeaders(5).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(6).Text = "Total Actual Amt"
    ListView2.ColumnHeaders(6).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(7).Text = "Total Approved Amt"
    ListView2.ColumnHeaders(7).Alignment = lvwColumnRight
        
    For AA = 8 To 63
        ListView2.ColumnHeaders(AA).Text = ""
    Next AA
    
ElseIf cboClaimCase.ListIndex = 0 And opt37 = True Then
    ListView2.ColumnHeaders(1).Text = "Payor"
    ListView2.ColumnHeaders(2).Text = "Insured Type"
    ListView2.ColumnHeaders(2).Alignment = lvwColumnCenter
    ListView2.ColumnHeaders(3).Text = "Race Code"
    ListView2.ColumnHeaders(3).Alignment = lvwColumnCenter
    ListView2.ColumnHeaders(4).Text = "Total Claims"
    ListView2.ColumnHeaders(4).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(5).Text = "Total Duration"
    ListView2.ColumnHeaders(5).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(6).Text = "Total Actual Amt"
    ListView2.ColumnHeaders(6).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(7).Text = "Total Approved Amt"
    ListView2.ColumnHeaders(7).Alignment = lvwColumnRight
        
    For AA = 8 To 63
        ListView2.ColumnHeaders(AA).Text = ""
    Next AA
    
ElseIf cboClaimCase.ListIndex = 1 And opt35 = True Then
    
    ListView2.ColumnHeaders(1).Text = "Payor"
    ListView2.ColumnHeaders(2).Text = "Ins Type"
    ListView2.ColumnHeaders(2).Alignment = lvwColumnCenter
    ListView2.ColumnHeaders(3).Text = "Total Cases"
    ListView2.ColumnHeaders(3).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(4).Text = "Total Duration"
    ListView2.ColumnHeaders(4).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(5).Text = "Total Actual Amt"
    ListView2.ColumnHeaders(5).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(6).Text = "Total Approved Amt"
    ListView2.ColumnHeaders(6).Alignment = lvwColumnRight
    For AA = 7 To 63
        ListView2.ColumnHeaders(AA).Text = ""
    Next AA

ElseIf cboClaimCase.ListIndex = 1 And opt34 = True Then
    ListView2.ColumnHeaders(1).Text = "Payor"
    ListView2.ColumnHeaders(2).Text = "Insured Type"
    ListView2.ColumnHeaders(2).Alignment = lvwColumnCenter
    ListView2.ColumnHeaders(3).Text = "Plan Type"
    ListView2.ColumnHeaders(3).Alignment = lvwColumnCenter
    ListView2.ColumnHeaders(4).Text = "No of Cases"
    ListView2.ColumnHeaders(4).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(5).Text = "Total Duration "
    ListView2.ColumnHeaders(5).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(6).Text = "Total Actual Amt"
    ListView2.ColumnHeaders(6).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(7).Text = "Total Approved Amt"
    ListView2.ColumnHeaders(7).Alignment = lvwColumnRight
    
    For AA = 8 To 63
        ListView2.ColumnHeaders(AA).Text = ""
    Next AA

ElseIf cboClaimCase.ListIndex = 1 And opt33 = True Then
    ListView2.ColumnHeaders(1).Text = "Payor"
    ListView2.ColumnHeaders(2).Text = "Insured Type"
    ListView2.ColumnHeaders(2).Alignment = lvwColumnCenter
    ListView2.ColumnHeaders(3).Text = "Ageband"
    ListView2.ColumnHeaders(3).Alignment = lvwColumnCenter
    ListView2.ColumnHeaders(4).Text = "Total Cases"
    ListView2.ColumnHeaders(4).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(5).Text = "Total Duration"
    ListView2.ColumnHeaders(5).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(6).Text = "Total Actual Amt"
    ListView2.ColumnHeaders(6).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(7).Text = "Total Approved Amt"
    ListView2.ColumnHeaders(7).Alignment = lvwColumnRight
        
    For AA = 8 To 63
        ListView2.ColumnHeaders(AA).Text = ""
    Next AA
ElseIf cboClaimCase.ListIndex = 1 And opt29 = True Then
    ListView2.ColumnHeaders(1).Text = "Payor"
    ListView2.ColumnHeaders(2).Text = "Insured Type"
    ListView2.ColumnHeaders(2).Alignment = lvwColumnCenter
    ListView2.ColumnHeaders(3).Text = "Gender"
    ListView2.ColumnHeaders(3).Alignment = lvwColumnCenter
    ListView2.ColumnHeaders(4).Text = "Total Cases"
    ListView2.ColumnHeaders(4).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(5).Text = "Total Duration"
    ListView2.ColumnHeaders(5).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(6).Text = "Total Actual Amt"
    ListView2.ColumnHeaders(6).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(7).Text = "Total Approved Amt"
    ListView2.ColumnHeaders(7).Alignment = lvwColumnRight
        
    For AA = 8 To 63
        ListView2.ColumnHeaders(AA).Text = ""
    Next AA
    
ElseIf cboClaimCase.ListIndex = 1 And opt37 = True Then
    ListView2.ColumnHeaders(1).Text = "Payor"
    ListView2.ColumnHeaders(2).Text = "Insured Type"
    ListView2.ColumnHeaders(2).Alignment = lvwColumnCenter
    ListView2.ColumnHeaders(3).Text = "Race Code"
    ListView2.ColumnHeaders(3).Alignment = lvwColumnCenter
    ListView2.ColumnHeaders(4).Text = "Total Cases"
    ListView2.ColumnHeaders(4).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(5).Text = "Total Duration"
    ListView2.ColumnHeaders(5).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(6).Text = "Total Actual Amt"
    ListView2.ColumnHeaders(6).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(7).Text = "Total Approved Amt"
    ListView2.ColumnHeaders(7).Alignment = lvwColumnRight
        
    For AA = 8 To 63
        ListView2.ColumnHeaders(AA).Text = ""
    Next AA
End If

End Sub

Private Sub ListViewHeader21()
If cboClaimCase.ListIndex = 0 And opt35 = True Then
    ListView2.ColumnHeaders(1).Text = "Payor"
    ListView2.ColumnHeaders(2).Text = "Plan Type"
    ListView2.ColumnHeaders(2).Alignment = lvwColumnCenter
    ListView2.ColumnHeaders(3).Text = "Total Claims"
    ListView2.ColumnHeaders(3).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(4).Text = "Total Duration"
    ListView2.ColumnHeaders(4).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(5).Text = "Total Actual Amt"
    ListView2.ColumnHeaders(5).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(6).Text = "Total Approved Amt"
    ListView2.ColumnHeaders(6).Alignment = lvwColumnRight
    
    For AA = 7 To 63
        ListView2.ColumnHeaders(AA).Text = ""
    Next AA

ElseIf cboClaimCase.ListIndex = 0 And opt31 = True Then
    ListView2.ColumnHeaders(1).Text = "Payor"
    ListView2.ColumnHeaders(2).Text = "Plan Type"
    ListView2.ColumnHeaders(2).Alignment = lvwColumnCenter
    ListView2.ColumnHeaders(3).Text = "Total Claims"
    ListView2.ColumnHeaders(3).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(4).Text = "Total Duration"
    ListView2.ColumnHeaders(4).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(5).Text = "Total Actual Amt"
    ListView2.ColumnHeaders(5).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(6).Text = "Total Approved Amt"
    ListView2.ColumnHeaders(6).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(7).Text = "Gross Premium"
    ListView2.ColumnHeaders(7).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(8).Text = "Net Premium"
    ListView2.ColumnHeaders(8).Alignment = lvwColumnRight
    
    For AA = 9 To 63
        ListView2.ColumnHeaders(AA).Text = ""
    Next AA
    
ElseIf cboClaimCase.ListIndex = 0 And opt34 = True Then
    ListView2.ColumnHeaders(1).Text = "Payor"
    ListView2.ColumnHeaders(2).Text = "Plan Type"
    ListView2.ColumnHeaders(2).Alignment = lvwColumnCenter
    ListView2.ColumnHeaders(3).Text = "Hospital Name"
    ListView2.ColumnHeaders(3).Alignment = lvwColumnLeft
    ListView2.ColumnHeaders(4).Text = "No of Claims"
    ListView2.ColumnHeaders(4).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(5).Text = "Total Duration"
    ListView2.ColumnHeaders(5).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(6).Text = "Total Actual Amt"
    ListView2.ColumnHeaders(6).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(7).Text = "Total Approved Amt"
    ListView2.ColumnHeaders(7).Alignment = lvwColumnRight
    
    For AA = 8 To 63
        ListView2.ColumnHeaders(AA).Text = ""
    Next AA
    
ElseIf cboClaimCase.ListIndex = 0 And opt33 = True Then
    ListView2.ColumnHeaders(1).Text = "Payor"
    ListView2.ColumnHeaders(2).Text = "Plan Type"
    ListView2.ColumnHeaders(2).Alignment = lvwColumnCenter
    ListView2.ColumnHeaders(3).Text = "Final Diag. Code"
    ListView2.ColumnHeaders(3).Alignment = lvwColumnCenter
    ListView2.ColumnHeaders(4).Text = "Descriptions"
    ListView2.ColumnHeaders(4).Alignment = lvwColumnLeft
    ListView2.ColumnHeaders(5).Text = "No of Claims"
    ListView2.ColumnHeaders(5).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(6).Text = "Total Duration "
    ListView2.ColumnHeaders(6).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(7).Text = "Total Actual Amt"
    ListView2.ColumnHeaders(7).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(8).Text = "Total Approved Amt"
    ListView2.ColumnHeaders(8).Alignment = lvwColumnRight
    
    For AA = 9 To 63
        ListView2.ColumnHeaders(AA).Text = ""
    Next AA
    
ElseIf cboClaimCase.ListIndex = 0 And opt29 = True Then
    ListView2.ColumnHeaders(1).Text = "Payor"
    ListView2.ColumnHeaders(2).Text = "Plan Type"
    ListView2.ColumnHeaders(2).Alignment = lvwColumnCenter
    ListView2.ColumnHeaders(3).Text = "Ageband"
    ListView2.ColumnHeaders(3).Alignment = lvwColumnCenter
    ListView2.ColumnHeaders(4).Text = "No of Claims"
    ListView2.ColumnHeaders(4).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(5).Text = "Total Duration "
    ListView2.ColumnHeaders(5).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(6).Text = "Total Actual Amt"
    ListView2.ColumnHeaders(6).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(7).Text = "Total Approved Amt"
    ListView2.ColumnHeaders(7).Alignment = lvwColumnRight
    
    For AA = 8 To 63
        ListView2.ColumnHeaders(AA).Text = ""
    Next AA
    
ElseIf cboClaimCase.ListIndex = 0 And opt37 = True Then
    ListView2.ColumnHeaders(1).Text = "Payor"
    ListView2.ColumnHeaders(2).Text = "Insured Type"
    ListView2.ColumnHeaders(2).Alignment = lvwColumnCenter
    ListView2.ColumnHeaders(3).Text = "Plan Type"
    ListView2.ColumnHeaders(3).Alignment = lvwColumnCenter
    ListView2.ColumnHeaders(4).Text = "No of Claims"
    ListView2.ColumnHeaders(4).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(5).Text = "Total Duration "
    ListView2.ColumnHeaders(5).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(6).Text = "Total Actual Amt"
    ListView2.ColumnHeaders(6).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(7).Text = "Total Approved Amt"
    ListView2.ColumnHeaders(7).Alignment = lvwColumnRight
    
    For AA = 8 To 63
        ListView2.ColumnHeaders(AA).Text = ""
    Next AA

ElseIf cboClaimCase.ListIndex = 0 And opt30 = True Then
    ListView2.ColumnHeaders(1).Text = "Payor"
    ListView2.ColumnHeaders(2).Text = "Insured Type"
    ListView2.ColumnHeaders(2).Alignment = lvwColumnCenter
    ListView2.ColumnHeaders(3).Text = "Plan Type"
    ListView2.ColumnHeaders(3).Alignment = lvwColumnCenter
    ListView2.ColumnHeaders(4).Text = "Ageband"
    ListView2.ColumnHeaders(4).Alignment = lvwColumnCenter
    ListView2.ColumnHeaders(5).Text = "No of Claims"
    ListView2.ColumnHeaders(5).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(6).Text = "Total Duration "
    ListView2.ColumnHeaders(6).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(7).Text = "Total Actual Amt"
    ListView2.ColumnHeaders(7).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(8).Text = "Total Approved Amt"
    ListView2.ColumnHeaders(8).Alignment = lvwColumnRight
    
    For AA = 9 To 63
        ListView2.ColumnHeaders(AA).Text = ""
    Next AA

ElseIf cboClaimCase.ListIndex = 1 And opt35 = True Then
    ListView2.ColumnHeaders(1).Text = "Payor"
    ListView2.ColumnHeaders(2).Text = "Plan Type"
    ListView2.ColumnHeaders(2).Alignment = lvwColumnCenter
    ListView2.ColumnHeaders(3).Text = "Total Cases"
    ListView2.ColumnHeaders(3).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(4).Text = "Total Duration"
    ListView2.ColumnHeaders(4).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(5).Text = "Total Actual Amt"
    ListView2.ColumnHeaders(5).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(6).Text = "Total Approved Amt"
    ListView2.ColumnHeaders(6).Alignment = lvwColumnRight
    
    For AA = 7 To 63
        ListView2.ColumnHeaders(AA).Text = ""
    Next AA

ElseIf cboClaimCase.ListIndex = 1 And opt31 = True Then
    ListView2.ColumnHeaders(1).Text = "Payor"
    ListView2.ColumnHeaders(2).Text = "Plan Type"
    ListView2.ColumnHeaders(2).Alignment = lvwColumnCenter
    ListView2.ColumnHeaders(3).Text = "Total Cases"
    ListView2.ColumnHeaders(3).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(4).Text = "Total Duration"
    ListView2.ColumnHeaders(4).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(5).Text = "Total Actual Amt"
    ListView2.ColumnHeaders(5).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(6).Text = "Total Approved Amt"
    ListView2.ColumnHeaders(6).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(7).Text = "Gross Premium"
    ListView2.ColumnHeaders(7).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(8).Text = "Net Premium"
    ListView2.ColumnHeaders(8).Alignment = lvwColumnRight
    
    For AA = 9 To 63
        ListView2.ColumnHeaders(AA).Text = ""
    Next AA
    
ElseIf cboClaimCase.ListIndex = 1 And opt34 = True Then
    ListView2.ColumnHeaders(1).Text = "Payor"
    ListView2.ColumnHeaders(2).Text = "Plan Type"
    ListView2.ColumnHeaders(2).Alignment = lvwColumnCenter
    ListView2.ColumnHeaders(3).Text = "Hospital Name"
    ListView2.ColumnHeaders(3).Alignment = lvwColumnLeft
    ListView2.ColumnHeaders(4).Text = "No of Cases"
    ListView2.ColumnHeaders(4).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(5).Text = "Total Duration"
    ListView2.ColumnHeaders(5).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(6).Text = "Total Actual Amt"
    ListView2.ColumnHeaders(6).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(7).Text = "Total Approved Amt"
    ListView2.ColumnHeaders(7).Alignment = lvwColumnRight
    
    For AA = 8 To 63
        ListView2.ColumnHeaders(AA).Text = ""
    Next AA
    
ElseIf cboClaimCase.ListIndex = 1 And opt33 = True Then
    ListView2.ColumnHeaders(1).Text = "Payor"
    ListView2.ColumnHeaders(2).Text = "Plan Type"
    ListView2.ColumnHeaders(2).Alignment = lvwColumnCenter
    ListView2.ColumnHeaders(3).Text = "Final Diag. Code"
    ListView2.ColumnHeaders(3).Alignment = lvwColumnCenter
    ListView2.ColumnHeaders(4).Text = "Descriptions"
    ListView2.ColumnHeaders(4).Alignment = lvwColumnLeft
    ListView2.ColumnHeaders(5).Text = "No of Cases"
    ListView2.ColumnHeaders(5).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(6).Text = "Total Duration "
    ListView2.ColumnHeaders(6).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(7).Text = "Total Actual Amt"
    ListView2.ColumnHeaders(7).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(8).Text = "Total Approved Amt"
    ListView2.ColumnHeaders(8).Alignment = lvwColumnRight
    
    For AA = 9 To 63
        ListView2.ColumnHeaders(AA).Text = ""
    Next AA
    
ElseIf cboClaimCase.ListIndex = 1 And opt29 = True Then
    ListView2.ColumnHeaders(1).Text = "Payor"
    ListView2.ColumnHeaders(2).Text = "Plan Type"
    ListView2.ColumnHeaders(2).Alignment = lvwColumnCenter
    ListView2.ColumnHeaders(3).Text = "Ageband"
    ListView2.ColumnHeaders(3).Alignment = lvwColumnCenter
    ListView2.ColumnHeaders(4).Text = "No of Cases"
    ListView2.ColumnHeaders(4).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(5).Text = "Total Duration "
    ListView2.ColumnHeaders(5).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(6).Text = "Total Actual Amt"
    ListView2.ColumnHeaders(6).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(7).Text = "Total Approved Amt"
    ListView2.ColumnHeaders(7).Alignment = lvwColumnRight
    
    For AA = 8 To 63
        ListView2.ColumnHeaders(AA).Text = ""
    Next AA
    
    
ElseIf cboClaimCase.ListIndex = 1 And opt37 = True Then
    ListView2.ColumnHeaders(1).Text = "Payor"
    ListView2.ColumnHeaders(2).Text = "Insured Type"
    ListView2.ColumnHeaders(2).Alignment = lvwColumnCenter
    ListView2.ColumnHeaders(3).Text = "Plan Type"
    ListView2.ColumnHeaders(3).Alignment = lvwColumnCenter
    ListView2.ColumnHeaders(4).Text = "No of Cases"
    ListView2.ColumnHeaders(4).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(5).Text = "Total Duration "
    ListView2.ColumnHeaders(5).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(6).Text = "Total Actual Amt"
    ListView2.ColumnHeaders(6).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(7).Text = "Total Approved Amt"
    ListView2.ColumnHeaders(7).Alignment = lvwColumnRight
    
    For AA = 8 To 63
        ListView2.ColumnHeaders(AA).Text = ""
    Next AA

ElseIf cboClaimCase.ListIndex = 1 And opt30 = True Then
    ListView2.ColumnHeaders(1).Text = "Payor"
    ListView2.ColumnHeaders(2).Text = "Insured Type"
    ListView2.ColumnHeaders(2).Alignment = lvwColumnCenter
    ListView2.ColumnHeaders(3).Text = "Plan Type"
    ListView2.ColumnHeaders(3).Alignment = lvwColumnCenter
    ListView2.ColumnHeaders(4).Text = "Ageband"
    ListView2.ColumnHeaders(4).Alignment = lvwColumnCenter
    ListView2.ColumnHeaders(5).Text = "No of Cases"
    ListView2.ColumnHeaders(5).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(6).Text = "Total Duration "
    ListView2.ColumnHeaders(6).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(7).Text = "Total Actual Amt"
    ListView2.ColumnHeaders(7).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(8).Text = "Total Approved Amt"
    ListView2.ColumnHeaders(8).Alignment = lvwColumnRight
    
    For AA = 9 To 63
        ListView2.ColumnHeaders(AA).Text = ""
    Next AA
End If
End Sub

Private Sub ListViewHeader27()
If cboClaimCase.ListIndex = 0 Then
    ListView2.ColumnHeaders(1).Text = "Payor"
    ListView2.ColumnHeaders(2).Text = "State"
    ListView2.ColumnHeaders(2).Alignment = lvwColumnLeft
    ListView2.ColumnHeaders(3).Text = "Total Claims"
    ListView2.ColumnHeaders(3).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(4).Text = "Total Duration"
    ListView2.ColumnHeaders(4).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(5).Text = "Total Actual Amt"
    ListView2.ColumnHeaders(5).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(6).Text = "Total Approved Amt"
    ListView2.ColumnHeaders(6).Alignment = lvwColumnRight
    
    For AA = 7 To 63
        ListView2.ColumnHeaders(AA).Text = ""
    Next AA
ElseIf cboClaimCase.ListIndex = 1 Then
    ListView2.ColumnHeaders(1).Text = "Payor"
    ListView2.ColumnHeaders(2).Text = "State"
    ListView2.ColumnHeaders(2).Alignment = lvwColumnLeft
    ListView2.ColumnHeaders(3).Text = "Total Cases"
    ListView2.ColumnHeaders(3).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(4).Text = "Total Duration"
    ListView2.ColumnHeaders(4).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(5).Text = "Total Actual Amt"
    ListView2.ColumnHeaders(5).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(6).Text = "Total Approved Amt"
    ListView2.ColumnHeaders(6).Alignment = lvwColumnRight
    
    For AA = 7 To 63
        ListView2.ColumnHeaders(AA).Text = ""
    Next AA
End If
End Sub

Private Sub ListViewHeader25()

If opt35 = True And cboClaimCase.ListIndex = 0 Then
    ListView2.ColumnHeaders(1).Text = "Repudiation Code"
    ListView2.ColumnHeaders(2).Text = "Description"
    ListView2.ColumnHeaders(2).Alignment = lvwColumnLeft
    ListView2.ColumnHeaders(3).Text = "No of Claims"
    ListView2.ColumnHeaders(3).Alignment = lvwColumnRight
    
    For AA = 4 To 63
        ListView2.ColumnHeaders(AA).Text = ""
    Next AA
    
ElseIf opt34 = True And cboClaimCase.ListIndex = 0 Then
    ListView2.ColumnHeaders(1).Text = "RepCategory"
    ListView2.ColumnHeaders(2).Text = "Description"
    ListView2.ColumnHeaders(2).Alignment = lvwColumnLeft
    ListView2.ColumnHeaders(3).Text = "No of Claims"
    ListView2.ColumnHeaders(3).Alignment = lvwColumnRight

    For AA = 4 To 63
        ListView2.ColumnHeaders(AA).Text = ""
    Next AA
    
ElseIf opt33 = True And cboClaimCase.ListIndex = 0 Then
    ListView2.ColumnHeaders(1).Text = "First Diag"
    ListView2.ColumnHeaders(2).Text = "Description"
    ListView2.ColumnHeaders(2).Alignment = lvwColumnLeft
    ListView2.ColumnHeaders(3).Text = "No of Claims"
    ListView2.ColumnHeaders(3).Alignment = lvwColumnRight
    
    For AA = 4 To 63
        ListView2.ColumnHeaders(AA).Text = ""
    Next AA
    
ElseIf opt29 = True And cboClaimCase.ListIndex = 0 Then
    ListView2.ColumnHeaders(1).Text = "Final Diag"
    ListView2.ColumnHeaders(2).Text = "Description"
    ListView2.ColumnHeaders(2).Alignment = lvwColumnLeft
    ListView2.ColumnHeaders(3).Text = "No of Claims"
    ListView2.ColumnHeaders(3).Alignment = lvwColumnRight

    For AA = 4 To 63
        ListView2.ColumnHeaders(AA).Text = ""
    Next AA
    
ElseIf opt37 = True And cboClaimCase.ListIndex = 0 Then
    ListView2.ColumnHeaders(1).Text = "Claim No"
    ListView2.ColumnHeaders(2).Text = "Bordx No"
    ListView2.ColumnHeaders(2).Alignment = lvwColumnLeft
    ListView2.ColumnHeaders(3).Text = "Bordx Date"
    ListView2.ColumnHeaders(3).Alignment = lvwColumnCenter
    ListView2.ColumnHeaders(4).Text = "Claimabity"
    ListView2.ColumnHeaders(4).Alignment = lvwColumnCenter
    ListView2.ColumnHeaders(5).Text = "Patient"
    ListView2.ColumnHeaders(5).Alignment = lvwColumnLeft
    ListView2.ColumnHeaders(6).Text = "Repudiate Cat"
    ListView2.ColumnHeaders(6).Alignment = lvwColumnCenter
    ListView2.ColumnHeaders(7).Text = "Description"
    ListView2.ColumnHeaders(7).Alignment = lvwColumnLeft
    ListView2.ColumnHeaders(8).Text = "Repudiate Code"
    ListView2.ColumnHeaders(8).Alignment = lvwColumnCenter
    ListView2.ColumnHeaders(9).Text = "Description"
    ListView2.ColumnHeaders(9).Alignment = lvwColumnLeft
    ListView2.ColumnHeaders(10).Text = "Diag1"
    ListView2.ColumnHeaders(10).Alignment = lvwColumnCenter
    ListView2.ColumnHeaders(11).Text = "Diag2"
    ListView2.ColumnHeaders(11).Alignment = lvwColumnCenter
    ListView2.ColumnHeaders(12).Text = "Hospital"
    ListView2.ColumnHeaders(12).Alignment = lvwColumnLeft
    ListView2.ColumnHeaders(13).Text = "DOA"
    ListView2.ColumnHeaders(13).Alignment = lvwColumnCenter
    ListView2.ColumnHeaders(14).Text = "DOD"
    ListView2.ColumnHeaders(14).Alignment = lvwColumnCenter
    
    For AA = 15 To 63
       ListView2.ColumnHeaders(AA).Text = ""
    Next AA
    
ElseIf opt30 = True And cboClaimCase.ListIndex = 0 Then
    ListView2.ColumnHeaders(1).Text = "Hospital"
    ListView2.ColumnHeaders(2).Text = "No of Claims"
    ListView2.ColumnHeaders(2).Alignment = lvwColumnRight
    
    For AA = 3 To 63
        ListView2.ColumnHeaders(AA).Text = ""
    Next AA
    
ElseIf opt31 = True And cboClaimCase.ListIndex = 0 Then
    ListView2.ColumnHeaders(1).Text = "Payor"
    ListView2.ColumnHeaders(2).Text = "No of Claims"
    ListView2.ColumnHeaders(2).Alignment = lvwColumnRight

    For AA = 3 To 63
        ListView2.ColumnHeaders(AA).Text = ""
    Next AA


    

    
ElseIf opt35 = True And cboClaimCase.ListIndex = 1 Then
    ListView2.ColumnHeaders(1).Text = "Repudiation Code"
    ListView2.ColumnHeaders(2).Text = "Description"
    ListView2.ColumnHeaders(2).Alignment = lvwColumnLeft
    ListView2.ColumnHeaders(3).Text = "No of Cases"
    ListView2.ColumnHeaders(3).Alignment = lvwColumnRight
    
    For AA = 4 To 63
        ListView2.ColumnHeaders(AA).Text = ""
    Next AA
    
ElseIf opt34 = True And cboClaimCase.ListIndex = 1 Then
    ListView2.ColumnHeaders(1).Text = "RepCategory"
    ListView2.ColumnHeaders(2).Text = "Description"
    ListView2.ColumnHeaders(2).Alignment = lvwColumnLeft
    ListView2.ColumnHeaders(3).Text = "No of Cases"
    ListView2.ColumnHeaders(3).Alignment = lvwColumnRight

    For AA = 4 To 63
        ListView2.ColumnHeaders(AA).Text = ""
    Next AA
    
ElseIf opt33 = True And cboClaimCase.ListIndex = 1 Then
    ListView2.ColumnHeaders(1).Text = "First Diag"
    ListView2.ColumnHeaders(2).Text = "Description"
    ListView2.ColumnHeaders(2).Alignment = lvwColumnLeft
    ListView2.ColumnHeaders(3).Text = "No of Cases"
    ListView2.ColumnHeaders(3).Alignment = lvwColumnRight
    
    For AA = 4 To 63
        ListView2.ColumnHeaders(AA).Text = ""
    Next AA
    
ElseIf opt29 = True And cboClaimCase.ListIndex = 1 Then
    ListView2.ColumnHeaders(1).Text = "Final Diag"
    ListView2.ColumnHeaders(2).Text = "Description"
    ListView2.ColumnHeaders(2).Alignment = lvwColumnLeft
    ListView2.ColumnHeaders(3).Text = "No of Cases"
    ListView2.ColumnHeaders(3).Alignment = lvwColumnRight

    For AA = 4 To 63
        ListView2.ColumnHeaders(AA).Text = ""
    Next AA
    
ElseIf opt37 = True And cboClaimCase.ListIndex = 1 Then
    ListView2.ColumnHeaders(1).Text = "Case No"
    ListView2.ColumnHeaders(2).Text = "Claimabity"
    ListView2.ColumnHeaders(2).Alignment = lvwColumnCenter
    ListView2.ColumnHeaders(3).Text = "Patient"
    ListView2.ColumnHeaders(3).Alignment = lvwColumnLeft
    ListView2.ColumnHeaders(4).Text = "Repudiate Cat"
    ListView2.ColumnHeaders(4).Alignment = lvwColumnCenter
    ListView2.ColumnHeaders(5).Text = "Description"
    ListView2.ColumnHeaders(5).Alignment = lvwColumnLeft
    ListView2.ColumnHeaders(6).Text = "Repudiate Code"
    ListView2.ColumnHeaders(6).Alignment = lvwColumnCenter
    ListView2.ColumnHeaders(7).Text = "Description"
    ListView2.ColumnHeaders(7).Alignment = lvwColumnLeft
    ListView2.ColumnHeaders(8).Text = "Diag1"
    ListView2.ColumnHeaders(8).Alignment = lvwColumnCenter
    ListView2.ColumnHeaders(9).Text = "Diag2"
    ListView2.ColumnHeaders(9).Alignment = lvwColumnCenter
    ListView2.ColumnHeaders(10).Text = "Hospital"
    ListView2.ColumnHeaders(10).Alignment = lvwColumnLeft
    ListView2.ColumnHeaders(11).Text = "DOA"
    ListView2.ColumnHeaders(11).Alignment = lvwColumnCenter
    ListView2.ColumnHeaders(12).Text = "DOD"
    ListView2.ColumnHeaders(12).Alignment = lvwColumnCenter
    
    For AA = 13 To 63
       ListView2.ColumnHeaders(AA).Text = ""
    Next AA
    
ElseIf opt30 = True And cboClaimCase.ListIndex = 1 Then
    ListView2.ColumnHeaders(1).Text = "Hospital"
    ListView2.ColumnHeaders(2).Text = "No of Cases"
    ListView2.ColumnHeaders(2).Alignment = lvwColumnRight
    
    For AA = 3 To 63
        ListView2.ColumnHeaders(AA).Text = ""
    Next AA
    
ElseIf opt31 = True And cboClaimCase.ListIndex = 1 Then
    ListView2.ColumnHeaders(1).Text = "Payor"
    ListView2.ColumnHeaders(2).Text = "No of Cases"
    ListView2.ColumnHeaders(2).Alignment = lvwColumnRight

    For AA = 3 To 63
        ListView2.ColumnHeaders(AA).Text = ""
    Next AA

    

End If

End Sub

Private Sub ListViewHeader26()

    ListView2.ColumnHeaders(1).Text = "Company Name"
    ListView2.ColumnHeaders(2).Text = "Principal Name"
    ListView2.ColumnHeaders(2).Alignment = lvwColumnLeft
    ListView2.ColumnHeaders(3).Text = "Patient Name"
    ListView2.ColumnHeaders(3).Alignment = lvwColumnLeft
    ListView2.ColumnHeaders(4).Text = "Patient I/C No"
    ListView2.ColumnHeaders(4).Alignment = lvwColumnLeft
    ListView2.ColumnHeaders(5).Text = "DOA"
    ListView2.ColumnHeaders(5).Alignment = lvwColumnCenter
    ListView2.ColumnHeaders(6).Text = "App Amt"
    ListView2.ColumnHeaders(6).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(7).Text = "Act Amt"
    ListView2.ColumnHeaders(7).Alignment = lvwColumnRight
            
    For AA = 8 To 63
       ListView2.ColumnHeaders(AA).Text = ""
    Next AA

End Sub

Private Sub ListViewHeader24()

    If cboClaimCase.ListIndex = 0 And opt35 = True Then
        ListView2.ColumnHeaders(1).Text = "Claim No"
        ListView2.ColumnHeaders(2).Text = "Doctor"
        ListView2.ColumnHeaders(2).Alignment = lvwColumnLeft
        ListView2.ColumnHeaders(3).Text = "Patient Name"
        ListView2.ColumnHeaders(3).Alignment = lvwColumnLeft
        ListView2.ColumnHeaders(4).Text = "Agent Code"
        ListView2.ColumnHeaders(4).Alignment = lvwColumnLeft
        ListView2.ColumnHeaders(5).Text = "Hospital"
        ListView2.ColumnHeaders(5).Alignment = lvwColumnLeft
        ListView2.ColumnHeaders(6).Text = "DOA"
        ListView2.ColumnHeaders(6).Alignment = lvwColumnCenter
        ListView2.ColumnHeaders(7).Text = "DOD"
        ListView2.ColumnHeaders(7).Alignment = lvwColumnCenter
        ListView2.ColumnHeaders(8).Text = "Inv Amt"
        ListView2.ColumnHeaders(8).Alignment = lvwColumnRight
        ListView2.ColumnHeaders(9).Text = "Final Diag"
        ListView2.ColumnHeaders(9).Alignment = lvwColumnCenter
        ListView2.ColumnHeaders(10).Text = "Description"
        ListView2.ColumnHeaders(10).Alignment = lvwColumnLeft
        For AA = 11 To 63
            ListView2.ColumnHeaders(AA).Text = ""
        Next AA
    ElseIf cboClaimCase.ListIndex = 1 And opt35 = True Then
        ListView2.ColumnHeaders(1).Text = "Case No"
        ListView2.ColumnHeaders(2).Text = "Doctor"
        ListView2.ColumnHeaders(2).Alignment = lvwColumnLeft
        ListView2.ColumnHeaders(3).Text = "Patient Name"
        ListView2.ColumnHeaders(3).Alignment = lvwColumnLeft
        ListView2.ColumnHeaders(4).Text = "Agent Code"
        ListView2.ColumnHeaders(4).Alignment = lvwColumnLeft
        ListView2.ColumnHeaders(5).Text = "Hospital"
        ListView2.ColumnHeaders(5).Alignment = lvwColumnLeft
        ListView2.ColumnHeaders(6).Text = "DOA"
        ListView2.ColumnHeaders(6).Alignment = lvwColumnCenter
        ListView2.ColumnHeaders(7).Text = "DOD"
        ListView2.ColumnHeaders(7).Alignment = lvwColumnCenter
        ListView2.ColumnHeaders(8).Text = "Inv Amt"
        ListView2.ColumnHeaders(8).Alignment = lvwColumnRight
        ListView2.ColumnHeaders(9).Text = "Final Diag"
        ListView2.ColumnHeaders(9).Alignment = lvwColumnCenter
        ListView2.ColumnHeaders(10).Text = "Description"
        ListView2.ColumnHeaders(10).Alignment = lvwColumnLeft
        For AA = 11 To 63
            ListView2.ColumnHeaders(AA).Text = ""
        Next AA
    ElseIf cboClaimCase.ListIndex = 0 And opt34 = True Then
        ListView2.ColumnHeaders(1).Text = "Doctor"
        ListView2.ColumnHeaders(2).Text = "Hospital"
        ListView2.ColumnHeaders(2).Alignment = lvwColumnLeft
        ListView2.ColumnHeaders(3).Text = "No of Claims"
        ListView2.ColumnHeaders(3).Alignment = lvwColumnRight
        ListView2.ColumnHeaders(4).Text = "Total Duration"
        ListView2.ColumnHeaders(4).Alignment = lvwColumnRight
        ListView2.ColumnHeaders(5).Text = "Total Actual"
        ListView2.ColumnHeaders(5).Alignment = lvwColumnRight
        ListView2.ColumnHeaders(6).Text = "Total Approved"
        ListView2.ColumnHeaders(6).Alignment = lvwColumnRight
        ListView2.ColumnHeaders(7).Text = "Average Actual"
        ListView2.ColumnHeaders(7).Alignment = lvwColumnRight
        ListView2.ColumnHeaders(8).Text = "Average Approved"
        ListView2.ColumnHeaders(8).Alignment = lvwColumnRight
        
        For AA = 9 To 63
            ListView2.ColumnHeaders(AA).Text = ""
        Next AA
    ElseIf cboClaimCase.ListIndex = 1 And opt34 = True Then
        ListView2.ColumnHeaders(1).Text = "Doctor"
        ListView2.ColumnHeaders(2).Text = "Hospital"
        ListView2.ColumnHeaders(2).Alignment = lvwColumnLeft
        ListView2.ColumnHeaders(3).Text = "No of Cases"
        ListView2.ColumnHeaders(3).Alignment = lvwColumnRight
        ListView2.ColumnHeaders(4).Text = "Total Duration"
        ListView2.ColumnHeaders(4).Alignment = lvwColumnRight
        ListView2.ColumnHeaders(5).Text = "Total Actual"
        ListView2.ColumnHeaders(5).Alignment = lvwColumnRight
        ListView2.ColumnHeaders(6).Text = "Total Approved"
        ListView2.ColumnHeaders(6).Alignment = lvwColumnRight
        ListView2.ColumnHeaders(7).Text = "Average Actual"
        ListView2.ColumnHeaders(7).Alignment = lvwColumnRight
        ListView2.ColumnHeaders(8).Text = "Average Approved"
        ListView2.ColumnHeaders(8).Alignment = lvwColumnRight
        For AA = 9 To 63
            ListView2.ColumnHeaders(AA).Text = ""
        Next AA
    End If

End Sub

Private Sub ListViewHeader20()
If cboClaimCase.ListIndex = 0 Then
    ListView2.ColumnHeaders(1).Text = "Payor"
    ListView2.ColumnHeaders(2).Text = "Total Claims"
    ListView2.ColumnHeaders(2).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(3).Text = "Total Duration"
    ListView2.ColumnHeaders(3).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(4).Text = "Total Actual Amt"
    ListView2.ColumnHeaders(4).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(5).Text = "Total Approved Amt"
    ListView2.ColumnHeaders(5).Alignment = lvwColumnRight
    
    For AA = 6 To 63
        ListView2.ColumnHeaders(AA).Text = ""
    Next AA
ElseIf cboClaimCase.ListIndex = 1 Then
    ListView2.ColumnHeaders(1).Text = "Payor"
    ListView2.ColumnHeaders(2).Text = "Total Cases"
    ListView2.ColumnHeaders(2).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(3).Text = "Total Duration"
    ListView2.ColumnHeaders(3).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(4).Text = "Total Actual Amt"
    ListView2.ColumnHeaders(4).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(5).Text = "Total Approved Amt"
    ListView2.ColumnHeaders(5).Alignment = lvwColumnRight
    
    For AA = 6 To 63
        ListView2.ColumnHeaders(AA).Text = ""
    Next AA
End If

End Sub

Private Sub Form_Load()
    intExit = 0
       ' medixmasterconn.Open "File Name=" & BOSPath & "Conn\mediconnect.UDL"
        Call ComboListing
    'medixmasterconn.Close
   ' Set medixmasterconn = Nothing
    'Call ListViewHeader15
    ListView2.Visible = True
    cboClaimCase.ListIndex = 0
    'opt35 = True
    'Option35.Caption = "Without Remarks"
End Sub

' =================== Command Button Click =============================
Private Sub CmdExit_Click()
    Unload Me
End Sub

Private Sub cmdPrint_Click()
    Screen.MousePointer = vbHourglass
        
    If ListView2.ListItems.Count <> 0 Then
            
        'If cboReportType.ListIndex = 0 And opt35 = true Then 'General Listing
            'Call ExportToExcelAnalysis15(ListView2)
        'ElseIf cboReportType.ListIndex = 0 And opt34 = true Then 'General Listing 2
            Call ExportListViewtoExcel(ListView2)
        'ElseIf cboReportType.ListIndex = 7 And opt35 = true Then 'Plan No 1
            'Call ExportToExcelAnalysis211(ListView2)
        'ElseIf cboReportType.ListIndex = 7 And opt31 = true Then 'Plan No 7
            'Call ExportToExcelAnalysis211(ListView2)
        
        'ElseIf cboReportType.ListIndex = 7 And opt34 = true Then 'Plan No 2
            'Call ExportToExcelAnalysis212(ListView2)
        'ElseIf cboReportType.ListIndex = 7 And opt33 = true Then 'Plan No 3
            'Call ExportToExcelAnalysis213(ListView2)
        'ElseIf cboReportType.ListIndex = 7 And opt29 = true Then 'Plan No 4
            'Call ExportToExcelAnalysis214(ListView2)
        'ElseIf cboReportType.ListIndex = 7 And opt37 = true Then ' Plan No 5
            'Call ExportToExcelAnalysis215(ListView2)
        'ElseIf cboReportType.ListIndex = 7 And opt30 = true Then ' Plan No 6
            'Call ExportToExcelAnalysis215(ListView2)
        'ElseIf cboReportType.ListIndex = 1 And opt35 = true Then 'Ageband 1
            'Call ExportToExcelAnalysis31(ListView2)
        'ElseIf cboReportType.ListIndex = 1 And opt34 = true Then 'Ageband 2
            'Call ExportToExcelAnalysis32(ListView2)
        'ElseIf cboReportType.ListIndex = 1 And opt33 = true Then 'Ageband 3
            'Call ExportToExcelAnalysis33(ListView2)
        'ElseIf cboReportType.ListIndex = 1 And opt29 = true Then 'Ageband 4
            'Call ExportToExcelAnalysis34(ListView2)
        'ElseIf cboReportType.ListIndex = 1 And opt37 = true Then 'Ageband 5
            'Call ExportToExcelAnalysis35(ListView2)
        'ElseIf cboReportType.ListIndex = 1 And opt30 = true Then 'Ageband 6
            'Call ExportToExcelAnalysis36(ListView2)
        'ElseIf cboReportType.ListIndex = 2 And opt35 = true Then 'Diagnosis 1
            'Call ExportToExcelAnalysis151(ListView2)
        'ElseIf cboReportType.ListIndex = 2 And opt34 = true Then 'Diagnosis 2
            'Call ExportToExcelAnalysis152(ListView2)
        'ElseIf cboReportType.ListIndex = 2 And opt33 = true Then 'Diagnosis 2
            'Call ExportToExcelAnalysis153(ListView2)
        'ElseIf cboReportType.ListIndex = 2 And opt29 = true Then 'Diagnosis 2
            'Call ExportToExcelAnalysis153(ListView2)
        'Else
            'Call ExportToExcelAnalysis1(ListView2)
        'End If
    
    Else
        MsgBox "Report cannot be printed without any record.", vbInformation
    End If
    
    Screen.MousePointer = vbDefault
End Sub

Private Sub CmdSearch_Click()
On Error GoTo errRun
Dim StrAppend As String
    StrAppend = ""
    CmdPrint.Enabled = False
    CmdClear.Enabled = False
    cmdExit.Enabled = False
    cmdSearch.Enabled = False
    Screen.MousePointer = vbHourglass
 '   medixmasterconn.Open "File Name=" & BOSPath & "Conn\mediconnect.UDL"
 '   medixmasterconnWS.Open "File Name=" & BOSPath & "Conn\mediconnectWS.UDL"

    
    Call SearchRecords
  '  medixmasterconn.Close
  '  Set medixmasterconn = Nothing
 '   medixmasterconnWS.Close
  '  Set medixmasterconnWS = Nothing
    Screen.MousePointer = vbDefault
    cmdSearch.Enabled = True
    Exit Sub
errRun:
    MsgBox Err.Description
    Screen.MousePointer = vbDefault
    ListView2.ListItems.Clear
    lblCurrent = ""
    CmdPrint.Enabled = True
    CmdClear.Enabled = True
    cmdExit.Enabled = True
    cmdSearch.Enabled = True
    'medixmasterconn.Close
   ' Set medixmasterconn = Nothing
   ' medixmasterconnWS.Close
   ' Set medixmasterconnWS = Nothing
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
            ctl.Value = False
    ElseIf TypeOf ctl Is MaskEdBox Then
        ctl.Text = "__/__/____"
    
    End If
Next ctl
End Function

Private Sub cmdClear_Click()
    Call ClearForm(Me)
    lblCurrent = ""
    lblTotalRecord = ""
    ListView2.ListItems.Clear
End Sub
' =================== Command Button Click / Change ==========================


'============== ComboBox Listing =================================
Private Sub ComboListing()
Dim PAY As New ADODB.Recordset
Dim INS As New ADODB.Recordset
Dim PolStatus As New ADODB.Recordset

    INS.Open "Select INSCode, INSDescription from INS", medixmasterconn, adOpenForwardOnly, adLockBatchOptimistic
    PAY.Open "Select PAYCode, PAYCompanyName from PAY", medixmasterconn, adOpenForwardOnly, adLockBatchOptimistic
    PolStatus.Open "Select pstatuscode ,pstatusDesc from PolicyStatus", medixmasterconn, adOpenForwardOnly, adLockBatchOptimistic
    
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
        .AddItem "0 - Open"
        .AddItem "1 - Open & Close"
        .AddItem "2 - Delete"
        .AddItem "3 - All"
        .Text = "1 - Open & Close"
    End With
    
    CboCasePending.AddItem "N - No"
    CboCasePending.AddItem "Y - Yes"
    
    CboClmPending.AddItem "B - Bordxed"
    CboClmPending.AddItem "C - Cancelled"
    CboClmPending.AddItem "P - Pending"
           
    Do Until PAY.EOF
        cboPayorCode.AddItem PAY!PAYCode & " - " & Trim(PAY!PAYCompanyName)
        PAY.MoveNext
    Loop
    PAY.Close
    Set PAY = Nothing
    
    cboSuppType.AddItem "Both"
    cboSuppType.AddItem "C" & " - " & "Combined Limit"
    cboSuppType.AddItem "S" & " - " & "Separated Limit"
    cboSuppType.Text = "Both"
    
    With CboStayType
        .AddItem "I - In-Patient"
        .AddItem "O - Out-Patient"
    End With
    
    With CboCashType
        .AddItem "C - Cashless"
        .AddItem "R - Reimbursement"
    End With
    
    With CboModality
        .AddItem "B - Both"
        .AddItem "N - Non-Surgical"
        .AddItem "S - Surgical"
    End With
    
    With cboClaimCase
        .AddItem "Claim"
        .AddItem "Case"
    End With

End Sub
'============== ComboBox Listing =================================
Private Sub cboClaimCase_Click()

cboReportType.Clear
cboReportSelection.Clear
cboReportSelection.Visible = True
If cboClaimCase.ListIndex = 0 Then
    Call Option4
    With cboReportType
        .AddItem "General Listing"
        .AddItem "Age Band"
        .AddItem "Diagnosis"
        .AddItem "Hospital"
        .AddItem "Month"
        .AddItem "Insured Type"
        .AddItem "Payor"
        .AddItem "Plan No"
        .AddItem "State"
        .AddItem "Efficiency"
    End With
    cboReportType.ListIndex = 0
    Call rptGeneralListing
    
ElseIf cboClaimCase.ListIndex = 1 Then
    Call Option14
    With cboReportType
        .AddItem "General Listing"
        .AddItem "Age Band"
        .AddItem "Diagnosis"
        .AddItem "Hospital"
        .AddItem "Month"
        .AddItem "Insured Type"
        .AddItem "Payor"
        .AddItem "Plan No"
        .AddItem "State"
    End With
    cboReportType.ListIndex = 0
    Call rptGeneralListing
End If
    
End Sub

Private Sub CboReportType_Click()
Call ClearBoolean
cboReportSelection.Clear
cboReportSelection.Visible = True

If cboReportType = cboReportType.List(0) Then
    Call rptGeneralListing
    cboReportSelection = cboReportSelection.List(0)
ElseIf cboReportType = cboReportType.List(1) Then
    Call rptAgeband
    cboReportSelection = cboReportSelection.List(0)
ElseIf cboReportType = cboReportType.List(2) Then
    Call rptDiagnosis
    cboReportSelection = cboReportSelection.List(0)
ElseIf cboReportType = cboReportType.List(3) Then
    Call rptHospital
    cboReportSelection = cboReportSelection.List(0)
ElseIf cboReportType = cboReportType.List(4) Then
    Call rptMonth
    cboReportSelection = cboReportSelection.List(0)
ElseIf cboReportType = cboReportType.List(5) Then
    Call rptInsuredType
    cboReportSelection = cboReportSelection.List(0)
ElseIf cboReportType = cboReportType.List(6) Then
    Call rptPayor
    cboReportSelection.Visible = False
ElseIf cboReportType = cboReportType.List(7) Then
    Call rptPlanNo
    cboReportSelection = cboReportSelection.List(0)
ElseIf cboReportType = cboReportType.List(8) Then
    Call rptState
    cboReportSelection.Visible = False
ElseIf cboReportType = cboReportType.List(9) Then
    Call rptEfficiency
    cboReportSelection = cboReportSelection.List(0)
End If

End Sub

Private Sub cboReportSelection_Click()
Call ClearBoolean
If cboReportType = "General Listing" Then
    opt35 = True
    Call Option35
ElseIf cboReportType = "Age Band" Then
    If cboReportSelection.ListIndex = 0 Then
        opt35 = True
        Call Option35
    ElseIf cboReportSelection.ListIndex = 1 Then
        opt34 = True
        Call Option34
    ElseIf cboReportSelection.ListIndex = 2 Then
        opt33 = True
        Call Option33
    ElseIf cboReportSelection.ListIndex = 3 Then
        opt29 = True
        Call Option29
    ElseIf cboReportSelection.ListIndex = 4 Then
        opt37 = True
        Call Option37
    ElseIf cboReportSelection.ListIndex = 5 Then
        opt30 = True
        Call Option30
    End If
ElseIf cboReportType = "Diagnosis" Then
    If cboReportSelection.ListIndex = 0 Then
        opt35 = True
        Call Option35
    ElseIf cboReportSelection.ListIndex = 1 Then
        opt34 = True
        Call Option34
    ElseIf cboReportSelection.ListIndex = 2 Then
        opt33 = True
        Call Option33
    ElseIf cboReportSelection.ListIndex = 3 Then
        opt29 = True
        Call Option29
    End If
ElseIf cboReportType = "Hospital" Then
    If cboReportSelection.ListIndex = 0 Then
        opt35 = True
        Call Option35
    ElseIf cboReportSelection.ListIndex = 1 Then
        opt34 = True
        Call Option34
    ElseIf cboReportSelection.ListIndex = 2 Then
        opt33 = True
        Call Option33
    ElseIf cboReportSelection.ListIndex = 3 Then
        opt29 = True
        Call Option29
    ElseIf cboReportSelection.ListIndex = 4 Then
        opt37 = True
        Call Option37
    ElseIf cboReportSelection.ListIndex = 5 Then
        opt31 = True
        Call Option31
    ElseIf cboReportSelection.ListIndex = 6 Then
        opt30 = True
        Call Option30
    'ElseIf cboReportSelection.ListIndex = 7 Then
    '    opt16 = True
    '    Call Option16
    End If
ElseIf cboReportType = "Month" Then
    If cboClaimCase.ListIndex = 0 Then
        If cboReportSelection.ListIndex = 0 Then
            opt35 = True
            Call Option35
        ElseIf cboReportSelection.ListIndex = 1 Then
            opt34 = True
            Call Option34
        ElseIf cboReportSelection.ListIndex = 2 Then
            opt33 = True
            Call Option33
        ElseIf cboReportSelection.ListIndex = 3 Then
            opt29 = True
            Call Option29
        ElseIf cboReportSelection.ListIndex = 4 Then
            opt37 = True
            Call Option37
        ElseIf cboReportSelection.ListIndex = 5 Then
            opt30 = True
            Call Option30
        ElseIf cboReportSelection.ListIndex = 6 Then
            Opt8 = True
            Call Option8
        ElseIf cboReportSelection.ListIndex = 7 Then
            Opt11 = True
            Call Option11
        ElseIf cboReportSelection.ListIndex = 8 Then
            opt12 = True
            Call Option12
        ElseIf cboReportSelection.ListIndex = 9 Then
            opt13 = True
            Call Option13
        ElseIf cboReportSelection.ListIndex = 10 Then
            Opt7 = True
            Call Option7
        ElseIf cboReportSelection.ListIndex = 11 Then
            opt16 = True
            Call Option16
        End If
    ElseIf cboClaimCase.ListIndex = 1 Then
        If cboReportSelection.ListIndex = 0 Then
            opt35 = True
            Call Option35
        ElseIf cboReportSelection.ListIndex = 1 Then
            opt34 = True
            Call Option34
        ElseIf cboReportSelection.ListIndex = 2 Then
            opt33 = True
            Call Option33
        ElseIf cboReportSelection.ListIndex = 3 Then
            opt29 = True
            Call Option29
        ElseIf cboReportSelection.ListIndex = 4 Then
            opt37 = True
            Call Option37
        ElseIf cboReportSelection.ListIndex = 5 Then
            Opt8 = True
            Call Option8
        ElseIf cboReportSelection.ListIndex = 6 Then
            Opt11 = True
            Call Option11
        ElseIf cboReportSelection.ListIndex = 7 Then
            opt12 = True
            Call Option12
        ElseIf cboReportSelection.ListIndex = 8 Then
            opt13 = True
            Call Option13
        ElseIf cboReportSelection.ListIndex = 9 Then
            Opt7 = True
            Call Option7
        End If
    End If
ElseIf cboReportType = "Insured Type" Then
    If cboReportSelection.ListIndex = 0 Then
        opt35 = True
        Call Option35
    ElseIf cboReportSelection.ListIndex = 1 Then
        opt34 = True
        Call Option34
    ElseIf cboReportSelection.ListIndex = 2 Then
        opt33 = True
        Call Option33
    ElseIf cboReportSelection.ListIndex = 3 Then
        opt29 = True
        Call Option29
    ElseIf cboReportSelection.ListIndex = 4 Then
        opt37 = True
        Call Option37
    End If
ElseIf cboReportType = "Plan No" Then
    If cboReportSelection.ListIndex = 0 Then
        opt35 = True
        Call Option35
    ElseIf cboReportSelection.ListIndex = 1 Then
        opt34 = True
        Call Option34
    ElseIf cboReportSelection.ListIndex = 2 Then
        opt33 = True
        Call Option33
    ElseIf cboReportSelection.ListIndex = 3 Then
        opt29 = True
        Call Option29
    ElseIf cboReportSelection.ListIndex = 4 Then
        opt37 = True
        Call Option37
    ElseIf cboReportSelection.ListIndex = 5 Then
        opt31 = True
        Call Option31
    ElseIf cboReportSelection.ListIndex = 6 Then
        opt30 = True
        Call Option30
    ElseIf cboReportSelection.ListIndex = 7 Then
        opt16 = True
        Call Option16
    End If
ElseIf cboReportType = "Efficiency" Then
    If cboReportSelection.ListIndex = 0 Then
        opt35 = True
        Call Option35
    ElseIf cboReportSelection.ListIndex = 1 Then
        opt34 = True
        Call Option34
    ElseIf cboReportSelection.ListIndex = 2 Then
        opt33 = True
        Call Option33
    ElseIf cboReportSelection.ListIndex = 3 Then
        opt29 = True
        Call Option29
    ElseIf cboReportSelection.ListIndex = 4 Then
        opt37 = True
        Call Option37
    End If
End If

End Sub

Private Sub ClearBoolean()
    Opt7 = False
    Opt8 = False
    Opt11 = False
    opt12 = False
    opt13 = False
    opt16 = False
    opt29 = False
    opt30 = False
    opt31 = False
    opt33 = False
    opt34 = False
    opt35 = False
    opt37 = False
End Sub

Private Sub DiagFinalICD(StrAppend1, strAppendCase, Record)
Dim rst3 As New ADODB.Recordset
Dim con As New ADODB.Command
Dim prm1 As ADODB.Parameter
Dim prm2 As ADODB.Parameter

Set con.ActiveConnection = medixmasterconn
    con.CommandText = "DCUReport"
    con.CommandType = adCmdStoredProc
    con.CommandTimeout = 300
    Set prm1 = con.CreateParameter("StrAppend", adVarChar, adParamInput, 2000, StrAppend1)
    con.Parameters.Append prm1
    Set prm2 = con.CreateParameter("StrAppendCase", adVarChar, adParamInput, 2000, strAppendCase)
    con.Parameters.Append prm2
    rst3.CursorLocation = adUseClient
    rst3.CursorType = adOpenForwardOnly
    rst3.CacheSize = 100
    Set rst3 = con.Execute
    
    If rst3.EOF = True Then
        Record.SubItems(1) = ""
        rst3.Close
        Set rst3 = Nothing
    Else
        Description = Trim(rst3!ICDDescription)
        Record.SubItems(1) = Trim(Description)
        rst3.Close
        Set rst3 = Nothing
    End If
    
End Sub

'==================== List Final Diag. Listing Resullts ==========================
Private Sub FinalDiagListing(Optional StrAppend As String)
Dim rst2 As New ADODB.Recordset
Dim cmd As New ADODB.Command
Dim prm1 As ADODB.Parameter
Dim prm2 As ADODB.Parameter
Dim strAppendCase As String
Dim StrAppend1 As String
Dim Record As ListItem
    
    DiagCode = ""
    lblCurrent = 0
    ListView2.ListItems.Clear
            
    If opt35 = True Then
    
        If cboClaimCase.ListIndex = 0 Then
            strAppendCase = "71"
        'sql = " SELECT Top 20 COUNT(CASNumber) AS Total, FinalDiag, Sum(CLMTotalApproved) As TotalApproved, " & _
              " Sum(CLMTotalActual) As TotalActual, " & _
              " SUM(NoofDay) AS NoofDay From VIEW_CLMReportDCU Where 0 = 0 " & StrAppend & _
              " GROUP BY FinalDiag ORDER BY Total DESC"
              
        ElseIf cboClaimCase.ListIndex = 1 Then
            strAppendCase = "72"
        'sql = "SELECT Top 20 COUNT(DISTINCT CASNumber) AS Total, FinalDiag, Sum(CLMTotalApproved) As TotalApproved, " & _
              " Sum(CLMTotalActual) As TotalActual " & _
              " From VIEW_CLMReportDCU Where 0 = 0 " & StrAppend & _
              " GROUP BY FinalDiag ORDER BY Total DESC"
        
        End If
    
    ElseIf opt34 = True Then
        
        If cboClaimCase.ListIndex = 0 Then
            strAppendCase = "73"
        'sql = "SELECT Top 20 Sum(CLMTotalApproved) As TotalApproved, COUNT(CASNumber) AS Total, FinalDiag, " & _
              " Sum(CLMTotalActual) As TotalActual, " & _
              " SUM(NoofDay) AS NoofDay From VIEW_CLMReportDCU Where 0 = 0 " & StrAppend & _
              " Group By FinalDiag Order By TotalApproved Desc "
              
        ElseIf cboClaimCase.ListIndex = 1 Then
            strAppendCase = "74"
        'sql = "SELECT Top 20 Sum(CLMTotalApproved) As TotalApproved, COUNT(DISTINCT CASNumber) AS Total, FinalDiag, " & _
              " Sum(CLMTotalActual) As TotalActual " & _
              " From VIEW_CLMReportDCU Where 0 = 0 " & StrAppend & _
              " Group By FinalDiag Order By TotalApproved Desc "
              
        End If
        
    
    Else
        
        If cboClaimCase.ListIndex = 0 Then
            strAppendCase = "75"
        'sql = "SELECT COUNT(CASNumber) AS Total, FinalDiag, Sum(CLMTotalApproved) As TotalApproved, " & _
              " Sum(CLMTotalActual) As TotalActual, " & _
              " SUM(NoofDay) AS NoofDay From VIEW_CLMReportDCU Where 0 = 0 " & StrAppend & _
              " GROUP BY FinalDiag ORDER BY Total DESC"
              
        ElseIf cboClaimCase.ListIndex = 1 Then
            strAppendCase = "76"
        'sql = "SELECT COUNT(DISTINCT CASNumber) AS Total, FinalDiag, Sum(CLMTotalApproved) As TotalApproved, " & _
              " Sum(CLMTotalActual) As TotalActual " & _
              " From VIEW_CLMReportDCU Where 0 = 0 " & StrAppend & _
              " GROUP BY FinalDiag ORDER BY Total DESC"
        
        End If
    End If
    'rst2.Open sql, medixmasterconnWS, adOpenForwardOnly, adLockBatchOptimistic
    Set cmd.ActiveConnection = medixmasterconnWS
        cmd.CommandText = "DCUReport"
        cmd.CommandType = adCmdStoredProc
        cmd.CommandTimeout = 300
        Set prm1 = cmd.CreateParameter("StrAppend", adVarChar, adParamInput, 2000, StrAppend)
        cmd.Parameters.Append prm1
        Set prm2 = cmd.CreateParameter("StrAppendCase", adVarChar, adParamInput, 2000, strAppendCase)
        cmd.Parameters.Append prm2
        rst2.CursorLocation = adUseClient
        rst2.CursorType = adOpenForwardOnly
        rst2.CacheSize = 100
        Set rst2 = cmd.Execute
        
        Do Until rst2.EOF
    
            With rst2
                
                If Len(rst2!FinalDiag) Then
                    Set Record = ListView2.ListItems.Add(, , rst2!FinalDiag)
                    DiagCode = rst2!FinalDiag
                Else
                    Set Record = ListView2.ListItems.Add(, , "Empty")
                    DiagCode = ""
                End If
                
                'sql2 = " SELECT ICDCode, ICDDescription From ICDOriginal Where ICDCode = '" & Trim(DiagCode) & "'"
                'rst3.Open sql2, medixmasterconn, adOpenForwardOnly, adLockBatchOptimistic
                strAppendCase = "1"
                StrAppend1 = " AND ICDCode = '" & Trim(DiagCode) & "'"
                Call DiagFinalICD(StrAppend1, strAppendCase, Record)
                                
                If Len(rst2!Total) Then
                    Record.SubItems(2) = rst2!Total
                End If
                
                If cboClaimCase.ListIndex = 0 Then
                
                    If rst2!NoofDay = "0" Then
                        Record.SubItems(3) = "1"
                    ElseIf Len(rst2!NoofDay) Then
                        Record.SubItems(3) = rst2!NoofDay
                    End If
                
                ElseIf cboClaimCase.ListIndex = 1 Then
                
                    If DiagCode = "" Then
                       strAppendCase = "77"
                       StrAppend1 = " AND (FinalDiag = '' OR FinalDiag IS NULL)"
                       StrAppend1 = StrAppend1 + StrAppend
                    'sql3 = " SELECT FinalDiag, SUM(DurationStay) AS NoofDay " & _
                           " From VIEW_CLMReportDCU2 Where (FinalDiag = '' OR FinalDiag IS NULL) "
                           
                    Else
                       strAppendCase = "77"
                       StrAppend1 = " AND FinalDiag = '" & Trim(DiagCode) & "'"
                       StrAppend1 = StrAppend1 + StrAppend
                    'sql3 = " SELECT FinalDiag, SUM(DurationStay) AS NoofDay " & _
                           " From VIEW_CLMReportDCU2 Where FinalDiag = '" & Trim(DiagCode) & "'"
                           
                    End If
                       
                       'sql2 = " Group By FinalDiag "
                       
                       'If Len(Trim(StrAppend)) Then
                       '     sql3 = sql3 + StrAppend + sql2
                       ' Else
                       '     sql3 = sql3 + sql2
                       ' End If
                                              
                       'rst1.Open sql3, medixmasterconnWS, adOpenForwardOnly, adLockBatchOptimistic
                        Call DurationStay(StrAppend1, strAppendCase, Record)
                End If
                
                If Len(rst2!TotalActual) Then
                   Record.SubItems(4) = Format(rst2!TotalActual, "##,#0.00")
                End If
                
                If Len(rst2!TotalApproved) Then
                    Record.SubItems(5) = Format(rst2!TotalApproved, "##,#0.00")
                End If
            
            End With
            lblCurrent = lblCurrent + 1
            rst2.MoveNext
        Loop
    If ListView2.ListItems.Count = 0 Then
        MsgBox "No record found!", vbOKOnly + vbCritical, DialogBoxTitle
        CmdPrint.Enabled = True
        CmdClear.Enabled = True
        cmdExit.Enabled = True
    End If
    rst2.Close
    Set rst2 = Nothing
    CmdPrint.Enabled = True
    CmdClear.Enabled = True
    cmdExit.Enabled = True
End Sub

Private Sub DiagHOSFinalICD(StrAppend1, strAppendCase, Record)
Dim rst3 As New ADODB.Recordset
Dim con As New ADODB.Command
Dim prm1 As ADODB.Parameter
Dim prm2 As ADODB.Parameter

Set con.ActiveConnection = medixmasterconn
    con.CommandText = "DCUReport"
    con.CommandType = adCmdStoredProc
    con.CommandTimeout = 300
    Set prm1 = con.CreateParameter("StrAppend", adVarChar, adParamInput, 2000, StrAppend1)
    con.Parameters.Append prm1
    Set prm2 = con.CreateParameter("StrAppendCase", adVarChar, adParamInput, 2000, strAppendCase)
    con.Parameters.Append prm2
    rst3.CursorLocation = adUseClient
    rst3.CursorType = adOpenForwardOnly
    rst3.CacheSize = 100
    Set rst3 = con.Execute
    
    If rst3.EOF = True Then
        Record.SubItems(1) = ""
        rst3.Close
        Set rst3 = Nothing
    Else
        Description = Trim(rst3!ICDDescription)
        Record.SubItems(1) = Trim(Description)
        rst3.Close
        Set rst3 = Nothing
    End If
    
End Sub

Private Sub FinalDiagbyHos(Optional StrAppend As String)
Dim rst2 As New ADODB.Recordset
Dim cmd As New ADODB.Command
Dim prm1 As ADODB.Parameter
Dim prm2 As ADODB.Parameter
Dim strAppendCase As String
Dim StrAppend1 As String
Dim Record As ListItem
    
    DiagCode = ""
    Hospital = ""
    HOspitals = ""
    lblCurrent = 0
    ListView2.ListItems.Clear
    
    If cboClaimCase.ListIndex = 0 Then
        strAppendCase = "78"
    'sql = " SELECT COUNT(CASNumber) AS Total, FinalDiag, Sum(CLMTotalApproved) As TotalApproved, " & _
          " Sum(CLMTotalActual) As TotalActual, HOSName, " & _
          " SUM(NoofDay) AS NoofDay From VIEW_CLMReportDCU Where 0 = 0 " & StrAppend & _
          " GROUP BY FinalDiag, HOSName ORDER BY FinalDiag "
          
    ElseIf cboClaimCase.ListIndex = 1 Then
        strAppendCase = "79"
    'sql = " SELECT COUNT(DISTINCT CASNumber) AS Total, FinalDiag, Sum(CLMTotalApproved) As TotalApproved, " & _
          " Sum(CLMTotalActual) As TotalActual, HOSName " & _
          " From VIEW_CLMReportDCU Where 0 = 0 " & StrAppend & _
          " GROUP BY FinalDiag, HOSName ORDER BY FinalDiag "
          
    End If
    
    'rst2.Open sql, medixmasterconnWS, adOpenForwardOnly, adLockBatchOptimistic
    Set cmd.ActiveConnection = medixmasterconnWS
        cmd.CommandText = "DCUReport"
        cmd.CommandType = adCmdStoredProc
        cmd.CommandTimeout = 300
        Set prm1 = cmd.CreateParameter("StrAppend", adVarChar, adParamInput, 2000, StrAppend)
        cmd.Parameters.Append prm1
        Set prm2 = cmd.CreateParameter("StrAppendCase", adVarChar, adParamInput, 2000, strAppendCase)
        cmd.Parameters.Append prm2
        rst2.CursorLocation = adUseClient
        rst2.CursorType = adOpenForwardOnly
        rst2.CacheSize = 100
        Set rst2 = cmd.Execute
               
        Do Until rst2.EOF
    
            With rst2
                               
                If Len(rst2!FinalDiag) Then
                    Set Record = ListView2.ListItems.Add(, , rst2!FinalDiag)
                    DiagCode = rst2!FinalDiag
                Else
                    Set Record = ListView2.ListItems.Add(, , "Empty")
                    DiagCode = ""
                End If
                
                'sql2 = " SELECT ICDCode, ICDDescription From ICDOriginal Where ICDCode = '" & Trim(DiagCode) & "'"
                'rst3.Open sql2, medixmasterconn, adOpenForwardOnly, adLockBatchOptimistic
                strAppendCase = "1"
                StrAppend1 = " AND ICDCode = '" & Trim(DiagCode) & "'"
                Call DiagHOSFinalICD(StrAppend1, strAppendCase, Record)
                           
                If Len(rst2!HOSName) Then
                    Record.SubItems(2) = Trim(rst2!HOSName)
                    Hospital = rst2!HOSName
                Else
                    Record.SubItems(2) = "Empty"
                    Hospital = ""
                End If
                
                If Len(rst2!Total) Then
                    Record.SubItems(3) = rst2!Total
                End If
                
                If cboClaimCase.ListIndex = 0 Then
                
                    If rst2!NoofDay = "0" Then
                        Record.SubItems(4) = "1"
                    ElseIf Len(rst2!NoofDay) Then
                        Record.SubItems(4) = rst2!NoofDay
                    End If
                
                ElseIf cboClaimCase.ListIndex = 1 Then
                
                    HOspitals = Replace(Hospital, "'", "''")
                    
                    If DiagCode = "" And HOspitals = "" Then
                    
                    'sql3 = " SELECT FinalDiag, HOSName, SUM(DurationStay) AS NoofDay " & _
                           " From VIEW_CLMReportDCU2 Where (FinalDiag = '' OR FinalDiag IS NULL) AND " & _
                           " (HOSName = '' OR HOSName IS NULL)"
                              
                    ElseIf DiagCode <> "" And HOspitals = "" Then
                    
                    'sql3 = " SELECT FinalDiag, HOSName, SUM(DurationStay) AS NoofDay " & _
                           " From VIEW_CLMReportDCU2 Where FinalDiag = '" & Trim(DiagCode) & "' AND " & _
                           " (HOSName = '' OR HOSName IS NULL)"
                              
                    ElseIf DiagCode = "" And HOspitals <> "" Then
                       strAppendCase = "80"
                       StrAppend1 = " AND (FinalDiag = '' OR FinalDiag IS NULL) AND " & _
                           " HOSName = '" & Trim(HOspitals) & "'"
                       StrAppend1 = StrAppend1 + StrAppend
                    'sql3 = " SELECT FinalDiag, HOSName, SUM(DurationStay) AS NoofDay " & _
                           " From VIEW_CLMReportDCU2 Where (FinalDiag = '' OR FinalDiag IS NULL) AND " & _
                           " HOSName = '" & Trim(Hospitals) & "'"
                              
                    Else
                       strAppendCase = "80"
                       StrAppend1 = " AND FinalDiag = '" & Trim(DiagCode) & "' AND " & _
                           " HOSName = '" & Trim(HOspitals) & "'"
                       StrAppend1 = StrAppend1 + StrAppend
                    'sql3 = " SELECT FinalDiag, HOSName, SUM(DurationStay) AS NoofDay " & _
                           " From VIEW_CLMReportDCU2 Where FinalDiag = '" & Trim(DiagCode) & "' AND " & _
                           " HOSName = '" & Trim(Hospitals) & "'"
                              
                    End If
                       
                       'sql2 = " Group By FinalDiag, HOSName ORDER BY FinalDiag "
                       
                       'If Len(Trim(StrAppend)) Then
                       '     sql3 = sql3 + StrAppend + sql2
                       ' Else
                       '     sql3 = sql3 + sql2
                       ' End If
                                              
                       'rst1.Open sql3, medixmasterconnWS, adOpenForwardOnly, adLockBatchOptimistic
                        Call DurationStay(StrAppend1, strAppendCase, Record)
                End If
                
                If Len(rst2!TotalActual) Then
                   Record.SubItems(5) = Format(rst2!TotalActual, "##,#0.00")
                End If
                
                If Len(rst2!TotalApproved) Then
                    Record.SubItems(6) = Format(rst2!TotalApproved, "##,#0.00")
                End If
            
            End With
            lblCurrent = lblCurrent + 1
            rst2.MoveNext
        Loop
    If ListView2.ListItems.Count = 0 Then
        MsgBox "No record found!", vbOKOnly + vbCritical, DialogBoxTitle
        CmdPrint.Enabled = True
        CmdClear.Enabled = True
        cmdExit.Enabled = True
    End If
    rst2.Close
    Set rst2 = Nothing
    CmdPrint.Enabled = True
    CmdClear.Enabled = True
    cmdExit.Enabled = True
End Sub

'=========== Option Button, Combo Box Change/KeyPress ===================

Private Sub cboPayorCode_Click()
Dim StrSql As String
Dim PAYGRP As New ADODB.Recordset
Dim cmd As New ADODB.Command
Dim prm1 As New ADODB.Parameter
Dim StrsqlAppend As String

    If Len(Mid(cboPayorCode, 1, 2)) Then
        If ChkGrpCom2.Value = 1 Then
            CboGrpCom.Clear
            'medixmasterconn.Open "File Name=" & BOSPath & "Conn\mediconnect.UDL"
           ' StrsqlAppend = StrsqlAppend & "And PAYCode = '" & Mid(cboPayorCode, 1, 2) & "' GROUP BY MBMGroupCompany, PAYCode"
            
            Set cmd.ActiveConnection = medixmasterconn
            cmd.CommandText = "SearchGrpCompanyPay"
            cmd.CommandType = adCmdStoredProc
            cmd.CommandTimeout = 15
            Set prm1 = cmd.CreateParameter("StrsqlAppend", adVarChar, adParamInput, 255, StrsqlAppend)
            cmd.Parameters.Append prm1
            PAYGRP.CursorLocation = adUseClient
            PAYGRP.CursorType = adOpenForwardOnly
            PAYGRP.CacheSize = 100
            Set PAYGRP = cmd.Execute
            
            Do Until PAYGRP.EOF
                If Len(PAYGRP!MBMGroupCompany) Then
                    CboGrpCom.AddItem Trim(PAYGRP!MBMGroupCompany)
                End If
                PAYGRP.MoveNext
            Loop
            PAYGRP.Close
            Set PAYGRP = Nothing
          '  medixmasterconn.Close
           ' Set medixmasterconn = Nothing
        End If
    End If
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



Private Sub cboHLTcode_KeyPress(KeyAscii As Integer)
Dim NewText As String
Dim ValidCount As Integer
Dim ValidValue As String
Dim i As Integer

    If KeyAscii >= 32 And KeyAscii <> 127 Then
    NewText = LCase(Left(cboHLTCode.Text, cboHLTCode.SelStart) + Chr(KeyAscii) + Mid(cboHLTCode.Text, cboHLTCode.SelStart + cboHLTCode.SelLength + 1))
 
    ValidCount = 0
    ValidValue = ""
 
    For i = 0 To cboHLTCode.ListCount - 1
 
        If NewText = LCase(Left(cboHLTCode.List(i), Len(NewText))) Then
                ValidCount = ValidCount + 1
                ValidValue = cboHLTCode.List(i)
        End If
    Next
        
        If ValidCount <= 1 Then KeyAscii = 0
            If ValidCount = 1 Then
                cboHLTCode.Text = ValidValue
                cboHLTCode.SelStart = 0
                cboHLTCode.SelLength = Len(ValidValue)
            End If
        End If
End Sub

Private Sub CboInsType_Change()

    If CboInsType.Text = "" Then
        If InStr(CboInsType.Text, "null") Then
           CboInsType.Enabled = False
        End If
    End If
End Sub

Private Sub CboInsType_KeyPress(KeyAscii As Integer)
Dim NewText As String
Dim ValidCount As Integer
Dim ValidValue As String
Dim i As Integer

    If KeyAscii >= 32 And KeyAscii <> 127 Then
    NewText = LCase(Left(CboInsType.Text, CboInsType.SelStart) + Chr(KeyAscii) + Mid(CboInsType.Text, CboInsType.SelStart + CboInsType.SelLength + 1))
 
    ValidCount = 0
    ValidValue = ""
 
    For i = 0 To CboInsType.ListCount - 1
 
        If NewText = LCase(Left(CboInsType.List(i), Len(NewText))) Then
                ValidCount = ValidCount + 1
                ValidValue = CboInsType.List(i)
        End If
    Next
        
        If ValidCount <= 1 Then KeyAscii = 0
            If ValidCount = 1 Then
                CboInsType.Text = ValidValue
                CboInsType.SelStart = 0
                CboInsType.SelLength = Len(ValidValue)
            End If
        End If
End Sub

Private Sub CboMemType_Change()

    If CboMemType.Text = "" Then
        If InStr(CboMemType.Text, "null") Then
           CboMemType.Enabled = False
        End If
    End If
End Sub

Private Sub cboSuppType_Change()

    If cboSuppType.Text = "" Then
        If InStr(cboSuppType.Text, "null") Then
           cboSuppType.Enabled = False
        End If
    End If
End Sub

Private Sub cboDataStatus_Change()

    If cboDataStatus.Text = "" Then
        If InStr(cboDataStatus.Text, "null") Then
           cboDataStatus.Enabled = False
        End If
    End If
End Sub

Private Sub cboDataStatus_KeyPress(KeyAscii As Integer)
Dim NewText As String
Dim ValidCount As Integer
Dim ValidValue As String
Dim i As Integer

    If KeyAscii >= 32 And KeyAscii <> 127 Then
    NewText = LCase(Left(cboDataStatus.Text, cboDataStatus.SelStart) + Chr(KeyAscii) + Mid(cboDataStatus.Text, cboDataStatus.SelStart + cboDataStatus.SelLength + 1))
 
    ValidCount = 0
    ValidValue = ""
 
    For i = 0 To cboDataStatus.ListCount - 1
 
        If NewText = LCase(Left(cboDataStatus.List(i), Len(NewText))) Then
                ValidCount = ValidCount + 1
                ValidValue = cboDataStatus.List(i)
        End If
    Next
        
        If ValidCount <= 1 Then KeyAscii = 0
            If ValidCount = 1 Then
                cboDataStatus.Text = ValidValue
                cboDataStatus.SelStart = 0
                cboDataStatus.SelLength = Len(ValidValue)
            End If
        End If
End Sub

Private Sub CboClaimability_Change()

    If CboClaimability.Text = "" Then
        If InStr(CboClaimability.Text, "null") Then
           CboClaimability.Enabled = False
        End If
    End If
End Sub

Private Sub CboClaimability_KeyPress(KeyAscii As Integer)
Dim NewText As String
Dim ValidCount As Integer
Dim ValidValue As String
Dim i As Integer

    If KeyAscii >= 32 And KeyAscii <> 127 Then
    NewText = LCase(Left(CboClaimability.Text, CboClaimability.SelStart) + Chr(KeyAscii) + Mid(CboClaimability.Text, CboClaimability.SelStart + CboClaimability.SelLength + 1))
 
    ValidCount = 0
    ValidValue = ""
 
    For i = 0 To CboClaimability.ListCount - 1
 
        If NewText = LCase(Left(CboClaimability.List(i), Len(NewText))) Then
                ValidCount = ValidCount + 1
                ValidValue = CboClaimability.List(i)
        End If
    Next
        
        If ValidCount <= 1 Then KeyAscii = 0
            If ValidCount = 1 Then
                CboClaimability.Text = ValidValue
                CboClaimability.SelStart = 0
                CboClaimability.SelLength = Len(ValidValue)
            End If
        End If
End Sub

Private Sub CboStatus_Change()

    If CboStatus.Text = "" Then
        If InStr(CboStatus.Text, "null") Then
           CboStatus.Enabled = False
        End If
    End If
End Sub

Private Sub cboStatus_KeyPress(KeyAscii As Integer)
Dim NewText As String
Dim ValidCount As Integer
Dim ValidValue As String
Dim i As Integer

    If KeyAscii >= 32 And KeyAscii <> 127 Then
    NewText = LCase(Left(CboStatus.Text, CboStatus.SelStart) + Chr(KeyAscii) + Mid(CboStatus.Text, CboStatus.SelStart + CboStatus.SelLength + 1))
 
    ValidCount = 0
    ValidValue = ""
 
    For i = 0 To CboStatus.ListCount - 1
 
        If NewText = LCase(Left(CboStatus.List(i), Len(NewText))) Then
                ValidCount = ValidCount + 1
                ValidValue = CboStatus.List(i)
        End If
    Next
        
        If ValidCount <= 1 Then KeyAscii = 0
            If ValidCount = 1 Then
                CboStatus.Text = ValidValue
                CboStatus.SelStart = 0
                CboStatus.SelLength = Len(ValidValue)
            End If
        End If
End Sub

Private Sub CboCaseStatus_Change()

    If CboCaseStatus.Text = "" Then
        If InStr(CboCaseStatus.Text, "null") Then
           CboCaseStatus.Enabled = False
        End If
    End If
End Sub

Private Sub CboCaseStatus_KeyPress(KeyAscii As Integer)
Dim NewText As String
Dim ValidCount As Integer
Dim ValidValue As String
Dim i As Integer

    If KeyAscii >= 32 And KeyAscii <> 127 Then
    NewText = LCase(Left(CboCaseStatus.Text, CboCaseStatus.SelStart) + Chr(KeyAscii) + Mid(CboCaseStatus.Text, CboCaseStatus.SelStart + CboCaseStatus.SelLength + 1))
 
    ValidCount = 0
    ValidValue = ""
 
    For i = 0 To CboCaseStatus.ListCount - 1
 
        If NewText = LCase(Left(CboCaseStatus.List(i), Len(NewText))) Then
                ValidCount = ValidCount + 1
                ValidValue = CboCaseStatus.List(i)
        End If
    Next
        
        If ValidCount <= 1 Then KeyAscii = 0
            If ValidCount = 1 Then
                CboCaseStatus.Text = ValidValue
                CboCaseStatus.SelStart = 0
                CboCaseStatus.SelLength = Len(ValidValue)
            End If
        End If
End Sub

Private Sub CboGrpCom_Change()

    If CboGrpCom.Text = "" Then
        If InStr(CboGrpCom.Text, "null") Then
           CboGrpCom.Enabled = False
        End If
    End If
End Sub

Private Sub CboGrpCom_KeyPress(KeyAscii As Integer)
Dim NewText As String
Dim ValidCount As Integer
Dim ValidValue As String
Dim i As Integer

    If KeyAscii >= 32 And KeyAscii <> 127 Then
    NewText = LCase(Left(CboGrpCom.Text, CboGrpCom.SelStart) + Chr(KeyAscii) + Mid(CboGrpCom.Text, CboGrpCom.SelStart + CboGrpCom.SelLength + 1))
 
    ValidCount = 0
    ValidValue = ""
 
    For i = 0 To CboGrpCom.ListCount - 1
 
        If NewText = LCase(Left(CboGrpCom.List(i), Len(NewText))) Then
                ValidCount = ValidCount + 1
                ValidValue = CboGrpCom.List(i)
        End If
    Next
        
        If ValidCount <= 1 Then KeyAscii = 0
            If ValidCount = 1 Then
                CboGrpCom.Text = ValidValue
                CboGrpCom.SelStart = 0
                CboGrpCom.SelLength = Len(ValidValue)
            End If
        End If
End Sub

Private Sub ListView2_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
m_blnDirection(ColumnHeader.Index) = Not m_blnDirection(ColumnHeader.Index)
    
If (cboReportType.ListIndex = 5 And opt35 = True) Or _
   cboReportType.ListIndex = 6 Then
        Select Case ColumnHeader.Index
        Case 1
            SortListView ListView2, ColumnHeader.Index, ldtString, m_blnDirection(1)
        Case 2
            SortListView ListView2, ColumnHeader.Index, ldtNumber, m_blnDirection(2)
        Case 3
            SortListView ListView2, ColumnHeader.Index, ldtNumber, m_blnDirection(3)
        Case 4
            SortListView ListView2, ColumnHeader.Index, ldtNumber, m_blnDirection(4)
        Case 5
            SortListView ListView2, ColumnHeader.Index, ldtNumber, m_blnDirection(5)
        End Select
ElseIf (cboReportType.ListIndex = 7 And opt35 = True) Then
        Select Case ColumnHeader.Index
        Case 1
            SortListView ListView2, ColumnHeader.Index, ldtString, m_blnDirection(1)
        Case 2
            SortListView ListView2, ColumnHeader.Index, ldtString, m_blnDirection(2)
        Case 3
            SortListView ListView2, ColumnHeader.Index, ldtNumber, m_blnDirection(3)
        Case 4
            SortListView ListView2, ColumnHeader.Index, ldtNumber, m_blnDirection(4)
        Case 5
            SortListView ListView2, ColumnHeader.Index, ldtNumber, m_blnDirection(5)
        Case 6
            SortListView ListView2, ColumnHeader.Index, ldtNumber, m_blnDirection(6)
        End Select
        
ElseIf cboReportType.ListIndex = 7 And opt31 = True Then
        Select Case ColumnHeader.Index
        Case 1
            SortListView ListView2, ColumnHeader.Index, ldtString, m_blnDirection(1)
        Case 2
            SortListView ListView2, ColumnHeader.Index, ldtString, m_blnDirection(2)
        Case 3
            SortListView ListView2, ColumnHeader.Index, ldtNumber, m_blnDirection(3)
        Case 4
            SortListView ListView2, ColumnHeader.Index, ldtNumber, m_blnDirection(4)
        Case 5
            SortListView ListView2, ColumnHeader.Index, ldtNumber, m_blnDirection(5)
        Case 6
            SortListView ListView2, ColumnHeader.Index, ldtNumber, m_blnDirection(6)
        Case 7
            SortListView ListView2, ColumnHeader.Index, ldtNumber, m_blnDirection(7)
        Case 8
            SortListView ListView2, ColumnHeader.Index, ldtNumber, m_blnDirection(8)
        End Select

ElseIf (cboReportType.ListIndex = 1 And opt34 = True) Or _
   (cboReportType.ListIndex = 1 And opt33 = True) Or (cboReportType.ListIndex = 1 And opt29 = True) Or _
   (cboReportType.ListIndex = 1 And opt30 = True) Or _
   (cboReportType.ListIndex = 5 And opt33 = True) Or cboReportType.ListIndex = 4 Then
        Select Case ColumnHeader.Index
        Case 1
            SortListView ListView2, ColumnHeader.Index, ldtString, m_blnDirection(1)
        Case 2
            SortListView ListView2, ColumnHeader.Index, ldtNumber, m_blnDirection(2)
        Case 3
            SortListView ListView2, ColumnHeader.Index, ldtNumber, m_blnDirection(3)
        Case 4
            SortListView ListView2, ColumnHeader.Index, ldtNumber, m_blnDirection(4)
        Case 5
            SortListView ListView2, ColumnHeader.Index, ldtNumber, m_blnDirection(5)
        Case 6
            SortListView ListView2, ColumnHeader.Index, ldtNumber, m_blnDirection(6)
        End Select
ElseIf (cboReportType.ListIndex = 7 And opt29 = True) Or _
    (cboReportType.ListIndex = 3 And opt34 = True) Or (cboReportType.ListIndex = 3 And opt37 = True) Or _
    (cboReportType.ListIndex = 3 And opt29 = True) Then
        Select Case ColumnHeader.Index
        Case 1
            SortListView ListView2, ColumnHeader.Index, ldtString, m_blnDirection(1)
        Case 2
            SortListView ListView2, ColumnHeader.Index, ldtString, m_blnDirection(2)
        Case 3
            SortListView ListView2, ColumnHeader.Index, ldtNumber, m_blnDirection(3)
        Case 4
            SortListView ListView2, ColumnHeader.Index, ldtNumber, m_blnDirection(4)
        Case 5
            SortListView ListView2, ColumnHeader.Index, ldtNumber, m_blnDirection(5)
        Case 6
            SortListView ListView2, ColumnHeader.Index, ldtNumber, m_blnDirection(6)
        Case 7
            SortListView ListView2, ColumnHeader.Index, ldtNumber, m_blnDirection(7)
        Case 8
            SortListView ListView2, ColumnHeader.Index, ldtNumber, m_blnDirection(8)
        Case 9
            SortListView ListView2, ColumnHeader.Index, ldtNumber, m_blnDirection(9)
        End Select
ElseIf (cboReportType.ListIndex = 3 And opt31 = True) Then
        Select Case ColumnHeader.Index
        Case 1
            SortListView ListView2, ColumnHeader.Index, ldtString, m_blnDirection(1)
        Case 2
            SortListView ListView2, ColumnHeader.Index, ldtNumber, m_blnDirection(2)
        Case 3
            SortListView ListView2, ColumnHeader.Index, ldtNumber, m_blnDirection(3)
        Case 4
            SortListView ListView2, ColumnHeader.Index, ldtNumber, m_blnDirection(4)
        Case 5
            SortListView ListView2, ColumnHeader.Index, ldtNumber, m_blnDirection(5)
        Case 6
            SortListView ListView2, ColumnHeader.Index, ldtNumber, m_blnDirection(6)
        Case 7
            SortListView ListView2, ColumnHeader.Index, ldtNumber, m_blnDirection(7)
        Case 8
            SortListView ListView2, ColumnHeader.Index, ldtNumber, m_blnDirection(8)
        End Select
ElseIf cboReportType.ListIndex = 3 And opt33 = True Then
        Select Case ColumnHeader.Index
        Case 1
            SortListView ListView2, ColumnHeader.Index, ldtString, m_blnDirection(1)
        Case 2
            SortListView ListView2, ColumnHeader.Index, ldtNumber, m_blnDirection(2)
        Case 3
            SortListView ListView2, ColumnHeader.Index, ldtNumber, m_blnDirection(3)
        Case 4
            SortListView ListView2, ColumnHeader.Index, ldtNumber, m_blnDirection(4)
        Case 5
            SortListView ListView2, ColumnHeader.Index, ldtNumber, m_blnDirection(5)
        Case 6
            SortListView ListView2, ColumnHeader.Index, ldtNumber, m_blnDirection(6)
        Case 7
            SortListView ListView2, ColumnHeader.Index, ldtNumber, m_blnDirection(7)
        End Select
ElseIf (cboReportType.ListIndex = 5 And opt34 = True) Or _
   (cboReportType.ListIndex = 5 And opt29 = True) Or (cboReportType.ListIndex = 5 And opt37 = True) Then
        Select Case ColumnHeader.Index
        Case 1
            SortListView ListView2, ColumnHeader.Index, ldtString, m_blnDirection(1)
        Case 2
            SortListView ListView2, ColumnHeader.Index, ldtString, m_blnDirection(2)
        Case 3
            SortListView ListView2, ColumnHeader.Index, ldtNumber, m_blnDirection(3)
        Case 4
            SortListView ListView2, ColumnHeader.Index, ldtNumber, m_blnDirection(4)
        Case 5
            SortListView ListView2, ColumnHeader.Index, ldtNumber, m_blnDirection(5)
        Case 6
            SortListView ListView2, ColumnHeader.Index, ldtNumber, m_blnDirection(6)
        End Select
ElseIf (cboReportType.ListIndex = 2 And opt35 = True) Or (cboReportType.ListIndex = 2 And opt34 = True) Or _
   (cboReportType.ListIndex = 2 And opt33 = True) Then
        Select Case ColumnHeader.Index
        Case 1
            SortListView ListView2, ColumnHeader.Index, ldtString, m_blnDirection(1)
        Case 2
            SortListView ListView2, ColumnHeader.Index, ldtString, m_blnDirection(2)
        Case 3
            SortListView ListView2, ColumnHeader.Index, ldtNumber, m_blnDirection(3)
        Case 4
            SortListView ListView2, ColumnHeader.Index, ldtNumber, m_blnDirection(4)
        Case 5
            SortListView ListView2, ColumnHeader.Index, ldtNumber, m_blnDirection(5)
        Case 6
            SortListView ListView2, ColumnHeader.Index, ldtNumber, m_blnDirection(6)
        End Select
ElseIf cboReportType.ListIndex = 2 And opt29 = True Then
        Select Case ColumnHeader.Index
        Case 1
            SortListView ListView2, ColumnHeader.Index, ldtString, m_blnDirection(1)
        Case 2
            SortListView ListView2, ColumnHeader.Index, ldtString, m_blnDirection(2)
        Case 3
            SortListView ListView2, ColumnHeader.Index, ldtString, m_blnDirection(3)
        Case 4
            SortListView ListView2, ColumnHeader.Index, ldtNumber, m_blnDirection(4)
        Case 5
            SortListView ListView2, ColumnHeader.Index, ldtNumber, m_blnDirection(5)
        Case 6
            SortListView ListView2, ColumnHeader.Index, ldtNumber, m_blnDirection(6)
        Case 7
            SortListView ListView2, ColumnHeader.Index, ldtNumber, m_blnDirection(7)
        End Select

ElseIf (cboReportType.ListIndex = 7 And opt34 = True) Or (cboReportType.ListIndex = 7 And opt37 = True) Then
        Select Case ColumnHeader.Index
        Case 1
            SortListView ListView2, ColumnHeader.Index, ldtString, m_blnDirection(1)
        Case 2
            SortListView ListView2, ColumnHeader.Index, ldtString, m_blnDirection(2)
        Case 3
            SortListView ListView2, ColumnHeader.Index, ldtString, m_blnDirection(3)
        Case 4
            SortListView ListView2, ColumnHeader.Index, ldtNumber, m_blnDirection(4)
        Case 5
            SortListView ListView2, ColumnHeader.Index, ldtNumber, m_blnDirection(5)
        Case 6
            SortListView ListView2, ColumnHeader.Index, ldtNumber, m_blnDirection(6)
        Case 7
            SortListView ListView2, ColumnHeader.Index, ldtNumber, m_blnDirection(7)
        End Select
ElseIf cboReportType.ListIndex = 7 And opt30 = True Then
        Select Case ColumnHeader.Index
        Case 1
            SortListView ListView2, ColumnHeader.Index, ldtString, m_blnDirection(1)
        Case 2
            SortListView ListView2, ColumnHeader.Index, ldtString, m_blnDirection(2)
        Case 3
            SortListView ListView2, ColumnHeader.Index, ldtString, m_blnDirection(3)
        Case 4
            SortListView ListView2, ColumnHeader.Index, ldtString, m_blnDirection(4)
        Case 5
            SortListView ListView2, ColumnHeader.Index, ldtNumber, m_blnDirection(5)
        Case 6
            SortListView ListView2, ColumnHeader.Index, ldtNumber, m_blnDirection(6)
        Case 7
            SortListView ListView2, ColumnHeader.Index, ldtNumber, m_blnDirection(7)
        Case 8
            SortListView ListView2, ColumnHeader.Index, ldtNumber, m_blnDirection(8)
        End Select
ElseIf cboReportType.ListIndex = 7 And opt33 = True Then
        Select Case ColumnHeader.Index
        Case 1
            SortListView ListView2, ColumnHeader.Index, ldtString, m_blnDirection(1)
        Case 2
            SortListView ListView2, ColumnHeader.Index, ldtString, m_blnDirection(2)
        Case 3
            SortListView ListView2, ColumnHeader.Index, ldtString, m_blnDirection(3)
        Case 4
            SortListView ListView2, ColumnHeader.Index, ldtString, m_blnDirection(4)
        Case 5
            SortListView ListView2, ColumnHeader.Index, ldtNumber, m_blnDirection(5)
        Case 6
            SortListView ListView2, ColumnHeader.Index, ldtNumber, m_blnDirection(6)
        Case 7
            SortListView ListView2, ColumnHeader.Index, ldtNumber, m_blnDirection(7)
        Case 8
            SortListView ListView2, ColumnHeader.Index, ldtNumber, m_blnDirection(8)
        End Select


ElseIf (cboReportType.ListIndex = 1 And opt35 = True) Then
        Select Case ColumnHeader.Index
        Case 1
            SortListView ListView2, ColumnHeader.Index, ldtNumber, m_blnDirection(1)
        Case 2
            SortListView ListView2, ColumnHeader.Index, ldtNumber, m_blnDirection(2)
        Case 3
            SortListView ListView2, ColumnHeader.Index, ldtNumber, m_blnDirection(3)
        Case 4
            SortListView ListView2, ColumnHeader.Index, ldtNumber, m_blnDirection(4)
        Case 5
            SortListView ListView2, ColumnHeader.Index, ldtNumber, m_blnDirection(5)
        End Select


ElseIf (cboReportType.ListIndex = 1 And opt37 = True) Then
        Select Case ColumnHeader.Index
        Case 1
            SortListView ListView2, ColumnHeader.Index, ldtString, m_blnDirection(1)
        Case 2
            SortListView ListView2, ColumnHeader.Index, ldtString, m_blnDirection(2)
        Case 3
            SortListView ListView2, ColumnHeader.Index, ldtNumber, m_blnDirection(3)
        Case 4
            SortListView ListView2, ColumnHeader.Index, ldtNumber, m_blnDirection(4)
        Case 5
            SortListView ListView2, ColumnHeader.Index, ldtNumber, m_blnDirection(5)
        Case 6
            SortListView ListView2, ColumnHeader.Index, ldtNumber, m_blnDirection(6)
        Case 7
            SortListView ListView2, ColumnHeader.Index, ldtNumber, m_blnDirection(7)
        End Select



ElseIf (cboReportType.ListIndex = 3 And opt35 = True) Then
        Select Case ColumnHeader.Index
        Case 1
            SortListView ListView2, ColumnHeader.Index, ldtString, m_blnDirection(1)
        Case 2
            SortListView ListView2, ColumnHeader.Index, ldtString, m_blnDirection(2)
        Case 3
            SortListView ListView2, ColumnHeader.Index, ldtString, m_blnDirection(3)
        Case 4
            SortListView ListView2, ColumnHeader.Index, ldtString, m_blnDirection(4)
        Case 5
            SortListView ListView2, ColumnHeader.Index, ldtNumber, m_blnDirection(5)
        Case 6
            SortListView ListView2, ColumnHeader.Index, ldtString, m_blnDirection(6)
        Case 7
            SortListView ListView2, ColumnHeader.Index, ldtNumber, m_blnDirection(7)
        Case 8
            SortListView ListView2, ColumnHeader.Index, ldtNumber, m_blnDirection(8)
        Case 9
            SortListView ListView2, ColumnHeader.Index, ldtNumber, m_blnDirection(9)
        Case 10
            SortListView ListView2, ColumnHeader.Index, ldtDateTime, m_blnDirection(10)
        Case 11
            SortListView ListView2, ColumnHeader.Index, ldtDateTime, m_blnDirection(11)
        Case 12
            SortListView ListView2, ColumnHeader.Index, ldtString, m_blnDirection(12)
        Case 13
            SortListView ListView2, ColumnHeader.Index, ldtString, m_blnDirection(13)
        Case 14
            SortListView ListView2, ColumnHeader.Index, ldtString, m_blnDirection(14)
        End Select
ElseIf (cboReportType.ListIndex = 3 And opt30 = True) Then
        Select Case ColumnHeader.Index
        Case 1
            SortListView ListView2, ColumnHeader.Index, ldtString, m_blnDirection(1)
        Case 2
            SortListView ListView2, ColumnHeader.Index, ldtString, m_blnDirection(2)
        Case 3
            SortListView ListView2, ColumnHeader.Index, ldtString, m_blnDirection(3)
        Case 4
            SortListView ListView2, ColumnHeader.Index, ldtNumber, m_blnDirection(4)
        Case 5
            SortListView ListView2, ColumnHeader.Index, ldtNumber, m_blnDirection(5)
        Case 6
            SortListView ListView2, ColumnHeader.Index, ldtNumber, m_blnDirection(6)
        Case 7
            SortListView ListView2, ColumnHeader.Index, ldtNumber, m_blnDirection(7)
        End Select
ElseIf cboReportType.ListIndex = 0 And opt35 = True And cboClaimCase.ListIndex = 0 Then
        Select Case ColumnHeader.Index
        Case 1
            SortListView ListView2, ColumnHeader.Index, ldtString, m_blnDirection(1)
        Case 2
            SortListView ListView2, ColumnHeader.Index, ldtString, m_blnDirection(2)
        Case 3
            SortListView ListView2, ColumnHeader.Index, ldtString, m_blnDirection(3)
        Case 4
            SortListView ListView2, ColumnHeader.Index, ldtDateTime, m_blnDirection(4)
        Case 5
            SortListView ListView2, ColumnHeader.Index, ldtString, m_blnDirection(5)
        Case 6
            SortListView ListView2, ColumnHeader.Index, ldtString, m_blnDirection(6)
        Case 7
            SortListView ListView2, ColumnHeader.Index, ldtNumber, m_blnDirection(7)
        Case 8
            SortListView ListView2, ColumnHeader.Index, ldtString, m_blnDirection(8)
        Case 9
            SortListView ListView2, ColumnHeader.Index, ldtString, m_blnDirection(9)
        Case 10
            SortListView ListView2, ColumnHeader.Index, ldtString, m_blnDirection(10)
        Case 11
            SortListView ListView2, ColumnHeader.Index, ldtString, m_blnDirection(11)
        Case 12
            SortListView ListView2, ColumnHeader.Index, ldtNumber, m_blnDirection(12)
        Case 13
            SortListView ListView2, ColumnHeader.Index, ldtDateTime, m_blnDirection(13)
        Case 14
            SortListView ListView2, ColumnHeader.Index, ldtString, m_blnDirection(14)
        Case 15
            SortListView ListView2, ColumnHeader.Index, ldtString, m_blnDirection(15)
        Case 16
            SortListView ListView2, ColumnHeader.Index, ldtNumber, m_blnDirection(16)
        Case 17
            SortListView ListView2, ColumnHeader.Index, ldtDateTime, m_blnDirection(17)
        Case 18
            SortListView ListView2, ColumnHeader.Index, ldtString, m_blnDirection(18)
        Case 19
            SortListView ListView2, ColumnHeader.Index, ldtString, m_blnDirection(19)
        Case 20
            SortListView ListView2, ColumnHeader.Index, ldtString, m_blnDirection(20)
        Case 21
            SortListView ListView2, ColumnHeader.Index, ldtString, m_blnDirection(21)
        Case 22
            SortListView ListView2, ColumnHeader.Index, ldtString, m_blnDirection(22)
        Case 23
            SortListView ListView2, ColumnHeader.Index, ldtString, m_blnDirection(23)
        Case 24
            SortListView ListView2, ColumnHeader.Index, ldtString, m_blnDirection(24)
        Case 25
            SortListView ListView2, ColumnHeader.Index, ldtNumber, m_blnDirection(25)
        Case 26
            SortListView ListView2, ColumnHeader.Index, ldtNumber, m_blnDirection(26)
        Case 27
            SortListView ListView2, ColumnHeader.Index, ldtNumber, m_blnDirection(27)
        Case 28
            SortListView ListView2, ColumnHeader.Index, ldtNumber, m_blnDirection(28)
        Case 29
            SortListView ListView2, ColumnHeader.Index, ldtString, m_blnDirection(29)
        Case 30
            SortListView ListView2, ColumnHeader.Index, ldtString, m_blnDirection(30)
        Case 31
            SortListView ListView2, ColumnHeader.Index, ldtString, m_blnDirection(31)
        Case 32
            SortListView ListView2, ColumnHeader.Index, ldtDateTime, m_blnDirection(32)
        Case 33
            SortListView ListView2, ColumnHeader.Index, ldtDateTime, m_blnDirection(33)
        Case 34
            SortListView ListView2, ColumnHeader.Index, ldtDateTime, m_blnDirection(34)
        Case 35
            SortListView ListView2, ColumnHeader.Index, ldtNumber, m_blnDirection(35)
        Case 36
            SortListView ListView2, ColumnHeader.Index, ldtDateTime, m_blnDirection(36)
        Case 37
            SortListView ListView2, ColumnHeader.Index, ldtDateTime, m_blnDirection(37)
        Case 38
            SortListView ListView2, ColumnHeader.Index, ldtDateTime, m_blnDirection(38)
        End Select
ElseIf cboReportType.ListIndex = 0 And opt34 = True And cboClaimCase.ListIndex = 0 Then
        Select Case ColumnHeader.Index
        Case 1
            SortListView ListView2, ColumnHeader.Index, ldtString, m_blnDirection(1)
        Case 2
            SortListView ListView2, ColumnHeader.Index, ldtString, m_blnDirection(2)
        Case 3
            SortListView ListView2, ColumnHeader.Index, ldtString, m_blnDirection(3)
        Case 4
            SortListView ListView2, ColumnHeader.Index, ldtDateTime, m_blnDirection(4)
        Case 5
            SortListView ListView2, ColumnHeader.Index, ldtString, m_blnDirection(5)
        Case 6
            SortListView ListView2, ColumnHeader.Index, ldtString, m_blnDirection(6)
        Case 7
            SortListView ListView2, ColumnHeader.Index, ldtNumber, m_blnDirection(7)
        Case 8
            SortListView ListView2, ColumnHeader.Index, ldtString, m_blnDirection(8)
        Case 9
            SortListView ListView2, ColumnHeader.Index, ldtString, m_blnDirection(9)
        Case 10
            SortListView ListView2, ColumnHeader.Index, ldtString, m_blnDirection(10)
        Case 11
            SortListView ListView2, ColumnHeader.Index, ldtString, m_blnDirection(11)
        Case 12
            SortListView ListView2, ColumnHeader.Index, ldtNumber, m_blnDirection(12)
        Case 13
            SortListView ListView2, ColumnHeader.Index, ldtDateTime, m_blnDirection(13)
        Case 14
            SortListView ListView2, ColumnHeader.Index, ldtString, m_blnDirection(14)
        Case 15
            SortListView ListView2, ColumnHeader.Index, ldtString, m_blnDirection(15)
        Case 16
            SortListView ListView2, ColumnHeader.Index, ldtNumber, m_blnDirection(16)
        Case 17
            SortListView ListView2, ColumnHeader.Index, ldtDateTime, m_blnDirection(17)
        Case 18
            SortListView ListView2, ColumnHeader.Index, ldtString, m_blnDirection(18)
        Case 19
            SortListView ListView2, ColumnHeader.Index, ldtString, m_blnDirection(19)
        Case 20
            SortListView ListView2, ColumnHeader.Index, ldtString, m_blnDirection(20)
        Case 21
            SortListView ListView2, ColumnHeader.Index, ldtString, m_blnDirection(21)
        Case 22
            SortListView ListView2, ColumnHeader.Index, ldtString, m_blnDirection(22)
        Case 23
            SortListView ListView2, ColumnHeader.Index, ldtString, m_blnDirection(23)
        Case 24
            SortListView ListView2, ColumnHeader.Index, ldtString, m_blnDirection(24)
        Case 25
            SortListView ListView2, ColumnHeader.Index, ldtNumber, m_blnDirection(25)
        Case 26
            SortListView ListView2, ColumnHeader.Index, ldtNumber, m_blnDirection(26)
        Case 27
            SortListView ListView2, ColumnHeader.Index, ldtNumber, m_blnDirection(27)
        Case 28
            SortListView ListView2, ColumnHeader.Index, ldtNumber, m_blnDirection(28)
        Case 29
            SortListView ListView2, ColumnHeader.Index, ldtString, m_blnDirection(29)
        Case 30
            SortListView ListView2, ColumnHeader.Index, ldtString, m_blnDirection(30)
        Case 31
            SortListView ListView2, ColumnHeader.Index, ldtString, m_blnDirection(31)
        Case 32
            SortListView ListView2, ColumnHeader.Index, ldtDateTime, m_blnDirection(32)
        Case 33
            SortListView ListView2, ColumnHeader.Index, ldtDateTime, m_blnDirection(33)
        Case 34
            SortListView ListView2, ColumnHeader.Index, ldtDateTime, m_blnDirection(34)
        Case 35
            SortListView ListView2, ColumnHeader.Index, ldtNumber, m_blnDirection(35)
        Case 36
            SortListView ListView2, ColumnHeader.Index, ldtDateTime, m_blnDirection(36)
        Case 37
            SortListView ListView2, ColumnHeader.Index, ldtDateTime, m_blnDirection(37)
        Case 38
            SortListView ListView2, ColumnHeader.Index, ldtDateTime, m_blnDirection(38)
        Case 39
            SortListView ListView2, ColumnHeader.Index, ldtString, m_blnDirection(39)
        End Select



    
    ElseIf cboReportType.ListIndex = 0 And opt35 = True And cboClaimCase.ListIndex = 1 Then
        Select Case ColumnHeader.Index
        Case 1
            SortListView ListView2, ColumnHeader.Index, ldtString, m_blnDirection(1)
        Case 2
            SortListView ListView2, ColumnHeader.Index, ldtString, m_blnDirection(2)
        Case 3
            SortListView ListView2, ColumnHeader.Index, ldtString, m_blnDirection(3)
        Case 4
            SortListView ListView2, ColumnHeader.Index, ldtDateTime, m_blnDirection(4)
        Case 5
            SortListView ListView2, ColumnHeader.Index, ldtString, m_blnDirection(5)
        Case 6
            SortListView ListView2, ColumnHeader.Index, ldtString, m_blnDirection(6)
        Case 7
            SortListView ListView2, ColumnHeader.Index, ldtNumber, m_blnDirection(7)
        Case 8
            SortListView ListView2, ColumnHeader.Index, ldtString, m_blnDirection(8)
        Case 9
            SortListView ListView2, ColumnHeader.Index, ldtString, m_blnDirection(9)
        Case 10
            SortListView ListView2, ColumnHeader.Index, ldtString, m_blnDirection(10)
        Case 11
            SortListView ListView2, ColumnHeader.Index, ldtString, m_blnDirection(11)
        Case 12
            SortListView ListView2, ColumnHeader.Index, ldtNumber, m_blnDirection(12)
        Case 13
            SortListView ListView2, ColumnHeader.Index, ldtDateTime, m_blnDirection(13)
        Case 14
            SortListView ListView2, ColumnHeader.Index, ldtString, m_blnDirection(14)
        Case 15
            SortListView ListView2, ColumnHeader.Index, ldtString, m_blnDirection(15)
        Case 16
            SortListView ListView2, ColumnHeader.Index, ldtNumber, m_blnDirection(16)
        Case 17
            SortListView ListView2, ColumnHeader.Index, ldtDateTime, m_blnDirection(17)
        Case 18
            SortListView ListView2, ColumnHeader.Index, ldtString, m_blnDirection(18)
        Case 19
            SortListView ListView2, ColumnHeader.Index, ldtString, m_blnDirection(19)
        Case 20
            SortListView ListView2, ColumnHeader.Index, ldtString, m_blnDirection(20)
        Case 21
            SortListView ListView2, ColumnHeader.Index, ldtString, m_blnDirection(21)
        Case 22
            SortListView ListView2, ColumnHeader.Index, ldtString, m_blnDirection(22)
        Case 23
            SortListView ListView2, ColumnHeader.Index, ldtString, m_blnDirection(23)
        Case 24
            SortListView ListView2, ColumnHeader.Index, ldtString, m_blnDirection(24)
        Case 25
            SortListView ListView2, ColumnHeader.Index, ldtNumber, m_blnDirection(25)
        Case 26
            SortListView ListView2, ColumnHeader.Index, ldtNumber, m_blnDirection(26)
        Case 27
            SortListView ListView2, ColumnHeader.Index, ldtNumber, m_blnDirection(27)
        Case 28
            SortListView ListView2, ColumnHeader.Index, ldtNumber, m_blnDirection(28)
        Case 29
            SortListView ListView2, ColumnHeader.Index, ldtNumber, m_blnDirection(29)
        Case 30
            SortListView ListView2, ColumnHeader.Index, ldtNumber, m_blnDirection(30)
        Case 31
            SortListView ListView2, ColumnHeader.Index, ldtString, m_blnDirection(31)
        Case 32
            SortListView ListView2, ColumnHeader.Index, ldtString, m_blnDirection(32)
        Case 33
            SortListView ListView2, ColumnHeader.Index, ldtString, m_blnDirection(33)
        Case 34
            SortListView ListView2, ColumnHeader.Index, ldtDateTime, m_blnDirection(34)
        Case 35
            SortListView ListView2, ColumnHeader.Index, ldtDateTime, m_blnDirection(35)
        Case 36
            SortListView ListView2, ColumnHeader.Index, ldtDateTime, m_blnDirection(36)
        Case 37
            SortListView ListView2, ColumnHeader.Index, ldtNumber, m_blnDirection(37)
        Case 38
            SortListView ListView2, ColumnHeader.Index, ldtDateTime, m_blnDirection(38)
        Case 39
            SortListView ListView2, ColumnHeader.Index, ldtDateTime, m_blnDirection(39)
        Case 40
            SortListView ListView2, ColumnHeader.Index, ldtDateTime, m_blnDirection(40)
        End Select
ElseIf cboReportType.ListIndex = 0 And opt34 = True And cboClaimCase.ListIndex = 1 Then
        Select Case ColumnHeader.Index
        Case 1
            SortListView ListView2, ColumnHeader.Index, ldtString, m_blnDirection(1)
        Case 2
            SortListView ListView2, ColumnHeader.Index, ldtString, m_blnDirection(2)
        Case 3
            SortListView ListView2, ColumnHeader.Index, ldtString, m_blnDirection(3)
        Case 4
            SortListView ListView2, ColumnHeader.Index, ldtDateTime, m_blnDirection(4)
        Case 5
            SortListView ListView2, ColumnHeader.Index, ldtString, m_blnDirection(5)
        Case 6
            SortListView ListView2, ColumnHeader.Index, ldtString, m_blnDirection(6)
        Case 7
            SortListView ListView2, ColumnHeader.Index, ldtNumber, m_blnDirection(7)
        Case 8
            SortListView ListView2, ColumnHeader.Index, ldtString, m_blnDirection(8)
        Case 9
            SortListView ListView2, ColumnHeader.Index, ldtString, m_blnDirection(9)
        Case 10
            SortListView ListView2, ColumnHeader.Index, ldtString, m_blnDirection(10)
        Case 11
            SortListView ListView2, ColumnHeader.Index, ldtString, m_blnDirection(11)
        Case 12
            SortListView ListView2, ColumnHeader.Index, ldtNumber, m_blnDirection(12)
        Case 13
            SortListView ListView2, ColumnHeader.Index, ldtDateTime, m_blnDirection(13)
        Case 14
            SortListView ListView2, ColumnHeader.Index, ldtString, m_blnDirection(14)
        Case 15
            SortListView ListView2, ColumnHeader.Index, ldtString, m_blnDirection(15)
        Case 16
            SortListView ListView2, ColumnHeader.Index, ldtNumber, m_blnDirection(16)
        Case 17
            SortListView ListView2, ColumnHeader.Index, ldtDateTime, m_blnDirection(17)
        Case 18
            SortListView ListView2, ColumnHeader.Index, ldtString, m_blnDirection(18)
        Case 19
            SortListView ListView2, ColumnHeader.Index, ldtString, m_blnDirection(19)
        Case 20
            SortListView ListView2, ColumnHeader.Index, ldtString, m_blnDirection(20)
        Case 21
            SortListView ListView2, ColumnHeader.Index, ldtString, m_blnDirection(21)
        Case 22
            SortListView ListView2, ColumnHeader.Index, ldtString, m_blnDirection(22)
        Case 23
            SortListView ListView2, ColumnHeader.Index, ldtString, m_blnDirection(23)
        Case 24
            SortListView ListView2, ColumnHeader.Index, ldtString, m_blnDirection(24)
        Case 25
            SortListView ListView2, ColumnHeader.Index, ldtNumber, m_blnDirection(25)
        Case 26
            SortListView ListView2, ColumnHeader.Index, ldtNumber, m_blnDirection(26)
        Case 27
            SortListView ListView2, ColumnHeader.Index, ldtNumber, m_blnDirection(27)
        Case 28
            SortListView ListView2, ColumnHeader.Index, ldtNumber, m_blnDirection(28)
        Case 29
            SortListView ListView2, ColumnHeader.Index, ldtNumber, m_blnDirection(29)
        Case 30
            SortListView ListView2, ColumnHeader.Index, ldtNumber, m_blnDirection(30)
        Case 31
            SortListView ListView2, ColumnHeader.Index, ldtString, m_blnDirection(31)
        Case 32
            SortListView ListView2, ColumnHeader.Index, ldtString, m_blnDirection(32)
        Case 33
            SortListView ListView2, ColumnHeader.Index, ldtString, m_blnDirection(33)
        Case 34
            SortListView ListView2, ColumnHeader.Index, ldtDateTime, m_blnDirection(34)
        Case 35
            SortListView ListView2, ColumnHeader.Index, ldtDateTime, m_blnDirection(35)
        Case 36
            SortListView ListView2, ColumnHeader.Index, ldtDateTime, m_blnDirection(36)
        Case 37
            SortListView ListView2, ColumnHeader.Index, ldtNumber, m_blnDirection(37)
        Case 38
            SortListView ListView2, ColumnHeader.Index, ldtDateTime, m_blnDirection(38)
        Case 39
            SortListView ListView2, ColumnHeader.Index, ldtDateTime, m_blnDirection(39)
        Case 40
            SortListView ListView2, ColumnHeader.Index, ldtDateTime, m_blnDirection(40)
        Case 41
            SortListView ListView2, ColumnHeader.Index, ldtString, m_blnDirection(41)
        End Select
End If

End Sub

'=========== Combo Box Change/KeyPress ===================

Private Sub rptGeneralListing()
    lblCurrent = ""
    lblTotalRecord = ""
    ListView2.ListItems.Clear
    cboReportSelection.Clear
    cboReportSelection.AddItem "Without Remarks"  '35
    cboReportSelection.ListIndex = 0
    opt35 = True
    Call ListViewHeader15
End Sub

Private Sub rptHospital()
    lblCurrent = ""
    lblTotalRecord = ""
    ListView2.ListItems.Clear
    If cboClaimCase.ListIndex = 0 Then
        cboReportSelection.AddItem "Detail Listing"     '35
        cboReportSelection.AddItem "General Summary (By Payor)"     '34
        cboReportSelection.AddItem "General Summary (All)"      '33
        cboReportSelection.AddItem "Top 20 by Claims/Cases"     '29
        cboReportSelection.AddItem "Top 20 by Amount (By Payor)"    '37
        cboReportSelection.AddItem "Top 20 by Amount (All Payor)"   '31
        cboReportSelection.AddItem "By Bordx Month"     '30
        'cboReportSelection.AddItem "By Bordx Month(Payor)"     '16
        opt35 = True
        Call ListViewHeader17
    ElseIf cboClaimCase.ListIndex = 1 Then
        cboReportSelection.AddItem "Detail Listing"     '35
        cboReportSelection.AddItem "General Summary (By Payor)"     '34
        cboReportSelection.AddItem "General Summary (All)"      '33
        cboReportSelection.AddItem "Top 20 by Claims/Cases"     '29
        cboReportSelection.AddItem "Top 20 by Amount (By Payor)"    '37
        cboReportSelection.AddItem "Top 20 by Amount (All Payor)"   '31
        opt35 = True
        Call ListViewHeader17
    End If
End Sub

Private Sub Option11()
    lblCurrent = ""
    lblTotalRecord = ""
    ListView2.ListItems.Clear
    
    If cboReportType.ListIndex = 4 Then
        Call ListViewHeader7
    End If
End Sub

Private Sub Option12()
    lblCurrent = ""
    lblTotalRecord = ""
    ListView2.ListItems.Clear
    
    If cboReportType.ListIndex = 4 Then
        Call ListViewHeader7
    End If
End Sub

Private Sub Option13()
    lblCurrent = ""
    lblTotalRecord = ""
    ListView2.ListItems.Clear
    
    If cboReportType.ListIndex = 4 Then
        Call ListViewHeader7
    End If
End Sub

Private Sub Option14()
    TxtBordxNo1.Text = ""
    TxtBordxNo2.Text = ""
    ChkBordx.Value = 0
    TxtBordxDate1.Text = "__/__/____"
    TxtBordxDate2.Text = "__/__/____"
    TxtVersion1.Text = ""
    TxtVersion2.Text = ""
    ChkBordxDate.Value = 0
    TxtBordxNo1.Enabled = False
    TxtBordxNo2.Enabled = False
    ChkBordx.Enabled = False
    TxtBordxDate1.Enabled = False
    TxtBordxDate2.Enabled = False
    TxtVersion1.Enabled = False
    TxtVersion2.Enabled = False
    ChkBordxDate.Enabled = False
    opt35 = True
    
    If (cboReportType.ListIndex = 0 And opt35 = True) Or _
       (cboReportType.ListIndex = 0 And opt34 = True) Then
        Call ListViewHeader15
    ElseIf (cboReportType.ListIndex = 5 And opt35 = True) Or _
       (cboReportType.ListIndex = 5 And opt34 = True) Or _
       (cboReportType.ListIndex = 5 And opt33 = True) Or _
       (cboReportType.ListIndex = 5 And opt29 = True) Or _
       (cboReportType.ListIndex = 5 And opt37 = True) Then
        Call ListViewHeader18
   ElseIf (cboReportType.ListIndex = 1 And opt35 = True) Or _
       (cboReportType.ListIndex = 1 And opt34 = True) Or _
       (cboReportType.ListIndex = 1 And opt33 = True) Or _
       (cboReportType.ListIndex = 1 And opt29 = True) Or _
       (cboReportType.ListIndex = 1 And opt37 = True) Or _
       (cboReportType.ListIndex = 1 And opt30 = True) Then
        Call ListViewHeader3
   
    ElseIf (cboReportType.ListIndex = 3 And opt35 = True) Or _
       (cboReportType.ListIndex = 3 And opt34 = True) Or _
       (cboReportType.ListIndex = 3 And opt33 = True) Or _
       (cboReportType.ListIndex = 3 And opt29 = True) Or _
       (cboReportType.ListIndex = 3 And opt37 = True) Or _
       (cboReportType.ListIndex = 3 And opt31 = True) Or _
       (cboReportType.ListIndex = 3 And opt30 = True) Then
        Call ListViewHeader17
   
    ElseIf (cboReportType.ListIndex = 2 And opt29 = True) Then
        Call ListViewHeader8
   
    ElseIf cboReportType.ListIndex = 6 Then
        Call ListViewHeader20
        
    ElseIf cboReportType.ListIndex = 8 Then
        Call ListViewHeader5
    
    ElseIf (cboReportType.ListIndex = 4 And opt35 = True) Or _
       (cboReportType.ListIndex = 4 And opt34 = True) Or _
       (cboReportType.ListIndex = 4 And opt33 = True) Or _
       (cboReportType.ListIndex = 4 And opt29 = True) Or _
       (cboReportType.ListIndex = 4 And opt37 = True) Or _
       (cboReportType.ListIndex = 4 And opt30 = True) Or _
       (cboReportType.ListIndex = 4 And Opt8 = True) Or _
       (cboReportType.ListIndex = 4 And Opt11 = True) Or _
       (cboReportType.ListIndex = 4 And opt12 = True) Or _
       (cboReportType.ListIndex = 4 And opt13 = True) Or _
       (cboReportType.ListIndex = 4 And Opt7 = True) Or _
       (cboReportType.ListIndex = 4 And opt16 = True) Then
        Call ListViewHeader7
    End If
    
End Sub

Private Sub rptDiagnosis()
    lblCurrent = ""
    lblTotalRecord = ""
    ListView2.ListItems.Clear
    cboReportSelection.AddItem "Top 20 by Claims/Cases" '35
    cboReportSelection.AddItem "Top 20 by Amount"   '34
    cboReportSelection.AddItem "General Summary"    '33
    cboReportSelection.AddItem "By Hospital"    '29
    opt35 = True
    Call ListViewHeader8
End Sub

Private Sub Option16()
    lblCurrent = ""
    lblTotalRecord = ""
    ListView2.ListItems.Clear
    
    'If cboClaimCase.ListIndex = 0 Then
    '    Option16.Caption = "By Bordx Month(Payor)"
    '    Option30.Caption = "By Bordx Month"
    'End If
    
    If cboReportType.ListIndex = 4 Then
        Call ListViewHeader7
    End If
End Sub

Private Sub rptMonth()
    lblCurrent = ""
    lblTotalRecord = ""
    ListView2.ListItems.Clear
    If cboClaimCase.ListIndex = 0 Then
        cboReportSelection.AddItem "By Reg Month"   '35
        cboReportSelection.AddItem "By Eff Month"   '34
        cboReportSelection.AddItem "By Exp Month"   '33
        cboReportSelection.AddItem "By DOA Month"   '29
        cboReportSelection.AddItem "By DOD Month"   '37
        cboReportSelection.AddItem "By Bordx Month" '30
        cboReportSelection.AddItem "By Reg Month (Payor)"   '8
        cboReportSelection.AddItem "By Eff Month (Payor)"   '11
        cboReportSelection.AddItem "By Exp Month (Payor)"   '12
        cboReportSelection.AddItem "By DOA Month (Payor)"   '13
        cboReportSelection.AddItem "By DOD Month (Payor)"   '7
        cboReportSelection.AddItem "By Bordx Month (Payor)"   '16
        opt35 = True
        Call ListViewHeader7
    ElseIf cboClaimCase.ListIndex = 1 Then
        cboReportSelection.AddItem "By Reg Month"   '35
        cboReportSelection.AddItem "By Eff Month"   '34
        cboReportSelection.AddItem "By Exp Month"   '33
        cboReportSelection.AddItem "By DOA Month"   '29
        cboReportSelection.AddItem "By DOD Month"   '37
        cboReportSelection.AddItem "By Reg Month (Payor)"   '8
        cboReportSelection.AddItem "By Eff Month (Payor)"   '11
        cboReportSelection.AddItem "By Exp Month (Payor)"   '12
        cboReportSelection.AddItem "By DOA Month (Payor)"   '13
        cboReportSelection.AddItem "By DOD Month (Payor)"   '7
        opt35 = True
        Call ListViewHeader7
    End If
    
End Sub

Private Sub rptEfficiency()
    lblCurrent = ""
    lblTotalRecord = ""
    ListView2.ListItems.Clear
    cboReportSelection.AddItem "Registration"   '35
    cboReportSelection.AddItem "Bordereaux (Claim)" '34
    cboReportSelection.AddItem "Invoice"    '33
    cboReportSelection.AddItem "Bordereaux (Inv.)"  '29
    cboReportSelection.AddItem "O/S Bordereaux"     '37
    opt35 = True
    Call ListViewHeader10
End Sub

Private Sub rptInsuredType()
    lblCurrent = ""
    lblTotalRecord = ""
    ListView2.ListItems.Clear
    cboReportSelection.AddItem "Summary"    '35
    cboReportSelection.AddItem "By Plan No"     '34
    cboReportSelection.AddItem "By Ageband"     '33
    cboReportSelection.AddItem "By Gender"      '29
    cboReportSelection.AddItem "By Race"    '37
    opt35 = True
    Call ListViewHeader18
End Sub

Private Sub Option29()
    
    lblCurrent = ""
    lblTotalRecord = ""
    ListView2.ListItems.Clear
    
    If cboReportType.ListIndex = 5 Then
        Call ListViewHeader18
    ElseIf cboReportType.ListIndex = 7 Then
        Call ListViewHeader21
    ElseIf cboReportType.ListIndex = 1 Then
        Call ListViewHeader3
    ElseIf cboReportType.ListIndex = 2 Then
        Call ListViewHeader8
    ElseIf cboReportType.ListIndex = 3 Then
         Call ListViewHeader17
    ElseIf cboReportType.ListIndex = 4 Then
        Call ListViewHeader7
    ElseIf cboReportType.ListIndex = 9 Then
        Call ListViewHeader10
    End If
    
End Sub

Private Sub rptPlanNo()
    lblCurrent = ""
    lblTotalRecord = ""
    ListView2.ListItems.Clear
    If cboClaimCase.ListIndex = 0 Then
        cboReportSelection.AddItem "General Summary"    '35
        cboReportSelection.AddItem "Top 10 Hospitals"   '34
        cboReportSelection.AddItem "Top 10 Diagnosis"   '33
        cboReportSelection.AddItem "By Ageband"     '29
        cboReportSelection.AddItem "By Insured Type"    '37
        cboReportSelection.AddItem "General Summary with Premium"   '31
        cboReportSelection.AddItem "By Bordx Month"     '30
        'cboReportSelection.AddItem "By Bordx Month(Payor)"     '16
        opt35 = True
        Call ListViewHeader21
    ElseIf cboClaimCase.ListIndex = 1 Then
        cboReportSelection.AddItem "General Summary"    '35
        cboReportSelection.AddItem "Top 10 Hospitals"   '34
        cboReportSelection.AddItem "Top 10 Diagnosis"   '33
        cboReportSelection.AddItem "By Ageband"     '29
        cboReportSelection.AddItem "By Insured Type"    '37
        cboReportSelection.AddItem "General Summary with Premium"   '31
        opt35 = True
        Call ListViewHeader21
    End If
    
End Sub

Private Sub Option30()
    
    lblCurrent = ""
    lblTotalRecord = ""
    ListView2.ListItems.Clear
    
    'If cboClaimCase.ListIndex = 0 Then
    '    Option16.Caption = "By Bordx Month(Payor)"
    '    Option30.Caption = "By Bordx Month"
    'End If
    
    If cboReportType.ListIndex = 1 Then
        Call ListViewHeader3
    ElseIf cboReportType.ListIndex = 7 Then
        Call ListViewHeader21
    ElseIf cboReportType.ListIndex = 3 Then
        Call ListViewHeader17
    ElseIf cboReportType.ListIndex = 4 Then
        Call ListViewHeader7
    End If
End Sub

Private Sub Option31()
    lblCurrent = ""
    lblTotalRecord = ""
    ListView2.ListItems.Clear


    If cboReportType.ListIndex = 7 Then
        ListViewHeader21
    ElseIf cboReportType.ListIndex = 3 Then
        ListViewHeader17
    End If
    
End Sub

Private Sub Option33()
    
    lblCurrent = ""
    lblTotalRecord = ""
    ListView2.ListItems.Clear
    
    If cboReportType.ListIndex = 5 Then
        Call ListViewHeader18
    ElseIf cboReportType.ListIndex = 7 Then
        Call ListViewHeader21
    ElseIf cboReportType.ListIndex = 1 Then
        Call ListViewHeader3
    ElseIf cboReportType.ListIndex = 3 Then
        Call ListViewHeader17
    ElseIf cboReportType.ListIndex = 4 Then
        Call ListViewHeader7
    ElseIf cboReportType.ListIndex = 9 Then
        Call ListViewHeader10
    End If
    
End Sub

Private Sub Option34()
    
    lblCurrent = ""
    lblTotalRecord = ""
    ListView2.ListItems.Clear
    
    If cboReportType.ListIndex = 0 Then
        Call ListViewHeader15
    ElseIf cboReportType.ListIndex = 5 Then
        Call ListViewHeader18
    ElseIf cboReportType.ListIndex = 7 Then
        Call ListViewHeader21
    ElseIf cboReportType.ListIndex = 1 Then
        Call ListViewHeader3
    ElseIf cboReportType.ListIndex = 3 Then
        Call ListViewHeader17
    ElseIf cboReportType.ListIndex = 4 Then
        Call ListViewHeader7
    ElseIf cboReportType.ListIndex = 9 Then
        Call ListViewHeader10
    End If
    
End Sub

Private Sub Option35()
    
    lblCurrent = ""
    lblTotalRecord = ""
    ListView2.ListItems.Clear
    
    If cboReportType.ListIndex = 0 Then
        Call ListViewHeader15
    ElseIf cboReportType.ListIndex = 5 Then
        Call ListViewHeader18
    ElseIf cboReportType.ListIndex = 7 Then
        Call ListViewHeader21
    ElseIf cboReportType.ListIndex = 1 Then
        Call ListViewHeader3
    ElseIf cboReportType.ListIndex = 3 Then
        Call ListViewHeader17
    ElseIf cboReportType.ListIndex = 4 Then
        Call ListViewHeader7
    ElseIf cboReportType.ListIndex = 9 Then
        Call ListViewHeader10
    End If
    
End Sub

Private Sub Option37()
    
    lblCurrent = ""
    lblTotalRecord = ""
    ListView2.ListItems.Clear
    
    If cboReportType.ListIndex = 5 Then
        Call ListViewHeader18
    ElseIf cboReportType.ListIndex = 7 Then
        Call ListViewHeader21
    ElseIf cboReportType.ListIndex = 1 Then
        Call ListViewHeader3
    ElseIf cboReportType.ListIndex = 3 Then
        Call ListViewHeader17
    ElseIf cboReportType.ListIndex = 4 Then
        Call ListViewHeader7
    ElseIf cboReportType.ListIndex = 9 Then
        Call ListViewHeader10
    End If
    
End Sub

Private Sub Option4()
    TxtBordxNo1.Enabled = True
    TxtBordxNo2.Enabled = True
    ChkBordx.Enabled = True
    TxtBordxDate1.Enabled = True
    TxtBordxDate2.Enabled = True
    ChkBordxDate.Enabled = True
    opt35 = True
        
    If (cboReportType.ListIndex = 0 And opt35 = True) Or _
       (cboReportType.ListIndex = 0 And opt34 = True) Then
        Call ListViewHeader15
    ElseIf (cboReportType.ListIndex = 5 And opt35 = True) Or _
       (cboReportType.ListIndex = 5 And opt34 = True) Or _
       (cboReportType.ListIndex = 5 And opt33 = True) Or _
       (cboReportType.ListIndex = 5 And opt29 = True) Or _
       (cboReportType.ListIndex = 5 And opt37 = True) Then
        Call ListViewHeader18
    
    ElseIf (cboReportType.ListIndex = 1 And opt35 = True) Or _
       (cboReportType.ListIndex = 1 And opt34 = True) Or _
       (cboReportType.ListIndex = 1 And opt33 = True) Or _
       (cboReportType.ListIndex = 1 And opt29 = True) Or _
       (cboReportType.ListIndex = 1 And opt37 = True) Or _
       (cboReportType.ListIndex = 1 And opt30 = True) Then
        Call ListViewHeader3
    
    ElseIf (cboReportType.ListIndex = 3 And opt35 = True) Or _
       (cboReportType.ListIndex = 3 And opt34 = True) Or _
       (cboReportType.ListIndex = 3 And opt33 = True) Or _
       (cboReportType.ListIndex = 3 And opt29 = True) Or _
       (cboReportType.ListIndex = 3 And opt37 = True) Or _
       (cboReportType.ListIndex = 3 And opt31 = True) Or _
       (cboReportType.ListIndex = 3 And opt30 = True) Then
        Call ListViewHeader17
    
    ElseIf (cboReportType.ListIndex = 2 And opt29 = True) Then
        Call ListViewHeader8
    
    ElseIf cboReportType.ListIndex = 6 Then
        Call ListViewHeader20
    
    ElseIf cboReportType.ListIndex = 8 Then
        Call ListViewHeader5
    
    ElseIf (cboReportType.ListIndex = 4 And opt35 = True) Or _
       (cboReportType.ListIndex = 4 And opt34 = True) Or _
       (cboReportType.ListIndex = 4 And opt33 = True) Or _
       (cboReportType.ListIndex = 4 And opt29 = True) Or _
       (cboReportType.ListIndex = 4 And opt37 = True) Or _
       (cboReportType.ListIndex = 4 And opt30 = True) Or _
       (cboReportType.ListIndex = 4 And Opt8 = True) Or _
       (cboReportType.ListIndex = 4 And Opt11 = True) Or _
       (cboReportType.ListIndex = 4 And opt12 = True) Or _
       (cboReportType.ListIndex = 4 And opt13 = True) Or _
       (cboReportType.ListIndex = 4 And Opt7 = True) Or _
       (cboReportType.ListIndex = 4 And opt16 = True) Then
        Call ListViewHeader7
        
    End If
End Sub

Private Sub rptAgeband()
    lblCurrent = ""
    lblTotalRecord = ""
    ListView2.ListItems.Clear
    If cboClaimCase.ListIndex = 0 Then
        cboReportSelection.AddItem "Summary"      '35
        cboReportSelection.AddItem "By Gender"   '34
        cboReportSelection.AddItem "By Insured Type"     '33
        cboReportSelection.AddItem "By Race"     '29
        cboReportSelection.AddItem "By Plan No"      '37
        cboReportSelection.AddItem "By Bordx Month"     '30
        'cboReportSelection.AddItem "By Bordx Month(Payor)"     '16
        opt35 = True
        Call ListViewHeader3
    ElseIf cboClaimCase.ListIndex = 1 Then
        lblCurrent = ""
        lblTotalRecord = ""
        ListView2.ListItems.Clear
        cboReportSelection.AddItem "Summary"     '35
        cboReportSelection.AddItem "By Gender"   '34
        cboReportSelection.AddItem "By Insured Type"     '33
        cboReportSelection.AddItem "By Race"     '29
        cboReportSelection.AddItem "By Plan No"      '37
        opt35 = True
        Call ListViewHeader3
    End If
End Sub

Private Sub rptState()
    lblCurrent = ""
    lblTotalRecord = ""
    ListView2.ListItems.Clear
    Call ListViewHeader5
End Sub

Private Sub Option7()
    lblCurrent = ""
    lblTotalRecord = ""
    ListView2.ListItems.Clear
    
    If cboReportType.ListIndex = 4 Then
        Call ListViewHeader7
    End If
End Sub

Private Sub Option8()
    lblCurrent = ""
    lblTotalRecord = ""
    ListView2.ListItems.Clear
    
    If cboReportType.ListIndex = 4 Then
        Call ListViewHeader7
    End If
End Sub

Private Sub rptPayor()
    lblCurrent = ""
    lblTotalRecord = ""
    ListView2.ListItems.Clear
    Call ListViewHeader20
End Sub

'=========== Check Date Format ===========================

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
'=========== Check Date Format ===========================

Private Function SearchRecords()
Dim sql As String
Dim StrAppend As String

    StrAppend = ""
    GrpCompany = ""
        
    If Len(cboClaimCase) = False Then
        MsgBox "Please select a Claim or Case Report.", vbOKOnly + vbCritical, DialogBoxTitle
        Exit Function
    End If
    
    If Len(Mid(cboHLTCode, 1, 1)) Then
        StrAppend = StrAppend + " AND HLTCode ='" & Trim(Mid(cboHLTCode, 1, 1)) & "'"
    End If

    If Len(Mid(CboInsType, 1, 1)) Then
        StrAppend = StrAppend & " AND INSCode = '" & Trim(Mid(CboInsType, 1, 1)) & "'"
    End If
       
    If Trim(Mid(CboMemType, 1, 1)) = "P" And Trim(cboSuppType) = "Both" Then
        StrAppend = StrAppend & " AND (MBMMemberType = 'P' Or MBMMemberType = 'Q') And MBMPatientCovID = '00' "
    ElseIf Trim(Mid(CboMemType, 1, 1)) = "P" And Trim(Mid(cboSuppType, 1, 1)) = "C" Then
        StrAppend = StrAppend & " AND MBMMemberType = 'P'  And MBMPatientCovID = '00'"
    ElseIf Trim(Mid(CboMemType, 1, 1)) = "P" And Trim(Mid(cboSuppType, 1, 1)) = "S" Then
        StrAppend = StrAppend & " AND MBMMemberType = 'Q' And MBMPatientCovID = '00'"
    ElseIf Trim(Mid(CboMemType, 1, 1)) = "S" And Trim(cboSuppType) = "Both" Then
        StrAppend = StrAppend & " AND (MBMCMemberType = 'S' OR MBMCMemberType = 'C')"
    ElseIf Trim(Mid(CboMemType, 1, 1)) = "S" And Trim(Mid(cboSuppType, 1, 1)) = "C" Then
        StrAppend = StrAppend & " AND MBMCMemberType = 'C'"
    ElseIf Trim(Mid(CboMemType, 1, 1)) = "S" And Trim(Mid(cboSuppType, 1, 1)) = "S" Then
        StrAppend = StrAppend & " AND MBMCMemberType = 'S'"
    ElseIf Trim(CboMemType) = "Both" And Trim(Mid(cboSuppType, 1, 1)) = "C" Then
        StrAppend = StrAppend & " AND (MBMMemberType = 'P' OR MBMCMemberType = 'C')"
    ElseIf Trim(CboMemType) = "Both" And Trim(Mid(cboSuppType, 1, 1)) = "S" Then
        StrAppend = StrAppend & " AND (MBMMemberType = 'Q' OR MBMCMemberType = 'S')"
    End If
    
    If Len(Mid(cboDataStatus, 1, 1)) Then
        StrAppend = StrAppend & " AND MBMDataStatus = '" & Trim(Mid(cboDataStatus, 1, 1)) & "'"
    End If
    
    If Len(Mid(CboClaimability, 1, 1)) Then
        StrAppend = StrAppend & " AND CASClaimability = '" & Trim(Mid(CboClaimability, 1, 1)) & "'"
    End If

    If Len(Mid(CboStatus, 1, 1)) Then
        StrAppend = StrAppend & " AND MBMStatus = '" & Trim(Mid(CboStatus, 1, 1)) & "'"
    End If

    If Mid(CboCaseStatus, 1, 1) = 0 Then
        StrAppend = StrAppend & " AND CASStatus = 'O'"
    ElseIf Mid(CboCaseStatus, 1, 1) = 1 Then
        StrAppend = StrAppend & " AND (CASStatus = 'O' OR CASStatus = 'C')"
    ElseIf Mid(CboCaseStatus, 1, 1) = 2 Then
        StrAppend = StrAppend & " AND CASStatus = 'D'"
    End If
    
    If Len(Trim(txtAgentCode)) Then
       StrAppend = StrAppend + " And  MBMAgentCode like '%" & ChkStr(txtAgentCode) & "%' "
    End If
    
    If Len(Trim(txtPolNo)) Then
       StrAppend = StrAppend + " And  MBMPolicyNo like '%" & ChkStr(txtPolNo) & "%' "
    End If
    
    If Len(Mid(cboPayorCode, 1, 2)) Then
        StrAppend = StrAppend & " AND PAYCode = '" & Trim(Mid(cboPayorCode, 1, 2)) & "'"
    Else
        'StrAppend = StrAppend & " AND Substring(PAYCode,1,1) <> 'Z'"
    End If
           
    If Len(Trim(CboGrpCom)) Then
        GrpCompany = Replace(CboGrpCom, "'", "''")
        StrAppend = StrAppend + " AND MBMGroupCompany = '" & Trim(GrpCompany) & "'"
    End If
    
    If Len(Trim(TxtDOA1)) > 0 And IsDate(TxtDOA1) = True Then
        StrAppend = StrAppend + " And CasDateAdmitted >= '" & Format(TxtDOA1, "mm/dd/yyyy") & "'"
    End If
    
    If Len(Trim(TxtDOA2)) > 0 And IsDate(TxtDOA2) = True Then
        StrAppend = StrAppend + " And CasDateAdmitted <= '" & Format(TxtDOA2, "mm/dd/yyyy") & "'"
    End If
    
    If Len(Trim(txtDOD1)) > 0 And IsDate(txtDOD1) = True Then
        StrAppend = StrAppend + " And CasDischargeDate >= '" & Format(txtDOD1, "mm/dd/yyyy") & "'"
    End If
    
    If Len(Trim(txtDOD2)) > 0 And IsDate(txtDOD2) = True Then
        StrAppend = StrAppend + " And CasDischargeDate <= '" & Format(txtDOD2, "mm/dd/yyyy") & "'"
    End If
    
    If Len(Trim(TxtDOR1)) > 0 And IsDate(TxtDOR1) = True Then
        StrAppend = StrAppend + " And CaseRegDate >= '" & Format(TxtDOR1, "mm/dd/yyyy") & "'"
    End If
    
    If Len(Trim(TxtDOR2)) > 0 And IsDate(TxtDOR2) = True Then
        StrAppend = StrAppend + " And CaseRegDate <= '" & Format(TxtDOR2, "mm/dd/yyyy") & "'"
    End If
    
    If Len(Trim(txtFinalDiag1)) Then
        StrAppend = StrAppend + " And FinalDiag >= '" & Trim(txtFinalDiag1) & "'"
    End If
    
    If Len(Trim(txtFinalDiag2)) Then
        StrAppend = StrAppend + " And FinalDiag <= '" & Trim(txtFinalDiag2) & "'"
    End If
    
    If Mid(CboCasePending, 1, 1) = "N" Then
        StrAppend = StrAppend + " AND (CASPending = 'N' OR CASPending = '' OR CASPending IS NULL)"
    ElseIf Mid(CboCasePending, 1, 1) = "Y" Then
        StrAppend = StrAppend + " AND CASPending = 'Y'"
    End If
    
    If Mid(CboClmPending, 1, 1) = "B" Then
        StrAppend = StrAppend + " AND CLMPending = 'B'"
    ElseIf Mid(CboClmPending, 1, 1) = "C" Then
        StrAppend = StrAppend + " AND CLMPending = 'C'"
    ElseIf Mid(CboClmPending, 1, 1) = "P" Then
        StrAppend = StrAppend + " AND CLMPending = 'P'"
    End If
    
    If Len(Trim(TxtPlan1)) Then
        StrAppend = StrAppend & " AND PLNCode >= '" & Trim(TxtPlan1) & "'"
    End If
    
    If Len(Trim(TxtPlan2)) Then
        StrAppend = StrAppend & " AND PLNCode <= '" & Trim(TxtPlan2) & "'"
    End If
    
    If Len(Trim(TxtBordxNo1)) Then
        StrAppend = StrAppend & " AND BordxRefNo >= '" & Trim(TxtBordxNo1) & "'"
    End If
    
    If Len(Trim(TxtBordxNo2)) Then
        StrAppend = StrAppend & " AND BordxRefNo <= '" & Trim(TxtBordxNo2) & "'"
    End If
    
    If Len(Trim(TxtPayEffDate1)) > 0 And IsDate(TxtPayEffDate1) = True Then
        StrAppend = StrAppend + " And MBMPayorEffDate >= '" & Format(TxtPayEffDate1, "mm/dd/yyyy") & "'"
    End If
    
    If Len(Trim(TxtPayEffDate2)) > 0 And IsDate(TxtPayEffDate2) = True Then
        StrAppend = StrAppend + " And MBMPayorEffDate <= '" & Format(TxtPayEffDate2, "mm/dd/yyyy") & "'"
    End If
    
    If Len(Trim(TxtPayExpDate1)) And IsDate(TxtPayExpDate1) = True Then
        StrAppend = StrAppend & " AND MBMPayorExpDate >= '" & Format(TxtPayExpDate1, "mm/dd/yyyy") & "'"
    End If
    
    If Len(Trim(TxtPayExpDate2)) And IsDate(TxtPayExpDate2) = True Then
        StrAppend = StrAppend & " AND MBMPayorExpDate <= '" & Format(TxtPayExpDate2, "mm/dd/yyyy") & "'"
    End If
    
    If Len(Trim(txtDateJoin1)) And IsDate(txtDateJoin1) = True Then
        StrAppend = StrAppend & " AND MBMFirstJoinedDate >= '" & Format(txtDateJoin1, "mm/dd/yyyy") & "'"
    End If
    
    If Len(Trim(txtDateJoin2)) And IsDate(txtDateJoin2) = True Then
        StrAppend = StrAppend & " AND MBMFirstJoinedDate <= '" & Format(txtDateJoin2, "mm/dd/yyyy") & "'"
    End If
    
    If Len(Trim(TxtBordxDate1)) And IsDate(TxtBordxDate1) = True Then
        StrAppend = StrAppend & " AND BordxDate >= '" & Format(TxtBordxDate1, "mm/dd/yyyy") & "'"
    End If
    
    If Len(Trim(TxtBordxDate2)) And IsDate(TxtBordxDate2) = True Then
        StrAppend = StrAppend & " AND BordxDate <= '" & Format(TxtBordxDate2, "mm/dd/yyyy") & "'"
    End If
    
    If Len(Trim(TxtCaseNo1)) Then
       StrAppend = StrAppend + " AND CASNumber >= '" & RemStr(Trim(TxtCaseNo1)) & "'"
    End If
    
    If Len(Trim(TxtCaseNo2)) Then
       StrAppend = StrAppend + " AND CASNumber <= '" & RemStr(Trim(TxtCaseNo2)) & "'"
    End If
    
    If Len(Trim(TxtVersion1)) Then
       StrAppend = StrAppend + " AND ClaimVersion >= '" & RemStr(Trim(TxtVersion1)) & "'"
    End If
    
    If Len(Trim(TxtVersion2)) Then
       StrAppend = StrAppend + " AND ClaimVersion <= '" & RemStr(Trim(TxtVersion2)) & "'"
    End If
    
    If ChkHLT.Value = 1 Then
        StrAppend = StrAppend & " AND (HLTCode Is null OR HLTCode = '')"
    End If
    
    If ChkINS.Value = 1 Then
        StrAppend = StrAppend & " AND (INSCode Is null OR INSCode = '')"
    End If
    
    If ChkDataStatus.Value = 1 Then
        StrAppend = StrAppend & " AND (MBMDataStatus Is null OR MBMDataStatus = '')"
    End If
    
    If ChkClaimability.Value = 1 Then
        StrAppend = StrAppend & " AND (CASClaimability Is null OR CASClaimability = '')"
    End If
    
    If ChkPolStatus.Value = 1 Then
        StrAppend = StrAppend & " AND (MBMStatus Is null OR MBMStatus = '')"
    End If
    
    If ChkCaseStatus.Value = 1 Then
        StrAppend = StrAppend & " AND (CASStatus Is null OR CASStatus = '')"
    End If
    
    If ChkAgentCode.Value = 1 Then
        StrAppend = StrAppend & " AND (MBMAgentCode Is null OR MBMAgentCode = '')"
    End If
    
    If ChkPolNo.Value = 1 Then
        StrAppend = StrAppend & " AND (MBMPolicyNo Is null OR MBMPolicyNo = '')"
    End If
    
    If ChkPayor.Value = 1 Then
        StrAppend = StrAppend & " AND (PAYCode Is null OR PAYCode = '')"
    End If
    
    If ChkGrpCom.Value = 1 Then
        StrAppend = StrAppend & " AND (MBMGroupCompany Is null OR MBMGroupCompany = '')"
    End If
            
    If ChkDOA.Value = 1 Then
        StrAppend = StrAppend & " AND (CASDateAdmitted Is null OR CASDateAdmitted = '')"
    End If
    
    If ChkDOD.Value = 1 Then
        StrAppend = StrAppend & " AND (CASDischargeDate Is null OR CASDischargeDate = '')"
    End If
    
    If chkDOR.Value = 1 Then
        StrAppend = StrAppend & " AND (CaseRegDate Is null OR CaseRegDate = '')"
    End If
    
    If ChkFinal.Value = 1 Then
        StrAppend = StrAppend & " AND (FinalDiag Is null OR FinalDiag = '')"
    End If
    
    If ChkPlan.Value = 1 Then
        StrAppend = StrAppend & " AND (PLNCode Is null OR PLNCode = '')"
    End If
    
    If ChkBordx.Value = 1 Then
        StrAppend = StrAppend & " AND (BordxRefNo Is null OR BordxRefNo = '')"
    End If
    
    If ChkEffDate.Value = 1 Then
        StrAppend = StrAppend & " AND (MBMPayorEffDate Is null OR MBMPayorEffDate = '')"
    End If
    
    If chkExpDate.Value = 1 Then
        StrAppend = StrAppend & " AND (MBMPayorExpDate Is null OR MBMPayorExpDate = '')"
    End If
    
    If ChkJoinDate.Value = 1 Then
        StrAppend = StrAppend & " AND (MBMFirstJoinedDate Is null OR MBMFirstJoinedDate = '')"
    End If
    
    If ChkBordxDate.Value = 1 Then
        StrAppend = StrAppend & " AND (BordxDate Is null OR BordxDate = '')"
    End If
    
    If Len(Trim(Mid(CboStayType, 1, 1))) Then
        StrAppend = StrAppend + " And CASStayType = '" & Trim(Mid(CboStayType, 1, 1)) & "'"
    End If
    
    If Len(Trim(Mid(CboCashType, 1, 1))) Then
        StrAppend = StrAppend + " And CASPayType = '" & Trim(Mid(CboCashType, 1, 1)) & "'"
    End If
    
    If Len(Trim(Mid(CboModality, 1, 1))) Then
        StrAppend = StrAppend + " And ModalityCode = '" & Trim(Mid(CboModality, 1, 1)) & "'"
    End If
    
    If cboClaimCase.ListIndex = 0 And cboReportType.ListIndex = 0 And opt35 = True Or _
           cboClaimCase.ListIndex = 1 And cboReportType.ListIndex = 0 And opt35 = True Then
        Call ResultListing(StrAppend)
    ElseIf cboClaimCase.ListIndex = 0 And cboReportType.ListIndex = 5 And opt35 = True Or _
           cboClaimCase.ListIndex = 1 And cboReportType.ListIndex = 5 And opt35 = True Then
        Call INSListing(StrAppend)
    ElseIf cboClaimCase.ListIndex = 0 And cboReportType.ListIndex = 5 And opt34 = True Or _
           cboClaimCase.ListIndex = 1 And cboReportType.ListIndex = 5 And opt34 = True Then
        Call PlanInsuredListing(StrAppend)
    ElseIf cboClaimCase.ListIndex = 0 And cboReportType.ListIndex = 5 And opt33 = True Or _
           cboClaimCase.ListIndex = 1 And cboReportType.ListIndex = 5 And opt33 = True Then
        Call AgeInsuredListing(StrAppend)
    ElseIf cboClaimCase.ListIndex = 0 And cboReportType.ListIndex = 5 And opt29 = True Or _
           cboClaimCase.ListIndex = 1 And cboReportType.ListIndex = 5 And opt29 = True Then
        Call SexInsuredListing(StrAppend)
    ElseIf cboClaimCase.ListIndex = 0 And cboReportType.ListIndex = 5 And opt37 = True Or _
           cboClaimCase.ListIndex = 1 And cboReportType.ListIndex = 5 And opt37 = True Then
        Call RaceInsuredListing(StrAppend)
    ElseIf cboClaimCase.ListIndex = 0 And cboReportType.ListIndex = 7 And opt31 = True Or _
           cboClaimCase.ListIndex = 1 And cboReportType.ListIndex = 7 And opt31 = True Then
        Call PlanPremiumListing(StrAppend)
    ElseIf cboClaimCase.ListIndex = 0 And cboReportType.ListIndex = 7 And opt35 = True Or _
           cboClaimCase.ListIndex = 1 And cboReportType.ListIndex = 7 And opt35 = True Then
        Call PlanListing(StrAppend)
    ElseIf cboClaimCase.ListIndex = 0 And cboReportType.ListIndex = 7 And opt34 = True Or _
           cboClaimCase.ListIndex = 1 And cboReportType.ListIndex = 7 And opt34 = True Then
        Call PlanTopHosListing(StrAppend)
    ElseIf cboClaimCase.ListIndex = 0 And cboReportType.ListIndex = 7 And opt33 = True Or _
           cboClaimCase.ListIndex = 1 And cboReportType.ListIndex = 7 And opt33 = True Then
        Call PlanTopDiagListing(StrAppend)
    ElseIf cboClaimCase.ListIndex = 0 And cboReportType.ListIndex = 7 And opt29 = True Or _
           cboClaimCase.ListIndex = 1 And cboReportType.ListIndex = 7 And opt29 = True Then
        Call PlanAgebandListing(StrAppend)
    ElseIf cboClaimCase.ListIndex = 0 And cboReportType.ListIndex = 7 And opt37 = True Or _
           cboClaimCase.ListIndex = 1 And cboReportType.ListIndex = 7 And opt37 = True Then
        Call PlanInsuredListing(StrAppend)
    ElseIf cboClaimCase.ListIndex = 0 And cboReportType.ListIndex = 7 And opt30 = True Or _
           cboClaimCase.ListIndex = 1 And cboReportType.ListIndex = 7 And opt30 = True Then
        Call PlanInsuredListing2(StrAppend)
    ElseIf cboClaimCase.ListIndex = 0 And cboReportType.ListIndex = 1 And opt35 = True Or _
           cboClaimCase.ListIndex = 1 And cboReportType.ListIndex = 1 And opt35 = True Then
        Call AgeGroupListing(StrAppend)
    ElseIf cboClaimCase.ListIndex = 0 And cboReportType.ListIndex = 1 And opt34 = True Or _
           cboClaimCase.ListIndex = 1 And cboReportType.ListIndex = 1 And opt34 = True Then
        Call AgeGenderListing(StrAppend)
    ElseIf cboClaimCase.ListIndex = 0 And cboReportType.ListIndex = 1 And opt33 = True Or _
           cboClaimCase.ListIndex = 1 And cboReportType.ListIndex = 1 And opt33 = True Then
        Call AgeInsuredListing(StrAppend)
    ElseIf cboClaimCase.ListIndex = 0 And cboReportType.ListIndex = 1 And opt29 = True Or _
           cboClaimCase.ListIndex = 1 And cboReportType.ListIndex = 1 And opt29 = True Then
        Call AgeRaceListing(StrAppend)
    ElseIf cboClaimCase.ListIndex = 0 And cboReportType.ListIndex = 1 And opt37 = True Or _
           cboClaimCase.ListIndex = 1 And cboReportType.ListIndex = 1 And opt37 = True Then
        Call AgePlanListing(StrAppend)
    ElseIf cboClaimCase.ListIndex = 0 And cboReportType.ListIndex = 1 And opt30 = True Or _
           cboClaimCase.ListIndex = 1 And cboReportType.ListIndex = 1 And opt30 = True Then
        Call AgeIllnessListing(StrAppend)
    ElseIf cboClaimCase.ListIndex = 0 And cboReportType.ListIndex = 6 Or _
           cboClaimCase.ListIndex = 1 And cboReportType.ListIndex = 6 Then
        Call PayorListing(StrAppend)
    ElseIf (cboClaimCase.ListIndex = 0 And cboReportType.ListIndex = 3 And opt29 = True) Or _
           (cboClaimCase.ListIndex = 0 And cboReportType.ListIndex = 3 And opt37 = True) Or _
           (cboClaimCase.ListIndex = 1 And cboReportType.ListIndex = 3 And opt29 = True) Or _
           (cboClaimCase.ListIndex = 1 And cboReportType.ListIndex = 3 And opt37 = True) Then
        Call HOSTop20(StrAppend)
    ElseIf (cboClaimCase.ListIndex = 0 And cboReportType.ListIndex = 3 And opt31 = True) Or _
           (cboClaimCase.ListIndex = 1 And cboReportType.ListIndex = 3 And opt31 = True) Then
        Call HOSTop202(StrAppend)
    ElseIf cboClaimCase.ListIndex = 0 And cboReportType.ListIndex = 3 And opt34 = True Or _
           cboClaimCase.ListIndex = 1 And cboReportType.ListIndex = 3 And opt34 = True Then
        Call HOSSumListing(StrAppend)
    ElseIf cboClaimCase.ListIndex = 0 And cboReportType.ListIndex = 3 And opt33 = True Or _
           cboClaimCase.ListIndex = 1 And cboReportType.ListIndex = 3 And opt33 = True Then
        Call HOSSumListing2(StrAppend)
    ElseIf cboClaimCase.ListIndex = 0 And cboReportType.ListIndex = 3 And opt35 = True Or _
           cboClaimCase.ListIndex = 1 And cboReportType.ListIndex = 3 And opt35 = True Then
        Call HOSDetailListing(StrAppend)
    ElseIf cboClaimCase.ListIndex = 0 And cboReportType.ListIndex = 3 And opt30 = True Or _
           cboClaimCase.ListIndex = 1 And cboReportType.ListIndex = 3 And opt30 = True Then
        Call HOSByDiag(StrAppend)
    ElseIf cboClaimCase.ListIndex = 0 And cboReportType.ListIndex = 2 And opt35 = True Or _
           cboClaimCase.ListIndex = 1 And cboReportType.ListIndex = 2 And opt35 = True Then
        Call FinalDiagListing(StrAppend)
    ElseIf cboClaimCase.ListIndex = 0 And cboReportType.ListIndex = 2 And opt34 = True Or _
           cboClaimCase.ListIndex = 1 And cboReportType.ListIndex = 2 And opt34 = True Then
        Call FinalDiagListing(StrAppend)
    ElseIf cboClaimCase.ListIndex = 0 And cboReportType.ListIndex = 2 And opt33 = True Or _
           cboClaimCase.ListIndex = 1 And cboReportType.ListIndex = 2 And opt33 = True Then
        Call FinalDiagListing(StrAppend)
    ElseIf cboClaimCase.ListIndex = 0 And cboReportType.ListIndex = 2 And opt29 = True Or _
           cboClaimCase.ListIndex = 1 And cboReportType.ListIndex = 2 And opt29 = True Then
        Call FinalDiagbyHos(StrAppend)
    ElseIf (cboReportType.ListIndex = 8 And cboClaimCase.ListIndex = 0) Or _
           (cboReportType.ListIndex = 8 And cboClaimCase.ListIndex = 1) Then
        Call StateListing(StrAppend)
    ElseIf cboClaimCase.ListIndex = 0 And cboReportType.ListIndex = 4 And opt35 = True Or _
           cboClaimCase.ListIndex = 1 And cboReportType.ListIndex = 4 And opt35 = True Then
        Call RegMonthAllPayor(StrAppend)
    ElseIf cboClaimCase.ListIndex = 0 And cboReportType.ListIndex = 4 And Opt8 = True Or _
           cboClaimCase.ListIndex = 1 And cboReportType.ListIndex = 4 And Opt8 = True Then
        Call RegMonthPayor(StrAppend)
    ElseIf cboClaimCase.ListIndex = 0 And cboReportType.ListIndex = 4 And opt34 = True Or _
           cboClaimCase.ListIndex = 1 And cboReportType.ListIndex = 4 And opt34 = True Then
        Call EffMonthAllPayor(StrAppend)
    ElseIf cboClaimCase.ListIndex = 0 And cboReportType.ListIndex = 4 And Opt11 = True Or _
           cboClaimCase.ListIndex = 1 And cboReportType.ListIndex = 4 And Opt11 = True Then
        Call EffMonthPayor(StrAppend)
    ElseIf cboClaimCase.ListIndex = 0 And cboReportType.ListIndex = 4 And opt33 = True Or _
           cboClaimCase.ListIndex = 1 And cboReportType.ListIndex = 4 And opt33 = True Then
        Call ExpMonthAllPayor(StrAppend)
    ElseIf cboClaimCase.ListIndex = 0 And cboReportType.ListIndex = 4 And opt12 = True Or _
           cboClaimCase.ListIndex = 1 And cboReportType.ListIndex = 4 And opt12 = True Then
        Call ExpMonthPayor(StrAppend)
    ElseIf cboClaimCase.ListIndex = 0 And cboReportType.ListIndex = 4 And opt29 = True Or _
           cboClaimCase.ListIndex = 1 And cboReportType.ListIndex = 4 And opt29 = True Then
        Call DOAMonthAllPayor(StrAppend)
    ElseIf cboClaimCase.ListIndex = 0 And cboReportType.ListIndex = 4 And opt13 = True Or _
           cboClaimCase.ListIndex = 1 And cboReportType.ListIndex = 4 And opt13 = True Then
        Call DOAMonthPayor(StrAppend)
    ElseIf cboClaimCase.ListIndex = 0 And cboReportType.ListIndex = 4 And opt37 = True Or _
           cboClaimCase.ListIndex = 1 And cboReportType.ListIndex = 4 And opt37 = True Then
        Call DODMonthAllPayor(StrAppend)
    ElseIf cboClaimCase.ListIndex = 0 And cboReportType.ListIndex = 4 And Opt7 = True Or _
           cboClaimCase.ListIndex = 1 And cboReportType.ListIndex = 4 And Opt7 = True Then
        Call DODMonthPayor(StrAppend)
    ElseIf cboClaimCase.ListIndex = 0 And cboReportType.ListIndex = 4 And opt30 = True Then
        Call BordxMonthAllPayor(StrAppend)
    ElseIf cboClaimCase.ListIndex = 0 And cboReportType.ListIndex = 4 And opt16 = True Then
        Call BordxMonthPayor(StrAppend)
    ElseIf cboClaimCase.ListIndex = 0 And cboReportType.ListIndex = 9 And opt35 = True Then
        Call DIREfficiencyRpt(StrAppend)
    ElseIf cboClaimCase.ListIndex = 0 And cboReportType.ListIndex = 9 And opt34 = True Then
        Call BordxEfficiencyRpt(StrAppend)
    ElseIf cboClaimCase.ListIndex = 0 And cboReportType.ListIndex = 9 And opt33 = True Then
        Call InvEfficiencyRpt(StrAppend)
    ElseIf cboClaimCase.ListIndex = 0 And cboReportType.ListIndex = 9 And opt29 = True Then
        Call BordxInvEfficiencyRpt(StrAppend)
    ElseIf cboClaimCase.ListIndex = 0 And cboReportType.ListIndex = 9 And opt37 = True Then
        Call OSBordxEfficiencyRpt(StrAppend)
    End If
End Function

Private Sub HOSTop20(Optional StrAppend As String)
Dim rst2 As New ADODB.Recordset
Dim cmd As New ADODB.Command
Dim prm1 As ADODB.Parameter
Dim prm2 As ADODB.Parameter
Dim strAppendCase As String
Dim StrAppend1 As String
Dim Record As ListItem

    Payor = ""
    Hospital = ""
    HOspitals = ""
    lblCurrent = 0
    ListView2.ListItems.Clear
        
    If cboClaimCase.ListIndex = 0 Then
        strAppendCase = "55"
    'sql = " SELECT TOP 20 PAYCode, HOSName, COUNT(HOSName) AS Total, SUM(NoofDay) AS NoofDay, " & _
          " SUM(CLMTotalActual) AS TotalActual, SUM(CLMTotalApproved) AS TotalApproved" & _
          " From VIEW_CLMReportDCU Where  0 = 0"
          
    ElseIf cboClaimCase.ListIndex = 1 Then
        strAppendCase = "56"
    'sql = " SELECT TOP 20 PAYCode, HOSName, COUNT(DISTINCT CASNumber) AS Total, " & _
          " SUM(CLMTotalActual) AS TotalActual, SUM(CLMTotalApproved) AS TotalApproved" & _
          " From VIEW_CLMReportDCU Where  0 = 0 "
          
    End If
    
    'If opt29 = True Then
    '    sql2 = " GROUP BY PAYCode, HOSName ORDER BY Total DESC "
    'ElseIf opt37 = True Then
    '    sql2 = " GROUP BY PAYCode, HOSName ORDER BY TotalActual DESC "
    'End If
    
    'sql = sql + StrAppend + sql2
    
    'rst2.Open sql, medixmasterconnWS, adOpenForwardOnly, adLockBatchOptimistic
    Set cmd.ActiveConnection = medixmasterconnWS
        cmd.CommandText = "DCUReport"
        cmd.CommandType = adCmdStoredProc
        cmd.CommandTimeout = 300
        Set prm1 = cmd.CreateParameter("StrAppend", adVarChar, adParamInput, 2000, StrAppend)
        cmd.Parameters.Append prm1
        Set prm2 = cmd.CreateParameter("StrAppendCase", adVarChar, adParamInput, 2000, strAppendCase)
        cmd.Parameters.Append prm2
        rst2.CursorLocation = adUseClient
        rst2.CursorType = adOpenForwardOnly
        rst2.CacheSize = 100
        Set rst2 = cmd.Execute
        
        Do Until rst2.EOF
    
            With rst2
                
                If Len(rst2!PAYCode) Then
                    Set Record = ListView2.ListItems.Add(, , rst2!PAYCode)
                    Payor = Trim(rst2!PAYCode)
                Else
                    Set Record = ListView2.ListItems.Add(, , "Empty")
                    Payor = ""
                End If
                
                If Len(rst2!HOSName) Then
                    Record.SubItems(1) = Trim(rst2!HOSName)
                    Hospital = Trim(rst2!HOSName)
                Else
                    Record.SubItems(1) = "Empty"
                    Hospital = ""
                End If
                                
                If Len(rst2!Total) Then
                    Record.SubItems(2) = rst2!Total
                End If
                               
                If cboClaimCase.ListIndex = 0 Then
                
                    If rst2!NoofDay = "0" Then
                        Record.SubItems(3) = "1"
                    ElseIf Len(rst2!NoofDay) Then
                        Record.SubItems(3) = rst2!NoofDay
                    End If
                    
                ElseIf cboClaimCase.ListIndex = 1 Then
                
                    HOspitals = Replace(Hospital, "'", "''")
                    
                    If Payor = "" And HOspitals = "" Then
                        strAppendCase = "57"
                        StrAppend1 = " AND (PAYCode = '' OR PAYCode IS NULL) AND " & _
                           " (HOSName = '' OR HOSName IS NULL)"
                        StrAppend1 = StrAppend1 + StrAppend
                    'sql3 = " SELECT PAYCode, HOSName, SUM(DurationStay) AS NoofDay " & _
                           " From VIEW_CLMReportDCU2 Where (PAYCode = '' OR PAYCode IS NULL) AND " & _
                           " (HOSName = '' OR HOSName IS NULL)"
                              
                    ElseIf Payor <> "" And HOspitals = "" Then
                        strAppendCase = "57"
                        StrAppend1 = " AND PAYCode = '" & Trim(Payor) & "' AND " & _
                           " (HOSName = '' OR HOSName IS NULL)"
                        StrAppend1 = StrAppend1 + StrAppend
                    'sql3 = " SELECT PAYCode, HOSName, SUM(DurationStay) AS NoofDay " & _
                           " From VIEW_CLMReportDCU2 Where PAYCode = '" & Trim(Payor) & "' AND " & _
                           " (HOSName = '' OR HOSName IS NULL)"
                              
                    ElseIf Payor = "" And HOspitals <> "" Then
                        strAppendCase = "57"
                        StrAppend1 = " AND (PAYCode = '' OR PAYCode IS NULL) AND " & _
                           " HOSName = '" & Trim(HOspitals) & "'"
                        StrAppend1 = StrAppend1 + StrAppend
                    'sql3 = " SELECT PAYCode, HOSName, SUM(DurationStay) AS NoofDay " & _
                           " From VIEW_CLMReportDCU2 Where (PAYCode = '' OR PAYCode IS NULL) AND " & _
                           " HOSName = '" & Trim(Hospitals) & "'"
                              
                    Else
                        strAppendCase = "57"
                        StrAppend1 = " AND PAYCode = '" & Trim(Payor) & "' AND " & _
                           " HOSName = '" & Trim(HOspitals) & "'"
                        StrAppend1 = StrAppend1 + StrAppend
                    'sql3 = " SELECT PAYCode, HOSName, SUM(DurationStay) AS NoofDay " & _
                           " From VIEW_CLMReportDCU2 Where PAYCode = '" & Trim(Payor) & "' AND " & _
                           " HOSName = '" & Trim(Hospitals) & "'"
                              
                    End If
                       
                       'sql2 = " Group By PAYCode, HOSName"
                       
                       'If Len(Trim(StrAppend)) Then
                       '     sql3 = sql3 + StrAppend + sql2
                       ' Else
                       '     sql3 = sql3 + sql2
                       ' End If
                                              
                       'rst1.Open sql3, medixmasterconnWS, adOpenForwardOnly, adLockBatchOptimistic
                        Call DurationStay(StrAppend1, strAppendCase, Record)
                End If
                
                
                If Len(rst2!TotalActual) Then
                   Record.SubItems(4) = Format(rst2!TotalActual, "#,##0.00")
                End If
                
                If Len(rst2!TotalApproved) Then
                    Record.SubItems(5) = Format(rst2!TotalApproved, "#,##0.00")
                End If
                            
                If rst2!Total <> 0 Then
                    Record.SubItems(6) = Format(Round(rst2!TotalActual / rst2!Total, 0), "#,##0.00")
                    Record.SubItems(7) = Format(Round(rst2!TotalApproved / rst2!Total, 0), "#,##0.00")
                Else
                    Record.SubItems(6) = Format(0, "#,##0.00")
                    Record.SubItems(7) = Format(0, "#,##0.00")
                End If
                
            End With
            lblCurrent = lblCurrent + 1
            rst2.MoveNext
        Loop
    If ListView2.ListItems.Count = 0 Then
        MsgBox "No record found!", vbOKOnly + vbCritical, DialogBoxTitle
        CmdPrint.Enabled = True
        CmdClear.Enabled = True
        cmdExit.Enabled = True
    End If
    rst2.Close
    Set rst2 = Nothing
    CmdPrint.Enabled = True
    CmdClear.Enabled = True
    cmdExit.Enabled = True
End Sub

Private Sub HOSTop202(Optional StrAppend As String)
Dim rst2 As New ADODB.Recordset
Dim cmd As New ADODB.Command
Dim prm1 As ADODB.Parameter
Dim prm2 As ADODB.Parameter
Dim strAppendCase As String
Dim StrAppend1 As String
Dim Record As ListItem

    Hospital = ""
    HOspitals = ""
    lblCurrent = 0
    ListView2.ListItems.Clear
        
    If cboClaimCase.ListIndex = 0 Then
        strAppendCase = "57"
    'sql = " SELECT TOP 20 HOSName, COUNT(HOSName) AS Total, SUM(NoofDay) AS NoofDay, " & _
          " SUM(CLMTotalActual) AS TotalActual, SUM(CLMTotalApproved) AS TotalApproved" & _
          " From VIEW_CLMReportDCU Where 0 = 0"
          
    ElseIf cboClaimCase.ListIndex = 1 Then
        strAppendCase = "58"
    'sql = " SELECT TOP 20 HOSName, COUNT(DISTINCT CASNumber) AS Total, " & _
          " SUM(CLMTotalActual) AS TotalActual, SUM(CLMTotalApproved) AS TotalApproved" & _
          " From VIEW_CLMReportDCU Where 0 = 0 "
          
    End If
    
    'sql2 = " GROUP BY HOSName ORDER BY TotalActual DESC "
        
    'sql = sql + StrAppend + sql2
   
    'rst2.Open sql, medixmasterconnWS, adOpenForwardOnly, adLockBatchOptimistic
    Set cmd.ActiveConnection = medixmasterconnWS
        cmd.CommandText = "DCUReport"
        cmd.CommandType = adCmdStoredProc
        cmd.CommandTimeout = 300
        Set prm1 = cmd.CreateParameter("StrAppend", adVarChar, adParamInput, 2000, StrAppend)
        cmd.Parameters.Append prm1
        Set prm2 = cmd.CreateParameter("StrAppendCase", adVarChar, adParamInput, 2000, strAppendCase)
        cmd.Parameters.Append prm2
        rst2.CursorLocation = adUseClient
        rst2.CursorType = adOpenForwardOnly
        rst2.CacheSize = 100
        Set rst2 = cmd.Execute
        
        Do Until rst2.EOF
    
            With rst2
                
                If Len(rst2!HOSName) Then
                    Set Record = ListView2.ListItems.Add(, , rst2!HOSName)
                    Hospital = Trim(rst2!HOSName)
                Else
                    Set Record = ListView2.ListItems.Add(, , "Empty")
                    Hospital = ""
                End If
                
                If Len(rst2!Total) Then
                    Record.SubItems(1) = rst2!Total
                End If
                
                If cboClaimCase.ListIndex = 0 Then
                
                    If rst2!NoofDay = "0" Then
                        Record.SubItems(2) = "1"
                    ElseIf Len(rst2!NoofDay) Then
                        Record.SubItems(2) = rst2!NoofDay
                    End If
                    
                ElseIf cboClaimCase.ListIndex = 1 Then
                
                    HOspitals = Replace(Hospital, "'", "''")
                    
                    If HOspitals = "" Then
                        strAppendCase = "59"
                        StrAppend1 = " AND (HOSName = '' OR HOSName IS NULL)"
                        StrAppend1 = StrAppend1 + StrAppend
                    'sql3 = " SELECT HOSName, SUM(DurationStay) AS NoofDay " & _
                           " From VIEW_CLMReportDCU2 Where (HOSName = '' OR HOSName IS NULL)"
                    
                    Else
                        strAppendCase = "59"
                        StrAppend1 = " AND HOSName = '" & Trim(HOspitals) & "'"
                        StrAppend1 = StrAppend1 + StrAppend
                    'sql3 = " SELECT HOSName, SUM(DurationStay) AS NoofDay " & _
                           " From VIEW_CLMReportDCU2 Where HOSName = '" & Trim(Hospitals) & "'"
                              
                    End If
                       
                       'sql2 = " Group By HOSName"
                       
                       'If Len(Trim(StrAppend)) Then
                       '     sql3 = sql3 + StrAppend + sql2
                       ' Else
                       '     sql3 = sql3 + sql2
                       ' End If
                       
                       'rst1.Open sql3, medixmasterconnWS, adOpenForwardOnly, adLockBatchOptimistic
                        Call DurationStay(StrAppend1, strAppendCase, Record)
                End If
                
                
                If Len(rst2!TotalActual) Then
                   Record.SubItems(3) = Format(rst2!TotalActual, "#,##0.00")
                End If
                
                If Len(rst2!TotalApproved) Then
                    Record.SubItems(4) = Format(rst2!TotalApproved, "#,##0.00")
                End If
                            
                If rst2!Total <> 0 Then
                    Record.SubItems(5) = Format(Round(rst2!TotalActual / rst2!Total, 0), "#,##0.00")
                    Record.SubItems(6) = Format(Round(rst2!TotalApproved / rst2!Total, 0), "#,##0.00")
                Else
                    Record.SubItems(5) = Format(0, "#,##0.00")
                    Record.SubItems(6) = Format(0, "#,##0.00")
                End If
                
            End With
            lblCurrent = lblCurrent + 1
            rst2.MoveNext
        Loop
    If ListView2.ListItems.Count = 0 Then
        MsgBox "No record found!", vbOKOnly + vbCritical, DialogBoxTitle
        CmdPrint.Enabled = True
        CmdClear.Enabled = True
        cmdExit.Enabled = True
    End If
    rst2.Close
    Set rst2 = Nothing
    CmdPrint.Enabled = True
    CmdClear.Enabled = True
    cmdExit.Enabled = True
End Sub

Private Sub HOSSumListing(Optional StrAppend As String)
Dim rst2 As New ADODB.Recordset
Dim cmd As New ADODB.Command
Dim prm1 As ADODB.Parameter
Dim prm2 As ADODB.Parameter
Dim strAppendCase As String
Dim StrAppend1 As String
Dim Record As ListItem

    Payor = ""
    Hospital = ""
    HOspitals = ""
    lblCurrent = 0
    ListView2.ListItems.Clear
        
    If cboClaimCase.ListIndex = 0 Then
        strAppendCase = "60"
    'sql = " SELECT PAYCode, HOSName, COUNT(HOSName) AS Total, SUM(NoofDay) AS NoofDay, " & _
          " SUM(CLMTotalActual) AS TotalActual, SUM(CLMTotalApproved) AS TotalApproved " & _
          " From VIEW_CLMReportDCU Where 0 = 0 "
    
    ElseIf cboClaimCase.ListIndex = 1 Then
        strAppendCase = "61"
    'sql = " SELECT PAYCode, HOSName, COUNT(DISTINCT CASNumber) AS Total, " & _
          " SUM(CLMTotalActual) AS TotalActual, SUM(CLMTotalApproved) AS TotalApproved " & _
          " From VIEW_CLMReportDCU Where 0 = 0 "
    
    End If
    
    'sql2 = " GROUP BY PAYCode, HOSName "
    
    'If Len(Trim(StrAppend)) Then
    '    sql = sql + StrAppend + sql2
    'Else
    '    sql = sql + sql2
    'End If
        
    'rst2.Open sql, medixmasterconnWS, adOpenForwardOnly, adLockBatchOptimistic
    Set cmd.ActiveConnection = medixmasterconnWS
        cmd.CommandText = "DCUReport"
        cmd.CommandType = adCmdStoredProc
        cmd.CommandTimeout = 300
        Set prm1 = cmd.CreateParameter("StrAppend", adVarChar, adParamInput, 2000, StrAppend)
        cmd.Parameters.Append prm1
        Set prm2 = cmd.CreateParameter("StrAppendCase", adVarChar, adParamInput, 2000, strAppendCase)
        cmd.Parameters.Append prm2
        rst2.CursorLocation = adUseClient
        rst2.CursorType = adOpenForwardOnly
        rst2.CacheSize = 100
        Set rst2 = cmd.Execute

        Do Until rst2.EOF
                
            With rst2
                If Len(rst2!PAYCode) Then
                    Set Record = ListView2.ListItems.Add(, , rst2!PAYCode)
                    Payor = Trim(rst2!PAYCode)
                Else
                    Set Record = ListView2.ListItems.Add(, , "Empty")
                    Payor = ""
                End If
                
                If Len(rst2!HOSName) Then
                    Record.SubItems(1) = Trim(rst2!HOSName)
                    Hospital = Trim(rst2!HOSName)
                Else
                    Record.SubItems(1) = "Empty"
                    Hospital = ""
                End If
                                
                If Len(rst2!Total) Then
                    Record.SubItems(2) = rst2!Total
                End If
                                
                If cboClaimCase.ListIndex = 0 Then
                
                    If rst2!NoofDay = "0" Then
                        Record.SubItems(3) = "1"
                    ElseIf Len(rst2!NoofDay) Then
                        Record.SubItems(3) = rst2!NoofDay
                    End If
                    
                ElseIf cboClaimCase.ListIndex = 1 Then
                
                    HOspitals = Replace(Hospital, "'", "''")
                    
                    If Payor = "" And HOspitals = "" Then
                        strAppendCase = "62"
                        StrAppend1 = " AND (PAYCode = '' OR PAYCode IS NULL) AND " & _
                           " (HOSName = '' OR HOSName IS NULL)"
                        StrAppend1 = StrAppend1 + StrAppend
                    'sql3 = " SELECT PAYCode, HOSName, SUM(DurationStay) AS NoofDay " & _
                           " From VIEW_CLMReportDCU2 Where (PAYCode = '' OR PAYCode IS NULL) AND " & _
                           " (HOSName = '' OR HOSName IS NULL)"
                              
                    ElseIf Payor <> "" And HOspitals = "" Then
                        strAppendCase = "62"
                        StrAppend1 = " AND PAYCode = '" & Trim(Payor) & "' AND " & _
                           " (HOSName = '' OR HOSName IS NULL)"
                        StrAppend1 = StrAppend1 + StrAppend
                    'sql3 = " SELECT PAYCode, HOSName, SUM(DurationStay) AS NoofDay " & _
                           " From VIEW_CLMReportDCU2 Where PAYCode = '" & Trim(Payor) & "' AND " & _
                           " (HOSName = '' OR HOSName IS NULL)"
                              
                    ElseIf Payor = "" And HOspitals <> "" Then
                        strAppendCase = "62"
                        StrAppend1 = " AND (PAYCode = '' OR PAYCode IS NULL) AND " & _
                           " HOSName = '" & Trim(HOspitals) & "'"
                        StrAppend1 = StrAppend1 + StrAppend
                    'sql3 = " SELECT PAYCode, HOSName, SUM(DurationStay) AS NoofDay " & _
                           " From VIEW_CLMReportDCU2 Where (PAYCode = '' OR PAYCode IS NULL) AND " & _
                           " HOSName = '" & Trim(Hospitals) & "'"
                              
                    Else
                        strAppendCase = "62"
                        StrAppend1 = " AND PAYCode = '" & Trim(Payor) & "' AND " & _
                           " HOSName = '" & Trim(HOspitals) & "'"
                        StrAppend1 = StrAppend1 + StrAppend
                    'sql3 = " SELECT PAYCode, HOSName, SUM(DurationStay) AS NoofDay " & _
                           " From VIEW_CLMReportDCU2 Where PAYCode = '" & Trim(Payor) & "' AND " & _
                           " HOSName = '" & Trim(Hospitals) & "'"
                              
                    End If
                       
                       'sql2 = " Group By PAYCode, HOSName"
                       
                       'If Len(Trim(StrAppend)) Then
                       '     sql3 = sql3 + StrAppend + sql2
                       ' Else
                       '     sql3 = sql3 + sql2
                       ' End If
                                              
                       'rst1.Open sql3, medixmasterconnWS, adOpenForwardOnly, adLockBatchOptimistic
                        Call DurationStay(StrAppend1, strAppendCase, Record)
                End If
                              
                
                If Len(rst2!TotalActual) Then
                   Record.SubItems(4) = Format(rst2!TotalActual, "#,##0.00")
                End If
                
                If Len(rst2!TotalApproved) Then
                    Record.SubItems(5) = Format(rst2!TotalApproved, "#,##0.00")
                End If
                
                If rst2!Total <> 0 Then
                    Record.SubItems(6) = Format(Round(rst2!TotalActual / rst2!Total, 0), "#,##0.00")
                    Record.SubItems(7) = Format(Round(rst2!TotalApproved / rst2!Total, 0), "#,##0.00")
                Else
                    Record.SubItems(6) = Format(0, "#,##0.00")
                    Record.SubItems(7) = Format(0, "#,##0.00")
                End If
                
                           
            End With
            lblCurrent = lblCurrent + 1
            rst2.MoveNext
        Loop
    If ListView2.ListItems.Count = 0 Then
        MsgBox "No record found!", vbOKOnly + vbCritical, DialogBoxTitle
        CmdPrint.Enabled = True
        CmdClear.Enabled = True
        cmdExit.Enabled = True
    End If
    rst2.Close
    Set rst2 = Nothing
    CmdPrint.Enabled = True
    CmdClear.Enabled = True
    cmdExit.Enabled = True
End Sub

Private Sub HOSSumListing2(Optional StrAppend As String)
Dim rst2 As New ADODB.Recordset
Dim cmd As New ADODB.Command
Dim prm1 As ADODB.Parameter
Dim prm2 As ADODB.Parameter
Dim strAppendCase As String
Dim StrAppend1 As String
Dim Record As ListItem
    
    Hospital = ""
    HOspitals = ""
    lblCurrent = 0
    ListView2.ListItems.Clear
        
    If cboClaimCase.ListIndex = 0 Then
        strAppendCase = "63"
    'sql = " SELECT HOSName, COUNT(HOSName) AS Total, SUM(NoofDay) AS NoofDay, " & _
          " SUM(CLMTotalActual) AS TotalActual, SUM(CLMTotalApproved) AS TotalApproved " & _
          " From VIEW_CLMReportDCU Where 0 = 0 "
          
    ElseIf cboClaimCase.ListIndex = 1 Then
        strAppendCase = "64"
    'sql = " SELECT HOSName, COUNT(DISTINCT CASNumber) AS Total, " & _
          " SUM(CLMTotalActual) AS TotalActual, SUM(CLMTotalApproved) AS TotalApproved " & _
          " From VIEW_CLMReportDCU Where 0 = 0 "
    
    End If
    
    'sql2 = " GROUP BY HOSName "
    
    'If Len(Trim(StrAppend)) Then
    '    sql = sql + StrAppend + sql2
    'Else
    '    sql = sql + sql2
    'End If
        
    'rst2.Open sql, medixmasterconnWS, adOpenForwardOnly, adLockBatchOptimistic
    Set cmd.ActiveConnection = medixmasterconnWS
        cmd.CommandText = "DCUReport"
        cmd.CommandType = adCmdStoredProc
        cmd.CommandTimeout = 300
        Set prm1 = cmd.CreateParameter("StrAppend", adVarChar, adParamInput, 2000, StrAppend)
        cmd.Parameters.Append prm1
        Set prm2 = cmd.CreateParameter("StrAppendCase", adVarChar, adParamInput, 2000, strAppendCase)
        cmd.Parameters.Append prm2
        rst2.CursorLocation = adUseClient
        rst2.CursorType = adOpenForwardOnly
        rst2.CacheSize = 100
        Set rst2 = cmd.Execute
        
        Do Until rst2.EOF
                
            With rst2
                If Len(rst2!HOSName) Then
                    Set Record = ListView2.ListItems.Add(, , Trim(rst2!HOSName))
                    Hospital = Trim(rst2!HOSName)
                Else
                    Set Record = ListView2.ListItems.Add(, , "Empty")
                    Hospital = ""
                End If
                               
                                
                If Len(rst2!Total) Then
                    Record.SubItems(1) = rst2!Total
                End If
                                
                If cboClaimCase.ListIndex = 0 Then
                
                    If rst2!NoofDay = "0" Then
                        Record.SubItems(2) = "1"
                    ElseIf Len(rst2!NoofDay) Then
                        Record.SubItems(2) = rst2!NoofDay
                    End If
                    
                ElseIf cboClaimCase.ListIndex = 1 Then
                
                    HOspitals = Replace(Hospital, "'", "''")
                    
                    If HOspitals = "" Then
                        strAppendCase = "65"
                        StrAppend1 = " AND (HOSName = '' OR HOSName IS NULL)"
                        StrAppend1 = StrAppend1 + StrAppend
                    'sql3 = " SELECT HOSName, SUM(DurationStay) AS NoofDay " & _
                           " From VIEW_CLMReportDCU2 Where (HOSName = '' OR HOSName IS NULL)"
                              
                    Else
                        strAppendCase = "65"
                        StrAppend1 = " AND HOSName = '" & Trim(HOspitals) & "'"
                        StrAppend1 = StrAppend1 + StrAppend
                    'sql3 = " SELECT HOSName, SUM(DurationStay) AS NoofDay " & _
                           " From VIEW_CLMReportDCU2 Where HOSName = '" & Trim(Hospitals) & "'"
                              
                    End If
                       
                       'sql2 = " Group By HOSName "
                       
                       'If Len(Trim(StrAppend)) Then
                       '     sql3 = sql3 + StrAppend + sql2
                       ' Else
                       '     sql3 = sql3 + sql2
                       ' End If
                                              
                       'rst1.Open sql3, medixmasterconnWS, adOpenForwardOnly, adLockBatchOptimistic
                        Call DurationStay(StrAppend1, strAppendCase, Record)
                End If
                
                If Len(rst2!TotalActual) Then
                   Record.SubItems(3) = Format(rst2!TotalActual, "#,##0.00")
                End If
                
                If Len(rst2!TotalApproved) Then
                    Record.SubItems(4) = Format(rst2!TotalApproved, "#,##0.00")
                End If
                                
                If rst2!Total <> 0 Then
                    Record.SubItems(5) = Format(Round(rst2!TotalActual / rst2!Total, 0), "#,##0.00")
                    Record.SubItems(6) = Format(Round(rst2!TotalApproved / rst2!Total, 0), "#,##0.00")
                Else
                    Record.SubItems(5) = Format(0, "#,##0.00")
                    Record.SubItems(6) = Format(0, "#,##0.00")

                End If
                           
            End With
            lblCurrent = lblCurrent + 1
            rst2.MoveNext
        Loop
    If ListView2.ListItems.Count = 0 Then
        MsgBox "No record found!", vbOKOnly + vbCritical, DialogBoxTitle
        CmdPrint.Enabled = True
        CmdClear.Enabled = True
        cmdExit.Enabled = True
    End If
    rst2.Close
    Set rst2 = Nothing
    CmdPrint.Enabled = True
    CmdClear.Enabled = True
    cmdExit.Enabled = True
End Sub

Private Sub HOSFinalICD(StrAppend1, strAppendCase, Record)
Dim rst3 As New ADODB.Recordset
Dim con As New ADODB.Command
Dim prm1 As ADODB.Parameter
Dim prm2 As ADODB.Parameter

Set con.ActiveConnection = medixmasterconn
    con.CommandText = "DCUReport"
    con.CommandType = adCmdStoredProc
    con.CommandTimeout = 300
    Set prm1 = con.CreateParameter("StrAppend", adVarChar, adParamInput, 2000, StrAppend1)
    con.Parameters.Append prm1
    Set prm2 = con.CreateParameter("StrAppendCase", adVarChar, adParamInput, 2000, strAppendCase)
    con.Parameters.Append prm2
    rst3.CursorLocation = adUseClient
    rst3.CursorType = adOpenForwardOnly
    rst3.CacheSize = 100
    Set rst3 = con.Execute
    
    If rst3.EOF = True Then
        Record.SubItems(2) = ""
        rst3.Close
        Set rst3 = Nothing
    Else
        Record.SubItems(2) = Trim(rst3!ICDDescription)
        rst3.Close
        Set rst3 = Nothing
    End If
End Sub

Private Sub HOSByDiag(Optional StrAppend As String)
Dim rst2 As New ADODB.Recordset
Dim cmd As New ADODB.Command
Dim prm1 As ADODB.Parameter
Dim prm2 As ADODB.Parameter
Dim strAppendCase As String
Dim StrAppend1 As String
Dim Record As ListItem
    
    DiagCode = ""
    Hospital = ""
    HOspitals = ""
    lblCurrent = 0
    ListView2.ListItems.Clear
    
    If cboClaimCase.ListIndex = 0 Then
        strAppendCase = "68"
    'sql = " SELECT COUNT(CASNumber) AS Total, FinalDiag, Sum(CLMTotalApproved) As TotalApproved, " & _
          " Sum(CLMTotalActual) As TotalActual, HOSName, " & _
          " SUM(NoofDay) AS NoofDay From VIEW_CLMReportDCU Where 0 = 0 " & StrAppend & _
          " GROUP BY FinalDiag, HOSName "
          
    ElseIf cboClaimCase.ListIndex = 1 Then
        strAppendCase = "69"
    'sql = " SELECT COUNT(DISTINCT CASNumber) AS Total, FinalDiag, Sum(CLMTotalApproved) As TotalApproved, " & _
          " Sum(CLMTotalActual) As TotalActual, HOSName " & _
          " From VIEW_CLMReportDCU Where 0 = 0 " & StrAppend & _
          " GROUP BY FinalDiag, HOSName "
          
    End If
    
    'rst2.Open sql, medixmasterconnWS, adOpenForwardOnly, adLockBatchOptimistic
    Set cmd.ActiveConnection = medixmasterconnWS
        cmd.CommandText = "DCUReport"
        cmd.CommandType = adCmdStoredProc
        cmd.CommandTimeout = 300
        Set prm1 = cmd.CreateParameter("StrAppend", adVarChar, adParamInput, 2000, StrAppend)
        cmd.Parameters.Append prm1
        Set prm2 = cmd.CreateParameter("StrAppendCase", adVarChar, adParamInput, 2000, strAppendCase)
        cmd.Parameters.Append prm2
        rst2.CursorLocation = adUseClient
        rst2.CursorType = adOpenForwardOnly
        rst2.CacheSize = 100
        Set rst2 = cmd.Execute
        
        Do Until rst2.EOF
    
            With rst2
                               
                If Len(rst2!HOSName) Then
                    Set Record = ListView2.ListItems.Add(, , Trim(rst2!HOSName))
                    Hospital = rst2!HOSName
                Else
                    Set Record = ListView2.ListItems.Add(, , "Empty")
                    Hospital = ""
                End If
                If Len(rst2!FinalDiag) Then
                    Record.SubItems(1) = Trim(rst2!FinalDiag)
                    DiagCode = rst2!FinalDiag
                Else
                    Record.SubItems(1) = "Empty"
                    DiagCode = ""
                End If
                
                'sql2 = " SELECT ICDCode, ICDDescription From ICDOriginal Where ICDCode = '" & Trim(DiagCode) & "'"
                'rst3.Open sql2, medixmasterconn, adOpenForwardOnly, adLockBatchOptimistic
                strAppendCase = "1"
                StrAppend1 = " AND ICDCode = '" & Trim(DiagCode) & "'"
                Call HOSFinalICD(StrAppend1, strAppendCase, Record)
                                   
                If Len(rst2!Total) Then
                    Record.SubItems(3) = rst2!Total
                End If
                
                If cboClaimCase.ListIndex = 0 Then
                
                    If rst2!NoofDay = "0" Then
                        Record.SubItems(4) = "1"
                    ElseIf Len(rst2!NoofDay) Then
                        Record.SubItems(4) = rst2!NoofDay
                    End If
                
                ElseIf cboClaimCase.ListIndex = 1 Then
                
                    HOspitals = Replace(Hospital, "'", "''")
                    
                    If DiagCode = "" And HOspitals = "" Then
                        strAppendCase = "70"
                        StrAppend1 = " AND (FinalDiag = '' OR FinalDiag IS NULL) AND " & _
                           " (HOSName = '' OR HOSName IS NULL)"
                        StrAppend1 = StrAppend1 + StrAppend
                    'sql3 = " SELECT FinalDiag, HOSName, SUM(DurationStay) AS NoofDay " & _
                           " From VIEW_CLMReportDCU2 Where (FinalDiag = '' OR FinalDiag IS NULL) AND " & _
                           " (HOSName = '' OR HOSName IS NULL)"
                              
                    ElseIf DiagCode <> "" And HOspitals = "" Then
                        strAppendCase = "70"
                        StrAppend1 = " AND FinalDiag = '" & Trim(DiagCode) & "' AND " & _
                           " (HOSName = '' OR HOSName IS NULL)"
                        StrAppend1 = StrAppend1 + StrAppend
                    'sql3 = " SELECT FinalDiag, HOSName, SUM(DurationStay) AS NoofDay " & _
                           " From VIEW_CLMReportDCU2 Where FinalDiag = '" & Trim(DiagCode) & "' AND " & _
                           " (HOSName = '' OR HOSName IS NULL)"
                              
                    ElseIf DiagCode = "" And HOspitals <> "" Then
                        strAppendCase = "70"
                        StrAppend1 = " AND (FinalDiag = '' OR FinalDiag IS NULL) AND " & _
                           " HOSName = '" & Trim(HOspitals) & "'"
                        StrAppend1 = StrAppend1 + StrAppend
                    'sql3 = " SELECT FinalDiag, HOSName, SUM(DurationStay) AS NoofDay " & _
                           " From VIEW_CLMReportDCU2 Where (FinalDiag = '' OR FinalDiag IS NULL) AND " & _
                           " HOSName = '" & Trim(Hospitals) & "'"
                              
                    Else
                        strAppendCase = "70"
                        StrAppend1 = " AND FinalDiag = '" & Trim(DiagCode) & "' AND " & _
                           " HOSName = '" & Trim(HOspitals) & "'"
                        StrAppend1 = StrAppend1 + StrAppend
                    'sql3 = " SELECT FinalDiag, HOSName, SUM(DurationStay) AS NoofDay " & _
                           " From VIEW_CLMReportDCU2 Where FinalDiag = '" & Trim(DiagCode) & "' AND " & _
                           " HOSName = '" & Trim(Hospitals) & "'"
                              
                    End If
                       
                       'sql2 = " Group By FinalDiag, HOSName"
                       
                       'If Len(Trim(StrAppend)) Then
                       '     sql3 = sql3 + StrAppend + sql2
                       ' Else
                       '     sql3 = sql3 + sql2
                       ' End If
                                              
                       'rst1.Open sql3, medixmasterconnWS, adOpenForwardOnly, adLockBatchOptimistic
                        Call DurationStay(StrAppend1, strAppendCase, Record)
                End If
                
                If Len(rst2!TotalActual) Then
                   Record.SubItems(5) = Format(rst2!TotalActual, "##,#0.00")
                End If
                
                If Len(rst2!TotalApproved) Then
                    Record.SubItems(6) = Format(rst2!TotalApproved, "##,#0.00")
                End If
            
            End With
            lblCurrent = lblCurrent + 1
            rst2.MoveNext
        Loop
    If ListView2.ListItems.Count = 0 Then
        MsgBox "No record found!", vbOKOnly + vbCritical, DialogBoxTitle
        CmdPrint.Enabled = True
        CmdClear.Enabled = True
        cmdExit.Enabled = True
    End If
    rst2.Close
    Set rst2 = Nothing
    CmdPrint.Enabled = True
    CmdClear.Enabled = True
    cmdExit.Enabled = True
End Sub

Private Sub HOSDetailListing(Optional StrAppend As String)
Dim rst2 As New ADODB.Recordset
Dim cmd As New ADODB.Command
Dim prm1 As ADODB.Parameter
Dim prm2 As ADODB.Parameter
Dim strAppendCase As String
Dim Record As ListItem
    
    ListView2.ListItems.Clear
    Current = 0
    Total = 0
            
    If cboClaimCase.ListIndex = 0 Then
        strAppendCase = "66"
    'sql = " SELECT PAYCode, HOSName, CASNumber, ClaimVersion, MBMNumber, MBMPatientCovID, " & _
          " PatientName, CLMTotalActual, CLMTotalApproved, " & _
          " RecoveryAmt, CASDateAdmitted, CASDischargeDate, " & _
          " CASStatus, CASClaimability, FinalDiag From VIEW_CLMReportDCU Where 0 = 0 "
    
    'If Len(Trim(StrAppend)) Then
    '    sql = sql + StrAppend
    'End If
            
    'rst2.Open sql, medixmasterconnWS, adOpenForwardOnly, adLockBatchOptimistic
    Set cmd.ActiveConnection = medixmasterconnWS
        cmd.CommandText = "DCUReport"
        cmd.CommandType = adCmdStoredProc
        cmd.CommandTimeout = 300
        Set prm1 = cmd.CreateParameter("StrAppend", adVarChar, adParamInput, 2000, StrAppend)
        cmd.Parameters.Append prm1
        Set prm2 = cmd.CreateParameter("StrAppendCase", adVarChar, adParamInput, 2000, strAppendCase)
        cmd.Parameters.Append prm2
        rst2.CursorLocation = adUseClient
        rst2.CursorType = adOpenForwardOnly
        rst2.CacheSize = 100
        Set rst2 = cmd.Execute
        
    Do
    
            Do Until rst2.EOF
            
            Total = rst2.recordCount
            lblTotalRecord = Total
                
            With rst2
                
                If Len(rst2!PAYCode) Then
                    Set Record = ListView2.ListItems.Add(, , rst2!PAYCode)
                Else
                    Set Record = ListView2.ListItems.Add(, , "Empty")
                End If
                
                If Len(rst2!HOSName) Then
                    Record.SubItems(1) = Trim(rst2!HOSName)
                Else
                    Record.SubItems(1) = "Empty"
                End If
                
                If Len(rst2!CasNumber & "-" & rst2!ClaimVersion) Then
                   Record.SubItems(2) = rst2!CasNumber & "-" & rst2!ClaimVersion
                End If
                                
                If Len(rst2!MBMNumber) Then
                    Record.SubItems(3) = rst2!MBMNumber
                End If
                
                If Len(rst2!MBMPatientCovID) Then
                    Record.SubItems(4) = rst2!MBMPatientCovID
                End If
                
                If Len(rst2!PatientName) Then
                    Record.SubItems(5) = Trim(rst2!PatientName)
                End If
                
                If Len(rst2!CLMTotalActual) Then
                    Record.SubItems(6) = Format(rst2!CLMTotalActual, "#,##0.00")
                End If
                
                If Len(rst2!CLMTotalApproved) Then
                    Record.SubItems(7) = Format(rst2!CLMTotalApproved, "#,##0.00")
                End If
                
                If Len(rst2!RecoveryAmt) Then
                    Record.SubItems(8) = Format(rst2!RecoveryAmt, "#,##0.00")
                End If
                
                If Len(rst2!CASDateAdmitted) Then
                    Record.SubItems(9) = Format(rst2!CASDateAdmitted, "dd/mm/yyyy")
                End If
                
                If Len(rst2!CASDischargeDate) Then
                    Record.SubItems(10) = Format(rst2!CASDischargeDate, "dd/mm/yyyy")
                End If
                
                If Len(rst2!CASStatus) Then
                   Record.SubItems(11) = rst2!CASStatus
                End If
                
                If Len(rst2!CASClaimability) Then
                    Record.SubItems(12) = rst2!CASClaimability
                End If
                                
                If Len(rst2!FinalDiag) Then
                    Record.SubItems(13) = rst2!FinalDiag
                End If
                
                Current = Current + 1
                lblCurrent = Current
                           
            End With
            rst2.MoveNext
            DoEvents
        
            If intExit = 1 Then
                Check = False
                Exit Do
            End If
            Loop
        Loop Until Check = False
    
    intExit = 0
    If ListView2.ListItems.Count = 0 Then
        MsgBox "No record found!", vbOKOnly + vbCritical, DialogBoxTitle
        CmdPrint.Enabled = True
        CmdClear.Enabled = True
        cmdExit.Enabled = True
    End If
    rst2.Close
    Set rst2 = Nothing
    CmdPrint.Enabled = True
    CmdClear.Enabled = True
    cmdExit.Enabled = True
    
    ElseIf cboClaimCase.ListIndex = 1 Then
        strAppendCase = "67"
   'sql = " SELECT PAYCode, HOSName, CASNumber, MBMNumber, MBMPatientCovID, " & _
          " PatientName, SUM(CLMTotalActual) AS CLMTotalActual, SUM(CLMTotalApproved) AS CLMTotalApproved, " & _
          " SUM(RecoveryAmt) AS RecoveryAmt, CASDateAdmitted, CASDischargeDate, " & _
          " CASStatus, CASClaimability, FinalDiag From VIEW_CLMReportDCU Where 0 = 0 "
    
    'sql2 = " Group By PAYCode, HOSName, CASNumber, MBMNumber, MBMPatientCovID, " & _
          " PatientName, CASDateAdmitted, CASDischargeDate, " & _
          " CASStatus, CASClaimability, FinalDiag "
    
    'If Len(Trim(StrAppend)) Then
    '    sql = sql + StrAppend + sql2
    'End If
            
    'rst2.Open sql, medixmasterconnWS, adOpenForwardOnly, adLockBatchOptimistic
    Set cmd.ActiveConnection = medixmasterconnWS
        cmd.CommandText = "DCUReport"
        cmd.CommandType = adCmdStoredProc
        cmd.CommandTimeout = 300
        Set prm1 = cmd.CreateParameter("StrAppend", adVarChar, adParamInput, 2000, StrAppend)
        cmd.Parameters.Append prm1
        Set prm2 = cmd.CreateParameter("StrAppendCase", adVarChar, adParamInput, 2000, strAppendCase)
        cmd.Parameters.Append prm2
        rst2.CursorLocation = adUseClient
        rst2.CursorType = adOpenForwardOnly
        rst2.CacheSize = 100
        Set rst2 = cmd.Execute
        
    Do
    
            Do Until rst2.EOF
            
            Total = rst2.recordCount
            lblTotalRecord = Total
                
            With rst2
                
                If Len(rst2!PAYCode) Then
                    Set Record = ListView2.ListItems.Add(, , rst2!PAYCode)
                Else
                    Set Record = ListView2.ListItems.Add(, , "Empty")
                End If
                
                If Len(rst2!HOSName) Then
                    Record.SubItems(1) = Trim(rst2!HOSName)
                Else
                    Record.SubItems(1) = "Empty"
                End If
                
                If Len(rst2!CasNumber) Then
                   Record.SubItems(2) = rst2!CasNumber
                End If
                                
                If Len(rst2!MBMNumber) Then
                    Record.SubItems(3) = rst2!MBMNumber
                End If
                
                If Len(rst2!MBMPatientCovID) Then
                    Record.SubItems(4) = rst2!MBMPatientCovID
                End If
                
                If Len(rst2!PatientName) Then
                    Record.SubItems(5) = Trim(rst2!PatientName)
                End If
                
                If Len(rst2!CLMTotalActual) Then
                    Record.SubItems(6) = Format(rst2!CLMTotalActual, "#,##0.00")
                End If
                
                If Len(rst2!CLMTotalApproved) Then
                    Record.SubItems(7) = Format(rst2!CLMTotalApproved, "#,##0.00")
                End If
                
                If Len(rst2!RecoveryAmt) Then
                    Record.SubItems(8) = Format(rst2!RecoveryAmt, "#,##0.00")
                End If
                
                If Len(rst2!CASDateAdmitted) Then
                    Record.SubItems(9) = Format(rst2!CASDateAdmitted, "dd/mm/yyyy")
                End If
                
                If Len(rst2!CASDischargeDate) Then
                    Record.SubItems(10) = Format(rst2!CASDischargeDate, "dd/mm/yyyy")
                End If
                
                If Len(rst2!CASStatus) Then
                   Record.SubItems(11) = rst2!CASStatus
                End If
                
                If Len(rst2!CASClaimability) Then
                    Record.SubItems(12) = rst2!CASClaimability
                End If
                                
                If Len(rst2!FinalDiag) Then
                    Record.SubItems(13) = rst2!FinalDiag
                End If
                
                Current = Current + 1
                lblCurrent = Current
                           
            End With
            rst2.MoveNext
            DoEvents
        
            If intExit = 1 Then
                Check = False
                Exit Do
            End If
            Loop
        Loop Until Check = False
    
    intExit = 0
    If ListView2.ListItems.Count = 0 Then
        MsgBox "No record found!", vbOKOnly + vbCritical, DialogBoxTitle
        CmdPrint.Enabled = True
        CmdClear.Enabled = True
        cmdExit.Enabled = True
    End If
    rst2.Close
    Set rst2 = Nothing
    CmdPrint.Enabled = True
    CmdClear.Enabled = True
    cmdExit.Enabled = True
    
    End If
    
End Sub

Private Sub GenFirstICD(StrAppend1, strAppendCase, Record, rst2)
Dim rst3 As New ADODB.Recordset
Dim con As New ADODB.Command
Dim prm1 As ADODB.Parameter
Dim prm2 As ADODB.Parameter

Set con.ActiveConnection = medixmasterconn
    con.CommandText = "DCUReport"
    con.CommandType = adCmdStoredProc
    con.CommandTimeout = 300
    Set prm1 = con.CreateParameter("StrAppend", adVarChar, adParamInput, 2000, StrAppend1)
    con.Parameters.Append prm1
    Set prm2 = con.CreateParameter("StrAppendCase", adVarChar, adParamInput, 2000, strAppendCase)
    con.Parameters.Append prm2
    rst3.CursorLocation = adUseClient
    rst3.CursorType = adOpenForwardOnly
    rst3.CacheSize = 100
    Set rst3 = con.Execute
    
    If cboClaimCase.ListIndex = 0 Then
        If rst3.EOF = True Then
            Record.SubItems(21) = ""
            rst3.Close
            Set rst3 = Nothing
        Else
            Record.SubItems(21) = Trim(rst3!ICDDescription)
            rst3.Close
            Set rst3 = Nothing
        End If
    ElseIf cboClaimCase.ListIndex = 1 Then
        If rst3.EOF = True Then
            Record.SubItems(19) = ""
            rst3.Close
            Set rst3 = Nothing
        Else
            Record.SubItems(19) = Trim(rst3!ICDDescription)
            rst3.Close
            Set rst3 = Nothing
        End If
    End If

End Sub

Private Sub GenFinalICD(StrAppend1, strAppendCase, Record, rst2)
Dim rst3 As New ADODB.Recordset
Dim con As New ADODB.Command
Dim prm1 As ADODB.Parameter
Dim prm2 As ADODB.Parameter

Set con.ActiveConnection = medixmasterconn
    con.CommandText = "DCUReport"
    con.CommandType = adCmdStoredProc
    con.CommandTimeout = 300
    Set prm1 = con.CreateParameter("StrAppend", adVarChar, adParamInput, 2000, StrAppend1)
    con.Parameters.Append prm1
    Set prm2 = con.CreateParameter("StrAppendCase", adVarChar, adParamInput, 2000, strAppendCase)
    con.Parameters.Append prm2
    rst3.CursorLocation = adUseClient
    rst3.CursorType = adOpenForwardOnly
    rst3.CacheSize = 100
    Set rst3 = con.Execute
    
    If cboClaimCase.ListIndex = 0 Then
        If rst3.EOF = True Then
            Record.SubItems(23) = ""
            rst3.Close
            Set rst3 = Nothing
        Else
            Record.SubItems(23) = Trim(rst3!ICDDescription)
            rst3.Close
            Set rst3 = Nothing
        End If
    ElseIf cboClaimCase.ListIndex = 1 Then
        If rst3.EOF = True Then
            Record.SubItems(21) = ""
            rst3.Close
            Set rst3 = Nothing
        Else
            Record.SubItems(21) = Trim(rst3!ICDDescription)
            rst3.Close
            Set rst3 = Nothing
        End If
    End If

End Sub

Private Sub GenBordxRefNo2(StrAppend1, strAppendCase, Record, rst2)
Dim rst1 As New ADODB.Recordset
Dim conn As New ADODB.Command
Dim prm1 As ADODB.Parameter
Dim prm2 As ADODB.Parameter
Dim rstCount, CaseNo

Set conn.ActiveConnection = medixmasterconnWS
    conn.CommandText = "DCUReport"
    conn.CommandType = adCmdStoredProc
    conn.CommandTimeout = 300
    Set prm1 = conn.CreateParameter("StrAppend", adVarChar, adParamInput, 2000, StrAppend1)
    conn.Parameters.Append prm1
    Set prm2 = conn.CreateParameter("StrAppendCase", adVarChar, adParamInput, 2000, strAppendCase)
    conn.Parameters.Append prm2
    rst1.CursorLocation = adUseClient
    rst1.CursorType = adOpenForwardOnly
    rst1.CacheSize = 100
    Set rst1 = conn.Execute
    
    CaseNo = rst2!CasNumber

        strCount = " select count(*) as totalRecord From Claim where CASNumber = '" & CaseNo & "' AND (BordxRefno <> '')"
        Set rstCount = medixmasterconnWS.Execute(strCount)

        If rstCount!TotalRecord <> 0 Then
            Record.SubItems(25) = Format(rst1!CLMTotalApproved, "#,##0.00;(#,##0.00)")
            rst1.Close
            Set rst1 = Nothing
        Else
            rst1.Close
            Set rst1 = Nothing
        End If

End Sub

Private Sub GenBordxRefNo1(StrAppend1, strAppendCase, Record, rst2)
Dim rst1 As New ADODB.Recordset
Dim conn As New ADODB.Command
Dim prm1 As ADODB.Parameter
Dim prm2 As ADODB.Parameter
Dim rstCount, CaseNo

Set conn.ActiveConnection = medixmasterconnWS
    conn.CommandText = "DCUReport"
    conn.CommandType = adCmdStoredProc
    conn.CommandTimeout = 300
    Set prm1 = conn.CreateParameter("StrAppend", adVarChar, adParamInput, 2000, StrAppend1)
    conn.Parameters.Append prm1
    Set prm2 = conn.CreateParameter("StrAppendCase", adVarChar, adParamInput, 2000, strAppendCase)
    conn.Parameters.Append prm2
    rst1.CursorLocation = adUseClient
    rst1.CursorType = adOpenForwardOnly
    rst1.CacheSize = 100
    Set rst1 = conn.Execute
    
    CaseNo = rst2!CasNumber

        strCount = " select count(*) as totalRecord From Claim where CASNumber = '" & CaseNo & "' AND (BordxRefno is null or BordxRefno = '')"
         Set rstCount = medixmasterconnWS.Execute(strCount)
                
        If rstCount!TotalRecord <> 0 Then
            Record.SubItems(24) = Format(rst1!CLMTotalApproved, "#,##0.00;(#,##0.00)")
            rst1.Close
            Set rst1 = Nothing
        Else
            rst1.Close
            Set rst1 = Nothing
        End If
End Sub

Private Sub ResultListing(Optional StrAppend As String)
Dim rst2 As New ADODB.Recordset
Dim cmd As New ADODB.Command
Dim prm1 As ADODB.Parameter
Dim prm2 As ADODB.Parameter
Dim strAppendCase As String
Dim StrAppend1 As String
Dim Record As ListItem
Dim FirstICDCode As String
Dim FinalICDCode As String
Dim FirstDescription As String
Dim FinalDescription As String
Dim CaseNo As String
        
    ListView2.ListItems.Clear
    Current = 0
    Total = 0
    
    If cboClaimCase.ListIndex = 0 Then
        
    strAppendCase = "1"
    Set cmd.ActiveConnection = medixmasterconnWS
        cmd.CommandText = "DCUReport"
        cmd.CommandType = adCmdStoredProc
        cmd.CommandTimeout = 300
        Set prm1 = cmd.CreateParameter("StrAppend", adVarChar, adParamInput, 2000, StrAppend)
        cmd.Parameters.Append prm1
        Set prm2 = cmd.CreateParameter("StrAppendCase", adVarChar, adParamInput, 2000, strAppendCase)
        cmd.Parameters.Append prm2
        rst2.CursorLocation = adUseClient
        rst2.CursorType = adOpenForwardOnly
        rst2.CacheSize = 100
        Set rst2 = cmd.Execute
    Do
              
       Do Until rst2.EOF
       
       Total = rst2.recordCount
       lblTotalRecord = Total
        
        FirstICDCode = ""
        FinalICDCode = ""
        FirstDescription = ""
        FinalDescription = ""
        
            With rst2
                    
                    If Len(!PAYCode) Then
                        Set Record = ListView2.ListItems.Add(, , !PAYCode)
                    Else
                        Set Record = ListView2.ListItems.Add(, , "Empty")
                    End If
                    
                    If Len(!CasNumber & "-" & !ClaimVersion) Then
                       Record.SubItems(1) = (!CasNumber & "-" & !ClaimVersion)
                    End If
                    
                    If Len(!BordxRefNo) Then
                       Record.SubItems(2) = !BordxRefNo
                    End If
                    
                    If Len(!BordxDate) Then
                       Record.SubItems(3) = Format(!BordxDate, "dd/mm/yyyy")
                    End If
                    
                    If Len(!MBMNumber) Then
                       Record.SubItems(4) = !MBMNumber
                    End If
                    
                    If Len(!MBMPolicyNo) Then
                       Record.SubItems(5) = Trim(!MBMPolicyNo)
                    End If
                                    
                    If Len(!MBMPatientCovID) Then
                       Record.SubItems(6) = !MBMPatientCovID
                    End If
                    
                    If Len(!MBMName) Then
                       Record.SubItems(7) = Trim(!MBMName)
                    End If
                    
                    If Len(!PatientName) Then
                       Record.SubItems(8) = Trim(!PatientName)
                    End If
                    
                    If Len(!MBMIcBcPp) Then
                       Record.SubItems(9) = !MBMIcBcPp
                    End If
                    
                    If Len(!MBMSex) Then
                       Record.SubItems(10) = !MBMSex
                    End If
                    
                    If Len(!AgeCode) Then
                       Record.SubItems(11) = !AgeCode
                    End If
                    
                    If Len(!MBMDOB) Then
                        Record.SubItems(12) = Format(!MBMDOB, "dd/mm/yyyy")
                    End If
                    
                    If Len(!ICNo) Then
                       Record.SubItems(13) = !ICNo
                    End If
                    
                    If Len(!Sex) Then
                       Record.SubItems(14) = !Sex
                    End If
                    
                    If Len(!AgeBand) Then
                       Record.SubItems(15) = !AgeBand
                    End If
                    
                    If Len(!DOB) Then
                        Record.SubItems(16) = Format(!DOB, "dd/mm/yyyy")
                    End If
                    
                    If Len(!INSCode) Then
                       Record.SubItems(17) = !INSCode
                    End If
                    
                    If Len(!PLNCode) Then
                       Record.SubItems(18) = !PLNCode
                    End If
                    
                    If Len(!HOSName) Then
                       Record.SubItems(19) = Trim(!HOSName)
                    End If
                    
                    If Len(!FirstDiag) Then
                       Record.SubItems(20) = !FirstDiag
                       FirstICDCode = Record.SubItems(20)
                       
                       strAppendCase = "1"
                       StrAppend1 = " AND ICDCode = '" & Trim(FirstICDCode) & "'"
                       Call GenFirstICD(StrAppend1, strAppendCase, Record, rst2)
                    
                    End If
                    
                    If Len(!FinalDiag) Then
                       Record.SubItems(22) = !FinalDiag
                       FinalICDCode = Record.SubItems(22)
                                              
                       strAppendCase = "1"
                       StrAppend1 = " AND ICDCode = '" & Trim(FinalICDCode) & "'"
                       Call GenFinalICD(StrAppend1, strAppendCase, Record, rst2)
                                            
                    End If
                    
                    If Len(!CLMTotalActual) Then
                       Record.SubItems(24) = Format(!CLMTotalActual, "#,##0.00;(#,##0.00)")
                    End If
                    
                    If Len(!CLMTotalApproved) Then
                       Record.SubItems(25) = Format(!CLMTotalApproved, "#,##0.00;(#,##0.00)")
                    End If
                    
                    If Len(!CASReserveAmount) Then
                       Record.SubItems(26) = Format(!CASReserveAmount, "#,##0.00;(#,##0.00)")
                    End If
                                        
                    If Len(!RecoveryAmt) Then
                       
                       If !RecoveryAmt < 0 Then
                            Record.SubItems(27) = Format(!RecoveryAmt, "#,##0.00;(#,##0.00)")
                       Else
                            Record.SubItems(27) = "0.00"
                       End If
                    End If
                    
                    If Len(!MBMPremium) Then
                       Record.SubItems(28) = Format(!MBMPremium, "#,##0.00;(#,##0.00)")
                    End If
                    
                    If Len(!CASStatus) Then
                       Record.SubItems(29) = !CASStatus
                    End If
                    
                    If Len(!CASClaimability) Then
                       Record.SubItems(30) = !CASClaimability
                    End If
                    
                    If Len(!MBMStatus) Then
                       Record.SubItems(31) = !MBMStatus
                    End If
                    
                    If Len(!CASDateAdmitted) Then
                        Record.SubItems(32) = Format(!CASDateAdmitted, "dd/mm/yyyy")
                    End If
                    
                    If Len(!CASDischargeDate) Then
                        Record.SubItems(33) = Format(!CASDischargeDate, "dd/mm/yyyy")
                    End If
                    
                    If Len(!CaseRegDate) Then
                        Record.SubItems(34) = Format(!CaseRegDate, "dd/mm/yyyy")
                    End If
                    
                    If !NoofDay = "0" Then
                        Record.SubItems(35) = "1"
                    ElseIf Len(!NoofDay) Then
                        Record.SubItems(35) = !NoofDay
                    End If
                    
                    If Len(!ClaimOpendate) Then
                        Record.SubItems(36) = Format(!ClaimOpendate, "dd/mm/yyyy")
                    End If
                    
                    If Len(!MBMPayorEffDate) Then
                        Record.SubItems(37) = Format(!MBMPayorEffDate, "dd/mm/yyyy")
                    End If
                    
                    If Len(!MBMPayorEXPDate) Then
                        Record.SubItems(38) = Format(!MBMPayorEXPDate, "dd/mm/yyyy")
                    End If
                    
                    If Len(!CASCptCode1) Then
                       Record.SubItems(39) = Trim(!CASCptCode1)
                    End If
                    
                    If Len(!CASOPROCRequired1) Then
                       Record.SubItems(40) = Trim(!CASOPROCRequired1)
                    End If
                    
                    If Len(!CASCptCode2) Then
                       Record.SubItems(41) = Trim(!CASCptCode2)
                    End If
                    
                    If Len(!CASOPROCRequired2) Then
                       Record.SubItems(42) = Trim(!CASOPROCRequired2)
                    End If
                    
                    If Len(!CASCptCode3) Then
                       Record.SubItems(43) = Trim(!CASCptCode3)
                    End If
                    
                    If Len(!CASOPROCRequired3) Then
                       Record.SubItems(44) = Trim(!CASOPROCRequired3)
                    End If
                    
                    If Len(!CASCptCode4) Then
                       Record.SubItems(45) = Trim(!CASCptCode4)
                    End If
                    
                    If Len(!CASOPROCRequired4) Then
                       Record.SubItems(46) = Trim(!CASOPROCRequired4)
                    End If
                    
                    If Len(!CASCptCode5) Then
                       Record.SubItems(47) = Trim(!CASCptCode5)
                    End If
                    
                    If Len(!CASOPROCRequired5) Then
                       Record.SubItems(48) = Trim(!CASOPROCRequired5)
                    End If
                    
                    Current = Current + 1
                    lblCurrent = Current
            End With
            rst2.MoveNext
            DoEvents
        
            If intExit = 1 Then
                Check = False
                Exit Do
            End If
            Loop
        Loop Until Check = False
    intExit = 0
    
    If ListView2.ListItems.Count = 0 Then
        MsgBox "No record found!", vbOKOnly + vbCritical, DialogBoxTitle
        CmdPrint.Enabled = True
        CmdClear.Enabled = True
        cmdExit.Enabled = True
    End If
    rst2.Close
    Set rst2 = Nothing
    CmdPrint.Enabled = True
    CmdClear.Enabled = True
    cmdExit.Enabled = True
    
    ElseIf cboClaimCase.ListIndex = 1 Then
    
    strAppendCase = "2"
    Set cmd.ActiveConnection = medixmasterconnWS
        cmd.CommandText = "DCUReport"
        cmd.CommandType = adCmdStoredProc
        cmd.CommandTimeout = 300
        Set prm1 = cmd.CreateParameter("StrAppend", adVarChar, adParamInput, 2000, StrAppend)
        cmd.Parameters.Append prm1
        Set prm2 = cmd.CreateParameter("StrAppendCase", adVarChar, adParamInput, 2000, strAppendCase)
        cmd.Parameters.Append prm2
        rst2.CursorLocation = adUseClient
        rst2.CursorType = adOpenForwardOnly
        rst2.CacheSize = 100
        Set rst2 = cmd.Execute
        
    Do
              
       Do Until rst2.EOF
       
       Total = rst2.recordCount
       lblTotalRecord = Total
        
        FirstICDCode = ""
        FinalICDCode = ""
        FirstDescription = ""
        FinalDescription = ""
        CaseNo = ""
        BordxRefNo = ""
        
            With rst2
                    
                    If Len(!PAYCode) Then
                        Set Record = ListView2.ListItems.Add(, , !PAYCode)
                    Else
                        Set Record = ListView2.ListItems.Add(, , "Empty")
                    End If
                    
                    If Len(!CasNumber) Then
                       Record.SubItems(1) = !CasNumber
                       CaseNo = Trim(!CasNumber)
                    End If
                    
                    If Len(!MBMNumber) Then
                       Record.SubItems(2) = !MBMNumber
                    End If
                    
                    If Len(!MBMPolicyNo) Then
                       Record.SubItems(3) = Trim(!MBMPolicyNo)
                    End If
                                    
                    If Len(!MBMPatientCovID) Then
                       Record.SubItems(4) = !MBMPatientCovID
                    End If
                    
                    If Len(!MBMName) Then
                       Record.SubItems(5) = Trim(!MBMName)
                    End If
                    
                    If Len(!PatientName) Then
                       Record.SubItems(6) = Trim(!PatientName)
                    End If
                    
                    If Len(!MBMIcBcPp) Then
                       Record.SubItems(7) = !MBMIcBcPp
                    End If
                    
                    If Len(!MBMSex) Then
                       Record.SubItems(8) = !MBMSex
                    End If
                    
                    If Len(!AgeCode) Then
                       Record.SubItems(9) = !AgeCode
                    End If
                    
                    If Len(!MBMDOB) Then
                        Record.SubItems(10) = Format(!MBMDOB, "dd/mm/yyyy")
                    End If
                    
                    If Len(!ICNo) Then
                       Record.SubItems(11) = !ICNo
                    End If
                    
                    If Len(!Sex) Then
                       Record.SubItems(12) = !Sex
                    End If
                    
                    If Len(!AgeBand) Then
                       Record.SubItems(13) = !AgeBand
                    End If
                    
                    If Len(!DOB) Then
                        Record.SubItems(14) = Format(!DOB, "dd/mm/yyyy")
                    End If
                    
                    If Len(!INSCode) Then
                       Record.SubItems(15) = !INSCode
                    End If
                    
                    If Len(!PLNCode) Then
                       Record.SubItems(16) = !PLNCode
                    End If
                    
                    If Len(!HOSName) Then
                       Record.SubItems(17) = Trim(!HOSName)
                    End If
                    
                    If Len(!FirstDiag) Then
                       Record.SubItems(18) = !FirstDiag
                       FirstICDCode = Record.SubItems(18)
                       
                       strAppendCase = "1"
                       StrAppend1 = " AND ICDCode = '" & Trim(FirstICDCode) & "'"
                       Call GenFirstICD(StrAppend1, strAppendCase, Record, rst2)
                       
                    End If
                                                            
                    If Len(!FinalDiag) Then
                       Record.SubItems(20) = !FinalDiag
                       FinalICDCode = Record.SubItems(20)
                       
                       strAppendCase = "1"
                       StrAppend1 = " AND ICDCode = '" & Trim(FinalICDCode) & "'"
                       Call GenFinalICD(StrAppend1, strAppendCase, Record, rst2)
                       
                    End If
                    
                    If Len(!CLMTotalActual) Then
                       Record.SubItems(22) = Format(!CLMTotalActual, "#,##0.00;(#,##0.00)")
                    End If
                    
                    If Len(!CLMTotalApproved) Then
                       Record.SubItems(23) = Format(!CLMTotalApproved, "#,##0.00;(#,##0.00)")
                    End If
                    
                    If Len(!CASReserveAmount) Then
                       Record.SubItems(26) = Format(!CASReserveAmount, "#,##0.00;(#,##0.00)")
                    End If
                                        
                    If Len(!RecoveryAmt) Then
                       Record.SubItems(27) = Format(!RecoveryAmt, "#,##0.00;(#,##0.00)")
                    End If
                    
                    If Len(!MBMPremium) Then
                       Record.SubItems(28) = Format(!MBMPremium, "#,##0.00;(#,##0.00)")
                    End If
                    
                    If Len(!CASStatus) Then
                       Record.SubItems(29) = !CASStatus
                    End If
                    
                    If Len(!CASClaimability) Then
                       Record.SubItems(30) = !CASClaimability
                    End If
                    
                    If Len(!MBMStatus) Then
                       Record.SubItems(31) = !MBMStatus
                    End If
                    
                    If Len(!CASDateAdmitted) Then
                        Record.SubItems(32) = Format(!CASDateAdmitted, "dd/mm/yyyy")
                    End If
                    
                    If Len(!CASDischargeDate) Then
                        Record.SubItems(33) = Format(!CASDischargeDate, "dd/mm/yyyy")
                    End If
                    
                    If Len(!CaseRegDate) Then
                        Record.SubItems(34) = Format(!CaseRegDate, "dd/mm/yyyy")
                    End If
                    
                    If !NoofDay = "0" Then
                        Record.SubItems(35) = "1"
                    ElseIf Len(!NoofDay) Then
                        Record.SubItems(35) = !NoofDay
                    End If
                    
                    If Len(!MBMPayorEffDate) Then
                        Record.SubItems(36) = Format(!MBMPayorEffDate, "dd/mm/yyyy")
                    End If
                    
                    If Len(!MBMPayorEXPDate) Then
                        Record.SubItems(37) = Format(!MBMPayorEXPDate, "dd/mm/yyyy")
                    End If
                    
                    If Len(!CASCptCode1) Then
                       Record.SubItems(38) = Trim(!CASCptCode1)
                    End If
                    
                    If Len(!CASOPROCRequired1) Then
                       Record.SubItems(39) = Trim(!CASOPROCRequired1)
                    End If
                    
                    If Len(!CASCptCode2) Then
                       Record.SubItems(40) = Trim(!CASCptCode2)
                    End If
                    
                    If Len(!CASOPROCRequired2) Then
                       Record.SubItems(41) = Trim(!CASOPROCRequired2)
                    End If
                    
                    If Len(!CASCptCode3) Then
                       Record.SubItems(42) = Trim(!CASCptCode3)
                    End If
                    
                    If Len(!CASOPROCRequired3) Then
                       Record.SubItems(43) = Trim(!CASOPROCRequired3)
                    End If
                    
                    If Len(!CASCptCode4) Then
                       Record.SubItems(44) = Trim(!CASCptCode4)
                    End If
                    
                    If Len(!CASOPROCRequired4) Then
                       Record.SubItems(45) = Trim(!CASOPROCRequired4)
                    End If
                    
                    If Len(!CASCptCode5) Then
                       Record.SubItems(46) = Trim(!CASCptCode5)
                    End If
                    
                    If Len(!CASOPROCRequired5) Then
                       Record.SubItems(47) = Trim(!CASOPROCRequired5)
                    End If
                    
                    If Len(CaseNo) Then
                       strAppendCase = "3"
                       StrAppend1 = " AND CASNumber = '" & Trim(CaseNo) & "'"
                       Call GenBordxRefNo1(StrAppend1, strAppendCase, Record, rst2)
                       Call GenBordxRefNo2(StrAppend1, strAppendCase, Record, rst2)
                    End If
                    Current = Current + 1
                    lblCurrent = Current
            End With
            rst2.MoveNext
            DoEvents
        
            If intExit = 1 Then
                Check = False
                Exit Do
            End If
            Loop
        Loop Until Check = False
    intExit = 0
  
    If ListView2.ListItems.Count = 0 Then
        MsgBox "No record found!", vbOKOnly + vbCritical, DialogBoxTitle
        CmdPrint.Enabled = True
        CmdClear.Enabled = True
        cmdExit.Enabled = True
    End If
    rst2.Close
    Set rst2 = Nothing
    CmdPrint.Enabled = True
    CmdClear.Enabled = True
    cmdExit.Enabled = True
    
    End If
End Sub

Private Sub ResultListing2(Optional StrAppend As String)
Dim rst1 As New ADODB.Recordset
Dim rst2 As New ADODB.Recordset
Dim rst3 As New ADODB.Recordset
Dim Record As ListItem
Dim sql As String
Dim sql2 As String
Dim sql3 As String
Dim sql4 As String
Dim sql5 As String
Dim FirstICDCode As String
Dim FinalICDCode As String
Dim FirstDescription As String
Dim FinalDescription As String
Dim CaseNo As String
Dim strCount
        
    ListView2.ListItems.Clear
    Current = 0
    Total = 0
          
    If cboClaimCase.ListIndex = 0 Then
    
    sql = " SELECT PAYCode, CASNumber, claimversion, MBMNumber, MBMPolicyNo, " & _
          " MBMPatientCovID, MBMName, MBMDOB, DOB, PatientName, MBMSex, AgeCode, Sex, Ageband, " & _
          " InsCode, PLNCode, HOSName, FirstDiag, FinalDiag, " & _
          " CLMTotalActual, CLMTotalApproved, CASReserveAmount, " & _
          " RecoveryAmt, BordxRefNo, BordxDate, ClaimOpenDate, " & _
          " CASClaimability, CASStatus, MBMStatus, CASDateAdmitted, " & _
          " CASDischargeDate, MBMPayorEffDate, MBMPayorExpDate, MBMPremium, " & _
          " NoofDay, CaseRegDate, CASORemarks, MBMIcBcPp, ICNo, CASCptCode1, CASCptCode2, CASCptCode3, " & _
          " CASCptCode4, CASCptCode5, CASOPROCRequired1, CASOPROCRequired2, CASOPROCRequired3, " & _
          " CASOPROCRequired4, CASOPROCRequired5 From VIEW_CLMReportDCU Where 0 = 0 "
    
    If Len(Trim(StrAppend)) Then
        sql = sql + StrAppend
    End If
            
    rst2.Open sql, medixmasterconnWS, adOpenForwardOnly, adLockBatchOptimistic
    
    Do
              
       Do Until rst2.EOF
       
       Total = rst2.recordCount
       lblTotalRecord = Total
        
        FirstICDCode = ""
        FinalICDCode = ""
        FirstDescription = ""
        FinalDescription = ""
        CaseNo = ""
        
            With rst2
                    
                    If Len(!PAYCode) Then
                        Set Record = ListView2.ListItems.Add(, , !PAYCode)
                    Else
                        Set Record = ListView2.ListItems.Add(, , "Empty")
                    End If
                    
                    If Len(rst2!CasNumber & "-" & rst2!ClaimVersion) Then
                       Record.SubItems(1) = (rst2!CasNumber & "-" & rst2!ClaimVersion)
                    End If
                    
                    If Len(rst2!BordxRefNo) Then
                       Record.SubItems(2) = rst2!BordxRefNo
                    End If
                    
                    If Len(rst2!BordxDate) Then
                       Record.SubItems(3) = Format(rst2!BordxDate, "dd/mm/yyyy")
                    End If
                    
                    If Len(rst2!MBMNumber) Then
                       Record.SubItems(4) = rst2!MBMNumber
                    End If
                    
                    If Len(rst2!MBMPolicyNo) Then
                       Record.SubItems(5) = Trim(rst2!MBMPolicyNo)
                    End If
                                    
                    If Len(rst2!MBMPatientCovID) Then
                       Record.SubItems(6) = rst2!MBMPatientCovID
                    End If
                    
                    If Len(rst2!MBMName) Then
                       Record.SubItems(7) = Trim(rst2!MBMName)
                    End If
                    
                    If Len(rst2!PatientName) Then
                       Record.SubItems(8) = Trim(rst2!PatientName)
                    End If
                    
                    If Len(!MBMIcBcPp) Then
                       Record.SubItems(9) = !MBMIcBcPp
                    End If
                    
                    If Len(!MBMSex) Then
                       Record.SubItems(10) = !MBMSex
                    End If
                    
                    If Len(!AgeCode) Then
                       Record.SubItems(11) = !AgeCode
                    End If
                    
                    If Len(!MBMDOB) Then
                        Record.SubItems(12) = Format(!MBMDOB, "dd/mm/yyyy")
                    End If
                    
                    If Len(!ICNo) Then
                       Record.SubItems(13) = !ICNo
                    End If
                    
                    If Len(!Sex) Then
                       Record.SubItems(14) = !Sex
                    End If
                    
                    If Len(!AgeBand) Then
                       Record.SubItems(15) = !AgeBand
                    End If
                    
                    If Len(!DOB) Then
                        Record.SubItems(16) = Format(!DOB, "dd/mm/yyyy")
                    End If
                    
                    If Len(!INSCode) Then
                       Record.SubItems(17) = !INSCode
                    End If
                    
                    If Len(!PLNCode) Then
                       Record.SubItems(18) = !PLNCode
                    End If
                    
                    If Len(!HOSName) Then
                       Record.SubItems(19) = Trim(!HOSName)
                    End If
                    
                    If Len(!FirstDiag) Then
                       Record.SubItems(20) = !FirstDiag
                       FirstICDCode = Record.SubItems(20)
                       sql3 = " SELECT ICDCode, ICDDescription From ICDOriginal Where ICDCode = '" & Trim(FirstICDCode) & "'"
                       rst3.Open sql3, medixmasterconn, adOpenForwardOnly, adLockBatchOptimistic
                       
                       If rst3.EOF = True Then
                            Record.SubItems(21) = ""
                            rst3.Close
                            Set rst3 = Nothing
                       Else
                            Record.SubItems(21) = Trim(rst3!ICDDescription)
                            rst3.Close
                            Set rst3 = Nothing
                       End If
                    End If
                                                            
                    If Len(!FinalDiag) Then
                       Record.SubItems(22) = !FinalDiag
                       FinalICDCode = Record.SubItems(22)
                       sql3 = " SELECT ICDCode, ICDDescription From ICDOriginal Where ICDCode = '" & Trim(FinalICDCode) & "'"
                       rst3.Open sql3, medixmasterconn, adOpenForwardOnly, adLockBatchOptimistic
                       
                       If rst3.EOF = True Then
                            Record.SubItems(23) = ""
                            rst3.Close
                            Set rst3 = Nothing
                       Else
                            Record.SubItems(23) = Trim(rst3!ICDDescription)
                            rst3.Close
                            Set rst3 = Nothing
                       End If
                    End If
                    
                    If Len(!CLMTotalActual) Then
                       Record.SubItems(24) = Format(!CLMTotalActual, "#,##0.00;(#,##0.00)")
                    End If
                    
                    If Len(!CLMTotalApproved) Then
                       Record.SubItems(25) = Format(!CLMTotalApproved, "#,##0.00;(#,##0.00)")
                    End If
                    
                    If Len(!CASReserveAmount) Then
                       Record.SubItems(26) = Format(!CASReserveAmount, "#,##0.00;(#,##0.00)")
                    End If
                                        
                    If Len(!RecoveryAmt) Then
                       
                       If !RecoveryAmt < 0 Then
                            Record.SubItems(27) = Format(!RecoveryAmt, "#,##0.00;(#,##0.00)")
                       Else
                            Record.SubItems(27) = "0.00"
                       End If
                    End If
                    
                    If Len(!MBMPremium) Then
                       Record.SubItems(28) = Format(!MBMPremium, "#,##0.00;(#,##0.00)")
                    End If
                    
                    If Len(!CASStatus) Then
                       Record.SubItems(29) = !CASStatus
                    End If
                    
                    If Len(!CASClaimability) Then
                       Record.SubItems(30) = !CASClaimability
                    End If
                    
                    If Len(!MBMStatus) Then
                       Record.SubItems(31) = !MBMStatus
                    End If
                    
                    If Len(!CASDateAdmitted) Then
                        Record.SubItems(32) = Format(!CASDateAdmitted, "dd/mm/yyyy")
                    End If
                    
                    If Len(!CASDischargeDate) Then
                        Record.SubItems(33) = Format(!CASDischargeDate, "dd/mm/yyyy")
                    End If
                    
                    If Len(!CaseRegDate) Then
                        Record.SubItems(34) = Format(!CaseRegDate, "dd/mm/yyyy")
                    End If
                    
                    If !NoofDay = "0" Then
                        Record.SubItems(35) = "1"
                    ElseIf Len(!NoofDay) Then
                        Record.SubItems(35) = !NoofDay
                    End If
                    
                    If Len(!ClaimOpendate) Then
                        Record.SubItems(36) = Format(!ClaimOpendate, "dd/mm/yyyy")
                    End If
                    
                    If Len(!MBMPayorEffDate) Then
                        Record.SubItems(37) = Format(!MBMPayorEffDate, "dd/mm/yyyy")
                    End If
                    
                    If Len(!MBMPayorEXPDate) Then
                        Record.SubItems(38) = Format(!MBMPayorEXPDate, "dd/mm/yyyy")
                    End If
                    
                    If Len(rst2!CASORemarks) Then
                        Record.SubItems(39) = Trim(rst2!CASORemarks)
                    End If
                    
                    If Len(!CASCptCode1) Then
                       Record.SubItems(40) = Trim(!CASCptCode1)
                    End If
                    
                    If Len(!CASOPROCRequired1) Then
                       Record.SubItems(41) = Trim(!CASOPROCRequired1)
                    End If
                    
                    If Len(!CASCptCode2) Then
                       Record.SubItems(42) = Trim(!CASCptCode2)
                    End If
                    
                    If Len(!CASOPROCRequired2) Then
                       Record.SubItems(43) = Trim(!CASOPROCRequired2)
                    End If
                    
                    If Len(!CASCptCode3) Then
                       Record.SubItems(44) = Trim(!CASCptCode3)
                    End If
                    
                    If Len(!CASOPROCRequired3) Then
                       Record.SubItems(45) = Trim(!CASOPROCRequired3)
                    End If
                    
                    If Len(!CASCptCode4) Then
                       Record.SubItems(46) = Trim(!CASCptCode4)
                    End If
                    
                    If Len(!CASOPROCRequired4) Then
                       Record.SubItems(47) = Trim(!CASOPROCRequired4)
                    End If
                    
                    If Len(!CASCptCode5) Then
                       Record.SubItems(48) = Trim(!CASCptCode5)
                    End If
                    
                    If Len(!CASOPROCRequired5) Then
                       Record.SubItems(49) = Trim(!CASOPROCRequired5)
                    End If
                    
                    Current = Current + 1
                    lblCurrent = Current
            End With
            rst2.MoveNext
            DoEvents
        
            If intExit = 1 Then
                Check = False
                Exit Do
            End If
            Loop
        Loop Until Check = False
    intExit = 0
    'End If
    If ListView2.ListItems.Count = 0 Then
        MsgBox "No record found!", vbOKOnly + vbCritical, DialogBoxTitle
        CmdPrint.Enabled = True
        CmdClear.Enabled = True
        cmdExit.Enabled = True
    End If
    rst2.Close
    Set rst2 = Nothing
    CmdPrint.Enabled = True
    CmdClear.Enabled = True
    cmdExit.Enabled = True
    
    ElseIf cboClaimCase.ListIndex = 1 Then
    
    sql = " SELECT PAYCode, CASNumber, MBMNumber, MBMPolicyNo, " & _
          " MBMPatientCovID, MBMName, MBMDOB, DOB, PatientName, MBMSex, AgeCode, Sex, Ageband, " & _
          " InsCode, PLNCode, HOSName, FirstDiag, FinalDiag, MBMPremium, " & _
          " SUM(CLMTotalActual) AS CLMTotalActual, SUM(CLMTotalApproved) AS CLMTotalApproved, " & _
          " SUM(CASReserveAmount) AS CASReserveAmount, SUM(RecoveryAmt) AS RecoveryAmt, " & _
          " CASClaimability, CASStatus, MBMStatus, CASDateAdmitted, " & _
          " CASDischargeDate, MBMPayorEffDate, MBMPayorExpDate, " & _
          " NoofDay, CaseRegDate, CASORemarks, MBMIcBcPp, ICNo, CASCptCode1, CASCptCode2, CASCptCode3, " & _
          " CASCptCode4, CASCptCode5, CASOPROCRequired1, CASOPROCRequired2, CASOPROCRequired3, " & _
          " CASOPROCRequired4, CASOPROCRequired5 From VIEW_CLMReportDCU Where 0 = 0 "
    
    sql2 = " Group By PAYCode, CASNumber, MBMNumber, MBMPolicyNo, " & _
           " MBMPatientCovID, MBMName, MBMDOB, DOB, PatientName, MBMSex, AgeCode, Sex, Ageband, " & _
           " InsCode, PLNCode, HOSName, FirstDiag, FinalDiag, MBMPremium, " & _
           " CASClaimability, CASStatus, MBMStatus, CASDateAdmitted, " & _
           " CASDischargeDate, MBMPayorEffDate, MBMPayorExpDate, " & _
           " NoofDay, CaseRegDate, CASORemarks, MBMIcBcPp, ICNo, CASCptCode1, CASCptCode2, CASCptCode3, " & _
          " CASCptCode4, CASCptCode5, CASOPROCRequired1, CASOPROCRequired2, CASOPROCRequired3, " & _
          " CASOPROCRequired4, CASOPROCRequired5 "
    
    If Len(Trim(StrAppend)) Then
        sql = sql + StrAppend + sql2
    End If
            
    rst2.Open sql, medixmasterconnWS, adOpenForwardOnly, adLockBatchOptimistic
    
    Do
              
       Do Until rst2.EOF
       
       Total = rst2.recordCount
       lblTotalRecord = Total
        
        FirstICDCode = ""
        FinalICDCode = ""
        FirstDescription = ""
        FinalDescription = ""
        
            With rst2
                    
                    If Len(!PAYCode) Then
                        Set Record = ListView2.ListItems.Add(, , !PAYCode)
                    Else
                        Set Record = ListView2.ListItems.Add(, , "Empty")
                    End If
                    
                    If Len(rst2!CasNumber) Then
                       Record.SubItems(1) = Trim(rst2!CasNumber)
                       CaseNo = Trim(rst2!CasNumber)
                    End If
                    
                    If Len(rst2!MBMNumber) Then
                       Record.SubItems(2) = rst2!MBMNumber
                    End If
                    
                    If Len(rst2!MBMPolicyNo) Then
                       Record.SubItems(3) = Trim(rst2!MBMPolicyNo)
                    End If
                                    
                    If Len(rst2!MBMPatientCovID) Then
                       Record.SubItems(4) = rst2!MBMPatientCovID
                    End If
                    
                    If Len(rst2!MBMName) Then
                       Record.SubItems(5) = Trim(rst2!MBMName)
                    End If
                    
                    If Len(rst2!PatientName) Then
                       Record.SubItems(6) = Trim(rst2!PatientName)
                    End If
                    
                    If Len(!MBMIcBcPp) Then
                       Record.SubItems(7) = !MBMIcBcPp
                    End If
                    
                    If Len(!MBMSex) Then
                       Record.SubItems(8) = !MBMSex
                    End If
                    
                    If Len(!AgeCode) Then
                       Record.SubItems(9) = !AgeCode
                    End If
                    
                    If Len(!MBMDOB) Then
                        Record.SubItems(10) = Format(!MBMDOB, "dd/mm/yyyy")
                    End If
                    
                    If Len(!ICNo) Then
                       Record.SubItems(11) = !ICNo
                    End If
                    
                    If Len(!Sex) Then
                       Record.SubItems(12) = !Sex
                    End If
                    
                    If Len(!AgeBand) Then
                       Record.SubItems(13) = !AgeBand
                    End If
                    
                    If Len(!DOB) Then
                        Record.SubItems(14) = Format(!DOB, "dd/mm/yyyy")
                    End If
                    
                    If Len(!INSCode) Then
                       Record.SubItems(15) = !INSCode
                    End If
                    
                    If Len(!PLNCode) Then
                       Record.SubItems(16) = !PLNCode
                    End If
                    
                    If Len(!HOSName) Then
                       Record.SubItems(17) = Trim(!HOSName)
                    End If
                    
                    If Len(!FirstDiag) Then
                       Record.SubItems(18) = !FirstDiag
                       FirstICDCode = Record.SubItems(18)
                       sql3 = " SELECT ICDCode, ICDDescription From ICDOriginal Where ICDCode = '" & Trim(FirstICDCode) & "'"
                       rst3.Open sql3, medixmasterconn, adOpenForwardOnly, adLockBatchOptimistic
                       
                       If rst3.EOF = True Then
                            Record.SubItems(19) = ""
                            rst3.Close
                            Set rst3 = Nothing
                       Else
                            Record.SubItems(19) = Trim(rst3!ICDDescription)
                            rst3.Close
                            Set rst3 = Nothing
                       End If
                    End If
                                                            
                    If Len(!FinalDiag) Then
                       Record.SubItems(20) = !FinalDiag
                       FinalICDCode = Record.SubItems(20)
                       sql3 = " SELECT ICDCode, ICDDescription From ICDOriginal Where ICDCode = '" & Trim(FinalICDCode) & "'"
                       rst3.Open sql3, medixmasterconn, adOpenForwardOnly, adLockBatchOptimistic
                       
                       If rst3.EOF = True Then
                            Record.SubItems(21) = ""
                            rst3.Close
                            Set rst3 = Nothing
                       Else
                            Record.SubItems(21) = Trim(rst3!ICDDescription)
                            rst3.Close
                            Set rst3 = Nothing
                       End If
                    End If
                    
                    If Len(!CLMTotalActual) Then
                       Record.SubItems(22) = Format(!CLMTotalActual, "#,##0.00;(#,##0.00)")
                    End If
                    
                    If Len(!CLMTotalApproved) Then
                       Record.SubItems(23) = Format(!CLMTotalApproved, "#,##0.00;(#,##0.00)")
                    End If
                    
                    If Len(!CASReserveAmount) Then
                       Record.SubItems(26) = Format(!CASReserveAmount, "#,##0.00;(#,##0.00)")
                    End If
                                        
                    If Len(!RecoveryAmt) Then
                       
                       If !RecoveryAmt < 0 Then
                            Record.SubItems(27) = Format(!RecoveryAmt, "#,##0.00;(#,##0.00)")
                       Else
                            Record.SubItems(27) = "0.00"
                       End If
                    End If
                    
                    If Len(!MBMPremium) Then
                       Record.SubItems(28) = Format(!MBMPremium, "#,##0.00;(#,##0.00)")
                    End If
                    
                    If Len(!CASStatus) Then
                       Record.SubItems(29) = !CASStatus
                    End If
                    
                    If Len(!CASClaimability) Then
                       Record.SubItems(30) = !CASClaimability
                    End If
                    
                    If Len(!MBMStatus) Then
                       Record.SubItems(31) = !MBMStatus
                    End If
                    
                    If Len(!CASDateAdmitted) Then
                        Record.SubItems(32) = Format(!CASDateAdmitted, "dd/mm/yyyy")
                    End If
                    
                    If Len(!CASDischargeDate) Then
                        Record.SubItems(33) = Format(!CASDischargeDate, "dd/mm/yyyy")
                    End If
                    
                    If Len(!CaseRegDate) Then
                        Record.SubItems(34) = Format(!CaseRegDate, "dd/mm/yyyy")
                    End If
                    
                    If !NoofDay = "0" Then
                        Record.SubItems(35) = "1"
                    ElseIf Len(!NoofDay) Then
                        Record.SubItems(35) = !NoofDay
                    End If
                    
                    If Len(!MBMPayorEffDate) Then
                        Record.SubItems(36) = Format(!MBMPayorEffDate, "dd/mm/yyyy")
                    End If
                    
                    If Len(!MBMPayorEXPDate) Then
                        Record.SubItems(37) = Format(!MBMPayorEXPDate, "dd/mm/yyyy")
                    End If
                    
                    If Len(rst2!CASORemarks) Then
                        Record.SubItems(38) = Trim(rst2!CASORemarks)
                    End If
                    
                    If Len(!CASCptCode1) Then
                       Record.SubItems(39) = Trim(!CASCptCode1)
                    End If
                    
                    If Len(!CASOPROCRequired1) Then
                       Record.SubItems(40) = Trim(!CASOPROCRequired1)
                    End If
                    
                    If Len(!CASCptCode2) Then
                       Record.SubItems(41) = Trim(!CASCptCode2)
                    End If
                    
                    If Len(!CASOPROCRequired2) Then
                       Record.SubItems(42) = Trim(!CASOPROCRequired2)
                    End If
                    
                    If Len(!CASCptCode3) Then
                       Record.SubItems(43) = Trim(!CASCptCode3)
                    End If
                    
                    If Len(!CASOPROCRequired3) Then
                       Record.SubItems(44) = Trim(!CASOPROCRequired3)
                    End If
                    
                    If Len(!CASCptCode4) Then
                       Record.SubItems(45) = Trim(!CASCptCode4)
                    End If
                    
                    If Len(!CASOPROCRequired4) Then
                       Record.SubItems(46) = Trim(!CASOPROCRequired4)
                    End If
                    
                    If Len(!CASCptCode5) Then
                       Record.SubItems(47) = Trim(!CASCptCode5)
                    End If
                    
                    If Len(!CASOPROCRequired5) Then
                       Record.SubItems(48) = Trim(!CASOPROCRequired5)
                    End If
                    
                    
                    If Len(CaseNo) Then
                       sql5 = " SELECT CASNumber, SUM(CLMTotalApproved) AS CLMTotalApproved From Claim Where CASNumber = '" & Trim(CaseNo) & "' AND (BordxRefno is null or BordxRefno = '')" & _
                              " Group by CASNumber  "
                       
                       strCount = " select count(*) as totalRecord From Claim where CASNumber = '" & Trim(CaseNo) & "' AND (BordxRefno is null or BordxRefno = '')"
                       Set rstCount = medixmasterconnWS.Execute(strCount)
                       rst1.Open sql5, medixmasterconnWS, adOpenForwardOnly, adLockBatchOptimistic
                       
                       If rstCount!TotalRecord <> 0 Then
                            Record.SubItems(24) = Format(rst1!CLMTotalApproved, "#,##0.00;(#,##0.00)")
                            rst1.Close
                            Set rst1 = Nothing
                       Else
                            rst1.Close
                            Set rst1 = Nothing
                       End If
                    End If
                    
                    If Len(CaseNo) Then
                       sql5 = " SELECT CASNumber, SUM(CLMTotalApproved) AS CLMTotalApproved From Claim Where CASNumber = '" & Trim(CaseNo) & "' AND (BordxRefno <> '')" & _
                              " Group by CASNumber "
                       
                       strCount = " select count(*) as totalRecord From Claim where CASNumber = '" & Trim(CaseNo) & "' AND (BordxRefno <> '')"
                       Set rstCount = medixmasterconnWS.Execute(strCount)
                       rst1.Open sql5, medixmasterconnWS, adOpenForwardOnly, adLockBatchOptimistic
                       
                       If rstCount!TotalRecord <> 0 Then
                            Record.SubItems(25) = Format(rst1!CLMTotalApproved, "#,##0.00;(#,##0.00)")
                            rst1.Close
                            Set rst1 = Nothing
                       Else
                            rst1.Close
                            Set rst1 = Nothing
                       End If
                    End If
                    
                    Current = Current + 1
                    lblCurrent = Current
            End With
            rst2.MoveNext
            DoEvents
        
            If intExit = 1 Then
                Check = False
                Exit Do
            End If
            Loop
        Loop Until Check = False
    intExit = 0
    'End If
    If ListView2.ListItems.Count = 0 Then
        MsgBox "No record found!", vbOKOnly + vbCritical, DialogBoxTitle
        CmdPrint.Enabled = True
        CmdClear.Enabled = True
        cmdExit.Enabled = True
    End If
    rst2.Close
    Set rst2 = Nothing
    CmdPrint.Enabled = True
    CmdClear.Enabled = True
    cmdExit.Enabled = True
    
    End If
End Sub

Private Sub INSListing(Optional StrAppend As String)
Dim rst2 As New ADODB.Recordset
Dim cmd As New ADODB.Command
Dim prm1 As ADODB.Parameter
Dim prm2 As ADODB.Parameter
Dim strAppendCase As String
Dim StrAppend1 As String
Dim Record As ListItem

    lblCurrent = 0
    Payor = ""
    INSCode = ""
    ListView2.ListItems.Clear
        
    If cboClaimCase.ListIndex = 0 Then
        strAppendCase = "4"
    'sql = " SELECT PAYCode, INSCode, COUNT(CASNumber) as TotalCases, SUM(clmtotalactual) AS totalActual, " & _
          " SUM(clmtotalapproved) As TotalApproved, SUM(NoofDay) AS NoofDay " & _
          " From VIEW_CLMReportDCU Where 0 = 0 "
          
    ElseIf cboClaimCase.ListIndex = 1 Then
        strAppendCase = "5"
    'sql = " SELECT PAYCode, INSCode, COUNT(DISTINCT CASNumber) as TotalCases, SUM(clmtotalactual) AS totalActual, " & _
          " SUM(clmtotalapproved) As TotalApproved " & _
          " From VIEW_CLMReportDCU Where 0 = 0 "
    
    End If
    
    'sql2 = " Group By PAYCode, INSCode "
    
    'If Len(Trim(StrAppend)) Then
    '    sql = sql + StrAppend + sql2
    'Else
    '    sql = sql + sql2
    'End If
    
    'rst2.Open sql, medixmasterconnWS, adOpenForwardOnly, adLockBatchOptimistic
    Set cmd.ActiveConnection = medixmasterconnWS
        cmd.CommandText = "DCUReport"
        cmd.CommandType = adCmdStoredProc
        cmd.CommandTimeout = 300
        Set prm1 = cmd.CreateParameter("StrAppend", adVarChar, adParamInput, 2000, StrAppend)
        cmd.Parameters.Append prm1
        Set prm2 = cmd.CreateParameter("StrAppendCase", adVarChar, adParamInput, 2000, strAppendCase)
        cmd.Parameters.Append prm2
        rst2.CursorLocation = adUseClient
        rst2.CursorType = adOpenForwardOnly
        rst2.CacheSize = 100
        Set rst2 = cmd.Execute
                
        Do Until rst2.EOF
        
            With rst2
                If Len(rst2!PAYCode) Then
                    Set Record = ListView2.ListItems.Add(, , rst2!PAYCode)
                    Payor = rst2!PAYCode
                Else
                    Set Record = ListView2.ListItems.Add(, , "Empty")
                    Payor = ""
                End If
                
                If Len(rst2!INSCode) Then
                    Record.SubItems(1) = rst2!INSCode
                    INSCode = rst2!INSCode
                Else
                    Record.SubItems(1) = "Empty"
                    INSCode = ""
                End If
                
                If Len(rst2!TotalCases) Then
                    Record.SubItems(2) = rst2!TotalCases
                End If
                
                If cboClaimCase.ListIndex = 0 Then
                
                    If rst2!NoofDay = "0" Then
                        Record.SubItems(3) = "1"
                    ElseIf Len(rst2!NoofDay) Then
                        Record.SubItems(3) = rst2!NoofDay
                    End If
                
                ElseIf cboClaimCase.ListIndex = 1 Then
                
                    If Payor = "" And INSCode = "" Then
                        strAppendCase = "6"
                        StrAppend1 = " AND (PAYCode = '' OR PAYCode IS NULL) AND " & _
                           " (INSCode = '' OR INSCode IS NULL)"
                        StrAppend1 = StrAppend1 + StrAppend
                    'sql3 = " SELECT PAYCode, INSCode, SUM(DurationStay) AS NoofDay " & _
                           " From VIEW_CLMReportDCU2 Where (PAYCode = '' OR PAYCode IS NULL) AND " & _
                           " (INSCode = '' OR INSCode IS NULL)"
                           
                    ElseIf Payor <> "" And INSCode = "" Then
                        strAppendCase = "6"
                        StrAppend1 = " AND PAYCode = '" & Trim(Payor) & "' AND " & _
                           " (INSCode = '' OR INSCode IS NULL)"
                        StrAppend1 = StrAppend1 + StrAppend
                    'sql3 = " SELECT PAYCode, INSCode, SUM(DurationStay) AS NoofDay " & _
                           " From VIEW_CLMReportDCU2 Where PAYCode = '" & Trim(Payor) & "' AND " & _
                           " (INSCode = '' OR INSCode IS NULL)"
                           
                    ElseIf Payor = "" And INSCode <> "" Then
                        strAppendCase = "6"
                        StrAppend1 = " AND (PAYCode = '' OR PAYCode IS NULL) AND " & _
                           " INSCode = '" & Trim(INSCode) & "'"
                        StrAppend1 = StrAppend1 + StrAppend
                    'sql3 = " SELECT PAYCode, INSCode, SUM(DurationStay) AS NoofDay " & _
                           " From VIEW_CLMReportDCU2 Where (PAYCode = '' OR PAYCode IS NULL) AND " & _
                           " INSCode = '" & Trim(INSCode) & "'"
                           
                    Else
                        strAppendCase = "6"
                        StrAppend1 = " AND PAYCode = '" & Trim(Payor) & "' AND " & _
                           " INSCode = '" & Trim(INSCode) & "'"
                        StrAppend1 = StrAppend1 + StrAppend
                    'sql3 = " SELECT PAYCode, INSCode, SUM(DurationStay) AS NoofDay " & _
                           " From VIEW_CLMReportDCU2 Where PAYCode = '" & Trim(Payor) & "' AND " & _
                           " INSCode = '" & Trim(INSCode) & "'"
                           
                    End If
                       
                       'sql2 = " Group By PAYCode, INSCode  "
                       
                       'If Len(Trim(StrAppend)) Then
                       '     sql3 = sql3 + StrAppend + sql2
                       ' Else
                       '     sql3 = sql3 + sql2
                       ' End If
                                              
                       'rst1.Open sql3, medixmasterconnWS, adOpenForwardOnly, adLockBatchOptimistic
                        Call DurationStay(StrAppend1, strAppendCase, Record)
                End If
                
                If Len(rst2!TotalActual) Then
                    Record.SubItems(4) = Format(rst2!TotalActual, "#,##0.00")
                End If
                
                If Len(rst2!TotalApproved) Then
                    Record.SubItems(5) = Format(rst2!TotalApproved, "#,##0.00")
                End If
                
            End With
            lblCurrent = lblCurrent + 1
            rst2.MoveNext
       Loop
    If ListView2.ListItems.Count = 0 Then
        MsgBox "No record found!", vbOKOnly + vbCritical, DialogBoxTitle
        CmdPrint.Enabled = True
        CmdClear.Enabled = True
        cmdExit.Enabled = True
    End If
    rst2.Close
    Set rst2 = Nothing
    CmdPrint.Enabled = True
    CmdClear.Enabled = True
    cmdExit.Enabled = True
End Sub

Private Sub PayorListing(Optional StrAppend As String)
Dim rst2 As New ADODB.Recordset
Dim cmd As New ADODB.Command
Dim prm1 As ADODB.Parameter
Dim prm2 As ADODB.Parameter
Dim strAppendCase As String
Dim StrAppend1 As String
Dim Record As ListItem

    lblCurrent = 0
    Payor = ""
    ListView2.ListItems.Clear
        
    If cboClaimCase.ListIndex = 0 Then
        strAppendCase = "52"
    'sql = " SELECT PAYCode, COUNT(CASNumber) as TotalCases, SUM(clmtotalactual) AS totalActual, " & _
          " SUM(clmtotalapproved) As TotalApproved, SUM(NoofDay) AS NoofDay From VIEW_CLMReportDCU Where 0 = 0 "
          
    ElseIf cboClaimCase.ListIndex = 1 Then
        strAppendCase = "53"
    'sql = " SELECT PAYCode, COUNT(DISTINCT CASNumber) as TotalCases, SUM(clmtotalactual) AS totalActual, " & _
          " SUM(clmtotalapproved) As TotalApproved From VIEW_CLMReportDCU Where 0 = 0 "
          
    End If
    
    'sql2 = " Group By PAYCode Order By PAYCode"
    
    'If Len(Trim(StrAppend)) Then
    '    sql = sql + StrAppend + sql2
    'Else
    '    sql = sql + sql2
    'End If
     
    'rst2.Open sql, medixmasterconnWS, adOpenForwardOnly, adLockBatchOptimistic
    Set cmd.ActiveConnection = medixmasterconnWS
        cmd.CommandText = "DCUReport"
        cmd.CommandType = adCmdStoredProc
        cmd.CommandTimeout = 300
        Set prm1 = cmd.CreateParameter("StrAppend", adVarChar, adParamInput, 2000, StrAppend)
        cmd.Parameters.Append prm1
        Set prm2 = cmd.CreateParameter("StrAppendCase", adVarChar, adParamInput, 2000, strAppendCase)
        cmd.Parameters.Append prm2
        rst2.CursorLocation = adUseClient
        rst2.CursorType = adOpenForwardOnly
        rst2.CacheSize = 100
        Set rst2 = cmd.Execute
        
        Do Until rst2.EOF
        
            With rst2
                If Len(rst2!PAYCode) Then
                    Set Record = ListView2.ListItems.Add(, , rst2!PAYCode)
                    Payor = rst2!PAYCode
                Else
                    Set Record = ListView2.ListItems.Add(, , "Empty")
                    Payor = ""
                End If
                
                If Len(rst2!TotalCases) Then
                    Record.SubItems(1) = rst2!TotalCases
                End If
                
                If cboClaimCase.ListIndex = 0 Then
                
                    If rst2!NoofDay = "0" Then
                        Record.SubItems(2) = "1"
                    ElseIf Len(rst2!NoofDay) Then
                        Record.SubItems(2) = rst2!NoofDay
                    End If
                
                ElseIf cboClaimCase.ListIndex = 1 Then
                
                    If Payor = "" Then
                        strAppendCase = "54"
                        StrAppend1 = " AND (PAYCode = '' OR PAYCode IS NULL)"
                        StrAppend1 = StrAppend1 + StrAppend
                    'sql3 = " SELECT PAYCode, SUM(DurationStay) AS NoofDay " & _
                           " From VIEW_CLMReportDCU2 Where (PAYCode = '' OR PAYCode IS NULL) "
                           
                    Else
                        strAppendCase = "54"
                        StrAppend1 = " AND PAYCode = '" & Trim(Payor) & "'"
                        StrAppend1 = StrAppend1 + StrAppend
                    'sql3 = " SELECT PAYCode, SUM(DurationStay) AS NoofDay " & _
                           " From VIEW_CLMReportDCU2 Where PAYCode = '" & Trim(Payor) & "'"
                           
                    End If
                       
                       'sql2 = " Group By PAYCode "
                       
                       'If Len(Trim(StrAppend)) Then
                       '     sql3 = sql3 + StrAppend + sql2
                       ' Else
                       '     sql3 = sql3 + sql2
                       ' End If
                                              
                       'rst1.Open sql3, medixmasterconnWS, adOpenForwardOnly, adLockBatchOptimistic
                        Call DurationStay(StrAppend1, strAppendCase, Record)
                End If
                
                If Len(rst2!TotalActual) Then
                    Record.SubItems(3) = Format(rst2!TotalActual, "#,##0.00")
                End If
                
                If Len(rst2!TotalApproved) Then
                    Record.SubItems(4) = Format(rst2!TotalApproved, "#,##0.00")
                End If
                
            End With
            lblCurrent = lblCurrent + 1
            rst2.MoveNext
       Loop
    If ListView2.ListItems.Count = 0 Then
        MsgBox "No record found!", vbOKOnly + vbCritical, DialogBoxTitle
        CmdPrint.Enabled = True
        CmdClear.Enabled = True
        cmdExit.Enabled = True
    End If
    rst2.Close
    Set rst2 = Nothing
    CmdPrint.Enabled = True
    CmdClear.Enabled = True
    cmdExit.Enabled = True
End Sub

Private Sub ListViewHeader6()
If opt35 = True Then
    ListView2.ColumnHeaders(1).Text = "Claim Type"
    ListView2.ColumnHeaders(2).Text = "No of Claims"
    ListView2.ColumnHeaders(2).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(3).Text = "Total Duration"
    ListView2.ColumnHeaders(3).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(4).Text = "Total Actual Amt"
    ListView2.ColumnHeaders(4).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(5).Text = "Total Approved Amt"
    ListView2.ColumnHeaders(5).Alignment = lvwColumnRight
    
    For AA = 6 To 63
        ListView2.ColumnHeaders(AA).Text = ""
    Next AA
ElseIf opt34 = True Then
    ListView2.ColumnHeaders(1).Text = "Claim Type"
    ListView2.ColumnHeaders(2).Text = "No of Claims"
    ListView2.ColumnHeaders(2).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(3).Text = "No of Follow-Up"
    ListView2.ColumnHeaders(3).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(4).Text = "No of Pre-Hosp."
    ListView2.ColumnHeaders(4).Alignment = lvwColumnRight
        
    For AA = 5 To 63
        ListView2.ColumnHeaders(AA).Text = ""
    Next AA
End If
End Sub

Private Sub ListViewHeader8()
    
If cboClaimCase.ListIndex = 0 And opt29 = True Then
    ListView2.ColumnHeaders(1).Text = "Diag. Code"
    ListView2.ColumnHeaders(2).Text = "Diag. Desc"
    ListView2.ColumnHeaders(2).Alignment = lvwColumnLeft
    ListView2.ColumnHeaders(3).Text = "Hospital"
    ListView2.ColumnHeaders(3).Alignment = lvwColumnLeft
    ListView2.ColumnHeaders(4).Text = "No of Claims"
    ListView2.ColumnHeaders(4).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(5).Text = "Total Duration"
    ListView2.ColumnHeaders(5).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(6).Text = "Total Actual Amt"
    ListView2.ColumnHeaders(6).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(7).Text = "Total Approved Amt"
    ListView2.ColumnHeaders(7).Alignment = lvwColumnRight
    
    For AA = 8 To 63
        ListView2.ColumnHeaders(AA).Text = ""
    Next AA

ElseIf cboClaimCase.ListIndex = 1 And opt29 = True Then
    ListView2.ColumnHeaders(1).Text = "Diag. Code"
    ListView2.ColumnHeaders(2).Text = "Diag. Desc"
    ListView2.ColumnHeaders(2).Alignment = lvwColumnLeft
    ListView2.ColumnHeaders(3).Text = "Hospital"
    ListView2.ColumnHeaders(3).Alignment = lvwColumnLeft
    ListView2.ColumnHeaders(4).Text = "No of Cases"
    ListView2.ColumnHeaders(4).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(5).Text = "Total Duration"
    ListView2.ColumnHeaders(5).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(6).Text = "Total Actual Amt"
    ListView2.ColumnHeaders(6).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(7).Text = "Total Approved Amt"
    ListView2.ColumnHeaders(7).Alignment = lvwColumnRight
    
    For AA = 8 To 63
        ListView2.ColumnHeaders(AA).Text = ""
    Next AA
ElseIf cboClaimCase.ListIndex = 0 And opt35 = True Then
    ListView2.ColumnHeaders(1).Text = "Diag. Code"
    ListView2.ColumnHeaders(2).Text = "Diag. Desc"
    ListView2.ColumnHeaders(2).Alignment = lvwColumnLeft
    ListView2.ColumnHeaders(3).Text = "No of Claims"
    ListView2.ColumnHeaders(3).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(4).Text = "Total Duration"
    ListView2.ColumnHeaders(4).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(5).Text = "Total Actual Amt"
    ListView2.ColumnHeaders(5).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(6).Text = "Total Approved Amt"
    ListView2.ColumnHeaders(6).Alignment = lvwColumnRight
    
    For AA = 7 To 63
        ListView2.ColumnHeaders(AA).Text = ""
    Next AA
ElseIf cboClaimCase.ListIndex = 0 And opt34 = True Then
    ListView2.ColumnHeaders(1).Text = "Diag. Code"
    ListView2.ColumnHeaders(2).Text = "Diag. Desc"
    ListView2.ColumnHeaders(2).Alignment = lvwColumnLeft
    ListView2.ColumnHeaders(3).Text = "No of Claims"
    ListView2.ColumnHeaders(3).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(4).Text = "Total Duration"
    ListView2.ColumnHeaders(4).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(5).Text = "Total Actual Amt"
    ListView2.ColumnHeaders(5).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(6).Text = "Total Approved Amt"
    ListView2.ColumnHeaders(6).Alignment = lvwColumnRight
    
    For AA = 7 To 63
        ListView2.ColumnHeaders(AA).Text = ""
    Next AA
ElseIf cboClaimCase.ListIndex = 0 And opt33 = True Then
    ListView2.ColumnHeaders(1).Text = "Diag. Code"
    ListView2.ColumnHeaders(2).Text = "Diag. Desc"
    ListView2.ColumnHeaders(2).Alignment = lvwColumnLeft
    ListView2.ColumnHeaders(3).Text = "No of Claims"
    ListView2.ColumnHeaders(3).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(4).Text = "Total Duration"
    ListView2.ColumnHeaders(4).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(5).Text = "Total Actual Amt"
    ListView2.ColumnHeaders(5).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(6).Text = "Total Approved Amt"
    ListView2.ColumnHeaders(6).Alignment = lvwColumnRight
    
    For AA = 7 To 63
        ListView2.ColumnHeaders(AA).Text = ""
    Next AA
ElseIf cboClaimCase.ListIndex = 1 And opt35 = True Then
    ListView2.ColumnHeaders(1).Text = "Diag. Code"
    ListView2.ColumnHeaders(2).Text = "Diag. Desc"
    ListView2.ColumnHeaders(2).Alignment = lvwColumnLeft
    ListView2.ColumnHeaders(3).Text = "No of Cases"
    ListView2.ColumnHeaders(3).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(4).Text = "Total Duration"
    ListView2.ColumnHeaders(4).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(5).Text = "Total Actual Amt"
    ListView2.ColumnHeaders(5).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(6).Text = "Total Approved Amt"
    ListView2.ColumnHeaders(6).Alignment = lvwColumnRight
    
    For AA = 7 To 63
        ListView2.ColumnHeaders(AA).Text = ""
    Next AA
ElseIf cboClaimCase.ListIndex = 1 And opt34 = True Then
    ListView2.ColumnHeaders(1).Text = "Diag. Code"
    ListView2.ColumnHeaders(2).Text = "Diag. Desc"
    ListView2.ColumnHeaders(2).Alignment = lvwColumnLeft
    ListView2.ColumnHeaders(3).Text = "No of Cases"
    ListView2.ColumnHeaders(3).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(4).Text = "Total Duration"
    ListView2.ColumnHeaders(4).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(5).Text = "Total Actual Amt"
    ListView2.ColumnHeaders(5).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(6).Text = "Total Approved Amt"
    ListView2.ColumnHeaders(6).Alignment = lvwColumnRight
    
    For AA = 7 To 63
        ListView2.ColumnHeaders(AA).Text = ""
    Next AA
ElseIf cboClaimCase.ListIndex = 1 And opt33 = True Then
    ListView2.ColumnHeaders(1).Text = "Diag. Code"
    ListView2.ColumnHeaders(2).Text = "Diag. Desc"
    ListView2.ColumnHeaders(2).Alignment = lvwColumnLeft
    ListView2.ColumnHeaders(3).Text = "No of Cases"
    ListView2.ColumnHeaders(3).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(4).Text = "Total Duration"
    ListView2.ColumnHeaders(4).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(5).Text = "Total Actual Amt"
    ListView2.ColumnHeaders(5).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(6).Text = "Total Approved Amt"
    ListView2.ColumnHeaders(6).Alignment = lvwColumnRight
    
    For AA = 7 To 63
        ListView2.ColumnHeaders(AA).Text = ""
    Next AA
End If
End Sub

Private Sub DurationStay(StrAppend1 As String, strAppendCase As String, Record)
Dim rst1 As New ADODB.Recordset
Dim conn As New ADODB.Command
Dim prm1 As ADODB.Parameter
Dim prm2 As ADODB.Parameter

Set conn.ActiveConnection = medixmasterconnWS
    conn.CommandText = "DCUReport"
    conn.CommandType = adCmdStoredProc
    conn.CommandTimeout = 300
    Set prm1 = conn.CreateParameter("StrAppend", adVarChar, adParamInput, 2000, StrAppend1)
    conn.Parameters.Append prm1
    Set prm2 = conn.CreateParameter("StrAppendCase", adVarChar, adParamInput, 2000, strAppendCase)
    conn.Parameters.Append prm2
    rst1.CursorLocation = adUseClient
    rst1.CursorType = adOpenForwardOnly
    rst1.CacheSize = 100
    Set rst1 = conn.Execute
    
    If cboReportType.ListIndex = 0 Then
    ElseIf cboReportType.ListIndex = 1 Then
        If cboReportSelection.ListIndex = 0 Then
            If rst1!NoofDay = "0" Then
                Record.SubItems(2) = "1"
                rst1.Close
                Set rst1 = Nothing
            Else
                Record.SubItems(2) = rst1!NoofDay
                rst1.Close
                Set rst1 = Nothing
            End If
        ElseIf cboReportSelection.ListIndex = 1 Then
            If rst1!NoofDay = "0" Then
                Record.SubItems(3) = "1"
                rst1.Close
                Set rst1 = Nothing
            Else
                Record.SubItems(3) = rst1!NoofDay
                rst1.Close
                Set rst1 = Nothing
            End If
        ElseIf cboReportSelection.ListIndex = 2 Then
            If rst1!NoofDay = "0" Then
                Record.SubItems(4) = "1"
                rst1.Close
                Set rst1 = Nothing
            Else
                Record.SubItems(4) = rst1!NoofDay
                rst1.Close
                Set rst1 = Nothing
            End If
        ElseIf cboReportSelection.ListIndex = 3 Then
            If rst1!NoofDay = "0" Then
                Record.SubItems(3) = "1"
                rst1.Close
                Set rst1 = Nothing
            Else
                Record.SubItems(3) = rst1!NoofDay
                rst1.Close
                Set rst1 = Nothing
            End If
        ElseIf cboReportSelection.ListIndex = 4 Then
            If rst1!NoofDay = "0" Then
                Record.SubItems(4) = "1"
                rst1.Close
                Set rst1 = Nothing
            Else
                Record.SubItems(4) = rst1!NoofDay
                rst1.Close
                Set rst1 = Nothing
            End If
        ElseIf cboReportSelection.ListIndex = 5 Then
            If rst1!NoofDay = "0" Then
                Record.SubItems(4) = "1"
                rst1.Close
                Set rst1 = Nothing
            Else
                Record.SubItems(4) = rst1!NoofDay
                rst1.Close
                Set rst1 = Nothing
            End If
        End If
        
    ElseIf cboReportType.ListIndex = 2 Then
        If cboReportSelection.ListIndex = 0 Then
            If rst1!NoofDay = "0" Then
                Record.SubItems(3) = "1"
                rst1.Close
                Set rst1 = Nothing
            Else
                Record.SubItems(3) = rst1!NoofDay
                rst1.Close
                Set rst1 = Nothing
            End If
        ElseIf cboReportSelection.ListIndex = 1 Then
            If rst1!NoofDay = "0" Then
                Record.SubItems(3) = "1"
                rst1.Close
                Set rst1 = Nothing
            Else
                Record.SubItems(3) = rst1!NoofDay
                rst1.Close
                Set rst1 = Nothing
            End If
        ElseIf cboReportSelection.ListIndex = 2 Then
            If rst1!NoofDay = "0" Then
                Record.SubItems(3) = "1"
                rst1.Close
                Set rst1 = Nothing
            Else
                Record.SubItems(3) = rst1!NoofDay
                rst1.Close
                Set rst1 = Nothing
            End If
        ElseIf cboReportSelection.ListIndex = 3 Then
            If rst1!NoofDay = "0" Then
                Record.SubItems(4) = "1"
                rst1.Close
                Set rst1 = Nothing
            Else
                Record.SubItems(4) = rst1!NoofDay
                rst1.Close
                Set rst1 = Nothing
            End If
        End If
        
    ElseIf cboReportType.ListIndex = 3 Then
        If cboReportSelection.ListIndex = 0 Then
            
        ElseIf cboReportSelection.ListIndex = 1 Then
            If rst1!NoofDay = "0" Then
                Record.SubItems(3) = "1"
                rst1.Close
                Set rst1 = Nothing
            Else
                Record.SubItems(3) = rst1!NoofDay
                rst1.Close
                Set rst1 = Nothing
            End If
        ElseIf cboReportSelection.ListIndex = 2 Then
            If rst1!NoofDay = "0" Then
                Record.SubItems(2) = "1"
                rst1.Close
                Set rst1 = Nothing
            Else
                Record.SubItems(2) = rst1!NoofDay
                rst1.Close
                Set rst1 = Nothing
            End If
        ElseIf cboReportSelection.ListIndex = 3 Then
            If rst1!NoofDay = "0" Then
                Record.SubItems(3) = "1"
                rst1.Close
                Set rst1 = Nothing
            Else
                Record.SubItems(3) = rst1!NoofDay
                rst1.Close
                Set rst1 = Nothing
            End If
        ElseIf cboReportSelection.ListIndex = 4 Then
            If rst1!NoofDay = "0" Then
                Record.SubItems(3) = "1"
                rst1.Close
                Set rst1 = Nothing
            Else
                Record.SubItems(3) = rst1!NoofDay
                rst1.Close
                Set rst1 = Nothing
            End If
        ElseIf cboReportSelection.ListIndex = 5 Then
            If rst1!NoofDay = "0" Then
                Record.SubItems(2) = "1"
                rst1.Close
                Set rst1 = Nothing
            Else
                Record.SubItems(2) = rst1!NoofDay
                rst1.Close
                Set rst1 = Nothing
            End If
        ElseIf cboReportSelection.ListIndex = 6 Then
            If rst1!NoofDay = "0" Then
                Record.SubItems(4) = "1"
                rst1.Close
                Set rst1 = Nothing
            Else
                Record.SubItems(4) = rst1!NoofDay
                rst1.Close
                Set rst1 = Nothing
            End If
        End If
        
    ElseIf cboReportType.ListIndex = 4 Then
        If cboClaimCase.ListIndex = 0 Then
            If cboReportSelection.ListIndex = 0 Then
                If rst1!NoofDay = "0" Then
                    Record.SubItems(2) = "1"
                    rst1.Close
                    Set rst1 = Nothing
                Else
                    Record.SubItems(2) = rst1!NoofDay
                    rst1.Close
                    Set rst1 = Nothing
                End If
            ElseIf cboReportSelection.ListIndex = 1 Then
                If rst1!NoofDay = "0" Then
                    Record.SubItems(2) = "1"
                    rst1.Close
                    Set rst1 = Nothing
                Else
                    Record.SubItems(2) = rst1!NoofDay
                    rst1.Close
                    Set rst1 = Nothing
                End If
            ElseIf cboReportSelection.ListIndex = 2 Then
                If rst1!NoofDay = "0" Then
                    Record.SubItems(2) = "1"
                    rst1.Close
                    Set rst1 = Nothing
                Else
                    Record.SubItems(2) = rst1!NoofDay
                    rst1.Close
                    Set rst1 = Nothing
                End If
            ElseIf cboReportSelection.ListIndex = 3 Then
                If rst1!NoofDay = "0" Then
                    Record.SubItems(2) = "1"
                    rst1.Close
                    Set rst1 = Nothing
                Else
                    Record.SubItems(2) = rst1!NoofDay
                    rst1.Close
                    Set rst1 = Nothing
                End If
            ElseIf cboReportSelection.ListIndex = 4 Then
                If rst1!NoofDay = "0" Then
                    Record.SubItems(2) = "1"
                    rst1.Close
                    Set rst1 = Nothing
                Else
                    Record.SubItems(2) = rst1!NoofDay
                    rst1.Close
                    Set rst1 = Nothing
                End If
            ElseIf cboReportSelection.ListIndex = 5 Then
            ElseIf cboReportSelection.ListIndex = 6 Then
                If rst1!NoofDay = "0" Then
                    Record.SubItems(3) = "1"
                    rst1.Close
                    Set rst1 = Nothing
                Else
                    Record.SubItems(3) = rst1!NoofDay
                    rst1.Close
                    Set rst1 = Nothing
                End If
            ElseIf cboReportSelection.ListIndex = 7 Then
                If rst1!NoofDay = "0" Then
                    Record.SubItems(3) = "1"
                    rst1.Close
                    Set rst1 = Nothing
                Else
                    Record.SubItems(3) = rst1!NoofDay
                    rst1.Close
                    Set rst1 = Nothing
                End If
            ElseIf cboReportSelection.ListIndex = 8 Then
                If rst1!NoofDay = "0" Then
                    Record.SubItems(3) = "1"
                    rst1.Close
                    Set rst1 = Nothing
                Else
                    Record.SubItems(3) = rst1!NoofDay
                    rst1.Close
                    Set rst1 = Nothing
                End If
            ElseIf cboReportSelection.ListIndex = 9 Then
                If rst1!NoofDay = "0" Then
                    Record.SubItems(3) = "1"
                    rst1.Close
                    Set rst1 = Nothing
                Else
                    Record.SubItems(3) = rst1!NoofDay
                    rst1.Close
                    Set rst1 = Nothing
                End If
            ElseIf cboReportSelection.ListIndex = 10 Then
                If rst1!NoofDay = "0" Then
                    Record.SubItems(3) = "1"
                    rst1.Close
                    Set rst1 = Nothing
                Else
                    Record.SubItems(3) = rst1!NoofDay
                    rst1.Close
                    Set rst1 = Nothing
                End If
            ElseIf cboReportSelection.ListIndex = 11 Then
            End If
            
        ElseIf cboClaimCase.ListIndex = 1 Then
            If cboReportSelection.ListIndex = 0 Then
                If rst1!NoofDay = "0" Then
                    Record.SubItems(2) = "1"
                    rst1.Close
                    Set rst1 = Nothing
                Else
                    Record.SubItems(2) = rst1!NoofDay
                    rst1.Close
                    Set rst1 = Nothing
                End If
            ElseIf cboReportSelection.ListIndex = 1 Then
                If rst1!NoofDay = "0" Then
                    Record.SubItems(2) = "1"
                    rst1.Close
                    Set rst1 = Nothing
                Else
                    Record.SubItems(2) = rst1!NoofDay
                    rst1.Close
                    Set rst1 = Nothing
                End If
            ElseIf cboReportSelection.ListIndex = 2 Then
                If rst1!NoofDay = "0" Then
                    Record.SubItems(2) = "1"
                    rst1.Close
                    Set rst1 = Nothing
                Else
                    Record.SubItems(2) = rst1!NoofDay
                    rst1.Close
                    Set rst1 = Nothing
                End If
            ElseIf cboReportSelection.ListIndex = 3 Then
                If rst1!NoofDay = "0" Then
                    Record.SubItems(2) = "1"
                    rst1.Close
                    Set rst1 = Nothing
                Else
                    Record.SubItems(2) = rst1!NoofDay
                    rst1.Close
                    Set rst1 = Nothing
                End If
            ElseIf cboReportSelection.ListIndex = 4 Then
                If rst1!NoofDay = "0" Then
                    Record.SubItems(2) = "1"
                    rst1.Close
                    Set rst1 = Nothing
                Else
                    Record.SubItems(2) = rst1!NoofDay
                    rst1.Close
                    Set rst1 = Nothing
                End If
            ElseIf cboReportSelection.ListIndex = 5 Then
                If rst1!NoofDay = "0" Then
                    Record.SubItems(3) = "1"
                    rst1.Close
                    Set rst1 = Nothing
                Else
                    Record.SubItems(3) = rst1!NoofDay
                    rst1.Close
                    Set rst1 = Nothing
                End If
            ElseIf cboReportSelection.ListIndex = 6 Then
                If rst1!NoofDay = "0" Then
                    Record.SubItems(3) = "1"
                    rst1.Close
                    Set rst1 = Nothing
                Else
                    Record.SubItems(3) = rst1!NoofDay
                    rst1.Close
                    Set rst1 = Nothing
                End If
            ElseIf cboReportSelection.ListIndex = 7 Then
                If rst1!NoofDay = "0" Then
                    Record.SubItems(3) = "1"
                    rst1.Close
                    Set rst1 = Nothing
                Else
                    Record.SubItems(3) = rst1!NoofDay
                    rst1.Close
                    Set rst1 = Nothing
                End If
            ElseIf cboReportSelection.ListIndex = 8 Then
                If rst1!NoofDay = "0" Then
                    Record.SubItems(3) = "1"
                    rst1.Close
                    Set rst1 = Nothing
                Else
                    Record.SubItems(3) = rst1!NoofDay
                    rst1.Close
                    Set rst1 = Nothing
                End If
            ElseIf cboReportSelection.ListIndex = 9 Then
                If rst1!NoofDay = "0" Then
                    Record.SubItems(3) = "1"
                    rst1.Close
                    Set rst1 = Nothing
                Else
                    Record.SubItems(3) = rst1!NoofDay
                    rst1.Close
                    Set rst1 = Nothing
                End If
            End If
        End If
      
    ElseIf cboReportType.ListIndex = 5 Then
        If cboReportSelection.ListIndex = 0 Then
            If rst1!NoofDay = "0" Then
                Record.SubItems(3) = "1"
                rst1.Close
                Set rst1 = Nothing
            Else
                Record.SubItems(3) = rst1!NoofDay
                rst1.Close
                Set rst1 = Nothing
            End If
        ElseIf cboReportSelection.ListIndex = 1 Then
            If rst1!NoofDay = "0" Then
                Record.SubItems(4) = "1"
                rst1.Close
                Set rst1 = Nothing
            Else
                Record.SubItems(4) = rst1!NoofDay
                rst1.Close
                Set rst1 = Nothing
            End If
        ElseIf cboReportSelection.ListIndex = 2 Then
            If rst1!NoofDay = "0" Then
                Record.SubItems(4) = "1"
                rst1.Close
                Set rst1 = Nothing
            Else
                Record.SubItems(4) = rst1!NoofDay
                rst1.Close
                Set rst1 = Nothing
            End If
        ElseIf cboReportSelection.ListIndex = 3 Then
            If rst1!NoofDay = "0" Then
                Record.SubItems(4) = "1"
                rst1.Close
                Set rst1 = Nothing
            Else
                Record.SubItems(4) = rst1!NoofDay
                rst1.Close
                Set rst1 = Nothing
            End If
        ElseIf cboReportSelection.ListIndex = 4 Then
            If rst1!NoofDay = "0" Then
                Record.SubItems(4) = "1"
                rst1.Close
                Set rst1 = Nothing
            Else
                Record.SubItems(4) = rst1!NoofDay
                rst1.Close
                Set rst1 = Nothing
            End If
        End If
        
    ElseIf cboReportType.ListIndex = 6 Then
        If rst1!NoofDay = "0" Then
            Record.SubItems(2) = "1"
            rst1.Close
            Set rst1 = Nothing
        Else
            Record.SubItems(2) = rst1!NoofDay
            rst1.Close
            Set rst1 = Nothing
        End If
        
    ElseIf cboReportType.ListIndex = 7 Then
        If cboReportSelection.ListIndex = 0 Then
            If rst1!NoofDay = "0" Then
                Record.SubItems(3) = "1"
                rst1.Close
                Set rst1 = Nothing
            Else
                Record.SubItems(3) = rst1!NoofDay
                rst1.Close
                Set rst1 = Nothing
            End If
        ElseIf cboReportSelection.ListIndex = 1 Then
            If rst1!NoofDay = "0" Then
                Record.SubItems(4) = "1"
                rst1.Close
                Set rst1 = Nothing
            Else
                Record.SubItems(4) = rst1!NoofDay
                rst1.Close
                Set rst1 = Nothing
            End If
        ElseIf cboReportSelection.ListIndex = 2 Then
            If rst1!NoofDay = "0" Then
                Record.SubItems(5) = "1"
                rst1.Close
                Set rst1 = Nothing
            Else
                Record.SubItems(5) = rst1!NoofDay
                rst1.Close
                Set rst1 = Nothing
            End If
        ElseIf cboReportSelection.ListIndex = 3 Then
            If rst1!NoofDay = "0" Then
                Record.SubItems(4) = "1"
                rst1.Close
                Set rst1 = Nothing
            Else
                Record.SubItems(4) = rst1!NoofDay
                rst1.Close
                Set rst1 = Nothing
            End If
        ElseIf cboReportSelection.ListIndex = 4 Then
            If rst1!NoofDay = "0" Then
                Record.SubItems(4) = "1"
                rst1.Close
                Set rst1 = Nothing
            Else
                Record.SubItems(4) = rst1!NoofDay
                rst1.Close
                Set rst1 = Nothing
            End If
        ElseIf cboReportSelection.ListIndex = 5 Then
            If rst1!NoofDay = "0" Then
                Record.SubItems(3) = "1"
                rst1.Close
                Set rst1 = Nothing
            Else
                Record.SubItems(3) = rst1!NoofDay
                rst1.Close
                Set rst1 = Nothing
            End If
        ElseIf cboReportSelection.ListIndex = 6 Then
            If rst1!NoofDay = "0" Then
                Record.SubItems(5) = "1"
                rst1.Close
                Set rst1 = Nothing
            Else
                Record.SubItems(5) = rst1!NoofDay
                rst1.Close
                Set rst1 = Nothing
            End If
        End If
            
    ElseIf cboReportType.ListIndex = 8 Then
        If rst1!NoofDay = "0" Then
            Record.SubItems(3) = "1"
            rst1.Close
            Set rst1 = Nothing
        Else
            Record.SubItems(3) = rst1!NoofDay
            rst1.Close
            Set rst1 = Nothing
        End If
        
    ElseIf cboReportType.ListIndex = 9 Then
    End If

End Sub

'********************** List Ageband Results *******************************
Private Sub AgeGroupListing(Optional StrAppend As String)
Dim rst2 As New ADODB.Recordset
Dim cmd As New ADODB.Command
Dim prm1 As ADODB.Parameter
Dim prm2 As ADODB.Parameter
Dim strAppendCase As String
Dim StrAppend1 As String
Dim Record As ListItem
Dim strCount

    lblCurrent = 0
    AgeCode = ""
    ListView2.ListItems.Clear
            
    If cboClaimCase.ListIndex = 0 Then
        strAppendCase = "37"
    'sql = " SELECT Ageband, COUNT(CASNumber) AS TotalCases, " & _
          " SUM(CLMTotalActual) AS TotalActualAmt, " & _
          " SUM(CLMTotalApproved) AS TotalApprovedAmt, " & _
          " SUM(NoofDay) AS NoofDay From VIEW_CLMReportDCU Where 0 = 0 "
          
    ElseIf cboClaimCase.ListIndex = 1 Then
        strAppendCase = "38"
    'sql = " SELECT Ageband, COUNT(DISTINCT CASNumber) AS TotalCases, " & _
          " SUM(CLMTotalActual) AS TotalActualAmt, " & _
          " SUM(CLMTotalApproved) AS TotalApprovedAmt " & _
          " From VIEW_CLMReportDCU Where 0 = 0 "
    
    End If
          
    'sql2 = " GROUP BY Ageband ORDER BY Ageband "
    
    'If Len(Trim(StrAppend)) Then
    '    sql = sql + StrAppend + sql2
    'Else
    '    sql = sql + sql2
    'End If
    
    'rst2.Open sql, medixmasterconnWS, adOpenForwardOnly, adLockBatchOptimistic
    Set cmd.ActiveConnection = medixmasterconnWS
        cmd.CommandText = "DCUReport"
        cmd.CommandType = adCmdStoredProc
        cmd.CommandTimeout = 300
        Set prm1 = cmd.CreateParameter("StrAppend", adVarChar, adParamInput, 2000, StrAppend)
        cmd.Parameters.Append prm1
        Set prm2 = cmd.CreateParameter("StrAppendCase", adVarChar, adParamInput, 2000, strAppendCase)
        cmd.Parameters.Append prm2
        rst2.CursorLocation = adUseClient
        rst2.CursorType = adOpenForwardOnly
        rst2.CacheSize = 100
        Set rst2 = cmd.Execute
        
        Do Until rst2.EOF
        
                If Len(rst2!AgeBand) Then
                    Set Record = ListView2.ListItems.Add(, , rst2!AgeBand)
                    AgeCode = rst2!AgeBand
                Else
                    Set Record = ListView2.ListItems.Add(, , "Empty")
                    AgeCode = ""
                End If
                
                If Len(rst2!TotalCases) Then
                   Record.SubItems(1) = rst2!TotalCases
                End If
                
                If cboClaimCase.ListIndex = 0 Then
                
                    If rst2!NoofDay = "0" Then
                        Record.SubItems(2) = "1"
                    ElseIf Len(rst2!NoofDay) Then
                        Record.SubItems(2) = rst2!NoofDay
                    End If
                
                ElseIf cboClaimCase.ListIndex = 1 Then
                
                    If AgeCode = "" Then
                        strAppendCase = "39"
                        StrAppend1 = " AND (Ageband = '' OR Ageband is NULL)"
                        StrAppend1 = StrAppend1 + StrAppend
                    'sql3 = " SELECT Ageband, SUM(DurationStay) AS NoofDay " & _
                           " From VIEW_CLMReportDCU2 Where (Ageband = '' OR Ageband is NULL)"
                    
                    Else
                        strAppendCase = "39"
                        StrAppend1 = " AND AGEBand = '" & Trim(AgeCode) & "'"
                        StrAppend1 = StrAppend1 + StrAppend
                    'sql3 = " SELECT Ageband, SUM(DurationStay) AS NoofDay " & _
                           " From VIEW_CLMReportDCU2 Where AGEBand = '" & Trim(AgeCode) & "'"
                           
                    End If
                       
                       'sql2 = " Group By Ageband "
                       
                       'If Len(Trim(StrAppend)) Then
                       '     sql3 = sql3 + StrAppend + sql2
                       ' Else
                       '     sql3 = sql3 + sql2
                       ' End If
                                              
                       'rst1.Open sql3, medixmasterconnWS, adOpenForwardOnly, adLockBatchOptimistic
                        Call DurationStay(StrAppend1, strAppendCase, Record)
                End If
                                
                If Len(rst2!TotalActualAmt) Then
                   Record.SubItems(3) = Format(rst2!TotalActualAmt, "#,##0.00")
                End If
                
                If Len(rst2!TotalApprovedAmt) Then
                   Record.SubItems(4) = Format(rst2!TotalApprovedAmt, "#,##0.00")
                End If
                            
            lblCurrent = lblCurrent + 1
            rst2.MoveNext
       Loop
    
    If ListView2.ListItems.Count = 0 Then
        MsgBox "No record found!", vbOKOnly + vbCritical, DialogBoxTitle
        CmdPrint.Enabled = True
        CmdClear.Enabled = True
        cmdExit.Enabled = True
    End If
    
    rst2.Close
    Set rst2 = Nothing
    CmdPrint.Enabled = True
    CmdClear.Enabled = True
    cmdExit.Enabled = True
End Sub
'********************** List Ageband Results *******************************

'********************** List Ageband By Gender Results *******************************
Private Sub AgeGenderListing(Optional StrAppend As String)
Dim rst2 As New ADODB.Recordset
Dim cmd As New ADODB.Command
Dim prm1 As ADODB.Parameter
Dim prm2 As ADODB.Parameter
Dim strAppendCase As String
Dim StrAppend1 As String
Dim Record As ListItem
Dim strCount
    
    lblCurrent = 0
    ListView2.ListItems.Clear
    SexCode = ""
    AgeCode = ""
    
    If cboClaimCase.ListIndex = 0 Then
        strAppendCase = "40"
    'sql = " SELECT Sex, Ageband, COUNT(CASNumber) AS TotalCases, " & _
          " SUM(CLMTotalActual) AS TotalActualAmt, " & _
          " SUM(CLMTotalApproved) AS TotalApprovedAmt, " & _
          " SUM(NoofDay) AS NoofDay From VIEW_CLMReportDCU Where 0 = 0 "
          
    ElseIf cboClaimCase.ListIndex = 1 Then
        strAppendCase = "41"
    'sql = " SELECT Sex, Ageband, COUNT(DISTINCT CASNumber) AS TotalCases, " & _
          " SUM(CLMTotalActual) AS TotalActualAmt, " & _
          " SUM(CLMTotalApproved) AS TotalApprovedAmt " & _
          " From VIEW_CLMReportDCU Where 0 = 0 "
          
    End If
          
    'sql2 = " GROUP BY Sex, Ageband ORDER BY Sex, Ageband "
   
    'If Len(Trim(StrAppend)) Then
    '    sql = sql + StrAppend + sql2
    'Else
    '    sql = sql + sql2
    'End If
    
    'rst2.Open sql, medixmasterconnWS, adOpenForwardOnly, adLockBatchOptimistic
    Set cmd.ActiveConnection = medixmasterconnWS
        cmd.CommandText = "DCUReport"
        cmd.CommandType = adCmdStoredProc
        cmd.CommandTimeout = 300
        Set prm1 = cmd.CreateParameter("StrAppend", adVarChar, adParamInput, 2000, StrAppend)
        cmd.Parameters.Append prm1
        Set prm2 = cmd.CreateParameter("StrAppendCase", adVarChar, adParamInput, 2000, strAppendCase)
        cmd.Parameters.Append prm2
        rst2.CursorLocation = adUseClient
        rst2.CursorType = adOpenForwardOnly
        rst2.CacheSize = 100
        Set rst2 = cmd.Execute
        
        Do Until rst2.EOF
            
                If Len(rst2!Sex) Then
                    Set Record = ListView2.ListItems.Add(, , rst2!Sex)
                    SexCode = rst2!Sex
                Else
                    Set Record = ListView2.ListItems.Add(, , "Empty")
                    SexCode = ""
                End If
                                              
                If Len(rst2!AgeBand) Then
                   Record.SubItems(1) = rst2!AgeBand
                   AgeCode = rst2!AgeBand
                Else
                   Record.SubItems(1) = "Empty"
                   AgeCode = ""
                End If
                                    
                If Len(rst2!TotalCases) Then
                   Record.SubItems(2) = rst2!TotalCases
                End If
                
                If cboClaimCase.ListIndex = 0 Then
                
                    If rst2!NoofDay = "0" Then
                        Record.SubItems(3) = "1"
                    ElseIf Len(rst2!NoofDay) Then
                        Record.SubItems(3) = rst2!NoofDay
                    End If
                
                ElseIf cboClaimCase.ListIndex = 1 Then
                
                    If SexCode = "" And AgeCode = "" Then
                        strAppendCase = "42"
                        StrAppend1 = " AND (Sex = '' OR Sex IS NULL) AND (Ageband = '' OR Ageband IS NULL) "
                        StrAppend1 = StrAppend1 + StrAppend
                    'sql3 = " SELECT Sex, Ageband, SUM(DurationStay) AS NoofDay " & _
                           " From VIEW_CLMReportDCU2 Where (Sex = '' OR Sex IS NULL) AND (Ageband = '' OR Ageband IS NULL) "
                           
                    ElseIf SexCode <> "" And AgeCode = "" Then
                        strAppendCase = "42"
                        StrAppend1 = " AND Sex = '" & Trim(SexCode) & "' AND (Ageband = '' OR Ageband is NULL)"
                        StrAppend1 = StrAppend1 + StrAppend
                    'sql3 = " SELECT Sex, Ageband, SUM(DurationStay) AS NoofDay " & _
                           " From VIEW_CLMReportDCU2 Where Sex = '" & Trim(SexCode) & "' AND (Ageband = '' OR Ageband is NULL)"
                           
                    ElseIf SexCode = "" And AgeCode <> "" Then
                        strAppendCase = "42"
                        StrAppend1 = " AND (Sex = '' OR Sex is NULL) AND Ageband = '" & Trim(AgeCode) & "'"
                        StrAppend1 = StrAppend1 + StrAppend
                    'sql3 = " SELECT Sex, Ageband, SUM(DurationStay) AS NoofDay " & _
                           " From VIEW_CLMReportDCU2 Where (Sex = '' OR Sex is NULL) AND Ageband = '" & Trim(AgeCode) & "' "
                           
                    Else
                        strAppendCase = "42"
                        StrAppend1 = " AND Sex = '" & Trim(SexCode) & "' AND Ageband = '" & Trim(AgeCode) & "'"
                        StrAppend1 = StrAppend1 + StrAppend
                    'sql3 = " SELECT Sex, Ageband, SUM(DurationStay) AS NoofDay " & _
                           " From VIEW_CLMReportDCU2 Where Sex = '" & Trim(SexCode) & "' AND Ageband = '" & Trim(AgeCode) & "'"
                    
                    End If
                       
                       'sql2 = " Group By Sex, Ageband "
                       
                       'If Len(Trim(StrAppend)) Then
                       '     sql3 = sql3 + StrAppend + sql2
                       ' Else
                       '     sql3 = sql3 + sql2
                       ' End If
                                              
                       'rst1.Open sql3, medixmasterconnWS, adOpenForwardOnly, adLockBatchOptimistic
                        Call DurationStay(StrAppend1, strAppendCase, Record)
                End If
                                
                If Len(rst2!TotalActualAmt) Then
                   Record.SubItems(4) = Format(rst2!TotalActualAmt, "#,##0.00")
                End If
                
                If Len(rst2!TotalApprovedAmt) Then
                   Record.SubItems(5) = Format(rst2!TotalApprovedAmt, "#,##0.00")
                End If
            
            lblCurrent = lblCurrent + 1
            rst2.MoveNext
       Loop
    
    If ListView2.ListItems.Count = 0 Then
        MsgBox "No record found!", vbOKOnly + vbCritical, DialogBoxTitle
        CmdPrint.Enabled = True
        CmdClear.Enabled = True
        cmdExit.Enabled = True
    End If
    
    rst2.Close
    Set rst2 = Nothing
    CmdPrint.Enabled = True
    CmdClear.Enabled = True
    cmdExit.Enabled = True
End Sub
'********************** List Ageband By Gender Results *******************************

'********************** List Ageband By Insured Type Results *******************************
Private Sub AgeInsuredListing(Optional StrAppend As String)
Dim rst2 As New ADODB.Recordset
Dim cmd As New ADODB.Command
Dim prm1 As ADODB.Parameter
Dim prm2 As ADODB.Parameter
Dim strAppendCase As String
Dim StrAppend1 As String
Dim Record As ListItem
   
    lblCurrent = 0
    Payor = ""
    INSCode = ""
    AgeCode = ""
    ListView2.ListItems.Clear
            
    If cboClaimCase.ListIndex = 0 Then
        strAppendCase = "10"
    'sql = " SELECT PAYCode, INSCode, Ageband, COUNT(CASNumber) AS TotalCases, " & _
          " SUM(CLMTotalActual) AS TotalActualAmt, " & _
          " SUM(CLMTotalApproved) AS TotalApprovedAmt, " & _
          " SUM(NoofDay) AS NoofDay From VIEW_CLMReportDCU Where 0 = 0 "
          
    ElseIf cboClaimCase.ListIndex = 1 Then
        strAppendCase = "11"
    'sql = " SELECT PAYCode, INSCode, Ageband, COUNT(DISTINCT CASNumber) AS TotalCases, " & _
          " SUM(CLMTotalActual) AS TotalActualAmt, " & _
          " SUM(CLMTotalApproved) AS TotalApprovedAmt " & _
          " From VIEW_CLMReportDCU Where 0 = 0 "
    
    End If
          
    'sql2 = " GROUP BY PAYCode, INSCode, Ageband "
    
    'If Len(Trim(StrAppend)) Then
    '    sql = sql + StrAppend + sql2
    'Else
    '    sql = sql + sql2
    'End If
    
    'rst2.Open sql, medixmasterconnWS, adOpenForwardOnly, adLockBatchOptimistic
    Set cmd.ActiveConnection = medixmasterconnWS
        cmd.CommandText = "DCUReport"
        cmd.CommandType = adCmdStoredProc
        cmd.CommandTimeout = 300
        Set prm1 = cmd.CreateParameter("StrAppend", adVarChar, adParamInput, 2000, StrAppend)
        cmd.Parameters.Append prm1
        Set prm2 = cmd.CreateParameter("StrAppendCase", adVarChar, adParamInput, 2000, strAppendCase)
        cmd.Parameters.Append prm2
        rst2.CursorLocation = adUseClient
        rst2.CursorType = adOpenForwardOnly
        rst2.CacheSize = 100
        Set rst2 = cmd.Execute
        
        Do Until rst2.EOF
        
                If Len(rst2!PAYCode) Then
                    Set Record = ListView2.ListItems.Add(, , rst2!PAYCode)
                    Payor = rst2!PAYCode
                Else
                    Set Record = ListView2.ListItems.Add(, , "Empty")
                    Payor = ""
                End If
                
                If Len(rst2!INSCode) Then
                   Record.SubItems(1) = rst2!INSCode
                   INSCode = rst2!INSCode
                Else
                    Record.SubItems(1) = "Empty"
                    INSCode = ""
                End If
                                              
                If Len(rst2!AgeBand) Then
                   Record.SubItems(2) = rst2!AgeBand
                   AgeCode = rst2!AgeBand
                Else
                    Record.SubItems(2) = "Empty"
                    AgeCode = ""
                End If
                         
                If Len(rst2!TotalCases) Then
                   Record.SubItems(3) = rst2!TotalCases
                End If
                
                If cboClaimCase.ListIndex = 0 Then
                
                    If rst2!NoofDay = "0" Then
                        Record.SubItems(4) = "1"
                    ElseIf Len(rst2!NoofDay) Then
                        Record.SubItems(4) = rst2!NoofDay
                    End If
                
                ElseIf cboClaimCase.ListIndex = 1 Then
                
                    If Payor = "" And INSCode = "" And AgeCode = "" Then
                        strAppendCase = "12"
                        StrAppend1 = " AND (PAYCode = '' OR PAYCode IS NULL) AND " & _
                           " (INSCode = '' OR INSCode IS NULL) AND (AGEBand = '' OR AGEBand IS NULL)"
                        StrAppend1 = StrAppend1 + StrAppend
                    'sql3 = " SELECT PAYCode, INSCode, Ageband, SUM(DurationStay) AS NoofDay " & _
                           " From VIEW_CLMReportDCU2 Where (PAYCode = '' OR PAYCode IS NULL) AND " & _
                           " (INSCode = '' OR INSCode IS NULL) AND (AGEBand = '' OR AGEBand IS NULL) "
                           
                    ElseIf Payor <> "" And INSCode = "" And AgeCode = "" Then
                        strAppendCase = "12"
                        StrAppend1 = " AND PAYCode = '" & Trim(Payor) & "' AND " & _
                           " (INSCode = '' OR INSCode IS NULL) AND (AGEBand = '' OR AGEBand IS NULL)"
                        StrAppend1 = StrAppend1 + StrAppend
                    'sql3 = " SELECT PAYCode, INSCode, Ageband, SUM(DurationStay) AS NoofDay " & _
                           " From VIEW_CLMReportDCU2 Where PAYCode = '" & Trim(Payor) & "' AND " & _
                           " (INSCode = '' OR INSCode IS NULL) AND (AGEBand = '' OR AGEBand IS NULL) "
                    
                    ElseIf Payor = "" And INSCode <> "" And AgeCode = "" Then
                        strAppendCase = "12"
                        StrAppend1 = " AND (PAYCode = '' OR PAYCode IS NULL) AND " & _
                           " INSCode = '" & Trim(INSCode) & "' AND (AGEBand = '' OR AGEBand IS NULL)"
                        StrAppend1 = StrAppend1 + StrAppend
                    'sql3 = " SELECT PAYCode, INSCode, Ageband, SUM(DurationStay) AS NoofDay " & _
                           " From VIEW_CLMReportDCU2 Where (PAYCode = '' OR PAYCode IS NULL) AND " & _
                           " INSCode = '" & Trim(INSCode) & "' AND (AGEBand = '' OR AGEBand IS NULL) "
                    
                    ElseIf Payor = "" And INSCode = "" And AgeCode <> "" Then
                        strAppendCase = "12"
                        StrAppend1 = " AND (PAYCode = '' OR PAYCode IS NULL) AND " & _
                           " (INSCode = '' OR INSCode IS NULL) AND AGEBand = '" & Trim(AgeCode) & "'"
                        StrAppend1 = StrAppend1 + StrAppend
                    'sql3 = " SELECT PAYCode, INSCode, Ageband, SUM(DurationStay) AS NoofDay " & _
                           " From VIEW_CLMReportDCU2 Where (PAYCode = '' OR PAYCode IS NULL) AND " & _
                           " (INSCode = '' OR INSCode IS NULL) AND AGEBand = '" & Trim(AgeCode) & "'"
                    
                    Else
                        strAppendCase = "12"
                        StrAppend1 = " AND PAYCode = '" & Trim(Payor) & "' AND " & _
                           " INSCode = '" & Trim(INSCode) & "' AND AGEBand = '" & Trim(AgeCode) & "'"
                        StrAppend1 = StrAppend1 + StrAppend
                    'sql3 = " SELECT PAYCode, INSCode, Ageband, SUM(DurationStay) AS NoofDay " & _
                           " From VIEW_CLMReportDCU2 Where PAYCode = '" & Trim(Payor) & "' AND " & _
                           " INSCode = '" & Trim(INSCode) & "' AND AGEBand = '" & Trim(AgeCode) & "'"
                    
                    End If
                       
                       'sql2 = " Group By PAYCode, INSCode, Ageband "
                       
                       'If Len(Trim(StrAppend)) Then
                       '     sql3 = sql3 + StrAppend + sql2
                       ' Else
                       '     sql3 = sql3 + sql2
                       ' End If
                                              
                       'rst1.Open sql3, medixmasterconnWS, adOpenForwardOnly, adLockBatchOptimistic
                        Call DurationStay(StrAppend1, strAppendCase, Record)
                End If
                                
                If Len(rst2!TotalActualAmt) Then
                   Record.SubItems(5) = Format(rst2!TotalActualAmt, "#,##0.00")
                End If
                
                If Len(rst2!TotalApprovedAmt) Then
                   Record.SubItems(6) = Format(rst2!TotalApprovedAmt, "#,##0.00")
                End If
                
            'End With
            
            lblCurrent = lblCurrent + 1
            
            rst2.MoveNext
       Loop
    'End If
    If ListView2.ListItems.Count = 0 Then
        MsgBox "No record found!", vbOKOnly + vbCritical, DialogBoxTitle
        CmdPrint.Enabled = True
        CmdClear.Enabled = True
        cmdExit.Enabled = True
    End If
    rst2.Close
    Set rst2 = Nothing
    CmdPrint.Enabled = True
    CmdClear.Enabled = True
    cmdExit.Enabled = True
End Sub
'********************** List Ageband By Insured Type Results *******************************

'********************** List Ageband By Race Type Results *******************************
Private Sub AgeRaceListing(Optional StrAppend As String)
Dim rst2 As New ADODB.Recordset
Dim cmd As New ADODB.Command
Dim prm1 As ADODB.Parameter
Dim prm2 As ADODB.Parameter
Dim strAppendCase As String
Dim StrAppend1 As String
Dim Record As ListItem
    
    lblCurrent = 0
    ListView2.ListItems.Clear
    RaceCode = ""
    AgeCode = ""
                
    If cboClaimCase.ListIndex = 0 Then
        strAppendCase = "43"
    'sql = " SELECT RACCode, Ageband, COUNT(CASNumber) AS TotalCases, " & _
          " SUM(CLMTotalActual) AS TotalActualAmt, " & _
          " SUM(CLMTotalApproved) AS TotalApprovedAmt, " & _
          " SUM(NoofDay) AS NoofDay From VIEW_CLMReportDCU Where 0 = 0 "
          
    ElseIf cboClaimCase.ListIndex = 1 Then
        strAppendCase = "44"
    'sql = " SELECT RACCode, Ageband, COUNT(DISTINCT CASNumber) AS TotalCases, " & _
          " SUM(CLMTotalActual) AS TotalActualAmt, " & _
          " SUM(CLMTotalApproved) AS TotalApprovedAmt " & _
          " From VIEW_CLMReportDCU Where 0 = 0 "
          
    End If
              
    'sql2 = " GROUP BY RACCode, Ageband ORDER BY RACCode, Ageband "
    
    'If Len(Trim(StrAppend)) Then
    '    sql = sql + StrAppend + sql2
    'Else
    '    sql = sql + sql2
    'End If
    
    'rst2.Open sql, medixmasterconnWS, adOpenForwardOnly, adLockBatchOptimistic
    Set cmd.ActiveConnection = medixmasterconnWS
        cmd.CommandText = "DCUReport"
        cmd.CommandType = adCmdStoredProc
        cmd.CommandTimeout = 300
        Set prm1 = cmd.CreateParameter("StrAppend", adVarChar, adParamInput, 2000, StrAppend)
        cmd.Parameters.Append prm1
        Set prm2 = cmd.CreateParameter("StrAppendCase", adVarChar, adParamInput, 2000, strAppendCase)
        cmd.Parameters.Append prm2
        rst2.CursorLocation = adUseClient
        rst2.CursorType = adOpenForwardOnly
        rst2.CacheSize = 100
        Set rst2 = cmd.Execute
        
        Do Until rst2.EOF
        
                If Len(rst2!RACCode) Then
                    Set Record = ListView2.ListItems.Add(, , rst2!RACCode)
                    RaceCode = rst2!RACCode
                Else
                    Set Record = ListView2.ListItems.Add(, , "Empty")
                    RaceCode = "'"
                End If
                                              
                If Len(rst2!AgeBand) Then
                   Record.SubItems(1) = rst2!AgeBand
                   AgeCode = rst2!AgeBand
                Else
                   Record.SubItems(1) = "Empty"
                   AgeCode = ""
                End If
                    
                If Len(rst2!TotalCases) Then
                   Record.SubItems(2) = rst2!TotalCases
                End If
                    
                If cboClaimCase.ListIndex = 0 Then
                
                    If rst2!NoofDay = "0" Then
                        Record.SubItems(3) = "1"
                    ElseIf Len(rst2!NoofDay) Then
                        Record.SubItems(3) = rst2!NoofDay
                    End If
                
                ElseIf cboClaimCase.ListIndex = 1 Then
                
                    If RaceCode = "" And AgeCode = "" Then
                        strAppendCase = "45"
                        StrAppend1 = " AND (RACCode = '' OR RACCode IS NULL) AND " & _
                           " (AGEBand = '' or AGEBand IS NULL) "
                        StrAppend1 = StrAppend1 + StrAppend
                    'sql3 = " SELECT RACCode, Ageband, SUM(DurationStay) AS NoofDay " & _
                           " From VIEW_CLMReportDCU2 Where (RACCode = '' OR RACCode IS NULL) AND " & _
                           " (AGEBand = '' or AGEBand IS NULL) "
                           
                    ElseIf RaceCode <> "" And AgeCode = "" Then
                        strAppendCase = "45"
                        StrAppend1 = " AND RACCode = '" & Trim(RaceCode) & "' AND " & _
                           " (AGEBand = '' or AGEBand IS NULL) "
                        StrAppend1 = StrAppend1 + StrAppend
                    'sql3 = " SELECT RACCode, Ageband, SUM(DurationStay) AS NoofDay " & _
                           " From VIEW_CLMReportDCU2 Where RACCode = '" & Trim(RaceCode) & "' AND " & _
                           " (AGEBand = '' or AGEBand IS NULL) "
                          
                    ElseIf RaceCode = "" And AgeCode <> "" Then
                        strAppendCase = "45"
                        StrAppend1 = " AND (RACCode = '' OR RACCode IS NULL) AND " & _
                           " AGEBand = '" & Trim(AgeCode) & "'"
                        StrAppend1 = StrAppend1 + StrAppend
                    'sql3 = " SELECT RACCode, Ageband, SUM(DurationStay) AS NoofDay " & _
                           " From VIEW_CLMReportDCU2 Where (RACCode = '' OR RACCode IS NULL) AND " & _
                           " AGEBand = '" & Trim(AgeCode) & "'"
                           
                    Else
                        strAppendCase = "45"
                        StrAppend1 = " AND RACCode = '" & Trim(RaceCode) & "' AND " & _
                           " AGEBand = '" & Trim(AgeCode) & "'"
                        StrAppend1 = StrAppend1 + StrAppend
                    'sql3 = " SELECT RACCode, Ageband, SUM(DurationStay) AS NoofDay " & _
                           " From VIEW_CLMReportDCU2 Where RACCode = '" & Trim(RaceCode) & "' AND " & _
                           " AGEBand = '" & Trim(AgeCode) & "'"
                           
                    End If
                       
                       'sql2 = " Group By RACCode, Ageband "
                       
                       'If Len(Trim(StrAppend)) Then
                       '     sql3 = sql3 + StrAppend + sql2
                       ' Else
                       '     sql3 = sql3 + sql2
                       ' End If
                                              
                       'rst1.Open sql3, medixmasterconnWS, adOpenForwardOnly, adLockBatchOptimistic
                        Call DurationStay(StrAppend1, strAppendCase, Record)
                End If
                
                If Len(rst2!TotalActualAmt) Then
                   Record.SubItems(4) = Format(rst2!TotalActualAmt, "#,##0.00")
                End If
                
                If Len(rst2!TotalApprovedAmt) Then
                   Record.SubItems(5) = Format(rst2!TotalApprovedAmt, "#,##0.00")
                End If
                
            'End With
            lblCurrent = lblCurrent + 1
            rst2.MoveNext
       Loop
    'End If
    If ListView2.ListItems.Count = 0 Then
        MsgBox "No record found!", vbOKOnly + vbCritical, DialogBoxTitle
        CmdPrint.Enabled = True
        CmdClear.Enabled = True
        cmdExit.Enabled = True
    End If
    rst2.Close
    Set rst2 = Nothing
    CmdPrint.Enabled = True
    CmdClear.Enabled = True
    cmdExit.Enabled = True
End Sub
'********************** List Ageband By Race Type Results *******************************

'********************** List Ageband By Plan Type Results *******************************
Private Sub AgePlanListing(Optional StrAppend As String)
Dim rst2 As New ADODB.Recordset
Dim cmd As New ADODB.Command
Dim prm1 As ADODB.Parameter
Dim prm2 As ADODB.Parameter
Dim strAppendCase As String
Dim StrAppend1 As String
Dim Record As ListItem

    lblCurrent = 0
    INSCode = ""
    plancode = ""
    AgeCode = ""
    ListView2.ListItems.Clear
    
    If cboClaimCase.ListIndex = 0 Then
        strAppendCase = "46"
    'sql = " SELECT INSCode, PLNCode, Ageband, COUNT(CASNumber) AS TotalCases, " & _
          " SUM(CLMTotalActual) AS TotalActualAmt, " & _
          " SUM(CLMTotalApproved) AS TotalApprovedAmt, " & _
          " SUM(NoofDay) AS NoofDay From VIEW_CLMReportDCU Where 0 = 0 "
          
    ElseIf cboClaimCase.ListIndex = 1 Then
        strAppendCase = "47"
    'sql = " SELECT INSCode, PLNCode, Ageband, COUNT(DISTINCT CASNumber) AS TotalCases, " & _
          " SUM(CLMTotalActual) AS TotalActualAmt, COUNT(CASNumber) AS TotalCases2, " & _
          " SUM(CLMTotalApproved) AS TotalApprovedAmt, " & _
          " SUM(NoofDay) AS NoofDay From VIEW_CLMReportDCU Where 0 = 0 "
          
    End If
          
    'sql2 = " GROUP BY INSCode, PLNCode, Ageband ORDER BY INSCode "
    
    'If Len(Trim(StrAppend)) Then
    '    sql = sql + StrAppend + sql2
    'Else
    '    sql = sql + sql2
    'End If
   
    'rst2.Open sql, medixmasterconnWS, adOpenForwardOnly, adLockBatchOptimistic
    Set cmd.ActiveConnection = medixmasterconnWS
        cmd.CommandText = "DCUReport"
        cmd.CommandType = adCmdStoredProc
        cmd.CommandTimeout = 300
        Set prm1 = cmd.CreateParameter("StrAppend", adVarChar, adParamInput, 2000, StrAppend)
        cmd.Parameters.Append prm1
        Set prm2 = cmd.CreateParameter("StrAppendCase", adVarChar, adParamInput, 2000, strAppendCase)
        cmd.Parameters.Append prm2
        rst2.CursorLocation = adUseClient
        rst2.CursorType = adOpenForwardOnly
        rst2.CacheSize = 100
        Set rst2 = cmd.Execute
        
        Do Until rst2.EOF
        
                If Len(rst2!INSCode) Then
                    Set Record = ListView2.ListItems.Add(, , rst2!INSCode)
                    INSCode = rst2!INSCode
                Else
                    Set Record = ListView2.ListItems.Add(, , "Empty")
                    INSCode = ""
                End If
                                              
                If Len(rst2!PLNCode) Then
                   Record.SubItems(1) = rst2!PLNCode
                   plancode = rst2!PLNCode
                Else
                    plancode = ""
                End If
                    
                If Len(rst2!AgeBand) Then
                   Record.SubItems(2) = rst2!AgeBand
                   AgeCode = rst2!AgeBand
                Else
                    AgeCode = ""
                End If
                    
                If Len(rst2!TotalCases) Then
                   Record.SubItems(3) = rst2!TotalCases
                End If
                
                If cboClaimCase.ListIndex = 0 Then
                
                    If rst2!NoofDay = "0" Then
                        Record.SubItems(4) = "1"
                    ElseIf Len(rst2!NoofDay) Then
                        Record.SubItems(4) = rst2!NoofDay
                    End If
                
                ElseIf cboClaimCase.ListIndex = 1 Then
                
                    If INSCode = "" And plancode = "" And AgeCode = "" Then
                        strAppendCase = "48"
                        StrAppend1 = " AND (INSCode = '' OR INSCode IS NULL) AND " & _
                           " (PLNCode = '' OR PLNCode IS NULL) AND (AGEBand = '' OR AGEBand IS NULL)"
                        StrAppend1 = StrAppend1 + StrAppend
                    'sql3 = " SELECT INSCode, PLNCode, Ageband, SUM(DurationStay) AS NoofDay " & _
                           " From VIEW_CLMReportDCU2 WHERE (INSCode = '' OR INSCode IS NULL) AND " & _
                           " (PLNCode = '' OR PLNCode IS NULL) AND (AGEBand = '' OR AGEBand IS NULL) "
                           
                    ElseIf INSCode <> "" And plancode = "" And AgeCode = "" Then
                        strAppendCase = "48"
                        StrAppend1 = " AND INSCode = '" & Trim(INSCode) & "' AND " & _
                           " (PLNCode = '' OR PLNCode IS NULL) AND (AGEBand = '' OR AGEBand IS NULL)"
                        StrAppend1 = StrAppend1 + StrAppend
                    'sql3 = " SELECT INSCode, PLNCode, Ageband, SUM(DurationStay) AS NoofDay " & _
                           " From VIEW_CLMReportDCU2 WHERE INSCode = '" & Trim(INSCode) & "' AND " & _
                           " (PLNCode = '' OR PLNCode IS NULL) AND (AGEBand = '' OR AGEBand IS NULL) "
                           
                    ElseIf INSCode = "" And plancode <> "" And AgeCode = "" Then
                        strAppendCase = "48"
                        StrAppend1 = " AND (INSCode = '' OR INSCode IS NULL) AND " & _
                           " PLNCode = '" & Trim(plancode) & "' AND (AGEBand = '' OR AGEBand IS NULL)"
                        StrAppend1 = StrAppend1 + StrAppend
                    'sql3 = " SELECT INSCode, PLNCode, Ageband, SUM(DurationStay) AS NoofDay " & _
                           " From VIEW_CLMReportDCU2 WHERE (INSCode = '' OR INSCode IS NULL) AND " & _
                           " PLNCode = '" & Trim(plancode) & "' AND (AGEBand = '' OR AGEBand IS NULL) "
                           
                    ElseIf INSCode = "" And plancode = "" And AgeCode <> "" Then
                        strAppendCase = "48"
                        StrAppend1 = " AND (INSCode = '' OR INSCode IS NULL) AND " & _
                           " (PLNCode = '' OR PLNCode IS NULL) AND AGEBand = '" & Trim(AgeCode) & "'"
                        StrAppend1 = StrAppend1 + StrAppend
                    'sql3 = " SELECT INSCode, PLNCode, Ageband, SUM(DurationStay) AS NoofDay " & _
                           " From VIEW_CLMReportDCU2 WHERE (INSCode = '' OR INSCode IS NULL) AND " & _
                           " (PLNCode = '' OR PLNCode IS NULL) AND AGEBand = '" & Trim(AgeCode) & "'"
                           
                    Else
                        strAppendCase = "48"
                        StrAppend1 = " AND INSCode = '" & Trim(INSCode) & "' AND " & _
                           " PLNCode = '" & Trim(plancode) & "' AND AGEBand = '" & Trim(AgeCode) & "'"
                        StrAppend1 = StrAppend1 + StrAppend
                    'sql3 = " SELECT INSCode, PLNCode, Ageband, SUM(DurationStay) AS NoofDay " & _
                           " From VIEW_CLMReportDCU2 WHERE INSCode = '" & Trim(INSCode) & "' AND " & _
                           " PLNCode = '" & Trim(plancode) & "' AND AGEBand = '" & Trim(AgeCode) & "'"
                           
                    End If
                       
                       'sql2 = " Group By INSCode, PLNCode, Ageband "
                      
                       'If Len(Trim(StrAppend)) Then
                       '     sql3 = sql3 + StrAppend + sql2
                       'Else
                       '     sql3 = sql3 + sql2
                       'End If
                                              
                       'rst1.Open sql3, medixmasterconnWS, adOpenForwardOnly, adLockBatchOptimistic
                        Call DurationStay(StrAppend1, strAppendCase, Record)
                End If
                                
                If Len(rst2!TotalActualAmt) Then
                   Record.SubItems(5) = Format(rst2!TotalActualAmt, "#,##0.00")
                End If
                
                If Len(rst2!TotalApprovedAmt) Then
                   Record.SubItems(6) = Format(rst2!TotalApprovedAmt, "#,##0.00")
                End If
                        
            lblCurrent = lblCurrent + 1
            rst2.MoveNext
       Loop
   
    If ListView2.ListItems.Count = 0 Then
        MsgBox "No record found!", vbOKOnly + vbCritical, DialogBoxTitle
        CmdPrint.Enabled = True
        CmdClear.Enabled = True
        cmdExit.Enabled = True
    End If
    rst2.Close
    Set rst2 = Nothing
    CmdPrint.Enabled = True
    CmdClear.Enabled = True
    cmdExit.Enabled = True
End Sub
'********************** List Ageband By Plan Type Results *******************************

'********************** List Ageband By Illness Results *******************************
Private Sub AgeIllnessListing(Optional StrAppend As String)
Dim rst2 As New ADODB.Recordset
Dim cmd As New ADODB.Command
Dim prm1 As ADODB.Parameter
Dim prm2 As ADODB.Parameter
Dim strAppendCase As String
Dim StrAppend1 As String
Dim Record As ListItem

    lblCurrent = 0
    AgeCode = ""
    ClaimNatureCode = ""
    ListView2.ListItems.Clear
            
    If cboClaimCase.ListIndex = 0 Then
        strAppendCase = "49"
    'sql = " SELECT ClaimNatureCode, Ageband, COUNT(CASNumber) AS TotalCases, " & _
          " SUM(CLMTotalActual) AS TotalActualAmt, " & _
          " SUM(CLMTotalApproved) AS TotalApprovedAmt, " & _
          " SUM(NoofDay) AS NoofDay From VIEW_CLMReportDCU Where 0 = 0 "
          
    ElseIf cboClaimCase.ListIndex = 1 Then
        strAppendCase = "50"
    'sql = " SELECT ClaimNatureCode, Ageband, COUNT(DISTINCT CASNumber) AS TotalCases, " & _
          " SUM(CLMTotalActual) AS TotalActualAmt, " & _
          " SUM(CLMTotalApproved) AS TotalApprovedAmt " & _
          " From VIEW_CLMReportDCU Where 0 = 0 "
          
    End If
          
    'sql2 = " GROUP BY ClaimNatureCode, Ageband ORDER BY ClaimNatureCode "
    
    'If Len(Trim(StrAppend)) Then
    '    sql = sql + StrAppend + sql2
    'Else
    '    sql = sql + sql2
    'End If
    
    'rst2.Open sql, medixmasterconnWS, adOpenForwardOnly, adLockBatchOptimistic
    Set cmd.ActiveConnection = medixmasterconnWS
        cmd.CommandText = "DCUReport"
        cmd.CommandType = adCmdStoredProc
        cmd.CommandTimeout = 300
        Set prm1 = cmd.CreateParameter("StrAppend", adVarChar, adParamInput, 2000, StrAppend)
        cmd.Parameters.Append prm1
        Set prm2 = cmd.CreateParameter("StrAppendCase", adVarChar, adParamInput, 2000, strAppendCase)
        cmd.Parameters.Append prm2
        rst2.CursorLocation = adUseClient
        rst2.CursorType = adOpenForwardOnly
        rst2.CacheSize = 100
        Set rst2 = cmd.Execute
        
        Do Until rst2.EOF
        
                If Len(rst2!ClaimNatureCode) Then
                    Set Record = ListView2.ListItems.Add(, , rst2!ClaimNatureCode)
                    ClaimNatureCode = rst2!ClaimNatureCode
                Else
                    Set Record = ListView2.ListItems.Add(, , "Empty")
                    ClaimNatureCode = ""
                End If
                                        
                If Len(rst2!AgeBand) Then
                   Record.SubItems(1) = rst2!AgeBand
                   AgeCode = rst2!AgeBand
                Else
                   AgeCode = ""
                End If
                    
                If Len(rst2!TotalCases) Then
                   Record.SubItems(2) = rst2!TotalCases
                End If
                    
                If cboClaimCase.ListIndex = 0 Then
                
                    If rst2!NoofDay = "0" Then
                        Record.SubItems(3) = "1"
                    ElseIf Len(rst2!NoofDay) Then
                        Record.SubItems(3) = rst2!NoofDay
                    End If
                
                ElseIf cboClaimCase.ListIndex = 1 Then
                
                    If ClaimNatureCode = "" And AgeCode = "" Then
                        strAppendCase = "51"
                        StrAppend1 = " AND (claimnaturecode = '' OR ClaimNatureCode IS NULL) AND (ageband = '' OR ageband IS NULL)"
                        StrAppend1 = StrAppend1 + StrAppend
                    'sql3 = " SELECT ClaimNatureCode, Ageband, SUM(DurationStay) AS NoofDay " & _
                           " From VIEW_CLMReportDCU2 Where (claimnaturecode = '' OR ClaimNatureCode IS NULL) AND (ageband = '' OR ageband IS NULL)"
                           
                    ElseIf ClaimNatureCode <> "" And AgeCode = "" Then
                        strAppendCase = "51"
                        StrAppend1 = " AND (claimnaturecode = '' OR ClaimNatureCode IS NULL) AND (ageband = '' OR ageband IS NULL)"
                        StrAppend1 = StrAppend1 + StrAppend
                    'sql3 = " SELECT ClaimNatureCode, Ageband, SUM(DurationStay) AS NoofDay " & _
                           " From VIEW_CLMReportDCU2 Where ClaimNatureCode = '" & Trim(ClaimNatureCode) & "' AND (ageband = '' OR Ageband IS NULL)"
                    
                    ElseIf ClaimNatureCode = "" And AgeCode <> "" Then
                        strAppendCase = "51"
                        StrAppend1 = " AND (ClaimNatureCode = '' OR ClaimNatureCode IS NULL) AND ageband = '" & Trim(AgeCode) & "'"
                        StrAppend1 = StrAppend1 + StrAppend
                    'sql3 = " SELECT ClaimNatureCode, Ageband, SUM(DurationStay) AS NoofDay " & _
                           " From VIEW_CLMReportDCU2 Where (ClaimNatureCode = '' OR ClaimNatureCode IS NULL) AND ageband = '" & Trim(AgeCode) & "'"
                    
                    Else
                        strAppendCase = "51"
                        StrAppend1 = " AND claimnaturecode = '" & Trim(ClaimNatureCode) & "' AND ageband = '" & Trim(AgeCode) & "'"
                        StrAppend1 = StrAppend1 + StrAppend
                    'sql3 = " SELECT ClaimNatureCode, Ageband, SUM(DurationStay) AS NoofDay " & _
                           " From VIEW_CLMReportDCU2 Where claimnaturecode = '" & Trim(ClaimNatureCode) & "' AND ageband = '" & Trim(AgeCode) & "'"
                    
                    End If
                       
                       'sql2 = " Group By ClaimNatureCode, Ageband "
                       
                       'If Len(Trim(StrAppend)) Then
                       '     sql3 = sql3 + StrAppend + sql2
                       ' Else
                       '     sql3 = sql3 + sql2
                       ' End If
                                              
                       'rst1.Open sql3, medixmasterconnWS, adOpenForwardOnly, adLockBatchOptimistic
                        Call DurationStay(StrAppend1, strAppendCase, Record)
                End If
                                    
                If Len(rst2!TotalActualAmt) Then
                   Record.SubItems(4) = Format(rst2!TotalActualAmt, "#,##0.00")
                End If
                
                If Len(rst2!TotalApprovedAmt) Then
                   Record.SubItems(5) = Format(rst2!TotalApprovedAmt, "#,##0.00")
                End If
               
            lblCurrent = lblCurrent + 1
            rst2.MoveNext
       Loop
       
    If ListView2.ListItems.Count = 0 Then
     MsgBox "No record found!", vbOKOnly + vbCritical, DialogBoxTitle
     CmdPrint.Enabled = True
     CmdClear.Enabled = True
     cmdExit.Enabled = True
    End If
   
    rst2.Close
    Set rst2 = Nothing
    CmdPrint.Enabled = True
    CmdClear.Enabled = True
    cmdExit.Enabled = True
End Sub
'********************** List Ageband By Illness Results *******************************

'********************** List Gender By Insured Type Results *******************************
Private Sub SexInsuredListing(Optional StrAppend As String)
Dim rst2 As New ADODB.Recordset
Dim cmd As New ADODB.Command
Dim prm1 As ADODB.Parameter
Dim prm2 As ADODB.Parameter
Dim strAppendCase As String
Dim StrAppend1 As String
Dim Record As ListItem
    
    lblCurrent = 0
    Payor = ""
    INSCode = ""
    SexCode = ""
    ListView2.ListItems.Clear
            
    If cboClaimCase.ListIndex = 0 Then
        strAppendCase = "13"
    'sql = " SELECT PAYCode, INSCode, Sex, COUNT(CASNumber) AS TotalCases, SUM(CLMTotalActual) " & _
          " AS totalActual, SUM(CLMTotalApproved) AS TotalApproved, " & _
          " SUM(NoofDay) AS NoofDay From VIEW_CLMReportDCU Where 0 = 0 "
          
    ElseIf cboClaimCase.ListIndex = 1 Then
        strAppendCase = "14"
    'sql = " SELECT PAYCode, INSCode, Sex, COUNT(DISTINCT CASNumber) AS TotalCases, SUM(CLMTotalActual) " & _
          " AS totalActual, SUM(CLMTotalApproved) AS TotalApproved " & _
          " From VIEW_CLMReportDCU Where 0 = 0 "
    
    End If
    
    'sql2 = " Group By PAYCode, INSCode, Sex "
    
    'If Len(Trim(StrAppend)) Then
    '    sql = sql + StrAppend + sql2
    'Else
    '    sql = sql + sql2
    'End If
      
    'rst2.Open sql, medixmasterconnWS, adOpenForwardOnly, adLockBatchOptimistic
    Set cmd.ActiveConnection = medixmasterconnWS
        cmd.CommandText = "DCUReport"
        cmd.CommandType = adCmdStoredProc
        cmd.CommandTimeout = 300
        Set prm1 = cmd.CreateParameter("StrAppend", adVarChar, adParamInput, 2000, StrAppend)
        cmd.Parameters.Append prm1
        Set prm2 = cmd.CreateParameter("StrAppendCase", adVarChar, adParamInput, 2000, strAppendCase)
        cmd.Parameters.Append prm2
        rst2.CursorLocation = adUseClient
        rst2.CursorType = adOpenForwardOnly
        rst2.CacheSize = 100
        Set rst2 = cmd.Execute
        
        Do Until rst2.EOF
        
            With rst2
                If Len(rst2!PAYCode) Then
                    Set Record = ListView2.ListItems.Add(, , rst2!PAYCode)
                    Payor = rst2!PAYCode
                Else
                    Set Record = ListView2.ListItems.Add(, , "Empty")
                    Payor = ""
                End If
                
                If Len(rst2!INSCode) Then
                    Record.SubItems(1) = rst2!INSCode
                    INSCode = rst2!INSCode
                Else
                    Record.SubItems(1) = "Empty"
                    INSCode = ""
                End If
                
                If Len(rst2!Sex) Then
                    Record.SubItems(2) = rst2!Sex
                    SexCode = rst2!Sex
                Else
                    Record.SubItems(2) = "Empty"
                    SexCode = ""
                End If

                If Len(rst2!TotalCases) Then
                    Record.SubItems(3) = rst2!TotalCases
                End If
                    
                If cboClaimCase.ListIndex = 0 Then
                
                    If rst2!NoofDay = "0" Then
                        Record.SubItems(4) = "1"
                    ElseIf Len(rst2!NoofDay) Then
                        Record.SubItems(4) = rst2!NoofDay
                    End If
                
                ElseIf cboClaimCase.ListIndex = 1 Then
                
                    If Payor = "" And INSCode = "" And SexCode = "" Then
                        strAppendCase = "15"
                        StrAppend1 = " AND (PAYCode = '' OR PAYCode IS NULL) AND " & _
                           " (INSCode = '' OR INSCOde IS NULL) AND (Sex = '' OR Sex IS NULL) "
                        StrAppend1 = StrAppend1 + StrAppend
                    'sql3 = " SELECT PAYCode, INSCode, Sex, SUM(DurationStay) AS NoofDay " & _
                           " From VIEW_CLMReportDCU2 Where (PAYCode = '' OR PAYCode IS NULL) AND " & _
                           " (INSCode = '' OR INSCOde IS NULL) AND (Sex = '' OR Sex IS NULL) "
                           
                    ElseIf Payor <> "" And INSCode = "" And SexCode = "" Then
                        strAppendCase = "15"
                        StrAppend1 = " AND PAYCode = '" & Trim(Payor) & "' AND " & _
                           " (INSCode = '' OR INSCOde IS NULL) AND (Sex = '' OR Sex IS NULL) "
                        StrAppend1 = StrAppend1 + StrAppend
                    'sql3 = " SELECT PAYCode, INSCode, Sex, SUM(DurationStay) AS NoofDay " & _
                           " From VIEW_CLMReportDCU2 Where PAYCode = '" & Trim(Payor) & "' AND " & _
                           " (INSCode = '' OR INSCOde IS NULL) AND (Sex = '' OR Sex IS NULL) "
                           
                    ElseIf Payor = "" And INSCode <> "" And SexCode = "" Then
                        strAppendCase = "15"
                        StrAppend1 = " AND (PAYCode = '' OR PAYCode IS NULL) AND " & _
                           " INSCode = '" & Trim(INSCode) & "' AND (Sex = '' OR Sex IS NULL) "
                        StrAppend1 = StrAppend1 + StrAppend
                    'sql3 = " SELECT PAYCode, INSCode, Sex, SUM(DurationStay) AS NoofDay " & _
                           " From VIEW_CLMReportDCU2 Where (PAYCode = '' OR PAYCode IS NULL) AND " & _
                           " INSCode = '" & Trim(INSCode) & "' AND (Sex = '' OR Sex IS NULL) "
                           
                    ElseIf Payor = "" And INSCode = "" And SexCode <> "" Then
                        strAppendCase = "15"
                        StrAppend1 = " AND (PAYCode = '' OR PAYCode IS NULL) AND " & _
                           " (INSCode = '' OR INSCOde IS NULL) AND Sex = '" & Trim(SexCode) & "'"
                        StrAppend1 = StrAppend1 + StrAppend
                    'sql3 = " SELECT PAYCode, INSCode, Sex, SUM(DurationStay) AS NoofDay " & _
                           " From VIEW_CLMReportDCU2 Where (PAYCode = '' OR PAYCode IS NULL) AND " & _
                           " (INSCode = '' OR INSCOde IS NULL) AND Sex = '" & Trim(SexCode) & "'"
                           
                    Else
                        strAppendCase = "15"
                        StrAppend1 = " AND PAYCode = '" & Trim(Payor) & "' AND " & _
                           " INSCode = '" & Trim(INSCode) & "' AND Sex = '" & Trim(SexCode) & "'"
                        StrAppend1 = StrAppend1 + StrAppend
                    'sql3 = " SELECT PAYCode, INSCode, Sex, SUM(DurationStay) AS NoofDay " & _
                           " From VIEW_CLMReportDCU2 Where PAYCode = '" & Trim(Payor) & "' AND " & _
                           " INSCode = '" & Trim(INSCode) & "' AND Sex = '" & Trim(SexCode) & "'"
                           
                    End If
                       
                       'sql2 = " Group By PAYCode, INSCode, Sex "
                       
                       'If Len(Trim(StrAppend)) Then
                       '     sql3 = sql3 + StrAppend + sql2
                       ' Else
                       '     sql3 = sql3 + sql2
                       ' End If
                                              
                       'rst1.Open sql3, medixmasterconnWS, adOpenForwardOnly, adLockBatchOptimistic
                        Call DurationStay(StrAppend1, strAppendCase, Record)
                End If
                    
                If Len(rst2!TotalActual) Then
                    Record.SubItems(5) = Format(rst2!TotalActual, "#,##0.00")
                End If
                    
                If Len(rst2!TotalApproved) Then
                    Record.SubItems(6) = Format(rst2!TotalApproved, "#,##0.00")
                End If
                
            End With
            lblCurrent = lblCurrent + 1
            rst2.MoveNext
       Loop
    If ListView2.ListItems.Count = 0 Then
        MsgBox "No record found!", vbOKOnly + vbCritical, DialogBoxTitle
        CmdPrint.Enabled = True
        CmdClear.Enabled = True
        cmdExit.Enabled = True
    End If
    rst2.Close
    Set rst2 = Nothing
    CmdPrint.Enabled = True
    CmdClear.Enabled = True
    cmdExit.Enabled = True
End Sub
'********************** List Gender By Insured Type Results *******************************

'********************** List Race By Insured Type Results *******************************
Private Sub RaceInsuredListing(Optional StrAppend As String)
Dim rst2 As New ADODB.Recordset
Dim cmd As New ADODB.Command
Dim prm1 As ADODB.Parameter
Dim prm2 As ADODB.Parameter
Dim strAppendCase As String
Dim StrAppend1 As String
Dim Record As ListItem

    lblCurrent = 0
    Payor = ""
    INSCode = ""
    RaceCode = ""
    ListView2.ListItems.Clear
            
    If cboClaimCase.ListIndex = 0 Then
        strAppendCase = "16"
    'sql = " SELECT PAYCode, INSCode, RACCode, COUNT(CASNumber) AS TotalCases, SUM(CLMTotalActual) " & _
          " AS totalActual, SUM(CLMTotalApproved) AS TotalApproved, " & _
          " SUM(NoofDay) AS NoofDay From VIEW_CLMReportDCU Where 0 = 0 "
    
    ElseIf cboClaimCase.ListIndex = 1 Then
        strAppendCase = "17"
    'sql = " SELECT PAYCode, INSCode, RACCode, COUNT(DISTINCT CASNumber) AS TotalCases, SUM(CLMTotalActual) " & _
          " AS totalActual, SUM(CLMTotalApproved) AS TotalApproved " & _
          " From VIEW_CLMReportDCU Where 0 = 0 "
    
    End If
    
    'sql2 = " Group By PAYCode, INSCode, RACCode "
    
    'If Len(Trim(StrAppend)) Then
    '    sql = sql + StrAppend + sql2
    'Else
    '    sql = sql + sql2
    'End If
    
    'rst2.Open sql, medixmasterconnWS, adOpenForwardOnly, adLockBatchOptimistic
    Set cmd.ActiveConnection = medixmasterconnWS
        cmd.CommandText = "DCUReport"
        cmd.CommandType = adCmdStoredProc
        cmd.CommandTimeout = 300
        Set prm1 = cmd.CreateParameter("StrAppend", adVarChar, adParamInput, 2000, StrAppend)
        cmd.Parameters.Append prm1
        Set prm2 = cmd.CreateParameter("StrAppendCase", adVarChar, adParamInput, 2000, strAppendCase)
        cmd.Parameters.Append prm2
        rst2.CursorLocation = adUseClient
        rst2.CursorType = adOpenForwardOnly
        rst2.CacheSize = 100
        Set rst2 = cmd.Execute
               
        Do Until rst2.EOF
        
            With rst2
                If Len(rst2!PAYCode) Then
                    Set Record = ListView2.ListItems.Add(, , rst2!PAYCode)
                    Payor = rst2!PAYCode
                Else
                    Set Record = ListView2.ListItems.Add(, , "Empty")
                    Payor = ""
                End If
                
                If Len(rst2!INSCode) Then
                    Record.SubItems(1) = rst2!INSCode
                    INSCode = rst2!INSCode
                Else
                    Record.SubItems(1) = "Empty"
                    INSCode = ""
                End If
                
                If Len(rst2!RACCode) Then
                    Record.SubItems(2) = rst2!RACCode
                    RaceCode = rst2!RACCode
                Else
                    Record.SubItems(2) = "Empty"
                    RaceCode = ""
                End If
                
                If Len(rst2!TotalCases) Then
                    Record.SubItems(3) = rst2!TotalCases
                End If
                    
                If cboClaimCase.ListIndex = 0 Then
                
                    If rst2!NoofDay = "0" Then
                        Record.SubItems(4) = "1"
                    ElseIf Len(rst2!NoofDay) Then
                        Record.SubItems(4) = rst2!NoofDay
                    End If
                
                ElseIf cboClaimCase.ListIndex = 1 Then
                
                    If Payor = "" And INSCode = "" And RaceCode = "" Then
                        strAppendCase = "18"
                        StrAppend1 = " AND (PAYCode = '' OR PAYCode IS NULL) AND " & _
                           " (INSCode = '' OR INSCOde IS NULL) AND (RACCode = '' OR RACCode IS NULL) "
                        StrAppend1 = StrAppend1 + StrAppend
                    'sql3 = " SELECT PAYCode, INSCode, RACCode, SUM(DurationStay) AS NoofDay " & _
                           " From VIEW_CLMReportDCU2 Where (PAYCode = '' OR PAYCode IS NULL) AND " & _
                           " (INSCode = '' OR INSCOde IS NULL) AND (RACCode = '' OR RACCode IS NULL) "
                           
                    ElseIf Payor <> "" And INSCode = "" And RaceCode = "" Then
                        strAppendCase = "18"
                        StrAppend1 = " AND PAYCode = '" & Trim(Payor) & "' AND " & _
                           " (INSCode = '' OR INSCOde IS NULL) AND (RACCode = '' OR RACCode IS NULL) "
                        StrAppend1 = StrAppend1 + StrAppend
                    'sql3 = " SELECT PAYCode, INSCode, RACCode, SUM(DurationStay) AS NoofDay " & _
                           " From VIEW_CLMReportDCU2 Where PAYCode = '" & Trim(Payor) & "' AND " & _
                           " (INSCode = '' OR INSCOde IS NULL) AND (RACCode = '' OR RACCode IS NULL) "
                           
                    ElseIf Payor = "" And INSCode <> "" And RaceCode = "" Then
                        strAppendCase = "18"
                        StrAppend1 = " AND (PAYCode = '' OR PAYCode IS NULL) AND " & _
                           " INSCode = '" & Trim(INSCode) & "' AND (RACCode = '' OR RACCode IS NULL) "
                        StrAppend1 = StrAppend1 + StrAppend
                    'sql3 = " SELECT PAYCode, INSCode, RACCode, SUM(DurationStay) AS NoofDay " & _
                           " From VIEW_CLMReportDCU2 Where (PAYCode = '' OR PAYCode IS NULL) AND " & _
                           " INSCode = '" & Trim(INSCode) & "' AND (RACCode = '' OR RACCode IS NULL) "
                           
                    ElseIf Payor = "" And INSCode = "" And RaceCode <> "" Then
                        strAppendCase = "18"
                        StrAppend1 = " AND (PAYCode = '' OR PAYCode IS NULL) AND " & _
                           " (INSCode = '' OR INSCOde IS NULL) AND RACCode = '" & Trim(RaceCode) & "'"
                        StrAppend1 = StrAppend1 + StrAppend
                    'sql3 = " SELECT PAYCode, INSCode, RACCode, SUM(DurationStay) AS NoofDay " & _
                           " From VIEW_CLMReportDCU2 Where (PAYCode = '' OR PAYCode IS NULL) AND " & _
                           " (INSCode = '' OR INSCOde IS NULL) AND RACCode = '" & Trim(RaceCode) & "'"
                           
                    Else
                        strAppendCase = "18"
                        StrAppend1 = " AND PAYCode = '" & Trim(Payor) & "' AND " & _
                           " INSCode = '" & Trim(INSCode) & "' AND RACCode = '" & Trim(RaceCode) & "'"
                        StrAppend1 = StrAppend1 + StrAppend
                    'sql3 = " SELECT PAYCode, INSCode, RACCode, SUM(DurationStay) AS NoofDay " & _
                           " From VIEW_CLMReportDCU2 Where PAYCode = '" & Trim(Payor) & "' AND " & _
                           " INSCode = '" & Trim(INSCode) & "' AND RACCode = '" & Trim(RaceCode) & "'"
                    
                    End If
                       
                       'sql2 = " Group By PAYCode, INSCode, RACCode  "
                       
                       'If Len(Trim(StrAppend)) Then
                       '     sql3 = sql3 + StrAppend + sql2
                       ' Else
                       '     sql3 = sql3 + sql2
                       ' End If
                                              
                       'rst1.Open sql3, medixmasterconnWS, adOpenForwardOnly, adLockBatchOptimistic
                        Call DurationStay(StrAppend1, strAppendCase, Record)
                End If
                    
                If Len(rst2!TotalActual) Then
                    Record.SubItems(5) = Format(rst2!TotalActual, "#,##0.00")
                End If
                
                If Len(rst2!TotalApproved) Then
                    Record.SubItems(6) = Format(rst2!TotalApproved, "#,##0.00")
                End If
                
            End With
            lblCurrent = lblCurrent + 1
            rst2.MoveNext
       Loop
    If ListView2.ListItems.Count = 0 Then
        MsgBox "No record found!", vbOKOnly + vbCritical, DialogBoxTitle
        CmdPrint.Enabled = True
        CmdClear.Enabled = True
        cmdExit.Enabled = True
    End If
    rst2.Close
    Set rst2 = Nothing
    CmdPrint.Enabled = True
    CmdClear.Enabled = True
    cmdExit.Enabled = True
End Sub
'********************** List Race By Insured Type Results *******************************

'********************** List Plan Summary Results *******************************
Private Sub PlanListing(Optional StrAppend As String)
Dim rst2 As New ADODB.Recordset
Dim cmd As New ADODB.Command
Dim prm1 As ADODB.Parameter
Dim prm2 As ADODB.Parameter
Dim strAppendCase As String
Dim StrAppend1 As String
Dim Record As ListItem

    lblCurrent = 0
    Payor = ""
    plancode = ""
    ListView2.ListItems.Clear
        
    If cboClaimCase.ListIndex = 0 Then
        strAppendCase = "22"
    'sql = " SELECT PAYCode, PLNCode, COUNT(CASNumber) as TotalCases, SUM(clmtotalactual) AS totalActual, " & _
          " SUM(clmtotalapproved) As TotalApproved, " & _
          " SUM(NoofDay) AS NoofDay From VIEW_CLMReportDCU Where 0 = 0 "
        
    ElseIf cboClaimCase.ListIndex = 1 Then
        strAppendCase = "23"
    'sql = " SELECT PAYCode, PLNCode, COUNT(DISTINCT CASNumber) as TotalCases, SUM(clmtotalactual) AS totalActual, " & _
          " SUM(clmtotalapproved) As TotalApproved " & _
          " From VIEW_CLMReportDCU Where 0 = 0 "
    
    End If
    
    'sql2 = " Group By PAYCode, PLNCode "
    
    'If Len(Trim(StrAppend)) Then
    '    sql = sql + StrAppend + sql2
    'Else
    '    sql = sql + sql2
    'End If
    
    'rst2.Open sql, medixmasterconnWS, adOpenForwardOnly, adLockBatchOptimistic
    Set cmd.ActiveConnection = medixmasterconnWS
        cmd.CommandText = "DCUReport"
        cmd.CommandType = adCmdStoredProc
        cmd.CommandTimeout = 300
        Set prm1 = cmd.CreateParameter("StrAppend", adVarChar, adParamInput, 2000, StrAppend)
        cmd.Parameters.Append prm1
        Set prm2 = cmd.CreateParameter("StrAppendCase", adVarChar, adParamInput, 2000, strAppendCase)
        cmd.Parameters.Append prm2
        rst2.CursorLocation = adUseClient
        rst2.CursorType = adOpenForwardOnly
        rst2.CacheSize = 100
        Set rst2 = cmd.Execute
        
        Do Until rst2.EOF
        
            With rst2
                If Len(rst2!PAYCode) Then
                    Set Record = ListView2.ListItems.Add(, , rst2!PAYCode)
                    Payor = rst2!PAYCode
                Else
                    Set Record = ListView2.ListItems.Add(, , "Empty")
                    Payor = ""
                End If
                
                If Len(rst2!PLNCode) Then
                    Record.SubItems(1) = rst2!PLNCode
                    plancode = rst2!PLNCode
                Else
                    Record.SubItems(1) = "Empty"
                    plancode = ""
                End If
                               
                If Len(rst2!TotalCases) Then
                    Record.SubItems(2) = rst2!TotalCases
                End If
                
                If cboClaimCase.ListIndex = 0 Then
                
                    If rst2!NoofDay = "0" Then
                        Record.SubItems(3) = "1"
                    ElseIf Len(rst2!NoofDay) Then
                        Record.SubItems(3) = rst2!NoofDay
                    End If
                
                ElseIf cboClaimCase.ListIndex = 1 Then
                
                    If Payor = "" And plancode = "" Then
                        strAppendCase = "24"
                        StrAppend1 = " AND (PAYCode = '' OR PAYCode IS NULL) AND " & _
                           " (PLNCode = '' OR PLNCode IS NULL) "
                        StrAppend1 = StrAppend1 + StrAppend
                    'sql3 = " SELECT PAYCode, PLNCode, SUM(DurationStay) AS NoofDay " & _
                           " From VIEW_CLMReportDCU2 Where (PAYCode = '' OR PAYCode IS NULL) AND " & _
                           " (PLNCode = '' OR PLNCode IS NULL) "
                           
                    ElseIf Payor <> "" And plancode = "" Then
                        strAppendCase = "24"
                        StrAppend1 = " AND PAYCode = '" & Trim(Payor) & "' AND " & _
                           " (PLNCode = '' OR PLNCode IS NULL) "
                        StrAppend1 = StrAppend1 + StrAppend
                    'sql3 = " SELECT PAYCode, PLNCode, SUM(DurationStay) AS NoofDay " & _
                           " From VIEW_CLMReportDCU2 Where PAYCode = '" & Trim(Payor) & "' AND " & _
                           " (PLNCode = '' OR PLNCode IS NULL) "
                    
                    ElseIf Payor = "" And plancode <> "" Then
                        strAppendCase = "24"
                        StrAppend1 = " AND (PAYCode = '' OR PAYCode IS NULL) AND " & _
                           " PLNCode = '" & Trim(plancode) & "'"
                        StrAppend1 = StrAppend1 + StrAppend
                    'sql3 = " SELECT PAYCode, PLNCode, SUM(DurationStay) AS NoofDay " & _
                           " From VIEW_CLMReportDCU2 Where (PAYCode = '' OR PAYCode IS NULL) AND " & _
                           " PLNCode = '" & Trim(plancode) & "'"
                    
                    Else
                        strAppendCase = "24"
                        StrAppend1 = " AND PAYCode = '" & Trim(Payor) & "' AND " & _
                           " PLNCode = '" & Trim(plancode) & "'"
                        StrAppend1 = StrAppend1 + StrAppend
                    'sql3 = " SELECT PAYCode, PLNCode, SUM(DurationStay) AS NoofDay " & _
                           " From VIEW_CLMReportDCU2 Where PAYCode = '" & Trim(Payor) & "' AND " & _
                           " PLNCode = '" & Trim(plancode) & "'"
                           
                    End If
                       
                       'sql2 = " Group By PAYCode, PLNCode  "
                       
                       'If Len(Trim(StrAppend)) Then
                       '     sql3 = sql3 + StrAppend + sql2
                       ' Else
                       '     sql3 = sql3 + sql2
                       ' End If
                                              
                       'rst1.Open sql3, medixmasterconnWS, adOpenForwardOnly, adLockBatchOptimistic
                        Call DurationStay(StrAppend1, strAppendCase, Record)
                End If
                
                If Len(rst2!TotalActual) Then
                    Record.SubItems(4) = Format(rst2!TotalActual, "#,##0.00")
                End If
                
                If Len(rst2!TotalApproved) Then
                    Record.SubItems(5) = Format(rst2!TotalApproved, "#,##0.00")
                End If
                
            End With
            lblCurrent = lblCurrent + 1
            rst2.MoveNext
       Loop
    If ListView2.ListItems.Count = 0 Then
        MsgBox "No record found!", vbOKOnly + vbCritical, DialogBoxTitle
        CmdPrint.Enabled = True
        CmdClear.Enabled = True
        cmdExit.Enabled = True
    End If
    rst2.Close
    Set rst2 = Nothing
    CmdPrint.Enabled = True
    CmdClear.Enabled = True
    cmdExit.Enabled = True
End Sub

Private Sub PlanPremiumListing(Optional StrAppend As String)
Dim rst2 As New ADODB.Recordset
Dim cmd As New ADODB.Command
Dim prm1 As ADODB.Parameter
Dim prm2 As ADODB.Parameter
Dim strAppendCase As String
Dim StrAppend1 As String
Dim Record As ListItem

    lblCurrent = 0
    Payor = ""
    plancode = ""
    ListView2.ListItems.Clear
        
    If cboClaimCase.ListIndex = 0 Then
        strAppendCase = "19"
    'sql = " SELECT PAYCode, PLNCode, COUNT(CASNumber) as TotalCases, SUM(clmtotalactual) AS totalActual, " & _
          " SUM(clmtotalapproved) As TotalApproved, " & _
          " SUM(NoofDay) AS NoofDay, SUM(MBMBasicPrem) AS BasicPremium, " & _
          " SUM(MBMPremium) AS Premium From VIEW_CLMReportDCU Where 0 = 0 "
          
    ElseIf cboClaimCase.ListIndex = 1 Then
        strAppendCase = "20"
    'sql = " SELECT PAYCode, PLNCode, COUNT(DISTINCT CASNumber) as TotalCases, SUM(clmtotalactual) AS totalActual, " & _
          " SUM(clmtotalapproved) As TotalApproved, " & _
          " SUM(MBMBasicPrem) AS BasicPremium, " & _
          " SUM(MBMPremium) AS Premium From VIEW_CLMReportDCU Where 0 = 0 "
    
    End If
    
    'sql2 = " Group By PAYCode, PLNCode "
   
    'If Len(Trim(StrAppend)) Then
    '    sql = sql + StrAppend + sql2
    'Else
    '    sql = sql + sql2
    'End If
    
    'rst2.Open sql, medixmasterconnWS, adOpenForwardOnly, adLockBatchOptimistic
    Set cmd.ActiveConnection = medixmasterconnWS
        cmd.CommandText = "DCUReport"
        cmd.CommandType = adCmdStoredProc
        cmd.CommandTimeout = 300
        Set prm1 = cmd.CreateParameter("StrAppend", adVarChar, adParamInput, 2000, StrAppend)
        cmd.Parameters.Append prm1
        Set prm2 = cmd.CreateParameter("StrAppendCase", adVarChar, adParamInput, 2000, strAppendCase)
        cmd.Parameters.Append prm2
        rst2.CursorLocation = adUseClient
        rst2.CursorType = adOpenForwardOnly
        rst2.CacheSize = 100
        Set rst2 = cmd.Execute

        Do Until rst2.EOF
        
            With rst2
                If Len(rst2!PAYCode) Then
                    Set Record = ListView2.ListItems.Add(, , rst2!PAYCode)
                    Payor = rst2!PAYCode
                Else
                    Set Record = ListView2.ListItems.Add(, , "Empty")
                    Payor = ""
                End If
                
                If Len(rst2!PLNCode) Then
                    Record.SubItems(1) = rst2!PLNCode
                    plancode = rst2!PLNCode
                Else
                    Record.SubItems(1) = "Empty"
                    plancode = ""
                End If
                               
                If Len(rst2!TotalCases) Then
                    Record.SubItems(2) = rst2!TotalCases
                End If
                
                If cboClaimCase.ListIndex = 0 Then
                
                    If rst2!NoofDay = "0" Then
                        Record.SubItems(3) = "1"
                    ElseIf Len(rst2!NoofDay) Then
                        Record.SubItems(3) = rst2!NoofDay
                    End If
                
                ElseIf cboClaimCase.ListIndex = 1 Then
                
                    If Payor = "" And plancode = "" Then
                        strAppendCase = "21"
                        StrAppend1 = " AND (PAYCode = '' OR PAYCode IS NULL) AND " & _
                           " (PLNCode = '' OR PLNCode IS NULL) "
                        StrAppend1 = StrAppend1 + StrAppend
                    'sql3 = " SELECT PAYCode, PLNCode, SUM(DurationStay) AS NoofDay " & _
                           " From VIEW_CLMReportDCU2 Where (PAYCode = '' OR PAYCode IS NULL) AND " & _
                           " (PLNCode = '' OR PLNCode IS NULL) "
                           
                    ElseIf Payor <> "" And plancode = "" Then
                        strAppendCase = "21"
                        StrAppend1 = " AND PAYCode = '" & Trim(Payor) & "' AND " & _
                           " (PLNCode = '' OR PLNCode IS NULL) "
                        StrAppend1 = StrAppend1 + StrAppend
                    'sql3 = " SELECT PAYCode, PLNCode, SUM(DurationStay) AS NoofDay " & _
                           " From VIEW_CLMReportDCU2 Where PAYCode = '" & Trim(Payor) & "' AND " & _
                           " (PLNCode = '' OR PLNCode IS NULL) "
                           
                    ElseIf Payor = "" And plancode <> "" Then
                        strAppendCase = "21"
                        StrAppend1 = " AND (PAYCode = '' OR PAYCode IS NULL) AND " & _
                           " PLNCode = '" & Trim(plancode) & "'"
                        StrAppend1 = StrAppend1 + StrAppend
                    'sql3 = " SELECT PAYCode, PLNCode, SUM(DurationStay) AS NoofDay " & _
                           " From VIEW_CLMReportDCU2 Where (PAYCode = '' OR PAYCode IS NULL) AND " & _
                           " PLNCode = '" & Trim(plancode) & "'"
                           
                    Else
                        strAppendCase = "21"
                        StrAppend1 = " AND PAYCode = '" & Trim(Payor) & "' AND " & _
                           " PLNCode = '" & Trim(plancode) & "'"
                        StrAppend1 = StrAppend1 + StrAppend
                    'sql3 = " SELECT PAYCode, PLNCode, SUM(DurationStay) AS NoofDay " & _
                           " From VIEW_CLMReportDCU2 Where PAYCode = '" & Trim(Payor) & "' AND " & _
                           " PLNCode = '" & Trim(plancode) & "'"
                           
                    End If
                       
                       'sql2 = " Group By PAYCode, PLNCode "
                       
                       'If Len(Trim(StrAppend)) Then
                       '     sql3 = sql3 + StrAppend + sql2
                       ' Else
                       '     sql3 = sql3 + sql2
                       ' End If
                                              
                       'rst1.Open sql3, medixmasterconnWS, adOpenForwardOnly, adLockBatchOptimistic
                        Call DurationStay(StrAppend1, strAppendCase, Record)
                End If
                
                If Len(rst2!TotalActual) Then
                    Record.SubItems(4) = Format(rst2!TotalActual, "#,##0.00")
                End If
                
                If Len(rst2!TotalApproved) Then
                    Record.SubItems(5) = Format(rst2!TotalApproved, "#,##0.00")
                End If
                
                If Len(rst2!BasicPremium) Then
                    Record.SubItems(6) = Format(rst2!BasicPremium, "#,##0.00")
                End If
                
                If Len(rst2!Premium) Then
                    Record.SubItems(7) = Format(rst2!Premium, "#,##0.00")
                End If
                
            End With
            lblCurrent = lblCurrent + 1
            rst2.MoveNext
       Loop
    If ListView2.ListItems.Count = 0 Then
        MsgBox "No record found!", vbOKOnly + vbCritical, DialogBoxTitle
        CmdPrint.Enabled = True
        CmdClear.Enabled = True
        cmdExit.Enabled = True
    End If
    rst2.Close
    Set rst2 = Nothing
    CmdPrint.Enabled = True
    CmdClear.Enabled = True
    cmdExit.Enabled = True
End Sub
'********************** List Plan Summary Results *******************************

'********************** List Plan Top 10 Hospital Results *******************************
Private Sub PlanTopHosListing(Optional StrAppend As String)
Dim rst2 As New ADODB.Recordset
Dim cmd As New ADODB.Command
Dim prm1 As ADODB.Parameter
Dim prm2 As ADODB.Parameter
Dim strAppendCase As String
Dim StrAppend1 As String
Dim Record As ListItem

    lblCurrent = 0
    ListView2.ListItems.Clear
    Payor = ""
    plancode = ""
    Hospital = ""
    HOspitals = ""
    
    If cboClaimCase.ListIndex = 0 Then
        strAppendCase = "25"
    'sql = " SELECT Top 10 PAYCode, PLNCode, HOSName, COUNT(CASNumber) as TotalCases, SUM(clmtotalactual) AS totalActual, " & _
          " SUM(clmtotalapproved) As TotalApproved, " & _
          " SUM(NoofDay) AS NoofDay From VIEW_CLMReportDCU Where 0 = 0 "
          
    ElseIf cboClaimCase.ListIndex = 1 Then
        strAppendCase = "26"
    'sql = " SELECT Top 10 PAYCode, PLNCode, HOSName, COUNT(DISTINCT CASNumber) as TotalCases, SUM(clmtotalactual) AS totalActual, " & _
          " SUM(clmtotalapproved) As TotalApproved " & _
          " From VIEW_CLMReportDCU Where 0 = 0 "
    
    End If
    
    'sql2 = " Group By PAYCode, PLNCode, HOSName ORDER BY TotalCases Desc"
    
    'If Len(Trim(StrAppend)) Then
    '    sql = sql + StrAppend + sql2
    'Else
    '    sql = sql + sql2
    'End If
           
    'rst2.Open sql, medixmasterconnWS, adOpenForwardOnly, adLockBatchOptimistic
    Set cmd.ActiveConnection = medixmasterconnWS
        cmd.CommandText = "DCUReport"
        cmd.CommandType = adCmdStoredProc
        cmd.CommandTimeout = 300
        Set prm1 = cmd.CreateParameter("StrAppend", adVarChar, adParamInput, 2000, StrAppend)
        cmd.Parameters.Append prm1
        Set prm2 = cmd.CreateParameter("StrAppendCase", adVarChar, adParamInput, 2000, strAppendCase)
        cmd.Parameters.Append prm2
        rst2.CursorLocation = adUseClient
        rst2.CursorType = adOpenForwardOnly
        rst2.CacheSize = 100
        Set rst2 = cmd.Execute
        
        Do Until rst2.EOF
        
            With rst2
                If Len(rst2!PAYCode) Then
                    Set Record = ListView2.ListItems.Add(, , rst2!PAYCode)
                    Payor = rst2!PAYCode
                Else
                    Set Record = ListView2.ListItems.Add(, , "Empty")
                    Payor = ""
                End If
                
                If Len(rst2!PLNCode) Then
                    Record.SubItems(1) = Trim(rst2!PLNCode)
                    plancode = rst2!PLNCode
                Else
                    Record.SubItems(1) = "Empty"
                    plancode = ""
                End If
                
                If Len(rst2!HOSName) Then
                    Record.SubItems(2) = Trim(rst2!HOSName)
                    Hospital = rst2!HOSName
                Else
                    Record.SubItems(2) = "Empty"
                    Hospital = ""
                End If
                           
                If Len(rst2!TotalCases) Then
                    Record.SubItems(3) = rst2!TotalCases
                End If
                
                If cboClaimCase.ListIndex = 0 Then
                
                    If rst2!NoofDay = "0" Then
                        Record.SubItems(4) = "1"
                    ElseIf Len(rst2!NoofDay) Then
                        Record.SubItems(4) = rst2!NoofDay
                    End If
                
                ElseIf cboClaimCase.ListIndex = 1 Then
                
                    HOspitals = Replace(Hospital, "'", "''")
                    
                    If Payor = "" And plancode = "" And HOspitals = "" Then
                        strAppendCase = "27"
                        StrAppend1 = " AND (PAYCode = '' OR PAYCode IS NULL) AND " & _
                           " (PLNCode = '' OR PLNCode IS NULL) AND (HOSName = '' OR HOSName IS NULL) "
                        StrAppend1 = StrAppend1 + StrAppend
                    'sql3 = " SELECT PAYCode, PLNCode, HOSName, SUM(DurationStay) AS NoofDay " & _
                           " From VIEW_CLMReportDCU2 Where (PAYCode = '' OR PAYCode IS NULL) AND " & _
                           " (PLNCode = '' OR PLNCode IS NULL) AND (HOSName = '' OR HOSName IS NULL) "
                           
                    ElseIf Payor <> "" And plancode = "" And HOspitals = "" Then
                        strAppendCase = "27"
                        StrAppend1 = " AND PAYCode = '" & Trim(HOspitals) & "' AND " & _
                           " (PLNCode = '' OR PLNCode IS NULL) AND (HOSName = '' OR HOSName IS NULL) "
                        StrAppend1 = StrAppend1 + StrAppend
                    'sql3 = " SELECT PAYCode, PLNCode, HOSName, SUM(DurationStay) AS NoofDay " & _
                           " From VIEW_CLMReportDCU2 Where PAYCode = '" & Trim(Hospitals) & "' AND " & _
                           " (PLNCode = '' OR PLNCode IS NULL) AND (HOSName = '' OR HOSName IS NULL) "
                           
                    ElseIf Payor = "" And plancode <> "" And HOspitals = "" Then
                        strAppendCase = "27"
                        StrAppend1 = " AND (PAYCode = '' OR PAYCode IS NULL) AND " & _
                           " PLNCode = '" & Trim(plancode) & "' AND (HOSName = '' OR HOSName IS NULL) "
                        StrAppend1 = StrAppend1 + StrAppend
                    'sql3 = " SELECT PAYCode, PLNCode, HOSName, SUM(DurationStay) AS NoofDay " & _
                           " From VIEW_CLMReportDCU2 Where (PAYCode = '' OR PAYCode IS NULL) AND " & _
                           " PLNCode = '" & Trim(plancode) & "' AND (HOSName = '' OR HOSName IS NULL) "
                           
                    ElseIf Payor = "" And plancode = "" And HOspitals <> "" Then
                        strAppendCase = "27"
                        StrAppend1 = " AND (PAYCode = '' OR PAYCode IS NULL) AND " & _
                           " (PLNCode = '' OR PLNCode IS NULL) AND HOSName = '" & Trim(HOspitals) & "'"
                        StrAppend1 = StrAppend1 + StrAppend
                    'sql3 = " SELECT PAYCode, PLNCode, HOSName, SUM(DurationStay) AS NoofDay " & _
                           " From VIEW_CLMReportDCU2 Where (PAYCode = '' OR PAYCode IS NULL) AND " & _
                           " (PLNCode = '' OR PLNCode IS NULL) AND HOSName = '" & Trim(Hospitals) & "'"
                           
                    Else
                        strAppendCase = "27"
                        StrAppend1 = " AND PAYCode = '" & Trim(Payor) & "' AND " & _
                           " PLNCode = '" & Trim(plancode) & "' AND HOSName = '" & Trim(HOspitals) & "'"
                        StrAppend1 = StrAppend1 + StrAppend
                    'sql3 = " SELECT PAYCode, PLNCode, HOSName, SUM(DurationStay) AS NoofDay " & _
                           " From VIEW_CLMReportDCU2 Where PAYCode = '" & Trim(Payor) & "' AND " & _
                           " PLNCode = '" & Trim(plancode) & "' AND HOSName = '" & Trim(Hospitals) & "'"
                           
                    End If
                       
                       'sql2 = " Group By PAYCode, PLNCode, HOSName "
                       
                       'If Len(Trim(StrAppend)) Then
                       '     sql3 = sql3 + StrAppend + sql2
                       ' Else
                       '     sql3 = sql3 + sql2
                       ' End If
                                              
                       'rst1.Open sql3, medixmasterconnWS, adOpenForwardOnly, adLockBatchOptimistic
                        Call DurationStay(StrAppend1, strAppendCase, Record)
                End If
                
                If Len(rst2!TotalActual) Then
                    Record.SubItems(5) = Format(rst2!TotalActual, "#,##0.00")
                End If
                
                If Len(rst2!TotalApproved) Then
                    Record.SubItems(6) = Format(rst2!TotalApproved, "#,##0.00")
                End If
            End With
            lblCurrent = lblCurrent + 1
            rst2.MoveNext
       Loop
    If ListView2.ListItems.Count = 0 Then
        MsgBox "No record found!", vbOKOnly + vbCritical, DialogBoxTitle
        CmdPrint.Enabled = True
        CmdClear.Enabled = True
        cmdExit.Enabled = True
    End If
    rst2.Close
    Set rst2 = Nothing
    CmdPrint.Enabled = True
    CmdClear.Enabled = True
    cmdExit.Enabled = True
End Sub
'********************** List Plan Top 10 Hospital Results *******************************

Private Sub PLNFinalICD(StrAppend1, strAppendCase, Record)
Dim rst3 As New ADODB.Recordset
Dim con As New ADODB.Command
Dim prm1 As ADODB.Parameter
Dim prm2 As ADODB.Parameter

Set con.ActiveConnection = medixmasterconn
    con.CommandText = "DCUReport"
    con.CommandType = adCmdStoredProc
    con.CommandTimeout = 300
    Set prm1 = con.CreateParameter("StrAppend", adVarChar, adParamInput, 2000, StrAppend1)
    con.Parameters.Append prm1
    Set prm2 = con.CreateParameter("StrAppendCase", adVarChar, adParamInput, 2000, strAppendCase)
    con.Parameters.Append prm2
    rst3.CursorLocation = adUseClient
    rst3.CursorType = adOpenForwardOnly
    rst3.CacheSize = 100
    Set rst3 = con.Execute
    
    If rst3.EOF = True Then
        Record.SubItems(3) = ""
        rst3.Close
        Set rst3 = Nothing
    Else
        Record.SubItems(3) = Trim(rst3!ICDDescription)
        rst3.Close
        Set rst3 = Nothing
    End If
End Sub

'********************** List Plan Top 10 Final Diag Results *******************************
Private Sub PlanTopDiagListing(Optional StrAppend As String)
Dim rst2 As New ADODB.Recordset
Dim cmd As New ADODB.Command
Dim prm1 As ADODB.Parameter
Dim prm2 As ADODB.Parameter
Dim strAppendCase As String
Dim StrAppend1 As String
Dim Record As ListItem

    lblCurrent = 0
    Payor = ""
    plancode = ""
    DiagCode = ""
    ListView2.ListItems.Clear
        
    If cboClaimCase.ListIndex = 0 Then
        strAppendCase = "28"
    'sql = " SELECT Top 10 PAYCode, PLNCode, FinalDiag, COUNT(CASNumber) as TotalCases, SUM(clmtotalactual) AS totalActual, " & _
          " SUM(clmtotalapproved) As TotalApproved, " & _
          " SUM(NoofDay) AS NoofDay From VIEW_CLMReportDCU Where 0 = 0 "
          
    ElseIf cboClaimCase.ListIndex = 1 Then
        strAppendCase = "29"
    'sql = " SELECT Top 10 PAYCode, PLNCode, FinalDiag, COUNT(DISTINCT CASNumber) as TotalCases, SUM(clmtotalactual) AS totalActual, " & _
          " SUM(clmtotalapproved) As TotalApproved " & _
          " From VIEW_CLMReportDCU Where 0 = 0 "
    
    End If
    
    'sql2 = " Group By PAYCode, PLNCode, FinalDiag ORDER BY TotalCases Desc"
    
    'If Len(Trim(StrAppend)) Then
    '    sql = sql + StrAppend + sql2
    'Else
    '    sql = sql + sql2
    'End If
            
    'rst2.Open sql, medixmasterconnWS, adOpenForwardOnly, adLockBatchOptimistic
    Set cmd.ActiveConnection = medixmasterconnWS
        cmd.CommandText = "DCUReport"
        cmd.CommandType = adCmdStoredProc
        cmd.CommandTimeout = 300
        Set prm1 = cmd.CreateParameter("StrAppend", adVarChar, adParamInput, 2000, StrAppend)
        cmd.Parameters.Append prm1
        Set prm2 = cmd.CreateParameter("StrAppendCase", adVarChar, adParamInput, 2000, strAppendCase)
        cmd.Parameters.Append prm2
        rst2.CursorLocation = adUseClient
        rst2.CursorType = adOpenForwardOnly
        rst2.CacheSize = 100
        Set rst2 = cmd.Execute
        
        Do Until rst2.EOF
        
            With rst2
                If Len(rst2!PAYCode) Then
                    Set Record = ListView2.ListItems.Add(, , rst2!PAYCode)
                    Payor = rst2!PAYCode
                Else
                    Set Record = ListView2.ListItems.Add(, , "Empty")
                    Payor = ""
                End If
                
                If Len(rst2!PLNCode) Then
                    Record.SubItems(1) = rst2!PLNCode
                    plancode = rst2!PLNCode
                Else
                    Record.SubItems(1) = "Empty"
                    plancode = ""
                End If
                
                If Len(rst2!FinalDiag) Then
                    Record.SubItems(2) = rst2!FinalDiag
                    DiagCode = rst2!FinalDiag
                    strAppendCase = "1"
                    StrAppend1 = " AND ICDCode = '" & Trim(DiagCode) & "'"
                    'sql3 = " SELECT ICDCode, ICDDescription From ICDOriginal Where ICDCode = '" & Trim(DiagCode) & "'"
                    'rst3.Open sql3, medixmasterconn, adOpenForwardOnly, adLockBatchOptimistic
                    Call PLNFinalICD(StrAppend1, strAppendCase, Record)
                Else
                    Record.SubItems(2) = "Empty"
                    DiagCode = ""
                End If
                    
                If Len(rst2!TotalCases) Then
                    Record.SubItems(4) = rst2!TotalCases
                End If
                
                If cboClaimCase.ListIndex = 0 Then
                    
                    If rst2!NoofDay = "0" Then
                        Record.SubItems(5) = "1"
                    ElseIf Len(rst2!NoofDay) Then
                        Record.SubItems(5) = rst2!NoofDay
                    End If
                
                ElseIf cboClaimCase.ListIndex = 1 Then
                
                    If Payor = "" And plancode And DiagCode = "" Then
                        strAppendCase = "30"
                        StrAppend1 = " AND (PAYCode = '' OR PAYCode IS NULL) AND " & _
                           " (PLNCode = '' OR PLNCode IS NULL) And (FinalDiag = '' OR FinalDiag IS NULL)"
                        StrAppend1 = StrAppend1 + StrAppend
                    'sql4 = " SELECT PAYCode, PLNCode, FinalDiag, SUM(DurationStay) AS NoofDay " & _
                           " From VIEW_CLMReportDCU2 Where (PAYCode = '' OR PAYCode IS NULL) AND " & _
                           " (PLNCode = '' OR PLNCode IS NULL) And (FinalDiag = '' OR FinalDiag IS NULL)"
                           
                    ElseIf Payor <> "" And plancode = "" And DiagCode = "" Then
                        strAppendCase = "30"
                        StrAppend1 = " AND PAYCode = '" & Trim(Payor) & "' AND " & _
                           " (PLNCode = '' OR PLNCode IS NULL) And (FinalDiag = '' OR FinalDiag IS NULL)"
                        StrAppend1 = StrAppend1 + StrAppend
                    'sql4 = " SELECT PAYCode, PLNCode, FinalDiag, SUM(DurationStay) AS NoofDay " & _
                           " From VIEW_CLMReportDCU2 Where PAYCode = '" & Trim(Payor) & "' AND " & _
                           " (PLNCode = '' OR PLNCode IS NULL) And (FinalDiag = '' OR FinalDiag IS NULL)"
                           
                    ElseIf Payor = "" And plancode <> "" And DiagCode = "" Then
                        strAppendCase = "30"
                        StrAppend1 = " AND (PAYCode = '' OR PAYCode IS NULL) AND " & _
                           " PLNCode = '" & Trim(plancode) & "' And (FinalDiag = '' OR FinalDiag IS NULL) "
                        StrAppend1 = StrAppend1 + StrAppend
                    'sql4 = " SELECT PAYCode, PLNCode, FinalDiag, SUM(DurationStay) AS NoofDay " & _
                           " From VIEW_CLMReportDCU2 Where (PAYCode = '' OR PAYCode IS NULL) AND " & _
                           " PLNCode = '" & Trim(plancode) & "' And (FinalDiag = '' OR FinalDiag IS NULL) "
                           
                    ElseIf Payor = "" And plancode = "" And DiagCode <> "" Then
                        strAppendCase = "30"
                        StrAppend1 = " AND (PAYCode = '' OR PAYCode IS NULL) AND " & _
                           " (PLNCode = '' OR PLNCode IS NULL) And FinalDiag = '" & Trim(DiagCode) & "'"
                        StrAppend1 = StrAppend1 + StrAppend
                    'sql4 = " SELECT PAYCode, PLNCode, FinalDiag, SUM(DurationStay) AS NoofDay " & _
                           " From VIEW_CLMReportDCU2 Where (PAYCode = '' OR PAYCode IS NULL) AND " & _
                           " (PLNCode = '' OR PLNCode IS NULL) And FinalDiag = '" & Trim(DiagCode) & "'"
                    
                    Else
                        strAppendCase = "30"
                        StrAppend1 = " AND PAYCode = '" & Trim(Payor) & "' AND " & _
                           " PLNCode = '" & Trim(plancode) & "' And FinalDiag = '" & Trim(DiagCode) & "'"
                        StrAppend1 = StrAppend1 + StrAppend
                    'sql4 = " SELECT PAYCode, PLNCode, FinalDiag, SUM(DurationStay) AS NoofDay " & _
                           " From VIEW_CLMReportDCU2 Where PAYCode = '" & Trim(Payor) & "' AND " & _
                           " PLNCode = '" & Trim(plancode) & "' And FinalDiag = '" & Trim(DiagCode) & "'"
                    
                    End If
                       
                       'sql2 = " Group By PAYCode, PLNCode, FinalDiag "
                       
                       'If Len(Trim(StrAppend)) Then
                       '     sql4 = sql4 + StrAppend + sql2
                       ' Else
                       '     sql4 = sql4 + sql2
                       ' End If
                                              
                       'rst1.Open sql4, medixmasterconnWS, adOpenForwardOnly, adLockBatchOptimistic
                        Call DurationStay(StrAppend1, strAppendCase, Record)
                End If
                    
                If Len(rst2!TotalActual) Then
                    Record.SubItems(6) = Format(rst2!TotalActual, "#,##0.00")
                End If
                
                If Len(rst2!TotalApproved) Then
                    Record.SubItems(7) = Format(rst2!TotalApproved, "#,##0.00")
                End If
                
            End With
            lblCurrent = lblCurrent + 1
            rst2.MoveNext
       Loop
    If ListView2.ListItems.Count = 0 Then
        MsgBox "No record found!", vbOKOnly + vbCritical, DialogBoxTitle
        CmdPrint.Enabled = True
        CmdClear.Enabled = True
        cmdExit.Enabled = True
    End If
    rst2.Close
    Set rst2 = Nothing
    CmdPrint.Enabled = True
    CmdClear.Enabled = True
    cmdExit.Enabled = True
End Sub
'********************** List Plan Top 10 Final Diag Results *******************************

'********************** List Plan Type by Ageband Results *******************************
Private Sub PlanAgebandListing(Optional StrAppend As String)
Dim rst2 As New ADODB.Recordset
Dim cmd As New ADODB.Command
Dim prm1 As ADODB.Parameter
Dim prm2 As ADODB.Parameter
Dim strAppendCase As String
Dim StrAppend1 As String
Dim Record As ListItem

    lblCurrent = 0
    Payor = ""
    plancode = ""
    AgeCode = ""
    ListView2.ListItems.Clear
    
    If cboClaimCase.ListIndex = 0 Then
        strAppendCase = "31"
    'sql = " SELECT PAYCode, PLNCode, Ageband, COUNT(CASNumber) AS TotalCases, " & _
          " SUM(CLMTotalActual) AS TotalActualAmt, " & _
          " SUM(CLMTotalApproved) AS TotalApprovedAmt, " & _
          " SUM(NoofDay) AS NoofDay From VIEW_CLMReportDCU Where 0 = 0 "
          
    ElseIf cboClaimCase.ListIndex = 1 Then
        strAppendCase = "32"
    'sql = " SELECT PAYCode, PLNCode, Ageband, COUNT(DISTINCT CASNumber) AS TotalCases, " & _
          " SUM(CLMTotalActual) AS TotalActualAmt, " & _
          " SUM(CLMTotalApproved) AS TotalApprovedAmt " & _
          " From VIEW_CLMReportDCU Where 0 = 0 "
    
    End If
          
    'sql2 = " GROUP BY PAYCode, PLNCode, Ageband "
    
    'If Len(Trim(StrAppend)) Then
    '    sql = sql + StrAppend + sql2
    'Else
    '    sql = sql + sql2
    'End If
    
    'rst2.Open sql, medixmasterconnWS, adOpenForwardOnly, adLockBatchOptimistic
    Set cmd.ActiveConnection = medixmasterconnWS
        cmd.CommandText = "DCUReport"
        cmd.CommandType = adCmdStoredProc
        cmd.CommandTimeout = 300
        Set prm1 = cmd.CreateParameter("StrAppend", adVarChar, adParamInput, 2000, StrAppend)
        cmd.Parameters.Append prm1
        Set prm2 = cmd.CreateParameter("StrAppendCase", adVarChar, adParamInput, 2000, strAppendCase)
        cmd.Parameters.Append prm2
        rst2.CursorLocation = adUseClient
        rst2.CursorType = adOpenForwardOnly
        rst2.CacheSize = 100
        Set rst2 = cmd.Execute
          
        Do Until rst2.EOF
        
            With rst2
                If Len(rst2!PAYCode) Then
                    Set Record = ListView2.ListItems.Add(, , rst2!PAYCode)
                    Payor = rst2!PAYCode
                Else
                    Set Record = ListView2.ListItems.Add(, , "Empty")
                    Payor = ""
                End If
                
                If Len(rst2!PLNCode) Then
                   Record.SubItems(1) = rst2!PLNCode
                   plancode = rst2!PLNCode
                Else
                   Record.SubItems(1) = "Empty"
                   plancode = ""
                End If
                    
                If Len(rst2!AgeBand) Then
                   Record.SubItems(2) = rst2!AgeBand
                   AgeCode = rst2!AgeBand
                Else
                    Record.SubItems(2) = "Empty"
                    AgeCode = ""
                End If
                                    
                If Len(rst2!TotalCases) Then
                   Record.SubItems(3) = rst2!TotalCases
                End If
                
                If cboClaimCase.ListIndex = 0 Then
                
                    If rst2!NoofDay = "0" Then
                        Record.SubItems(4) = "1"
                    ElseIf Len(rst2!NoofDay) Then
                        Record.SubItems(4) = rst2!NoofDay
                    End If
                
                ElseIf cboClaimCase.ListIndex = 1 Then
                
                    If Payor = "" And plancode = "" And AgeCode = "" Then
                        strAppendCase = "33"
                        StrAppend1 = " AND (PAYCode = '' OR PAYCode IS NULL) AND " & _
                           " (PLNCode = '' OR PLNCode IS NULL) AND (AGEBand = '' OR AGEBand IS NULL)"
                        StrAppend1 = StrAppend1 + StrAppend
                    'sql3 = " SELECT PAYCode, PLNCode, AGEBand, SUM(DurationStay) AS NoofDay " & _
                           " From VIEW_CLMReportDCU2 Where (PAYCode = '' OR PAYCode IS NULL) AND " & _
                           " (PLNCode = '' OR PLNCode IS NULL) AND (AGEBand = '' OR AGEBand IS NULL)"
                           
                    ElseIf Payor <> "" And plancode = "" And AgeCode = "" Then
                        strAppendCase = "33"
                        StrAppend1 = " AND PAYCode = '" & Trim(Payor) & "' AND " & _
                           " (PLNCode = '' OR PLNCode IS NULL) AND (AGEBand = '' OR AGEBand IS NULL)"
                        StrAppend1 = StrAppend1 + StrAppend
                    'sql3 = " SELECT PAYCode, PLNCode, AGEBand, SUM(DurationStay) AS NoofDay " & _
                           " From VIEW_CLMReportDCU2 Where PAYCode = '" & Trim(Payor) & "' AND " & _
                           " (PLNCode = '' OR PLNCode IS NULL) AND (AGEBand = '' OR AGEBand IS NULL)"
                           
                    ElseIf Payor = "" And plancode <> "" And AgeCode = "" Then
                        strAppendCase = "33"
                        StrAppend1 = " AND (PAYCode = '' OR PAYCode IS NULL) AND " & _
                           " PLNCode = '" & Trim(plancode) & "' AND (AGEBand = '' OR AGEBand IS NULL)"
                        StrAppend1 = StrAppend1 + StrAppend
                    'sql3 = " SELECT PAYCode, PLNCode, AGEBand, SUM(DurationStay) AS NoofDay " & _
                           " From VIEW_CLMReportDCU2 Where (PAYCode = '' OR PAYCode IS NULL) AND " & _
                           " PLNCode = '" & Trim(plancode) & "' AND (AGEBand = '' OR AGEBand IS NULL)"
                           
                    ElseIf Payor = "" And plancode = "" And AgeCode <> "" Then
                        strAppendCase = "33"
                        StrAppend1 = " AND (PAYCode = '' OR PAYCode IS NULL) AND " & _
                           " (PLNCode = '' OR PLNCode IS NULL) AND AGEBand = '" & Trim(AgeCode) & "'"
                        StrAppend1 = StrAppend1 + StrAppend
                    'sql3 = " SELECT PAYCode, PLNCode, AGEBand, SUM(DurationStay) AS NoofDay " & _
                           " From VIEW_CLMReportDCU2 Where (PAYCode = '' OR PAYCode IS NULL) AND " & _
                           " (PLNCode = '' OR PLNCode IS NULL) AND AGEBand = '" & Trim(AgeCode) & "'"
                           
                    Else
                        strAppendCase = "33"
                        StrAppend1 = " AND PAYCode = '" & Trim(Payor) & "' AND " & _
                           " PLNCode = '" & Trim(plancode) & "' AND AGEBand = '" & Trim(AgeCode) & "'"
                        StrAppend1 = StrAppend1 + StrAppend
                    'sql3 = " SELECT PAYCode, PLNCode, AGEBand, SUM(DurationStay) AS NoofDay " & _
                           " From VIEW_CLMReportDCU2 Where PAYCode = '" & Trim(Payor) & "' AND " & _
                           " PLNCode = '" & Trim(plancode) & "' AND AGEBand = '" & Trim(AgeCode) & "'"
                           
                    End If
                       
                       'sql2 = " Group By PAYCode, PLNCode, AGEBand "
                       
                       'If Len(Trim(StrAppend)) Then
                       '     sql3 = sql3 + StrAppend + sql2
                       ' Else
                       '     sql3 = sql3 + sql2
                       ' End If
                                              
                       'rst1.Open sql3, medixmasterconnWS, adOpenForwardOnly, adLockBatchOptimistic
                        Call DurationStay(StrAppend1, strAppendCase, Record)
                End If
                                
                If Len(rst2!TotalActualAmt) Then
                   Record.SubItems(5) = Format(rst2!TotalActualAmt, "#,##0.00")
                End If
                
                If Len(rst2!TotalApprovedAmt) Then
                   Record.SubItems(6) = Format(rst2!TotalApprovedAmt, "#,##0.00")
                End If
                
            End With
            lblCurrent = lblCurrent + 1
            rst2.MoveNext
       Loop
    'End If
    If ListView2.ListItems.Count = 0 Then
        MsgBox "No record found!", vbOKOnly + vbCritical, DialogBoxTitle
        CmdPrint.Enabled = True
        CmdClear.Enabled = True
        cmdExit.Enabled = True
    End If
    rst2.Close
    Set rst2 = Nothing
    CmdPrint.Enabled = True
    CmdClear.Enabled = True
    cmdExit.Enabled = True
End Sub

'********************** List Plan Type By Ageband Results *******************************

'********************** List Plan Type by Insured Results *******************************
Private Sub PlanInsuredListing(Optional StrAppend As String)
Dim rst2 As New ADODB.Recordset
Dim cmd As New ADODB.Command
Dim prm1 As ADODB.Parameter
Dim prm2 As ADODB.Parameter
Dim strAppendCase As String
Dim StrAppend1 As String
Dim Record As ListItem
Dim strCount

    lblCurrent = 0
    Payor = ""
    plancode = ""
    INSCode = ""
    ListView2.ListItems.Clear
            
    If cboClaimCase.ListIndex = 0 Then
        strAppendCase = "7"
    'sql = " SELECT PAYCode, PLNCode, INSCode, COUNT(CASNumber) AS TotalCases, " & _
          " SUM(CLMTotalActual) AS TotalActualAmt, " & _
          " SUM(CLMTotalApproved) AS TotalApprovedAmt, " & _
          " SUM(NoofDay) AS NoofDay From VIEW_CLMReportDCU Where 0 = 0 "
          
    ElseIf cboClaimCase.ListIndex = 1 Then
        strAppendCase = "8"
        'sql = " SELECT PAYCode, PLNCode, INSCode, COUNT(DISTINCT CASNumber) AS TotalCases, " & _
              " SUM(CLMTotalActual) AS TotalActualAmt, " & _
              " SUM(CLMTotalApproved) AS TotalApprovedAmt " & _
              " From VIEW_CLMReportDCU Where 0 = 0 "
    
    End If
          
    'sql2 = " GROUP BY PAYCode, PLNCode, INSCode "
    
    'If Len(Trim(StrAppend)) Then
    '    sql = sql + StrAppend + sql2
    'Else
    '    sql = sql + sql2
    'End If
    
    'rst2.Open sql, medixmasterconnWS, adOpenForwardOnly, adLockBatchOptimistic
    Set cmd.ActiveConnection = medixmasterconnWS
        cmd.CommandText = "DCUReport"
        cmd.CommandType = adCmdStoredProc
        cmd.CommandTimeout = 300
        Set prm1 = cmd.CreateParameter("StrAppend", adVarChar, adParamInput, 2000, StrAppend)
        cmd.Parameters.Append prm1
        Set prm2 = cmd.CreateParameter("StrAppendCase", adVarChar, adParamInput, 2000, strAppendCase)
        cmd.Parameters.Append prm2
        rst2.CursorLocation = adUseClient
        rst2.CursorType = adOpenForwardOnly
        rst2.CacheSize = 100
        Set rst2 = cmd.Execute
        
        Do Until rst2.EOF
        
            With rst2
                If Len(rst2!PAYCode) Then
                    Set Record = ListView2.ListItems.Add(, , rst2!PAYCode)
                    Payor = rst2!PAYCode
                Else
                    Set Record = ListView2.ListItems.Add(, , "Empty")
                    Payor = ""
                End If
                
                If Len(rst2!INSCode) Then
                   Record.SubItems(1) = rst2!INSCode
                   INSCode = rst2!INSCode
                Else
                   Record.SubItems(1) = "Empty"
                   INSCode = ""
                End If
                    
                If Len(rst2!PLNCode) Then
                   Record.SubItems(2) = rst2!PLNCode
                   plancode = rst2!PLNCode
                Else
                   Record.SubItems(2) = "Empty"
                   plancode = ""
                End If
                    
                If Len(rst2!TotalCases) Then
                   Record.SubItems(3) = rst2!TotalCases
                End If
                    
                If cboClaimCase.ListIndex = 0 Then
                
                    If rst2!NoofDay = "0" Then
                        Record.SubItems(4) = "1"
                    ElseIf Len(rst2!NoofDay) Then
                        Record.SubItems(4) = rst2!NoofDay
                    End If
                
                ElseIf cboClaimCase.ListIndex = 1 Then
                
                    If Payor = "" And plancode = "" And INSCode = "" Then
                        strAppendCase = "9"
                        StrAppend1 = " AND (PAYCode = '' OR PAYCode IS NULL) AND " & _
                           " (PLNCode = '' OR PLNCode IS NULL) AND (INSCode = '' OR INSCode IS NULL)"
                        StrAppend1 = StrAppend1 + StrAppend
                    'sql3 = " SELECT PAYCode, PLNCode, INSCode, SUM(DurationStay) AS NoofDay " & _
                           " From VIEW_CLMReportDCU2 Where (PAYCode = '' OR PAYCode IS NULL) AND " & _
                           " (PLNCode = '' OR PLNCode IS NULL) AND (INSCode = '' OR INSCode IS NULL)"
                           
                    ElseIf Payor <> "" And plancode = "" And INSCode = "" Then
                        strAppendCase = "9"
                        StrAppend1 = " AND PAYCode = '" & Trim(Payor) & "' AND " & _
                           " (PLNCode = '' OR PLNCode IS NULL) AND (INSCode = '' OR INSCode IS NULL)"
                        StrAppend1 = StrAppend1 + StrAppend
                    'sql3 = " SELECT PAYCode, PLNCode, INSCode, SUM(DurationStay) AS NoofDay " & _
                           " From VIEW_CLMReportDCU2 Where PAYCode = '" & Trim(Payor) & "' AND " & _
                           " (PLNCode = '' OR PLNCode IS NULL) AND (INSCode = '' OR INSCode IS NULL)"
                           
                    ElseIf Payor = "" And plancode <> "" And INSCode = "" Then
                        strAppendCase = "9"
                        StrAppend1 = " AND (PAYCode = '' OR PAYCode IS NULL) AND " & _
                           " PLNCode = '" & Trim(plancode) & "' AND (INSCode = '' OR INSCode IS NULL)"
                        StrAppend1 = StrAppend1 + StrAppend
                    'sql3 = " SELECT PAYCode, PLNCode, INSCode, SUM(DurationStay) AS NoofDay " & _
                           " From VIEW_CLMReportDCU2 Where (PAYCode = '' OR PAYCode IS NULL) AND " & _
                           " PLNCode = '" & Trim(plancode) & "' AND (INSCode = '' OR INSCode IS NULL)"
                           
                    ElseIf Payor = "" And plancode = "" And INSCode <> "" Then
                        strAppendCase = "9"
                        StrAppend1 = " AND (PAYCode = '' OR PAYCode IS NULL) AND " & _
                           " (PLNCode = '' OR PLNCode IS NULL) AND INSCode = '" & Trim(INSCode) & "'"
                        StrAppend1 = StrAppend1 + StrAppend
                    'sql3 = " SELECT PAYCode, PLNCode, INSCode, SUM(DurationStay) AS NoofDay " & _
                           " From VIEW_CLMReportDCU2 Where (PAYCode = '' OR PAYCode IS NULL) AND " & _
                           " (PLNCode = '' OR PLNCode IS NULL) AND INSCode = '" & Trim(INSCode) & "'"
                           
                    Else
                        strAppendCase = "9"
                        StrAppend1 = " AND PAYCode = '" & Trim(Payor) & "' AND " & _
                           " PLNCode = '" & Trim(plancode) & "' AND INSCode = '" & Trim(INSCode) & "'"
                        StrAppend1 = StrAppend1 + StrAppend
                    'sql3 = " SELECT PAYCode, PLNCode, INSCode, SUM(DurationStay) AS NoofDay " & _
                           " From VIEW_CLMReportDCU2 Where PAYCode = '" & Trim(Payor) & "' AND " & _
                           " PLNCode = '" & Trim(plancode) & "' AND INSCode = '" & Trim(INSCode) & "'"
                           
                    End If
                       
                       'sql2 = " Group By PAYCode, PLNCode, INSCode "
                       
                       'If Len(Trim(StrAppend)) Then
                       '     sql3 = sql3 + StrAppend + sql2
                       ' Else
                       '     sql3 = sql3 + sql2
                       ' End If
                                              
                       'rst1.Open sql3, medixmasterconnWS, adOpenForwardOnly, adLockBatchOptimistic
                        Call DurationStay(StrAppend1, strAppendCase, Record)
                End If
                                
                If Len(rst2!TotalActualAmt) Then
                   Record.SubItems(5) = Format(rst2!TotalActualAmt, "#,##0.00")
                End If
                
                If Len(rst2!TotalApprovedAmt) Then
                   Record.SubItems(6) = Format(rst2!TotalApprovedAmt, "#,##0.00")
                End If
                    
            End With
            lblCurrent = lblCurrent + 1
            rst2.MoveNext
       Loop
    
    If ListView2.ListItems.Count = 0 Then
        MsgBox "No record found!", vbOKOnly + vbCritical, DialogBoxTitle
        CmdPrint.Enabled = True
        CmdClear.Enabled = True
        cmdExit.Enabled = True
    End If
    rst2.Close
    Set rst2 = Nothing
    CmdPrint.Enabled = True
    CmdClear.Enabled = True
    cmdExit.Enabled = True
End Sub

Private Sub PlanInsuredListing2(Optional StrAppend As String)
Dim rst2 As New ADODB.Recordset
Dim cmd As New ADODB.Command
Dim prm1 As ADODB.Parameter
Dim prm2 As ADODB.Parameter
Dim strAppendCase As String
Dim StrAppend1 As String
Dim Record As ListItem

    lblCurrent = 0
    Payor = ""
    plancode = ""
    INSCode = ""
    AgeCode = ""
    ListView2.ListItems.Clear
            
    If cboClaimCase.ListIndex = 0 Then
        strAppendCase = "34"
    'sql = " SELECT PAYCode, PLNCode, INSCode, Ageband, COUNT(CASNumber) AS TotalCases, " & _
          " SUM(CLMTotalActual) AS TotalActualAmt, " & _
          " SUM(CLMTotalApproved) AS TotalApprovedAmt, " & _
          " SUM(NoofDay) AS NoofDay From VIEW_CLMReportDCU Where 0 = 0 "
          
    ElseIf cboClaimCase.ListIndex = 1 Then
        strAppendCase = "35"
    'sql = " SELECT PAYCode, PLNCode, INSCode, Ageband, COUNT(DISTINCT CASNumber) AS TotalCases, " & _
          " SUM(CLMTotalActual) AS TotalActualAmt, " & _
          " SUM(CLMTotalApproved) AS TotalApprovedAmt " & _
          " From VIEW_CLMReportDCU Where 0 = 0 "
    
    End If
          
    'sql2 = " GROUP BY PAYCode, PLNCode, INSCode, Ageband "
    
    'If Len(Trim(StrAppend)) Then
    '    sql = sql + StrAppend + sql2
    'Else
    '    sql = sql + sql2
    'End If
    
    'rst2.Open sql, medixmasterconnWS, adOpenForwardOnly, adLockBatchOptimistic
    Set cmd.ActiveConnection = medixmasterconnWS
        cmd.CommandText = "DCUReport"
        cmd.CommandType = adCmdStoredProc
        cmd.CommandTimeout = 300
        Set prm1 = cmd.CreateParameter("StrAppend", adVarChar, adParamInput, 2000, StrAppend)
        cmd.Parameters.Append prm1
        Set prm2 = cmd.CreateParameter("StrAppendCase", adVarChar, adParamInput, 2000, strAppendCase)
        cmd.Parameters.Append prm2
        rst2.CursorLocation = adUseClient
        rst2.CursorType = adOpenForwardOnly
        rst2.CacheSize = 100
        Set rst2 = cmd.Execute
        
        Do Until rst2.EOF
        
            With rst2
                If Len(rst2!PAYCode) Then
                    Set Record = ListView2.ListItems.Add(, , rst2!PAYCode)
                    Payor = rst2!PAYCode
                Else
                    Set Record = ListView2.ListItems.Add(, , "Empty")
                    Payor = ""
                End If
                
                If Len(rst2!INSCode) Then
                   Record.SubItems(1) = rst2!INSCode
                   INSCode = rst2!INSCode
                Else
                    Record.SubItems(1) = "Empty"
                    INSCode = ""
                End If
                    
                If Len(rst2!PLNCode) Then
                   Record.SubItems(2) = rst2!PLNCode
                    plancode = rst2!PLNCode
                Else
                    Record.SubItems(2) = "Empty"
                    plancode = ""
                End If
                
                If Len(rst2!AgeBand) Then
                   Record.SubItems(3) = rst2!AgeBand
                   AgeCode = rst2!AgeBand
                Else
                    Record.SubItems(3) = "Empty"
                    AgeCode = ""
                End If
                    
                If Len(rst2!TotalCases) Then
                   Record.SubItems(4) = rst2!TotalCases
                End If
                    
                If cboClaimCase.ListIndex = 0 Then
                
                    If rst2!NoofDay = "0" Then
                        Record.SubItems(5) = "1"
                    ElseIf Len(rst2!NoofDay) Then
                        Record.SubItems(5) = rst2!NoofDay
                    End If
                
                ElseIf cboClaimCase.ListIndex = 1 Then
                
                    If Payor = "" And plancode = "" And INSCode = "" And AgeCode = "" Then
                        strAppendCase = "36"
                        StrAppend1 = " AND (PAYCode = '' OR PAYCode IS NULL) AND " & _
                           " (PLNCode = '' OR PLNCode IS NULL) AND (INSCode = '' OR INSCode IS NULL) " & _
                           " AND (AGEBand = '' OR AGEBand IS NULL) "
                        StrAppend1 = StrAppend1 + StrAppend
                    'sql3 = " SELECT PAYCode, PLNCode, INSCode, SUM(DurationStay) AS NoofDay " & _
                           " From VIEW_CLMReportDCU2 Where (PAYCode = '' OR PAYCode IS NULL) AND " & _
                           " (PLNCode = '' OR PLNCode IS NULL) AND (INSCode = '' OR INSCode IS NULL) " & _
                           " AND (AGEBand = '' OR AGEBand IS NULL) "
                           
                    ElseIf Payor <> "" And plancode = "" And INSCode = "" And AgeCode = "" Then
                        strAppendCase = "36"
                        StrAppend1 = " AND PAYCode = '" & Trim(Payor) & "' AND " & _
                           " (PLNCode = '' OR PLNCode IS NULL) AND (INSCode = '' OR INSCode IS NULL)" & _
                           " AND (AGEBand = '' OR AGEBand IS NULL) "
                        StrAppend1 = StrAppend1 + StrAppend
                    'sql3 = " SELECT PAYCode, PLNCode, INSCode, SUM(DurationStay) AS NoofDay " & _
                           " From VIEW_CLMReportDCU2 Where PAYCode = '" & Trim(Payor) & "' AND " & _
                           " (PLNCode = '' OR PLNCode IS NULL) AND (INSCode = '' OR INSCode IS NULL)" & _
                           " AND (AGEBand = '' OR AGEBand IS NULL) "
                           
                    ElseIf Payor = "" And plancode <> "" And INSCode = "" And AgeCode = "" Then
                        strAppendCase = "36"
                        StrAppend1 = " AND (PAYCode = '' OR PAYCode IS NULL) AND " & _
                           " PLNCode = '" & Trim(plancode) & "' AND (INSCode = '' OR INSCode IS NULL)" & _
                           " AND (AGEBand = '' OR AGEBand IS NULL) "
                        StrAppend1 = StrAppend1 + StrAppend
                    'sql3 = " SELECT PAYCode, PLNCode, INSCode, SUM(DurationStay) AS NoofDay " & _
                           " From VIEW_CLMReportDCU2 Where (PAYCode = '' OR PAYCode IS NULL) AND " & _
                           " PLNCode = '" & Trim(plancode) & "' AND (INSCode = '' OR INSCode IS NULL)" & _
                           " AND (AGEBand = '' OR AGEBand IS NULL) "
                           
                    ElseIf Payor = "" And plancode = "" And INSCode <> "" And AgeCode = "" Then
                        strAppendCase = "36"
                        StrAppend1 = " AND (PAYCode = '' OR PAYCode IS NULL) AND " & _
                           " (PLNCode = '' OR PLNCode IS NULL) AND INSCode = '" & Trim(INSCode) & "'" & _
                           " AND (AGEBand = '' OR AGEBand IS NULL) "
                        StrAppend1 = StrAppend1 + StrAppend
                    'sql3 = " SELECT PAYCode, PLNCode, INSCode, SUM(DurationStay) AS NoofDay " & _
                           " From VIEW_CLMReportDCU2 Where (PAYCode = '' OR PAYCode IS NULL) AND " & _
                           " (PLNCode = '' OR PLNCode IS NULL) AND INSCode = '" & Trim(INSCode) & "'" & _
                           " AND (AGEBand = '' OR AGEBand IS NULL) "
                    
                    ElseIf Payor = "" And plancode = "" And INSCode = "" And AgeCode <> "" Then
                        strAppendCase = "36"
                        StrAppend1 = " AND (PAYCode = '' OR PAYCode IS NULL) AND " & _
                           " (PLNCode = '' OR PLNCode IS NULL) AND (INSCode = '' OR INSCode IS NULL) " & _
                           " AND AGEBand = '" & Trim(AgeCode) & "'"
                        StrAppend1 = StrAppend1 + StrAppend
                    'sql3 = " SELECT PAYCode, PLNCode, INSCode, SUM(DurationStay) AS NoofDay " & _
                           " From VIEW_CLMReportDCU2 Where (PAYCode = '' OR PAYCode IS NULL) AND " & _
                           " (PLNCode = '' OR PLNCode IS NULL) AND (INSCode = '' OR INSCode IS NULL) " & _
                           " AND AGEBand = '" & Trim(AgeCode) & "'"
                    
                    Else
                        strAppendCase = "36"
                        StrAppend1 = " AND PAYCode = '" & Trim(Payor) & "'  AND " & _
                           " PLNCode = '" & Trim(plancode) & "' AND INSCode = '" & Trim(INSCode) & "'" & _
                           " AND AGEBand = '" & Trim(AgeCode) & "'"
                        StrAppend1 = StrAppend1 + StrAppend
                    'sql3 = " SELECT PAYCode, PLNCode, INSCode, SUM(DurationStay) AS NoofDay " & _
                           " From VIEW_CLMReportDCU2 Where PAYCode = '" & Trim(Payor) & "'  AND " & _
                           " PLNCode = '" & Trim(plancode) & "' AND INSCode = '" & Trim(INSCode) & "'" & _
                           " AND AGEBand = '" & Trim(AgeCode) & "'"
                           
                    End If
                       
                       'sql2 = " Group By PAYCode, PLNCode, INSCode, AGEBand "
                       
                       'If Len(Trim(StrAppend)) Then
                       '     sql3 = sql3 + StrAppend + sql2
                       ' Else
                       '     sql3 = sql3 + sql2
                       ' End If
                                              
                       'rst1.Open sql3, medixmasterconnWS, adOpenForwardOnly, adLockBatchOptimistic
                        Call DurationStay(StrAppend1, strAppendCase, Record)
                End If
                                
                If Len(rst2!TotalActualAmt) Then
                   Record.SubItems(6) = Format(rst2!TotalActualAmt, "#,##0.00")
                End If
                
                If Len(rst2!TotalApprovedAmt) Then
                   Record.SubItems(7) = Format(rst2!TotalApprovedAmt, "#,##0.00")
                End If
                    
            End With
            lblCurrent = lblCurrent + 1
            rst2.MoveNext
       Loop
    
    If ListView2.ListItems.Count = 0 Then
        MsgBox "No record found!", vbOKOnly + vbCritical, DialogBoxTitle
        CmdPrint.Enabled = True
        CmdClear.Enabled = True
        cmdExit.Enabled = True
    End If
    rst2.Close
    Set rst2 = Nothing
    CmdPrint.Enabled = True
    CmdClear.Enabled = True
    cmdExit.Enabled = True
End Sub

'********************** List Plan Type By Insured Results *******************************

Private Sub CboCasePending_Change()
    If InStr(CboCasePending.Text, "null") Then
       CboCasePending.Enabled = False
    End If
End Sub

Private Sub CboCasePending_KeyPress(KeyAscii As Integer)
Dim NewText As String
Dim ValidCount As Integer
Dim ValidValue As String
Dim i As Integer
    If KeyAscii >= 32 And KeyAscii <> 127 Then
    NewText = LCase(Left(CboCasePending.Text, CboCasePending.SelStart) + Chr(KeyAscii) + Mid(CboCasePending.Text, CboCasePending.SelStart + CboCasePending.SelLength + 1))
 
    ValidCount = 0
    ValidValue = ""
 
    For i = 0 To CboCasePending.ListCount - 1
 
        If NewText = LCase(Left(CboCasePending.List(i), Len(NewText))) Then
                ValidCount = ValidCount + 1
                ValidValue = CboCasePending.List(i)
        End If
        Next
        If ValidCount <= 1 Then KeyAscii = 0
            If ValidCount = 1 Then
                CboCasePending.Text = ValidValue
                CboCasePending.SelStart = 0
                CboCasePending.SelLength = Len(ValidValue)
            End If
        End If
End Sub

Private Sub CboClmPending_Change()
    If InStr(CboClmPending.Text, "null") Then
       CboClmPending.Enabled = False
    End If
End Sub

Private Sub CboClmPending_KeyPress(KeyAscii As Integer)
Dim NewText As String
Dim ValidCount As Integer
Dim ValidValue As String
Dim i As Integer
    If KeyAscii >= 32 And KeyAscii <> 127 Then
    NewText = LCase(Left(CboClmPending.Text, CboClmPending.SelStart) + Chr(KeyAscii) + Mid(CboClmPending.Text, CboClmPending.SelStart + CboClmPending.SelLength + 1))
 
    ValidCount = 0
    ValidValue = ""
 
    For i = 0 To CboClmPending.ListCount - 1
 
        If NewText = LCase(Left(CboClmPending.List(i), Len(NewText))) Then
                ValidCount = ValidCount + 1
                ValidValue = CboClmPending.List(i)
        End If
        Next
        If ValidCount <= 1 Then KeyAscii = 0
            If ValidCount = 1 Then
                CboClmPending.Text = ValidValue
                CboClmPending.SelStart = 0
                CboClmPending.SelLength = Len(ValidValue)
            End If
        End If
End Sub

Private Sub CboStayType_Change()
    If InStr(CboStayType.Text, "null") Then
       CboStayType.Enabled = False
    End If
End Sub

Private Sub CboStayType_KeyPress(KeyAscii As Integer)
Dim NewText As String
Dim ValidCount As Integer
Dim ValidValue As String
Dim i As Integer
    If KeyAscii >= 32 And KeyAscii <> 127 Then
    NewText = LCase(Left(CboStayType.Text, CboStayType.SelStart) + Chr(KeyAscii) + Mid(CboStayType.Text, CboStayType.SelStart + CboStayType.SelLength + 1))
 
    ValidCount = 0
    ValidValue = ""
 
    For i = 0 To CboStayType.ListCount - 1
 
        If NewText = LCase(Left(CboStayType.List(i), Len(NewText))) Then
                ValidCount = ValidCount + 1
                ValidValue = CboStayType.List(i)
        End If
        Next
        If ValidCount <= 1 Then KeyAscii = 0
            If ValidCount = 1 Then
                CboStayType.Text = ValidValue
                CboStayType.SelStart = 0
                CboStayType.SelLength = Len(ValidValue)
            End If
        End If
End Sub

Private Sub CboCashType_Change()
    If InStr(CboCashType.Text, "null") Then
       CboCashType.Enabled = False
    End If
End Sub

Private Sub CboCashType_KeyPress(KeyAscii As Integer)
Dim NewText As String
Dim ValidCount As Integer
Dim ValidValue As String
Dim i As Integer
    If KeyAscii >= 32 And KeyAscii <> 127 Then
    NewText = LCase(Left(CboCashType.Text, CboCashType.SelStart) + Chr(KeyAscii) + Mid(CboCashType.Text, CboCashType.SelStart + CboCashType.SelLength + 1))
 
    ValidCount = 0
    ValidValue = ""
 
    For i = 0 To CboStayType.ListCount - 1
 
        If NewText = LCase(Left(CboCashType.List(i), Len(NewText))) Then
                ValidCount = ValidCount + 1
                ValidValue = CboCashType.List(i)
        End If
        Next
        If ValidCount <= 1 Then KeyAscii = 0
            If ValidCount = 1 Then
                CboCashType.Text = ValidValue
                CboCashType.SelStart = 0
                CboCashType.SelLength = Len(ValidValue)
            End If
        End If
End Sub

Private Sub CboModality_Change()
    If InStr(CboModality.Text, "null") Then
       CboModality.Enabled = False
    End If
End Sub

Private Sub CboModality_KeyPress(KeyAscii As Integer)
Dim NewText As String
Dim ValidCount As Integer
Dim ValidValue As String
Dim i As Integer
    If KeyAscii >= 32 And KeyAscii <> 127 Then
    NewText = LCase(Left(CboModality.Text, CboModality.SelStart) + Chr(KeyAscii) + Mid(CboModality.Text, CboModality.SelStart + CboModality.SelLength + 1))
 
    ValidCount = 0
    ValidValue = ""
 
    For i = 0 To CboModality.ListCount - 1
 
        If NewText = LCase(Left(CboModality.List(i), Len(NewText))) Then
                ValidCount = ValidCount + 1
                ValidValue = CboModality.List(i)
        End If
        Next
        If ValidCount <= 1 Then KeyAscii = 0
            If ValidCount = 1 Then
                CboModality.Text = ValidValue
                CboModality.SelStart = 0
                CboModality.SelLength = Len(ValidValue)
            End If
        End If
End Sub

Private Sub StateListing(Optional StrAppend As String)
Dim rst2 As New ADODB.Recordset
Dim cmd As New ADODB.Command
Dim prm1 As ADODB.Parameter
Dim prm2 As ADODB.Parameter
Dim strAppendCase As String
Dim StrAppend1 As String
Dim Record As ListItem
    
    lblCurrent = 0
    StateCode = ""
    Payor = ""
    ListView2.ListItems.Clear
    
    If cboClaimCase.ListIndex = 0 Then
        strAppendCase = "81"
    'sql = " SELECT PAYCode, STACode, COUNT(CASNumber) as TotalCases, SUM(clmtotalactual) AS totalActual, " & _
          " SUM(clmtotalapproved) As TotalApproved, " & _
          " SUM(NoofDay) AS NoofDay From VIEW_CLMReportDCU Where 0 = 0 "
    
    ElseIf cboClaimCase.ListIndex = 1 Then
        strAppendCase = "82"
    'sql = " SELECT PAYCode, STACode, COUNT(DISTINCT CASNumber) as TotalCases, SUM(clmtotalactual) AS totalActual, " & _
          " SUM(clmtotalapproved) As TotalApproved " & _
          " From VIEW_CLMReportDCU Where 0 = 0 "
          
    End If
    
    'sql2 = " Group By PAYCode, STACode "
    
    'If Len(Trim(StrAppend)) Then
    '    sql = sql + StrAppend + sql2
    'Else
    '    sql = sql + sql2
    'End If
            
    'rst2.Open sql, medixmasterconnWS, adOpenForwardOnly, adLockBatchOptimistic
    Set cmd.ActiveConnection = medixmasterconnWS
        cmd.CommandText = "DCUReport"
        cmd.CommandType = adCmdStoredProc
        cmd.CommandTimeout = 300
        Set prm1 = cmd.CreateParameter("StrAppend", adVarChar, adParamInput, 2000, StrAppend)
        cmd.Parameters.Append prm1
        Set prm2 = cmd.CreateParameter("StrAppendCase", adVarChar, adParamInput, 2000, strAppendCase)
        cmd.Parameters.Append prm2
        rst2.CursorLocation = adUseClient
        rst2.CursorType = adOpenForwardOnly
        rst2.CacheSize = 100
        Set rst2 = cmd.Execute
            
        Do Until rst2.EOF
        
            With rst2
                If Len(rst2!PAYCode) Then
                    Set Record = ListView2.ListItems.Add(, , rst2!PAYCode)
                    Payor = rst2!PAYCode
                Else
                    Set Record = ListView2.ListItems.Add(, , "Empty")
                    Payor = ""
                End If
                
                If Len(rst2!STACode) Then
                    Record.SubItems(1) = Trim(rst2!STACode)
                    StateCode = rst2!STACode
                Else
                    Record.SubItems(1) = "Empty"
                    StateCode = ""
                End If
                
                If Len(rst2!TotalCases) Then
                    Record.SubItems(2) = rst2!TotalCases
                End If
                
                If cboClaimCase.ListIndex = 0 Then
                
                    If rst2!NoofDay = "0" Then
                        Record.SubItems(3) = "1"
                    ElseIf Len(rst2!NoofDay) Then
                        Record.SubItems(3) = rst2!NoofDay
                    End If
                
                ElseIf cboClaimCase.ListIndex = 1 Then
                
                    If Payor = "" And StateCode = "" Then
                       strAppendCase = "83"
                       StrAppend1 = " AND (PAYCode = '' OR PAYCode IS NULL) AND " & _
                           " (STACode = '' OR STACode IS NULL)"
                       StrAppend1 = StrAppend1 + StrAppend
                   'sql3 = " SELECT PAYCode, STACode, SUM(DurationStay) AS NoofDay " & _
                           " From VIEW_CLMReportDCU2 Where (PAYCode = '' OR PAYCode IS NULL) AND " & _
                           " (STACode = '' OR STACode IS NULL) "
                           
                    ElseIf Payor <> "" And StateCode = "" Then
                       strAppendCase = "83"
                       StrAppend1 = " AND PAYCode = '" & Trim(Payor) & "' AND " & _
                           " (STACode = '' OR STACode IS NULL)"
                       StrAppend1 = StrAppend1 + StrAppend
                    'sql3 = " SELECT PAYCode, STACode, SUM(DurationStay) AS NoofDay " & _
                           " From VIEW_CLMReportDCU2 Where PAYCode = '" & Trim(Payor) & "' AND " & _
                           " (STACode = '' OR STACode IS NULL) "
                           
                    ElseIf Payor = "" And StateCode <> "" Then
                       strAppendCase = "83"
                       StrAppend1 = " AND (PAYCode = '' OR PAYCode IS NULL) AND " & _
                           " STACode = '" & Trim(StateCode) & "'"
                       StrAppend1 = StrAppend1 + StrAppend
                    'sql3 = " SELECT PAYCode, STACode, SUM(DurationStay) AS NoofDay " & _
                           " From VIEW_CLMReportDCU2 Where (PAYCode = '' OR PAYCode IS NULL) AND " & _
                           " STACode = '" & Trim(StateCode) & "'"
                           
                    Else
                       strAppendCase = "83"
                       StrAppend1 = " AND PAYCode = '" & Trim(Payor) & "' AND " & _
                           " STACode = '" & Trim(StateCode) & "'"
                       StrAppend1 = StrAppend1 + StrAppend
                    'sql3 = " SELECT PAYCode, STACode, SUM(DurationStay) AS NoofDay " & _
                           " From VIEW_CLMReportDCU2 Where PAYCode = '" & Trim(Payor) & "' AND " & _
                           " STACode = '" & Trim(StateCode) & "'"
                    
                    End If
                       
                       'sql2 = " Group By PAYCode, STACode  "
                       
                       'If Len(Trim(StrAppend)) Then
                       '     sql3 = sql3 + StrAppend + sql2
                       ' Else
                       '     sql3 = sql3 + sql2
                       ' End If
                                              
                       'rst1.Open sql3, medixmasterconnWS, adOpenForwardOnly, adLockBatchOptimistic
                        Call DurationStay(StrAppend1, strAppendCase, Record)
                End If
                
                If Len(rst2!TotalActual) Then
                    Record.SubItems(4) = Format(rst2!TotalActual, "#,##0.00")
                End If
                
                If Len(rst2!TotalApproved) Then
                    Record.SubItems(5) = Format(rst2!TotalApproved, "#,##0.00")
                End If
                
            End With
            lblCurrent = lblCurrent + 1
            rst2.MoveNext
       Loop
    If ListView2.ListItems.Count = 0 Then
        MsgBox "No record found!", vbOKOnly + vbCritical, DialogBoxTitle
        CmdPrint.Enabled = True
        CmdClear.Enabled = True
        cmdExit.Enabled = True
    End If
    rst2.Close
    Set rst2 = Nothing
    CmdPrint.Enabled = True
    CmdClear.Enabled = True
    cmdExit.Enabled = True
End Sub

Private Sub ListViewHeader5()
If cboClaimCase.ListIndex = 0 Then
    ListView2.ColumnHeaders(1).Text = "Payor"
    ListView2.ColumnHeaders(2).Text = "State"
    ListView2.ColumnHeaders(2).Alignment = lvwColumnLeft
    ListView2.ColumnHeaders(3).Text = "Total Claims"
    ListView2.ColumnHeaders(3).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(4).Text = "Total Duration"
    ListView2.ColumnHeaders(4).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(5).Text = "Total Actual Amt"
    ListView2.ColumnHeaders(5).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(6).Text = "Total Approved Amt"
    ListView2.ColumnHeaders(6).Alignment = lvwColumnRight
    
    For AA = 7 To 63
        ListView2.ColumnHeaders(AA).Text = ""
    Next AA
ElseIf cboClaimCase.ListIndex = 1 Then
    ListView2.ColumnHeaders(1).Text = "Payor"
    ListView2.ColumnHeaders(2).Text = "State"
    ListView2.ColumnHeaders(2).Alignment = lvwColumnLeft
    ListView2.ColumnHeaders(3).Text = "Total Cases"
    ListView2.ColumnHeaders(3).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(4).Text = "Total Duration"
    ListView2.ColumnHeaders(4).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(5).Text = "Total Actual Amt"
    ListView2.ColumnHeaders(5).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(6).Text = "Total Approved Amt"
    ListView2.ColumnHeaders(6).Alignment = lvwColumnRight
    
    For AA = 7 To 63
        ListView2.ColumnHeaders(AA).Text = ""
    Next AA
End If
End Sub

Private Sub ListViewHeader7()

    'If cboClaimCase.ListIndex = 0 Then
    '    Option16.Caption = "By Bordx Month(Payor)"
    '    Option30.Caption = "By Bordx Month"
    'End If

If cboClaimCase.ListIndex = 0 And opt35 = True Then
    ListView2.ColumnHeaders(1).Text = "Reg. Month"
    ListView2.ColumnHeaders(2).Text = "Total Claims"
    ListView2.ColumnHeaders(2).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(3).Text = "Total Duration"
    ListView2.ColumnHeaders(3).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(4).Text = "Total Actual Amt"
    ListView2.ColumnHeaders(4).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(5).Text = "Total Approved Amt"
    ListView2.ColumnHeaders(5).Alignment = lvwColumnRight
    
    For AA = 6 To 63
        ListView2.ColumnHeaders(AA).Text = ""
    Next AA
    
ElseIf cboClaimCase.ListIndex = 1 And opt35 = True Then
    ListView2.ColumnHeaders(1).Text = "Reg. Month"
    ListView2.ColumnHeaders(2).Text = "Total Cases"
    ListView2.ColumnHeaders(2).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(3).Text = "Total Duration"
    ListView2.ColumnHeaders(3).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(4).Text = "Total Actual Amt"
    ListView2.ColumnHeaders(4).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(5).Text = "Total Approved Amt"
    ListView2.ColumnHeaders(5).Alignment = lvwColumnRight
    
    For AA = 6 To 63
        ListView2.ColumnHeaders(AA).Text = ""
    Next AA
    
ElseIf cboClaimCase.ListIndex = 0 And Opt8 = True Then
    ListView2.ColumnHeaders(1).Text = "Payor"
    ListView2.ColumnHeaders(2).Text = "Reg. Month"
    ListView2.ColumnHeaders(2).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(3).Text = "Total Claims"
    ListView2.ColumnHeaders(3).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(4).Text = "Total Duration"
    ListView2.ColumnHeaders(4).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(5).Text = "Total Actual Amt"
    ListView2.ColumnHeaders(5).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(6).Text = "Total Approved Amt"
    ListView2.ColumnHeaders(6).Alignment = lvwColumnRight
    
    For AA = 7 To 63
        ListView2.ColumnHeaders(AA).Text = ""
    Next AA
    
ElseIf cboClaimCase.ListIndex = 1 And Opt8 = True Then
    ListView2.ColumnHeaders(1).Text = "Payor"
    ListView2.ColumnHeaders(2).Text = "Reg. Month"
    ListView2.ColumnHeaders(2).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(3).Text = "Total Cases"
    ListView2.ColumnHeaders(3).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(4).Text = "Total Duration"
    ListView2.ColumnHeaders(4).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(5).Text = "Total Actual Amt"
    ListView2.ColumnHeaders(5).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(6).Text = "Total Approved Amt"
    ListView2.ColumnHeaders(6).Alignment = lvwColumnRight
    
    For AA = 7 To 63
        ListView2.ColumnHeaders(AA).Text = ""
    Next AA
    
ElseIf cboClaimCase.ListIndex = 0 And opt34 = True Then
    ListView2.ColumnHeaders(1).Text = "Eff. Month"
    ListView2.ColumnHeaders(2).Text = "Total Claims"
    ListView2.ColumnHeaders(2).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(3).Text = "Total Duration"
    ListView2.ColumnHeaders(3).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(4).Text = "Total Actual Amt"
    ListView2.ColumnHeaders(4).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(5).Text = "Total Approved Amt"
    ListView2.ColumnHeaders(5).Alignment = lvwColumnRight
    
    For AA = 6 To 63
        ListView2.ColumnHeaders(AA).Text = ""
    Next AA
    
ElseIf cboClaimCase.ListIndex = 1 And opt34 = True Then
    ListView2.ColumnHeaders(1).Text = "Eff. Month"
    ListView2.ColumnHeaders(2).Text = "Total Cases"
    ListView2.ColumnHeaders(2).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(3).Text = "Total Duration"
    ListView2.ColumnHeaders(3).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(4).Text = "Total Actual Amt"
    ListView2.ColumnHeaders(4).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(5).Text = "Total Approved Amt"
    ListView2.ColumnHeaders(5).Alignment = lvwColumnRight
    
    For AA = 6 To 63
        ListView2.ColumnHeaders(AA).Text = ""
    Next AA
    
ElseIf cboClaimCase.ListIndex = 0 And Opt11 = True Then
    ListView2.ColumnHeaders(1).Text = "Payor"
    ListView2.ColumnHeaders(2).Text = "Eff. Month"
    ListView2.ColumnHeaders(2).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(3).Text = "Total Claims"
    ListView2.ColumnHeaders(3).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(4).Text = "Total Duration"
    ListView2.ColumnHeaders(4).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(5).Text = "Total Actual Amt"
    ListView2.ColumnHeaders(5).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(6).Text = "Total Approved Amt"
    ListView2.ColumnHeaders(6).Alignment = lvwColumnRight
    
    For AA = 7 To 63
        ListView2.ColumnHeaders(AA).Text = ""
    Next AA
    
ElseIf cboClaimCase.ListIndex = 1 And Opt11 = True Then
    ListView2.ColumnHeaders(1).Text = "Payor"
    ListView2.ColumnHeaders(2).Text = "Eff. Month"
    ListView2.ColumnHeaders(2).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(3).Text = "Total Cases"
    ListView2.ColumnHeaders(3).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(4).Text = "Total Duration"
    ListView2.ColumnHeaders(4).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(5).Text = "Total Actual Amt"
    ListView2.ColumnHeaders(5).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(6).Text = "Total Approved Amt"
    ListView2.ColumnHeaders(6).Alignment = lvwColumnRight
    
    For AA = 7 To 63
        ListView2.ColumnHeaders(AA).Text = ""
    Next AA
    
ElseIf cboClaimCase.ListIndex = 0 And opt33 = True Then
    ListView2.ColumnHeaders(1).Text = "Exp. Month"
    ListView2.ColumnHeaders(2).Text = "Total Claims"
    ListView2.ColumnHeaders(2).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(3).Text = "Total Duration"
    ListView2.ColumnHeaders(3).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(4).Text = "Total Actual Amt"
    ListView2.ColumnHeaders(4).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(5).Text = "Total Approved Amt"
    ListView2.ColumnHeaders(5).Alignment = lvwColumnRight
    
    For AA = 6 To 63
        ListView2.ColumnHeaders(AA).Text = ""
    Next AA
    
ElseIf cboClaimCase.ListIndex = 1 And opt33 = True Then
    ListView2.ColumnHeaders(1).Text = "Exp. Month"
    ListView2.ColumnHeaders(2).Text = "Total Cases"
    ListView2.ColumnHeaders(2).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(3).Text = "Total Duration"
    ListView2.ColumnHeaders(3).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(4).Text = "Total Actual Amt"
    ListView2.ColumnHeaders(4).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(5).Text = "Total Approved Amt"
    ListView2.ColumnHeaders(5).Alignment = lvwColumnRight
    
    For AA = 6 To 63
        ListView2.ColumnHeaders(AA).Text = ""
    Next AA
    
ElseIf cboClaimCase.ListIndex = 0 And opt12 = True Then
    ListView2.ColumnHeaders(1).Text = "Payor"
    ListView2.ColumnHeaders(2).Text = "Exp. Month"
    ListView2.ColumnHeaders(2).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(3).Text = "Total Claims"
    ListView2.ColumnHeaders(3).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(4).Text = "Total Duration"
    ListView2.ColumnHeaders(4).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(5).Text = "Total Actual Amt"
    ListView2.ColumnHeaders(5).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(6).Text = "Total Approved Amt"
    ListView2.ColumnHeaders(6).Alignment = lvwColumnRight
    
    For AA = 7 To 63
        ListView2.ColumnHeaders(AA).Text = ""
    Next AA
    
ElseIf cboClaimCase.ListIndex = 1 And opt12 = True Then
    ListView2.ColumnHeaders(1).Text = "Payor"
    ListView2.ColumnHeaders(2).Text = "Exp. Month"
    ListView2.ColumnHeaders(2).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(3).Text = "Total Cases"
    ListView2.ColumnHeaders(3).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(4).Text = "Total Duration"
    ListView2.ColumnHeaders(4).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(5).Text = "Total Actual Amt"
    ListView2.ColumnHeaders(5).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(6).Text = "Total Approved Amt"
    ListView2.ColumnHeaders(6).Alignment = lvwColumnRight
    
    For AA = 7 To 63
        ListView2.ColumnHeaders(AA).Text = ""
    Next AA

ElseIf cboClaimCase.ListIndex = 0 And opt29 = True Then
    ListView2.ColumnHeaders(1).Text = "DOA Month"
    ListView2.ColumnHeaders(2).Text = "Total Claims"
    ListView2.ColumnHeaders(2).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(3).Text = "Total Duration"
    ListView2.ColumnHeaders(3).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(4).Text = "Total Actual Amt"
    ListView2.ColumnHeaders(4).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(5).Text = "Total Approved Amt"
    ListView2.ColumnHeaders(5).Alignment = lvwColumnRight
    
    For AA = 6 To 63
        ListView2.ColumnHeaders(AA).Text = ""
    Next AA
    
ElseIf cboClaimCase.ListIndex = 1 And opt29 = True Then
    ListView2.ColumnHeaders(1).Text = "DOA Month"
    ListView2.ColumnHeaders(2).Text = "Total Cases"
    ListView2.ColumnHeaders(2).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(3).Text = "Total Duration"
    ListView2.ColumnHeaders(3).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(4).Text = "Total Actual Amt"
    ListView2.ColumnHeaders(4).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(5).Text = "Total Approved Amt"
    ListView2.ColumnHeaders(5).Alignment = lvwColumnRight
    
    For AA = 6 To 63
        ListView2.ColumnHeaders(AA).Text = ""
    Next AA
    
ElseIf cboClaimCase.ListIndex = 0 And opt13 = True Then
    ListView2.ColumnHeaders(1).Text = "Payor"
    ListView2.ColumnHeaders(2).Text = "DOA Month"
    ListView2.ColumnHeaders(2).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(3).Text = "Total Claims"
    ListView2.ColumnHeaders(3).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(4).Text = "Total Duration"
    ListView2.ColumnHeaders(4).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(5).Text = "Total Actual Amt"
    ListView2.ColumnHeaders(5).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(6).Text = "Total Approved Amt"
    ListView2.ColumnHeaders(6).Alignment = lvwColumnRight
    
    For AA = 7 To 63
        ListView2.ColumnHeaders(AA).Text = ""
    Next AA
    
ElseIf cboClaimCase.ListIndex = 1 And opt13 = True Then
    ListView2.ColumnHeaders(1).Text = "Payor"
    ListView2.ColumnHeaders(2).Text = "DOA Month"
    ListView2.ColumnHeaders(2).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(3).Text = "Total Cases"
    ListView2.ColumnHeaders(3).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(4).Text = "Total Duration"
    ListView2.ColumnHeaders(4).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(5).Text = "Total Actual Amt"
    ListView2.ColumnHeaders(5).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(6).Text = "Total Approved Amt"
    ListView2.ColumnHeaders(6).Alignment = lvwColumnRight
    
    For AA = 7 To 63
        ListView2.ColumnHeaders(AA).Text = ""
    Next AA
    
ElseIf cboClaimCase.ListIndex = 0 And opt37 = True Then
    ListView2.ColumnHeaders(1).Text = "DOD Month"
    ListView2.ColumnHeaders(2).Text = "Total Claims"
    ListView2.ColumnHeaders(2).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(3).Text = "Total Duration"
    ListView2.ColumnHeaders(3).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(4).Text = "Total Actual Amt"
    ListView2.ColumnHeaders(4).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(5).Text = "Total Approved Amt"
    ListView2.ColumnHeaders(5).Alignment = lvwColumnRight
    
    For AA = 6 To 63
        ListView2.ColumnHeaders(AA).Text = ""
    Next AA
    
ElseIf cboClaimCase.ListIndex = 1 And opt37 = True Then
    ListView2.ColumnHeaders(1).Text = "DOD Month"
    ListView2.ColumnHeaders(2).Text = "Total Cases"
    ListView2.ColumnHeaders(2).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(3).Text = "Total Duration"
    ListView2.ColumnHeaders(3).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(4).Text = "Total Actual Amt"
    ListView2.ColumnHeaders(4).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(5).Text = "Total Approved Amt"
    ListView2.ColumnHeaders(5).Alignment = lvwColumnRight
    
    For AA = 6 To 63
        ListView2.ColumnHeaders(AA).Text = ""
    Next AA
    
ElseIf cboClaimCase.ListIndex = 0 And Opt7 = True Then
    ListView2.ColumnHeaders(1).Text = "Payor"
    ListView2.ColumnHeaders(2).Text = "DOD Month"
    ListView2.ColumnHeaders(2).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(3).Text = "Total Claims"
    ListView2.ColumnHeaders(3).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(4).Text = "Total Duration"
    ListView2.ColumnHeaders(4).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(5).Text = "Total Actual Amt"
    ListView2.ColumnHeaders(5).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(6).Text = "Total Approved Amt"
    ListView2.ColumnHeaders(6).Alignment = lvwColumnRight
    
    For AA = 7 To 63
        ListView2.ColumnHeaders(AA).Text = ""
    Next AA
    
ElseIf cboClaimCase.ListIndex = 1 And Opt7 = True Then
    ListView2.ColumnHeaders(1).Text = "Payor"
    ListView2.ColumnHeaders(2).Text = "DOD Month"
    ListView2.ColumnHeaders(2).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(3).Text = "Total Cases"
    ListView2.ColumnHeaders(3).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(4).Text = "Total Duration"
    ListView2.ColumnHeaders(4).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(5).Text = "Total Actual Amt"
    ListView2.ColumnHeaders(5).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(6).Text = "Total Approved Amt"
    ListView2.ColumnHeaders(6).Alignment = lvwColumnRight
    
    For AA = 7 To 63
        ListView2.ColumnHeaders(AA).Text = ""
    Next AA
    
ElseIf cboClaimCase.ListIndex = 0 And opt30 = True Then
    ListView2.ColumnHeaders(1).Text = "Bordx Month"
    ListView2.ColumnHeaders(2).Text = "Total Claims"
    ListView2.ColumnHeaders(2).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(3).Text = "Total Duration"
    ListView2.ColumnHeaders(3).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(4).Text = "Total Actual Amt"
    ListView2.ColumnHeaders(4).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(5).Text = "Total Approved Amt"
    ListView2.ColumnHeaders(5).Alignment = lvwColumnRight
    
    For AA = 6 To 63
        ListView2.ColumnHeaders(AA).Text = ""
    Next AA
    
ElseIf cboClaimCase.ListIndex = 1 And opt30 = True Then
    ListView2.ColumnHeaders(1).Text = "Bordx Month"
    ListView2.ColumnHeaders(2).Text = "Total Cases"
    ListView2.ColumnHeaders(2).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(3).Text = "Total Duration"
    ListView2.ColumnHeaders(3).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(4).Text = "Total Actual Amt"
    ListView2.ColumnHeaders(4).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(5).Text = "Total Approved Amt"
    ListView2.ColumnHeaders(5).Alignment = lvwColumnRight
    
    For AA = 6 To 63
        ListView2.ColumnHeaders(AA).Text = ""
    Next AA
    
ElseIf cboClaimCase.ListIndex = 0 And opt16 = True Then
    ListView2.ColumnHeaders(1).Text = "Payor"
    ListView2.ColumnHeaders(2).Text = "Bordx Month"
    ListView2.ColumnHeaders(2).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(3).Text = "Total Claims"
    ListView2.ColumnHeaders(3).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(4).Text = "Total Duration"
    ListView2.ColumnHeaders(4).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(5).Text = "Total Actual Amt"
    ListView2.ColumnHeaders(5).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(6).Text = "Total Approved Amt"
    ListView2.ColumnHeaders(6).Alignment = lvwColumnRight
    
    For AA = 7 To 63
        ListView2.ColumnHeaders(AA).Text = ""
    Next AA
    
ElseIf cboClaimCase.ListIndex = 1 And opt16 = True Then
    ListView2.ColumnHeaders(1).Text = "Payor"
    ListView2.ColumnHeaders(2).Text = "Bordx Month"
    ListView2.ColumnHeaders(2).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(3).Text = "Total Cases"
    ListView2.ColumnHeaders(3).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(4).Text = "Total Duration"
    ListView2.ColumnHeaders(4).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(5).Text = "Total Actual Amt"
    ListView2.ColumnHeaders(5).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(6).Text = "Total Approved Amt"
    ListView2.ColumnHeaders(6).Alignment = lvwColumnRight
    
    For AA = 7 To 63
        ListView2.ColumnHeaders(AA).Text = ""
    Next AA
    
End If
End Sub

Private Sub RegMonthPayor(Optional StrAppend As String)
Dim rst2 As New ADODB.Recordset
Dim cmd As New ADODB.Command
Dim prm1 As ADODB.Parameter
Dim prm2 As ADODB.Parameter
Dim strAppendCase As String
Dim StrAppend1 As String
Dim Record As ListItem
Dim Total As String
Dim Current As String

    Current = 0
    Month = ""
    Month2 = ""
    Year = ""
    ListView2.ListItems.Clear

    If cboClaimCase.ListIndex = 0 Then
        strAppendCase = "87"
        'sql = " SELECT PAYCode, MONTH(CaseRegDate) AS MONTH, YEAR(CaseRegDate) AS YEAR, COUNT(CASNumber) AS TotalCases, " & _
              " SUM(CLMTotalActual) AS TotalActual, SUM(CLMTotalApproved) As TotalApproved, SUM(NoofDay) AS NoofDay " & _
              " From VIEW_CLMReportDCU where 0 = 0 "
          
    ElseIf cboClaimCase.ListIndex = 1 Then
        strAppendCase = "88"
        'sql = " SELECT PAYCode, MONTH(CaseRegDate) AS MONTH, YEAR(CaseRegDate) AS YEAR, COUNT(Distinct CASNumber) AS TotalCases, " & _
              " SUM(CLMTotalActual) AS TotalActual, SUM(CLMTotalApproved) As TotalApproved, SUM(NoofDay) AS NoofDay " & _
              " From VIEW_CLMReportDCU where 0 = 0 "
          
    End If
    
    'sql2 = " GROUP BY PAYCode, MONTH(CaseRegDate), YEAR(CaseRegDate)"
    
    'If Len(Trim(StrAppend)) Then
    '    sql = sql + StrAppend + sql2
    'Else
    '    sql = sql + sql2
    'End If

    'rst2.Open sql, medixmasterconnWS, adOpenForwardOnly, adLockBatchOptimistic
    Set cmd.ActiveConnection = medixmasterconnWS
        cmd.CommandText = "DCUReport"
        cmd.CommandType = adCmdStoredProc
        cmd.CommandTimeout = 300
        Set prm1 = cmd.CreateParameter("StrAppend", adVarChar, adParamInput, 2000, StrAppend)
        cmd.Parameters.Append prm1
        Set prm2 = cmd.CreateParameter("StrAppendCase", adVarChar, adParamInput, 2000, strAppendCase)
        cmd.Parameters.Append prm2
        rst2.CursorLocation = adUseClient
        rst2.CursorType = adOpenForwardOnly
        rst2.CacheSize = 100
        Set rst2 = cmd.Execute
    
    Do
        Do Until rst2.EOF

            With rst2
                
                If Len(!PAYCode) Then
                    Set Record = ListView2.ListItems.Add(, , !PAYCode)
                    Payor = !PAYCode
                Else
                    Set Record = ListView2.ListItems.Add(, , "Empty")
                    Payor = ""
                End If
                
                If Len(!Month & "/" & !Year) Then
                    If Len(!Month) = 1 Then
                        Month = "0" & !Month
                        Record.SubItems(1) = (!Year & "/" & Month)
                        Month2 = Month
                        Year = !Year
                    Else
                        Record.SubItems(1) = (!Year & "/" & !Month)
                        Month2 = !Month
                        Year = !Year
                    End If
                  
                End If
                
                If Len(!TotalCases) Then
                    Record.SubItems(2) = !TotalCases
                End If
                
                If cboClaimCase.ListIndex = 0 Then
                
                    If rst2!NoofDay = "0" Then
                        Record.SubItems(3) = "1"
                    ElseIf Len(rst2!NoofDay) Then
                        Record.SubItems(3) = rst2!NoofDay
                    End If
                
                ElseIf cboClaimCase.ListIndex = 1 Then
                                    
                       If Payor = "" Then
                            strAppendCase = "89"
                            StrAppend1 = " AND (PAYCode IS NULL OR PAYCode = '') AND MONTH(CaseRegDate) = '" & Trim(Month2) & "' AND " & _
                                   " Year(CaseRegDate) = '" & Trim(Year) & "'"
                            StrAppend1 = StrAppend1 + StrAppend
                            'sql3 = " SELECT MONTH(CaseRegDate) AS MONTH, YEAR(CaseRegDate) AS YEAR, SUM(DurationStay) AS NoofDay " & _
                                   " From VIEW_CLMReportDCU2 Where (PAYCode IS NULL OR PAYCode = '') AND MONTH(CaseRegDate) = '" & Trim(Month2) & "' AND " & _
                                   " Year(CaseRegDate) = '" & Trim(Year) & "'"
                       Else
                            strAppendCase = "89"
                            StrAppend1 = " AND PAYCode = '" & Trim(Payor) & "' AND MONTH(CaseRegDate) = '" & Trim(Month2) & "' AND " & _
                                   " Year(CaseRegDate) = '" & Trim(Year) & "'"
                            StrAppend1 = StrAppend1 + StrAppend
                            'sql3 = " SELECT MONTH(CaseRegDate) AS MONTH, YEAR(CaseRegDate) AS YEAR, SUM(DurationStay) AS NoofDay " & _
                                   " From VIEW_CLMReportDCU2 Where PAYCode = '" & Trim(Payor) & "' AND MONTH(CaseRegDate) = '" & Trim(Month2) & "' AND " & _
                                   " Year(CaseRegDate) = '" & Trim(Year) & "'"
                       End If
                                       
                       'sql2 = " GROUP BY PAYCode, MONTH(CaseRegDate), YEAR(CaseRegDate)"
                       
                       'If Len(Trim(StrAppend)) Then
                       '     sql3 = sql3 + StrAppend + sql2
                       'Else
                       '     sql3 = sql3 + sql2
                       'End If
                                              
                       'rst1.Open sql3, medixmasterconnWS, adOpenForwardOnly, adLockBatchOptimistic
                        Call DurationStay(StrAppend1, strAppendCase, Record)
                End If
                
                If Len(!TotalActual) Then
                    Record.SubItems(4) = Format(!TotalActual, "#,##0.00")
                Else
                    Record.SubItems(4) = "0.00"
                End If
                
                If Len(!TotalApproved) Then
                    Record.SubItems(5) = Format(!TotalApproved, "#,##0.00")
                Else
                    Record.SubItems(5) = "0.00"
                End If
                                
                Current = Current + 1
                lblCurrent = Current
            End With
            rst2.MoveNext
            DoEvents
                If intExit = 1 Then
                   Check = False
                   Exit Do
                End If
        Loop
    Loop Until Check = False
    
    If ListView2.ListItems.Count = 0 Then
        MsgBox "No record found!", vbOKOnly + vbCritical, DialogBoxTitle
        CmdClear.Enabled = True
        cmdExit.Enabled = True
        CmdPrint.Enabled = True
        lblCurrent = 0
    End If
    
    intExit = 0
    rst2.Close
    Set rst2 = Nothing
    CmdClear.Enabled = True
    cmdExit.Enabled = True
    CmdPrint.Enabled = True
End Sub

Private Sub RegMonthAllPayor(Optional StrAppend As String)
Dim rst2 As New ADODB.Recordset
Dim cmd As New ADODB.Command
Dim prm1 As ADODB.Parameter
Dim prm2 As ADODB.Parameter
Dim strAppendCase As String
Dim StrAppend1 As String
Dim Record As ListItem
Dim Total As String
Dim Current As String

    Current = 0
    Month = ""
    Month2 = ""
    Year = ""
    ListView2.ListItems.Clear

    If cboClaimCase.ListIndex = 0 Then
        strAppendCase = "84"
        'sql = " SELECT MONTH(CaseRegDate) AS MONTH, YEAR(CaseRegDate) AS YEAR, COUNT(CASNumber) AS TotalCases, " & _
              " SUM(CLMTotalActual) AS TotalActual, SUM(CLMTotalApproved) As TotalApproved, SUM(NoofDay) AS NoofDay " & _
              " From VIEW_CLMReportDCU where 0 = 0 "
          
    ElseIf cboClaimCase.ListIndex = 1 Then
        strAppendCase = "85"
        'sql = " SELECT MONTH(CaseRegDate) AS MONTH, YEAR(CaseRegDate) AS YEAR, COUNT(Distinct CASNumber) AS TotalCases, " & _
              " SUM(CLMTotalActual) AS TotalActual, SUM(CLMTotalApproved) As TotalApproved, SUM(NoofDay) AS NoofDay " & _
              " From VIEW_CLMReportDCU where 0 = 0 "
          
    End If
    
    'sql2 = " GROUP BY MONTH(CaseRegDate), YEAR(CaseRegDate)"
    
    'If Len(Trim(StrAppend)) Then
    '    sql = sql + StrAppend + sql2
    'Else
    '    sql = sql + sql2
    'End If

    'rst2.Open sql, medixmasterconnWS, adOpenForwardOnly, adLockBatchOptimistic
    Set cmd.ActiveConnection = medixmasterconnWS
        cmd.CommandText = "DCUReport"
        cmd.CommandType = adCmdStoredProc
        cmd.CommandTimeout = 300
        Set prm1 = cmd.CreateParameter("StrAppend", adVarChar, adParamInput, 2000, StrAppend)
        cmd.Parameters.Append prm1
        Set prm2 = cmd.CreateParameter("StrAppendCase", adVarChar, adParamInput, 2000, strAppendCase)
        cmd.Parameters.Append prm2
        rst2.CursorLocation = adUseClient
        rst2.CursorType = adOpenForwardOnly
        rst2.CacheSize = 100
        Set rst2 = cmd.Execute
    
    Do
        Do Until rst2.EOF

            With rst2
                
                If Len(!Month & "/" & !Year) Then
                    If Len(!Month) = 1 Then
                        Month = "0" & !Month
                        Set Record = ListView2.ListItems.Add(, , !Year & "/" & Month)
                        Month2 = Month
                        Year = !Year
                    Else
                        Set Record = ListView2.ListItems.Add(, , !Year & "/" & !Month)
                        Month2 = !Month
                        Year = !Year
                    End If
                  
                End If
                
                If Len(!TotalCases) Then
                    Record.SubItems(1) = !TotalCases
                End If
                
                If cboClaimCase.ListIndex = 0 Then
                
                    If rst2!NoofDay = "0" Then
                        Record.SubItems(2) = "1"
                    ElseIf Len(rst2!NoofDay) Then
                        Record.SubItems(2) = rst2!NoofDay
                    End If
                
                ElseIf cboClaimCase.ListIndex = 1 Then
                        strAppendCase = "86"
                        StrAppend1 = " AND MONTH(CaseRegDate) = '" & Trim(Month2) & "' AND " & _
                               " Year(CaseRegDate) = '" & Trim(Year) & "'"
                        StrAppend1 = StrAppend1 + StrAppend
                       'sql3 = " SELECT MONTH(CaseRegDate) AS MONTH, YEAR(CaseRegDate) AS YEAR, SUM(DurationStay) AS NoofDay " & _
                               " From VIEW_CLMReportDCU2 Where MONTH(CaseRegDate) = '" & Trim(Month2) & "' AND " & _
                               " Year(CaseRegDate) = '" & Trim(Year) & "'"
                                       
                       'sql2 = " GROUP BY MONTH(CaseRegDate), YEAR(CaseRegDate)"
                       
                       'If Len(Trim(StrAppend)) Then
                       '     sql3 = sql3 + StrAppend + sql2
                       'Else
                       '     sql3 = sql3 + sql2
                       'End If
                                              
                       'rst1.Open sql3, medixmasterconnWS, adOpenForwardOnly, adLockBatchOptimistic
                        Call DurationStay(StrAppend1, strAppendCase, Record)
                End If
                
                If Len(!TotalActual) Then
                    Record.SubItems(3) = Format(!TotalActual, "#,##0.00")
                Else
                    Record.SubItems(3) = "0.00"
                End If
                
                If Len(!TotalApproved) Then
                    Record.SubItems(4) = Format(!TotalApproved, "#,##0.00")
                Else
                    Record.SubItems(4) = "0.00"
                End If
                                
                Current = Current + 1
                lblCurrent = Current
            End With
            rst2.MoveNext
            DoEvents
                If intExit = 1 Then
                   Check = False
                   Exit Do
                End If
        Loop
    Loop Until Check = False
    
    If ListView2.ListItems.Count = 0 Then
        MsgBox "No record found!", vbOKOnly + vbCritical, DialogBoxTitle
        CmdClear.Enabled = True
        cmdExit.Enabled = True
        CmdPrint.Enabled = True
        lblCurrent = 0
    End If
    
    intExit = 0
    rst2.Close
    Set rst2 = Nothing
    CmdClear.Enabled = True
    cmdExit.Enabled = True
    CmdPrint.Enabled = True
    
End Sub

Private Sub EffMonthPayor(Optional StrAppend As String)
Dim rst2 As New ADODB.Recordset
Dim cmd As New ADODB.Command
Dim prm1 As ADODB.Parameter
Dim prm2 As ADODB.Parameter
Dim strAppendCase As String
Dim StrAppend1 As String
Dim Record As ListItem
Dim Total As String
Dim Current As String

    Current = 0
    Month = ""
    Month2 = ""
    Year = ""
    ListView2.ListItems.Clear

    If cboClaimCase.ListIndex = 0 Then
        strAppendCase = "93"
        'sql = " SELECT PAYCode, MONTH(MBMPayorEffDate) AS MONTH, YEAR(MBMPayorEffDate) AS YEAR, COUNT(CASNumber) AS TotalCases, " & _
              " SUM(CLMTotalActual) AS TotalActual, SUM(CLMTotalApproved) As TotalApproved, SUM(NoofDay) AS NoofDay " & _
              " From VIEW_CLMReportDCU where 0 = 0 "
          
    ElseIf cboClaimCase.ListIndex = 1 Then
        strAppendCase = "94"
        'sql = " SELECT PAYCode, MONTH(MBMPayorEffDate) AS MONTH, YEAR(MBMPayorEffDate) AS YEAR, COUNT(Distinct CASNumber) AS TotalCases, " & _
              " SUM(CLMTotalActual) AS TotalActual, SUM(CLMTotalApproved) As TotalApproved, SUM(NoofDay) AS NoofDay " & _
              " From VIEW_CLMReportDCU where 0 = 0 "
          
    End If
    
    'sql2 = " GROUP BY PAYCode, MONTH(MBMPayorEffDate), YEAR(MBMPayorEffDate)"
    
    'If Len(Trim(StrAppend)) Then
    '    sql = sql + StrAppend + sql2
    'Else
    '    sql = sql + sql2
    'End If

    'rst2.Open sql, medixmasterconnWS, adOpenForwardOnly, adLockBatchOptimistic
    Set cmd.ActiveConnection = medixmasterconnWS
        cmd.CommandText = "DCUReport"
        cmd.CommandType = adCmdStoredProc
        cmd.CommandTimeout = 300
        Set prm1 = cmd.CreateParameter("StrAppend", adVarChar, adParamInput, 2000, StrAppend)
        cmd.Parameters.Append prm1
        Set prm2 = cmd.CreateParameter("StrAppendCase", adVarChar, adParamInput, 2000, strAppendCase)
        cmd.Parameters.Append prm2
        rst2.CursorLocation = adUseClient
        rst2.CursorType = adOpenForwardOnly
        rst2.CacheSize = 100
        Set rst2 = cmd.Execute
            
    Do
        Do Until rst2.EOF

            With rst2
                
                If Len(!PAYCode) Then
                    Set Record = ListView2.ListItems.Add(, , !PAYCode)
                    Payor = !PAYCode
                Else
                    Set Record = ListView2.ListItems.Add(, , "Empty")
                    Payor = ""
                End If
                
                If Len(!Month & "/" & !Year) Then
                    If Len(!Month) = 1 Then
                        Month = "0" & !Month
                        Record.SubItems(1) = (!Year & "/" & Month)
                        Month2 = Month
                        Year = !Year
                    Else
                        Record.SubItems(1) = (!Year & "/" & !Month)
                        Month2 = !Month
                        Year = !Year
                    End If
                  
                End If
                
                If Len(!TotalCases) Then
                    Record.SubItems(2) = !TotalCases
                End If
                
                If cboClaimCase.ListIndex = 0 Then
                
                    If rst2!NoofDay = "0" Then
                        Record.SubItems(3) = "1"
                    ElseIf Len(rst2!NoofDay) Then
                        Record.SubItems(3) = rst2!NoofDay
                    End If
                
                ElseIf cboClaimCase.ListIndex = 1 Then
                                    
                       If Payor = "" Then
                            strAppendCase = "95"
                            StrAppend1 = " AND (PAYCode IS NULL OR PAYCode = '') AND MONTH(MBMPayorEffDate) = '" & Trim(Month2) & "' AND " & _
                                   " Year(MBMPayorEffDate) = '" & Trim(Year) & "'"
                            StrAppend1 = StrAppend1 + StrAppend
                            'sql3 = " SELECT MONTH(MBMPayorEffDate) AS MONTH, YEAR(MBMPayorEffDate) AS YEAR, SUM(DurationStay) AS NoofDay " & _
                                   " From VIEW_CLMReportDCU2 Where (PAYCode IS NULL OR PAYCode = '') AND MONTH(MBMPayorEffDate) = '" & Trim(Month2) & "' AND " & _
                                   " Year(MBMPayorEffDate) = '" & Trim(Year) & "'"
                       Else
                            strAppendCase = "95"
                            StrAppend1 = " AND PAYCode = '" & Trim(Payor) & "' AND MONTH(MBMPayorEffDate) = '" & Trim(Month2) & "' AND " & _
                                   " Year(MBMPayorEffDate) = '" & Trim(Year) & "'"
                            StrAppend1 = StrAppend1 + StrAppend
                            'sql3 = " SELECT MONTH(MBMPayorEffDate) AS MONTH, YEAR(MBMPayorEffDate) AS YEAR, SUM(DurationStay) AS NoofDay " & _
                                   " From VIEW_CLMReportDCU2 Where PAYCode = '" & Trim(Payor) & "' AND MONTH(MBMPayorEffDate) = '" & Trim(Month2) & "' AND " & _
                                   " Year(MBMPayorEffDate) = '" & Trim(Year) & "'"
                       End If
                                       
                       'sql2 = " GROUP BY PAYCode, MONTH(MBMPayorEffDate), YEAR(MBMPayorEffDate)"
                       
                       'If Len(Trim(StrAppend)) Then
                       '     sql3 = sql3 + StrAppend + sql2
                       'Else
                       '     sql3 = sql3 + sql2
                       'End If
                                              
                       'rst1.Open sql3, medixmasterconnWS, adOpenForwardOnly, adLockBatchOptimistic
                        Call DurationStay(StrAppend1, strAppendCase, Record)
                End If
                
                If Len(!TotalActual) Then
                    Record.SubItems(4) = Format(!TotalActual, "#,##0.00")
                Else
                    Record.SubItems(4) = "0.00"
                End If
                
                If Len(!TotalApproved) Then
                    Record.SubItems(5) = Format(!TotalApproved, "#,##0.00")
                Else
                    Record.SubItems(5) = "0.00"
                End If
                                
                Current = Current + 1
                lblCurrent = Current
            End With
            rst2.MoveNext
            DoEvents
                If intExit = 1 Then
                   Check = False
                   Exit Do
                End If
        Loop
    Loop Until Check = False
    
    If ListView2.ListItems.Count = 0 Then
        MsgBox "No record found!", vbOKOnly + vbCritical, DialogBoxTitle
        CmdClear.Enabled = True
        cmdExit.Enabled = True
        CmdPrint.Enabled = True
        lblCurrent = 0
    End If
    
    intExit = 0
    rst2.Close
    Set rst2 = Nothing
    CmdClear.Enabled = True
    cmdExit.Enabled = True
    CmdPrint.Enabled = True
End Sub

Private Sub EffMonthAllPayor(Optional StrAppend As String)
Dim rst2 As New ADODB.Recordset
Dim cmd As New ADODB.Command
Dim prm1 As ADODB.Parameter
Dim prm2 As ADODB.Parameter
Dim strAppendCase As String
Dim StrAppend1 As String
Dim Record As ListItem
Dim Total As String
Dim Current As String

    Current = 0
    Month = ""
    Month2 = ""
    Year = ""
    ListView2.ListItems.Clear

    If cboClaimCase.ListIndex = 0 Then
        strAppendCase = "90"
        'sql = " SELECT MONTH(MBMPayorEffDate) AS MONTH, YEAR(MBMPayorEffDate) AS YEAR, COUNT(CASNumber) AS TotalCases, " & _
              " SUM(CLMTotalActual) AS TotalActual, SUM(CLMTotalApproved) As TotalApproved, SUM(NoofDay) AS NoofDay " & _
              " From VIEW_CLMReportDCU where 0 = 0 "
          
    ElseIf cboClaimCase.ListIndex = 1 Then
        strAppendCase = "91"
        'sql = " SELECT MONTH(MBMPayorEffDate) AS MONTH, YEAR(MBMPayorEffDate) AS YEAR, COUNT(Distinct CASNumber) AS TotalCases, " & _
              " SUM(CLMTotalActual) AS TotalActual, SUM(CLMTotalApproved) As TotalApproved, SUM(NoofDay) AS NoofDay " & _
              " From VIEW_CLMReportDCU where 0 = 0 "
          
    End If
    
    'sql2 = " GROUP BY MONTH(MBMPayorEffDate), YEAR(MBMPayorEffDate)"
    
    'If Len(Trim(StrAppend)) Then
    '    sql = sql + StrAppend + sql2
    'Else
    '    sql = sql + sql2
    'End If

    'rst2.Open sql, medixmasterconnWS, adOpenForwardOnly, adLockBatchOptimistic
    Set cmd.ActiveConnection = medixmasterconnWS
        cmd.CommandText = "DCUReport"
        cmd.CommandType = adCmdStoredProc
        cmd.CommandTimeout = 300
        Set prm1 = cmd.CreateParameter("StrAppend", adVarChar, adParamInput, 2000, StrAppend)
        cmd.Parameters.Append prm1
        Set prm2 = cmd.CreateParameter("StrAppendCase", adVarChar, adParamInput, 2000, strAppendCase)
        cmd.Parameters.Append prm2
        rst2.CursorLocation = adUseClient
        rst2.CursorType = adOpenForwardOnly
        rst2.CacheSize = 100
        Set rst2 = cmd.Execute
    
    Do
        Do Until rst2.EOF

            With rst2
                
                If Len(!Month & "/" & !Year) Then
                    If Len(!Month) = 1 Then
                        Month = "0" & !Month
                        Set Record = ListView2.ListItems.Add(, , !Year & "/" & Month)
                        Month2 = Month
                        Year = !Year
                    Else
                        Set Record = ListView2.ListItems.Add(, , !Year & "/" & !Month)
                        Month2 = !Month
                        Year = !Year
                    End If
                  
                End If
                
                If Len(!TotalCases) Then
                    Record.SubItems(1) = !TotalCases
                End If
                
                If cboClaimCase.ListIndex = 0 Then
                
                    If rst2!NoofDay = "0" Then
                        Record.SubItems(2) = "1"
                    ElseIf Len(rst2!NoofDay) Then
                        Record.SubItems(2) = rst2!NoofDay
                    End If
                
                ElseIf cboClaimCase.ListIndex = 1 Then
                       strAppendCase = "92"
                       StrAppend1 = " AND MONTH(MBMPayorEffDate) = '" & Trim(Month2) & "' AND " & _
                               " Year(MBMPayorEffDate) = '" & Trim(Year) & "'"
                       StrAppend1 = StrAppend1 + StrAppend
                       'sql3 = " SELECT MONTH(MBMPayorEffDate) AS MONTH, YEAR(MBMPayorEffDate) AS YEAR, SUM(DurationStay) AS NoofDay " & _
                               " From VIEW_CLMReportDCU2 Where MONTH(MBMPayorEffDate) = '" & Trim(Month2) & "' AND " & _
                               " Year(MBMPayorEffDate) = '" & Trim(Year) & "'"
                                       
                       'sql2 = " GROUP BY MONTH(MBMPayorEffDate), YEAR(MBMPayorEffDate)"
                       
                       'If Len(Trim(StrAppend)) Then
                       '     sql3 = sql3 + StrAppend + sql2
                       'Else
                       '     sql3 = sql3 + sql2
                       'End If
                                              
                       'rst1.Open sql3, medixmasterconnWS, adOpenForwardOnly, adLockBatchOptimistic
                        Call DurationStay(StrAppend1, strAppendCase, Record)
                End If
                
                If Len(!TotalActual) Then
                    Record.SubItems(3) = Format(!TotalActual, "#,##0.00")
                Else
                    Record.SubItems(3) = "0.00"
                End If
                
                If Len(!TotalApproved) Then
                    Record.SubItems(4) = Format(!TotalApproved, "#,##0.00")
                Else
                    Record.SubItems(4) = "0.00"
                End If
                                
                Current = Current + 1
                lblCurrent = Current
            End With
            rst2.MoveNext
            DoEvents
                If intExit = 1 Then
                   Check = False
                   Exit Do
                End If
        Loop
    Loop Until Check = False
    
    If ListView2.ListItems.Count = 0 Then
        MsgBox "No record found!", vbOKOnly + vbCritical, DialogBoxTitle
        CmdClear.Enabled = True
        cmdExit.Enabled = True
        CmdPrint.Enabled = True
        'CmdPrintFile.Enabled = True
        lblCurrent = 0
    End If
    
    intExit = 0
    rst2.Close
    Set rst2 = Nothing
    CmdClear.Enabled = True
    cmdExit.Enabled = True
    CmdPrint.Enabled = True
    
End Sub

Private Sub ExpMonthPayor(Optional StrAppend As String)
Dim rst2 As New ADODB.Recordset
Dim cmd As New ADODB.Command
Dim prm1 As ADODB.Parameter
Dim prm2 As ADODB.Parameter
Dim strAppendCase As String
Dim StrAppend1 As String
Dim Record As ListItem
Dim Total As String
Dim Current As String

    Current = 0
    Month = ""
    Month2 = ""
    Year = ""
    ListView2.ListItems.Clear

    If cboClaimCase.ListIndex = 0 Then
        strAppendCase = "99"
        'sql = " SELECT PAYCode, MONTH(MBMPayorExpDate) AS MONTH, YEAR(MBMPayorExpDate) AS YEAR, COUNT(CASNumber) AS TotalCases, " & _
              " SUM(CLMTotalActual) AS TotalActual, SUM(CLMTotalApproved) As TotalApproved, SUM(NoofDay) AS NoofDay " & _
              " From VIEW_CLMReportDCU where 0 = 0 "
          
    ElseIf cboClaimCase.ListIndex = 1 Then
        strAppendCase = "100"
        'sql = " SELECT PAYCode, MONTH(MBMPayorExpDate) AS MONTH, YEAR(MBMPayorExpDate) AS YEAR, COUNT(Distinct CASNumber) AS TotalCases, " & _
              " SUM(CLMTotalActual) AS TotalActual, SUM(CLMTotalApproved) As TotalApproved, SUM(NoofDay) AS NoofDay " & _
              " From VIEW_CLMReportDCU where 0 = 0 "
          
    End If
    
    'sql2 = " GROUP BY PAYCode, MONTH(MBMPayorExpDate), YEAR(MBMPayorExpDate)"
    
    'If Len(Trim(StrAppend)) Then
    '    sql = sql + StrAppend + sql2
    'Else
    '    sql = sql + sql2
    'End If

    'rst2.Open sql, medixmasterconnWS, adOpenForwardOnly, adLockBatchOptimistic
    Set cmd.ActiveConnection = medixmasterconnWS
        cmd.CommandText = "DCUReport"
        cmd.CommandType = adCmdStoredProc
        cmd.CommandTimeout = 300
        Set prm1 = cmd.CreateParameter("StrAppend", adVarChar, adParamInput, 2000, StrAppend)
        cmd.Parameters.Append prm1
        Set prm2 = cmd.CreateParameter("StrAppendCase", adVarChar, adParamInput, 2000, strAppendCase)
        cmd.Parameters.Append prm2
        rst2.CursorLocation = adUseClient
        rst2.CursorType = adOpenForwardOnly
        rst2.CacheSize = 100
        Set rst2 = cmd.Execute
    
    Do
        Do Until rst2.EOF

            With rst2
                
                If Len(!PAYCode) Then
                    Set Record = ListView2.ListItems.Add(, , !PAYCode)
                    Payor = !PAYCode
                Else
                    Set Record = ListView2.ListItems.Add(, , "Empty")
                    Payor = ""
                End If
                
                If Len(!Month & "/" & !Year) Then
                    If Len(!Month) = 1 Then
                        Month = "0" & !Month
                        Record.SubItems(1) = (!Year & "/" & Month)
                        Month2 = Month
                        Year = !Year
                    Else
                        Record.SubItems(1) = (!Year & "/" & !Month)
                        Month2 = !Month
                        Year = !Year
                    End If
                  
                End If
                
                If Len(!TotalCases) Then
                    Record.SubItems(2) = !TotalCases
                End If
                
                If cboClaimCase.ListIndex = 0 Then
                
                    If rst2!NoofDay = "0" Then
                        Record.SubItems(3) = "1"
                    ElseIf Len(rst2!NoofDay) Then
                        Record.SubItems(3) = rst2!NoofDay
                    End If
                
                ElseIf cboClaimCase.ListIndex = 1 Then
                                    
                       If Payor = "" Then
                            strAppendCase = "101"
                            StrAppend1 = " AND (PAYCode IS NULL OR PAYCode = '') AND MONTH(MBMPayorExpDate) = '" & Trim(Month2) & "' AND " & _
                                   " Year(MBMPayorExpDate) = '" & Trim(Year) & "'"
                            StrAppend1 = StrAppend1 + StrAppend
                            'sql3 = " SELECT MONTH(MBMPayorExpDate) AS MONTH, YEAR(MBMPayorExpDate) AS YEAR, SUM(DurationStay) AS NoofDay " & _
                                   " From VIEW_CLMReportDCU2 Where (PAYCode IS NULL OR PAYCode = '') AND MONTH(MBMPayorExpDate) = '" & Trim(Month2) & "' AND " & _
                                   " Year(MBMPayorExpDate) = '" & Trim(Year) & "'"
                       Else
                            strAppendCase = "101"
                            StrAppend1 = " AND PAYCode = '" & Trim(Payor) & "' AND MONTH(MBMPayorExpDate) = '" & Trim(Month2) & "' AND " & _
                                   " Year(MBMPayorExpDate) = '" & Trim(Year) & "'"
                            StrAppend1 = StrAppend1 + StrAppend
                            'sql3 = " SELECT MONTH(MBMPayorExpDate) AS MONTH, YEAR(MBMPayorExpDate) AS YEAR, SUM(DurationStay) AS NoofDay " & _
                                   " From VIEW_CLMReportDCU2 Where PAYCode = '" & Trim(Payor) & "' AND MONTH(MBMPayorExpDate) = '" & Trim(Month2) & "' AND " & _
                                   " Year(MBMPayorExpDate) = '" & Trim(Year) & "'"
                       End If
                                       
                       'sql2 = " GROUP BY PAYCode, MONTH(MBMPayorExpDate), YEAR(MBMPayorExpDate)"
                       
                       'If Len(Trim(StrAppend)) Then
                       '     sql3 = sql3 + StrAppend + sql2
                       'Else
                       '     sql3 = sql3 + sql2
                       'End If
                                              
                       'rst1.Open sql3, medixmasterconnWS, adOpenForwardOnly, adLockBatchOptimistic
                        Call DurationStay(StrAppend1, strAppendCase, Record)
                End If
                
                If Len(!TotalActual) Then
                    Record.SubItems(4) = Format(!TotalActual, "#,##0.00")
                Else
                    Record.SubItems(4) = "0.00"
                End If
                
                If Len(!TotalApproved) Then
                    Record.SubItems(5) = Format(!TotalApproved, "#,##0.00")
                Else
                    Record.SubItems(5) = "0.00"
                End If
                                
                Current = Current + 1
                lblCurrent = Current
            End With
            rst2.MoveNext
            DoEvents
                If intExit = 1 Then
                   Check = False
                   Exit Do
                End If
        Loop
    Loop Until Check = False
    
    If ListView2.ListItems.Count = 0 Then
        MsgBox "No record found!", vbOKOnly + vbCritical, DialogBoxTitle
        CmdClear.Enabled = True
        cmdExit.Enabled = True
        CmdPrint.Enabled = True
        lblCurrent = 0
    End If
    
    intExit = 0
    rst2.Close
    Set rst2 = Nothing
    CmdClear.Enabled = True
    cmdExit.Enabled = True
    CmdPrint.Enabled = True
End Sub

Private Sub ExpMonthAllPayor(Optional StrAppend As String)
Dim rst2 As New ADODB.Recordset
Dim cmd As New ADODB.Command
Dim prm1 As ADODB.Parameter
Dim prm2 As ADODB.Parameter
Dim strAppendCase As String
Dim StrAppend1 As String
Dim Record As ListItem
Dim Total As String
Dim Current As String

    Current = 0
    Month = ""
    Month2 = ""
    Year = ""
    ListView2.ListItems.Clear

    If cboClaimCase.ListIndex = 0 Then
        strAppendCase = "96"
        'sql = " SELECT MONTH(MBMPayorExpDate) AS MONTH, YEAR(MBMPayorExpDate) AS YEAR, COUNT(CASNumber) AS TotalCases, " & _
              " SUM(CLMTotalActual) AS TotalActual, SUM(CLMTotalApproved) As TotalApproved, SUM(NoofDay) AS NoofDay " & _
              " From VIEW_CLMReportDCU where 0 = 0 "
          
    ElseIf cboClaimCase.ListIndex = 1 Then
        strAppendCase = "97"
        'sql = " SELECT MONTH(MBMPayorExpDate) AS MONTH, YEAR(MBMPayorExpDate) AS YEAR, COUNT(Distinct CASNumber) AS TotalCases, " & _
              " SUM(CLMTotalActual) AS TotalActual, SUM(CLMTotalApproved) As TotalApproved, SUM(NoofDay) AS NoofDay " & _
              " From VIEW_CLMReportDCU where 0 = 0 "
          
    End If
    
    'sql2 = " GROUP BY MONTH(MBMPayorExpDate), YEAR(MBMPayorExpDate)"
    
    'If Len(Trim(StrAppend)) Then
    '    sql = sql + StrAppend + sql2
    'Else
    '    sql = sql + sql2
    'End If

    'rst2.Open sql, medixmasterconnWS, adOpenForwardOnly, adLockBatchOptimistic
    Set cmd.ActiveConnection = medixmasterconnWS
        cmd.CommandText = "DCUReport"
        cmd.CommandType = adCmdStoredProc
        cmd.CommandTimeout = 300
        Set prm1 = cmd.CreateParameter("StrAppend", adVarChar, adParamInput, 2000, StrAppend)
        cmd.Parameters.Append prm1
        Set prm2 = cmd.CreateParameter("StrAppendCase", adVarChar, adParamInput, 2000, strAppendCase)
        cmd.Parameters.Append prm2
        rst2.CursorLocation = adUseClient
        rst2.CursorType = adOpenForwardOnly
        rst2.CacheSize = 100
        Set rst2 = cmd.Execute
    
    Do
        Do Until rst2.EOF

            With rst2
                
                If Len(!Month & "/" & !Year) Then
                    If Len(!Month) = 1 Then
                        Month = "0" & !Month
                        Set Record = ListView2.ListItems.Add(, , !Year & "/" & Month)
                        Month2 = Month
                        Year = !Year
                    Else
                        Set Record = ListView2.ListItems.Add(, , !Year & "/" & !Month)
                        Month2 = !Month
                        Year = !Year
                    End If
                  
                End If
                
                If Len(!TotalCases) Then
                    Record.SubItems(1) = !TotalCases
                End If
                
                If cboClaimCase.ListIndex = 0 Then
                
                    If rst2!NoofDay = "0" Then
                        Record.SubItems(2) = "1"
                    ElseIf Len(rst2!NoofDay) Then
                        Record.SubItems(2) = rst2!NoofDay
                    End If
                
                ElseIf cboClaimCase.ListIndex = 1 Then
                      strAppendCase = "98"
                      StrAppend1 = " AND MONTH(MBMPayorExpDate) = '" & Trim(Month2) & "' AND " & _
                               " Year(MBMPayorExpDate) = '" & Trim(Year) & "'"
                      StrAppend1 = StrAppend1 + StrAppend
                       'sql3 = " SELECT MONTH(MBMPayorExpDate) AS MONTH, YEAR(MBMPayorExpDate) AS YEAR, SUM(DurationStay) AS NoofDay " & _
                               " From VIEW_CLMReportDCU2 Where MONTH(MBMPayorExpDate) = '" & Trim(Month2) & "' AND " & _
                               " Year(MBMPayorExpDate) = '" & Trim(Year) & "'"
                                       
                       'sql2 = " GROUP BY MONTH(MBMPayorExpDate), YEAR(MBMPayorExpDate)"
                       
                       'If Len(Trim(StrAppend)) Then
                       '     sql3 = sql3 + StrAppend + sql2
                       'Else
                       '     sql3 = sql3 + sql2
                       'End If
                                              
                       'rst1.Open sql3, medixmasterconnWS, adOpenForwardOnly, adLockBatchOptimistic
                        Call DurationStay(StrAppend1, strAppendCase, Record)
                End If
                
                If Len(!TotalActual) Then
                    Record.SubItems(3) = Format(!TotalActual, "#,##0.00")
                Else
                    Record.SubItems(3) = "0.00"
                End If
                
                If Len(!TotalApproved) Then
                    Record.SubItems(4) = Format(!TotalApproved, "#,##0.00")
                Else
                    Record.SubItems(4) = "0.00"
                End If
                                
                Current = Current + 1
                lblCurrent = Current
            End With
            rst2.MoveNext
            DoEvents
                If intExit = 1 Then
                   Check = False
                   Exit Do
                End If
        Loop
    Loop Until Check = False
    
    If ListView2.ListItems.Count = 0 Then
        MsgBox "No record found!", vbOKOnly + vbCritical, DialogBoxTitle
        CmdClear.Enabled = True
        cmdExit.Enabled = True
        CmdPrint.Enabled = True
        lblCurrent = 0
    End If
    
    intExit = 0
    rst2.Close
    Set rst2 = Nothing
    CmdClear.Enabled = True
    cmdExit.Enabled = True
    CmdPrint.Enabled = True
    
End Sub

Private Sub DOAMonthPayor(Optional StrAppend As String)
Dim rst2 As New ADODB.Recordset
Dim cmd As New ADODB.Command
Dim prm1 As ADODB.Parameter
Dim prm2 As ADODB.Parameter
Dim strAppendCase As String
Dim StrAppend1 As String
Dim Record As ListItem
Dim Total As String
Dim Current As String

    Current = 0
    Month = ""
    Month2 = ""
    Year = ""
    ListView2.ListItems.Clear

    If cboClaimCase.ListIndex = 0 Then
        strAppendCase = "105"
        'sql = " SELECT PAYCode, MONTH(CASDateAdmitted) AS MONTH, YEAR(CASDateAdmitted) AS YEAR, COUNT(CASNumber) AS TotalCases, " & _
              " SUM(CLMTotalActual) AS TotalActual, SUM(CLMTotalApproved) As TotalApproved, SUM(NoofDay) AS NoofDay " & _
              " From VIEW_CLMReportDCU where 0 = 0 "
          
    ElseIf cboClaimCase.ListIndex = 1 Then
        strAppendCase = "106"
        'sql = " SELECT PAYCode, MONTH(CASDateAdmitted) AS MONTH, YEAR(CASDateAdmitted) AS YEAR, COUNT(Distinct CASNumber) AS TotalCases, " & _
              " SUM(CLMTotalActual) AS TotalActual, SUM(CLMTotalApproved) As TotalApproved, SUM(NoofDay) AS NoofDay " & _
              " From VIEW_CLMReportDCU where 0 = 0 "
          
    End If
    
    'sql2 = " GROUP BY PAYCode, MONTH(CASDateAdmitted), YEAR(CASDateAdmitted)"
    
    'If Len(Trim(StrAppend)) Then
    '    sql = sql + StrAppend + sql2
    'Else
    '    sql = sql + sql2
    'End If

    'rst2.Open sql, medixmasterconnWS, adOpenForwardOnly, adLockBatchOptimistic
    Set cmd.ActiveConnection = medixmasterconnWS
        cmd.CommandText = "DCUReport"
        cmd.CommandType = adCmdStoredProc
        cmd.CommandTimeout = 300
        Set prm1 = cmd.CreateParameter("StrAppend", adVarChar, adParamInput, 2000, StrAppend)
        cmd.Parameters.Append prm1
        Set prm2 = cmd.CreateParameter("StrAppendCase", adVarChar, adParamInput, 2000, strAppendCase)
        cmd.Parameters.Append prm2
        rst2.CursorLocation = adUseClient
        rst2.CursorType = adOpenForwardOnly
        rst2.CacheSize = 100
        Set rst2 = cmd.Execute
        
    Do
        Do Until rst2.EOF

            With rst2
                
                If Len(!PAYCode) Then
                    Set Record = ListView2.ListItems.Add(, , !PAYCode)
                    Payor = !PAYCode
                Else
                    Set Record = ListView2.ListItems.Add(, , "Empty")
                    Payor = ""
                End If
                
                If Len(!Month & "/" & !Year) Then
                    If Len(!Month) = 1 Then
                        Month = "0" & !Month
                        Record.SubItems(1) = (!Year & "/" & Month)
                        Month2 = Month
                        Year = !Year
                    Else
                        Record.SubItems(1) = (!Year & "/" & !Month)
                        Month2 = !Month
                        Year = !Year
                    End If
                  
                End If
                
                If Len(!TotalCases) Then
                    Record.SubItems(2) = !TotalCases
                End If
                
                If cboClaimCase.ListIndex = 0 Then
                
                    If rst2!NoofDay = "0" Then
                        Record.SubItems(3) = "1"
                    ElseIf Len(rst2!NoofDay) Then
                        Record.SubItems(3) = rst2!NoofDay
                    End If
                
                ElseIf cboClaimCase.ListIndex = 1 Then
                                    
                       If Payor = "" Then
                            strAppendCase = "107"
                            StrAppend1 = " AND (PAYCode IS NULL OR PAYCode = '') AND MONTH(CASDateAdmitted) = '" & Trim(Month2) & "' AND " & _
                                   " Year(CASDateAdmitted) = '" & Trim(Year) & "'"
                            StrAppend1 = StrAppend1 + StrAppend
                            'sql3 = " SELECT MONTH(CASDateAdmitted) AS MONTH, YEAR(CASDateAdmitted) AS YEAR, SUM(DurationStay) AS NoofDay " & _
                                   " From VIEW_CLMReportDCU2 Where (PAYCode IS NULL OR PAYCode = '') AND MONTH(CASDateAdmitted) = '" & Trim(Month2) & "' AND " & _
                                   " Year(CASDateAdmitted) = '" & Trim(Year) & "'"
                       Else
                            strAppendCase = "107"
                            StrAppend1 = " AND PAYCode = '" & Trim(Payor) & "' AND MONTH(CASDateAdmitted) = '" & Trim(Month2) & "' AND " & _
                                   " Year(CASDateAdmitted) = '" & Trim(Year) & "'"
                            StrAppend1 = StrAppend1 + StrAppend
                            'sql3 = " SELECT MONTH(CASDateAdmitted) AS MONTH, YEAR(CASDateAdmitted) AS YEAR, SUM(DurationStay) AS NoofDay " & _
                                   " From VIEW_CLMReportDCU2 Where PAYCode = '" & Trim(Payor) & "' AND MONTH(CASDateAdmitted) = '" & Trim(Month2) & "' AND " & _
                                   " Year(CASDateAdmitted) = '" & Trim(Year) & "'"
                       End If
                                       
                       'sql2 = " GROUP BY PAYCode, MONTH(CASDateAdmitted), YEAR(CASDateAdmitted)"
                       
                       'If Len(Trim(StrAppend)) Then
                       '     sql3 = sql3 + StrAppend + sql2
                       'Else
                       '     sql3 = sql3 + sql2
                       'End If
                                              
                       'rst1.Open sql3, medixmasterconnWS, adOpenForwardOnly, adLockBatchOptimistic
                        Call DurationStay(StrAppend1, strAppendCase, Record)
                End If
                
                If Len(!TotalActual) Then
                    Record.SubItems(4) = Format(!TotalActual, "#,##0.00")
                Else
                    Record.SubItems(4) = "0.00"
                End If
                
                If Len(!TotalApproved) Then
                    Record.SubItems(5) = Format(!TotalApproved, "#,##0.00")
                Else
                    Record.SubItems(5) = "0.00"
                End If
                                
                Current = Current + 1
                lblCurrent = Current
            End With
            rst2.MoveNext
            DoEvents
                If intExit = 1 Then
                   Check = False
                   Exit Do
                End If
        Loop
    Loop Until Check = False
    
    If ListView2.ListItems.Count = 0 Then
        MsgBox "No record found!", vbOKOnly + vbCritical, DialogBoxTitle
        CmdClear.Enabled = True
        cmdExit.Enabled = True
        CmdPrint.Enabled = True
        lblCurrent = 0
    End If
    
    intExit = 0
    rst2.Close
    Set rst2 = Nothing
    CmdClear.Enabled = True
    cmdExit.Enabled = True
    CmdPrint.Enabled = True
End Sub

Private Sub DOAMonthAllPayor(Optional StrAppend As String)
Dim rst2 As New ADODB.Recordset
Dim cmd As New ADODB.Command
Dim prm1 As ADODB.Parameter
Dim prm2 As ADODB.Parameter
Dim strAppendCase As String
Dim StrAppend1 As String
Dim Record As ListItem
Dim Total As String
Dim Current As String
        
    Current = 0
    Month = ""
    Month2 = ""
    Year = ""
    ListView2.ListItems.Clear
    
    If cboClaimCase.ListIndex = 0 Then
        strAppendCase = "102"
        'sql = " SELECT MONTH(CASDateAdmitted) AS MONTH, YEAR(CASDateAdmitted) AS YEAR, COUNT(CASNumber) AS TotalCases, " & _
              " SUM(CLMTotalActual) AS TotalActual, SUM(CLMTotalApproved) As TotalApproved, SUM(NoofDay) AS NoofDay " & _
              " From VIEW_CLMReportDCU where 0 = 0 "
          
    ElseIf cboClaimCase.ListIndex = 1 Then
        strAppendCase = "103"
        'sql = " SELECT MONTH(CASDateAdmitted) AS MONTH, YEAR(CASDateAdmitted) AS YEAR, COUNT(Distinct CASNumber) AS TotalCases, " & _
              " SUM(CLMTotalActual) AS TotalActual, SUM(CLMTotalApproved) As TotalApproved, SUM(NoofDay) AS NoofDay " & _
              " From VIEW_CLMReportDCU where 0 = 0 "
          
    End If
    
    'sql2 = " GROUP BY MONTH(CASDateAdmitted), YEAR(CASDateAdmitted)"
    
    'If Len(Trim(StrAppend)) Then
    '    sql = sql + StrAppend + sql2
    'Else
    '    sql = sql + sql2
    'End If

    'rst2.Open sql, medixmasterconnWS, adOpenForwardOnly, adLockBatchOptimistic
    Set cmd.ActiveConnection = medixmasterconnWS
        cmd.CommandText = "DCUReport"
        cmd.CommandType = adCmdStoredProc
        cmd.CommandTimeout = 300
        Set prm1 = cmd.CreateParameter("StrAppend", adVarChar, adParamInput, 2000, StrAppend)
        cmd.Parameters.Append prm1
        Set prm2 = cmd.CreateParameter("StrAppendCase", adVarChar, adParamInput, 2000, strAppendCase)
        cmd.Parameters.Append prm2
        rst2.CursorLocation = adUseClient
        rst2.CursorType = adOpenForwardOnly
        rst2.CacheSize = 100
        Set rst2 = cmd.Execute
        
    Do
        Do Until rst2.EOF

            With rst2
                
                If Len(!Month & "/" & !Year) Then
                    If Len(!Month) = 1 Then
                        Month = "0" & !Month
                        Set Record = ListView2.ListItems.Add(, , !Year & "/" & Month)
                        Month2 = Month
                        Year = !Year
                    Else
                        Set Record = ListView2.ListItems.Add(, , !Year & "/" & !Month)
                        Month2 = !Month
                        Year = !Year
                    End If
                  
                End If
                
                If Len(!TotalCases) Then
                    Record.SubItems(1) = !TotalCases
                End If
                
                If cboClaimCase.ListIndex = 0 Then
                
                    If rst2!NoofDay = "0" Then
                        Record.SubItems(2) = "1"
                    ElseIf Len(rst2!NoofDay) Then
                        Record.SubItems(2) = rst2!NoofDay
                    End If
                
                ElseIf cboClaimCase.ListIndex = 1 Then
                        strAppendCase = "104"
                        StrAppend1 = " AND MONTH(CASDateAdmitted) = '" & Trim(Month2) & "' AND " & _
                               " Year(CASDateAdmitted) = '" & Trim(Year) & "'"
                        StrAppend1 = StrAppend1 + StrAppend
                       'sql3 = " SELECT MONTH(CASDateAdmitted) AS MONTH, YEAR(CASDateAdmitted) AS YEAR, SUM(DurationStay) AS NoofDay " & _
                               " From VIEW_CLMReportDCU2 Where MONTH(CASDateAdmitted) = '" & Trim(Month2) & "' AND " & _
                               " Year(CASDateAdmitted) = '" & Trim(Year) & "'"
                                       
                       'sql2 = " GROUP BY MONTH(CASDateAdmitted), YEAR(CASDateAdmitted)"
                       
                       'If Len(Trim(StrAppend)) Then
                       '     sql3 = sql3 + StrAppend + sql2
                       'Else
                       '     sql3 = sql3 + sql2
                       'End If
                                              
                       'rst1.Open sql3, medixmasterconnWS, adOpenForwardOnly, adLockBatchOptimistic
                        Call DurationStay(StrAppend1, strAppendCase, Record)
                End If
                
                If Len(!TotalActual) Then
                    Record.SubItems(3) = Format(!TotalActual, "#,##0.00")
                Else
                    Record.SubItems(3) = "0.00"
                End If
                
                If Len(!TotalApproved) Then
                    Record.SubItems(4) = Format(!TotalApproved, "#,##0.00")
                Else
                    Record.SubItems(4) = "0.00"
                End If
                                
                Current = Current + 1
                lblCurrent = Current
            End With
            rst2.MoveNext
            DoEvents
                If intExit = 1 Then
                   Check = False
                   Exit Do
                End If
        Loop
    Loop Until Check = False
    
    If ListView2.ListItems.Count = 0 Then
        MsgBox "No record found!", vbOKOnly + vbCritical, DialogBoxTitle
        CmdClear.Enabled = True
        cmdExit.Enabled = True
        CmdPrint.Enabled = True
        lblCurrent = 0
    End If
    
    intExit = 0
    rst2.Close
    Set rst2 = Nothing
    CmdClear.Enabled = True
    cmdExit.Enabled = True
    CmdPrint.Enabled = True
    
End Sub

Private Sub DODMonthPayor(Optional StrAppend As String)
Dim rst2 As New ADODB.Recordset
Dim cmd As New ADODB.Command
Dim prm1 As ADODB.Parameter
Dim prm2 As ADODB.Parameter
Dim strAppendCase As String
Dim StrAppend1 As String
Dim Record As ListItem
Dim Total As String
Dim Current As String

    Current = 0
    Month = ""
    Month2 = ""
    Year = ""
    ListView2.ListItems.Clear

    If cboClaimCase.ListIndex = 0 Then
        strAppendCase = "111"
        'sql = " SELECT PAYCode, MONTH(CASDischargeDate) AS MONTH, YEAR(CASDischargeDate) AS YEAR, COUNT(CASNumber) AS TotalCases, " & _
              " SUM(CLMTotalActual) AS TotalActual, SUM(CLMTotalApproved) As TotalApproved, SUM(NoofDay) AS NoofDay " & _
              " From VIEW_CLMReportDCU where 0 = 0 "
          
    ElseIf cboClaimCase.ListIndex = 1 Then
        strAppendCase = "112"
        'sql = " SELECT PAYCode, MONTH(CASDischargeDate) AS MONTH, YEAR(CASDischargeDate) AS YEAR, COUNT(Distinct CASNumber) AS TotalCases, " & _
              " SUM(CLMTotalActual) AS TotalActual, SUM(CLMTotalApproved) As TotalApproved, SUM(NoofDay) AS NoofDay " & _
              " From VIEW_CLMReportDCU where 0 = 0 "
          
    End If
    
    'sql2 = " GROUP BY PAYCode, MONTH(CASDischargeDate), YEAR(CASDischargeDate)"
    
    'If Len(Trim(StrAppend)) Then
    '    sql = sql + StrAppend + sql2
    'Else
    '    sql = sql + sql2
    'End If

    'rst2.Open sql, medixmasterconnWS, adOpenForwardOnly, adLockBatchOptimistic
    Set cmd.ActiveConnection = medixmasterconnWS
        cmd.CommandText = "DCUReport"
        cmd.CommandType = adCmdStoredProc
        cmd.CommandTimeout = 300
        Set prm1 = cmd.CreateParameter("StrAppend", adVarChar, adParamInput, 2000, StrAppend)
        cmd.Parameters.Append prm1
        Set prm2 = cmd.CreateParameter("StrAppendCase", adVarChar, adParamInput, 2000, strAppendCase)
        cmd.Parameters.Append prm2
        rst2.CursorLocation = adUseClient
        rst2.CursorType = adOpenForwardOnly
        rst2.CacheSize = 100
        Set rst2 = cmd.Execute
    
    Do
        Do Until rst2.EOF

            With rst2
                
                If Len(!PAYCode) Then
                    Set Record = ListView2.ListItems.Add(, , !PAYCode)
                    Payor = !PAYCode
                Else
                    Set Record = ListView2.ListItems.Add(, , "Empty")
                    Payor = ""
                End If
                
                If Len(!Month & "/" & !Year) Then
                    If Len(!Month) = 1 Then
                        Month = "0" & !Month
                        Record.SubItems(1) = (!Year & "/" & Month)
                        Month2 = Month
                        Year = !Year
                    Else
                        Record.SubItems(1) = (!Year & "/" & !Month)
                        Month2 = !Month
                        Year = !Year
                    End If
                  
                End If
                
                If Len(!TotalCases) Then
                    Record.SubItems(2) = !TotalCases
                End If
                
                If cboClaimCase.ListIndex = 0 Then
                
                    If rst2!NoofDay = "0" Then
                        Record.SubItems(3) = "1"
                    ElseIf Len(rst2!NoofDay) Then
                        Record.SubItems(3) = rst2!NoofDay
                    End If
                
                ElseIf cboClaimCase.ListIndex = 1 Then
                                    
                       If Payor = "" Then
                            strAppendCase = "113"
                            StrAppend1 = " AND (PAYCode IS NULL OR PAYCode = '') AND MONTH(CASDischargeDate) = '" & Trim(Month2) & "' AND " & _
                                   " Year(CASDischargeDate) = '" & Trim(Year) & "'"
                            StrAppend1 = StrAppend1 + StrAppend
                            'sql3 = " SELECT MONTH(CASDischargeDate) AS MONTH, YEAR(CASDischargeDate) AS YEAR, SUM(DurationStay) AS NoofDay " & _
                                   " From VIEW_CLMReportDCU2 Where (PAYCode IS NULL OR PAYCode = '') AND MONTH(CASDischargeDate) = '" & Trim(Month2) & "' AND " & _
                                   " Year(CASDischargeDate) = '" & Trim(Year) & "'"
                       Else
                            strAppendCase = "113"
                            StrAppend1 = " AND PAYCode = '" & Trim(Payor) & "' AND MONTH(CASDischargeDate) = '" & Trim(Month2) & "' AND " & _
                                   " Year(CASDischargeDate) = '" & Trim(Year) & "'"
                            StrAppend1 = StrAppend1 + StrAppend
                            'sql3 = " SELECT MONTH(CASDischargeDate) AS MONTH, YEAR(CASDischargeDate) AS YEAR, SUM(DurationStay) AS NoofDay " & _
                                   " From VIEW_CLMReportDCU2 Where PAYCode = '" & Trim(Payor) & "' AND MONTH(CASDischargeDate) = '" & Trim(Month2) & "' AND " & _
                                   " Year(CASDischargeDate) = '" & Trim(Year) & "'"
                       End If
                                       
                       'sql2 = " GROUP BY PAYCode, MONTH(CASDischargeDate), YEAR(CASDischargeDate)"
                       
                       'If Len(Trim(StrAppend)) Then
                       '     sql3 = sql3 + StrAppend + sql2
                       'Else
                       '     sql3 = sql3 + sql2
                       'End If
                                              
                       'rst1.Open sql3, medixmasterconnWS, adOpenForwardOnly, adLockBatchOptimistic
                        Call DurationStay(StrAppend1, strAppendCase, Record)
                End If
                
                If Len(!TotalActual) Then
                    Record.SubItems(4) = Format(!TotalActual, "#,##0.00")
                Else
                    Record.SubItems(4) = "0.00"
                End If
                
                If Len(!TotalApproved) Then
                    Record.SubItems(5) = Format(!TotalApproved, "#,##0.00")
                Else
                    Record.SubItems(5) = "0.00"
                End If
                                
                Current = Current + 1
                lblCurrent = Current
            End With
            rst2.MoveNext
            DoEvents
                If intExit = 1 Then
                   Check = False
                   Exit Do
                End If
        Loop
    Loop Until Check = False
    
    If ListView2.ListItems.Count = 0 Then
        MsgBox "No record found!", vbOKOnly + vbCritical, DialogBoxTitle
        CmdClear.Enabled = True
        cmdExit.Enabled = True
        CmdPrint.Enabled = True
        lblCurrent = 0
    End If
    
    intExit = 0
    rst2.Close
    Set rst2 = Nothing
    CmdClear.Enabled = True
    cmdExit.Enabled = True
    CmdPrint.Enabled = True
End Sub

Private Sub DODMonthAllPayor(Optional StrAppend As String)
Dim rst2 As New ADODB.Recordset
Dim cmd As New ADODB.Command
Dim prm1 As ADODB.Parameter
Dim prm2 As ADODB.Parameter
Dim strAppendCase As String
Dim StrAppend1 As String
Dim Record As ListItem
Dim Total As String
Dim Current As String

    Current = 0
    Month = ""
    Month2 = ""
    Year = ""
    ListView2.ListItems.Clear

    If cboClaimCase.ListIndex = 0 Then
        strAppendCase = "108"
        'sql = " SELECT MONTH(CASDischargeDate) AS MONTH, YEAR(CASDischargeDate) AS YEAR, COUNT(CASNumber) AS TotalCases, " & _
              " SUM(CLMTotalActual) AS TotalActual, SUM(CLMTotalApproved) As TotalApproved, SUM(NoofDay) AS NoofDay " & _
              " From VIEW_CLMReportDCU where 0 = 0 "
          
    ElseIf cboClaimCase.ListIndex = 1 Then
        strAppendCase = "109"
        'sql = " SELECT MONTH(CASDischargeDate) AS MONTH, YEAR(CASDischargeDate) AS YEAR, COUNT(Distinct CASNumber) AS TotalCases, " & _
              " SUM(CLMTotalActual) AS TotalActual, SUM(CLMTotalApproved) As TotalApproved, SUM(NoofDay) AS NoofDay " & _
              " From VIEW_CLMReportDCU where 0 = 0 "
          
    End If
    
    'sql2 = " GROUP BY MONTH(CASDischargeDate), YEAR(CASDischargeDate)"
    
    'If Len(Trim(StrAppend)) Then
    '    sql = sql + StrAppend + sql2
    'Else
    '    sql = sql + sql2
    'End If

    'rst2.Open sql, medixmasterconnWS, adOpenForwardOnly, adLockBatchOptimistic
    Set cmd.ActiveConnection = medixmasterconnWS
        cmd.CommandText = "DCUReport"
        cmd.CommandType = adCmdStoredProc
        cmd.CommandTimeout = 300
        Set prm1 = cmd.CreateParameter("StrAppend", adVarChar, adParamInput, 2000, StrAppend)
        cmd.Parameters.Append prm1
        Set prm2 = cmd.CreateParameter("StrAppendCase", adVarChar, adParamInput, 2000, strAppendCase)
        cmd.Parameters.Append prm2
        rst2.CursorLocation = adUseClient
        rst2.CursorType = adOpenForwardOnly
        rst2.CacheSize = 100
        Set rst2 = cmd.Execute
    
    Do
        Do Until rst2.EOF

            With rst2
                
                If Len(!Month & "/" & !Year) Then
                    If Len(!Month) = 1 Then
                        Month = "0" & !Month
                        Set Record = ListView2.ListItems.Add(, , !Year & "/" & Month)
                        Month2 = Month
                        Year = !Year
                    Else
                        Set Record = ListView2.ListItems.Add(, , !Year & "/" & !Month)
                        Month2 = !Month
                        Year = !Year
                    End If
                  
                End If
                
                If Len(!TotalCases) Then
                    Record.SubItems(1) = !TotalCases
                End If
                
                If cboClaimCase.ListIndex = 0 Then
                
                    If rst2!NoofDay = "0" Then
                        Record.SubItems(2) = "1"
                    ElseIf Len(rst2!NoofDay) Then
                        Record.SubItems(2) = rst2!NoofDay
                    End If
                
                ElseIf cboClaimCase.ListIndex = 1 Then
                      strAppendCase = "110"
                      StrAppend1 = " AND MONTH(CASDischargeDate) = '" & Trim(Month2) & "' AND " & _
                               " Year(CASDischargeDate) = '" & Trim(Year) & "'"
                      StrAppend1 = StrAppend1 + StrAppend
                       'sql3 = " SELECT MONTH(CASDischargeDate) AS MONTH, YEAR(CASDischargeDate) AS YEAR, SUM(DurationStay) AS NoofDay " & _
                               " From VIEW_CLMReportDCU2 Where MONTH(CASDischargeDate) = '" & Trim(Month2) & "' AND " & _
                               " Year(CASDischargeDate) = '" & Trim(Year) & "'"
                                       
                       'sql2 = " GROUP BY MONTH(CASDischargeDate), YEAR(CASDischargeDate)"
                       
                       'If Len(Trim(StrAppend)) Then
                       '     sql3 = sql3 + StrAppend + sql2
                       'Else
                       '     sql3 = sql3 + sql2
                       'End If
                                              
                       'rst1.Open sql3, medixmasterconnWS, adOpenForwardOnly, adLockBatchOptimistic
                        Call DurationStay(StrAppend1, strAppendCase, Record)
                End If
                
                If Len(!TotalActual) Then
                    Record.SubItems(3) = Format(!TotalActual, "#,##0.00")
                Else
                    Record.SubItems(3) = "0.00"
                End If
                
                If Len(!TotalApproved) Then
                    Record.SubItems(4) = Format(!TotalApproved, "#,##0.00")
                Else
                    Record.SubItems(4) = "0.00"
                End If
                                
                Current = Current + 1
                lblCurrent = Current
            End With
            rst2.MoveNext
            DoEvents
                If intExit = 1 Then
                   Check = False
                   Exit Do
                End If
        Loop
    Loop Until Check = False
    
    If ListView2.ListItems.Count = 0 Then
        MsgBox "No record found!", vbOKOnly + vbCritical, DialogBoxTitle
        CmdClear.Enabled = True
        cmdExit.Enabled = True
        CmdPrint.Enabled = True
        lblCurrent = 0
    End If
    
    intExit = 0
    rst2.Close
    Set rst2 = Nothing
    CmdClear.Enabled = True
    cmdExit.Enabled = True
    CmdPrint.Enabled = True
    
End Sub

Private Sub BordxMonthPayor(Optional StrAppend As String)
Dim rst2 As New ADODB.Recordset
Dim cmd As New ADODB.Command
Dim prm1 As ADODB.Parameter
Dim prm2 As ADODB.Parameter
Dim strAppendCase As String
Dim Record As ListItem
Dim Total As String
Dim Current As String

    Current = 0
    Month = ""
    Month2 = ""
    Year = ""
    ListView2.ListItems.Clear
      
    strAppendCase = "115"
    'sql = " SELECT PAYCode, MONTH(BordxDate) AS MONTH, YEAR(BordxDate) AS YEAR, COUNT(CASNumber) AS TotalCases, " & _
          " SUM(CLMTotalActual) AS TotalActual, SUM(CLMTotalApproved) As TotalApproved, SUM(NoofDay) AS NoofDay " & _
          " From VIEW_CLMReportDCU where 0 = 0 "
    
    
    'sql2 = " GROUP BY PAYCode, MONTH(BordxDate), YEAR(BordxDate)"
    
    'If Len(Trim(StrAppend)) Then
    '    sql = sql + StrAppend + sql2
    'Else
    '    sql = sql + sql2
    'End If

    'rst2.Open sql, medixmasterconnWS, adOpenForwardOnly, adLockBatchOptimistic
    Set cmd.ActiveConnection = medixmasterconnWS
        cmd.CommandText = "DCUReport"
        cmd.CommandType = adCmdStoredProc
        cmd.CommandTimeout = 300
        Set prm1 = cmd.CreateParameter("StrAppend", adVarChar, adParamInput, 2000, StrAppend)
        cmd.Parameters.Append prm1
        Set prm2 = cmd.CreateParameter("StrAppendCase", adVarChar, adParamInput, 2000, strAppendCase)
        cmd.Parameters.Append prm2
        rst2.CursorLocation = adUseClient
        rst2.CursorType = adOpenForwardOnly
        rst2.CacheSize = 100
        Set rst2 = cmd.Execute
        
    Do
        Do Until rst2.EOF

            With rst2
                
                If Len(!PAYCode) Then
                    Set Record = ListView2.ListItems.Add(, , !PAYCode)
                    Payor = !PAYCode
                Else
                    Set Record = ListView2.ListItems.Add(, , "Empty")
                    Payor = ""
                End If
                
                If Len(!Month) Or Len(!Year) Then
                    If Len(!Month) = 1 Then
                        Month = "0" & !Month
                        Record.SubItems(1) = (!Year & "/" & Month)
                        Month2 = Month
                        Year = !Year
                    Else
                        Record.SubItems(1) = (!Year & "/" & !Month)
                        Month2 = !Month
                        Year = !Year
                    End If
                Else
                    Record.SubItems(1) = ("" & "/" & "")
                    Month2 = ""
                    Year = ""
                End If
                
                If Len(!TotalCases) Then
                    Record.SubItems(2) = !TotalCases
                End If
                
                If rst2!NoofDay = "0" Then
                    Record.SubItems(3) = "1"
                ElseIf Len(rst2!NoofDay) Then
                    Record.SubItems(3) = rst2!NoofDay
                End If
                                
                If Len(!TotalActual) Then
                    Record.SubItems(4) = Format(!TotalActual, "#,##0.00")
                Else
                    Record.SubItems(4) = "0.00"
                End If
                
                If Len(!TotalApproved) Then
                    Record.SubItems(5) = Format(!TotalApproved, "#,##0.00")
                Else
                    Record.SubItems(5) = "0.00"
                End If
                                
                Current = Current + 1
                lblCurrent = Current
            End With
            rst2.MoveNext
            DoEvents
                If intExit = 1 Then
                   Check = False
                   Exit Do
                End If
        Loop
    Loop Until Check = False
    
    If ListView2.ListItems.Count = 0 Then
        MsgBox "No record found!", vbOKOnly + vbCritical, DialogBoxTitle
        CmdClear.Enabled = True
        cmdExit.Enabled = True
        CmdPrint.Enabled = True
        lblCurrent = 0
    End If
    
    intExit = 0
    rst2.Close
    Set rst2 = Nothing
    CmdClear.Enabled = True
    cmdExit.Enabled = True
    CmdPrint.Enabled = True
End Sub

Private Sub BordxMonthAllPayor(Optional StrAppend As String)
Dim rst2 As New ADODB.Recordset
Dim cmd As New ADODB.Command
Dim prm1 As ADODB.Parameter
Dim prm2 As ADODB.Parameter
Dim strAppendCase As String
Dim Record As ListItem
Dim Total As String
Dim Current As String
Dim strCheck As Boolean

    Current = 0
    Month = ""
    Month2 = ""
    Year = ""
    ListView2.ListItems.Clear
    strCheck = False
    
    strAppendCase = "114"
    'sql = " SELECT MONTH(BordxDate) AS MONTH, YEAR(BordxDate) AS YEAR, COUNT(CASNumber) AS TotalCases, " & _
          " SUM(CLMTotalActual) AS TotalActual, SUM(CLMTotalApproved) As TotalApproved, SUM(NoofDay) AS NoofDay " & _
          " From VIEW_CLMReportDCU where 0 = 0 "
    
    'sql2 = " GROUP BY MONTH(BordxDate), YEAR(BordxDate)"
    
    'If Len(Trim(StrAppend)) Then
    '    sql = sql + StrAppend + sql2
    'Else
    '    sql = sql + sql2
    'End If

    'rst2.Open sql, medixmasterconnWS, adOpenForwardOnly, adLockBatchOptimistic
    Set cmd.ActiveConnection = medixmasterconnWS
        cmd.CommandText = "DCUReport"
        cmd.CommandType = adCmdStoredProc
        cmd.CommandTimeout = 300
        Set prm1 = cmd.CreateParameter("StrAppend", adVarChar, adParamInput, 2000, StrAppend)
        cmd.Parameters.Append prm1
        Set prm2 = cmd.CreateParameter("StrAppendCase", adVarChar, adParamInput, 2000, strAppendCase)
        cmd.Parameters.Append prm2
        rst2.CursorLocation = adUseClient
        rst2.CursorType = adOpenForwardOnly
        rst2.CacheSize = 100
        Set rst2 = cmd.Execute
        
    Do
        Do Until rst2.EOF

            With rst2
                
                If Len(!Month) Or Len(!Year) Then
                    If Len(!Month) = 1 Then
                        Month = "0" & !Month
                        Set Record = ListView2.ListItems.Add(, , !Year & "/" & Month)
                        Month2 = Month
                        Year = !Year
                    Else
                        Set Record = ListView2.ListItems.Add(, , !Year & "/" & !Month)
                        Month2 = !Month
                        Year = !Year
                    End If
                Else
                    Set Record = ListView2.ListItems.Add(, , "" & "/" & "")
                    Month2 = ""
                    Year = ""
                End If
                
                If Len(!TotalCases) Then
                    Record.SubItems(1) = !TotalCases
                End If
                
                If rst2!NoofDay = "0" Then
                    Record.SubItems(2) = "1"
                ElseIf Len(rst2!NoofDay) Then
                    Record.SubItems(2) = rst2!NoofDay
                End If
                
                If Len(!TotalActual) Then
                    Record.SubItems(3) = Format(!TotalActual, "#,##0.00")
                Else
                    Record.SubItems(3) = "0.00"
                End If
                
                If Len(!TotalApproved) Then
                    Record.SubItems(4) = Format(!TotalApproved, "#,##0.00")
                Else
                    Record.SubItems(4) = "0.00"
                End If
                                
                Current = Current + 1
                lblCurrent = Current
            End With
            rst2.MoveNext
            DoEvents
                If intExit = 1 Then
                   Check = False
                   Exit Do
                End If
        Loop
    Loop Until Check = False
    
    If ListView2.ListItems.Count = 0 Then
        MsgBox "No record found!", vbOKOnly + vbCritical, DialogBoxTitle
        CmdClear.Enabled = True
        cmdExit.Enabled = True
        CmdPrint.Enabled = True
        lblCurrent = 0
    End If
    
    intExit = 0
    rst2.Close
    Set rst2 = Nothing
    CmdClear.Enabled = True
    cmdExit.Enabled = True
    CmdPrint.Enabled = True
    
End Sub

Private Sub ListViewHeader10()

If cboClaimCase.ListIndex = 0 And opt35 = True Then
    
    ListView2.ColumnHeaders(1).Text = "DIR Month"
    ListView2.ColumnHeaders(2).Text = "Total"
    ListView2.ColumnHeaders(2).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(3).Text = "0-15"
    ListView2.ColumnHeaders(3).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(4).Text = "15-30"
    ListView2.ColumnHeaders(4).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(5).Text = "31-60"
    ListView2.ColumnHeaders(5).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(6).Text = "61-90"
    ListView2.ColumnHeaders(6).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(7).Text = ">90"
    ListView2.ColumnHeaders(7).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(8).Text = "O/S"
    ListView2.ColumnHeaders(8).Alignment = lvwColumnRight
    For AA = 9 To 63
        ListView2.ColumnHeaders(AA).Text = ""
    Next AA

ElseIf cboClaimCase.ListIndex = 0 And opt34 = True Then
    ListView2.ColumnHeaders(1).Text = "DOR Month"
    ListView2.ColumnHeaders(2).Text = "Total"
    ListView2.ColumnHeaders(2).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(3).Text = "0-15"
    ListView2.ColumnHeaders(3).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(4).Text = "15-30"
    ListView2.ColumnHeaders(4).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(5).Text = "31-60"
    ListView2.ColumnHeaders(5).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(6).Text = "61-90"
    ListView2.ColumnHeaders(6).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(7).Text = ">90"
    ListView2.ColumnHeaders(7).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(8).Text = "O/S"
    ListView2.ColumnHeaders(8).Alignment = lvwColumnRight
    For AA = 9 To 63
        ListView2.ColumnHeaders(AA).Text = ""
    Next AA

ElseIf cboClaimCase.ListIndex = 0 And opt33 = True Then
    ListView2.ColumnHeaders(1).Text = "Invoice Date Month"
    ListView2.ColumnHeaders(2).Text = "Total"
    ListView2.ColumnHeaders(2).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(3).Text = "0-15"
    ListView2.ColumnHeaders(3).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(4).Text = "15-30"
    ListView2.ColumnHeaders(4).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(5).Text = "31-60"
    ListView2.ColumnHeaders(5).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(6).Text = "61-90"
    ListView2.ColumnHeaders(6).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(7).Text = ">90"
    ListView2.ColumnHeaders(7).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(8).Text = "O/S"
    ListView2.ColumnHeaders(8).Alignment = lvwColumnRight
    For AA = 9 To 63
        ListView2.ColumnHeaders(AA).Text = ""
    Next AA
    
ElseIf cboClaimCase.ListIndex = 0 And opt29 = True Then
    ListView2.ColumnHeaders(1).Text = "Invoice Date Month"
    ListView2.ColumnHeaders(2).Text = "Total"
    ListView2.ColumnHeaders(2).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(3).Text = "0-15"
    ListView2.ColumnHeaders(3).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(4).Text = "15-30"
    ListView2.ColumnHeaders(4).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(5).Text = "31-60"
    ListView2.ColumnHeaders(5).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(6).Text = "61-90"
    ListView2.ColumnHeaders(6).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(7).Text = ">90"
    ListView2.ColumnHeaders(7).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(8).Text = "O/S"
    ListView2.ColumnHeaders(8).Alignment = lvwColumnRight
    For AA = 9 To 63
        ListView2.ColumnHeaders(AA).Text = ""
    Next AA
    
ElseIf cboClaimCase.ListIndex = 0 And opt37 = True Then
    ListView2.ColumnHeaders(1).Text = "Reg. Month"
    ListView2.ColumnHeaders(2).Text = "B/F"
    ListView2.ColumnHeaders(2).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(3).Text = "Registered"
    ListView2.ColumnHeaders(3).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(4).Text = "Bordx"
    ListView2.ColumnHeaders(4).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(5).Text = "O/S"
    ListView2.ColumnHeaders(5).Alignment = lvwColumnRight
            
    For AA = 6 To 63
        ListView2.ColumnHeaders(AA).Text = ""
    Next AA

End If

End Sub

Private Sub DIREfficiencyRpt(Optional StrAppend As String)
Dim rst2 As New ADODB.Recordset
Dim cmd As New ADODB.Command
Dim conn As New ADODB.Command
Dim prm1 As ADODB.Parameter
Dim prm2 As ADODB.Parameter
Dim strAppendCase As String
Dim Record As ListItem
Dim Total As String
Dim Current As String

    Current = 0
    Month = ""
    ListView2.ListItems.Clear
    
    strAppendCase = "116"
    'sql = " SELECT MONTH(CLMINVRecDate) AS MONTH, Count(CLMINVRecDate) as Total " & _
          " From VIEW_CLMReportDCUNew where CLMINVRecDate IS NOT NULL"
        
    'sql2 = " GROUP BY MONTH(CLMINVRecDate)"
    
    'If Len(Trim(StrAppend)) Then
    '    sql = sql + StrAppend + sql2
    'Else
    '    sql = sql + sql2
    'End If

    'rst2.Open sql, medixmasterconnWS, adOpenForwardOnly, adLockBatchOptimistic
    Set cmd.ActiveConnection = medixmasterconnWS
        cmd.CommandText = "DCUReport"
        cmd.CommandType = adCmdStoredProc
        cmd.CommandTimeout = 300
        Set prm1 = cmd.CreateParameter("StrAppend", adVarChar, adParamInput, 2000, StrAppend)
        cmd.Parameters.Append prm1
        Set prm2 = cmd.CreateParameter("StrAppendCase", adVarChar, adParamInput, 2000, strAppendCase)
        cmd.Parameters.Append prm2
        rst2.CursorLocation = adUseClient
        rst2.CursorType = adOpenForwardOnly
        rst2.CacheSize = 100
        Set rst2 = cmd.Execute
        
    Do
        Do Until rst2.EOF

            With rst2
                
                If Len(!Month) Then
                    If !Month = 1 Then
                        Set Record = ListView2.ListItems.Add(, , "January")
                        Month = !Month
                    ElseIf !Month = 2 Then
                        Set Record = ListView2.ListItems.Add(, , "February")
                        Month = !Month
                    ElseIf !Month = 3 Then
                        Set Record = ListView2.ListItems.Add(, , "March")
                        Month = !Month
                    ElseIf !Month = 4 Then
                        Set Record = ListView2.ListItems.Add(, , "April")
                        Month = !Month
                    ElseIf !Month = 5 Then
                        Set Record = ListView2.ListItems.Add(, , "May")
                        Month = !Month
                    ElseIf !Month = 6 Then
                        Set Record = ListView2.ListItems.Add(, , "June")
                        Month = !Month
                    ElseIf !Month = 7 Then
                        Set Record = ListView2.ListItems.Add(, , "July")
                        Month = !Month
                    ElseIf !Month = 8 Then
                        Set Record = ListView2.ListItems.Add(, , "August")
                        Month = !Month
                    ElseIf !Month = 9 Then
                        Set Record = ListView2.ListItems.Add(, , "September")
                        Month = !Month
                    ElseIf !Month = 10 Then
                        Set Record = ListView2.ListItems.Add(, , "October")
                        Month = !Month
                    ElseIf !Month = 11 Then
                        Set Record = ListView2.ListItems.Add(, , "November")
                        Month = !Month
                    ElseIf !Month = 12 Then
                        Set Record = ListView2.ListItems.Add(, , "December")
                        Month = !Month
                    End If
                Else
                    Set Record = ListView2.ListItems.Add(, , "Empty")
                    Month = ""
                End If
                
                If Len(rst2!Total) Then
                   Record.SubItems(1) = rst2!Total
                Else
                    Record.SubItems(1) = "0"
                End If
            
                Call CountDIR0to15
                Record.SubItems(2) = NoInv
                
                Call CountDIR15to30
                Record.SubItems(3) = NoInv2
                
                Call CountDIR31to60
                Record.SubItems(4) = NoInv3
                
                Call CountDIR61to90
                Record.SubItems(5) = NoInv4
                
                Call CountDIR90
                Record.SubItems(6) = NoInv5
                
                TotalInv = (CDbl(Record.SubItems(1)) - CDbl(Record.SubItems(2)) - CDbl(Record.SubItems(3)) - CDbl(Record.SubItems(4)) - CDbl(Record.SubItems(5)) - CDbl(Record.SubItems(6)))
                Record.SubItems(7) = TotalInv
                
                Current = Current + 1
                lblCurrent = Current
                TotalInv = ""
                
            End With
            rst2.MoveNext
            DoEvents
                If intExit = 1 Then
                   Check = False
                   Exit Do
                End If
        Loop
    Loop Until Check = False
    
    If ListView2.ListItems.Count = 0 Then
        MsgBox "No record found!", vbOKOnly + vbCritical, DialogBoxTitle
        CmdClear.Enabled = True
        cmdExit.Enabled = True
        CmdPrint.Enabled = True
        'CmdPrintFile.Enabled = True
        lblCurrent = 0
    End If
    
    intExit = 0
    rst2.Close
    Set rst2 = Nothing
    CmdClear.Enabled = True
    cmdExit.Enabled = True
    CmdPrint.Enabled = True
End Sub

Private Sub CountDIR0to15(Optional StrAppend As String)
Dim rst3 As New ADODB.Recordset
Dim Record As ListItem
Dim sql As String
Dim sql2 As String
Dim sql3 As String
    
    NoInv = 0
               
    sql = " SELECT MONTH(CLMINVRecDate) AS MONTH, COUNT(CLMINVRecDate) AS NoofClm " & _
          " From VIEW_CLMReportDCUNew Where MONTH(CLMINVRecDate) = '" & Trim(Month) & "' AND DIRDay between 0 AND 15"
           
    sql2 = " GROUP BY MONTH(CLMINVRecDate) "
           
    If Len(Trim(StrAppend)) Then
        sql = sql + StrAppend + sql2
    Else
        sql = sql + sql2
    End If
    
    rst3.Open sql, medixmasterconnWS, adOpenForwardOnly, adLockBatchOptimistic
    
    NoInv = IIf(IsNumeric(rst3!NoofClm) = True, rst3!NoofClm, 0)
                    
    rst3.Close
    Set rst3 = Nothing
End Sub

Private Sub CountDIR15to30(Optional StrAppend As String)
Dim rst3 As New ADODB.Recordset
Dim Record As ListItem
Dim sql As String
Dim sql2 As String
Dim sql3 As String
    
    NoInv2 = 0
               
    sql = " SELECT MONTH(CLMINVRecDate) AS MONTH, COUNT(CLMINVRecDate) AS NoofClm " & _
          " From VIEW_CLMReportDCUNew Where MONTH(CLMINVRecDate) = '" & Trim(Month) & "' AND DIRDay between 15 AND 30"
           
    sql2 = " GROUP BY MONTH(CLMINVRecDate) "
           
    If Len(Trim(StrAppend)) Then
        sql = sql + StrAppend + sql2
    Else
        sql = sql + sql2
    End If
    
    rst3.Open sql, medixmasterconnWS, adOpenForwardOnly, adLockBatchOptimistic
    
    NoInv2 = IIf(IsNumeric(rst3!NoofClm) = True, rst3!NoofClm, 0)
                    
    rst3.Close
    Set rst3 = Nothing
End Sub

Private Sub CountDIR31to60(Optional StrAppend As String)
Dim rst3 As New ADODB.Recordset
Dim Record As ListItem
Dim sql As String
Dim sql2 As String
Dim sql3 As String
    
    NoInv3 = 0
               
    sql = " SELECT MONTH(CLMINVRecDate) AS MONTH, COUNT(CLMINVRecDate) AS NoofClm " & _
          " From VIEW_CLMReportDCUNew Where MONTH(CLMINVRecDate) = '" & Trim(Month) & "' AND DIRDay between 31 AND 60"
           
    sql2 = " GROUP BY MONTH(CLMINVRecDate) "
           
    If Len(Trim(StrAppend)) Then
        sql = sql + StrAppend + sql2
    Else
        sql = sql + sql2
    End If
    
    rst3.Open sql, medixmasterconnWS, adOpenForwardOnly, adLockBatchOptimistic
    
    NoInv3 = IIf(IsNumeric(rst3!NoofClm) = True, rst3!NoofClm, 0)
                    
    rst3.Close
    Set rst3 = Nothing
End Sub

Private Sub CountDIR61to90(Optional StrAppend As String)
Dim rst3 As New ADODB.Recordset
Dim Record As ListItem
Dim sql As String
Dim sql2 As String
Dim sql3 As String
    
    NoInv4 = 0
               
    sql = " SELECT MONTH(CLMINVRecDate) AS MONTH, COUNT(CLMINVRecDate) AS NoofClm " & _
          " From VIEW_CLMReportDCUNew Where MONTH(CLMINVRecDate) = '" & Trim(Month) & "' AND DIRDay between 61 AND 90"
           
    sql2 = " GROUP BY MONTH(CLMINVRecDate) "
           
    If Len(Trim(StrAppend)) Then
        sql = sql + StrAppend + sql2
    Else
        sql = sql + sql2
    End If
    
    rst3.Open sql, medixmasterconnWS, adOpenForwardOnly, adLockBatchOptimistic
    
    NoInv4 = IIf(IsNumeric(rst3!NoofClm) = True, rst3!NoofClm, 0)
                    
    rst3.Close
    Set rst3 = Nothing
End Sub

Private Sub CountDIR90(Optional StrAppend As String)
Dim rst3 As New ADODB.Recordset
Dim Record As ListItem
Dim sql As String
Dim sql2 As String
Dim sql3 As String
    
    NoInv5 = 0
               
    sql = " SELECT MONTH(CLMINVRecDate) AS MONTH, COUNT(CLMINVRecDate) AS NoofClm " & _
          " From VIEW_CLMReportDCUNew Where MONTH(CLMINVRecDate) = '" & Trim(Month) & "' AND DIRDay >90"
           
    sql2 = " GROUP BY MONTH(CLMINVRecDate) "
           
    If Len(Trim(StrAppend)) Then
        sql = sql + StrAppend + sql2
    Else
        sql = sql + sql2
    End If
    
    rst3.Open sql, medixmasterconnWS, adOpenForwardOnly, adLockBatchOptimistic
    
    NoInv5 = IIf(IsNumeric(rst3!NoofClm) = True, rst3!NoofClm, 0)
                    
    rst3.Close
    Set rst3 = Nothing
End Sub

Private Sub BordxEfficiencyRpt(Optional StrAppend As String)
Dim rst2 As New ADODB.Recordset
Dim cmd As New ADODB.Command
Dim prm1 As ADODB.Parameter
Dim prm2 As ADODB.Parameter
Dim strAppendCase As String
Dim Record As ListItem
Dim Total As String
Dim Current As String

    Current = 0
    Month = ""
    ListView2.ListItems.Clear
    
    strAppendCase = "117"
    'sql = " SELECT MONTH(BordxDate) AS MONTH, Count(BordxDate) as Total " & _
          " From VIEW_CLMReportDCU where BordxDate IS NOT NULL"
        
    'sql2 = " GROUP BY MONTH(BordxDate)"
    
    'If Len(Trim(StrAppend)) Then
    '    sql = sql + StrAppend + sql2
    'Else
    '    sql = sql + sql2
    'End If

    'rst2.Open sql, medixmasterconnWS, adOpenForwardOnly, adLockBatchOptimistic
    Set cmd.ActiveConnection = medixmasterconnWS
        cmd.CommandText = "DCUReport"
        cmd.CommandType = adCmdStoredProc
        cmd.CommandTimeout = 300
        Set prm1 = cmd.CreateParameter("StrAppend", adVarChar, adParamInput, 2000, StrAppend)
        cmd.Parameters.Append prm1
        Set prm2 = cmd.CreateParameter("StrAppendCase", adVarChar, adParamInput, 2000, strAppendCase)
        cmd.Parameters.Append prm2
        rst2.CursorLocation = adUseClient
        rst2.CursorType = adOpenForwardOnly
        rst2.CacheSize = 100
        Set rst2 = cmd.Execute
        
    Do
        Do Until rst2.EOF

            With rst2
                
                If Len(!Month) Then
                    If !Month = 1 Then
                        Set Record = ListView2.ListItems.Add(, , "January")
                        Month = !Month
                    ElseIf !Month = 2 Then
                        Set Record = ListView2.ListItems.Add(, , "February")
                        Month = !Month
                    ElseIf !Month = 3 Then
                        Set Record = ListView2.ListItems.Add(, , "March")
                        Month = !Month
                    ElseIf !Month = 4 Then
                        Set Record = ListView2.ListItems.Add(, , "April")
                        Month = !Month
                    ElseIf !Month = 5 Then
                        Set Record = ListView2.ListItems.Add(, , "May")
                        Month = !Month
                    ElseIf !Month = 6 Then
                        Set Record = ListView2.ListItems.Add(, , "June")
                        Month = !Month
                    ElseIf !Month = 7 Then
                        Set Record = ListView2.ListItems.Add(, , "July")
                        Month = !Month
                    ElseIf !Month = 8 Then
                        Set Record = ListView2.ListItems.Add(, , "August")
                        Month = !Month
                    ElseIf !Month = 9 Then
                        Set Record = ListView2.ListItems.Add(, , "September")
                        Month = !Month
                    ElseIf !Month = 10 Then
                        Set Record = ListView2.ListItems.Add(, , "October")
                        Month = !Month
                    ElseIf !Month = 11 Then
                        Set Record = ListView2.ListItems.Add(, , "November")
                        Month = !Month
                    ElseIf !Month = 12 Then
                        Set Record = ListView2.ListItems.Add(, , "December")
                        Month = !Month
                    End If
                Else
                    Set Record = ListView2.ListItems.Add(, , "Empty")
                    Month = ""
                End If
                
                If Len(rst2!Total) Then
                   Record.SubItems(1) = rst2!Total
                Else
                    Record.SubItems(1) = "0"
                End If
            
                Call CountBordx0to15
                Record.SubItems(2) = NoInv
                
                Call CountBordx15to30
                Record.SubItems(3) = NoInv2
                
                Call CountBordx31to60
                Record.SubItems(4) = NoInv3
                
                Call CountBordx61to90
                Record.SubItems(5) = NoInv4
                
                Call CountBordx90
                Record.SubItems(6) = NoInv5
                
                TotalInv = (CDbl(Record.SubItems(1)) - CDbl(Record.SubItems(2)) - CDbl(Record.SubItems(3)) - CDbl(Record.SubItems(4)) - CDbl(Record.SubItems(5)) - CDbl(Record.SubItems(6)))
                Record.SubItems(7) = TotalInv
                
                Current = Current + 1
                lblCurrent = Current
                TotalInv = ""
                
            End With
            rst2.MoveNext
            DoEvents
                If intExit = 1 Then
                   Check = False
                   Exit Do
                End If
        Loop
    Loop Until Check = False
    
    If ListView2.ListItems.Count = 0 Then
        MsgBox "No record found!", vbOKOnly + vbCritical, DialogBoxTitle
        CmdClear.Enabled = True
        cmdExit.Enabled = True
        CmdPrint.Enabled = True
        lblCurrent = 0
    End If
    
    intExit = 0
    rst2.Close
    Set rst2 = Nothing
    CmdClear.Enabled = True
    cmdExit.Enabled = True
    CmdPrint.Enabled = True
End Sub

Private Sub CountBordx0to15(Optional StrAppend As String)
Dim rst3 As New ADODB.Recordset
Dim Record As ListItem
Dim sql As String
Dim sql2 As String
Dim sql3 As String
    
    NoInv = 0
               
    sql = " SELECT MONTH(BordxDate) AS MONTH, COUNT(BordxDate) AS NoofClm " & _
          " From VIEW_CLMReportDCU Where MONTH(BordxDate) = '" & Trim(Month) & "' AND BordxDay between 0 AND 15"
           
    sql2 = " GROUP BY MONTH(BordxDate) "
           
    If Len(Trim(StrAppend)) Then
        sql = sql + StrAppend + sql2
    Else
        sql = sql + sql2
    End If
    
    rst3.Open sql, medixmasterconnWS, adOpenForwardOnly, adLockBatchOptimistic
    
    NoInv = IIf(IsNumeric(rst3!NoofClm) = True, rst3!NoofClm, 0)
                    
    rst3.Close
    Set rst3 = Nothing
End Sub

Private Sub CountBordx15to30(Optional StrAppend As String)
Dim rst3 As New ADODB.Recordset
Dim Record As ListItem
Dim sql As String
Dim sql2 As String
Dim sql3 As String
    
    NoInv2 = 0
               
    sql = " SELECT MONTH(BordxDate) AS MONTH, COUNT(BordxDate) AS NoofClm " & _
          " From VIEW_CLMReportDCU Where MONTH(BordxDate) = '" & Trim(Month) & "' AND BordxDay between 15 AND 30"
           
    sql2 = " GROUP BY MONTH(BordxDate) "
           
    If Len(Trim(StrAppend)) Then
        sql = sql + StrAppend + sql2
    Else
        sql = sql + sql2
    End If
    
    rst3.Open sql, medixmasterconnWS, adOpenForwardOnly, adLockBatchOptimistic
    
    NoInv2 = IIf(IsNumeric(rst3!NoofClm) = True, rst3!NoofClm, 0)
                    
    rst3.Close
    Set rst3 = Nothing
End Sub

Private Sub CountBordx31to60(Optional StrAppend As String)
Dim rst3 As New ADODB.Recordset
Dim Record As ListItem
Dim sql As String
Dim sql2 As String
Dim sql3 As String
    
    NoInv3 = 0
               
    sql = " SELECT MONTH(BordxDate) AS MONTH, COUNT(BordxDate) AS NoofClm " & _
          " From VIEW_CLMReportDCU Where MONTH(BordxDate) = '" & Trim(Month) & "' AND BordxDay between 31 AND 60"
           
    sql2 = " GROUP BY MONTH(BordxDate) "
           
    If Len(Trim(StrAppend)) Then
        sql = sql + StrAppend + sql2
    Else
        sql = sql + sql2
    End If
    
    rst3.Open sql, medixmasterconnWS, adOpenForwardOnly, adLockBatchOptimistic
    
    NoInv3 = IIf(IsNumeric(rst3!NoofClm) = True, rst3!NoofClm, 0)
                    
    rst3.Close
    Set rst3 = Nothing
End Sub

Private Sub CountBordx61to90(Optional StrAppend As String)
Dim rst3 As New ADODB.Recordset
Dim Record As ListItem
Dim sql As String
Dim sql2 As String
Dim sql3 As String
    
    NoInv4 = 0
               
    sql = " SELECT MONTH(BordxDate) AS MONTH, COUNT(BordxDate) AS NoofClm " & _
          " From VIEW_CLMReportDCU Where MONTH(BordxDate) = '" & Trim(Month) & "' AND BordxDay between 61 AND 90"
           
    sql2 = " GROUP BY MONTH(BordxDate) "
           
    If Len(Trim(StrAppend)) Then
        sql = sql + StrAppend + sql2
    Else
        sql = sql + sql2
    End If
    
    rst3.Open sql, medixmasterconnWS, adOpenForwardOnly, adLockBatchOptimistic
    
    NoInv4 = IIf(IsNumeric(rst3!NoofClm) = True, rst3!NoofClm, 0)
                    
    rst3.Close
    Set rst3 = Nothing
End Sub

Private Sub CountBordx90(Optional StrAppend As String)
Dim rst3 As New ADODB.Recordset
Dim Record As ListItem
Dim sql As String
Dim sql2 As String
Dim sql3 As String
    
    NoInv5 = 0
               
    sql = " SELECT MONTH(BordxDate) AS MONTH, COUNT(BordxDate) AS NoofClm " & _
          " From VIEW_CLMReportDCU Where MONTH(BordxDate) = '" & Trim(Month) & "' AND BordxDay >90"
           
    sql2 = " GROUP BY MONTH(BordxDate) "
           
    If Len(Trim(StrAppend)) Then
        sql = sql + StrAppend + sql2
    Else
        sql = sql + sql2
    End If
    
    rst3.Open sql, medixmasterconnWS, adOpenForwardOnly, adLockBatchOptimistic
    
    NoInv5 = IIf(IsNumeric(rst3!NoofClm) = True, rst3!NoofClm, 0)
                    
    rst3.Close
    Set rst3 = Nothing
End Sub

Private Sub InvEfficiencyRpt(Optional StrAppend As String)
Dim rst2 As New ADODB.Recordset
Dim cmd As New ADODB.Command
Dim prm1 As ADODB.Parameter
Dim prm2 As ADODB.Parameter
Dim strAppendCase As String
Dim Record As ListItem
Dim Total As String
Dim Current As String

    Current = 0
    Month = ""
    ListView2.ListItems.Clear
      
     strAppendCase = "118"
    'sql = " SELECT MONTH(INVEnteredDate) AS MONTH, Count(INVEnteredDate) as Total " & _
          " From VIEW_CLMReportDCUNew where INVEnteredDate IS NOT NULL"
        
    'sql2 = " GROUP BY MONTH(INVEnteredDate)"
    
    'If Len(Trim(StrAppend)) Then
    '    sql = sql + StrAppend + sql2
    'Else
    '    sql = sql + sql2
    'End If

    'rst2.Open sql, medixmasterconnWS, adOpenForwardOnly, adLockBatchOptimistic
    Set cmd.ActiveConnection = medixmasterconnWS
        cmd.CommandText = "DCUReport"
        cmd.CommandType = adCmdStoredProc
        cmd.CommandTimeout = 300
        Set prm1 = cmd.CreateParameter("StrAppend", adVarChar, adParamInput, 2000, StrAppend)
        cmd.Parameters.Append prm1
        Set prm2 = cmd.CreateParameter("StrAppendCase", adVarChar, adParamInput, 2000, strAppendCase)
        cmd.Parameters.Append prm2
        rst2.CursorLocation = adUseClient
        rst2.CursorType = adOpenForwardOnly
        rst2.CacheSize = 100
        Set rst2 = cmd.Execute
        
    Do
        Do Until rst2.EOF

            With rst2
                
                If Len(!Month) Then
                    If !Month = 1 Then
                        Set Record = ListView2.ListItems.Add(, , "January")
                        Month = !Month
                    ElseIf !Month = 2 Then
                        Set Record = ListView2.ListItems.Add(, , "February")
                        Month = !Month
                    ElseIf !Month = 3 Then
                        Set Record = ListView2.ListItems.Add(, , "March")
                        Month = !Month
                    ElseIf !Month = 4 Then
                        Set Record = ListView2.ListItems.Add(, , "April")
                        Month = !Month
                    ElseIf !Month = 5 Then
                        Set Record = ListView2.ListItems.Add(, , "May")
                        Month = !Month
                    ElseIf !Month = 6 Then
                        Set Record = ListView2.ListItems.Add(, , "June")
                        Month = !Month
                    ElseIf !Month = 7 Then
                        Set Record = ListView2.ListItems.Add(, , "July")
                        Month = !Month
                    ElseIf !Month = 8 Then
                        Set Record = ListView2.ListItems.Add(, , "August")
                        Month = !Month
                    ElseIf !Month = 9 Then
                        Set Record = ListView2.ListItems.Add(, , "September")
                        Month = !Month
                    ElseIf !Month = 10 Then
                        Set Record = ListView2.ListItems.Add(, , "October")
                        Month = !Month
                    ElseIf !Month = 11 Then
                        Set Record = ListView2.ListItems.Add(, , "November")
                        Month = !Month
                    ElseIf !Month = 12 Then
                        Set Record = ListView2.ListItems.Add(, , "December")
                        Month = !Month
                    End If
                Else
                    Set Record = ListView2.ListItems.Add(, , "Empty")
                    Month = ""
                End If
                
                If Len(rst2!Total) Then
                   Record.SubItems(1) = rst2!Total
                Else
                    Record.SubItems(1) = "0"
                End If
            
                Call CountInv0to15
                Record.SubItems(2) = NoInv
                
                Call CountInv15to30
                Record.SubItems(3) = NoInv2
                
                Call CountInv31to60
                Record.SubItems(4) = NoInv3
                
                Call CountInv61to90
                Record.SubItems(5) = NoInv4
                
                Call CountInv90
                Record.SubItems(6) = NoInv5
                
                TotalInv = (CDbl(Record.SubItems(1)) - CDbl(Record.SubItems(2)) - CDbl(Record.SubItems(3)) - CDbl(Record.SubItems(4)) - CDbl(Record.SubItems(5)) - CDbl(Record.SubItems(6)))
                Record.SubItems(7) = TotalInv
                
                Current = Current + 1
                lblCurrent = Current
                TotalInv = ""
                
            End With
            rst2.MoveNext
            DoEvents
                If intExit = 1 Then
                   Check = False
                   Exit Do
                End If
        Loop
    Loop Until Check = False
    
    If ListView2.ListItems.Count = 0 Then
        MsgBox "No record found!", vbOKOnly + vbCritical, DialogBoxTitle
        CmdClear.Enabled = True
        cmdExit.Enabled = True
        CmdPrint.Enabled = True
        'CmdPrintFile.Enabled = True
        lblCurrent = 0
    End If
    
    intExit = 0
    rst2.Close
    Set rst2 = Nothing
    CmdClear.Enabled = True
    cmdExit.Enabled = True
    CmdPrint.Enabled = True
End Sub

Private Sub CountInv0to15(Optional StrAppend As String)
Dim rst3 As New ADODB.Recordset
Dim Record As ListItem
Dim sql As String
Dim sql2 As String
Dim sql3 As String
    
    NoInv = 0
               
    sql = " SELECT MONTH(INVEnteredDate) AS MONTH, COUNT(INVEnteredDate) AS NoofClm " & _
          " From VIEW_CLMReportDCUNew Where MONTH(INVEnteredDate) = '" & Trim(Month) & "' AND InvDay between 0 AND 15"
           
    sql2 = " GROUP BY MONTH(INVEnteredDate) "
           
    If Len(Trim(StrAppend)) Then
        sql = sql + StrAppend + sql2
    Else
        sql = sql + sql2
    End If
    
    rst3.Open sql, medixmasterconnWS, adOpenForwardOnly, adLockBatchOptimistic
    
    NoInv = IIf(IsNumeric(rst3!NoofClm) = True, rst3!NoofClm, 0)
                    
    rst3.Close
    Set rst3 = Nothing
End Sub

Private Sub CountInv15to30(Optional StrAppend As String)
Dim rst3 As New ADODB.Recordset
Dim Record As ListItem
Dim sql As String
Dim sql2 As String
Dim sql3 As String
    
    NoInv2 = 0
               
    sql = " SELECT MONTH(INVEnteredDate) AS MONTH, COUNT(INVEnteredDate) AS NoofClm " & _
          " From VIEW_CLMReportDCUNew Where MONTH(INVEnteredDate) = '" & Trim(Month) & "' AND InvDay between 15 AND 30"
           
    sql2 = " GROUP BY MONTH(INVEnteredDate) "
           
    If Len(Trim(StrAppend)) Then
        sql = sql + StrAppend + sql2
    Else
        sql = sql + sql2
    End If
    
    rst3.Open sql, medixmasterconnWS, adOpenForwardOnly, adLockBatchOptimistic
    
    NoInv2 = IIf(IsNumeric(rst3!NoofClm) = True, rst3!NoofClm, 0)
                    
    rst3.Close
    Set rst3 = Nothing
End Sub

Private Sub CountInv31to60(Optional StrAppend As String)
Dim rst3 As New ADODB.Recordset
Dim Record As ListItem
Dim sql As String
Dim sql2 As String
Dim sql3 As String
    
    NoInv3 = 0
               
    sql = " SELECT MONTH(INVEnteredDate) AS MONTH, COUNT(INVEnteredDate) AS NoofClm " & _
          " From VIEW_CLMReportDCUNew Where MONTH(INVEnteredDate) = '" & Trim(Month) & "' AND InvDay between 31 AND 60"
           
    sql2 = " GROUP BY MONTH(INVEnteredDate) "
           
    If Len(Trim(StrAppend)) Then
        sql = sql + StrAppend + sql2
    Else
        sql = sql + sql2
    End If
    
    rst3.Open sql, medixmasterconnWS, adOpenForwardOnly, adLockBatchOptimistic
    
    NoInv3 = IIf(IsNumeric(rst3!NoofClm) = True, rst3!NoofClm, 0)
                    
    rst3.Close
    Set rst3 = Nothing
End Sub

Private Sub CountInv61to90(Optional StrAppend As String)
Dim rst3 As New ADODB.Recordset
Dim Record As ListItem
Dim sql As String
Dim sql2 As String
Dim sql3 As String
    
    NoInv4 = 0
               
    sql = " SELECT MONTH(INVEnteredDate) AS MONTH, COUNT(INVEnteredDate) AS NoofClm " & _
          " From VIEW_CLMReportDCUNew Where MONTH(INVEnteredDate) = '" & Trim(Month) & "' AND InvDay between 61 AND 90"
           
    sql2 = " GROUP BY MONTH(INVEnteredDate) "
           
    If Len(Trim(StrAppend)) Then
        sql = sql + StrAppend + sql2
    Else
        sql = sql + sql2
    End If
    
    rst3.Open sql, medixmasterconnWS, adOpenForwardOnly, adLockBatchOptimistic
    
    NoInv4 = IIf(IsNumeric(rst3!NoofClm) = True, rst3!NoofClm, 0)
                    
    rst3.Close
    Set rst3 = Nothing
End Sub

Private Sub CountInv90(Optional StrAppend As String)
Dim rst3 As New ADODB.Recordset
Dim Record As ListItem
Dim sql As String
Dim sql2 As String
Dim sql3 As String
    
    NoInv5 = 0
               
    sql = " SELECT MONTH(INVEnteredDate) AS MONTH, COUNT(INVEnteredDate) AS NoofClm " & _
          " From VIEW_CLMReportDCUNew Where MONTH(INVEnteredDate) = '" & Trim(Month) & "' AND InvDay >90"
           
    sql2 = " GROUP BY MONTH(INVEnteredDate) "
           
    If Len(Trim(StrAppend)) Then
        sql = sql + StrAppend + sql2
    Else
        sql = sql + sql2
    End If
    
    rst3.Open sql, medixmasterconnWS, adOpenForwardOnly, adLockBatchOptimistic
    
    NoInv5 = IIf(IsNumeric(rst3!NoofClm) = True, rst3!NoofClm, 0)
                    
    rst3.Close
    Set rst3 = Nothing
End Sub

Private Sub BordxInvEfficiencyRpt(Optional StrAppend As String)
Dim rst2 As New ADODB.Recordset
Dim cmd As New ADODB.Command
Dim prm1 As ADODB.Parameter
Dim prm2 As ADODB.Parameter
Dim strAppendCase As String
Dim Record As ListItem
Dim Total As String
Dim Current As String

    Current = 0
    Month = ""
    ListView2.ListItems.Clear
      
    strAppendCase = "119"
    'sql = " SELECT MONTH(BordxDate) AS MONTH, Count(BordxDate) as Total " & _
          " From VIEW_CLMReportDCUNew where BordxDate IS NOT NULL"
    
    'sql2 = " GROUP BY MONTH(BordxDate)"
    
    'If Len(Trim(StrAppend)) Then
    '    sql = sql + StrAppend + sql2
    'Else
    '    sql = sql + sql2
    'End If

    'rst2.Open sql, medixmasterconnWS, adOpenForwardOnly, adLockBatchOptimistic
    Set cmd.ActiveConnection = medixmasterconnWS
        cmd.CommandText = "DCUReport"
        cmd.CommandType = adCmdStoredProc
        cmd.CommandTimeout = 300
        Set prm1 = cmd.CreateParameter("StrAppend", adVarChar, adParamInput, 2000, StrAppend)
        cmd.Parameters.Append prm1
        Set prm2 = cmd.CreateParameter("StrAppendCase", adVarChar, adParamInput, 2000, strAppendCase)
        cmd.Parameters.Append prm2
        rst2.CursorLocation = adUseClient
        rst2.CursorType = adOpenForwardOnly
        rst2.CacheSize = 100
        Set rst2 = cmd.Execute
    
    Do
        Do Until rst2.EOF

            With rst2
                
                If Len(!Month) Then
                    If !Month = 1 Then
                        Set Record = ListView2.ListItems.Add(, , "January")
                        Month = !Month
                    ElseIf !Month = 2 Then
                        Set Record = ListView2.ListItems.Add(, , "February")
                        Month = !Month
                    ElseIf !Month = 3 Then
                        Set Record = ListView2.ListItems.Add(, , "March")
                        Month = !Month
                    ElseIf !Month = 4 Then
                        Set Record = ListView2.ListItems.Add(, , "April")
                        Month = !Month
                    ElseIf !Month = 5 Then
                        Set Record = ListView2.ListItems.Add(, , "May")
                        Month = !Month
                    ElseIf !Month = 6 Then
                        Set Record = ListView2.ListItems.Add(, , "June")
                        Month = !Month
                    ElseIf !Month = 7 Then
                        Set Record = ListView2.ListItems.Add(, , "July")
                        Month = !Month
                    ElseIf !Month = 8 Then
                        Set Record = ListView2.ListItems.Add(, , "August")
                        Month = !Month
                    ElseIf !Month = 9 Then
                        Set Record = ListView2.ListItems.Add(, , "September")
                        Month = !Month
                    ElseIf !Month = 10 Then
                        Set Record = ListView2.ListItems.Add(, , "October")
                        Month = !Month
                    ElseIf !Month = 11 Then
                        Set Record = ListView2.ListItems.Add(, , "November")
                        Month = !Month
                    ElseIf !Month = 12 Then
                        Set Record = ListView2.ListItems.Add(, , "December")
                        Month = !Month
                    End If
                Else
                    Set Record = ListView2.ListItems.Add(, , "Empty")
                    Month = ""
                End If
                
                If Len(rst2!Total) Then
                   Record.SubItems(1) = rst2!Total
                Else
                    Record.SubItems(1) = "0"
                End If
            
                Call CountBordxInv0to15
                Record.SubItems(2) = NoInv
                
                Call CountBordxInv15to30
                Record.SubItems(3) = NoInv2
                
                Call CountBordxInv31to60
                Record.SubItems(4) = NoInv3
                
                Call CountBordxInv61to90
                Record.SubItems(5) = NoInv4
                
                Call CountBordxInv90
                Record.SubItems(6) = NoInv5
                
                TotalInv = (CDbl(Record.SubItems(1)) - CDbl(Record.SubItems(2)) - CDbl(Record.SubItems(3)) - CDbl(Record.SubItems(4)) - CDbl(Record.SubItems(5)) - CDbl(Record.SubItems(6)))
                Record.SubItems(7) = TotalInv
                
                Current = Current + 1
                lblCurrent = Current
                TotalInv = ""
                
            End With
            rst2.MoveNext
            DoEvents
                If intExit = 1 Then
                   Check = False
                   Exit Do
                End If
        Loop
    Loop Until Check = False
    
    If ListView2.ListItems.Count = 0 Then
        MsgBox "No record found!", vbOKOnly + vbCritical, DialogBoxTitle
        CmdClear.Enabled = True
        cmdExit.Enabled = True
        CmdPrint.Enabled = True
        'CmdPrintFile.Enabled = True
        lblCurrent = 0
    End If
    
    intExit = 0
    rst2.Close
    Set rst2 = Nothing
    CmdClear.Enabled = True
    cmdExit.Enabled = True
    CmdPrint.Enabled = True
End Sub

Private Sub CountBordxInv0to15(Optional StrAppend As String)
Dim rst3 As New ADODB.Recordset
Dim Record As ListItem
Dim sql As String
Dim sql2 As String
Dim sql3 As String
    
    NoInv = 0
               
    sql = " SELECT MONTH(BordxDate) AS MONTH, COUNT(BordxDate) AS NoofClm " & _
          " From VIEW_CLMReportDCUNew Where MONTH(BordxDate) = '" & Trim(Month) & "' AND BordxDay between 0 AND 15"
           
    sql2 = " GROUP BY MONTH(BordxDate) "
           
    If Len(Trim(StrAppend)) Then
        sql = sql + StrAppend + sql2
    Else
        sql = sql + sql2
    End If
    
    rst3.Open sql, medixmasterconnWS, adOpenForwardOnly, adLockBatchOptimistic
    
    NoInv = IIf(IsNumeric(rst3!NoofClm) = True, rst3!NoofClm, 0)
                    
    rst3.Close
    Set rst3 = Nothing
End Sub

Private Sub CountBordxInv15to30(Optional StrAppend As String)
Dim rst3 As New ADODB.Recordset
Dim Record As ListItem
Dim sql As String
Dim sql2 As String
Dim sql3 As String
    
    NoInv2 = 0
               
    sql = " SELECT MONTH(BordxDate) AS MONTH, COUNT(BordxDate) AS NoofClm " & _
          " From VIEW_CLMReportDCUNew Where MONTH(BordxDate) = '" & Trim(Month) & "' AND BordxDay between 15 AND 30"
           
    sql2 = " GROUP BY MONTH(BordxDate) "
           
    If Len(Trim(StrAppend)) Then
        sql = sql + StrAppend + sql2
    Else
        sql = sql + sql2
    End If
    
    rst3.Open sql, medixmasterconnWS, adOpenForwardOnly, adLockBatchOptimistic
    
    NoInv2 = IIf(IsNumeric(rst3!NoofClm) = True, rst3!NoofClm, 0)
                    
    rst3.Close
    Set rst3 = Nothing
End Sub

Private Sub CountBordxInv31to60(Optional StrAppend As String)
Dim rst3 As New ADODB.Recordset
Dim Record As ListItem
Dim sql As String
Dim sql2 As String
Dim sql3 As String
    
    NoInv3 = 0
               
    sql = " SELECT MONTH(BordxDate) AS MONTH, COUNT(BordxDate) AS NoofClm " & _
          " From VIEW_CLMReportDCUNew Where MONTH(BordxDate) = '" & Trim(Month) & "' AND BordxDay between 31 AND 60"
           
    sql2 = " GROUP BY MONTH(BordxDate) "
           
    If Len(Trim(StrAppend)) Then
        sql = sql + StrAppend + sql2
    Else
        sql = sql + sql2
    End If
    
    rst3.Open sql, medixmasterconnWS, adOpenForwardOnly, adLockBatchOptimistic
    
    NoInv3 = IIf(IsNumeric(rst3!NoofClm) = True, rst3!NoofClm, 0)
                    
    rst3.Close
    Set rst3 = Nothing
End Sub

Private Sub CountBordxInv61to90(Optional StrAppend As String)
Dim rst3 As New ADODB.Recordset
Dim Record As ListItem
Dim sql As String
Dim sql2 As String
Dim sql3 As String
    
    NoInv4 = 0
               
    sql = " SELECT MONTH(BordxDate) AS MONTH, COUNT(BordxDate) AS NoofClm " & _
          " From VIEW_CLMReportDCUNew Where MONTH(BordxDate) = '" & Trim(Month) & "' AND BordxDay between 61 AND 90"
           
    sql2 = " GROUP BY MONTH(BordxDate) "
           
    If Len(Trim(StrAppend)) Then
        sql = sql + StrAppend + sql2
    Else
        sql = sql + sql2
    End If
    
    rst3.Open sql, medixmasterconnWS, adOpenForwardOnly, adLockBatchOptimistic
    
    NoInv4 = IIf(IsNumeric(rst3!NoofClm) = True, rst3!NoofClm, 0)
                    
    rst3.Close
    Set rst3 = Nothing
End Sub

Private Sub CountBordxInv90(Optional StrAppend As String)
Dim rst3 As New ADODB.Recordset
Dim Record As ListItem
Dim sql As String
Dim sql2 As String
Dim sql3 As String
    
    NoInv5 = 0
               
    sql = " SELECT MONTH(BordxDate) AS MONTH, COUNT(BordxDate) AS NoofClm " & _
          " From VIEW_CLMReportDCUNew Where MONTH(BordxDate) = '" & Trim(Month) & "' AND BordxDay >90"
           
    sql2 = " GROUP BY MONTH(BordxDate) "
           
    If Len(Trim(StrAppend)) Then
        sql = sql + StrAppend + sql2
    Else
        sql = sql + sql2
    End If
    
    rst3.Open sql, medixmasterconnWS, adOpenForwardOnly, adLockBatchOptimistic
    
    NoInv5 = IIf(IsNumeric(rst3!NoofClm) = True, rst3!NoofClm, 0)
                    
    rst3.Close
    Set rst3 = Nothing
End Sub

Private Sub OSBordxEfficiencyRpt(Optional StrAppend As String)
Dim rst2 As New ADODB.Recordset
Dim cmd As New ADODB.Command
Dim prm1 As ADODB.Parameter
Dim prm2 As ADODB.Parameter
Dim strAppendCase As String
Dim Record As ListItem
Dim Total As String
Dim Current As String

    Current = 0
    NoInv2 = 0
    Month = ""
    Year = ""
    ListView2.ListItems.Clear

    strAppendCase = "120"
    'sql = " SELECT YEAR(ClaimOpenDate) AS Year, MONTH(ClaimOpenDate) AS MONTH, Count(CASNumber) as Total " & _
          " From VIEW_CLMReportDCU where ClaimOpenDate IS NOT NULL"
    
    'sql2 = " GROUP BY YEAR(ClaimOpenDate), MONTH(ClaimOpenDate)"
    
    'If Len(Trim(StrAppend)) Then
    '    sql = sql + StrAppend + sql2
    'Else
    '    sql = sql + sql2
    'End If

    'rst2.Open sql, medixmasterconnWS, adOpenForwardOnly, adLockBatchOptimistic
    Set cmd.ActiveConnection = medixmasterconnWS
        cmd.CommandText = "DCUReport"
        cmd.CommandType = adCmdStoredProc
        cmd.CommandTimeout = 300
        Set prm1 = cmd.CreateParameter("StrAppend", adVarChar, adParamInput, 2000, StrAppend)
        cmd.Parameters.Append prm1
        Set prm2 = cmd.CreateParameter("StrAppendCase", adVarChar, adParamInput, 2000, strAppendCase)
        cmd.Parameters.Append prm2
        rst2.CursorLocation = adUseClient
        rst2.CursorType = adOpenForwardOnly
        rst2.CacheSize = 100
        Set rst2 = cmd.Execute
        
    Do
        Do Until rst2.EOF

            With rst2
                
                If Len(!Month & "/" & !Year) Then
                    If Len(!Month) = 1 Then
                        Month = "0" & !Month
                        Set Record = ListView2.ListItems.Add(, , Trim(!Year & "/" & Month))
                        Month = Trim(!Month)
                        Year = Trim(!Year)
                    Else
                        Set Record = ListView2.ListItems.Add(, , Trim(!Year & "/" & !Month))
                        Month = Trim(!Month)
                        Year = Trim(!Year)
                    End If
                Else
                    Set Record = ListView2.ListItems.Add(, , "Empty")
                    Month = ""
                    Year = ""
                End If
                
                If NoInv2 = 0 Then
                    Record.SubItems(1) = "0"
                Else
                    Record.SubItems(1) = NoInv2
                End If
                
                If Len(rst2!Total) Then
                   Record.SubItems(2) = rst2!Total
                Else
                    Record.SubItems(2) = "0"
                End If
                            
                Call CountOSBordx
                Record.SubItems(3) = NoInv
                
                TotalInv = (CDbl(Record.SubItems(1)) + CDbl(Record.SubItems(2)) - CDbl(Record.SubItems(3)))
                Record.SubItems(4) = TotalInv
                NoInv2 = Record.SubItems(4)
                
                Current = Current + 1
                lblCurrent = Current
                TotalInv = ""
                
            End With
            rst2.MoveNext
            DoEvents
                If intExit = 1 Then
                   Check = False
                   Exit Do
                End If
        Loop
    Loop Until Check = False
    
    If ListView2.ListItems.Count = 0 Then
        MsgBox "No record found!", vbOKOnly + vbCritical, DialogBoxTitle
        CmdClear.Enabled = True
        cmdExit.Enabled = True
        CmdPrint.Enabled = True
        'CmdPrintFile.Enabled = True
        lblCurrent = 0
    End If
    
    intExit = 0
    rst2.Close
    Set rst2 = Nothing
    CmdClear.Enabled = True
    cmdExit.Enabled = True
    CmdPrint.Enabled = True
End Sub

Private Sub CountOSBordx(Optional StrAppend As String)
Dim rst3 As New ADODB.Recordset
Dim Record As ListItem
Dim sql As String
Dim sql2 As String
Dim sql3 As String
    
    NoInv = 0
               
    sql = " SELECT Year(BordxDate) AS YEAR, MONTH(BordxDate) AS MONTH, COUNT(CASNumber) AS NoofClm " & _
          " From VIEW_CLMReportDCU Where MONTH(BordxDate) = '" & Trim(Month) & "' AND YEAR(BordxDate) = '" & Trim(Year) & "'"
           
    sql2 = " GROUP BY MONTH(BordxDate), YEAR(BordxDate) "
           
    If Len(Trim(StrAppend)) Then
        sql = sql + StrAppend + sql2
    Else
        sql = sql + sql2
    End If
    
    rst3.Open sql, medixmasterconnWS, adOpenForwardOnly, adLockBatchOptimistic
    
    NoInv = IIf(IsNumeric(rst3!NoofClm) = True, rst3!NoofClm, 0)
                    
    rst3.Close
    Set rst3 = Nothing
End Sub


