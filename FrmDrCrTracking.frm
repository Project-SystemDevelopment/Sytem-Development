VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "CRYSTL32.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form FrmDrCrTracking 
   Caption         =   "Debit Credit Note Tracking"
   ClientHeight    =   8595
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   11880
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   MDIChild        =   -1  'True
   ScaleHeight     =   8595
   ScaleWidth      =   11880
   WindowState     =   2  'Maximized
   Begin VB.Frame Frame1 
      Height          =   5385
      Left            =   120
      TabIndex        =   97
      Top             =   240
      Width           =   11655
      Begin VB.OptionButton OptCredit 
         Caption         =   "Credit Note"
         Height          =   255
         Left            =   2880
         TabIndex        =   156
         Top             =   240
         Width           =   1215
      End
      Begin VB.OptionButton OptDebit 
         Caption         =   "Debit Note"
         Height          =   255
         Left            =   1680
         TabIndex        =   155
         Top             =   240
         Value           =   -1  'True
         Width           =   1215
      End
      Begin VB.CheckBox ChkPayor 
         Caption         =   "Check3"
         Height          =   255
         Left            =   8280
         TabIndex        =   42
         Top             =   825
         Width           =   255
      End
      Begin VB.CheckBox ChkCr 
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
         Left            =   3960
         TabIndex        =   25
         Top             =   3480
         Width           =   255
      End
      Begin VB.CheckBox ChkDr 
         Caption         =   "Check3"
         Height          =   255
         Left            =   3960
         TabIndex        =   22
         Top             =   3240
         Width           =   255
      End
      Begin VB.CheckBox ChkMemNo 
         Caption         =   "Check3"
         Height          =   255
         Left            =   3960
         TabIndex        =   17
         Top             =   2720
         Width           =   255
      End
      Begin VB.CheckBox ChkPVDate 
         Caption         =   "Check3"
         Height          =   255
         Left            =   3960
         TabIndex        =   34
         Top             =   4200
         Width           =   255
      End
      Begin VB.CheckBox ChkRefNo 
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
         Left            =   3960
         TabIndex        =   37
         Top             =   4440
         Width           =   255
      End
      Begin VB.CheckBox ChkPatient 
         Caption         =   "Check3"
         Height          =   255
         Left            =   3960
         TabIndex        =   19
         Top             =   2970
         Width           =   255
      End
      Begin VB.CheckBox ChkPVNo 
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
         Left            =   3960
         TabIndex        =   31
         Top             =   3960
         Width           =   255
      End
      Begin VB.CheckBox ChkAmt 
         Caption         =   "Check3"
         Height          =   255
         Left            =   3960
         TabIndex        =   28
         Top             =   3740
         Width           =   255
      End
      Begin VB.CheckBox ChkRecDate 
         Caption         =   "Check3"
         Height          =   255
         Left            =   3960
         TabIndex        =   40
         Top             =   4680
         Width           =   255
      End
      Begin VB.CheckBox ChkCaseStatus 
         Caption         =   "Check3"
         Height          =   255
         Left            =   3945
         TabIndex        =   9
         Top             =   1695
         Width           =   255
      End
      Begin VB.CheckBox ChkWksStatus 
         Caption         =   "Check3"
         Height          =   255
         Left            =   3940
         TabIndex        =   11
         Top             =   1905
         Width           =   255
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
         Left            =   3940
         TabIndex        =   13
         Top             =   2190
         Width           =   255
      End
      Begin VB.CheckBox ChkINS 
         Caption         =   "Check3"
         Height          =   255
         Left            =   3940
         TabIndex        =   5
         Top             =   1120
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
         Left            =   3940
         TabIndex        =   7
         Top             =   1380
         Width           =   255
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
         Left            =   8280
         TabIndex        =   81
         Top             =   4680
         Width           =   255
      End
      Begin VB.CheckBox ChkRepCat 
         Caption         =   "Check3"
         Height          =   255
         Left            =   8280
         TabIndex        =   59
         Top             =   2760
         Width           =   255
      End
      Begin VB.CheckBox ChkRep 
         Caption         =   "Check3"
         Height          =   255
         Left            =   8280
         TabIndex        =   62
         Top             =   3030
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
         Left            =   8280
         TabIndex        =   66
         Top             =   3280
         Width           =   255
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
         Left            =   8280
         TabIndex        =   69
         Top             =   3550
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
         Height          =   280
         Left            =   8280
         TabIndex        =   72
         Top             =   3820
         Width           =   255
      End
      Begin VB.CheckBox ChkHOSGrp 
         Caption         =   "Check3"
         Height          =   255
         Left            =   8280
         TabIndex        =   48
         Top             =   1620
         Width           =   255
      End
      Begin VB.CheckBox ChkHOS 
         Caption         =   "Check3"
         Height          =   255
         Left            =   8280
         TabIndex        =   46
         Top             =   1335
         Width           =   255
      End
      Begin VB.CheckBox ChkPolNo 
         Caption         =   "Check3"
         Height          =   255
         Left            =   3940
         TabIndex        =   15
         Top             =   2460
         Width           =   255
      End
      Begin VB.CheckBox ChkGrpCom 
         Caption         =   "Check3"
         Height          =   255
         Left            =   8280
         TabIndex        =   44
         Top             =   1080
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
         Left            =   8280
         TabIndex        =   78
         Top             =   4390
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
         Left            =   8280
         TabIndex        =   57
         Top             =   2460
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
         Left            =   8280
         TabIndex        =   75
         Top             =   4110
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
         Left            =   8280
         TabIndex        =   54
         Top             =   2160
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
         Left            =   8280
         TabIndex        =   51
         Top             =   1880
         Width           =   255
      End
      Begin VB.CheckBox ChkHLT 
         Caption         =   "Check3"
         Height          =   255
         Left            =   3940
         TabIndex        =   3
         Top             =   850
         Width           =   255
      End
      Begin MSComDlg.CommonDialog CommonDialog1 
         Left            =   7560
         Top             =   240
         _ExtentX        =   847
         _ExtentY        =   847
         _Version        =   393216
      End
      Begin Crystal.CrystalReport crptSummary 
         Left            =   8040
         Top             =   240
         _ExtentX        =   741
         _ExtentY        =   741
         _Version        =   348160
         PrintFileLinesPerPage=   60
      End
      Begin VB.ComboBox cboHLTCode 
         Height          =   315
         Left            =   1200
         Sorted          =   -1  'True
         TabIndex        =   2
         Top             =   840
         Width           =   2670
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
         TabIndex        =   4
         Top             =   1100
         Width           =   2670
      End
      Begin VB.ComboBox CboPayor 
         Height          =   315
         Left            =   5520
         TabIndex        =   41
         Top             =   795
         Width           =   2670
      End
      Begin VB.ComboBox CboGrpCom 
         Height          =   315
         Left            =   5520
         TabIndex        =   43
         Top             =   1080
         Width           =   2670
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
         Left            =   5520
         Sorted          =   -1  'True
         TabIndex        =   45
         Top             =   1320
         Width           =   2670
      End
      Begin VB.ComboBox cboHOSGrp 
         Height          =   315
         Left            =   5520
         TabIndex        =   47
         Top             =   1605
         Width           =   2670
      End
      Begin MSMask.MaskEdBox TxtDOA2 
         Height          =   315
         Left            =   7070
         TabIndex        =   50
         Top             =   1890
         Width           =   1125
         _ExtentX        =   1984
         _ExtentY        =   556
         _Version        =   393216
         MaxLength       =   10
         Mask            =   "##/##/####"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox txtDOD2 
         Height          =   315
         Left            =   7070
         TabIndex        =   53
         Top             =   2170
         Width           =   1125
         _ExtentX        =   1984
         _ExtentY        =   556
         _Version        =   393216
         MaxLength       =   10
         Mask            =   "##/##/####"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox TxtDOR2 
         Height          =   315
         Left            =   7070
         TabIndex        =   56
         Top             =   2460
         Width           =   1125
         _ExtentX        =   1984
         _ExtentY        =   556
         _Version        =   393216
         MaxLength       =   10
         Mask            =   "##/##/####"
         PromptChar      =   "_"
      End
      Begin VB.ComboBox CboClaimability 
         Height          =   315
         Left            =   1200
         Sorted          =   -1  'True
         TabIndex        =   6
         Top             =   1380
         Width           =   2670
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
         TabIndex        =   8
         Top             =   1650
         Width           =   2670
      End
      Begin VB.ComboBox CboWsStatus 
         Height          =   315
         Left            =   1200
         TabIndex        =   10
         Top             =   1890
         Width           =   2670
      End
      Begin VB.ComboBox CboState 
         Height          =   315
         Left            =   1200
         TabIndex        =   12
         Top             =   2180
         Width           =   2670
      End
      Begin VB.TextBox txtPolNo 
         Height          =   285
         Left            =   1200
         MaxLength       =   25
         TabIndex        =   14
         Top             =   2450
         Width           =   2670
      End
      Begin VB.TextBox TxtMBMNo 
         Height          =   285
         Left            =   1200
         MaxLength       =   15
         TabIndex        =   16
         Top             =   2700
         Width           =   2670
      End
      Begin VB.TextBox TxtPatientName 
         Height          =   285
         Left            =   1200
         MaxLength       =   60
         TabIndex        =   18
         Top             =   2960
         Width           =   2670
      End
      Begin VB.TextBox txtDrNoTo 
         Height          =   285
         Left            =   2740
         MaxLength       =   20
         TabIndex        =   21
         Top             =   3200
         Width           =   1120
      End
      Begin VB.TextBox txtCRNoTo 
         Height          =   285
         Left            =   2740
         MaxLength       =   20
         TabIndex        =   24
         Top             =   3450
         Width           =   1120
      End
      Begin VB.TextBox txtDRNoFr 
         Height          =   285
         Left            =   1200
         MaxLength       =   20
         TabIndex        =   20
         Top             =   3200
         Width           =   1120
      End
      Begin VB.TextBox txtCRNoFr 
         Height          =   285
         Left            =   1200
         MaxLength       =   20
         TabIndex        =   23
         Top             =   3450
         Width           =   1120
      End
      Begin MSMask.MaskEdBox txtDOA1 
         Height          =   315
         Left            =   5520
         TabIndex        =   49
         Top             =   1890
         Width           =   1125
         _ExtentX        =   1984
         _ExtentY        =   556
         _Version        =   393216
         MaxLength       =   10
         Mask            =   "##/##/####"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox txtDOD1 
         Height          =   315
         Left            =   5520
         TabIndex        =   52
         Top             =   2180
         Width           =   1125
         _ExtentX        =   1984
         _ExtentY        =   556
         _Version        =   393216
         MaxLength       =   10
         Mask            =   "##/##/####"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox TxtDOR1 
         Height          =   315
         Left            =   5520
         TabIndex        =   55
         Top             =   2460
         Width           =   1125
         _ExtentX        =   1984
         _ExtentY        =   556
         _Version        =   393216
         MaxLength       =   10
         Mask            =   "##/##/####"
         PromptChar      =   "_"
      End
      Begin VB.ComboBox cboRepCat1 
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
         Left            =   5520
         Sorted          =   -1  'True
         TabIndex        =   58
         Top             =   2740
         Width           =   1125
      End
      Begin VB.ComboBox cboRepCode1 
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
         Left            =   5520
         Sorted          =   -1  'True
         TabIndex        =   61
         Top             =   3000
         Width           =   1125
      End
      Begin VB.TextBox TxtPlan1 
         Height          =   315
         Left            =   5520
         MaxLength       =   4
         TabIndex        =   64
         Top             =   3250
         Width           =   1120
      End
      Begin VB.TextBox TxtBordxNo1 
         Height          =   315
         Left            =   5520
         MaxLength       =   10
         TabIndex        =   67
         Top             =   3540
         Width           =   1120
      End
      Begin MSMask.MaskEdBox TxtPayEffDate1 
         Height          =   315
         Left            =   5520
         TabIndex        =   70
         Top             =   3820
         Width           =   1125
         _ExtentX        =   1984
         _ExtentY        =   556
         _Version        =   393216
         MaxLength       =   10
         Mask            =   "##/##/####"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox txtPayExpDate1 
         Height          =   315
         Left            =   5520
         TabIndex        =   73
         Top             =   4110
         Width           =   1125
         _ExtentX        =   1984
         _ExtentY        =   556
         _Version        =   393216
         MaxLength       =   10
         Mask            =   "##/##/####"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox TxtDateJoin1 
         Height          =   315
         Left            =   5520
         TabIndex        =   76
         Top             =   4400
         Width           =   1125
         _ExtentX        =   1984
         _ExtentY        =   556
         _Version        =   393216
         MaxLength       =   10
         Mask            =   "##/##/####"
         PromptChar      =   "_"
      End
      Begin VB.TextBox txtPOLYear1 
         Height          =   315
         Left            =   5520
         MaxLength       =   2
         TabIndex        =   79
         Top             =   4680
         Width           =   1120
      End
      Begin VB.ComboBox cboRepCat2 
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
         Left            =   7070
         Sorted          =   -1  'True
         TabIndex        =   60
         Top             =   2740
         Width           =   1125
      End
      Begin VB.ComboBox cboRepCode2 
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
         Left            =   7070
         Sorted          =   -1  'True
         TabIndex        =   63
         Top             =   3015
         Width           =   1125
      End
      Begin VB.TextBox TxtPlan2 
         Height          =   315
         Left            =   7070
         MaxLength       =   4
         TabIndex        =   65
         Top             =   3260
         Width           =   1120
      End
      Begin VB.TextBox TxtBordxNo2 
         Height          =   315
         Left            =   7070
         MaxLength       =   10
         TabIndex        =   68
         Top             =   3540
         Width           =   1120
      End
      Begin MSMask.MaskEdBox TxtPayEffDate2 
         Height          =   315
         Left            =   7070
         TabIndex        =   71
         Top             =   3820
         Width           =   1125
         _ExtentX        =   1984
         _ExtentY        =   556
         _Version        =   393216
         MaxLength       =   10
         Mask            =   "##/##/####"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox txtPayExpDate2 
         Height          =   315
         Left            =   7070
         TabIndex        =   74
         Top             =   4100
         Width           =   1125
         _ExtentX        =   1984
         _ExtentY        =   556
         _Version        =   393216
         MaxLength       =   10
         Mask            =   "##/##/####"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox txtDateJoin2 
         Height          =   315
         Left            =   7070
         TabIndex        =   77
         Top             =   4370
         Width           =   1125
         _ExtentX        =   1984
         _ExtentY        =   556
         _Version        =   393216
         MaxLength       =   10
         Mask            =   "##/##/####"
         PromptChar      =   "_"
      End
      Begin VB.TextBox txtPOLYear2 
         Height          =   315
         Left            =   7070
         MaxLength       =   2
         TabIndex        =   80
         Top             =   4650
         Width           =   1120
      End
      Begin VB.TextBox txtAmtFr 
         Height          =   285
         Left            =   1200
         MaxLength       =   8
         TabIndex        =   26
         Top             =   3710
         Width           =   1120
      End
      Begin VB.TextBox txtAmtTo 
         Height          =   285
         Left            =   2745
         MaxLength       =   8
         TabIndex        =   27
         Top             =   3710
         Width           =   1120
      End
      Begin VB.TextBox txtPVNo1 
         Height          =   285
         Left            =   1200
         MaxLength       =   20
         TabIndex        =   29
         Top             =   3960
         Width           =   1120
      End
      Begin VB.TextBox txtPVNo2 
         Height          =   285
         Left            =   2740
         MaxLength       =   20
         TabIndex        =   30
         Top             =   3960
         Width           =   1120
      End
      Begin MSMask.MaskEdBox TxtPVDate1 
         Height          =   285
         Left            =   1200
         TabIndex        =   32
         Top             =   4200
         Width           =   1125
         _ExtentX        =   1984
         _ExtentY        =   503
         _Version        =   393216
         MaxLength       =   10
         Mask            =   "##/##/####"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox TxtPVDate2 
         Height          =   285
         Left            =   2760
         TabIndex        =   33
         Top             =   4200
         Width           =   1125
         _ExtentX        =   1984
         _ExtentY        =   503
         _Version        =   393216
         MaxLength       =   10
         Mask            =   "##/##/####"
         PromptChar      =   "_"
      End
      Begin VB.TextBox TxtRecNo1 
         Height          =   285
         Left            =   1200
         MaxLength       =   15
         TabIndex        =   35
         Top             =   4440
         Width           =   1120
      End
      Begin VB.TextBox TxtRecNo2 
         Height          =   285
         Left            =   2745
         MaxLength       =   15
         TabIndex        =   36
         Top             =   4440
         Width           =   1120
      End
      Begin MSMask.MaskEdBox txtRecDate1 
         Height          =   285
         Left            =   1200
         TabIndex        =   38
         Top             =   4680
         Width           =   1125
         _ExtentX        =   1984
         _ExtentY        =   503
         _Version        =   393216
         MaxLength       =   10
         Mask            =   "##/##/####"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox txtRecDate2 
         Height          =   285
         Left            =   2760
         TabIndex        =   39
         Top             =   4680
         Width           =   1125
         _ExtentX        =   1984
         _ExtentY        =   503
         _Version        =   393216
         MaxLength       =   10
         Mask            =   "##/##/####"
         PromptChar      =   "_"
      End
      Begin VB.Label Label58 
         Caption         =   "*"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000C0&
         Height          =   135
         Left            =   8320
         TabIndex        =   154
         Top             =   720
         Width           =   135
      End
      Begin VB.Label Label52 
         Caption         =   "*"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000C0&
         Height          =   135
         Left            =   3990
         TabIndex        =   153
         Top             =   750
         Width           =   135
      End
      Begin VB.Label Label55 
         Caption         =   "* Search Empty Data"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000C0&
         Height          =   255
         Left            =   120
         TabIndex        =   152
         Top             =   5040
         Width           =   1935
      End
      Begin VB.Label Label62 
         Caption         =   "Patient Name"
         Height          =   255
         Left            =   120
         TabIndex        =   151
         Top             =   3010
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
         TabIndex        =   150
         Top             =   1680
         Width           =   945
      End
      Begin VB.Label Label61 
         Caption         =   "To"
         Height          =   255
         Left            =   6720
         TabIndex        =   149
         Top             =   3600
         Width           =   255
      End
      Begin VB.Label Label60 
         Caption         =   "Bordx No"
         Height          =   255
         Left            =   4440
         TabIndex        =   148
         Top             =   3700
         Width           =   855
      End
      Begin VB.Label Label57 
         Caption         =   "To"
         Height          =   255
         Left            =   6720
         TabIndex        =   147
         Top             =   4720
         Width           =   255
      End
      Begin VB.Label Label56 
         Caption         =   "Pol. Year"
         Height          =   255
         Left            =   4440
         TabIndex        =   146
         Top             =   4750
         Width           =   1095
      End
      Begin VB.Label Label44 
         Caption         =   "To"
         Height          =   255
         Left            =   6720
         TabIndex        =   145
         Top             =   4440
         Width           =   255
      End
      Begin VB.Label Label43 
         Caption         =   "Date Join"
         Height          =   255
         Left            =   4440
         TabIndex        =   144
         ToolTipText     =   "Payor Expiry Date"
         Top             =   4480
         Width           =   1215
      End
      Begin VB.Label Label42 
         Caption         =   "Hosp. Grp"
         Height          =   255
         Left            =   4440
         TabIndex        =   143
         Top             =   1680
         Width           =   1095
      End
      Begin VB.Label Label41 
         Caption         =   "Hospital"
         Height          =   255
         Left            =   4440
         TabIndex        =   142
         Top             =   1440
         Width           =   975
      End
      Begin VB.Label Label40 
         Caption         =   "To"
         Height          =   255
         Left            =   6720
         TabIndex        =   141
         Top             =   3075
         Width           =   255
      End
      Begin VB.Label Label39 
         Caption         =   "To"
         Height          =   255
         Left            =   6720
         TabIndex        =   140
         Top             =   2805
         Width           =   255
      End
      Begin VB.Label Label38 
         Caption         =   "Repud Code"
         Height          =   255
         Left            =   4440
         TabIndex        =   139
         Top             =   3120
         Width           =   975
      End
      Begin VB.Label Label37 
         Caption         =   "Repud Cat"
         Height          =   255
         Left            =   4440
         TabIndex        =   138
         Top             =   2840
         Width           =   975
      End
      Begin VB.Label Label36 
         Caption         =   "Grp Company"
         Height          =   255
         Left            =   4440
         TabIndex        =   137
         Top             =   1140
         Width           =   1095
      End
      Begin VB.Label Label53 
         Caption         =   "Claimability"
         Height          =   255
         Left            =   120
         TabIndex        =   136
         Top             =   1410
         Width           =   975
      End
      Begin VB.Label Label51 
         Caption         =   "Plan Code"
         Height          =   255
         Left            =   4440
         TabIndex        =   135
         Top             =   3420
         Width           =   855
      End
      Begin VB.Label Label46 
         Caption         =   "To"
         Height          =   255
         Left            =   6720
         TabIndex        =   134
         Top             =   3340
         Width           =   255
      End
      Begin VB.Label Label35 
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
         TabIndex        =   133
         Top             =   1140
         Width           =   690
      End
      Begin VB.Label Label34 
         Caption         =   "Eff Date"
         Height          =   255
         Left            =   4440
         TabIndex        =   132
         ToolTipText     =   "Payor Effective Date"
         Top             =   3980
         Width           =   1215
      End
      Begin VB.Label Label33 
         Caption         =   "Exp Date"
         Height          =   255
         Left            =   4440
         TabIndex        =   131
         ToolTipText     =   "Payor Expiry Date"
         Top             =   4230
         Width           =   975
      End
      Begin VB.Label Label32 
         Caption         =   "DOR"
         Height          =   255
         Left            =   4440
         TabIndex        =   130
         ToolTipText     =   "Member Value Date"
         Top             =   2560
         Width           =   855
      End
      Begin VB.Label Label31 
         Caption         =   "To"
         Height          =   255
         Left            =   6720
         TabIndex        =   129
         Top             =   3860
         Width           =   255
      End
      Begin VB.Label Label26 
         Caption         =   "To"
         Height          =   255
         Left            =   6720
         TabIndex        =   128
         Top             =   4160
         Width           =   255
      End
      Begin VB.Label Label24 
         Caption         =   "To"
         Height          =   255
         Left            =   6720
         TabIndex        =   127
         Top             =   2520
         Width           =   255
      End
      Begin VB.Label Label47 
         Caption         =   "DOA"
         Height          =   255
         Left            =   4440
         TabIndex        =   126
         ToolTipText     =   "Date of Admission"
         Top             =   1960
         Width           =   735
      End
      Begin VB.Label Label48 
         Caption         =   "DOD"
         Height          =   255
         Left            =   4440
         TabIndex        =   125
         ToolTipText     =   "Date of Discharged"
         Top             =   2280
         Width           =   615
      End
      Begin VB.Label Label49 
         Caption         =   "To"
         Height          =   255
         Left            =   6720
         TabIndex        =   124
         Top             =   1960
         Width           =   255
      End
      Begin VB.Label Label50 
         Caption         =   "To"
         Height          =   255
         Left            =   6720
         TabIndex        =   123
         Top             =   2235
         Width           =   255
      End
      Begin VB.Label Label18 
         Caption         =   "WS Status"
         Height          =   255
         Left            =   120
         TabIndex        =   122
         Top             =   1950
         Width           =   975
      End
      Begin VB.Label Label12 
         Caption         =   "State"
         Height          =   255
         Left            =   120
         TabIndex        =   121
         Top             =   2220
         Width           =   855
      End
      Begin VB.Label Label30 
         Caption         =   "Policy No."
         Height          =   255
         Left            =   120
         TabIndex        =   120
         Top             =   2480
         Width           =   975
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
         TabIndex        =   119
         Top             =   850
         Width           =   705
      End
      Begin VB.Label Label15 
         Caption         =   "To"
         Height          =   255
         Left            =   2400
         TabIndex        =   118
         Top             =   3280
         Width           =   255
      End
      Begin VB.Label Label20 
         Caption         =   "DR No "
         Height          =   255
         Left            =   120
         TabIndex        =   117
         Top             =   3280
         Width           =   1095
      End
      Begin VB.Label Label13 
         Caption         =   "To"
         Height          =   255
         Left            =   2415
         TabIndex        =   116
         Top             =   4320
         Width           =   255
      End
      Begin VB.Label Label17 
         Caption         =   "PV Date "
         Height          =   255
         Left            =   120
         TabIndex        =   115
         Top             =   4320
         Width           =   1095
      End
      Begin VB.Label Label16 
         Caption         =   "PV No "
         Height          =   255
         Left            =   120
         TabIndex        =   114
         Top             =   4050
         Width           =   1215
      End
      Begin VB.Label Label11 
         Caption         =   "To"
         Height          =   255
         Left            =   2415
         TabIndex        =   113
         Top             =   4050
         Width           =   255
      End
      Begin VB.Label Label6 
         Caption         =   "To"
         Height          =   255
         Left            =   2415
         TabIndex        =   112
         Top             =   3530
         Width           =   255
      End
      Begin VB.Label Label19 
         Caption         =   "CR No "
         Height          =   255
         Left            =   120
         TabIndex        =   111
         Top             =   3530
         Width           =   1095
      End
      Begin VB.Label txtLocalFileLocation 
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   1680
         TabIndex        =   0
         Top             =   480
         Width           =   1665
      End
      Begin VB.Label txtFilName 
         BackColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   4200
         TabIndex        =   1
         Top             =   480
         Width           =   1305
      End
      Begin VB.Label Label25 
         Caption         =   "File Name"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   3480
         TabIndex        =   109
         Top             =   480
         Width           =   765
      End
      Begin VB.Label Label5 
         Caption         =   "File on PC"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   240
         TabIndex        =   108
         Top             =   480
         Width           =   765
      End
      Begin VB.Label Label27 
         Caption         =   "Member No"
         Height          =   255
         Left            =   120
         TabIndex        =   107
         Top             =   2740
         Width           =   1215
      End
      Begin VB.Label Label14 
         Caption         =   "To"
         Height          =   255
         Left            =   2400
         TabIndex        =   105
         Top             =   4800
         Width           =   255
      End
      Begin VB.Label Label8 
         Caption         =   "Receipt Date"
         Height          =   255
         Left            =   120
         TabIndex        =   104
         Top             =   4800
         Width           =   1095
      End
      Begin VB.Label Label7 
         Caption         =   "Receipt No"
         Height          =   255
         Left            =   120
         TabIndex        =   102
         Top             =   4560
         Width           =   1095
      End
      Begin VB.Label Label1 
         Caption         =   "To"
         Height          =   255
         Left            =   2415
         TabIndex        =   101
         Top             =   3800
         Width           =   255
      End
      Begin VB.Label Label3 
         Caption         =   "DR/CR Amt"
         Height          =   255
         Left            =   120
         TabIndex        =   100
         Top             =   3800
         Width           =   855
      End
      Begin VB.Label Label2 
         Caption         =   "Payor"
         Height          =   255
         Left            =   4440
         TabIndex        =   99
         Top             =   840
         Width           =   1215
      End
      Begin VB.Label Label4 
         Caption         =   "DN / CN"
         Height          =   255
         Left            =   240
         TabIndex        =   98
         Top             =   240
         Width           =   1215
      End
      Begin VB.Label Label9 
         Caption         =   "To"
         Height          =   255
         Left            =   2415
         TabIndex        =   103
         Top             =   4560
         Width           =   240
      End
   End
   Begin VB.Frame Frame2 
      Height          =   615
      Left            =   120
      TabIndex        =   90
      Top             =   7320
      Width           =   11655
      Begin VB.CommandButton CmdView 
         Caption         =   "&View"
         Height          =   375
         Left            =   8400
         TabIndex        =   85
         Top             =   180
         Width           =   615
      End
      Begin VB.CommandButton CmdBack 
         Caption         =   "&Back"
         Height          =   375
         Left            =   7800
         TabIndex        =   84
         Top             =   180
         Width           =   615
      End
      Begin MSMask.MaskEdBox txtDate 
         Height          =   255
         Left            =   5280
         TabIndex        =   110
         Top             =   480
         Visible         =   0   'False
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   450
         _Version        =   393216
         PromptChar      =   "_"
      End
      Begin VB.CommandButton CmdExit 
         Cancel          =   -1  'True
         Caption         =   "E&xit"
         Height          =   375
         Left            =   10920
         TabIndex        =   89
         Top             =   180
         Width           =   615
      End
      Begin VB.CommandButton CmdClear 
         Caption         =   "&Clear"
         Height          =   375
         Left            =   10320
         TabIndex        =   88
         Top             =   180
         Width           =   615
      End
      Begin VB.CommandButton cmdFile 
         Caption         =   "&Export"
         Height          =   375
         Left            =   9600
         TabIndex        =   87
         Top             =   180
         Width           =   735
      End
      Begin VB.CommandButton CmdPrint 
         Caption         =   "&Print"
         Height          =   375
         Left            =   9000
         TabIndex        =   86
         Top             =   180
         Width           =   615
      End
      Begin VB.CommandButton cmdStop 
         Caption         =   "S&top"
         Height          =   375
         Left            =   7200
         TabIndex        =   83
         Top             =   180
         Width           =   615
      End
      Begin VB.CommandButton CmdSearch 
         Caption         =   "&Search"
         Default         =   -1  'True
         Height          =   375
         Left            =   6480
         TabIndex        =   82
         Top             =   180
         Width           =   735
      End
      Begin VB.Label Label29 
         Caption         =   "Total Amt"
         Height          =   255
         Left            =   2880
         TabIndex        =   96
         Top             =   240
         Width           =   855
      End
      Begin VB.Label lblTotal 
         Alignment       =   1  'Right Justify
         BackColor       =   &H80000009&
         Height          =   255
         Left            =   1080
         TabIndex        =   95
         Top             =   600
         Visible         =   0   'False
         Width           =   735
      End
      Begin VB.Label Label22 
         Caption         =   "of"
         Height          =   255
         Left            =   840
         TabIndex        =   94
         Top             =   600
         Visible         =   0   'False
         Width           =   375
      End
      Begin VB.Label lblCurrent 
         Alignment       =   1  'Right Justify
         BackColor       =   &H80000009&
         Height          =   255
         Left            =   120
         TabIndex        =   93
         Top             =   240
         Width           =   1575
      End
      Begin VB.Label Label21 
         Caption         =   "Records"
         Height          =   255
         Left            =   1800
         TabIndex        =   92
         Top             =   240
         Width           =   735
      End
      Begin VB.Label LblTotalAmt 
         Alignment       =   1  'Right Justify
         BackColor       =   &H80000009&
         Height          =   255
         Left            =   3720
         TabIndex        =   91
         Top             =   240
         Width           =   1455
      End
   End
   Begin MSComctlLib.ListView lstDebitCreditList 
      Height          =   1695
      Left            =   120
      TabIndex        =   106
      Top             =   5640
      Width           =   11655
      _ExtentX        =   20558
      _ExtentY        =   2990
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
      NumItems        =   8
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Ref No"
         Object.Width           =   2117
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   1
         Text            =   "Date"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "Member Name"
         Object.Width           =   5292
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Text            =   "Patient Name"
         Object.Width           =   5292
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   4
         Text            =   "Amount"
         Object.Width           =   2117
      EndProperty
      BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   5
         Text            =   "Hospital"
         Object.Width           =   2646
      EndProperty
      BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   6
         Text            =   "Payor"
         Object.Width           =   1764
      EndProperty
      BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   7
         Text            =   "Claim No"
         Object.Width           =   2540
      EndProperty
   End
   Begin VB.Label Label28 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00404040&
      Caption         =   "Debit/Credit Note Tracking"
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
      TabIndex        =   157
      Top             =   0
      Width           =   11805
   End
End
Attribute VB_Name = "FrmDrCrTracking"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public DNCNCriteria As String
Public intExit As Integer
Dim DataString As String
Public Hospitals As String
Private m_blnDirection(1 To 8) As Boolean
Dim path As String

Private Sub CmdBack_Click()
    Frame1.Visible = True
    lstDebitCreditList.Top = 6000
    lstDebitCreditList.Width = 11655
    lstDebitCreditList.Left = 120
    lstDebitCreditList.Height = 1335
End Sub

Private Sub CmdView_Click()
    Frame1.Visible = False
    lstDebitCreditList.Top = 480
    lstDebitCreditList.Width = 11655
    lstDebitCreditList.Left = 120
    lstDebitCreditList.Height = 6855
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

Private Sub cboHLTCode_KeyPress(KeyAscii As Integer)
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


Private Sub cboHOSGrp_KeyPress(KeyAscii As Integer)
Dim NewText As String
Dim ValidCount As Integer
Dim ValidValue As String
Dim i As Integer
    If KeyAscii >= 32 And KeyAscii <> 127 Then
    NewText = LCase(Left(cboHOSGrp.Text, cboHOSGrp.SelStart) + Chr(KeyAscii) + Mid(cboHOSGrp.Text, cboHOSGrp.SelStart + cboHOSGrp.SelLength + 1))
    
    ValidCount = 0
    ValidValue = ""
    
    For i = 0 To cboHOSGrp.ListCount - 1
        
        If NewText = LCase(Left(cboHOSGrp.List(i), Len(NewText))) Then
                ValidCount = ValidCount + 1
                ValidValue = cboHOSGrp.List(i)
        End If
    Next
    If ValidCount <= 1 Then KeyAscii = 0
        If ValidCount = 1 Then
            cboHOSGrp.Text = ValidValue
            cboHOSGrp.SelStart = 0
            cboHOSGrp.SelLength = Len(ValidValue)
        End If
    End If

End Sub

Private Sub cboHOSName_KeyPress(KeyAscii As Integer)
Dim NewText As String
Dim ValidCount As Integer
Dim ValidValue As String
Dim i As Integer
    If KeyAscii >= 32 And KeyAscii <> 127 Then
    NewText = LCase(Left(cboHOSName.Text, cboHOSName.SelStart) + Chr(KeyAscii) + Mid(cboHOSName.Text, cboHOSName.SelStart + cboHOSName.SelLength + 1))
    
    ValidCount = 0
    ValidValue = ""
    
    For i = 0 To cboHOSName.ListCount - 1
        
        If NewText = LCase(Left(cboHOSName.List(i), Len(NewText))) Then
                ValidCount = ValidCount + 1
                ValidValue = cboHOSName.List(i)
        End If
    Next
    If ValidCount <= 1 Then KeyAscii = 0
        If ValidCount = 1 Then
            cboHOSName.Text = ValidValue
            cboHOSName.SelStart = 0
            cboHOSName.SelLength = Len(ValidValue)
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

Private Sub CboPayor_KeyPress(KeyAscii As Integer)
Dim NewText As String
Dim ValidCount As Integer
Dim ValidValue As String
Dim i As Integer
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
End Sub


Private Sub cboRepCat1_KeyPress(KeyAscii As Integer)
Dim NewText As String
Dim ValidCount As Integer
Dim ValidValue As String
Dim i As Integer
    If KeyAscii >= 32 And KeyAscii <> 127 Then
    NewText = LCase(Left(cboRepCat1.Text, cboRepCat1.SelStart) + Chr(KeyAscii) + Mid(cboRepCat1.Text, cboRepCat1.SelStart + cboRepCat1.SelLength + 1))
    
    ValidCount = 0
    ValidValue = ""
    
    For i = 0 To cboRepCat1.ListCount - 1
        
        If NewText = LCase(Left(cboRepCat1.List(i), Len(NewText))) Then
                ValidCount = ValidCount + 1
                ValidValue = cboRepCat1.List(i)
        End If
    Next
    If ValidCount <= 1 Then KeyAscii = 0
        If ValidCount = 1 Then
            cboRepCat1.Text = ValidValue
            cboRepCat1.SelStart = 0
            cboRepCat1.SelLength = Len(ValidValue)
        End If
    End If

End Sub


Private Sub cboRepCat2_KeyPress(KeyAscii As Integer)
Dim NewText As String
Dim ValidCount As Integer
Dim ValidValue As String
Dim i As Integer
    If KeyAscii >= 32 And KeyAscii <> 127 Then
    NewText = LCase(Left(cboRepCat2.Text, cboRepCat2.SelStart) + Chr(KeyAscii) + Mid(cboRepCat2.Text, cboRepCat2.SelStart + cboRepCat2.SelLength + 1))
    
    ValidCount = 0
    ValidValue = ""
    
    For i = 0 To cboRepCat2.ListCount - 1
        
        If NewText = LCase(Left(cboRepCat2.List(i), Len(NewText))) Then
                ValidCount = ValidCount + 1
                ValidValue = cboRepCat2.List(i)
        End If
    Next
    If ValidCount <= 1 Then KeyAscii = 0
        If ValidCount = 1 Then
            cboRepCat2.Text = ValidValue
            cboRepCat2.SelStart = 0
            cboRepCat2.SelLength = Len(ValidValue)
        End If
    End If
End Sub


Private Sub cboRepCode1_KeyPress(KeyAscii As Integer)
Dim NewText As String
Dim ValidCount As Integer
Dim ValidValue As String
Dim i As Integer
    If KeyAscii >= 32 And KeyAscii <> 127 Then
    NewText = LCase(Left(cboRepCode1.Text, cboRepCode1.SelStart) + Chr(KeyAscii) + Mid(cboRepCode1.Text, cboRepCode1.SelStart + cboRepCode1.SelLength + 1))
    
    ValidCount = 0
    ValidValue = ""
    
    For i = 0 To cboRepCode1.ListCount - 1
        
        If NewText = LCase(Left(cboRepCode1.List(i), Len(NewText))) Then
                ValidCount = ValidCount + 1
                ValidValue = cboRepCode1.List(i)
        End If
    Next
    If ValidCount <= 1 Then KeyAscii = 0
        If ValidCount = 1 Then
            cboRepCode1.Text = ValidValue
            cboRepCode1.SelStart = 0
            cboRepCode1.SelLength = Len(ValidValue)
        End If
    End If
End Sub


Private Sub cboRepCode2_KeyPress(KeyAscii As Integer)
Dim NewText As String
Dim ValidCount As Integer
Dim ValidValue As String
Dim i As Integer
    If KeyAscii >= 32 And KeyAscii <> 127 Then
    NewText = LCase(Left(cboRepCode2.Text, cboRepCode2.SelStart) + Chr(KeyAscii) + Mid(cboRepCode2.Text, cboRepCode2.SelStart + cboRepCode2.SelLength + 1))
    
    ValidCount = 0
    ValidValue = ""
    
    For i = 0 To cboRepCode2.ListCount - 1
        
        If NewText = LCase(Left(cboRepCode2.List(i), Len(NewText))) Then
                ValidCount = ValidCount + 1
                ValidValue = cboRepCode2.List(i)
        End If
    Next
    If ValidCount <= 1 Then KeyAscii = 0
        If ValidCount = 1 Then
            cboRepCode2.Text = ValidValue
            cboRepCode2.SelStart = 0
            cboRepCode2.SelLength = Len(ValidValue)
        End If
    End If
End Sub

Private Sub CboState_KeyPress(KeyAscii As Integer)
Dim NewText As String
Dim ValidCount As Integer
Dim ValidValue As String
Dim i As Integer
    If KeyAscii >= 32 And KeyAscii <> 127 Then
    NewText = LCase(Left(CboState.Text, CboState.SelStart) + Chr(KeyAscii) + Mid(CboState.Text, CboState.SelStart + CboState.SelLength + 1))
    
    ValidCount = 0
    ValidValue = ""
    
    For i = 0 To CboState.ListCount - 1
        
        If NewText = LCase(Left(CboState.List(i), Len(NewText))) Then
                ValidCount = ValidCount + 1
                ValidValue = CboState.List(i)
        End If
    Next
    If ValidCount <= 1 Then KeyAscii = 0
        If ValidCount = 1 Then
            CboState.Text = ValidValue
            CboState.SelStart = 0
            CboState.SelLength = Len(ValidValue)
        End If
    End If
End Sub

Private Sub CboWsStatus_KeyPress(KeyAscii As Integer)
Dim NewText As String
Dim ValidCount As Integer
Dim ValidValue As String
Dim i As Integer
    If KeyAscii >= 32 And KeyAscii <> 127 Then
    NewText = LCase(Left(CboWsStatus.Text, CboWsStatus.SelStart) + Chr(KeyAscii) + Mid(CboWsStatus.Text, CboWsStatus.SelStart + CboWsStatus.SelLength + 1))
    
    ValidCount = 0
    ValidValue = ""
    
    For i = 0 To CboWsStatus.ListCount - 1
        
        If NewText = LCase(Left(CboWsStatus.List(i), Len(NewText))) Then
                ValidCount = ValidCount + 1
                ValidValue = CboWsStatus.List(i)
        End If
    Next
    If ValidCount <= 1 Then KeyAscii = 0
        If ValidCount = 1 Then
            CboWsStatus.Text = ValidValue
            CboWsStatus.SelStart = 0
            CboWsStatus.SelLength = Len(ValidValue)
        End If
    End If
End Sub

Private Sub Form_Load()
    txtDate = Date
    intExit = 0
    lblCurrent = 0
    lblTotal = 0
    LblTotalAmt = 0
    DNCNCriteria = ""
    path = "C:\" 'SRVPath & "IT"
    txtLocalFileLocation = path
    Call ComboListing
End Sub

'******************* Start Command Button ************************************

Private Sub CmdClear_Click()
    Call ClearForm(Me)
    lstDebitCreditList.listitems.Clear
    lblCurrent = 0
    lblTotal = 0
    LblTotalAmt = 0
    Call ComboListing
End Sub
Public Function ClearForm(frmCall)
'On Error Resume Next
Dim ctl As Control

For Each ctl In frmCall.Controls
    If TypeOf ctl Is TextBox Then
        ctl.Text = ""
    ElseIf TypeOf ctl Is ListBox Then
        ctl.Clear
    ElseIf TypeOf ctl Is ComboBox Then
        ctl.Clear
    ElseIf TypeOf ctl Is CheckBox Then
        ctl.value = False
    ElseIf TypeOf ctl Is MaskEdBox Then
        ctl.mask = ""
        ctl.Text = ""
        ctl.mask = "##/##/####"
    End If
Next ctl
End Function
Private Sub cmdStop_Click()
    intExit = 1
    CmdPrint.Enabled = True
    cmdFile.Enabled = True
    CmdClear.Enabled = True
    CmdExit.Enabled = True
End Sub

Private Sub cmdExit_Click()
'Dim RecordExit
'    RecordExit = MsgBox("Quit this form?", vbYesNo + vbExclamation)
'    If RecordExit = vbYes Then
'        bExit = True
        Unload Me
'    End If
End Sub

Private Sub cmdSearch_Click()
    CmdPrint.Enabled = False
    cmdFile.Enabled = False
    CmdClear.Enabled = False
    CmdExit.Enabled = False
    CmdSearch.Enabled = False
    Screen.MousePointer = vbHourglass
    medixmasterconnWS.Open "File Name=" & BOSPath & "Conn\mediconnectWS.UDL"
    Call SearchRecord
    medixmasterconnWS.Close
    Set medixmasterconnWS = Nothing
    Screen.MousePointer = Default
    CmdSearch.Enabled = True
End Sub

Private Sub CmdPrint_Click()
    
Screen.MousePointer = vbHourglass
    
    If lstDebitCreditList.listitems.count <> 0 Then
        
                        
        If OptDebit.value = True Then
               crptSummary.ReportFileName = crdPath & "\Report\DebitTracking.rpt"
               crptSummary.Destination = crptToWindow
               crptSummary.WindowState = crptMaximized
               crptSummary.SelectionFormula = "{VIEW_Debit_Tracking.PAYCODE} = '" & Trim(Mid(CboPayor, 1, 2)) & "'"
               crptSummary.ParameterFields(0) = "txtPayor;" & CboPayor & ";true"
               crptSummary.Action = 1
               
            
        ElseIf OptCredit.value = True Then
               crptSummary.ReportFileName = crdPath & "\Report\CreditTracking.rpt"
               crptSummary.Destination = crptToWindow
               crptSummary.WindowState = crptMaximized
               crptSummary.SelectionFormula = "{VIEW_Credit_Tracking.PAYCODE} = '" & Trim(Mid(CboPayor, 1, 2)) & "'"
               crptSummary.ParameterFields(0) = "txtPayor;" & CboPayor & ";true"
               crptSummary.Action = 1
        End If
            
    Else
        MsgBox "Report cannot be printed without any record.", vbInformation
    End If
    
Screen.MousePointer = vbDefault

End Sub

Private Sub cmdFile_Click()
    If lstDebitCreditList.listitems.count <> 0 Then
        txtFilName = ""
        If OptDebit.value = True Then
            Call ExportToFile
        Else
            Call ExportToFileC
        End If
    Else
        MsgBox "Report cannot be printed without any record.", vbInformation
    End If
End Sub
'******************* End Command Button ************************************


'******************* Start SearchRecord ************************************
Private Function SearchRecord()
Dim strsql As String
Dim strAppend As String
Dim sql As String

    lblCurrent = 0
    lblTotal = 0
    LblTotalAmt = 0
    
    strAppend = ""
    Hospitals = ""
        
    
    If Len(Mid(cboHLTCode, 1, 1)) Then
        strAppend = strAppend + " AND HLTCode ='" & Trim(Mid(cboHLTCode, 1, 1)) & "'"
    End If
    
    If ChkHLT.value = 1 Then
        strAppend = strAppend & " AND (HLTCode Is null OR HLTCode = '')"
    End If
    
    If Len(Mid(CboInsType, 1, 1)) Then
        strAppend = strAppend & " AND INSCode = '" & Trim(Mid(CboInsType, 1, 1)) & "'"
    End If
    
    If ChkINS.value = 1 Then
        strAppend = strAppend & " AND (INSCode Is null OR INSCode = '')"
    End If
    
    If Len(Mid(CboClaimability, 1, 1)) Then
        strAppend = strAppend & " AND CASClaimability = '" & Trim(Mid(CboClaimability, 1, 1)) & "'"
    End If
    
    If ChkClaimability.value = 1 Then
        strAppend = strAppend & " AND (CASClaimability Is null OR CASClaimability = '')"
    End If
    
    If Len(Mid(CboCaseStatus, 1, 1)) Then
        strAppend = strAppend & " AND CASStatus = '" & Trim(Mid(CboCaseStatus, 1, 1)) & "'"
    End If
    
    If ChkCaseStatus.value = 1 Then
        strAppend = strAppend & " AND (CASStatus Is null OR CASStatus = '')"
    End If
    
    If Len(Mid(CboWsStatus, 1, 1)) Then
        strAppend = strAppend & " AND ClaimStatus = '" & Trim(Mid(CboWsStatus, 1, 1)) & "'"
    End If
    
    If ChkWksStatus.value = 1 Then
        strAppend = strAppend & " AND (ClaimStatus Is null OR ClaimStatus = '')"
    End If
    
    If Len(Trim(CboState)) Then
          strAppend = strAppend & "AND STACode= '" & Trim(CboState) & "' "
    End If
    
    If ChkState.value = 1 Then
        strAppend = strAppend & " AND (STACode Is null OR STACode = '')"
    End If
    
    If Len(Trim(txtPolNo)) Then
       strAppend = strAppend + " And  MBMPolicyNo like '%" & ChkStr(txtPolNo) & "%' "
    End If
    
    If ChkPolNo.value = 1 Then
        strAppend = strAppend & " AND (MBMPolicyNo Is null OR MBMPolicyNo = '')"
    End If
    
    If Len(Trim(TxtMBMNo)) Then
        strAppend = strAppend + " And MBMNumber like '%" & ChkStr(TxtMBMNo) & "%'"
    End If
    
    If ChkMemNo.value = 1 Then
        strAppend = strAppend & " AND (MBMNumber IS NULL OR MBMNumber = '')"
    End If
    
    If Len(Trim(TxtPatientName)) Then
        strAppend = strAppend + " And PatientName like '%" & ChkStr(TxtPatientName) & "%'"
    End If
    
    If ChkPatient.value = 1 Then
        strAppend = strAppend & " AND (PatientName IS NULL OR PatientName = '')"
    End If
    
    If OptDebit.value = True Then
        If Len(Trim(txtDRNoFr)) Then
            strAppend = strAppend + " AND DebitCreditRefNo >= '" & RemStr(Trim(txtDRNoFr)) & "'"
        End If

        If Len(Trim(txtDrNoTo)) Then
            strAppend = strAppend + " AND DebitCreditRefNo <= '" & RemStr(Trim(txtDrNoTo)) & "'"
        End If
        
        If IsNumeric(txtAmtFr) Then
           strAppend = strAppend + " AND CLMPayableByMedix2M >= " & Trim(txtAmtFr) & ""
        End If
        
        If IsNumeric(txtAmtTo) Then
           strAppend = strAppend + " AND CLMPayableByMedix2M <= " & Trim(txtAmtTo) & ""
        End If
        
        If ChkAmt.value = 1 Then
            strAppend = strAppend & " AND (CLMPayableByMedix2M Is null OR CLMPayableByMedix2M = '')"
        End If
        
        If ChkDr.value = 1 Then
            strAppend = strAppend & " AND (DebitCreditRefNo Is null OR DebitCreditRefNo = '')"
        End If
    Else
        If Len(Trim(txtCRNoFr)) Then
            strAppend = strAppend + " AND DebitCreditRefNo >= '" & RemStr(Trim(txtCRNoFr)) & "'"
        End If

        If Len(Trim(txtCRNoTo)) Then
            strAppend = strAppend + " AND DebitCreditRefNo <= '" & RemStr(Trim(txtCRNoTo)) & "'"
        End If
        
        If IsNumeric(txtAmtFr) Then
           strAppend = strAppend + " AND CLMPayableByMedix2M >= " & Trim(txtAmtFr) & ""
        End If
        
        If IsNumeric(txtAmtTo) Then
           strAppend = strAppend + " AND CLMPayableByMedix2M <= " & Trim(txtAmtTo) & ""
        End If
        
        If ChkAmt.value = 1 Then
            strAppend = strAppend & " AND (CLMPayableByMedix2M Is null OR CLMPayableByMedix2M = '')"
        End If
            
        If ChkCr.value = 1 Then
            strAppend = strAppend & " AND (DebitCreditRefNo Is null OR DebitCreditRefNo = '')"
        End If
    End If
    
    If IsDate(TxtPVDate1) = True Then
        strAppend = strAppend & " AND PVDate >= '" & Format(TxtPVDate1, "mm/dd/yyyy") & "'"
    End If
    
    If IsDate(TxtPVDate2) = True Then
        strAppend = strAppend & " AND PVDate <= '" & Format(TxtPVDate2, "mm/dd/yyyy") & "'"
    End If
    
    If ChkPVDate.value = 1 Then
        strAppend = strAppend & " AND (PVDate Is null OR PVDate = '')"
    End If
    
    If Len(Trim(txtPVNo1)) Then
        strAppend = strAppend & " AND PVNo >= '" & Trim(txtPVNo1) & "'"
    End If
    
    If Len(Trim(txtPVNo2)) Then
        strAppend = strAppend & " AND PVNo <= '" & Trim(txtPVNo2) & "'"
    End If
    
    If ChkPVNo.value = 1 Then
        strAppend = strAppend & " AND (PVNo Is null OR PVNo = '')"
    End If
    
    If Len(Trim(TxtRecNo1)) Then
        strAppend = strAppend & " AND ReceiptNo >= '" & Trim(TxtRecNo1) & "'"
    End If
    
    If Len(Trim(TxtRecNo2)) Then
        strAppend = strAppend & " AND ReceiptNo <= '" & Trim(TxtRecNo2) & "'"
    End If
    
    If ChkRefNo.value = 1 Then
        strAppend = strAppend & " AND (ReceiptNo Is null OR ReceiptNo = '')"
    End If
    
    If Len(Trim(txtPOLYear2)) Then
        strAppend = strAppend & " AND MBMPolicyYear <= '" & Trim(txtPOLYear2) & "'"
    End If
    
    If ChkHLT.value = 1 Then
        strAppend = strAppend & " AND (HLTCode Is null OR HLTCode = '')"
    End If
    
    
    If IsDate(txtRecDate1) = True Then
        strAppend = strAppend & " AND ReceiptDate >= '" & Format(txtRecDate1, "mm/dd/yyyy") & "'"
    End If
    
    If IsDate(txtRecDate2) = True Then
        strAppend = strAppend & " AND ReceiptDate <= '" & Format(txtRecDate2, "mm/dd/yyyy") & "'"
    End If
    
    If ChkRecDate.value = 1 Then
        strAppend = strAppend & " AND (ReceiptDate Is null OR ReceiptDate = '')"
    End If

    If Len(Mid(CboPayor, 1, 2)) Then
        strAppend = strAppend & " AND PAYCode = '" & Trim(Mid(CboPayor, 1, 2)) & "'"
    End If
           
    If Len(Trim(CboGrpCom)) Then
        strAppend = strAppend & " AND MBMGroupCompany = '" & Trim(CboGrpCom) & "'"
    End If
    
    If Len(Trim(cboHOSName)) Then
        Hospitals = Replace(cboHOSName, "'", "''")
        strAppend = strAppend + " AND HOSName = '" & Trim(Hospitals) & "'"
    End If
    
    If Len(Trim(cboHOSGrp)) Then
        strAppend = strAppend + " And HOSGRPName = '" & Trim(cboHOSGrp) & "'"
    End If
       
    If IsDate(txtDOA1) = True Then
        strAppend = strAppend + " And CasDateAdmitted >= '" & Format(txtDOA1, "mm/dd/yyyy") & "'"
    End If
    
    If IsDate(TxtDOA2) = True Then
        strAppend = strAppend + " And CasDateAdmitted <= '" & Format(TxtDOA2, "mm/dd/yyyy") & "'"
    End If
       
    If IsDate(txtDOD1) = True Then
        strAppend = strAppend + " And CasDischargeDate >= '" & Format(txtDOD1, "mm/dd/yyyy") & "'"
    End If
    
    If IsDate(txtDOD2) = True Then
        strAppend = strAppend + " And CasDischargeDate <= '" & Format(txtDOD2, "mm/dd/yyyy") & "'"
    End If
    
    If IsDate(TxtDOR1) = True Then
        strAppend = strAppend + " And CaseRegDate >= '" & Format(TxtDOR1, "mm/dd/yyyy") & "'"
    End If
    
    If IsDate(TxtDOR2) = True Then
        strAppend = strAppend + " And CaseRegDate <= '" & Format(TxtDOR2, "mm/dd/yyyy") & "'"
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
    
    If IsDate(TxtPayEffDate1) = True Then
        strAppend = strAppend & " AND MBMPayorEffDate >= '" & Format(TxtPayEffDate1, "m/dd/yyy") & "'"
    End If
    
    If IsDate(TxtPayEffDate2) = True Then
        strAppend = strAppend & " AND MBMPayorEffDate <= '" & Format(TxtPayEffDate2, "m/dd/yyy") & "'"
    End If
    
    If IsDate(txtPayExpDate1) = True Then
        strAppend = strAppend & " AND MBMPayorExpDate >= '" & Format(txtPayExpDate1, "m/dd/yyy") & "'"
    End If
    
    If IsDate(txtPayExpDate2) = True Then
        strAppend = strAppend & " AND MBMPayorExpDate <= '" & Format(txtPayExpDate2, "m/dd/yyy") & "'"
    End If
    
    If IsDate(TxtDateJoin1) = True Then
        strAppend = strAppend & " AND MBMFirstJoinedDate >= '" & Format(TxtDateJoin1, "m/dd/yyy") & "'"
    End If
    
    If IsDate(txtDateJoin2) = True Then
        strAppend = strAppend & " AND MBMFirstJoinedDate <= '" & Format(txtDateJoin2, "m/dd/yyy") & "'"
    End If
    
    If Len(Trim(txtPOLYear1)) Then
        strAppend = strAppend & " AND MBMPolicyYear >= '" & Trim(txtPOLYear1) & "'"
    End If
    
    If Len(Trim(txtPOLYear2)) Then
        strAppend = strAppend & " AND MBMPolicyYear <= '" & Trim(txtPOLYear2) & "'"
    End If
    
    If ChkPayor.value = 1 Then
        strAppend = strAppend & " AND (PAYCode Is null OR PAYCode = '')"
    End If
    
    If ChkGrpCom.value = 1 Then
        strAppend = strAppend & " AND (MBMGroupCompany Is null OR MBMGroupCompany = '') "
    End If
    
    If ChkHOS.value = 1 Then
        strAppend = strAppend & " AND (HOSName Is null OR HOSName = '') "
    End If
    
    If ChkHOSGrp.value = 1 Then
        strAppend = strAppend & " AND (HOSGRPName Is null OR HOSGRPName = '') "
    End If
           
    If ChkDOA.value = 1 Then
        strAppend = strAppend & " AND (CASDateAdmitted Is null OR CASDateAdmitted = '') "
    End If
    
    If ChkDOD.value = 1 Then
        strAppend = strAppend & " AND (CASDischargeDate Is null OR CASDischargeDate = '') "
    End If
    
    If chkDOR.value = 1 Then
        strAppend = strAppend & " AND (CaseRegDate Is null OR CaseRegDate = '') "
    End If
        
    If ChkRep.value = 1 Then
        strAppend = strAppend & " AND (CASRepCode1 Is null OR CASRepCode1 = '') "
    End If
    
    If ChkRepCat.value = 1 Then
        strAppend = strAppend & " AND (RepCatCode Is null OR RepCatCode = '') "
    End If
    
    If ChkPlan.value = 1 Then
        strAppend = strAppend & " AND (PLNCode Is null OR PLNCode = '') "
    End If
    
    If ChkBordx.value = 1 Then
        strAppend = strAppend & " AND (BordxRefNo Is null OR BordxRefNo = '') "
    End If
    
    If ChkEffDate.value = 1 Then
        strAppend = strAppend & " AND (MBMPayorEffDate Is null OR MBMPayorEffDate = '') "
    End If
    
    If chkExpDate.value = 1 Then
        strAppend = strAppend & " AND (MBMPayorExpDate Is null OR MBMPayorExpDate = '') "
    End If
    
    If ChkJoinDate.value = 1 Then
        strAppend = strAppend & " AND (MBMFirstJoinedDate Is null OR MBMFirstJoinedDate = '') "
    End If
       
    If ChkPolYear.value = 1 Then
        strAppend = strAppend & " AND (MBMPolicyYear Is null OR MBMPolicyYear = '') "
    End If
    
    If OptDebit.value = True Then
        DNCNCriteria = strAppend
        Call DNListing(strAppend)
    Else
        DNCNCriteria = strAppend
        Call CRListing(strAppend)
    End If
    
End Function
'******************* End SearchRecord ************************************

'******************* Start DNListing ************************************

Private Sub DNListing(Optional strAppend As String)
Dim rst2 As New ADODB.Recordset
Dim Record As ListItem
Dim sql As String
Dim Total As String
Dim Current As String
Dim TotalAmt As Variant
    
    Total = 0
    Current = 0
    
    lstDebitCreditList.listitems.Clear
            
    sql = " Select DebitCreditRefNo, DebitCreditDate, " & _
          " CASNumber, ClaimVersion, MBMName, PatientName, " & _
          " HOSName, PAYCode, CLMPayableByMedix2M " & _
          " From VIEW_Debit_Tracking where 0 = 0 "
     
    If Len(Trim(strAppend)) Then
        sql = sql + strAppend
    End If
    
    rst2.Open sql, medixmasterconnWS, adOpenForwardOnly, adLockBatchOptimistic
    
    If rst2.EOF = True Then
        MsgBox "No record found!", vbOKOnly + vbCritical, DialogBoxTitle
        CmdPrint.Enabled = True
        cmdFile.Enabled = True
        CmdClear.Enabled = True
        CmdExit.Enabled = True
    Else
         
    Do
        Do Until rst2.EOF
        Total = rst2.recordCount
        lblTotal = Total
        
            With rst2
                Set Record = lstDebitCreditList.listitems.Add(, , rst2!DebitCreditRefNo)
                If Len(rst2!DebitCreditDate) Then
                   Record.SubItems(1) = rst2!DebitCreditDate
                End If
                If Len(rst2!MBMName) Then
                   Record.SubItems(2) = rst2!MBMName
                End If
                If Len(rst2!PatientName) Then
                    Record.SubItems(3) = rst2!PatientName
                End If
                If Len(rst2!CLMPayableByMedix2M) Then
                    Record.SubItems(4) = Format(rst2!CLMPayableByMedix2M, "#,##0.00;(#,##0.00)")
                End If
                If Len(rst2!HOSName) Then
                    Record.SubItems(5) = rst2!HOSName
                End If
                If Len(rst2!PAYCode) Then
                    Record.SubItems(6) = rst2!PAYCode
                End If
                If Len(rst2!CASNumber & "-" & rst2!claimversion) Then
                    Record.SubItems(7) = (rst2!CASNumber & "-" & rst2!claimversion)
                End If
                
                Current = Current + 1
                lblCurrent = Current
                If Len(rst2!CLMPayableByMedix2M) Then
                    TotalAmt = TotalAmt + rst2!CLMPayableByMedix2M
                End If
                LblTotalAmt = Format(TotalAmt, "#,##0.00;(#,##0.00)")
                
            End With
            rst2.MoveNext
            DoEvents
                If intExit = 1 Then
                   Check = False
                   Exit Do
                End If
        Loop
    Loop Until Check = False
    End If
    intExit = 0
    rst2.Close
    Set rst2 = Nothing
    CmdPrint.Enabled = True
    cmdFile.Enabled = True
    CmdClear.Enabled = True
    CmdExit.Enabled = True
End Sub
'******************* End DNListing ************************************

'******************* Start CRListing ************************************

Private Sub CRListing(Optional strAppend As String)
Dim rst2 As New ADODB.Recordset
Dim Record As ListItem
Dim sql As String
Dim Total As String
Dim Current As String
Dim TotalAmt As Variant

    Total = 0
    Current = 0
    
    lstDebitCreditList.listitems.Clear
            
    sql = " Select DebitCreditRefNo, DebitCreditDate, " & _
          " CASNumber, ClaimVersion, MBMName, PatientName, " & _
          " HOSName, PAYCode, CLMPayableByMedix2M " & _
          " From VIEW_Credit_Tracking where 0 = 0 "
     
    If Len(Trim(strAppend)) Then
        sql = sql + strAppend
    End If
    
    rst2.Open sql, medixmasterconnWS, adOpenForwardOnly, adLockBatchOptimistic
    
    If rst2.EOF = True Then
        MsgBox "No record found!", vbOKOnly + vbCritical, DialogBoxTitle
        CmdPrint.Enabled = True
        cmdFile.Enabled = True
        CmdClear.Enabled = True
        CmdExit.Enabled = True
    Else
         
    Do
        Do Until rst2.EOF
        Total = rst2.recordCount
        lblTotal = Total
        
            With rst2
                Set Record = lstDebitCreditList.listitems.Add(, , rst2!DebitCreditRefNo)
                If Len(rst2!DebitCreditDate) Then
                   Record.SubItems(1) = rst2!DebitCreditDate
                End If
                If Len(rst2!MBMName) Then
                   Record.SubItems(2) = rst2!MBMName
                End If
                If Len(rst2!MBMName) Then
                    Record.SubItems(3) = rst2!MBMName
                End If
                If Len(rst2!CLMPayableByMedix2M) Then
                    Record.SubItems(4) = Format(rst2!CLMPayableByMedix2M, "#,##0.00;(#,##0.00)")
                End If
                If Len(rst2!HOSName) Then
                    Record.SubItems(5) = rst2!HOSName
                End If
                If Len(rst2!PAYCode) Then
                    Record.SubItems(6) = rst2!PAYCode
                End If
                If Len(rst2!CASNumber & "-" & rst2!claimversion) Then
                    Record.SubItems(7) = (rst2!CASNumber & "-" & rst2!claimversion)
                End If
                
                Current = Current + 1
                lblCurrent = Current
                If Len(rst2!CLMPayableByMedix2M) Then
                    TotalAmt = TotalAmt + rst2!CLMPayableByMedix2M
                End If
                LblTotalAmt = Format(TotalAmt, "#,##0.00;(#,##0.00)")
                
            End With
            rst2.MoveNext
            DoEvents
                If intExit = 1 Then
                   Check = False
                   Exit Do
                End If
        Loop
    Loop Until Check = False
    End If
    intExit = 0
    rst2.Close
    Set rst2 = Nothing
    CmdPrint.Enabled = True
    cmdFile.Enabled = True
    CmdClear.Enabled = True
    CmdExit.Enabled = True
End Sub
'******************* End CRListing ************************************

'******************* Start PrintDebit ************************************

Private Sub ExportToFile()
Dim OutputStr As String
Dim i As Integer
Dim totalrecord As Integer
Dim Record As ListItem
Dim FileDate As String

    Screen.MousePointer = vbHourglass
    
    FileDate = Format(txtDate, "ddmmyyyy")
    txtFilName = "D" & Trim(Mid(CboPayor, 1, 2)) & Trim(FileDate)
    'Open path & "\" & txtFilName For Output As #1
    Open path & txtFilName For Output As #1
    header = Trim(Mid(CboPayor, 1, 2)) & "-" & "Debit Note Listing"
    Print #1, Trim(header)
    totalrecord = lstDebitCreditList.listitems.count
    For i = 1 To totalrecord
        Set Record = lstDebitCreditList.listitems.Item(i)
        OutputStr = Trim(Record.Text)
        OutputStr = OutputStr & vbTab & Trim(Record.SubItems(1))
        OutputStr = OutputStr & vbTab & Record.SubItems(2)
        OutputStr = OutputStr & vbTab & Trim(Record.SubItems(3))
        OutputStr = OutputStr & vbTab & Trim(Record.SubItems(4))
        OutputStr = OutputStr & vbTab & Trim(Record.SubItems(5))
        OutputStr = OutputStr & vbTab & Trim(Record.SubItems(6))
        OutputStr = OutputStr & vbTab & Trim(Record.SubItems(7))
        Print #1, OutputStr
    Next
    
    Close #1
    MsgBox "Print to File Sucessfully...", vbOKOnly + vbInformation, DialogBoxTitle
    Screen.MousePointer = vbDefault
End Sub
'******************* End PrintDebit ************************************

'******************* Start PrintCredit ************************************

Private Sub ExportToFileC()
Dim OutputStr As String
Dim i As Integer
Dim totalrecord As Integer
Dim Record As ListItem
Dim FileDate As String

    Screen.MousePointer = vbHourglass
    FileDate = Format(txtDate, "ddmmyyyy")
    txtFilName = "C" & Trim(Mid(CboPayor, 1, 2)) & Trim(FileDate)
    
    'Open path & "\" & txtFilName For Output As #1
    Open path & txtFilName For Output As #1
    header = Trim(Mid(CboPayor, 1, 2)) & "-" & "Credit Note Listing"
    Print #1, Trim(header)
    totalrecord = lstDebitCreditList.listitems.count
    For i = 1 To totalrecord
        Set Record = lstDebitCreditList.listitems.Item(i)
        OutputStr = Trim(Record.Text)
        OutputStr = OutputStr & vbTab & Trim(Record.SubItems(1))
        OutputStr = OutputStr & vbTab & Record.SubItems(2)
        OutputStr = OutputStr & vbTab & Trim(Record.SubItems(3))
        OutputStr = OutputStr & vbTab & Trim(Record.SubItems(4))
        OutputStr = OutputStr & vbTab & Trim(Record.SubItems(5))
        OutputStr = OutputStr & vbTab & Trim(Record.SubItems(6))
        OutputStr = OutputStr & vbTab & Trim(Record.SubItems(7))
        Print #1, OutputStr
    Next
    
    Close #1
    MsgBox "Print to File Sucessfully...", vbOKOnly + vbInformation, DialogBoxTitle
    Screen.MousePointer = vbDefault
End Sub


'Private Sub PrintCredit()
'D 'im strsql As String
'Di'm DebitCredit As New ADODB.Recordset

'Screen.MousePointer = vbHourglass

'strsql = "Select * From VIEW_DebitCreditListing where substring(DebitCreditRefNo ,1 ,1) = 'C'" & _
'"AND PAYCode = '" & Trim(Mid(CboPayor, 1, 2)) & "'"

'DebitCredit.Open strsql, medixmasterconnWS, adOpenKeyset, adLockOptimistic
'
 '   txtFilName = "C" & Trim(Mid(CboPayor, 1, 2))
  '  Open path & "\" & txtFilName For Output As #1
    'Open txtLocalFileLocation & txtFilName For Output As #2
   ' header = Trim(Mid(CboPayor, 1, 2)) & "-" & "Credit Note Listing"
    'Print #1, Trim(header)
    'Print #2, Trim(header)
    
'    Do Until DebitCredit.EOF
    
 '       DataString = DataString & l(DebitCredit!DebitCreditRefNo, 15)
  '      If Len(DebitCredit!DebitCreditDate) Then
   '     DataString = DataString & l(DebitCredit!DebitCreditDate, 10)
    '    End If
     '   DataString = DataString & l(DebitCredit!MBMName, 60)
        ''DataString = DataString & l(DebitCredit!PatientName, 60)
'        DataString = DataString & l(DebitCredit!CLMPayableByMedix2M, 8)
 '       DataString = DataString & l(DebitCredit!CASHOSCode, 6)
    '    DataString = DataString & l(DebitCredit!PAYCode, 2)
  ''      DataString = DataString & l(DebitCredit!PayTypeCode, 2)
     '   DebitCredit.MoveNext
        
'        Print #1, DataString
        'Print #2, DataString
 '   Loop
  '      Close #1
        'Close #2
   '     MsgBox "Print to File Sucessfully...", vbOKOnly + vbInformation, DialogBoxTitle
    '    Screen.MousePointer = vbDefault
'End Sub
'******************* End PrintCredit ************************************

'******************* Start ComboListing ************************************
Private Sub ComboListing()
Dim PAY As New ADODB.Recordset
Dim INS As New ADODB.Recordset
Dim REL As New ADODB.Recordset
Dim OCP As New ADODB.Recordset
Dim PolStatus As New ADODB.Recordset
Dim State As New ADODB.Recordset
Dim HOS As New ADODB.Recordset
Dim HosGrp As New ADODB.Recordset
Dim PAYGRP As New ADODB.Recordset
Dim RepCat As New ADODB.Recordset
Dim RepCode As New ADODB.Recordset

    INS.Open "Select INSCode, INSDescription from INS", medixmasterconn, adOpenKeyset, adLockOptimistic
    PAY.Open "Select PAYCode, PAYCompanyName from PAY", medixmasterconn, adOpenKeyset, adLockOptimistic
    PolStatus.Open "Select pstatuscode ,pstatusDesc from PolicyStatus", medixmasterconn, adOpenKeyset, adLockOptimistic
    State.Open "Select STADescription from STA", medixmasterconn, adOpenKeyset, adLockOptimistic
    HOS.Open "Select HOSCode, HOSName from HOS", medixmasterconn, adOpenKeyset, adLockOptimistic
    HosGrp.Open "Select HOSGRPCode, HOSGRPName from HOSGRP", medixmasterconn, adOpenKeyset, adLockOptimistic
    RepCat.Open "Select REpCatCode, REpCatName from RepCodeCategory", medixmasterconn, adOpenKeyset, adLockOptimistic
    RepCode.Open "Select REpCode, REpDescription from RepCode", medixmasterconn, adOpenKeyset, adLockOptimistic

    cboHLTCode.AddItem "S" & " - " & "Sihat"
    cboHLTCode.AddItem "N" & " - " & "Non-Sihat"
    
    Do Until INS.EOF
        CboInsType.AddItem ChkStr(INS!INSCode) & " - " & ChkStr(INS!INSDescription)
        INS.MoveNext
    Loop
    INS.Close
    Set INS = Nothing
    
    With CboClaimability
        .AddItem "A - Accepted"
        .AddItem "R - Rejected"
    End With
    
    With CboCaseStatus
        .AddItem "O - Open"
        .AddItem "D - Delete"
        .AddItem "A - All"
    End With
    
    With CboWsStatus
        .AddItem "O - Open"
        .AddItem "C - Close"
    End With
    
    Do Until State.EOF
        CboState.AddItem ChkStr(State!STADescription)
        State.MoveNext
    Loop
    State.Close
    Set State = Nothing
        
    Do Until PAY.EOF
        CboPayor.AddItem PAY!PAYCode & " - " & Trim(PAY!PAYCompanyName)
        PAY.MoveNext
    Loop
    PAY.Close
    Set PAY = Nothing
    
    Do Until HOS.EOF
        cboHOSName.AddItem Trim(HOS!HOSName)
        HOS.MoveNext
    Loop
    HOS.Close
    Set HOS = Nothing
    
    Do Until HosGrp.EOF
        cboHOSGrp.AddItem Trim(HosGrp!HOSGRPName)
        HosGrp.MoveNext
    Loop
    HosGrp.Close
    Set HosGrp = Nothing
      
    Do Until RepCat.EOF
        cboRepCat1.AddItem RepCat!REpCatCode & " - " & Trim(RepCat!RepCatNAme)
        cboRepCat2.AddItem RepCat!REpCatCode & " - " & Trim(RepCat!RepCatNAme)
        RepCat.MoveNext
    Loop
    RepCat.Close
    Set RepCat = Nothing
    
    Do Until RepCode.EOF
        cboRepCode1.AddItem RepCode!RepCode & " - " & Trim(RepCode!RepDescription)
        cboRepCode2.AddItem RepCode!RepCode & " - " & Trim(RepCode!RepDescription)
        RepCode.MoveNext
    Loop
    RepCode.Close
    Set RepCode = Nothing
    
End Sub
'============== ComboBox Listing ============
'******************* End ComboListing ************************************

'**********************Start Sorted ListView *******************************
Private Sub lstDebitCreditList_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
    m_blnDirection(ColumnHeader.Index) = Not m_blnDirection(ColumnHeader.Index)
    Select Case ColumnHeader.Index
    
    Case 1
        SortListView lstDebitCreditList, ColumnHeader.Index, ldtString, m_blnDirection(1)
    Case 2
        SortListView lstDebitCreditList, ColumnHeader.Index, ldtDateTime, m_blnDirection(2)
    Case 3
        SortListView lstDebitCreditList, ColumnHeader.Index, ldtString, m_blnDirection(3)
    Case 4
        SortListView lstDebitCreditList, ColumnHeader.Index, ldtString, m_blnDirection(4)
    Case 5
        SortListView lstDebitCreditList, ColumnHeader.Index, ldtNumber, m_blnDirection(5)
    Case 6
        SortListView lstDebitCreditList, ColumnHeader.Index, ldtString, m_blnDirection(6)
    Case 7
        SortListView lstDebitCreditList, ColumnHeader.Index, ldtString, m_blnDirection(7)
    Case 8
        SortListView lstDebitCreditList, ColumnHeader.Index, ldtString, m_blnDirection(8)
    'Case 9
        'SortListView lstDebitCreditList, ColumnHeader.Index, ldtString, m_blnDirection(9)
    End Select
End Sub
'**********************End Sorted ListView *******************************

'******************* Start ComboBox Change/keyPress ************************************

Private Sub CboPayor_Click()
Dim strsql As String
Dim PAYGRP As New ADODB.Recordset

    CboGrpCom.Clear
    medixmasterconn.Open "File Name=" & BOSPath & "Conn\mediconnect.UDL"

    strsql = "select MBMGroupCompany from View_PAYGrpCom where PAYCode= '" & Mid(CboPayor, 1, 2) & "'"
    PAYGRP.Open strsql, medixmasterconn, adOpenKeyset, adLockOptimistic
    
    If PAYGRP.recordCount Then
        Do Until PAYGRP.EOF
           CboGrpCom.AddItem Trim(PAYGRP!MBMGroupCompany)
           PAYGRP.MoveNext
        Loop
    End If
    PAYGRP.Close
    Set PAYGRP = Nothing
    medixmasterconn.Close
    Set medixmasterconn = Nothing

If CboPayor.Text = "" Then
    lblCurrent = 0
    lblTotal = 0
    LblTotalAmt = 0
    lstDebitCreditList.listitems.Clear
End If

End Sub

Private Sub CboPayor_Change()

If CboPayor.Text = "" Then
    If InStr(CboPayor.Text, "null") Then
       CboPayor.Enabled = False
    End If
End If
    lstDebitCreditList.listitems.Clear
    lblCurrent = 0
    lblTotal = 0
    LblTotalAmt = 0

End Sub




'Private Sub CboPayor_KeyPress(KeyAscii As Integer)
'Dim NewText As String
'Dim ValidCount As Integer
'Dim ValidValue As String
'Dim i As Integer
 '   If KeyAscii >= 32 And KeyAscii <> 127 Then
  '  NewText = LCase(Left(CboPayor.Text, CboPayor.SelStart) + Chr(KeyAscii) + Mid(CboPayor.Text, CboPayor.SelStart + CboPayor.SelLength + 1))
 
   ' ValidCount = 0
    'ValidValue = ""
'
  '  For i = 0 To CboPayor.ListCount - 1
 '
   '     If NewText = LCase(Left(CboPayor.list(i), Len(NewText))) Then
    '            ValidCount = ValidCount + 1
     '           ValidValue = CboPayor.list(i)
      '  End If
'        Next
 '       If ValidCount <= 1 Then KeyAscii = 0
  '          If ValidCount = 1 Then
   '             CboPayor.Text = ValidValue
    '            CboPayor.SelStart = 0
     '           CboPayor.SelLength = Len(ValidValue)
      '      End If
'        End If
'E'nd Sub
'******************* End ComboBox Change/keyPress ************************************

'*************************** Start Diable Non Integer Value
Private Sub txtAmtFr_KeyPress(KeyAscii As Integer)
    KeyAscii = KeyNumber(KeyAscii)
End Sub

Private Sub txtAmtTo_KeyPress(KeyAscii As Integer)
    KeyAscii = KeyNumber(KeyAscii)
End Sub
'*************************** End Diable Non Integer Value


Private Sub txtPVDate1_LostFocus()
    If IsDate(TxtPVDate1) Then
    Else
        If TxtPVDate1 <> "__/__/____" Then
            MsgBox "Invalid Date Format! ", vbCritical
            TxtPVDate1.SetFocus
        End If
    End If
End Sub

Private Sub txtPVDate2_LostFocus()
    If IsDate(TxtPVDate2) Then
    Else
        If TxtPVDate2 <> "__/__/____" Then
            MsgBox "Invalid Date Format! ", vbCritical
            TxtPVDate2.SetFocus
        End If
    End If
End Sub


Private Sub TxtDateJoin1_LostFocus()
   If IsDate(TxtDateJoin1) Then
    Else
        If TxtDateJoin1 <> "__/__/____" Then
            MsgBox "Invalid Date Format! ", vbCritical
            TxtDateJoin1.SetFocus
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


Private Sub TxtDOA1_LostFocus()
   If IsDate(txtDOA1) Then
    Else
        If txtDOA1 <> "__/__/____" Then
            MsgBox "Invalid Date Format! ", vbCritical
            txtDOA1.SetFocus
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
   If IsDate(txtPayExpDate1) Then
    Else
        If txtPayExpDate1 <> "__/__/____" Then
            MsgBox "Invalid Date Format! ", vbCritical
            txtPayExpDate1.SetFocus
        End If
    End If
End Sub

Private Sub TxtPayExpDate2_LostFocus()
   If IsDate(txtPayExpDate2) Then
    Else
        If txtPayExpDate2 <> "__/__/____" Then
            MsgBox "Invalid Date Format! ", vbCritical
            txtPayExpDate2.SetFocus
        End If
    End If
End Sub

Private Sub txtPOLYear1_KeyPress(KeyAscii As Integer)
    KeyAscii = KeyNumber(KeyAscii)
End Sub

Private Sub txtPOLYear2_KeyPress(KeyAscii As Integer)
    KeyAscii = KeyNumber(KeyAscii)
End Sub

Private Sub txtRecDate1_LostFocus()
   If IsDate(txtRecDate1) Then
    Else
        If txtRecDate1 <> "__/__/____" Then
            MsgBox "Invalid Date Format! ", vbCritical
            txtRecDate1.SetFocus
        End If
    End If
End Sub

Private Sub txtRecDate2_LostFocus()
   If IsDate(txtRecDate2) Then
    Else
        If txtRecDate2 <> "__/__/____" Then
            MsgBox "Invalid Date Format! ", vbCritical
            txtRecDate2.SetFocus
        End If
    End If
End Sub
