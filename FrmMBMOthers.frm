VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form FrmMBMOthers 
   Caption         =   "Member Others"
   ClientHeight    =   5550
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   11160
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   5550
   ScaleWidth      =   11160
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame Frame3 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   315
      Left            =   0
      TabIndex        =   68
      Top             =   0
      Width           =   12120
      Begin VB.Frame Frame8 
         Caption         =   "Frame8"
         Height          =   15
         Left            =   240
         TabIndex        =   69
         Top             =   480
         Width           =   7215
      End
      Begin VB.Label Label3 
         Appearance      =   0  'Flat
         BackColor       =   &H00000000&
         Caption         =   "Member Others"
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
         Height          =   255
         Left            =   240
         TabIndex        =   70
         Top             =   0
         Width           =   6615
      End
   End
   Begin VB.Frame Frame5 
      Height          =   4635
      Left            =   0
      TabIndex        =   40
      Top             =   360
      Width           =   11100
      Begin VB.ComboBox cboMBMOccupation 
         Height          =   315
         Left            =   1320
         TabIndex        =   0
         Top             =   120
         Width           =   3640
      End
      Begin VB.TextBox txtMBMAllergic 
         Height          =   1785
         Left            =   5040
         MaxLength       =   500
         MultiLine       =   -1  'True
         TabIndex        =   37
         Top             =   2520
         Width           =   5940
      End
      Begin VB.Frame Frame12 
         Height          =   495
         Left            =   5040
         TabIndex        =   50
         Top             =   1680
         Width           =   3255
         Begin MSMask.MaskEdBox TxtClmNonPayFr 
            BeginProperty DataFormat 
               Type            =   1
               Format          =   "dd/MM/yyyy"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1033
               SubFormatType   =   3
            EndProperty
            Height          =   250
            Left            =   555
            TabIndex        =   33
            Top             =   150
            Width           =   1095
            _ExtentX        =   1931
            _ExtentY        =   450
            _Version        =   393216
            MaxLength       =   10
            Mask            =   "##/##/####"
            PromptChar      =   "_"
         End
         Begin MSMask.MaskEdBox TxtClmNonPayTo 
            BeginProperty DataFormat 
               Type            =   1
               Format          =   "dd/MM/yyyy"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1033
               SubFormatType   =   3
            EndProperty
            Height          =   255
            Left            =   2040
            TabIndex        =   34
            Top             =   150
            Width           =   1095
            _ExtentX        =   1931
            _ExtentY        =   450
            _Version        =   393216
            MaxLength       =   10
            Mask            =   "##/##/####"
            PromptChar      =   "_"
         End
         Begin VB.Label Label23 
            Alignment       =   2  'Center
            Caption         =   "To"
            Height          =   255
            Left            =   1680
            TabIndex        =   52
            Top             =   200
            Width           =   285
         End
         Begin VB.Label Label36 
            Alignment       =   2  'Center
            Caption         =   "From"
            Height          =   255
            Left            =   80
            TabIndex        =   51
            Top             =   200
            Width           =   405
         End
      End
      Begin VB.Frame Frame11 
         Height          =   1335
         Left            =   5040
         TabIndex        =   42
         Top             =   120
         Width           =   6015
         Begin VB.TextBox TxtGenWP 
            Height          =   285
            Left            =   1320
            MaxLength       =   50
            TabIndex        =   26
            Top             =   240
            Width           =   705
         End
         Begin VB.TextBox txtPREWP 
            Height          =   285
            Left            =   1320
            MaxLength       =   5
            TabIndex        =   28
            Top             =   480
            Width           =   705
         End
         Begin VB.TextBox TxtSpeWP 
            Height          =   300
            Left            =   1320
            MaxLength       =   50
            TabIndex        =   30
            Top             =   720
            Width           =   705
         End
         Begin VB.ComboBox cboStaffDisc 
            Height          =   315
            Left            =   1320
            Sorted          =   -1  'True
            TabIndex        =   32
            Top             =   960
            Width           =   705
         End
         Begin MSMask.MaskEdBox txtGENWPEnd 
            BeginProperty DataFormat 
               Type            =   1
               Format          =   "dd/MM/yyyy"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1033
               SubFormatType   =   3
            EndProperty
            Height          =   315
            Left            =   3240
            TabIndex        =   27
            Top             =   240
            Width           =   1095
            _ExtentX        =   1931
            _ExtentY        =   556
            _Version        =   393216
            Enabled         =   0   'False
            MaxLength       =   10
            Mask            =   "##/##/####"
            PromptChar      =   "_"
         End
         Begin MSMask.MaskEdBox txtPREWPEnd 
            BeginProperty DataFormat 
               Type            =   1
               Format          =   "dd/MM/yyyy"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1033
               SubFormatType   =   3
            EndProperty
            Height          =   315
            Left            =   3240
            TabIndex        =   29
            Top             =   480
            Width           =   1095
            _ExtentX        =   1931
            _ExtentY        =   556
            _Version        =   393216
            Enabled         =   0   'False
            MaxLength       =   10
            Mask            =   "##/##/####"
            PromptChar      =   "_"
         End
         Begin MSMask.MaskEdBox txtSPEWPEnd 
            BeginProperty DataFormat 
               Type            =   1
               Format          =   "dd/MM/yyyy"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1033
               SubFormatType   =   3
            EndProperty
            Height          =   315
            Left            =   3240
            TabIndex        =   31
            Top             =   720
            Width           =   1095
            _ExtentX        =   1931
            _ExtentY        =   556
            _Version        =   393216
            Enabled         =   0   'False
            MaxLength       =   10
            Mask            =   "##/##/####"
            PromptChar      =   "_"
         End
         Begin VB.Label Label99 
            Caption         =   "Staff Disc Req"
            Height          =   255
            Left            =   120
            TabIndex        =   49
            Top             =   960
            Width           =   1065
         End
         Begin VB.Label Label100 
            Caption         =   "End of Period"
            Height          =   255
            Left            =   2160
            TabIndex        =   48
            Top             =   765
            Width           =   1005
         End
         Begin VB.Label Label98 
            Caption         =   "End of Period"
            Height          =   255
            Left            =   2160
            TabIndex        =   47
            Top             =   480
            Width           =   1005
         End
         Begin VB.Label Label97 
            Caption         =   "End of Period"
            Height          =   255
            Left            =   2160
            TabIndex        =   46
            Top             =   240
            Width           =   1005
         End
         Begin VB.Label Label96 
            Caption         =   "Gen W. P."
            Height          =   255
            Left            =   120
            TabIndex        =   45
            Top             =   240
            Width           =   1620
         End
         Begin VB.Label Label95 
            Caption         =   "Pre W. P."
            Height          =   255
            Left            =   120
            TabIndex        =   44
            Top             =   480
            Width           =   1380
         End
         Begin VB.Label Label94 
            Caption         =   "Spe. Ill. W. P."
            Height          =   255
            Left            =   120
            TabIndex        =   43
            Top             =   720
            Width           =   2025
         End
      End
      Begin VB.ComboBox txtMBMClinical 
         Height          =   315
         Left            =   7200
         TabIndex        =   41
         Top             =   5400
         Width           =   1290
      End
      Begin VB.Frame Frame14 
         Height          =   495
         Left            =   8400
         TabIndex        =   53
         Top             =   1680
         Width           =   2535
         Begin VB.TextBox txtPLNCode 
            Height          =   225
            Left            =   480
            MaxLength       =   4
            TabIndex        =   35
            Top             =   150
            Width           =   615
         End
         Begin MSMask.MaskEdBox txtNewPolEffDate 
            BeginProperty DataFormat 
               Type            =   1
               Format          =   "dd/MM/yyyy"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1033
               SubFormatType   =   3
            EndProperty
            Height          =   270
            Left            =   1440
            TabIndex        =   36
            Top             =   150
            Width           =   975
            _ExtentX        =   1720
            _ExtentY        =   476
            _Version        =   393216
            MaxLength       =   10
            Mask            =   "##/##/####"
            PromptChar      =   "_"
         End
         Begin VB.Label Label39 
            Alignment       =   2  'Center
            Caption         =   "PLN"
            Height          =   255
            Left            =   80
            TabIndex        =   55
            Top             =   165
            Width           =   405
         End
         Begin VB.Label Label38 
            Alignment       =   2  'Center
            Caption         =   "Eff"
            Height          =   255
            Left            =   1150
            TabIndex        =   54
            Top             =   165
            Width           =   285
         End
      End
      Begin VB.ComboBox txtMBMWorkNature 
         Height          =   315
         Left            =   1320
         TabIndex        =   1
         Top             =   405
         Width           =   3640
      End
      Begin VB.TextBox txtMBMWeight 
         Height          =   285
         Left            =   1320
         MaxLength       =   3
         TabIndex        =   2
         Top             =   690
         Width           =   1215
      End
      Begin VB.TextBox txtMBMHeight 
         Height          =   285
         Left            =   3785
         MaxLength       =   3
         TabIndex        =   3
         Top             =   690
         Width           =   1185
      End
      Begin VB.TextBox txtMBMEmail 
         Height          =   285
         Left            =   1320
         MaxLength       =   3
         TabIndex        =   4
         Top             =   950
         Width           =   3645
      End
      Begin VB.CheckBox Check1 
         Caption         =   "Check1"
         Height          =   255
         Left            =   1320
         TabIndex        =   25
         Top             =   4275
         Width           =   255
      End
      Begin VB.ComboBox cboLGNCode 
         Height          =   315
         Left            =   1320
         TabIndex        =   5
         Top             =   1185
         Width           =   1215
      End
      Begin VB.ComboBox cboMBMSmoking 
         Height          =   315
         Left            =   1320
         TabIndex        =   7
         Top             =   1470
         Width           =   1215
      End
      Begin VB.ComboBox cboMBMBlood 
         Height          =   315
         Left            =   3780
         TabIndex        =   6
         Top             =   1185
         Width           =   1185
      End
      Begin VB.ComboBox cboMBMAlcohol 
         Height          =   315
         Left            =   3780
         TabIndex        =   8
         Top             =   1470
         Width           =   1185
      End
      Begin VB.TextBox txtMBMTelNoH 
         Height          =   315
         Left            =   1320
         MaxLength       =   12
         TabIndex        =   9
         Top             =   1755
         Width           =   1215
      End
      Begin VB.TextBox txtMBMTelNoM 
         Height          =   315
         Left            =   3765
         MaxLength       =   15
         TabIndex        =   10
         Top             =   1755
         Width           =   1200
      End
      Begin VB.TextBox txtMBMTelNoO 
         Height          =   315
         Left            =   1320
         MaxLength       =   15
         TabIndex        =   11
         Top             =   2040
         Width           =   1215
      End
      Begin VB.TextBox txtMBMRBAmt 
         Height          =   330
         Left            =   3765
         MaxLength       =   8
         TabIndex        =   12
         Top             =   2040
         Width           =   1200
      End
      Begin VB.TextBox txtMBMPrevPlan 
         Height          =   330
         Left            =   1320
         MaxLength       =   15
         TabIndex        =   13
         Top             =   2325
         Width           =   1215
      End
      Begin VB.ComboBox cboCOPayRB 
         Height          =   315
         ItemData        =   "FrmMBMOthers.frx":0000
         Left            =   3765
         List            =   "FrmMBMOthers.frx":0007
         TabIndex        =   14
         Top             =   2340
         Width           =   1200
      End
      Begin VB.TextBox txtMBMPrevINSCom 
         Height          =   330
         Left            =   1320
         MaxLength       =   60
         TabIndex        =   15
         Top             =   2625
         Width           =   3645
      End
      Begin VB.TextBox txtMBMPAYRemarks 
         Height          =   300
         Left            =   1320
         MaxLength       =   60
         MultiLine       =   -1  'True
         TabIndex        =   16
         Top             =   2925
         Width           =   3645
      End
      Begin VB.ComboBox cboBTBG 
         Height          =   315
         ItemData        =   "FrmMBMOthers.frx":000F
         Left            =   1320
         List            =   "FrmMBMOthers.frx":0016
         TabIndex        =   17
         Top             =   3195
         Width           =   1215
      End
      Begin VB.TextBox TxtPOLSubNo 
         Height          =   300
         Left            =   3840
         MaxLength       =   60
         MultiLine       =   -1  'True
         TabIndex        =   18
         Top             =   3195
         Width           =   1125
      End
      Begin VB.TextBox txtPayPrem 
         Height          =   285
         Left            =   1320
         MaxLength       =   8
         TabIndex        =   19
         Top             =   3480
         Width           =   1215
      End
      Begin VB.TextBox txtPayPremNett 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   3840
         MaxLength       =   8
         TabIndex        =   20
         Top             =   3465
         Width           =   1125
      End
      Begin VB.TextBox txtPayMCOGross 
         Height          =   285
         Left            =   1320
         MaxLength       =   8
         TabIndex        =   21
         Top             =   3735
         Width           =   1215
      End
      Begin VB.TextBox txtPolicyDisc 
         Height          =   285
         Left            =   1320
         MaxLength       =   8
         TabIndex        =   23
         Top             =   3990
         Width           =   1215
      End
      Begin VB.TextBox txtPayMCONett 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   3840
         MaxLength       =   8
         TabIndex        =   22
         Top             =   3720
         Width           =   1125
      End
      Begin VB.TextBox txtRenewalDisc 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   3840
         MaxLength       =   8
         TabIndex        =   24
         Top             =   3975
         Width           =   1125
      End
      Begin VB.Label Label15 
         Caption         =   "Email"
         Height          =   255
         Left            =   120
         TabIndex        =   90
         Top             =   960
         Width           =   1005
      End
      Begin VB.Label Label31 
         Caption         =   "Occupation"
         Height          =   255
         Left            =   120
         TabIndex        =   89
         Top             =   180
         Width           =   1005
      End
      Begin VB.Label Label14 
         Caption         =   "Work Nature"
         Height          =   255
         Left            =   120
         TabIndex        =   88
         Top             =   435
         Width           =   1140
      End
      Begin VB.Label Label35 
         Caption         =   "Height (cm)"
         Height          =   255
         Left            =   2640
         TabIndex        =   83
         Top             =   725
         Width           =   1005
      End
      Begin VB.Label Label34 
         Caption         =   "Weight (kgs)"
         Height          =   255
         Left            =   120
         TabIndex        =   82
         Top             =   720
         Width           =   1005
      End
      Begin VB.Label Label105 
         Caption         =   "New Plan Code"
         Height          =   255
         Left            =   8400
         TabIndex        =   67
         Top             =   1500
         Width           =   1245
      End
      Begin VB.Label Label41 
         Caption         =   "Allergic"
         Height          =   255
         Left            =   5040
         TabIndex        =   66
         Top             =   2280
         Width           =   735
      End
      Begin VB.Label Label13 
         Caption         =   "Claim Non-Pay Date"
         Height          =   255
         Left            =   5040
         TabIndex        =   65
         Top             =   1500
         Width           =   1725
      End
      Begin VB.Label Label12 
         Caption         =   "Alcohol (Y/N)"
         Height          =   255
         Left            =   2640
         TabIndex        =   87
         Top             =   1500
         Width           =   1275
      End
      Begin VB.Label Label11 
         Caption         =   "Smoking (Y/N)"
         Height          =   255
         Left            =   120
         TabIndex        =   86
         Top             =   1545
         Width           =   1275
      End
      Begin VB.Label Label37 
         Caption         =   "Blood Type"
         Height          =   255
         Left            =   2640
         TabIndex        =   85
         Top             =   1260
         Width           =   1005
      End
      Begin VB.Label Label10 
         Caption         =   "Pref. Language"
         Height          =   255
         Left            =   120
         TabIndex        =   84
         Top             =   1200
         Width           =   1140
      End
      Begin VB.Label Label9 
         Caption         =   "Renewal Disc."
         Height          =   255
         Left            =   2640
         TabIndex        =   81
         Top             =   3990
         Width           =   1485
      End
      Begin VB.Label Label8 
         Caption         =   "Policy Disc."
         Height          =   255
         Left            =   120
         TabIndex        =   80
         Top             =   4020
         Width           =   1365
      End
      Begin VB.Label Label6 
         Caption         =   "Pay MCO Nett"
         Height          =   255
         Left            =   2640
         TabIndex        =   79
         Top             =   3705
         Width           =   1245
      End
      Begin VB.Label Label5 
         Caption         =   "Pay MCO Gross"
         Height          =   255
         Left            =   120
         TabIndex        =   78
         Top             =   3720
         Width           =   1365
      End
      Begin VB.Label Label112 
         Caption         =   "Pay Prem Nett"
         Height          =   255
         Left            =   2640
         TabIndex        =   77
         Top             =   3495
         Width           =   1245
      End
      Begin VB.Label Label110 
         Caption         =   "Pay Prem Gross"
         Height          =   255
         Left            =   120
         TabIndex        =   74
         Top             =   3465
         Width           =   1365
      End
      Begin VB.Label Label106 
         Caption         =   "BTB Guarantee"
         Height          =   255
         Left            =   120
         TabIndex        =   73
         Top             =   3195
         Width           =   1140
      End
      Begin VB.Label Label1 
         Caption         =   "Pol Sub No"
         Height          =   255
         Left            =   2640
         TabIndex        =   72
         Top             =   3270
         Width           =   1125
      End
      Begin VB.Label Label93 
         Caption         =   "Co-Pay RB"
         Height          =   240
         Left            =   2640
         TabIndex        =   64
         Top             =   2355
         Width           =   1140
      End
      Begin VB.Label Label7 
         Caption         =   "Label Status"
         Height          =   255
         Left            =   120
         TabIndex        =   63
         Top             =   4275
         Width           =   885
      End
      Begin VB.Label Label101 
         Caption         =   "Payor Remarks"
         Height          =   255
         Left            =   120
         TabIndex        =   62
         Top             =   2955
         Width           =   1095
      End
      Begin VB.Label Label88 
         Caption         =   "T/O Pln Code"
         Height          =   255
         Left            =   120
         TabIndex        =   61
         Top             =   2355
         Width           =   1125
      End
      Begin VB.Label Label87 
         Caption         =   "Prev Ins Com"
         Height          =   255
         Left            =   120
         TabIndex        =   60
         Top             =   2670
         Width           =   1140
      End
      Begin VB.Label Label86 
         Caption         =   "R&&B Amt"
         Height          =   255
         Left            =   2640
         TabIndex        =   59
         Top             =   2130
         Width           =   780
      End
      Begin VB.Label Label28 
         Caption         =   "Tel Num (H)"
         Height          =   255
         Left            =   120
         TabIndex        =   58
         Top             =   1800
         Width           =   1005
      End
      Begin VB.Label Label32 
         Caption         =   "H/P Num"
         Height          =   255
         Left            =   2640
         TabIndex        =   57
         Top             =   1800
         Width           =   1005
      End
      Begin VB.Label Label33 
         Caption         =   "Tel Num (O)"
         Height          =   255
         Left            =   120
         TabIndex        =   56
         Top             =   2085
         Width           =   1005
      End
   End
   Begin VB.Frame Frame2 
      Height          =   615
      Left            =   0
      TabIndex        =   71
      Top             =   4920
      Width           =   11100
      Begin VB.CommandButton cmdClose 
         Cancel          =   -1  'True
         Caption         =   "&Close"
         Height          =   375
         Left            =   9240
         TabIndex        =   39
         Top             =   165
         Visible         =   0   'False
         Width           =   735
      End
      Begin VB.CommandButton cmdOK 
         Caption         =   "&OK"
         Height          =   375
         Left            =   9960
         TabIndex        =   38
         Top             =   160
         Width           =   735
      End
      Begin VB.Label Label4 
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
         Left            =   1080
         TabIndex        =   76
         Top             =   240
         Width           =   1695
      End
      Begin VB.Label Label46 
         Caption         =   "Member No:"
         Height          =   255
         Left            =   120
         TabIndex        =   75
         Top             =   240
         Width           =   975
      End
   End
End
Attribute VB_Name = "FrmMBMOthers"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public Source As Integer

Public Sub cmdClose_Click()
'Dim RecordExit
'
'    RecordExit = MsgBox("Quit this form?", vbYesNo + vbExclamation)
'    If RecordExit = vbYes Then
'        bExit = True
'       Unload Me
'    End If
    Call showMBMOthers
End Sub

Private Sub cmdOK_Click()
    If Source = 3 Then
        If txtPolicyDisc <> "" Then
            frmMemberRegistration.PolDics = txtPolicyDisc
        Else
            frmMemberRegistration.PolDics = "0"
        End If
        
        If txtRenewalDisc <> "" Then
            frmMemberRegistration.RenDisc = txtRenewalDisc
        Else
            frmMemberRegistration.RenDisc = "0"
        End If
        frmMemberRegistration.MemberCallFees
    ElseIf Source = 1 Then
        If txtPolicyDisc <> "" Then
            frmMBMAdjustment.PolDics = txtPolicyDisc
        Else
            frmMBMAdjustment.PolDics = "0"
        End If
        
        If txtRenewalDisc <> "" Then
            frmMBMAdjustment.RenDisc = txtRenewalDisc
        Else
            frmMBMAdjustment.RenDisc = "0"
        End If
        frmMBMAdjustment.MemberCallFees
    End If
    FrmMBMOthers.Visible = False
End Sub

Public Function showMBMOthers() As Integer
Dim sqlstr As String
Dim MBM As New ADODB.Recordset
Dim strCount As String
Dim rstCount As New ADODB.Recordset
Dim MBMNumber As String
Dim Search As Boolean
    
    Call ClearForm(Me)
    
    If Source <> 2 Then
        Call ComboListing
    End If
    
    Label4 = ""
    Search = False
    sqlstr = " Select MBMNumber, MBMTelnoH, MBMTelNoM, MBMPayorPrem, " & _
            " MBMTelNoO, MBMPrevPLNCode, MBMExclusion, MBMAllergic, " & _
            " LabelStatus, MBMPrevMemNo, MBMPrevPolNo, MBMPrevPLNCode, MBMRBAmt, " & _
            " MBMPAYRemarks, MBMPrevInsCom, MBMRemarks, MBMCoPayRB, MBMGenWP, MBMGenEndWP, " & _
            " MBMPreExWP, MBMPreExEndWP, MBMSpeIllWP, MBMSpeIllEndWP, MBMStaffDec, MBMValueDate, " & _
            " MBMCLMNonPayFr, MBMCLMNonPayTo, MBMNewPlanCode, MBMNewPolEffDate, MBMPOLSubNo1, MBMBTBG, " & _
            " MBMOccupation, MBMWorkNature, MBMWeight, MBMHeight, MBMEmail, MBMBTBG, LGNCode, MBMBlood, " & _
            " MBMAlcohol, MBMPrevTakeOverPLNCode,MBMPayorPremGross,MBMPayorPremNet,MBMPayorMCOGross,MBMPayorMCONet, " & _
            " MBMSmoking, MBMPolicyDisc, MBMRenewalDisc "
    
    If Source = 1 Then
    
        sqlstr = sqlstr & " From View_MBM_MBMOthers Where MBMNumber ='" & Trim(frmMBMAdjustment.txtMBMNumber) & "'"
        MBM.Open sqlstr, medixmasterconn, adOpenKeyset, adLockOptimistic
    
        strCount = "select count(*) as totalRecord From View_MBM_MBMOthers  Where MBMNumber = '" & Trim(frmMBMAdjustment.txtMBMNumber) & "'"
        strCount = strCount + sqlstr
        
        MBMNumber = frmMBMAdjustment.txtMBMNumber
        Set rstCount = medixmasterconn.Execute(strCount)
        Search = True
    ElseIf Source = 2 Then
    
        sqlstr = sqlstr & " From View_MBM_MBMOthers Where MBMNumber ='" & Trim(frmMBMEnquiry.txtMBMNumber) & "'"
        MBM.Open sqlstr, medixmasterconn, adOpenKeyset, adLockOptimistic
    
        strCount = "select count(*) as totalRecord From View_MBM_MBMOthers  Where MBMNumber = '" & Trim(frmMBMEnquiry.txtMBMNumber) & "'"
        strCount = strCount + sqlstr
    
        MBMNumber = frmMBMEnquiry.txtMBMNumber
        Set rstCount = medixmasterconn.Execute(strCount)
        Search = True
    ElseIf Source = 3 And Trim(frmMemberRegistration.txtMBMPrevMEMNo) <> "" Then
    
        sqlstr = sqlstr & " From View_MBM_MBMOthers Where MBMNumber ='" & Trim(frmMemberRegistration.txtMBMPrevMEMNo) & "'"
        MBM.Open sqlstr, medixmasterconn, adOpenKeyset, adLockOptimistic
    
        strCount = "select count(*) as totalRecord From View_MBM_MBMOthers  Where MBMNumber = '" & Trim(frmMemberRegistration.txtMBMPrevMEMNo) & "'"
        strCount = strCount + sqlstr
        MBMNumber = frmMemberRegistration.txtMBMPrevMEMNo
        Set rstCount = medixmasterconn.Execute(strCount)
        Search = True
    End If
    
    If Search = True Then
        If rstCount!totalrecord = 0 Then
            MsgBox "No record found!", vbOKOnly + vbCritical, DialogBoxTitle
        Else
        
            With MBM
                
                Label4 = Trim(MBMNumber)
                
                If Trim(!MBMTelnoH) <> "" Then
                    txtMBMTelNoH = Trim(!MBMTelnoH)
                End If
                
                If Trim(!MBMTelNoM) <> "" Then
                    txtMBMTelNoM = Trim(!MBMTelNoM)
                End If
                
                If Trim(!MBMTelNoO) <> "" Then
                    txtMBMTelNoO = Trim(!MBMTelNoO)
                End If
                
                If Trim(!MBMCLMNonPayFr) <> "" Then
                    TxtClmNonPayFr = Format(!MBMCLMNonPayFr, "dd/mm/yyyy")
                End If
                If Trim(!MBMCLMNonPayTo) <> "" Then
                    TxtClmNonPayTo = Format(!MBMCLMNonPayTo, "dd/mm/yyyy")
                End If
                
                If Trim(!LabelStatus) = "Y" Then
                    Check1.value = 1
                End If
                
                If Trim(!MBMPolSubNo1) <> "" Then
                    TxtPOLSubNo = Trim(!MBMPolSubNo1)
                End If
                
                'If Trim(!MBMExclusion) <> "" Then
                '    txtMBMExclusion = Trim(!MBMExclusion)
                'End If
                
                If Trim(!MBMAllergic) <> "" Then
                    txtMBMAllergic = Trim(!MBMAllergic)
                End If
                
                If Len(!MBMPrevPLNCode) Then
                    txtMBMPrevPlan = !MBMPrevPLNCode
                End If
                
                If Len(!MBMRBAmt) Then
                    txtMBMRBAmt = !MBMRBAmt
                End If
                
                If Len(!MBMPAYRemarks) Then
                    txtMBMPAYRemarks = !MBMPAYRemarks
                End If
                
                If Len(!MBMPrevInsCom) Then
                    txtMBMPrevINSCom = !MBMPrevInsCom
                End If
            
                'If Len(!MBMRemarks) Then
                '    txtRemarks = Trim(!MBMRemarks)
                'End If
    
                If Len(!MBMCoPayRB) Then
                    cboCOPayRB = !MBMCoPayRB
                End If
                
                If Len(!MBMGenWP) Then
                    TxtGenWP = !MBMGenWP
                End If
                
                If Len(!MBMPreExWP) Then
                    txtPREWP = !MBMPreExWP
                End If
                
                If Len(!MBMSpeIllWP) Then
                    TxtSpeWP = !MBMSpeIllWP
                End If
                
                If Len(!MBMStaffDec) Then
                    cboStaffDisc = !MBMStaffDec
                End If
                
                If Len(!MBMGenEndWP) Then
                    txtGENWPEnd = !MBMGenEndWP
                End If
                
                If Len(!MBMPreExEndWP) Then
                    txtPREWPEnd = !MBMPreExEndWP
                End If
                
                If Len(!MBMSpeIllEndWP) Then
                    txtSPEWPEnd = !MBMSpeIllEndWP
                End If
                    
                If Len(!MBMNewPlanCode) Then
                    txtPLNCode = !MBMNewPlanCode
                End If
                
                If Len(!MBMNewPolEffDate) Then
                    txtNewPolEffDate = Format(!MBMNewPolEffDate, "dd/mm/yyyy")
                End If
                
                If Len(!MBMBTBG) Then
                    cboBTBG = !MBMBTBG
                End If
                
                'If Len(!MBMPayorPrem) Then
                '    txtPayPrem = !MBMPayorPrem
                'End If
                
                If Len(!MBMOccupation) Then
                    cboMBMOccupation = !MBMOccupation
                End If
                
                If Len(!MBMWorkNature) Then
                    txtMBMWorkNature = !MBMWorkNature
                End If
                
                If Len(!MBMWeight) Then
                    txtMBMWeight = !MBMWeight
                End If
                
                If Len(!MBMHeight) Then
                    txtMBMHeight = !MBMHeight
                End If
                
                If Len(!MBMEmail) Then
                    txtMBMEmail = !MBMEmail
                End If
                
                If Len(!LGNCode) Then
                    cboLGNCode = !LGNCode
                End If
                
                If Len(!MBMBlood) Then
                    cboMBMBlood = !MBMBlood
                End If
                             
                If Len(!MBMSmoking) Then
                    cboMBMSmoking = !MBMSmoking
                End If
                
                If Len(!MBMAlcohol) Then
                    cboMBMAlcohol = !MBMAlcohol
                End If
                
                If Len(!MBMPayorPremGross) Then
                    txtPayPrem = !MBMPayorPremGross
                End If
                
                If Len(!MBMPayorPremNet) Then
                    txtPayPremNett = !MBMPayorPremNet
                End If
                
                If Len(!MBMPayorMCOGross) Then
                    txtPayMCOGross = !MBMPayorMCOGross
                End If
                
                If Len(!MBMPayorMCONet) Then
                    txtPayMCONett = !MBMPayorMCONet
                End If
                
                If Len(!MBMPolicyDisc) Then
                    txtPolicyDisc = !MBMPolicyDisc
                End If
                
                If Len(!MBMRenewalDisc) Then
                    txtRenewalDisc = !MBMRenewalDisc
                End If
            End With
                    
        End If
        MBM.Close
        Set MBM = Nothing
    End If
    Search = False
End Function


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




Private Sub TxtClmNonPayFr_LostFocus()
    If IsDate(TxtClmNonPayFr) Then
    Else
    If TxtClmNonPayFr <> "__/__/____" Then
        MsgBox "Invalid Date Format! ", vbCritical
        TxtClmNonPayFr.SetFocus
    End If
    End If
End Sub


Private Sub TxtClmNonPayTo_LostFocus()
    If IsDate(TxtClmNonPayTo) Then
    Else
    If TxtClmNonPayTo <> "__/__/____" Then
        MsgBox "Invalid Date Format! ", vbCritical
        TxtClmNonPayTo.SetFocus
    End If
    End If

End Sub

Private Sub TxtGenWP_KeyPress(KeyAscii As Integer)
    KeyAscii = KeyNumber(KeyAscii)
End Sub

Private Sub TxtGenWP_LostFocus()
    If TxtGenWP <> "" Then
        If TxtGenWP > 0 Then
            If txtMBMPayorEffDate <> "__/__/____" Then
                txtGENWPEnd = Format(DateAdd("d", Trim(TxtGenWP), txtMBMPayorEffDate), "dd/mm/yyyy")
            End If
        Else
            txtGENWPEnd.mask = ""
            txtGENWPEnd = ""
            txtGENWPEnd.mask = "##/##/####"
        End If
    Else
        txtGENWPEnd.mask = ""
        txtGENWPEnd = ""
        txtGENWPEnd.mask = "##/##/####"
    End If

End Sub

Private Sub txtMBMHeight_KeyPress(KeyAscii As Integer)
    KeyAscii = KeyNumber(KeyAscii)
End Sub

Private Sub txtMBMRBAmt_KeyPress(KeyAscii As Integer)
    KeyAscii = KeyNumber(KeyAscii)
End Sub

Private Sub txtMBMTelNoH_KeyPress(KeyAscii As Integer)
    KeyAscii = KeyNumber(KeyAscii)
End Sub

Private Sub txtMBMTelNoM_KeyPress(KeyAscii As Integer)
    KeyAscii = KeyNumber(KeyAscii)
End Sub

Private Sub txtMBMTelNoO_KeyPress(KeyAscii As Integer)
    KeyAscii = KeyNumber(KeyAscii)
End Sub

Private Sub txtMBMWeight_KeyPress(KeyAscii As Integer)
    KeyAscii = KeyNumber(KeyAscii)
End Sub



Private Sub txtNewPolEffDate_LostFocus()
    If IsDate(txtNewPolEffDate) Then
    Else
    If txtNewPolEffDate <> "__/__/____" Then
        MsgBox "Invalid Date Format! ", vbCritical
        txtNewPolEffDate.SetFocus
    End If
    End If

End Sub

Private Sub txtPayMCOGross_KeyPress(KeyAscii As Integer)
    KeyAscii = KeyNumber(KeyAscii)
End Sub

Private Sub txtPayMCONett_KeyPress(KeyAscii As Integer)
    KeyAscii = KeyNumber(KeyAscii)
End Sub

Private Sub txtPayPrem_KeyPress(KeyAscii As Integer)
    KeyAscii = KeyNumber(KeyAscii)
End Sub

Private Sub txtPayPremNett_KeyPress(KeyAscii As Integer)
    KeyAscii = KeyNumber(KeyAscii)
End Sub

Private Sub txtPolicyDisc_KeyPress(KeyAscii As Integer)
    KeyAscii = KeyNumber(KeyAscii)
End Sub

Private Sub txtPolicyDisc_LostFocus()
    If Source = 3 Then
        If txtPolicyDisc <> "" Then
            frmMemberRegistration.PolDics = txtPolicyDisc
        End If
        
        If txtRenewalDisc <> "" Then
            frmMemberRegistration.RenDisc = txtRenewalDisc
        End If
    End If
End Sub

Private Sub txtPREWP_KeyPress(KeyAscii As Integer)
    KeyAscii = KeyNumber(KeyAscii)
End Sub

Private Sub txtPREWP_LostFocus()
    If txtPREWP <> "" Then
        If txtPREWP > 0 Then
            If txtMBMPayorEffDate <> "__/__/____" Then
                txtPREWPEnd = Format(DateAdd("d", Trim(txtPREWP), txtMBMPayorEffDate), "dd/mm/yyyy")
            End If
        Else
            txtPREWPEnd.mask = ""
            txtPREWPEnd = ""
            txtPREWPEnd.mask = "##/##/####"
        End If
    Else
        txtPREWPEnd.mask = ""
        txtPREWPEnd = ""
        txtPREWPEnd.mask = "##/##/####"
    End If

End Sub

Private Sub txtRenewalDisc_KeyPress(KeyAscii As Integer)
    KeyAscii = KeyNumber(KeyAscii)
End Sub

Private Sub txtRenewalDisc_LostFocus()
    If Source = 3 Then
        If txtPolicyDisc <> "" Then
            frmMemberRegistration.PolDics = txtPolicyDisc
        End If
        
        If txtRenewalDisc <> "" Then
            frmMemberRegistration.RenDisc = txtRenewalDisc
        End If
    End If
End Sub

Private Sub TxtSpeWP_KeyPress(KeyAscii As Integer)
    KeyAscii = KeyNumber(KeyAscii)
End Sub

Private Sub TxtSpeWP_LostFocus()
    If TxtSpeWP <> "" Then
        If TxtSpeWP > 0 Then
            If txtMBMPayorEffDate <> "__/__/____" Then
                txtSPEWPEnd = Format(DateAdd("d", Trim(TxtSpeWP), txtMBMPayorEffDate), "dd/mm/yyyy")
            End If
        Else
            txtSPEWPEnd.mask = ""
            txtSPEWPEnd = ""
            txtSPEWPEnd.mask = "##/##/####"
        End If
    Else
        txtSPEWPEnd.mask = ""
        txtSPEWPEnd = ""
        txtSPEWPEnd.mask = "##/##/####"
    End If

End Sub


Private Sub ComboListing()
    Dim LGN As New ADODB.Recordset
    Dim OCP As New ADODB.Recordset
    Dim Work As New ADODB.Recordset
    
    cboLGNCode.Clear
    txtMBMWorkNature.Clear
    cboMBMOccupation.Clear
    cboMBMBlood.Clear
    cboMBMSmoking.Clear
    cboStaffDisc.Clear
    cboBTBG.Clear
    cboMBMAlcohol.Clear
    OCP.Open "select OCPCode, OCPDescription from OCP", medixmasterconn, adOpenKeyset, adLockOptimistic
    LGN.Open "select LGNCode, LGNDescription from LGN", medixmasterconnMaint, adOpenKeyset, adLockOptimistic
    Work.Open "Select  WorkID, WorkDescription from WorkNature ", medixmasterconnMaint, adOpenKeyset, adLockOptimistic
    
    Do Until LGN.EOF
        cboLGNCode.AddItem ChkStr(LGN!LGNCode) & " - " & LGN!LGNDescription
        LGN.MoveNext
    Loop

    Do Until Work.EOF
        txtMBMWorkNature.AddItem Work!WorkID & " - " & Work!WorkDescription
        Work.MoveNext
    Loop

    Do Until OCP.EOF
        cboMBMOccupation.AddItem ChkStr(OCP!OCPCode) & " - " & OCP!OCPDescription
        OCP.MoveNext
    Loop
    
    cboMBMBlood.AddItem "A"
    cboMBMBlood.AddItem "B"
    cboMBMBlood.AddItem "AB"
    cboMBMBlood.AddItem "O"
        
    cboMBMSmoking.AddItem "Y"
    cboMBMSmoking.AddItem "N"
    
  
    cboStaffDisc.AddItem "N - No"
    cboStaffDisc.AddItem "Y - Yes"
    
    cboMBMAlcohol.AddItem "Y"
    cboMBMAlcohol.AddItem "N"
    
    cboBTBG.AddItem "Y - Yes"
    cboBTBG.AddItem "N - No"
    
    OCP.Close
    Set OCP = Nothing
    LGN.Close
    Set LGN = Nothing
End Sub

