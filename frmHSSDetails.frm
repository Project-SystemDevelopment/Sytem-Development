VERSION 5.00
Begin VB.Form frmHSSDetails 
   Caption         =   "HSS Breakdowns Entry"
   ClientHeight    =   2385
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   5610
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2385
   ScaleWidth      =   5610
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      Height          =   1575
      Left            =   50
      TabIndex        =   11
      Top             =   240
      Width           =   5535
      Begin VB.TextBox txtRadiology 
         Alignment       =   1  'Right Justify
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
         Left            =   4080
         TabIndex        =   5
         Top             =   240
         Width           =   1335
      End
      Begin VB.TextBox txtInvPharmacy 
         Alignment       =   1  'Right Justify
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
         Left            =   1080
         TabIndex        =   0
         Top             =   240
         Width           =   1335
      End
      Begin VB.TextBox txtMedSupp 
         Alignment       =   1  'Right Justify
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
         Left            =   1080
         TabIndex        =   1
         Top             =   480
         Width           =   1335
      End
      Begin VB.TextBox txtOTSupp 
         Alignment       =   1  'Right Justify
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
         Left            =   4080
         TabIndex        =   6
         Top             =   480
         Width           =   1335
      End
      Begin VB.TextBox txtEquipChrg 
         Alignment       =   1  'Right Justify
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
         Left            =   4080
         TabIndex        =   7
         Top             =   720
         Width           =   1335
      End
      Begin VB.TextBox txtMisc 
         Alignment       =   1  'Right Justify
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
         Left            =   4080
         TabIndex        =   8
         Top             =   960
         Width           =   1335
      End
      Begin VB.TextBox txtNursingCare 
         Alignment       =   1  'Right Justify
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
         Left            =   1080
         TabIndex        =   2
         Top             =   720
         Width           =   1335
      End
      Begin VB.TextBox txtNursingServ 
         Alignment       =   1  'Right Justify
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
         Left            =   1080
         TabIndex        =   3
         Top             =   980
         Width           =   1335
      End
      Begin VB.TextBox txtLab 
         Alignment       =   1  'Right Justify
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
         Left            =   1080
         TabIndex        =   4
         Top             =   1230
         Width           =   1335
      End
      Begin VB.Label Label1 
         Caption         =   "Nursing Serv"
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
         Index           =   9
         Left            =   120
         TabIndex        =   23
         Top             =   1000
         Width           =   1575
      End
      Begin VB.Label Label1 
         Caption         =   "Misc Charges"
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
         Index           =   7
         Left            =   2640
         TabIndex        =   20
         Top             =   960
         Width           =   1095
      End
      Begin VB.Label Label1 
         Caption         =   "Equipment Charges"
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
         Index           =   6
         Left            =   2640
         TabIndex        =   19
         Top             =   720
         Width           =   1575
      End
      Begin VB.Label Label1 
         Caption         =   "OT Supplies"
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
         Index           =   5
         Left            =   2640
         TabIndex        =   18
         Top             =   480
         Width           =   1095
      End
      Begin VB.Label Label1 
         Caption         =   "Radiology"
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
         Index           =   4
         Left            =   2640
         TabIndex        =   17
         Top             =   240
         Width           =   735
      End
      Begin VB.Label Label1 
         Caption         =   "Lab"
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
         Index           =   3
         Left            =   120
         TabIndex        =   16
         Top             =   1250
         Width           =   1095
      End
      Begin VB.Label Label1 
         Caption         =   "Nursing Care"
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
         Index           =   2
         Left            =   120
         TabIndex        =   15
         Top             =   750
         Width           =   1095
      End
      Begin VB.Label Label1 
         Caption         =   "Med Supplies"
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
         Index           =   1
         Left            =   120
         TabIndex        =   14
         Top             =   500
         Width           =   1095
      End
      Begin VB.Label Label1 
         Caption         =   "Pharmacy"
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
         Index           =   0
         Left            =   120
         TabIndex        =   13
         Top             =   240
         Width           =   735
      End
   End
   Begin VB.Frame Frame2 
      Height          =   570
      Left            =   50
      TabIndex        =   12
      Top             =   1800
      Width           =   5535
      Begin VB.CommandButton cmdExit 
         Caption         =   "Exit"
         Height          =   255
         Left            =   4800
         TabIndex        =   9
         Top             =   240
         Width           =   615
      End
      Begin VB.TextBox txtHSS 
         Alignment       =   1  'Right Justify
         Enabled         =   0   'False
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
         Left            =   1080
         TabIndex        =   21
         Top             =   200
         Width           =   1335
      End
      Begin VB.Label Label1 
         Caption         =   "Total HSS"
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
         Index           =   8
         Left            =   120
         TabIndex        =   22
         Top             =   240
         Width           =   1095
      End
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00404040&
      Caption         =   "HSS Breakdowns Entry"
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
      TabIndex        =   10
      Top             =   0
      Width           =   5655
   End
End
Attribute VB_Name = "frmHSSDetails"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim tmpTotal As Double

Private Sub cmdExit_Click()
    frmHSSDetails.Visible = False
    frmInvoice.A3 = tmpTotal
End Sub

Private Sub txtEquipChrg_Change()
    Call Calculations
End Sub

Private Sub txtInvPharmacy_Change()
    Call Calculations
End Sub

Private Sub Calculations()
Dim a, b, c, d, e, f, g, H, i As Double

    a = IIf(IsNumeric(txtInvPharmacy), txtInvPharmacy, 0)
    b = IIf(IsNumeric(txtMedSupp), txtMedSupp, 0)
    c = IIf(IsNumeric(txtNursingCare), txtNursingCare, 0)
    d = IIf(IsNumeric(txtNursingServ), txtNursingServ, 0)
    e = IIf(IsNumeric(txtLab), txtLab, 0)
    f = IIf(IsNumeric(txtRadiology), txtRadiology, 0)
    g = IIf(IsNumeric(txtOTSupp), txtOTSupp, 0)
    H = IIf(IsNumeric(txtEquipChrg), txtEquipChrg, 0)
    i = IIf(IsNumeric(txtMisc), txtMisc, 0)
    
    tmpTotal = CDbl(a) + CDbl(b) + CDbl(c) + CDbl(d) + CDbl(e) + CDbl(f) + CDbl(g) + CDbl(H) + CDbl(i)
    txtHSS = tmpTotal
End Sub

Private Sub txtLab_Change()
    Call Calculations
End Sub

Private Sub txtMedSupp_Change()
    Call Calculations
End Sub

Private Sub txtMisc_Change()
    Call Calculations
End Sub

Private Sub txtNursingCare_Change()
    Call Calculations
End Sub

Private Sub txtNursingServ_Change()
    Call Calculations
End Sub

Private Sub txtOTSupp_Change()
    Call Calculations
End Sub

Private Sub txtRadiology_Change()
    Call Calculations
End Sub
