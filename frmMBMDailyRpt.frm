VERSION 5.00
Begin VB.Form frmMBMDailyRpt 
   Caption         =   "Membership Daily Statistic"
   ClientHeight    =   3675
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   3600
   ClipControls    =   0   'False
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3675
   ScaleWidth      =   3600
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdprint 
      Caption         =   "&Print"
      Height          =   375
      Left            =   840
      TabIndex        =   1
      Top             =   3000
      Width           =   975
   End
   Begin VB.CommandButton cmdOK 
      Cancel          =   -1  'True
      Caption         =   "&Close"
      Default         =   -1  'True
      Height          =   375
      Left            =   1800
      TabIndex        =   0
      Top             =   3000
      Width           =   975
   End
   Begin VB.Label Label7 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      Caption         =   "Label1"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   1200
      TabIndex        =   12
      Top             =   720
      Width           =   1035
   End
   Begin VB.Label lblTotalPrem 
      Alignment       =   1  'Right Justify
      Caption         =   "Label12"
      Height          =   375
      Left            =   1800
      TabIndex        =   11
      Top             =   2160
      Width           =   1575
   End
   Begin VB.Label Label6 
      Caption         =   "Total Premium"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   120
      TabIndex        =   10
      Top             =   2160
      Width           =   1815
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      Caption         =   "Mem'ship Daily Statistic"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   240
      TabIndex        =   9
      Top             =   0
      Width           =   2955
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      Caption         =   "Label1"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   1320
      TabIndex        =   8
      Top             =   360
      Width           =   1035
   End
   Begin VB.Label Label5 
      Caption         =   "Total Principal"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   120
      TabIndex        =   7
      Top             =   1440
      Width           =   1815
   End
   Begin VB.Label Label4 
      Caption         =   "Total Supplementary"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   120
      TabIndex        =   6
      Top             =   1800
      Width           =   1935
   End
   Begin VB.Label Label3 
      Caption         =   "Total MCO Fees"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   120
      TabIndex        =   5
      Top             =   2160
      Visible         =   0   'False
      Width           =   1815
   End
   Begin VB.Label lblTotalPrin 
      Alignment       =   1  'Right Justify
      Caption         =   "Label10"
      Height          =   375
      Left            =   2280
      TabIndex        =   4
      Top             =   1440
      Width           =   1095
   End
   Begin VB.Label lblTotalSupp 
      Alignment       =   1  'Right Justify
      Caption         =   "Label11"
      Height          =   375
      Left            =   2280
      TabIndex        =   3
      Top             =   1800
      Width           =   1095
   End
   Begin VB.Label lblTotalMCO 
      Alignment       =   1  'Right Justify
      Caption         =   "Label12"
      Height          =   375
      Left            =   2280
      TabIndex        =   2
      Top             =   2160
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.Line Line1 
      X1              =   0
      X2              =   3480
      Y1              =   1080
      Y2              =   1080
   End
End
Attribute VB_Name = "frmMBMDailyRpt"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Function getDailyMBM() As Integer

Dim MBMsql As String
Dim MBMCovsql As String
Dim MBMOthersql As String
Dim MBM As New ADODB.Recordset
Dim MBMCov As New ADODB.Recordset
Dim MBMOthers As New ADODB.Recordset
Dim TotPrin As Variant
Dim TotMCO As Variant
Dim TotPREM As Variant
Dim TotMBMCov As Variant
Dim total As Variant

MBMsql = "SELECT MBMNumber FROM MBM WHERE MBMValueDate = '" & Format(Now, "mm/dd/yyyy") & "'"
MBM.Open MBMsql, medixmasterconn, adOpenKeyset, adLockOptimistic

    If MBM.recordCount = 0 Then
        lblTotalPrin = 0
        lblTotalSupp = 0
        lblTotalMCO = Format(0, "#,##0.00")
        lblTotalPrem = Format(0, "#,##0.00")
    Else
        Do Until MBM.EOF
            TotPrin = MBM.recordCount
         With MBM
            
            MBMCovsql = "SELECT MBMCNUMBER FROM MBMCoveredPersons WHERE MBMCNUMBER='" & MBM!MBMNumber & "'"
            MBMCov.Open MBMCovsql, medixmasterconn, adOpenKeyset, adLockOptimistic
            total = MBMCov.recordCount
            TotMBMCov = TotMBMCov + total
            MBMCov.Close
            Set MBMCov = Nothing
            
            MBMOthersql = "SELECT MBMPremium, MBMMCOFee FROM MBMOthers WHERE MBMNUMBER='" & MBM!MBMNumber & "'"
            MBMOthers.Open MBMOthersql, medixmasterconn, adOpenKeyset, adLockOptimistic
            TotMCO = TotMCO + MBMOthers!MBMMCOFee
            TotPREM = TotPREM + MBMOthers!MBMPremium
            MBMOthers.Close
            Set MBMOthers = Nothing

        End With
        MBM.MoveNext
    Loop
    lblTotalPrin = TotPrin
    lblTotalSupp = TotMBMCov
    lblTotalMCO = Format(TotMCO, "#,##0.00")
    lblTotalPrem = Format(TotPREM, "#,##0.00")
    MBM.Close
    Set MBM = Nothing
End If
End Function

Private Sub cmdOK_Click()
    Unload Me
    frmDateSelection.show
End Sub

Private Sub cmdprint_Click()
    frmMBMDailyRpt.PrintForm
End Sub

Private Sub Form_Load()
    'Label1.Caption = "Hotel Statistics "
    Label2.Caption = Format(Now, "long date")
    
   ' Call getDailyMBM
    'lblGuestReserved = getGuestReserved
    'lblGuestCheckout = getGuestCheckout
    'lblInhouseGuests = getInhouseGuests
    'lbltotalreservations = gettotalreservation
    'lblOccupied = getoccupied
    'lblvacant = getvacant
End Sub

