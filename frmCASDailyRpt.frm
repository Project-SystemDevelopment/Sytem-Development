VERSION 5.00
Begin VB.Form frmCASDailyRpt 
   Caption         =   "Daily Claim Statistic"
   ClientHeight    =   3825
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   3525
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3825
   ScaleWidth      =   3525
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdOK 
      Caption         =   "Close"
      Height          =   375
      Left            =   1800
      TabIndex        =   1
      Top             =   3360
      Width           =   855
   End
   Begin VB.CommandButton cmdprint 
      Caption         =   "Print"
      Height          =   375
      Left            =   840
      TabIndex        =   0
      Top             =   3360
      Width           =   975
   End
   Begin VB.Label Label7 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      Caption         =   "Label2"
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
      Left            =   1290
      TabIndex        =   16
      Top             =   720
      Width           =   855
   End
   Begin VB.Label Label9 
      Caption         =   "Total Hosp Amt"
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
      TabIndex        =   15
      Top             =   2280
      Width           =   1815
   End
   Begin VB.Label lblTotalHOSAmt 
      Alignment       =   1  'Right Justify
      Caption         =   "Label12"
      Height          =   375
      Left            =   2040
      TabIndex        =   14
      Top             =   2280
      Width           =   1335
   End
   Begin VB.Label Label8 
      Caption         =   "Total Member Amt"
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
      TabIndex        =   13
      Top             =   2640
      Width           =   1815
   End
   Begin VB.Label lblTotalMBMAmt 
      Alignment       =   1  'Right Justify
      Caption         =   "Label12"
      Height          =   375
      Left            =   2040
      TabIndex        =   12
      Top             =   2640
      Width           =   1335
   End
   Begin VB.Label lblTotalResAmt 
      Alignment       =   1  'Right Justify
      Caption         =   "Label12"
      Height          =   375
      Left            =   2040
      TabIndex        =   11
      Top             =   3000
      Width           =   1335
   End
   Begin VB.Label Label6 
      Caption         =   "Total Reserved Amt"
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
      Top             =   3000
      Width           =   1815
   End
   Begin VB.Line Line1 
      X1              =   0
      X2              =   3960
      Y1              =   1080
      Y2              =   1080
   End
   Begin VB.Label lblTotalAppAmt 
      Alignment       =   1  'Right Justify
      Caption         =   "Label12"
      Height          =   375
      Left            =   2040
      TabIndex        =   9
      Top             =   1920
      Width           =   1335
   End
   Begin VB.Label lblTotalActAmt 
      Alignment       =   1  'Right Justify
      Caption         =   "Label11"
      Height          =   375
      Left            =   2040
      TabIndex        =   8
      Top             =   1560
      Width           =   1335
   End
   Begin VB.Label lblTotalCases 
      Alignment       =   1  'Right Justify
      Caption         =   "Label10"
      Height          =   375
      Left            =   2040
      TabIndex        =   7
      Top             =   1200
      Width           =   1335
   End
   Begin VB.Label Label3 
      Caption         =   "Total Approved Amt"
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
      Top             =   1920
      Width           =   1815
   End
   Begin VB.Label Label4 
      Caption         =   "Total Actual Amt"
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
      Top             =   1560
      Width           =   1935
   End
   Begin VB.Label Label5 
      Caption         =   "Total Claim"
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
      TabIndex        =   4
      Top             =   1200
      Width           =   1815
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
      Left            =   1200
      TabIndex        =   3
      Top             =   360
      Width           =   1035
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      Caption         =   "Claim Statistic"
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
      Left            =   840
      TabIndex        =   2
      Top             =   0
      Width           =   1815
   End
End
Attribute VB_Name = "frmCASDailyRpt"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public Source As Integer


Private Sub cmdOK_Click()
    Unload Me
    frmDateSelection.show
End Sub

Private Sub cmdprint_Click()
    frmCASDailyRpt.PrintForm
End Sub

Private Sub Form_Load()
    'If Source = 1 Then
    '    Label1.Caption = "Daily Claim Statistic"
        Label2.Caption = Format(Now, "long date")
    'Else
        'Label1.Caption = "Claim Statistic"
        'Label2.Caption = Format(Now, "mm/yyyy")
    'End If
    
    'Screen.MousePointer = vbHourglass
    '    Call getTotalCases
    'Screen.MousePointer = vbDefault
End Sub

Private Function getTotalCases() As Integer
Dim strsql As String
Dim CLM As New ADODB.Recordset
Dim RegDate  As Date
    lblTotalCases = 0
    lblTotalActAmt = 0
    lblTotalAppAmt = 0
    lblTotalResAmt = 0
    TotalMedix2M = 0
    TotalMedix2H = 0
    
    If Source = 1 Then
        sql = " SELECT  COUNT(CASNumber) AS TotalCases, SUM(CLMTotalActual)AS totalActual,  " & _
              " SUM(CLMTotalApproved) AS TotalApproved, SUM(CASReserveAmount) AS TotalReserved," & _
              " SUM(CLMPayableByMedix2M) AS TotalMedix2M, SUM(CLMPayableByMedix2H) AS TotalMedix2H," & _
              " SUM(NoofDay) AS NoofDay From VIEW_CLMReport Where CaseRegDate = '" & Format(Date, "mm/dd/yyyy") & "'"
    Else
        sql = " SELECT  COUNT(CASNumber) AS TotalCases, SUM(CLMTotalActual)AS totalActual,  " & _
          " SUM(CLMTotalApproved) AS TotalApproved, SUM(CASReserveAmount) AS TotalReserved," & _
          " SUM(CLMPayableByMedix2M) AS TotalMedix2M, SUM(CLMPayableByMedix2H) AS TotalMedix2H," & _
          " SUM(NoofDay) AS NoofDay From VIEW_CLMReport Where CaseRegDate = '" & Format(Now, "mm/yyyy") & "'"
    End If
    
    CLM.Open sql, medixmasterconnWS, adOpenKeyset, adLockOptimistic
    
    'If CLM.recordCount Then
        lblTotalCases = CLM!TotalCases
        lblTotalActAmt = "RM" & Format(CLM!TotalActual, "#,##0.00")
        lblTotalAppAmt = "RM" & Format(CLM!TotalApproved, "#,##0.00")
        lblTotalResAmt = "RM" & Format(CLM!TotalReserved, "#,##0.00")
        lblTotalMBMAmt = "RM" & Format(CLM!TotalMedix2M, "#,##0.00")
        lblTotalHOSAmt = "RM" & Format(CLM!TotalMedix2H, "#,##0.00")
    'End If
    CLM.Close
    Set CLM = Nothing
End Function

