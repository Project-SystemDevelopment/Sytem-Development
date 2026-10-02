VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmDateSelection 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Date Selection"
   ClientHeight    =   3240
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7350
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3240
   ScaleWidth      =   7350
   StartUpPosition =   2  'CenterScreen
   Begin MSComCtl2.MonthView MTHViewfr 
      Height          =   2370
      Left            =   1080
      TabIndex        =   2
      Top             =   720
      Width           =   2700
      _ExtentX        =   4763
      _ExtentY        =   4180
      _Version        =   393216
      ForeColor       =   -2147483630
      BackColor       =   -2147483633
      Appearance      =   1
      MultiSelect     =   -1  'True
      StartOfWeek     =   24576001
      TitleForeColor  =   12582912
      CurrentDate     =   37390
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Index           =   0
      Left            =   5160
      TabIndex        =   17
      Top             =   3720
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Index           =   1
      Left            =   3720
      TabIndex        =   15
      Top             =   3240
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Index           =   2
      Left            =   3720
      TabIndex        =   14
      Top             =   3480
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Index           =   3
      Left            =   5160
      TabIndex        =   13
      Top             =   3240
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Index           =   4
      Left            =   5160
      TabIndex        =   12
      Top             =   3480
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Index           =   5
      Left            =   3720
      TabIndex        =   11
      Top             =   3720
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Index           =   6
      Left            =   3720
      TabIndex        =   10
      Top             =   3960
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.ComboBox cboPayor 
      Height          =   315
      Left            =   1080
      TabIndex        =   0
      Top             =   40
      Width           =   4575
   End
   Begin VB.CommandButton cmdGO 
      Caption         =   "&Go"
      Default         =   -1  'True
      Height          =   360
      Left            =   6000
      TabIndex        =   4
      Top             =   45
      Width           =   615
   End
   Begin VB.ComboBox cboFormsSelection 
      Height          =   315
      Left            =   1080
      TabIndex        =   1
      Top             =   300
      Width           =   4575
   End
   Begin MSComCtl2.MonthView MTHViewTo 
      Height          =   2370
      Left            =   4485
      TabIndex        =   3
      Top             =   720
      Width           =   2700
      _ExtentX        =   4763
      _ExtentY        =   4180
      _Version        =   393216
      ForeColor       =   -2147483630
      BackColor       =   -2147483633
      Appearance      =   1
      StartOfWeek     =   24576001
      TitleForeColor  =   12582912
      CurrentDate     =   37390
   End
   Begin VB.CommandButton cmdExit 
      Caption         =   "&Exit"
      Height          =   365
      Left            =   6600
      TabIndex        =   5
      Top             =   45
      Width           =   615
   End
   Begin MSMask.MaskEdBox txtEffDate1 
      Height          =   300
      Left            =   1080
      TabIndex        =   27
      Top             =   4440
      Visible         =   0   'False
      Width           =   1815
      _ExtentX        =   3201
      _ExtentY        =   529
      _Version        =   393216
      MaxLength       =   10
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Mask            =   "##/##/####"
      PromptChar      =   "_"
   End
   Begin MSMask.MaskEdBox txtEffDate2 
      Height          =   300
      Left            =   3840
      TabIndex        =   28
      Top             =   4440
      Visible         =   0   'False
      Width           =   1815
      _ExtentX        =   3201
      _ExtentY        =   529
      _Version        =   393216
      MaxLength       =   10
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Mask            =   "##/##/####"
      PromptChar      =   "_"
   End
   Begin MSMask.MaskEdBox txtExpDate1 
      Height          =   300
      Left            =   1080
      TabIndex        =   29
      Top             =   4680
      Visible         =   0   'False
      Width           =   1815
      _ExtentX        =   3201
      _ExtentY        =   529
      _Version        =   393216
      MaxLength       =   10
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Mask            =   "##/##/####"
      PromptChar      =   "_"
   End
   Begin MSMask.MaskEdBox txtExpDate2 
      Height          =   300
      Left            =   3840
      TabIndex        =   30
      Top             =   4680
      Visible         =   0   'False
      Width           =   1815
      _ExtentX        =   3201
      _ExtentY        =   529
      _Version        =   393216
      MaxLength       =   10
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Mask            =   "##/##/####"
      PromptChar      =   "_"
   End
   Begin VB.Label Label8 
      Caption         =   "To"
      Height          =   255
      Left            =   3240
      TabIndex        =   26
      Top             =   4680
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.Label Label7 
      Caption         =   "To"
      Height          =   255
      Left            =   3240
      TabIndex        =   25
      Top             =   4440
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.Label Label6 
      Caption         =   "ExpDate"
      Height          =   255
      Left            =   120
      TabIndex        =   24
      Top             =   4800
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.Label Label5 
      Caption         =   "EffDate"
      Height          =   255
      Left            =   120
      TabIndex        =   23
      Top             =   4440
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.Label Label1 
      Appearance      =   0  'Flat
      BackColor       =   &H0090D0D8&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Label1"
      ForeColor       =   &H80000008&
      Height          =   240
      Index           =   6
      Left            =   840
      TabIndex        =   22
      Top             =   4200
      Visible         =   0   'False
      Width           =   2640
   End
   Begin VB.Label Label1 
      Appearance      =   0  'Flat
      BackColor       =   &H0090D0D8&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Label1"
      ForeColor       =   &H80000008&
      Height          =   240
      Index           =   5
      Left            =   840
      TabIndex        =   21
      Top             =   3960
      Visible         =   0   'False
      Width           =   2640
   End
   Begin VB.Label Label1 
      Appearance      =   0  'Flat
      BackColor       =   &H0090D0D8&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Label1"
      ForeColor       =   &H80000008&
      Height          =   240
      Index           =   4
      Left            =   840
      TabIndex        =   20
      Top             =   3720
      Visible         =   0   'False
      Width           =   2640
   End
   Begin VB.Label Label1 
      Appearance      =   0  'Flat
      BackColor       =   &H0090D0D8&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Label1"
      ForeColor       =   &H80000008&
      Height          =   240
      Index           =   3
      Left            =   840
      TabIndex        =   19
      Top             =   3480
      Visible         =   0   'False
      Width           =   2640
   End
   Begin VB.Label Label1 
      Appearance      =   0  'Flat
      BackColor       =   &H0090D0D8&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Label1"
      ForeColor       =   &H80000008&
      Height          =   240
      Index           =   2
      Left            =   840
      TabIndex        =   18
      Top             =   3240
      Visible         =   0   'False
      Width           =   2640
   End
   Begin VB.Label Label1 
      Appearance      =   0  'Flat
      BackColor       =   &H0090D0D8&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Label1"
      ForeColor       =   &H80000008&
      Height          =   240
      Index           =   0
      Left            =   8520
      TabIndex        =   16
      Top             =   3120
      Visible         =   0   'False
      Width           =   2640
   End
   Begin VB.Label Label4 
      Caption         =   "To"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   4080
      TabIndex        =   9
      Top             =   720
      Width           =   375
   End
   Begin VB.Label Label3 
      Caption         =   "From"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   8
      Top             =   720
      Width           =   615
   End
   Begin VB.Label Label2 
      Caption         =   "Report"
      Height          =   255
      Left            =   120
      TabIndex        =   7
      Top             =   380
      Width           =   615
   End
   Begin VB.Label Label1 
      Caption         =   "Payor"
      Height          =   255
      Index           =   1
      Left            =   120
      TabIndex        =   6
      Top             =   80
      Width           =   615
   End
End
Attribute VB_Name = "frmDateSelection"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim DayCensus As Boolean
Dim total As String
Dim Subtot(1 To 7) As String
Dim sDay As String
Dim eDay As String
Dim SelDay As Integer
Dim cnn As ADODB.Connection
Dim rst As ADODB.Recordset
Dim SvDate As String
Dim isDirty As Boolean
Dim StrAppend1 As String
Dim MBMIns As New ADODB.Recordset
Dim MBMSuppIns As New ADODB.Recordset
Dim CLMIns As New ADODB.Recordset
Dim MBMPAY As New ADODB.Recordset
Dim MBMSuppPAY As New ADODB.Recordset
Dim CLMPAY As New ADODB.Recordset

Dim TotPrinc, TotSupp As Integer
Dim TotMBM, TotPREM As Double
Dim TotCases As Integer
Dim TotApp As Double
Dim BB As Integer
Dim PYCode As String
Dim PAYRec As New ADODB.Recordset
Dim GrandTPrinc, GrandTSupp, GrandTPrem, GrandTCases As Double
Dim GrandTApproved, GrandMBMTotal As Double
    

Private Sub cboFormsSelection_Click()
If Mid(cboFormsSelection, 1, 1) = 0 Or Mid(cboFormsSelection, 1, 1) = 1 Then
        DayCensus = False
        Label3.Visible = True
        Label4.Visible = True
        MTHViewfr.Visible = True
        MTHViewTo.Visible = True
        MTHViewfr.Enabled = True
        MTHViewTo.Enabled = True
        Label5.Visible = False
        Label6.Visible = False
        Label7.Visible = False
        Label8.Visible = False
        txtEffDate1.Visible = False
        txtEffDate2.Visible = False
        txtExpDate1.Visible = False
        txtExpDate2.Visible = False
    ElseIf Mid(cboFormsSelection, 1, 1) = 2 Then
        DayCensus = True
        Label3.Visible = True
        Label4.Visible = False
        MTHViewfr.Visible = True
        MTHViewTo.Visible = True
        MTHViewfr.Enabled = True
        MTHViewTo.Enabled = False
        Label5.Visible = False
        Label6.Visible = False
        Label7.Visible = False
        Label8.Visible = False
        txtEffDate1.Visible = False
        txtEffDate2.Visible = False
        txtExpDate1.Visible = False
        txtExpDate2.Visible = False
    ElseIf Mid(cboFormsSelection, 1, 1) = 3 Or Mid(cboFormsSelection, 1, 1) = 4 Then
        DayCensus = False
        Label3.Visible = False
        Label4.Visible = False
        MTHViewfr.Visible = False
        MTHViewTo.Visible = False
        Label5.Visible = True
        Label6.Visible = True
        Label7.Visible = True
        Label8.Visible = True
        txtEffDate1.Visible = True
        txtEffDate2.Visible = True
        txtExpDate1.Visible = True
        txtExpDate2.Visible = True
        Label5.Top = 660
        Label7.Top = 660
        Label6.Top = 920
        Label8.Top = 920
        txtEffDate1.Top = 580
        txtEffDate2.Top = 580
        txtExpDate1.Top = 840
        txtExpDate2.Top = 840
    End If
End Sub

Private Sub cboFormsSelection_KeyPress(KeyAscii As Integer)
Dim NewText As String
Dim ValidCount As Integer
Dim ValidValue As String
Dim i As Integer

    If KeyAscii >= 32 And KeyAscii <> 127 Then
    NewText = LCase(Left(cboFormsSelection.Text, cboFormsSelection.SelStart) + Chr(KeyAscii) + Mid(cboFormsSelection.Text, cboFormsSelection.SelStart + cboFormsSelection.SelLength + 1))
 
    ValidCount = 0
    ValidValue = ""
 
    For i = 0 To cboFormsSelection.ListCount - 1
 
        If NewText = LCase(Left(cboFormsSelection.List(i), Len(NewText))) Then
                ValidCount = ValidCount + 1
                ValidValue = cboFormsSelection.List(i)
        End If
    Next
        
        If ValidCount <= 1 Then KeyAscii = 0
            If ValidCount = 1 Then
                cboFormsSelection.Text = ValidValue
                cboFormsSelection.SelStart = 0
                cboFormsSelection.SelLength = Len(ValidValue)
            End If
        End If
End Sub

Private Sub cboPayor_KeyPress(KeyAscii As Integer)
Dim NewText As String
Dim ValidCount As Integer
Dim ValidValue As String
Dim i As Integer

    If KeyAscii >= 32 And KeyAscii <> 127 Then
    NewText = LCase(Left(cboPayor.Text, cboPayor.SelStart) + Chr(KeyAscii) + Mid(cboPayor.Text, cboPayor.SelStart + cboPayor.SelLength + 1))
 
    ValidCount = 0
    ValidValue = ""
 
    For i = 0 To cboPayor.ListCount - 1
 
        If NewText = LCase(Left(cboPayor.List(i), Len(NewText))) Then
                ValidCount = ValidCount + 1
                ValidValue = cboPayor.List(i)
        End If
    Next
        
        If ValidCount <= 1 Then KeyAscii = 0
            If ValidCount = 1 Then
                cboPayor.Text = ValidValue
                cboPayor.SelStart = 0
                cboPayor.SelLength = Len(ValidValue)
            End If
        End If
End Sub

Private Sub CmdExit_Click()
    Unload Me
End Sub

Private Sub cmdGo_Click()
Dim sql, sql2, sql3 As String
Dim strAppend As String
Dim CLM As New ADODB.Recordset
Dim MBM As New ADODB.Recordset
Dim MBMSupp As New ADODB.Recordset

    If cboFormsSelection = "" Then
        MsgBox "Please Select a Report", vbOKOnly + vbCritical, DialogBoxTitle
        cboFormsSelection.SetFocus
    Else
        Screen.MousePointer = vbHourglass
        If Mid(cboFormsSelection, 1, 1) = 0 Then
            sql = " SELECT  COUNT(CASNumber) AS TotalCases, SUM(CLMTotalActual)AS totalActual,  " & _
              " SUM(CLMTotalApproved) AS TotalApproved, SUM(CASReserveAmount) AS TotalReserved," & _
              " SUM(CLMPayableByMedix2M) AS TotalMedix2M, SUM(CLMPayableByMedix2H) AS TotalMedix2H," & _
              " SUM(NoofDay) AS NoofDay From VIEW_CLMReport "
            sql = sql & "WHERE CaseRegDate >= CONVERT(DATETIME,  '" & MTHViewfr.value & "', 103) "
            sql = sql & "AND CaseRegDate <= CONVERT(DATETIME,  '" & MTHViewTo.value & "', 103) "
            If Len(cboPayor) Then
                sql = sql & "AND PAYCode = '" & Mid(cboPayor, 1, 2) & "'"
                frmCASDailyRpt.Label7 = cboPayor
            Else
                frmCASDailyRpt.Label7 = "ALL Payor"
            End If
            
            'rst.Open Sql, medixmasterconnWS, adOpenKeyset, adLockOptimistic

            CLM.Open sql, medixmasterconnWS, adOpenKeyset, adLockOptimistic
                      
            frmCASDailyRpt.lblTotalCases = CLM!TotalCases
            frmCASDailyRpt.lblTotalActAmt = "RM" & Format(CLM!TotalActual, "#,##0.00")
            frmCASDailyRpt.lblTotalAppAmt = "RM" & Format(CLM!TotalApproved, "#,##0.00")
            frmCASDailyRpt.lblTotalResAmt = "RM" & Format(CLM!TotalReserved, "#,##0.00")
            frmCASDailyRpt.lblTotalMBMAmt = "RM" & Format(CLM!TotalMedix2M, "#,##0.00")
            frmCASDailyRpt.lblTotalHOSAmt = "RM" & Format(CLM!TotalMedix2H, "#,##0.00")
            frmCASDailyRpt.show
            Unload frmDateSelection
        
            CLM.Close
            Set CLM = Nothing
        ElseIf Mid(cboFormsSelection, 1, 1) = 1 Then
            sql = " SELECT COUNT(MBMNumber) AS TotalPrinc, SUM(MBMPremium) AS TotalPrem From VIEW_MBMStatistic "
            
            sql2 = "WHERE MBMValueDate >= CONVERT(DATETIME,  '" & MTHViewfr.value & "', 103) "
            sql2 = sql2 & "AND MBMValueDate <= CONVERT(DATETIME,  '" & MTHViewTo.value & "', 103) "
            If Len(cboPayor) Then
                sql2 = sql2 & "AND PAYCode = '" & Mid(cboPayor, 1, 2) & "'"
                frmMBMDailyRpt.Label7 = cboPayor
            Else
                frmMBMDailyRpt.Label7 = "ALL Payor"
            End If
            
            sql = sql + sql2
            MBM.Open sql, medixmasterconn, adOpenKeyset, adLockOptimistic
            
            sql = ""
            sql = " SELECT COUNT(MBMCNumber) as TotalSupp FROM VIEW_MBMSUPP_STATISTIC " & sql2
            MBMSupp.Open sql, medixmasterconn, adOpenKeyset, adLockOptimistic
            
            frmMBMDailyRpt.lblTotalPrin = MBM!TotalPrinc
            frmMBMDailyRpt.lblTotalSupp = MBMSupp!TotalSupp
            frmMBMDailyRpt.lblTotalPrem = "RM " & Format(MBM!TotalPrem, "#,##0.00")
            frmMBMDailyRpt.show
            Unload frmDateSelection
            
            MBM.Close
            Set MBM = Nothing
            MBMSupp.Close
            Set MBMSupp = Nothing
        ElseIf Mid(cboFormsSelection, 1, 1) = 2 Then
                frmDailyCensus.Label2 = Text1(0).Text + " - " + Text1(6).Text
            If Len(cboPayor) Then
                StrAppend1 = "AND PAYCode = '" & Mid(cboPayor, 1, 2) & "'"
                frmDailyCensus.Label7 = cboPayor
            Else
                StrAppend1 = ""
                frmDailyCensus.Label7 = "ALL Payor"
            End If
            
            frmDailyCensus.Fgrid.Col = 1
            strAppend = " SELECT COUNT(CaseRegDate) AS NoofCase FROM VIEW_Admission WHERE 0 = 0 " & _
                  " AND CaseRegDate = '"
            Call NoofCase(strAppend)
            
            frmDailyCensus.Fgrid.Col = 2
            strAppend = " SELECT COUNT(CasDischargeDate) AS NoofCase FROM VIEW_Admission WHERE 0 = 0 " & _
                  " AND CasDischargeDate = '"
             Call NoofCase(strAppend)
            
            frmDailyCensus.Fgrid.Col = 3
            strAppend = " SELECT COUNT(CasClaimability) AS NoofCase FROM VIEW_Admission WHERE CASClaimability = 'R' " & _
                  " AND CaseRegDate = '"
             Call NoofCase(strAppend)
            
            frmDailyCensus.Fgrid.Col = 5
            strAppend = " SELECT COUNT(CasClaimAdmin) AS NoofCase FROM VIEW_Admission WHERE CasClaimAdmin = 1 " & _
                  " AND CaseRegDate = '"
             Call NoofCase(strAppend)
            
            frmDailyCensus.Fgrid.Col = 6
            strAppend = " SELECT COUNT(CasClaimAdmin) AS NoofCase FROM VIEW_Admission WHERE CasClaimAdmin = 0 " & _
                  " AND CaseRegDate = '"
            Call NoofCase(strAppend)
            
            frmDailyCensus.show
            Unload frmDateSelection
        ElseIf Mid(cboFormsSelection, 1, 1) = 3 Then ' MBM & CLaims by Payor
            Call SearchMBMCasPayor
        ElseIf Mid(cboFormsSelection, 1, 1) = 4 Then ' MBM & CLaims by Ins Type
            Call SearchMBMCasIns
        End If
        Screen.MousePointer = vbDefault
    End If
    
End Sub

Private Sub Form_Load()
    Call ComboListing
    DayCensus = False
End Sub

Private Sub ComboListing()
Dim PAY As New ADODB.Recordset

    PAY.Open "Select PAYCode, PAYCompanyName from PAY", medixmasterconn, adOpenKeyset, adLockOptimistic

    Do Until PAY.EOF
        cboPayor.AddItem PAY!PAYCode & " - " & Trim(PAY!PAYCompanyName)
        PAY.MoveNext
    Loop

    cboFormsSelection.AddItem "0" & " - " & "Claim"
    cboFormsSelection.AddItem "1" & " - " & "MBM'Ship"
    cboFormsSelection.AddItem "2" & " - " & "Daily Census"
    cboFormsSelection.AddItem "3" & " - " & "MBM & Claim"
    cboFormsSelection.AddItem "4" & " - " & "MBM & Claim (Ins Type)"
    
    PAY.Close
    Set PAY = Nothing

End Sub

Private Sub MTHViewfr_DateClick(ByVal DateClicked As Date)
    If DayCensus = True Then
        SelectWeek CStr(DateClicked)
        Call DailyCensusInitial
    End If
End Sub
Private Sub NoofCase(strAppend As String)
Dim rst2 As New ADODB.Recordset
Dim Record As ListItem
Dim ListCount As Integer
Dim listloop As Integer
Dim sql As String
    
    total = 0
      
    ListCount = 7
    For listloop = 1 To ListCount
        sql = ""
        sql = strAppend & Format(Text1(listloop - 1), "mm/dd/yyyy") & "'" & StrAppend1
    
        rst2.Open sql, medixmasterconn, adOpenKeyset, adLockOptimistic
    
        With rst2
            frmDailyCensus.Fgrid.Row = listloop
            frmDailyCensus.Fgrid.ColAlignment(1) = 6
            Subtot(listloop) = 0
            If Len(!NoofCase) Then
                Subtot(listloop) = !NoofCase
                frmDailyCensus.Fgrid = !NoofCase
                total = total + !NoofCase
            End If
        End With
        rst2.Close
        Set rst2 = Nothing
    Next listloop
    
    frmDailyCensus.Fgrid.Row = 8
    frmDailyCensus.Fgrid.CellFontBold = True
    frmDailyCensus.Fgrid.CellFontSize = 10
    frmDailyCensus.Fgrid = CDbl(total)
End Sub

Private Sub SelectWeek(sDate As String)
Dim X As Integer

MTHViewfr.value = sDate
SelDay = MTHViewfr.DayOfWeek

If SelDay = 1 Then
  eDay = DateAdd("y", 6, MTHViewfr.value)
  MTHViewfr.SelStart = MTHViewfr.value
  MTHViewfr.SelEnd = eDay
  For X = 0 To 6
    frmDailyCensus.Label1(X).Caption = Format(MTHViewfr.value + X, "long date")
    Text1(X).Text = Format(MTHViewfr.value + X, "short date")
  Next X
Else
  sDay = MTHViewfr.value - (MTHViewfr.DayOfWeek - 1)
  MTHViewfr.value = sDay
  eDay = DateAdd("y", 6, MTHViewfr.value)
  MTHViewfr.SelStart = MTHViewfr.value
  MTHViewfr.SelEnd = eDay
  For X = 0 To 6
    frmDailyCensus.Label1(X).Caption = Format(MTHViewfr.value + X, "long date")
    Text1(X).Text = Format(MTHViewfr.value + X, "short date")
  Next X
End If

' were switching to a new date so save the previous date
SvDate = MTHViewfr.value

End Sub
Private Sub DailyCensusInitial()
Dim aa As Integer
    
    frmDailyCensus.Fgrid.Row = 0
    frmDailyCensus.Fgrid.Col = 0
    frmDailyCensus.Fgrid = "Date"
    frmDailyCensus.Fgrid.Col = 1
    frmDailyCensus.Fgrid = "No of Reg"
    frmDailyCensus.Fgrid.Col = 2
    frmDailyCensus.Fgrid = "No of Dis"
    frmDailyCensus.Fgrid.Col = 3
    frmDailyCensus.Fgrid = "Repudiated"
    frmDailyCensus.Fgrid.Col = 4
    frmDailyCensus.Fgrid = "Manual Reg"
    frmDailyCensus.Fgrid.Col = 5
    frmDailyCensus.Fgrid = "Files Pass Out"
    frmDailyCensus.Fgrid.Col = 6
    frmDailyCensus.Fgrid = "Files Not Pass"
   

    frmDailyCensus.Fgrid.Col = 0
    
    frmDailyCensus.Fgrid.ColAlignment(0) = 1
    frmDailyCensus.Fgrid.ColAlignment(1) = 6
    frmDailyCensus.Fgrid.ColAlignment(2) = 6
    frmDailyCensus.Fgrid.ColAlignment(3) = 6
    frmDailyCensus.Fgrid.ColAlignment(4) = 6
    
    For aa = 1 To 7
        frmDailyCensus.Fgrid.ColWidth(aa - 1) = 1180
        If aa = 1 Then
            frmDailyCensus.Fgrid.ColWidth(aa - 1) = 2100
        End If
        frmDailyCensus.Fgrid.Row = aa
        frmDailyCensus.Fgrid = frmDailyCensus.Label1(aa - 1).Caption
    Next aa
    frmDailyCensus.Fgrid.Row = 8
    frmDailyCensus.Fgrid.CellFontBold = True
    frmDailyCensus.Fgrid.CellFontSize = 10
    frmDailyCensus.Fgrid = "Total"
End Sub
Private Function SearchMBMCasPayor()
Dim sql1, sql2, sql3 As String
Dim Payorsql1 As String
Dim strAppend As String
Dim MBMsql, CLMSql, SuppSql As String

Dim PayCnt As Integer

'GrandTPrinc = 0
'GrandTSupp = 0
'GrandMBMTotal = 0
'GrandTPrem = 0
'GrandTApproved = 0


    strAppend = ""
    If IsDate(txtEffDate1) = True Then
        strAppend = strAppend & " WHERE MBMPayorEffDate >= '" & Format(txtEffDate1, "mm/dd/yyyy") & "'"
    End If
    If IsDate(txtEffDate2) = True Then
        strAppend = strAppend & " AND MBMPayorEffDate <= '" & Format(txtEffDate2, "mm/dd/yyyy") & "'"
    End If
    If IsDate(txtExpDate1) = True Then
        strAppend = strAppend & " AND MBMPayorEXPDate >= '" & Format(txtExpDate1, "mm/dd/yyyy") & "'"
    End If
    If IsDate(txtExpDate2) = True Then
        strAppend = strAppend & " AND MBMPayorEXPDate <= '" & Format(txtExpDate2, "mm/dd/yyyy") & "'"
    End If
    
    BB = 0
    If Len(cboPayor) Then
        strAppend = strAppend & "AND PAYCode = '" & Mid(cboPayor, 1, 2) & "'"
        FrmDailyMBMCASByPayor.Label2 = cboPayor
        
        sql1 = " SELECT COUNT(MBMNumber) AS TotalPrinc, SUM(MBMPremium) AS TotalPrem From VIEW_MBMInsStatistic "
        sql2 = " SELECT COUNT(MBMCNumber) as TotalSupp FROM VIEW_MBMSuppInsStatistic "
        sql3 = " SELECT  COUNT(CASNumber) AS TotalCases, SUM(CLMTotalApproved) AS TotalApproved " & _
               " From VIEW_CLMReport "
        
        PYCode = Mid(cboPayor, 1, 2)
        Payorsql1 = " WHERE PAYCode = '" & PYCode & "'"
        sql1 = sql1 + Payorsql1
        sql2 = sql2 + Payorsql1
        sql3 = sql3 + Payorsql1
        
        MBMPAY.Open sql1, medixmasterconn, adOpenKeyset, adLockOptimistic
        MBMSuppPAY.Open sql2, medixmasterconn, adOpenKeyset, adLockOptimistic
        CLMPAY.Open sql3, medixmasterconn, adOpenKeyset, adLockOptimistic
            
        FrmDailyMBMCASByPayor.Label2 = PYCode
        
        BB = BB + 1
        PayCnt = 1
        
        Call MBMCasPayor
        
        MBMPAY.Close
        Set MBMPAY = Nothing
        MBMSuppPAY.Close
        Set MBMSuppPAY = Nothing
        CLMPAY.Close
        Set CLMPAY = Nothing
    
    Else
    
        PAYRec.Open "PAY", medixmasterconn, adOpenKeyset, adLockOptimistic
       ' PAYRec.Open "SELECT PAYCODE FROM PAY WHERE (PAYCODE='AG')  OR (PAYCODE='AL')", medixmasterconn, adOpenKeyset, adLockOptimistic
        PayCnt = PAYRec.recordCount + 2
        FrmDailyMBMCASByPayor.PayorFgrid.Rows = PayCnt
        
        Do Until PAYRec.EOF
            sql1 = ""
            sql2 = ""
            sql3 = ""
            Payorsql1 = ""
            
            sql1 = " SELECT COUNT(MBMNumber) AS TotalPrinc, SUM(MBMPremium) AS TotalPrem From VIEW_MBMInsStatistic "
            sql2 = " SELECT COUNT(MBMCNumber)As TotalSupp FROM VIEW_MBMSuppInsStatistic "
            sql3 = " SELECT COUNT(CASNumber) AS TotalCases, SUM(CLMTotalApproved) AS TotalApproved " & _
                  " From VIEW_CLMReport "
            
            PYCode = PAYRec!PAYCode
            Payorsql1 = " And PAYCode = '" & PYCode & "'"
       
            sql1 = sql1 + strAppend + Payorsql1
            sql2 = sql2 + strAppend + Payorsql1
            sql3 = sql3 + strAppend + Payorsql1
            
            MBMPAY.Open sql1, medixmasterconn, adOpenKeyset, adLockOptimistic
            MBMSuppPAY.Open sql2, medixmasterconn, adOpenKeyset, adLockOptimistic
            CLMPAY.Open sql3, medixmasterconn, adOpenKeyset, adLockOptimistic
            
            FrmDailyMBMCASByPayor.Label2 = "ALL Payor"
            
            BB = BB + 1
            Call MBMCasPayor
            
            MBMPAY.Close
            Set MBMPAY = Nothing
            MBMSuppPAY.Close
            Set MBMSuppPAY = Nothing
            CLMPAY.Close
            Set CLMPAY = Nothing
            PAYRec.MoveNext
        Loop
        PAYRec.Close
        Set PAYRec = Nothing
    End If
    BB = BB + 1
    FrmDailyMBMCASByPayor.PayorFgrid.Row = BB
    
    'FrmDailyMBMCASByPayor.PayorFgrid.Row = PayCnt
    FrmDailyMBMCASByPayor.PayorFgrid.Col = 0
    FrmDailyMBMCASByPayor.PayorFgrid.CellFontBold = True
    FrmDailyMBMCASByPayor.PayorFgrid = "Total"
    
    FrmDailyMBMCASByPayor.PayorFgrid.Col = 1
    FrmDailyMBMCASByPayor.PayorFgrid = GrandTPrinc
    
    FrmDailyMBMCASByPayor.PayorFgrid.Col = 2
    FrmDailyMBMCASByPayor.PayorFgrid = GrandTSupp
    
    FrmDailyMBMCASByPayor.PayorFgrid.Col = 3
    FrmDailyMBMCASByPayor.PayorFgrid = GrandMBMTotal
    
    FrmDailyMBMCASByPayor.PayorFgrid.Col = 4
    FrmDailyMBMCASByPayor.PayorFgrid = Format(GrandTPrem, "#,##0.00")
    
    FrmDailyMBMCASByPayor.PayorFgrid.Col = 5
    FrmDailyMBMCASByPayor.PayorFgrid = GrandTCases
    
    FrmDailyMBMCASByPayor.PayorFgrid.Col = 6
    FrmDailyMBMCASByPayor.PayorFgrid = Format(GrandTApproved, "#,##0.00")
    
    FrmDailyMBMCASByPayor.show
    Unload frmDateSelection

End Function
Private Function SearchMBMCasIns()
Dim sql1, sql2, sql3 As String
Dim strAppend As String
Dim INSsql1, INSsql2, INSsql3, INSsql4 As String
Dim MBMsql, CLMSql, SuppSql As String

        
    strAppend = ""
    If IsDate(txtEffDate1) = True Then
        strAppend = strAppend & " WHERE MBMPayorEffDate >= '" & Format(txtEffDate1, "mm/dd/yyyy") & "'"
    End If
    If IsDate(txtEffDate2) = True Then
        strAppend = strAppend & " AND MBMPayorEffDate <= '" & Format(txtEffDate2, "mm/dd/yyyy") & "'"
    End If
    If IsDate(txtExpDate1) = True Then
        strAppend = strAppend & " AND MBMPayorEXPDate >= '" & Format(txtExpDate1, "mm/dd/yyyy") & "'"
    End If
    If IsDate(txtExpDate2) = True Then
        strAppend = strAppend & " AND MBMPayorEXPDate <= '" & Format(txtExpDate2, "mm/dd/yyyy") & "'"
    End If
    
    If Len(cboPayor) Then
        strAppend = strAppend & "AND PAYCode = '" & Mid(cboPayor, 1, 2) & "'"
        frmDailyMBMCASByIns.Label2 = cboPayor
    Else
        frmDailyMBMCASByIns.Label2 = "ALL Payor"
    End If
    
    INSsql1 = " AND INSCODE='I' "
    INSsql2 = " AND INSCODE='F' "
    INSsql3 = " AND INSCODE='G' "
    INSsql4 = " AND INSCODE='H' "
    
    sql1 = " SELECT COUNT(MBMNumber) AS TotalPrinc, SUM(MBMPremium) AS TotalPrem From VIEW_MBMInsStatistic "
    sql2 = " SELECT COUNT(MBMCNumber) as TotalSupp FROM VIEW_MBMSuppInsStatistic "
    sql3 = " SELECT  COUNT(CASNumber) AS TotalCases, SUM(CLMTotalApproved) AS TotalApproved " & _
           " From VIEW_CLMReport "

' --------- Individual type
    MBMsql = ""
    MBMsql = sql1 + strAppend + INSsql1
    MBMIns.Open MBMsql, medixmasterconn, adOpenKeyset, adLockOptimistic
    
    SuppSql = ""
    SuppSql = sql2 + strAppend + INSsql1
    MBMSuppIns.Open SuppSql, medixmasterconn, adOpenKeyset, adLockOptimistic
    
    CLMSql = ""
    CLMSql = sql3 + strAppend + INSsql1
    CLMIns.Open CLMSql, medixmasterconnWS, adOpenKeyset, adLockOptimistic

    frmDailyMBMCASByIns.InsFgrid.Row = 1
    Call MBMCasIns
    
    MBMIns.Close
    Set MBMIns = Nothing
    MBMSuppIns.Close
    Set MBMSuppIns = Nothing
    CLMIns.Close
    Set CLMIns = Nothing
    
' --------- Family type
    MBMsql = ""
    MBMsql = sql1 + strAppend + INSsql2
    MBMIns.Open MBMsql, medixmasterconn, adOpenKeyset, adLockOptimistic
    
    SuppSql = ""
    SuppSql = sql2 + strAppend + INSsql2
    MBMSuppIns.Open SuppSql, medixmasterconn, adOpenKeyset, adLockOptimistic
    
    CLMSql = ""
    CLMSql = sql3 + strAppend + INSsql2
    CLMIns.Open CLMSql, medixmasterconnWS, adOpenKeyset, adLockOptimistic
    
    frmDailyMBMCASByIns.InsFgrid.Row = 2
    Call MBMCasIns
    
    MBMIns.Close
    Set MBMIns = Nothing
    MBMSuppIns.Close
    Set MBMSuppIns = Nothing
    CLMIns.Close
    Set CLMIns = Nothing

' --------- Group Individual
    MBMsql = ""
    MBMsql = sql1 + strAppend + INSsql3
    MBMIns.Open MBMsql, medixmasterconn, adOpenKeyset, adLockOptimistic
    
    SuppSql = ""
    SuppSql = sql2 + strAppend + INSsql3
    MBMSuppIns.Open SuppSql, medixmasterconn, adOpenKeyset, adLockOptimistic
    
    CLMSql = ""
    CLMSql = sql3 + strAppend + INSsql3
    CLMIns.Open CLMSql, medixmasterconnWS, adOpenKeyset, adLockOptimistic

    frmDailyMBMCASByIns.InsFgrid.Row = 3
    Call MBMCasIns
    
    MBMIns.Close
    Set MBMIns = Nothing
    MBMSuppIns.Close
    Set MBMSuppIns = Nothing
    CLMIns.Close
    Set CLMIns = Nothing
        
' --------- Group Family type
    MBMsql = ""
    MBMsql = sql1 + strAppend + INSsql4
    MBMIns.Open MBMsql, medixmasterconn, adOpenKeyset, adLockOptimistic
    
    SuppSql = ""
    SuppSql = sql2 + strAppend + INSsql4
    MBMSuppIns.Open SuppSql, medixmasterconn, adOpenKeyset, adLockOptimistic
    
    CLMSql = ""
    CLMSql = sql3 + strAppend + INSsql4
    CLMIns.Open CLMSql, medixmasterconnWS, adOpenKeyset, adLockOptimistic
    
    frmDailyMBMCASByIns.InsFgrid.Row = 4
    Call MBMCasIns
    
    MBMIns.Close
    Set MBMIns = Nothing
    MBMSuppIns.Close
    Set MBMSuppIns = Nothing
    CLMIns.Close
    Set CLMIns = Nothing
    
'-------- Total ----------

    frmDailyMBMCASByIns.InsFgrid.Row = 5
    frmDailyMBMCASByIns.InsFgrid.Col = 1
    frmDailyMBMCASByIns.InsFgrid = TotPrinc
                        
    frmDailyMBMCASByIns.InsFgrid.Col = 2
    frmDailyMBMCASByIns.InsFgrid = TotSupp
    
    frmDailyMBMCASByIns.InsFgrid.Col = 3
    frmDailyMBMCASByIns.InsFgrid = TotMBM
    
    frmDailyMBMCASByIns.InsFgrid.Col = 4
    frmDailyMBMCASByIns.InsFgrid = Format(TotPREM, "#,##0.00")
    
    frmDailyMBMCASByIns.InsFgrid.Col = 5
    frmDailyMBMCASByIns.InsFgrid = TotCases
    
    frmDailyMBMCASByIns.InsFgrid.Col = 6
    frmDailyMBMCASByIns.InsFgrid = Format(TotApp, "#,##0.00")
     
    frmDailyMBMCASByIns.InsFgrid.Col = 7
    frmDailyMBMCASByIns.InsFgrid = Format(CDbl((TotApp / TotPREM) * 100), "###0.00")
    
    frmDailyMBMCASByIns.show
    Unload frmDateSelection
    
End Function

Private Sub MBMCasIns()
    frmDailyMBMCASByIns.InsFgrid.Col = 1
    frmDailyMBMCASByIns.InsFgrid = MBMIns!TotalPrinc
                        
    frmDailyMBMCASByIns.InsFgrid.Col = 2
    frmDailyMBMCASByIns.InsFgrid = MBMSuppIns!TotalSupp
    
    frmDailyMBMCASByIns.InsFgrid.Col = 3
    frmDailyMBMCASByIns.InsFgrid = MBMIns!TotalPrinc + MBMSuppIns!TotalSupp
    
    frmDailyMBMCASByIns.InsFgrid.Col = 4
'    frmDailyMBMCASByIns.InsFgrid = "RM" & Format(MBMIns!TotalPrem, "#,##0.00")
    frmDailyMBMCASByIns.InsFgrid = Format(MBMIns!TotalPrem, "#,##0.00")
    
    frmDailyMBMCASByIns.InsFgrid.Col = 5
    frmDailyMBMCASByIns.InsFgrid = CLMIns!TotalCases

    If IsNumeric(CLMIns!TotalApproved) Then
        frmDailyMBMCASByIns.InsFgrid.Col = 6
        frmDailyMBMCASByIns.InsFgrid = Format(CLMIns!TotalApproved, "#,##0.00")
    
        frmDailyMBMCASByIns.InsFgrid.Col = 7
        frmDailyMBMCASByIns.InsFgrid = Format(CDbl(((MBMIns!TotalPrem - CLMIns!TotalApproved) / MBMIns!TotalPrem) * 100), "###0.00")
        TotApp = TotApp + CLMIns!TotalApproved
    Else
        frmDailyMBMCASByIns.InsFgrid.Col = 6
        frmDailyMBMCASByIns.InsFgrid = 0
        
        frmDailyMBMCASByIns.InsFgrid.Col = 7
        frmDailyMBMCASByIns.InsFgrid = 0
        TotApp = TotApp + 0
    End If
    
    TotPrinc = TotPrinc + MBMIns!TotalPrinc
    TotSupp = TotSupp + MBMSuppIns!TotalSupp
    TotMBM = TotMBM + MBMIns!TotalPrinc + MBMSuppIns!TotalSupp
    TotPREM = TotPREM + MBMIns!TotalPrem
    TotCases = TotCases + CLMIns!TotalCases
    
End Sub

Private Function MBMCasPayor()
    FrmDailyMBMCASByPayor.PayorFgrid.Row = BB
    FrmDailyMBMCASByPayor.PayorFgrid.Col = 0
    FrmDailyMBMCASByPayor.PayorFgrid = PYCode

    FrmDailyMBMCASByPayor.PayorFgrid.Col = 1
    FrmDailyMBMCASByPayor.PayorFgrid = MBMPAY!TotalPrinc

    FrmDailyMBMCASByPayor.PayorFgrid.Col = 2
    FrmDailyMBMCASByPayor.PayorFgrid = MBMSuppPAY!TotalSupp

    FrmDailyMBMCASByPayor.PayorFgrid.Col = 3
    FrmDailyMBMCASByPayor.PayorFgrid = MBMPAY!TotalPrinc + MBMSuppPAY!TotalSupp
    
    FrmDailyMBMCASByPayor.PayorFgrid.Col = 4
    'FrmDailyMBMCASByPayor.PayorFgrid = MBMPAY!TotalPrem
    FrmDailyMBMCASByPayor.PayorFgrid = IIf(IsNumeric(MBMPAY!TotalPrem) = True, MBMPAY!TotalPrem, 0)
    
    FrmDailyMBMCASByPayor.PayorFgrid.Col = 5
    FrmDailyMBMCASByPayor.PayorFgrid = CLMPAY!TotalCases
        
    FrmDailyMBMCASByPayor.PayorFgrid.Col = 6
    'FrmDailyMBMCASByPayor.PayorFgrid = CLMPAY!TotalApproved
    FrmDailyMBMCASByPayor.PayorFgrid = IIf(IsNumeric(CLMPAY!TotalApproved) = True, CLMPAY!TotalApproved, 0)

    FrmDailyMBMCASByPayor.PayorFgrid.Col = 7
   FrmDailyMBMCASByPayor.PayorFgrid = Format(CDbl((MBMPAY!TotalPrem - IIf(IsNumeric(CLMPAY!TotalApproved) = True, CLMPAY!TotalApproved, 0)) / MBMPAY!TotalPrem * 100), "#,##0.00")
   'FrmDailyMBMCASByPayor.PayorFgrid = Format(CDbl((MBMPAY!TotalPrem - CLMPAY!TotalApproved) / MBMPAY!TotalPrem * 100), "#,##0.00")
       
    GrandTPrinc = GrandTPrinc + MBMPAY!TotalPrinc
    GrandTSupp = GrandTSupp + MBMSuppPAY!TotalSupp
    GrandMBMTotal = GrandMBMTotal + MBMPAY!TotalPrinc + MBMSuppPAY!TotalSupp
    GrandTPrem = GrandTPrem + MBMPAY!TotalPrem
    GrandTCases = GrandTCases + CLMPAY!TotalCases
    'GrandTApproved = GrandTApproved + CLMPAY!TotalApproved
    GrandTApproved = GrandTApproved + IIf(IsNumeric(CLMPAY!TotalApproved) = True, CLMPAY!TotalApproved, 0)
End Function
