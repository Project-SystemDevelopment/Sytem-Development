VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form frmMBMReintstatement 
   Caption         =   "Membership Reinstatement"
   ClientHeight    =   3180
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   7035
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   3180
   ScaleWidth      =   7035
   Begin VB.Frame Frame3 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
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
      Left            =   0
      TabIndex        =   10
      Top             =   0
      Width           =   7335
      Begin VB.Label Label3 
         Alignment       =   2  'Center
         BackColor       =   &H00000000&
         Caption         =   "Membership Reinstatement"
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
         Height          =   315
         Left            =   1200
         TabIndex        =   11
         Top             =   0
         Width           =   4335
      End
   End
   Begin VB.Frame Frame8 
      Height          =   615
      Left            =   60
      TabIndex        =   8
      Top             =   240
      Width           =   6915
      Begin VB.CommandButton cmdSearch 
         Caption         =   "Search"
         Default         =   -1  'True
         Height          =   255
         Left            =   3360
         TabIndex        =   1
         Top             =   240
         Width           =   735
      End
      Begin VB.TextBox txtMBMNumber 
         Height          =   285
         Left            =   1320
         TabIndex        =   0
         Top             =   240
         Width           =   1935
      End
      Begin VB.Label Label24 
         Caption         =   "Member No"
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
         Left            =   120
         TabIndex        =   9
         Top             =   240
         Width           =   1125
      End
   End
   Begin VB.Frame Frame2 
      Enabled         =   0   'False
      Height          =   1320
      Left            =   60
      TabIndex        =   13
      Top             =   800
      Width           =   6915
      Begin VB.TextBox txtMBMName 
         Height          =   285
         Left            =   1320
         TabIndex        =   7
         Top             =   180
         Width           =   5535
      End
      Begin VB.TextBox txtPolNo 
         Height          =   285
         Left            =   1320
         TabIndex        =   21
         Top             =   440
         Width           =   5535
      End
      Begin MSMask.MaskEdBox txtEffDate 
         Height          =   285
         Left            =   1320
         TabIndex        =   14
         Top             =   675
         Width           =   1065
         _ExtentX        =   1879
         _ExtentY        =   503
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
      Begin MSMask.MaskEdBox txtExpDate 
         Height          =   285
         Left            =   3240
         TabIndex        =   15
         Top             =   675
         Width           =   1065
         _ExtentX        =   1879
         _ExtentY        =   503
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
      Begin MSMask.MaskEdBox txtLapsedDate 
         Height          =   285
         Left            =   5760
         TabIndex        =   27
         Top             =   675
         Width           =   1065
         _ExtentX        =   1879
         _ExtentY        =   503
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
      Begin VB.TextBox txtMBMStatus 
         Height          =   285
         Left            =   3240
         TabIndex        =   20
         Top             =   915
         Width           =   1065
      End
      Begin MSMask.MaskEdBox txtFinalExpDate 
         Height          =   285
         Left            =   1320
         TabIndex        =   29
         Top             =   915
         Width           =   1065
         _ExtentX        =   1879
         _ExtentY        =   503
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
      Begin VB.Label Label4 
         Caption         =   "Final Exp Date"
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
         Index           =   1
         Left            =   120
         TabIndex        =   30
         Top             =   960
         Width           =   1035
      End
      Begin VB.Label Label26 
         Caption         =   "Cancellation Date"
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
         Index           =   1
         Left            =   4440
         TabIndex        =   28
         Top             =   720
         Width           =   1275
      End
      Begin VB.Label Label25 
         Caption         =   "Policy No"
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
         Index           =   1
         Left            =   120
         TabIndex        =   22
         Top             =   470
         Width           =   1245
      End
      Begin VB.Label Label4 
         Caption         =   "Exp Date"
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
         Index           =   0
         Left            =   2520
         TabIndex        =   19
         Top             =   720
         Width           =   915
      End
      Begin VB.Label Label26 
         Caption         =   "Eff Date"
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
         Index           =   0
         Left            =   120
         TabIndex        =   18
         Top             =   720
         Width           =   1275
      End
      Begin VB.Label Label18 
         Caption         =   "Status"
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
         Index           =   0
         Left            =   2520
         TabIndex        =   17
         Top             =   960
         Width           =   555
      End
      Begin VB.Label Label25 
         Caption         =   "MBM Name"
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
         Index           =   0
         Left            =   120
         TabIndex        =   16
         Top             =   200
         Width           =   1245
      End
   End
   Begin VB.Frame Frame4 
      Height          =   525
      Left            =   60
      TabIndex        =   23
      Top             =   2040
      Width           =   6915
      Begin VB.TextBox txtBatchNo 
         Height          =   285
         Left            =   3240
         TabIndex        =   3
         Top             =   150
         Width           =   975
      End
      Begin MSMask.MaskEdBox txtendtbordxdate 
         Height          =   285
         Left            =   1320
         TabIndex        =   2
         Top             =   150
         Width           =   1065
         _ExtentX        =   1879
         _ExtentY        =   503
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
      Begin MSMask.MaskEdBox txtReinstateDate 
         Height          =   285
         Left            =   5760
         TabIndex        =   4
         Top             =   165
         Width           =   1065
         _ExtentX        =   1879
         _ExtentY        =   503
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
      Begin VB.Label Label18 
         Caption         =   "Batch No"
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
         Index           =   1
         Left            =   2520
         TabIndex        =   26
         Top             =   180
         Width           =   1035
      End
      Begin VB.Label Label4 
         Caption         =   "Reinstatement Date"
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
         Index           =   2
         Left            =   4320
         TabIndex        =   25
         Top             =   165
         Width           =   1455
      End
      Begin VB.Label Label26 
         Caption         =   "Endt Bordx Date"
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
         Index           =   2
         Left            =   120
         TabIndex        =   24
         Top             =   160
         Width           =   1275
      End
   End
   Begin VB.Frame Frame1 
      Height          =   615
      Left            =   60
      TabIndex        =   12
      Top             =   2520
      Width           =   6915
      Begin VB.CommandButton cmdReintstate 
         Caption         =   "&Reinstate"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   5160
         TabIndex        =   5
         Top             =   180
         Width           =   840
      End
      Begin VB.CommandButton CmdExit 
         Cancel          =   -1  'True
         Caption         =   "E&xit"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   6000
         TabIndex        =   6
         Top             =   180
         Width           =   825
      End
   End
End
Attribute VB_Name = "frmMBMReintstatement"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim tmpINSType As String
Dim tmpPLNCode As String
Dim tmpAdultChild As String
Dim tmpPayCode As String
Dim tmpHLTType As String
Dim tmpReinstateDate As String
Dim tmpLapsedDate As String
Dim tmpAGECode As String
Dim TempPremium, tempMCOFee As Double
Dim tmpSex As String
Dim tempOriMCO, tempOriPrem As Double
Dim tmpICNo As String

Private Sub CmdExit_Click()
    Unload Me
End Sub

Private Sub cmdReintstate_Click()
Dim tmpEndtBordxDate As String
Dim MBMCov As New ADODB.Recordset
Dim strsqlMBMCov As String
Dim strsql As String
Dim tmpExpDate As String

    If Trim(txtMBMStatus) <> "C" Then
        MsgBox "Membership is active! ", vbCritical
        Exit Sub
    Else
    '    medixmasterconn.Open "File Name=" & BOSPath & "Conn\mediconnect.UDL"
    '    medixmasterconnMaint.Open "File Name=" & BOSPath & "Conn\MediConnHISMaint.UDL"

        Screen.MousePointer = vbHourglass
            tmpEndtBordxDate = Format(txtendtbordxdate, "YYYY/MM/DD")
            tmpReinstateDate = Format(txtReinstateDate, "YYYY/MM/DD")
            tmpLapsedDate = Format(txtLapsedDate, "YYYY/MM/DD")
            tmpExpDate = Format(txtExpDate, "YYYY/MM/DD")
            Call MemberCallFees
            tempOriMCO = CDbl(tempOriMCO) + CDbl(tempMCOFee)
            tempOriMCO = Round(tempOriMCO)
            tempOriPrem = CDbl(tempOriPrem) + CDbl(TempPremium)
            strsql = ""
            'shan 12/06/2009
            strsql = " insert into MBMReinstatement (MBMNumber, MBMCoverID, MBMEndtBordxDate, MBMMCOFee, MBMPremium, " & _
                    " MBMLapsedDate, MBMReinstateDate, LastUpdateUSR, LastUpdateDate, MBMEndtBordxBatch)" & _
                    " Values('" & txtMBMNumber & "', '00', '" & Format(tmpEndtBordxDate, "yyyy/mm/dd") & "', " & tempMCOFee & ", " & CDbl(TempPremium) & ", " & _
                    " '" & Format(tmpLapsedDate, "yyyy/mm/dd") & "', '" & Format(tmpReinstateDate, "yyyy/mm/dd") & "', '" & ChkStr(Logon) & "', '" & Format(Date, "yyyy/mm/dd") & "', " & txtBatchNo & ")"
            medixmasterconn.Execute strsql
            
            strsql = ""
            strsql = "Update MBM set MBMStatus = 'A', MBMFinalExpiryDate = '" & tmpExpDate & "'  Where MBMNumber = '" & Trim(txtMBMNumber) & "'"
            medixmasterconn.Execute strsql
            
            strsql = ""
            strsql = "Update MBMCrossReference set MBMStatus = 'A' , MBMFinalExpiryDate = '" & tmpExpDate & "'  Where MBMNumber = '" & Trim(txtMBMNumber) & "'"
            medixmasterconnMaint.Execute strsql
            
            
            strsql = ""
            strsql = "Update MBMCoveredPersons set MBMStatus = 'A', MBMCExpDate  = '" & tmpExpDate & "'  Where MBMCNumber = '" & Trim(txtMBMNumber) & "'"
            medixmasterconn.Execute strsql
            
            'strsql = ""
            'strsql = "Update MBMOthers set MBMMCOFee = " & tempOriMCO & ", MBMPremium = " & tempOriPrem & "  Where MBMNumber = '" & Trim(txtMBMNumber) & "'"
            'medixmasterconn.Execute strsql
            
            
            If Trim(tmpINSType) = "F" Or Trim(tmpINSType) = "H" Then
                strsqlMBMCov = "SELECT MBMCNumber, MBMCCoverID, MBMCAdultChild,MBMCAgeband,MBMCSex,MBMCMCO,MBMCPremium From MBMCoveredPersons  Where MBMCNumber = '" & Trim(txtMBMNumber) & "'"
                MBMCov.Open strsqlMBMCov, medixmasterconn, adOpenKeyset, adLockOptimistic
                
                Do While Not MBMCov.EOF
                    strsql = ""
                    tmpAdultChild = MBMCov!MBMCAdultChild
                    tmpAGECode = MBMCov!MBMCAgeband
                    tmpSex = IIf(Len(MBMCov!MBMCSex), MBMCov!MBMCSex, "")
                    Call MemberCallFees
                    tempOriMCO = IIf(IsNumeric(MBMCov!MBMCMCO), MBMCov!MBMCMCO, 0)
                    tempOriPrem = IIf(IsNumeric(MBMCov!MBMCPremium), MBMCov!MBMCPremium, 0)
                    tempOriMCO = CDbl(tempOriMCO) + CDbl(tempMCOFee)
                    tempOriMCO = Round(tempOriMCO)
                    tempOriPrem = CDbl(tempOriPrem) + CDbl(TempPremium)
                    'shan 12/06/2009
                    strsql = " insert into MBMReinstatement (MBMNumber, MBMCoverID, MBMEndtBordxDate, MBMMCOFee, MBMPremium, " & _
                        " MBMLapsedDate, MBMReinstateDate, LastUpdateUSR, LastUpdateDate, MBMEndtBordxBatch)" & _
                        " Values('" & txtMBMNumber & "', '" & MBMCov!MBMCCoverID & "' , '" & tmpEndtBordxDate & "',  " & tempMCOFee & ", " & CDbl(TempPremium) & ", " & _
                        "'" & tmpLapsedDate & "', '" & tmpReinstateDate & "', '" & ChkStr(Logon) & "', '" & Format(Date, "yyyy/mm/dd") & "', " & txtBatchNo & " )"
                    medixmasterconn.Execute strsql
                    
                    strsql = ""
                    strsql = "Update MBMCoveredPersons set  MBMCStatus = 'A'  Where MBMCNumber = '" & Trim(txtMBMNumber) & "' And MBMCCoverID = '" & MBMCov!MBMCCoverID & "'"
                    medixmasterconn.Execute strsql
                MBMCov.MoveNext
                Loop
            End If
            
            strsql = ""
            strsql = " insert into MBMAmendHistory (MBMNumber, MBMName, MBMIcBcPp, MBMPlanCode, MBMPolicyNo, " & _
                    " MBMExpiryDate, MBMAmendUser, MBMAmendEdtType, MBMAmendDate, MBMAEndtBordxDate,MBMAEndtBatchNo,MBMAEndtEffDate)" & _
                    " Values('" & txtMBMNumber & "', '" & Replace(txtMBMName, "'", "''") & "', '" & tmpICNo & "', '" & tmpPLNCode & "', '" & txtPolNo & "', " & _
                    " '" & Format(txtExpDate, "yyyy/mm/dd") & "', '" & ChkStr(Logon) & "', 'R', '" & Format(Date, "yyyy/mm/dd") & "', '" & Format(txtendtbordxdate, "yyyy/mm/dd") & "',  " & txtBatchNo & ", '" & Format(txtReinstateDate, "yyyy/mm/dd") & "')"
            medixmasterconn.Execute strsql

      '      medixmasterconn.Close
     '       Set medixmasterconn = Nothing
       '     medixmasterconnMaint.Close
       '     Set medixmasterconnMaint = Nothing
            MsgBox "Membership Reinstated Successfully! ", vbCritical
            cmdSearch_Click
        Screen.MousePointer = Default
    End If
End Sub

Public Sub cmdSearch_Click()
Dim rst2 As New ADODB.Recordset
Dim Record As ListItem
Dim sqlstr As String
Dim strCriteria As String
Dim cmd As New ADODB.Command
Dim Prm1 As New ADODB.Parameter
Dim tmpMBMNumber As String
Dim strsqlTwo As String
Dim MBMTwo As New ADODB.Recordset

    strCriteria = ""
    If Trim(txtMBMNumber) = "" Then
        MsgBox "Please Enter Membership Number ", vbCritical
    Else
        Screen.MousePointer = vbHourglass
        tmpMBMNumber = txtMBMNumber
        Call ClearEntries
        txtMBMNumber = tmpMBMNumber
        
        'medixmasterconn.Open "File Name=" & BOSPath & "Conn\mediconnect.UDL"
        
        strCriteria = strCriteria & " AND MBMNumber =  '" & RemStr(txtMBMNumber) & "'"
        Set cmd.ActiveConnection = medixmasterconn
        cmd.CommandText = "Search_MBM"
        cmd.CommandType = adCmdStoredProc
        cmd.CommandTimeout = 300
        Set Prm1 = cmd.CreateParameter("strAppend", adVarChar, adParamInput, 2000, strCriteria)
        cmd.Parameters.Append Prm1
        rst2.CursorLocation = adUseClient
        rst2.CursorType = adOpenForwardOnly
        rst2.CacheSize = 100
        Set rst2 = cmd.Execute
            Do While Not rst2.EOF
                With rst2
                    txtMBMName = IIf(Len(!MBMName), Trim(!MBMName), "")
                    txtPolNo = IIf(Len(!MBMPolicyNo), Trim(!MBMPolicyNo), "")
                    txtMBMStatus = IIf(Len(!MBMStatus), Trim(!MBMStatus), "")
                    If Len(!MBMPayorEffDate) Then
                        txtEffDate = Format(!MBMPayorEffDate, "dd/mm/yyyy")
                    End If
                    If Len(!MBMPayorEXPDate) Then
                       txtExpDate = Format(!MBMPayorEXPDate, "dd/mm/yyyy")
                    End If
                    tmpINSType = IIf(Len(!INSCode), Trim(!INSCode), "")
                    tmpPLNCode = IIf(Len(!PLNCode), Trim(!PLNCode), "")
                    tmpAdultChild = IIf(Len(!MBMAdultChild), !MBMAdultChild, "")
                    tmpPayCode = IIf(Len(!PAYCode), !PAYCode, "")
                    tmpHLTType = IIf(Len(!HLTCode), !HLTCode, "")
                    tmpAGECode = IIf(Len(!AGECode), !AGECode, "")
                    tmpSex = IIf(Len(!MBMSex), !MBMSex, "")
                    tempOriMCO = IIf(IsNumeric(!MBMMCOFee), !MBMMCOFee, 0)
                    tempOriPrem = IIf(IsNumeric(!MBMPremium), !MBMPremium, 0)
                    tmpICNo = IIf(Len(!MBMIcBcPp), !MBMIcBcPp, "")
                    If Len(!MBMFinalExpiryDate) Then
                        txtFinalExpDate = Format(!MBMFinalExpiryDate, "dd/mm/yyyy")
                    End If
                End With
                rst2.MoveNext
            Loop
        rst2.Close
        Set rst2 = Nothing
        
        strsqlTwo = "SELECT MBMNumber, MBMCancelDate from MBMTwo where MBMNumber='" & Trim(txtMBMNumber) & "'"
        MBMTwo.Open strsqlTwo, medixmasterconn, adOpenKeyset, adLockOptimistic
        
            If MBMTwo.recordCount Then
                If Len(MBMTwo!MBMCancelDate) Then
                    txtLapsedDate = MBMTwo!MBMCancelDate
                End If
            End If
        MBMTwo.Close
        Set MBMTwo = Nothing
        
        'medixmasterconn.Close
       ' Set medixmasterconn = Nothing
        Screen.MousePointer = Default
    End If
End Sub

Private Sub ClearEntries()
   Dim ctl As Control
    For Each ctl In Me.Controls
        If TypeOf ctl Is TextBox Then
            ctl.Text = ""
        ElseIf TypeOf ctl Is ComboBox Then
        ElseIf TypeOf ctl Is ListView Then
            ctl.ListItems.Clear
        ElseIf TypeOf ctl Is MaskEdBox Then
            ctl.mask = ""
            ctl.Text = ""
            ctl.mask = "##/##/####"
        ElseIf TypeOf ctl Is Label Then
        End If
    Next ctl
End Sub
Public Sub MemberCallFees()
Dim PLNRec As New ADODB.Recordset
Dim SqlPln As String
Dim tmpPremSexStatus As String
Dim RoundUpStatus, PLNProRate As String

    If Trim(tmpLapsedDate) <> "__/__/____" _
        And Trim(tmpHLTType) <> "__/__/____" _
        And Trim(tmpPLNCode) <> "" And Trim(tmpINSType) <> "" Then

        SqlPln = "SELECT PLNRoundUpStatus, PLNProRate,PLNPremGenderStatus from PLN where PlnCode='" & Trim(tmpPLNCode) & _
        "' and HLTCode = '" & Trim(tmpHLTType) & _
        "' AND PLNEffDate <='" & Format(tmpLapsedDate, "mm/dd/yyyy") & "'"
        
        If Trim(tmpHLTType) = "N" Then
            SqlPln = SqlPln & " and PAYCode='" & Trim(tmpPayCode) & "'"
        End If

        PLNRec.Open SqlPln, medixmasterconn, adOpenKeyset, adLockOptimistic
        RoundUpStatus = "Y"
        PLNProRate = "Y"
        If PLNRec.recordCount Then
            RoundUpStatus = IIf(Len(Trim(PLNRec!PLNRoundUpStatus)), Trim(PLNRec!PLNRoundUpStatus), "Y")
            PLNProRate = IIf(Len(Trim(PLNRec!PLNProRate)), Trim(PLNRec!PLNProRate), "Y")
            tmpPremSexStatus = IIf(Len(Trim(PLNRec!PLNPremGenderStatus)), Trim(PLNRec!PLNPremGenderStatus), "N")
        End If
        PLNRec.Close
        Set PLNRec = Nothing
        
        If Trim(tmpAdultChild) <> "" Then
            tempMCOFee = Format(GetMCO(Trim(Mid(tmpINSType, 1, 1)), Trim(tmpHLTType), tmpAdultChild, CDate(tmpLapsedDate), tmpPLNCode), "#,##0.00")
            If PLNProRate <> "N" Then
                tempMCOFee = InsertMCOFees(CDate(txtExpDate), CDate(tmpLapsedDate), Format(tempMCOFee, "#,##0.00"))
            End If
            tempMCOFee = Format(tempMCOFee, "#,##0.00")
        End If

        If RoundUpStatus <> "N" Then
            tempMCOFee = Round(tempMCOFee)
        End If
        
        TempPremium = Format(GetPremium(Trim(Mid(tmpINSType, 1, 1)), Trim(Mid(tmpPayCode, 1, 2)), Trim(Mid(tmpAGECode, 1, 2)), Trim(tmpPLNCode), Trim(Mid(tmpHLTType, 1, 1)), CDate(tmpLapsedDate), Mid(tmpSex, 1, 1)), "#,##0.00")
        
        TempPremium = Format(InsertPremium(CDate(txtExpDate), CDate(tmpLapsedDate), Format(TempPremium, "#,##0.00")), "#,##0.00")
    End If
End Sub

Private Sub txtendtbordxdate_LostFocus()
    If IsDate(txtendtbordxdate) Then
    Else
        If txtendtbordxdate <> "__/__/____" Then
            MsgBox "Invalid Date Format! ", vbCritical
            txtendtbordxdate.SetFocus
        End If
    End If
End Sub

Private Sub txtReinstateDate_LostFocus()
    If IsDate(txtReinstateDate) Then
    Else
        If txtReinstateDate <> "__/__/____" Then
            MsgBox "Invalid Date Format! ", vbCritical
            txtReinstateDate.SetFocus
        End If
    End If
End Sub
