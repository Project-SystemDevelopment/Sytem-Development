VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form frmMembershipExcelPluto 
   Caption         =   "Generate Membership Excel (Pluto)"
   ClientHeight    =   1665
   ClientLeft      =   60
   ClientTop       =   510
   ClientWidth     =   7530
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   1665
   ScaleWidth      =   7530
   WindowState     =   2  'Maximized
   Begin VB.TextBox txtFileName 
      Enabled         =   0   'False
      Height          =   525
      Left            =   150
      TabIndex        =   1
      Top             =   420
      Width           =   6615
   End
   Begin VB.CommandButton cmdFileSearch 
      Caption         =   ">"
      Height          =   375
      Left            =   6840
      TabIndex        =   0
      Top             =   480
      Width           =   375
   End
   Begin VB.Frame Frame1 
      Height          =   1335
      Left            =   120
      TabIndex        =   2
      Top             =   240
      Width           =   7335
      Begin VB.CommandButton cmdExit 
         Caption         =   "Exit"
         Height          =   375
         Left            =   5880
         TabIndex        =   4
         Top             =   840
         Width           =   735
      End
      Begin VB.CommandButton cmdReturnDate 
         Caption         =   "Upload"
         Height          =   375
         Left            =   5040
         TabIndex        =   3
         Top             =   840
         Width           =   855
      End
      Begin MSComDlg.CommonDialog CommonDialog1 
         Left            =   120
         Top             =   480
         _ExtentX        =   847
         _ExtentY        =   847
         _Version        =   393216
      End
   End
   Begin VB.Label Label28 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00404040&
      Caption         =   "Generate Membership Excel (Pluto)"
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
      TabIndex        =   5
      Top             =   0
      Width           =   8445
   End
End
Attribute VB_Name = "frmMembershipExcelPluto"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim xl As New excel.Application
Dim xlsheet As excel.WORKSHEET
Dim xlwbook As excel.Workbook
Dim objExcel As excel.Application
Dim objWorkbook As excel.Workbook
Dim objWorksheet As excel.WORKSHEET
Dim tmpPathName As String

Private Sub cmdExit_Click()
    Unload Me
End Sub

Private Sub cmdFileSearch_Click()
  CommonDialog1.CancelError = True
  On Error GoTo ErrHandler
  CommonDialog1.InitDir = tmpPathName
  CommonDialog1.Flags = cdlOFNHideReadOnly
  CommonDialog1.Filter = "All Files (*.*)|*.*|Excel Files (*.xsl)|*.xsl"
  'CommonDialog1.FilterIndex = 2
  CommonDialog1.ShowOpen
  txtFileName = tmpPathName & CommonDialog1.FileTitle
  Exit Sub
  
ErrHandler:
  'User pressed the Cancel button
  Exit Sub

End Sub



Private Sub cmdReturnDate_Click()
Dim BordxNo As String
Dim AppCode As String
Dim AppDate As String
Dim ReturnDate As String
Dim AppAmt As Double
Dim commitTrans As Boolean
Dim Val1 As Integer
Dim strsqlupdate As String
Dim Number As Integer
Dim Record As ListItem
Dim RecordSaved
Dim tempPath As String
Dim strsqlREC As String
Dim tmpDept As String

    Screen.MousePointer = vbHourglass

    strsqlupdate = ""
    commitTrans = False
    Val1 = 2
    
    tempPath = "\\PLUTO\HISSaturn\Conn\MediConnectPluto.UDL"

    medixmasterconnPluto.Open "File Name=" & tempPath
    
    If txtFileName <> "" Then
        RecordSaved = MsgBox("Do you want to Upload the records?", vbYesNo + vbExclamation)
        If RecordSaved = vbYes Then

            Set xlwbook = xl.Workbooks.Open(txtFileName)
            Set xlsheet = xlwbook.Sheets.Item(1)
        
            
                medixmasterconnPluto.beginTrans
                On Err GoTo CisErrorSave
                Number = 1
                'Replace(cboHOS, "'", "''")
                Do Until Val1 = 60000
                    If xlsheet.Cells(Val1, 1) <> "" Then
                        tmpDept = Replace(xlsheet.Cells(Val1, 12), "'", "''")
                        strsqlREC = ""
                        strsqlREC = "Insert Into MBMExcellUploadShelly(No, RELCode, MBMPolicyNo," & _
                             "MBMName, MBMAddress1, MBMAddress2, MBMAddress3, " & _
                             "CTYCode, STACode,MBMPostCode,MBMGroupCompany," & _
                             "MBMDepartment, BRNCode, MBMSex, MBMRenewFlag," & _
                             "MBMIcBcPp, MBMDOB, PLNCode,INSCode, " & _
                             "MBMPayorEffDate, MBMPayorExpDate,MBMExclusion,MBMCLNPlanCode, " & _
                             "MBMCLNInsType,MBMCLNEffDate,MBMCLNExpDate,MBMCLNExclusion,PAYCode,MBMBordxDate,MBMRemarks,Subsidiary,costcentre)" & _
                             "Values('" & xlsheet.Cells(Val1, 1) & "','" & xlsheet.Cells(Val1, 2) & "' ,'" & xlsheet.Cells(Val1, 3) & _
                             "','" & xlsheet.Cells(Val1, 4) & "','" & xlsheet.Cells(Val1, 5) & "','" & xlsheet.Cells(Val1, 6) & "','" & xlsheet.Cells(Val1, 7) & _
                             "','" & xlsheet.Cells(Val1, 8) & "','" & xlsheet.Cells(Val1, 9) & "','" & xlsheet.Cells(Val1, 10) & "','" & xlsheet.Cells(Val1, 11) & _
                             "','" & tmpDept & "','" & xlsheet.Cells(Val1, 13) & "','" & xlsheet.Cells(Val1, 14) & "','" & xlsheet.Cells(Val1, 15) & _
                             "','" & xlsheet.Cells(Val1, 16) & "','" & Format(xlsheet.Cells(Val1, 17), "mm/dd/yyyy") & "','" & xlsheet.Cells(Val1, 18) & "','" & xlsheet.Cells(Val1, 19) & _
                             "','" & Format(xlsheet.Cells(Val1, 20), "mm/dd/yyyy") & "','" & Format(xlsheet.Cells(Val1, 21), "mm/dd/yyyy") & "','" & xlsheet.Cells(Val1, 22) & "','" & xlsheet.Cells(Val1, 23) & _
                             "','" & xlsheet.Cells(Val1, 24) & "','" & Format(xlsheet.Cells(Val1, 25), "mm/dd/yyyy") & "','" & Format(xlsheet.Cells(Val1, 26), "mm/dd/yyyy") & "','" & xlsheet.Cells(Val1, 27) & _
                             "','" & xlsheet.Cells(Val1, 28) & "','" & Format(xlsheet.Cells(Val1, 29), "mm/dd/yyyy") & "','" & xlsheet.Cells(Val1, 30) & "','" & xlsheet.Cells(Val1, 31) & "','" & xlsheet.Cells(Val1, 32) & "')"


                        medixmasterconnPluto.Execute strsqlREC
                        
                    Else
                        xl.DisplayAlerts = True
                        xl.Interactive = True
                        Set xlsheet = Nothing
                        Set xlwbook = Nothing
                        xl.Quit
                        Set xl = Nothing
                        commitTrans = True
                        medixmasterconnPluto.commitTrans
                        MsgBox "Records have been uploaded!", vbOKOnly + vbCritical, DialogBoxTitle
                        medixmasterconnPluto.Close
                        Set medixmasterconnPluto = Nothing
                        Screen.MousePointer = vbDefault
                        Exit Sub
                    End If
                    Val1 = Val1 + 1
                    Number = Number + 1
                Loop
        End If
    End If
  '  medixmasterconnWS.Close
  '  Set medixmasterconnWS = Nothing
  '  medixmasterconn.Close
  '  Set medixmasterconn = Nothing
    Screen.MousePointer = vbDefault

    Exit Sub
CisErrorSave:
    If commitTrans = False Then
        medixmasterconnWS.RollbackTrans
       ' medixmasterconnWS.Close
        'Set medixmasterconnWS = Nothing
       ' medixmasterconn.Close
       ' Set medixmasterconn = Nothing
    End If
    MsgBox Err.Description
    Screen.MousePointer = vbDefault
End Sub



Private Sub Form_Load()
    tmpPathName = "\\svr-doc\membership\"
End Sub

Private Sub UpdateSent(tBordxNo As String, SentDate As String)
Dim strsql As String
Dim strsqlREM As String
Dim rst As New ADODB.Recordset
Dim strsql2 As String
Dim rst2 As New ADODB.Recordset
Dim tmpRemarks1 As String
Dim tmpRemarks2 As String
Dim tmpRemarks3 As String
Dim tmpRemarks4 As String
Dim tmpRemarks5 As String
Dim strsqlWeb As String
Dim WEB As New ADODB.Recordset

    
    
    If Mid(tBordxNo, 1, 1) = "S" Then
        strsql = " update Bordx set " & _
                 " BordxMNRBSentDate = '" & Format(SentDate, "yyyy/mm/dd") & "', " & _
                 " BordxMNRBToStatus = '1', " & _
                 " BordxMNRBSentBy = '" & Logon & "' " & _
                 " Where BordxRefNo = '" & Trim(tBordxNo) & "'"
        
        medixmasterconnWS.Execute strsql
        
        strsql = " update Bordx set " & _
             " BordxPAYSentDate = '" & Format(SentDate, "yyyy/mm/dd") & "', " & _
             " BordxPayToStatus = '1', " & _
             " BordxPAYSentBy = '" & Logon & "' " & _
             " Where BordxRefNo = '" & Trim(tBordxNo) & "'"
    
        medixmasterconnWS.Execute strsql
        
    Else
        strsql = " update Bordx set " & _
             " BordxPAYSentDate = '" & Format(SentDate, "yyyy/mm/dd") & "', " & _
             " BordxPayToStatus = '1', " & _
             " BordxPAYSentBy = '" & Logon & "' " & _
             " Where BordxRefNo = '" & Trim(tBordxNo) & "'"
    
        medixmasterconnWS.Execute strsql
    End If
    
    strsql = " update Bordx set " & _
             " BordxSendACDate = '" & Format(SentDate, "yyyy/mm/dd") & "', " & _
             " BordxSendToACStatus = '1', " & _
             " BordxSendACBy = '" & Logon & "' " & _
             " Where BordxRefNo = '" & Trim(tBordxNo) & "'"
    medixmasterconnWS.Execute strsql
    
    strsqlREM = ""
    strsqlREM = "SELECT dbo.CASRemarks.CASNumber, " & _
                " HISWorksheet.dbo.Claim.BordxRefNo " & _
                " FROM dbo.CASRemarks  INNER JOIN " & _
                " HISWorksheet.dbo.Claim ON " & _
                " dbo.CasRemarks.CasNumber = HISWorksheet.dbo.Claim.CasNumber " & _
                " where HISWorksheet.dbo.Claim.BordxRefNo = '" & Trim(tBordxNo) & "'" & _
                " GROUP BY dbo.CASRemarks.CASNumber, " & _
                " HISWorksheet.dbo.Claim.BordxRefNo, " & _
                " HISWorksheet.dbo.Claim.CasNumber "
    rst.Open strsqlREM, medixmasterconn, adOpenKeyset, adLockOptimistic

    Do While Not rst.EOF
        strsqlWeb = ""
        strsqlWeb = " update CASRemarks set " & _
                " CASWebStatus = 'Y' " & _
                " Where CASNumber = '" & Trim(rst!CasNumber) & "'"
        medixmasterconn.Execute strsqlWeb
    rst.MoveNext
    Loop
    rst.Close
    Set rst = Nothing
    
    strsqlREM = ""
    strsqlREM = " SELECT HISWorksheet.dbo.Claim.BordxRefNo, " & _
                " dbo.CASRemarksTwo.CASNumber " & _
                " FROM HISWorksheet.dbo.Claim INNER JOIN " & _
                " dbo.CASRemarksTwo ON HISWorksheet.dbo.Claim.CasNumber = dbo.CASRemarksTwo.CasNumber " & _
                " where HISWorksheet.dbo.Claim.BordxRefNo = '" & Trim(tBordxNo) & "'" & _
                " GROUP BY HISWorksheet.dbo.Claim.BordxRefNo, HISWorksheet.dbo.Claim.CASNumber, " & _
                " dbo.CASRemarksTwo.CASNumber"
    rst.Open strsqlREM, medixmasterconn, adOpenKeyset, adLockOptimistic

    Do While Not rst.EOF
        strsqlWeb = ""
        strsqlWeb = " update CASRemarksTwo set " & _
                " CASWebStatus = 'Y' " & _
                " Where CASNumber = '" & Trim(rst!CasNumber) & "'"
        medixmasterconn.Execute strsqlWeb
    rst.MoveNext
    Loop
    rst.Close
    Set rst = Nothing
    
    strsqlREM = ""
    strsqlREM = " SELECT HISWorksheet.dbo.Claim.BordxRefNo, " & _
                " dbo.CASRemarkThree.CASNumber " & _
                " FROM HISWorksheet.dbo.Claim INNER JOIN dbo.CASRemarkThree ON " & _
                " HISWorksheet.dbo.Claim.CasNumber = dbo.CASRemarkThree.CasNumber " & _
                " where HISWorksheet.dbo.Claim.BordxRefNo = '" & Trim(tBordxNo) & "'" & _
                " GROUP BY HISWorksheet.dbo.Claim.BordxRefNo, dbo.CASRemarkThree.CASNumber "
    rst.Open strsqlREM, medixmasterconn, adOpenKeyset, adLockOptimistic

    Do While Not rst.EOF
        strsqlWeb = ""
        strsqlWeb = " update CASRemarkThree set " & _
                " CASWebStatus = 'Y' " & _
                " Where CASNumber = '" & Trim(rst!CasNumber) & "'"
        medixmasterconn.Execute strsqlWeb
    rst.MoveNext
    Loop
    rst.Close
    Set rst = Nothing
    
End Sub


Private Function UpdateReturnPayor(BordxNo As String, ReturnDate As String)
Dim strsql As String
Dim strsql2 As String
'Dim FRPayor As New ADODB.Recordset
Dim RecDocDate As String
Dim ii As Integer
Dim TempBodrxNo As String
Dim TempCASNumber As String
Dim TempCLMVersion As String
Dim CLM As New ADODB.Recordset
Dim strsqlRem1 As String
Dim strsqlRem2 As String
Dim strsqlRem3 As String
Dim REM1 As New ADODB.Recordset
Dim REM2 As New ADODB.Recordset
Dim REM3 As New ADODB.Recordset
Dim RemLen1 As Integer
Dim RemLen2 As Integer
Dim RemLen3 As Integer
Dim RemLen4 As Integer
Dim RemLen5 As Integer
Dim RemLen6 As Integer
Dim RemLen7 As Integer
Dim Remarks1 As String
Dim Remarks2 As String
Dim Remarks3 As String
Dim Remarks4 As String
Dim Remarks5 As String
Dim Remarks6 As String
Dim Remarks7 As String
Dim AddRemark As String
Dim strsqlREM As String

    ' update bordx table
    RecDocDate = Format(ReturnDate, "YYYY/MM/DD")
       
        
        strsql2 = ""
        strsql2 = " update Claim set " & _
                    " BordxPAYRecDocDate = '" & Format(RecDocDate, "yyyy/mm/dd") & "', " & _
                    " CLMAppStatus = 'A' " & _
                    " Where BordxRefNo = '" & Trim(BordxNo) & "'"
        medixmasterconnWS.Execute strsql2
        
                        
        medixmasterconn.Execute strsqlREM
            
    
    
End Function


