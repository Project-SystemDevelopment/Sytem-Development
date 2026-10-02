VERSION 5.00
Begin VB.Form frmUpdatePolYear 
   Caption         =   "Form1"
   ClientHeight    =   1470
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4695
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1470
   ScaleWidth      =   4695
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdUpdate 
      Caption         =   "&Update"
      Height          =   375
      Left            =   3120
      TabIndex        =   5
      Top             =   1080
      Width           =   735
   End
   Begin VB.CommandButton cmdExit 
      Caption         =   "&Exit"
      Height          =   375
      Left            =   3840
      TabIndex        =   4
      Top             =   1080
      Width           =   735
   End
   Begin VB.Frame Frame1 
      Height          =   675
      Left            =   0
      TabIndex        =   1
      Top             =   240
      Width           =   4695
      Begin VB.ComboBox CboPayor 
         Height          =   315
         Left            =   840
         TabIndex        =   2
         Top             =   240
         Width           =   3615
      End
      Begin VB.Label Label1 
         Caption         =   "Payor"
         Height          =   255
         Left            =   240
         TabIndex        =   3
         Top             =   240
         Width           =   1215
      End
   End
   Begin VB.Label Label28 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00404040&
      Caption         =   "Update of Policy Year"
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
      Height          =   255
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   4725
   End
End
Attribute VB_Name = "frmUpdatePolYear"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdUpdate_Click()
Dim strsql As String
Dim MBM As New ADODB.Recordset
Dim ICNo As String
Dim strsqlUpd As String
Dim MBMUpd As New ADODB.Recordset
Dim PolCount As String
Dim RecCount As String
    medixmasterconn.Open "File Name=" & BOSPath & "Conn\mediconnect.UDL"
    medixmasterconnMaint.Open "File Name=" & BOSPath & "Conn\MediConnHISMaint.UDL"
    
    strsql = "Select PAYCode,MBMIcBcPp from MBMCrossReference where PAYCode = '" & Mid(CboPayor, 1, 2) & "'"
    MBM.Open strsql, medixmasterconnMaint, adOpenKeyset, adLockOptimistic
    
    'strsql = "Select PAYCode,MBMIcBcPp,MBMNumber,MBMPayorEffDate,MBMPolicyYear From MBM Where PAYCode = '" & Mid(CboPayor, 1, 2) & "'"
    'MBM.Open strsql, medixmasterconn, adOpenForwardOnly, adLockOptimistic
   
    If MBM.recordCount > 0 Then
        Do Until MBM.EOF
            If Trim(ICNo) <> Trim(MBM!MBMIcBcPp) Then
                ICNo = MBM!MBMIcBcPp
                strsqlUpd = " Select PAYCode,MBMIcBcPp,MBMNumber,MBMPayorEffDate,MBMPolicyYear From MBM Where MBMIcBcPp = '" & MBM!MBMIcBcPp & "' ORDER BY MBMPayorEffDate"
                MBMUpd.Open strsqlUpd, medixmasterconn, adOpenKeyset, adLockOptimistic
                    PolCount = 0
                    RecCount = 0
                    If MBMUpd.recordCount > 0 Then
                        'RecCount = RecCount + 1
                        
                        Do Until MBMUpd.EOF
                                PolCount = CDbl(PolCount) + 1
                                MBMUpd!MBMPolicyYear = PolCount
                            
                            MBMUpd.Update
                            MBMUpd.MoveNext
                        Loop

                    End If
                    MBMUpd.Close
                    Set MBMUpd = Nothing
            End If
                
        MBM.MoveNext
        Loop
    End If
    MBM.Close
    Set MBM = Nothing
    medixmasterconn.Close
    Set medixmasterconn = Nothing
    medixmasterconnMaint.Close
    Set medixmasterconnMaint = Nothing
End Sub

Private Sub Form_Load()
    medixmasterconn.Open "File Name=" & BOSPath & "Conn\mediconnect.UDL"
        Call ComboListing
    medixmasterconn.Close
    Set medixmasterconn = Nothing
End Sub

Private Sub ComboListing()
Dim Pay As New ADODB.Recordset
Dim cmd As New ADODB.Command

    Set cmd.ActiveConnection = medixmasterconn
    cmd.CommandText = "SearchPayor"
    cmd.CommandType = adCmdStoredProc
    cmd.CommandTimeout = 15
    Pay.CursorLocation = adUseClient
    Pay.CursorType = adOpenForwardOnly
    Pay.CacheSize = 100
    Set Pay = cmd.Execute
    
    Do Until Pay.EOF
        With Pay
            CboPayor.AddItem Pay!PAYCode & " - " & Trim(Pay!PAYCompanyName)
            Pay.MoveNext
        End With
    Loop
    Pay.Close
    Set Pay = Nothing
End Sub

