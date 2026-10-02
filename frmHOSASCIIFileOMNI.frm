VERSION 5.00
Begin VB.Form frmHOSASCIIFileOMNI 
   Caption         =   "Form1"
   ClientHeight    =   3030
   ClientLeft      =   60
   ClientTop       =   510
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   3030
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.TextBox txtFilName 
      Height          =   495
      Left            =   2160
      TabIndex        =   2
      Text            =   "Text1"
      Top             =   240
      Width           =   1935
   End
   Begin VB.TextBox txtLocalFileLocation 
      Height          =   375
      Left            =   1080
      TabIndex        =   1
      Text            =   "Text1"
      Top             =   840
      Width           =   2055
   End
   Begin VB.CommandButton cmdGenerate 
      Caption         =   "Command1"
      Height          =   495
      Left            =   1440
      TabIndex        =   0
      Top             =   1920
      Width           =   1695
   End
End
Attribute VB_Name = "frmHOSASCIIFileOMNI"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit




Private Sub cmdGenerate_Click()
    Dim HOS As New ADODB.Recordset
    Dim PAY As New ADODB.Recordset
    Dim tmpcnt As Double
    Dim Trailer As String
    Dim tmpfilename As String
    Dim tmpLocalFileLocation As String
    Dim DataString As String
    
    tmpcnt = 0
    
    HOS.Open "select * from HOS", medixmasterconn, adOpenKeyset, adLockOptimistic
    'Do Until HOS.EOF
    '    CboHos.AddItem HOS!HOSCode & " - " & Trim(HOS!HosName)
    '    HOS.MoveNext
    'Loop
    

    'tmpLocalFileLocation = "C:\"
        
         
    If HOS.EOF Then
        MsgBox "No record found!", vbOKOnly + vbCritical, DialogBoxTitle
    Else
            txtLocalFileLocation = "C:\MedicaGenClaimASCII\"
            
            'txtFilName = Date & "_OUTPATIENTPROVIDER"
            'Open txtLocalFileLocation & txtFilName & ".txt" For Output As #1
            txtFilName = Format(Date, "yyyymmdd")
            txtFilName = txtFilName & "_OUTPATIENTPROVIDER" & ".txt"
            Open txtLocalFileLocation & "\" & txtFilName For Output As #1
            
            
        Do Until HOS.EOF
             
                tmpcnt = tmpcnt + 1
           ' DataString = ""
                DataString = "02" & "|"
                
                DataString = DataString & "N" & "|"
                           
                DataString = DataString & HOS!HOSCode & "|"
                
                DataString = DataString & HOS!HosName & "|"
                
                DataString = DataString & HOS!HOSTelNo & "|"
                
                DataString = DataString & HOS!HOSAdd1 & "|"
                
                DataString = DataString & HOS!HOSAdd2 & "|"
                
                DataString = DataString & HOS!HOSAdd3 & "|"
                
                DataString = DataString & HOS!HOSTown & "|"
                
                DataString = DataString & "" & "|"
                
                DataString = DataString & HOS!HOSState & "|"

                DataString = DataString & "MALAYSIA" & "|"
                
                DataString = DataString & "1" & "|"
                
                If HOS!HOSStatus = "P" Then
                    DataString = DataString & "N" & "|"
                ElseIf HOS!HOSStatus = "G" Then
                    DataString = DataString & "Y" & "|"
                Else
                    DataString = DataString & "N" & "|"
                End If
                
                DataString = DataString & "Y"
                

                
                Print #1, DataString
                
                DataString = DataString & Chr(13)
            HOS.MoveNext
            Loop

                Trailer = "3" & "|" & tmpcnt
                'Debug.Print Trailer
                Print #1, Trailer
                'Print #2, Trailer
                'cmdGenerate.Enabled = False
                Close #1
                
                MsgBox "File Generated Sucessfully...", vbOKOnly + vbInformation, DialogBoxTitle
                cmdGenerate.Enabled = True
        End If
    HOS.Close
    Set HOS = Nothing

End Sub
