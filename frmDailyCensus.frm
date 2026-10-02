VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form frmDailyCensus 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Daily Census"
   ClientHeight    =   4230
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9330
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4230
   ScaleWidth      =   9330
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdOK 
      Cancel          =   -1  'True
      Caption         =   "&Close"
      Default         =   -1  'True
      Height          =   375
      Left            =   5040
      TabIndex        =   2
      Top             =   3840
      Width           =   975
   End
   Begin VB.CommandButton cmdprintFile 
      Caption         =   "Print to &File"
      Height          =   375
      Left            =   4080
      TabIndex        =   1
      Top             =   3840
      Width           =   975
   End
   Begin MSFlexGridLib.MSFlexGrid Fgrid 
      Height          =   2535
      Left            =   0
      TabIndex        =   0
      Top             =   1200
      Width           =   10935
      _ExtentX        =   19288
      _ExtentY        =   4471
      _Version        =   393216
      Rows            =   9
      Cols            =   7
      BackColorFixed  =   -2147483638
      AllowUserResizing=   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
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
      Height          =   255
      Left            =   2400
      TabIndex        =   12
      Top             =   360
      Width           =   4815
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      Caption         =   "Daily Census"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2400
      TabIndex        =   11
      Top             =   40
      Width           =   4575
   End
   Begin VB.Label Label1 
      Appearance      =   0  'Flat
      BackColor       =   &H0090D0D8&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Label1"
      ForeColor       =   &H80000008&
      Height          =   240
      Index           =   0
      Left            =   4920
      TabIndex        =   10
      Top             =   4680
      Visible         =   0   'False
      Width           =   2160
   End
   Begin VB.Label Label1 
      Appearance      =   0  'Flat
      BackColor       =   &H0090D0D8&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Label1"
      ForeColor       =   &H80000008&
      Height          =   240
      Index           =   1
      Left            =   4920
      TabIndex        =   9
      Top             =   4440
      Visible         =   0   'False
      Width           =   2160
   End
   Begin VB.Label Label1 
      Appearance      =   0  'Flat
      BackColor       =   &H0090D0D8&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Label1"
      ForeColor       =   &H80000008&
      Height          =   240
      Index           =   2
      Left            =   240
      TabIndex        =   8
      Top             =   4320
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
      Left            =   240
      TabIndex        =   7
      Top             =   4560
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
      Left            =   240
      TabIndex        =   6
      Top             =   4800
      Visible         =   0   'False
      Width           =   2760
   End
   Begin VB.Label Label1 
      Appearance      =   0  'Flat
      BackColor       =   &H0090D0D8&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Label1"
      ForeColor       =   &H80000008&
      Height          =   240
      Index           =   5
      Left            =   240
      TabIndex        =   5
      Top             =   5040
      Visible         =   0   'False
      Width           =   2760
   End
   Begin VB.Label Label1 
      Appearance      =   0  'Flat
      BackColor       =   &H0090D0D8&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Label1"
      ForeColor       =   &H80000008&
      Height          =   240
      Index           =   6
      Left            =   240
      TabIndex        =   4
      Top             =   5280
      Visible         =   0   'False
      Width           =   2760
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
      Left            =   2400
      TabIndex        =   3
      Top             =   720
      Width           =   4635
   End
End
Attribute VB_Name = "frmDailyCensus"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Subtot(1 To 7) As String
Dim total As Integer
Dim PrevAmt As String

Private Sub cmdOK_Click()
    Unload Me
    frmDateSelection.show
End Sub


Private Sub cmdprintFile_Click()
 Call ExportToExcelWorksheet
End Sub

Private Sub FGrid_Click()
Dim temp, tempCol, tempRow, tempTitle As String
Dim days, value As Double
    
    If Fgrid.Col = 4 And Fgrid.Row <> 8 Then
        If IsNumeric(Fgrid) = False Then
            PrevAmt = 0
        Else
            PrevAmt = Fgrid
        End If
        tempCol = Fgrid.Col
        tempRow = Fgrid.Row
        Fgrid.Row = 0
        tempTitle = Fgrid.Text
        Fgrid.Row = tempRow
        Fgrid.Col = 4
        temp = InputBox(tempTitle, Fgrid.Text)
        Fgrid.Col = tempCol
        
        If IsNumeric(temp) = False Then
            temp = 0
        End If
        If temp <> "" Then
            Fgrid = temp
            If tempCol = 4 Then
                Fgrid.Col = 4
                Subtot(tempRow) = Fgrid
            End If
            total = total - CDbl(PrevAmt) + CDbl(Subtot(tempRow))
           
            Fgrid.Row = 8
            Fgrid.CellFontBold = True
            Fgrid.CellFontSize = 10
            Fgrid = CDbl(total)
        End If
    End If
End Sub

Public Function ExportToExcelWorksheet()
Dim objExcel As Excel.Application
Dim objWorkbook As Excel.Workbook
Dim objWorksheet As Excel.Worksheet

    Screen.MousePointer = vbHourglass
    
    'Set objExcel = New excel.Application
    If objExcel Is Nothing Then
        Set objExcel = CreateObject("Excel.application")
    End If


    If ERR.Number > 0 Then
        MsgBox "Microsoft Excel is Not loaded On this machine.", vbOKOnly + vbCritical, "Error Loading Excel"
        GoTo ExitFunction
    End If
    
    On Error GoTo HANDLE_ERROR
    ' Don't allow user to affect workbook
    objExcel.Interactive = False


    If objExcel.Visible = False Then
        objExcel.Visible = True
    End If
    
    objExcel.WindowState = xlMaximized

    Set objWorkbook = objExcel.Workbooks.Add '(crdPath & "Template\Bordx\Return~1.XLT")
    Set objWorksheet = objWorkbook.Sheets(1)
    'Turn off the alerts, otherwise user will have to confirm my actions.
    objExcel.DisplayAlerts = False
    
    Set objWorksheet = objWorkbook.Worksheets.Item(1)
    objWorksheet.Cells(3, "A") = "Date"
    objWorksheet.Cells(3, "B") = "No of Registered"
    objWorksheet.Cells(3, "C") = "No of Discharged"
    objWorksheet.Cells(3, "D") = "No of Repudiated"
    objWorksheet.Cells(3, "E") = "No of Manual Register"
    objWorksheet.Cells(3, "F") = "No of FPO"
    objWorksheet.Cells(3, "G") = "No of FNPO"
    
    Fgrid.Col = 0
    Fgrid.Row = 1
    If Trim(Fgrid) <> "" Then
        objWorksheet.Cells(4, "A") = Fgrid
    End If
    
    Fgrid.Row = 2
    If Trim(Fgrid) <> "" Then
        objWorksheet.Cells(5, "A") = Fgrid
    End If
    Fgrid.Row = 3
    If Trim(Fgrid) <> "" Then
        objWorksheet.Cells(6, "A") = Fgrid
    End If
    Fgrid.Row = 4
    If Trim(Fgrid) <> "" Then
        objWorksheet.Cells(7, "A") = Fgrid
    End If
    Fgrid.Row = 5
    If Trim(Fgrid) <> "" Then
        objWorksheet.Cells(8, "A") = Fgrid
    End If
    
    Fgrid.Row = 6
    If Trim(Fgrid) <> "" Then
        objWorksheet.Cells(9, "A") = Fgrid
    End If
    Fgrid.Row = 7
    If Trim(Fgrid) <> "" Then
        objWorksheet.Cells(10, "A") = Fgrid
    End If
    Fgrid.Row = 8
    If Trim(Fgrid) <> "" Then
        objWorksheet.Cells(11, "A") = Fgrid
    End If
    
    Fgrid.Col = 1
    Fgrid.Row = 1
    If Trim(Fgrid) <> "" Then
        objWorksheet.Cells(4, "B") = Fgrid
    End If
    
    Fgrid.Row = 2
    If Trim(Fgrid) <> "" Then
        objWorksheet.Cells(5, "B") = Fgrid
    End If
    Fgrid.Row = 3
    If Trim(Fgrid) <> "" Then
        objWorksheet.Cells(6, "B") = Fgrid
    End If
    Fgrid.Row = 4
    If Trim(Fgrid) <> "" Then
        objWorksheet.Cells(7, "B") = Fgrid
    End If
    Fgrid.Row = 5
    If Trim(Fgrid) <> "" Then
        objWorksheet.Cells(8, "B") = Fgrid
    End If
    
    Fgrid.Row = 6
    If Trim(Fgrid) <> "" Then
        objWorksheet.Cells(9, "B") = Fgrid
    End If
    Fgrid.Row = 7
    If Trim(Fgrid) <> "" Then
        objWorksheet.Cells(10, "B") = Fgrid
    End If
    Fgrid.Row = 8
    If Trim(Fgrid) <> "" Then
        objWorksheet.Cells(11, "B") = Fgrid
    End If
    
    Fgrid.Col = 2
    Fgrid.Row = 1
    If Trim(Fgrid) <> "" Then
        objWorksheet.Cells(4, "C") = Fgrid
    End If
    
    Fgrid.Row = 2
    If Trim(Fgrid) <> "" Then
        objWorksheet.Cells(5, "C") = Fgrid
    End If
    Fgrid.Row = 3
    If Trim(Fgrid) <> "" Then
        objWorksheet.Cells(6, "C") = Fgrid
    End If
    Fgrid.Row = 4
    If Trim(Fgrid) <> "" Then
        objWorksheet.Cells(7, "C") = Fgrid
    End If
    Fgrid.Row = 5
    If Trim(Fgrid) <> "" Then
        objWorksheet.Cells(8, "C") = Fgrid
    End If
    
    Fgrid.Row = 6
    If Trim(Fgrid) <> "" Then
        objWorksheet.Cells(9, "C") = Fgrid
    End If
    Fgrid.Row = 7
    If Trim(Fgrid) <> "" Then
        objWorksheet.Cells(10, "C") = Fgrid
    End If
    Fgrid.Row = 8
    If Trim(Fgrid) <> "" Then
        objWorksheet.Cells(11, "C") = Fgrid
    End If
    
    Fgrid.Col = 3
    Fgrid.Row = 1
    If Trim(Fgrid) <> "" Then
        objWorksheet.Cells(4, "D") = Fgrid
    End If
    
    Fgrid.Row = 2
    If Trim(Fgrid) <> "" Then
        objWorksheet.Cells(5, "D") = Fgrid
    End If
    Fgrid.Row = 3
    If Trim(Fgrid) <> "" Then
        objWorksheet.Cells(6, "D") = Fgrid
    End If
    Fgrid.Row = 4
    If Trim(Fgrid) <> "" Then
        objWorksheet.Cells(7, "D") = Fgrid
    End If
    Fgrid.Row = 5
    If Trim(Fgrid) <> "" Then
        objWorksheet.Cells(8, "D") = Fgrid
    End If
    
    Fgrid.Row = 6
    If Trim(Fgrid) <> "" Then
        objWorksheet.Cells(9, "D") = Fgrid
    End If
    Fgrid.Row = 7
    If Trim(Fgrid) <> "" Then
        objWorksheet.Cells(10, "D") = Fgrid
    End If
    Fgrid.Row = 8
    If Trim(Fgrid) <> "" Then
        objWorksheet.Cells(11, "D") = Fgrid
    End If
    
    Fgrid.Col = 4
    Fgrid.Row = 1
    If Trim(Fgrid) <> "" Then
        objWorksheet.Cells(4, "E") = Fgrid
    End If
    
    Fgrid.Row = 2
    If Trim(Fgrid) <> "" Then
        objWorksheet.Cells(5, "E") = Fgrid
    End If
    Fgrid.Row = 3
    If Trim(Fgrid) <> "" Then
        objWorksheet.Cells(6, "E") = Fgrid
    End If
    Fgrid.Row = 4
    If Trim(Fgrid) <> "" Then
        objWorksheet.Cells(7, "E") = Fgrid
    End If
    Fgrid.Row = 5
    If Trim(Fgrid) <> "" Then
        objWorksheet.Cells(8, "E") = Fgrid
    End If
    
    Fgrid.Row = 6
    If Trim(Fgrid) <> "" Then
        objWorksheet.Cells(9, "E") = Fgrid
    End If
    Fgrid.Row = 7
    If Trim(Fgrid) <> "" Then
        objWorksheet.Cells(10, "E") = Fgrid
    End If
    Fgrid.Row = 8
    If Trim(Fgrid) <> "" Then
        objWorksheet.Cells(11, "E") = Fgrid
    End If
    
    Fgrid.Col = 5
    Fgrid.Row = 1
    If Trim(Fgrid) <> "" Then
        objWorksheet.Cells(4, "F") = Fgrid
    End If
    
    Fgrid.Row = 2
    If Trim(Fgrid) <> "" Then
        objWorksheet.Cells(5, "F") = Fgrid
    End If
    Fgrid.Row = 3
    If Trim(Fgrid) <> "" Then
        objWorksheet.Cells(6, "F") = Fgrid
    End If
    Fgrid.Row = 4
    If Trim(Fgrid) <> "" Then
        objWorksheet.Cells(7, "F") = Fgrid
    End If
    Fgrid.Row = 5
    If Trim(Fgrid) <> "" Then
        objWorksheet.Cells(8, "F") = Fgrid
    End If
    
    Fgrid.Row = 6
    If Trim(Fgrid) <> "" Then
        objWorksheet.Cells(9, "F") = Fgrid
    End If
    Fgrid.Row = 7
    If Trim(Fgrid) <> "" Then
        objWorksheet.Cells(10, "F") = Fgrid
    End If
    Fgrid.Row = 8
    If Trim(Fgrid) <> "" Then
        objWorksheet.Cells(11, "F") = Fgrid
    End If
    
    Fgrid.Col = 6
    Fgrid.Row = 1
    If Trim(Fgrid) <> "" Then
        objWorksheet.Cells(4, "G") = Fgrid
    End If
    
    Fgrid.Row = 2
    If Trim(Fgrid) <> "" Then
        objWorksheet.Cells(5, "G") = Fgrid
    End If
    Fgrid.Row = 3
    If Trim(Fgrid) <> "" Then
        objWorksheet.Cells(6, "G") = Fgrid
    End If
    Fgrid.Row = 4
    If Trim(Fgrid) <> "" Then
        objWorksheet.Cells(7, "G") = Fgrid
    End If
    Fgrid.Row = 5
    If Trim(Fgrid) <> "" Then
        objWorksheet.Cells(8, "G") = Fgrid
    End If
    
    Fgrid.Row = 6
    If Trim(Fgrid) <> "" Then
        objWorksheet.Cells(9, "G") = Fgrid
    End If
    Fgrid.Row = 7
    If Trim(Fgrid) <> "" Then
        objWorksheet.Cells(10, "G") = Fgrid
    End If
    Fgrid.Row = 8
    If Trim(Fgrid) <> "" Then
        objWorksheet.Cells(11, "G") = Fgrid
    End If
        
    Screen.MousePointer = vbDefault
    
    objExcel.Interactive = True
    objExcel.DisplayAlerts = True
    
    Exit Function
ExitFunction:
    objExcel.Interactive = True
    objExcel.DisplayAlerts = True
    Screen.MousePointer = vbDefault
    objExcel.Quit
    Exit Function
HANDLE_ERROR:
    MsgBox "Export To Excel failed. Encountered the following Error" & vbCrLf & vbCrLf & _
    ERR.Number & ": " & ERR.Description, vbOKOnly + vbCritical, "Error Exporting To Excel"
    Set objWorksheet = Nothing
    Set objWorkbook = Nothing
    GoTo ExitFunction
    
End Function

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If Shift = 0 Then
        Select Case KeyCode
            Case vbKeyF6
                KeyCode = 0
                PrintScreenSnapShot
        End Select
    End If
total = 0
End Sub
