VERSION 5.00
Begin VB.Form frmMMA13thSchDesc 
   Caption         =   "MMA and 13th Schedule Description"
   ClientHeight    =   3090
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   9945
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3090
   ScaleWidth      =   9945
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame3 
      Height          =   1665
      Left            =   80
      TabIndex        =   6
      Top             =   720
      Width           =   9855
      Begin VB.TextBox lblBody13th 
         BackColor       =   &H8000000F&
         Height          =   495
         Left            =   960
         ScrollBars      =   2  'Vertical
         TabIndex        =   28
         Top             =   1080
         Width           =   1095
      End
      Begin VB.TextBox lblBodyMMA 
         BackColor       =   &H8000000F&
         Height          =   495
         Left            =   960
         ScrollBars      =   2  'Vertical
         TabIndex        =   27
         Top             =   480
         Width           =   1095
      End
      Begin VB.TextBox lblMMADesc 
         BackColor       =   &H8000000F&
         Height          =   495
         Left            =   2040
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   25
         Top             =   480
         Width           =   4215
      End
      Begin VB.TextBox lbl13thSchDesc 
         BackColor       =   &H8000000F&
         Height          =   495
         Left            =   2040
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   26
         Top             =   1080
         Width           =   4215
      End
      Begin VB.Label lblPage13th 
         Alignment       =   1  'Right Justify
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Left            =   9120
         TabIndex        =   23
         Top             =   1080
         Width           =   615
      End
      Begin VB.Label lblPageMMA 
         Alignment       =   1  'Right Justify
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Left            =   9120
         TabIndex        =   9
         Top             =   480
         Width           =   615
      End
      Begin VB.Label lblAnae13th 
         Alignment       =   1  'Right Justify
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Left            =   8160
         TabIndex        =   22
         Top             =   1080
         Width           =   975
      End
      Begin VB.Label lblAnaeMMA 
         Alignment       =   1  'Right Justify
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Left            =   8160
         TabIndex        =   15
         Top             =   480
         Width           =   975
      End
      Begin VB.Label lblSurg13thMax 
         Alignment       =   1  'Right Justify
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Left            =   7200
         TabIndex        =   21
         Top             =   1080
         Width           =   975
      End
      Begin VB.Label lblSurgMMA 
         Alignment       =   1  'Right Justify
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Left            =   7200
         TabIndex        =   13
         Top             =   480
         Width           =   975
      End
      Begin VB.Label lblSurg13thMin 
         Alignment       =   1  'Right Justify
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Left            =   6240
         TabIndex        =   20
         Top             =   1080
         Width           =   975
      End
      Begin VB.Label lblSurgMMAMin 
         Alignment       =   1  'Right Justify
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Left            =   6240
         TabIndex        =   18
         Top             =   480
         Width           =   975
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         Caption         =   "Surg Fee (Min)"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   10
         Left            =   6240
         TabIndex        =   19
         Top             =   120
         Width           =   855
      End
      Begin VB.Label Label1 
         Caption         =   "MMA"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   9
         Left            =   120
         TabIndex        =   17
         Top             =   600
         Width           =   495
      End
      Begin VB.Label Label1 
         Caption         =   "Anae Fee"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   5
         Left            =   8280
         TabIndex        =   16
         Top             =   240
         Width           =   855
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         Caption         =   "Surg Fee (Max)"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   4
         Left            =   7200
         TabIndex        =   14
         Top             =   120
         Width           =   1095
      End
      Begin VB.Label Label1 
         Caption         =   "Body Part"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   2
         Left            =   1080
         TabIndex        =   10
         Top             =   240
         Width           =   855
      End
      Begin VB.Label Label1 
         Caption         =   "Description"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   0
         Left            =   2040
         TabIndex        =   8
         Top             =   240
         Width           =   1575
      End
      Begin VB.Label Label1 
         Caption         =   "Page"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   1
         Left            =   9240
         TabIndex        =   7
         Top             =   240
         Width           =   495
      End
      Begin VB.Label Label1 
         Caption         =   "13th Sch"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   3
         Left            =   120
         TabIndex        =   24
         Top             =   1200
         Width           =   1215
      End
   End
   Begin VB.Frame Frame2 
      Height          =   495
      Left            =   80
      TabIndex        =   3
      Top             =   240
      Width           =   9855
      Begin VB.CommandButton cmdNext 
         Caption         =   "Next"
         Height          =   255
         Left            =   3360
         TabIndex        =   12
         Top             =   160
         Width           =   615
      End
      Begin VB.CommandButton cmdPrev 
         Caption         =   "Prev"
         Height          =   255
         Left            =   2760
         TabIndex        =   11
         Top             =   160
         Width           =   615
      End
      Begin VB.Label lblMMACode 
         BorderStyle     =   1  'Fixed Single
         Caption         =   "No Code Available !"
         Height          =   255
         Left            =   960
         TabIndex        =   5
         Top             =   150
         Width           =   1575
      End
      Begin VB.Label Label2 
         Caption         =   "MMA Code"
         Height          =   255
         Left            =   120
         TabIndex        =   4
         Top             =   150
         Width           =   975
      End
   End
   Begin VB.Frame Frame1 
      Height          =   615
      Left            =   120
      TabIndex        =   1
      Top             =   2400
      Width           =   9855
      Begin VB.CommandButton cmdExit 
         Caption         =   "Exit"
         Height          =   345
         Left            =   9000
         TabIndex        =   2
         Top             =   180
         Width           =   735
      End
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      BackColor       =   &H00404040&
      Caption         =   "MMA and 13th Schedule Description"
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
      TabIndex        =   0
      Top             =   0
      Width           =   9975
   End
End
Attribute VB_Name = "frmMMA13thSchDesc"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdExit_Click()
    Unload Me
End Sub

Private Sub CmdNext_Click()
Dim strsqlMMA As String
Dim MMA As New ADODB.Recordset
Dim strOrderby As String
Dim MMAStatus As Boolean

    strOrderby = ""
    strsqlMMA = ""
    MMAStatus = False
    
   ' medixmasterconnMaint.Open "File Name=" & BOSPath & "Conn\MediConnHISMaint.UDL"
    strsqlMMA = "Select OPCSCode, OPCSNarrative, ProcedureDescription,ProcMainBPCode,BodyParts,ProcPageNo,Page,  " & _
                "SurgeonsFees, AnaesthetistsFees, Surgeon13thMin,Surgeon13th,Anesthetist13th  " & _
                "from VIEW_MMA_Schedule where  OPCSCode > '" & Trim(lblMMACode) & "'"
    strOrderby = "Order by OPCSCode"
    strsqlMMA = strsqlMMA + strOrderby
    MMA.Open strsqlMMA, medixmasterconnMaint, adOpenKeyset, adLockOptimistic
        If Not MMA.EOF Then
            lblMMACode = IIf(Len(MMA!OPCSCode), MMA!OPCSCode, "")
            lblMMADesc = IIf(Len(MMA!OPCSNarrative), MMA!OPCSNarrative, "")
            lbl13thSchDesc = IIf(Len(MMA!ProcedureDescription), MMA!ProcedureDescription, "")
            lblBodyMMA = IIf(Len(MMA!ProcMainBPCode), MMA!ProcMainBPCode, "")
            lblPageMMA = IIf(Len(MMA!ProcPageNo), MMA!ProcPageNo, "")
            lblBody13th = IIf(Len(MMA!BodyParts), MMA!BodyParts, "")
            lblPage13th = IIf(Len(MMA!Page), MMA!Page, "")
            lblSurgMMA = IIf(Len(MMA!SurgeonsFees), MMA!SurgeonsFees, 0)
            lblAnaeMMA = IIf(Len(MMA!AnaesthetistsFees), MMA!AnaesthetistsFees, 0)
            lblSurg13thMin = IIf(Len(MMA!Surgeon13thMin), MMA!Surgeon13thMin, 0)
            lblSurg13thMax = IIf(Len(MMA!Surgeon13th), MMA!Surgeon13th, 0)
            lblAnae13th = IIf(Len(MMA!Anesthetist13th), MMA!Anesthetist13th, 0)
        Else
            MsgBox "This is last record!", vbOKOnly + vbCritical, DialogBoxTitle
        End If
    MMA.Close
    Set MMA = Nothing
'    medixmasterconnMaint.Close
 '   Set medixmasterconnMaint = Nothing

End Sub

Private Sub cmdPrev_Click()
Dim strsqlMMA As String
Dim MMA As New ADODB.Recordset
Dim strOrderby As String

  '  medixmasterconnMaint.Open "File Name=" & BOSPath & "Conn\MediConnHISMaint.UDL"
    strOrderby = ""
    
    strsqlMMA = "Select OPCSCode, OPCSNarrative, ProcedureDescription,ProcMainBPCode,BodyParts,ProcPageNo,Page,  " & _
                "SurgeonsFees, AnaesthetistsFees, Surgeon13thMin,Surgeon13th,Anesthetist13th  " & _
                "from VIEW_MMA_Schedule where  OPCSCode < '" & Trim(lblMMACode) & "'"
    strOrderby = "Order by OPCSCode DESC"
    strsqlMMA = strsqlMMA + strOrderby
    MMA.Open strsqlMMA, medixmasterconnMaint, adOpenKeyset, adLockOptimistic
        If Not MMA.EOF Then
            lblMMACode = IIf(Len(MMA!OPCSCode), MMA!OPCSCode, "")
            lblMMADesc = IIf(Len(MMA!OPCSNarrative), MMA!OPCSNarrative, "")
            lbl13thSchDesc = IIf(Len(MMA!ProcedureDescription), MMA!ProcedureDescription, "")
            lblBodyMMA = IIf(Len(MMA!ProcMainBPCode), MMA!ProcMainBPCode, "")
            lblPageMMA = IIf(Len(MMA!ProcPageNo), MMA!ProcPageNo, "")
            lblBody13th = IIf(Len(MMA!BodyParts), MMA!BodyParts, "")
            lblPage13th = IIf(Len(MMA!Page), MMA!Page, "")
            lblSurgMMA = IIf(Len(MMA!SurgeonsFees), MMA!SurgeonsFees, 0)
            lblAnaeMMA = IIf(Len(MMA!AnaesthetistsFees), MMA!AnaesthetistsFees, 0)
            lblSurg13thMin = IIf(Len(MMA!Surgeon13thMin), MMA!Surgeon13thMin, 0)
            lblSurg13thMax = IIf(Len(MMA!Surgeon13th), MMA!Surgeon13th, 0)
            lblAnae13th = IIf(Len(MMA!Anesthetist13th), MMA!Anesthetist13th, 0)
        Else
            MsgBox "This is first record!", vbOKOnly + vbCritical, DialogBoxTitle
        End If
    MMA.Close
    Set MMA = Nothing
  '  medixmasterconnMaint.Close
   ' Set medixmasterconnMaint = Nothing
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If Shift = 0 Then
        Select Case KeyCode
            Case vbKeyF6
                KeyCode = 0
                PrintScreenSnapShot
        End Select
    End If

End Sub

