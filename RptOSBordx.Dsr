VERSION 5.00
Begin {78E93846-85FD-11D0-8487-00A0C90DC8A9} RptOSBordx 
   Caption         =   "Print Outstanding Bordx"
   ClientHeight    =   8595
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   11880
   MDIChild        =   -1  'True
   WindowState     =   2  'Maximized
   _ExtentX        =   20955
   _ExtentY        =   15161
   _Version        =   393216
   _DesignerVersion=   100685828
   ReportWidth     =   8784
   BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
      Name            =   "Arial"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   GridX           =   10
   GridY           =   10
   LeftMargin      =   1440
   RightMargin     =   1440
   TopMargin       =   1440
   BottomMargin    =   1440
   NumSections     =   5
   SectionCode0    =   1
   BeginProperty Section0 {1C13A8E0-A0B6-11D0-848E-00A0C90DC8A9} 
      _Version        =   393216
      Name            =   "Section4"
      Object.Height          =   720
      NumControls     =   2
      ItemType0       =   3
      BeginProperty Item0 {1C13A8E1-A0B6-11D0-848E-00A0C90DC8A9} 
         _Version        =   393216
         Name            =   "Label1"
         Object.Left            =   144
         Object.Top             =   144
         Object.Width           =   2880
         Object.Height          =   285
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Object.Caption         =   "OUTSTANDING BORDX"
      EndProperty
      ItemType1       =   5
      BeginProperty Item1 {1C13A8E3-A0B6-11D0-848E-00A0C90DC8A9} 
         _Version        =   393216
         Name            =   "Line1"
         Object.Left            =   144
         Object.Top             =   576
         Object.Width           =   8505
      EndProperty
   EndProperty
   SectionCode1    =   2
   BeginProperty Section1 {1C13A8E0-A0B6-11D0-848E-00A0C90DC8A9} 
      _Version        =   393216
      Name            =   "Section2"
      Object.Height          =   795
      NumControls     =   0
   EndProperty
   SectionCode2    =   4
   BeginProperty Section2 {1C13A8E0-A0B6-11D0-848E-00A0C90DC8A9} 
      _Version        =   393216
      Name            =   "Section1"
      Object.Height          =   1440
      NumControls     =   0
   EndProperty
   SectionCode3    =   7
   BeginProperty Section3 {1C13A8E0-A0B6-11D0-848E-00A0C90DC8A9} 
      _Version        =   393216
      Name            =   "Section3"
      Object.Height          =   360
      NumControls     =   0
   EndProperty
   SectionCode4    =   8
   BeginProperty Section4 {1C13A8E0-A0B6-11D0-848E-00A0C90DC8A9} 
      _Version        =   393216
      Name            =   "Section5"
      Object.Height          =   360
      NumControls     =   0
   EndProperty
End
Attribute VB_Name = "RptOSBordx"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub DataReport_Initialize()
    
    If DataEnvironment2.rsOSBordx.State = 1 Then
       DataEnvironment2.rsOSBordx.Requery
    End If
       strsql = IIf(InStr(1, ChkStr(FrmOSBordx.OSBORDXCriteria), "AND") = 0, "", Mid(Trim(FrmOSBordx.OSBORDXCriteria), InStr(1, ChkStr(FrmOSBordx.OSBORDXCriteria), "AND") + 3, Len(Trim(FrmOSBordx.OSBORDXCriteria))))
       DataEnvironment2.rsOSBordx.Filter = strsql
       
End Sub
