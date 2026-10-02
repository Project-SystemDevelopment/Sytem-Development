VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "tabctl32.ocx"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Begin VB.Form FrmStorage 
   Caption         =   "Storage Maintenance"
   ClientHeight    =   7935
   ClientLeft      =   60
   ClientTop       =   1980
   ClientWidth     =   8880
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   7935
   ScaleWidth      =   8880
   WindowState     =   2  'Maximized
   Begin TabDlg.SSTab SSTab1 
      Height          =   6975
      Left            =   120
      TabIndex        =   0
      Top             =   240
      Width           =   11700
      _ExtentX        =   20638
      _ExtentY        =   12303
      _Version        =   393216
      Tabs            =   2
      Tab             =   1
      TabsPerRow      =   2
      TabHeight       =   520
      TabCaption(0)   =   "Carton Details"
      TabPicture(0)   =   "FrmStorage.frx":0000
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "Lstcrt"
      Tab(0).Control(1)=   "Frame4"
      Tab(0).ControlCount=   2
      TabCaption(1)   =   "Storage Details"
      TabPicture(1)   =   "FrmStorage.frx":001C
      Tab(1).ControlEnabled=   -1  'True
      Tab(1).Control(0)=   "Lststr"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).Control(1)=   "Frame2"
      Tab(1).Control(1).Enabled=   0   'False
      Tab(1).ControlCount=   2
      Begin VB.Frame Frame2 
         Height          =   2775
         Left            =   360
         TabIndex        =   41
         Top             =   480
         Width           =   11055
         Begin VB.CheckBox ChkScan 
            Caption         =   "Use Scan"
            Height          =   255
            Left            =   5640
            TabIndex        =   56
            Top             =   360
            Value           =   1  'Checked
            Width           =   1095
         End
         Begin VB.TextBox TxtstrCasNo2 
            Height          =   285
            Left            =   3720
            MaxLength       =   20
            TabIndex        =   11
            Top             =   360
            Width           =   1815
         End
         Begin VB.TextBox TxtstrRemarks 
            Height          =   1095
            Left            =   1560
            MaxLength       =   4000
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   19
            Top             =   1440
            Width           =   8895
         End
         Begin VB.TextBox TxtstrCasNo 
            Height          =   285
            Left            =   1560
            MaxLength       =   20
            TabIndex        =   10
            Top             =   360
            Width           =   1815
         End
         Begin VB.TextBox TxtstrClmNo 
            Height          =   285
            Left            =   4560
            MaxLength       =   3
            TabIndex        =   12
            Top             =   360
            Visible         =   0   'False
            Width           =   975
         End
         Begin VB.ComboBox CboFileStatus 
            Height          =   315
            ItemData        =   "FrmStorage.frx":0038
            Left            =   4560
            List            =   "FrmStorage.frx":003A
            TabIndex        =   14
            Text            =   "I - In"
            Top             =   600
            Width           =   975
         End
         Begin MSMask.MaskEdBox MaskstrDORqst 
            Height          =   285
            Left            =   8640
            TabIndex        =   16
            Top             =   360
            Width           =   1695
            _ExtentX        =   2990
            _ExtentY        =   503
            _Version        =   393216
            MaxLength       =   10
            Mask            =   "##/##/####"
            PromptChar      =   "_"
         End
         Begin MSMask.MaskEdBox MaskstrDODelivery 
            Height          =   285
            Left            =   8640
            TabIndex        =   17
            Top             =   600
            Width           =   1695
            _ExtentX        =   2990
            _ExtentY        =   503
            _Version        =   393216
            MaxLength       =   10
            Mask            =   "##/##/####"
            PromptChar      =   "_"
         End
         Begin MSMask.MaskEdBox MaskstrDOReturn 
            Height          =   285
            Left            =   8640
            TabIndex        =   18
            Top             =   840
            Width           =   1695
            _ExtentX        =   2990
            _ExtentY        =   503
            _Version        =   393216
            MaxLength       =   10
            Mask            =   "##/##/####"
            PromptChar      =   "_"
         End
         Begin VB.ComboBox CboCarton 
            Height          =   315
            Left            =   1560
            TabIndex        =   13
            Top             =   600
            Width           =   2055
         End
         Begin VB.ComboBox CboUser 
            Height          =   315
            Left            =   1560
            Sorted          =   -1  'True
            TabIndex        =   15
            Top             =   880
            Width           =   3975
         End
         Begin VB.Label Label15 
            Caption         =   "To"
            Height          =   255
            Left            =   3460
            TabIndex        =   54
            Top             =   360
            Width           =   375
         End
         Begin VB.Label Label4 
            Caption         =   "Request By"
            Height          =   255
            Left            =   480
            TabIndex        =   50
            Top             =   960
            Width           =   975
         End
         Begin VB.Label Label11 
            Caption         =   "File Status"
            Height          =   255
            Left            =   3720
            TabIndex        =   49
            Top             =   660
            Width           =   855
         End
         Begin VB.Label Label1 
            Caption         =   "Case No"
            Height          =   255
            Left            =   480
            TabIndex        =   48
            Top             =   380
            Width           =   990
         End
         Begin VB.Label Label2 
            Caption         =   "Carton No"
            Height          =   255
            Left            =   480
            TabIndex        =   47
            Top             =   640
            Width           =   975
         End
         Begin VB.Label Label3 
            Caption         =   "Date Of Request"
            Height          =   255
            Left            =   7080
            TabIndex        =   46
            Top             =   375
            Width           =   1575
         End
         Begin VB.Label Label5 
            Caption         =   "Date Of Delivery"
            Height          =   255
            Left            =   7080
            TabIndex        =   45
            Top             =   645
            Width           =   1575
         End
         Begin VB.Label Label6 
            Caption         =   "Version No"
            Height          =   255
            Left            =   3720
            TabIndex        =   44
            Top             =   360
            Visible         =   0   'False
            Width           =   855
         End
         Begin VB.Label Label7 
            Caption         =   "Date Of Return"
            Height          =   255
            Left            =   7080
            TabIndex        =   43
            Top             =   915
            Width           =   1455
         End
         Begin VB.Label Label8 
            Caption         =   "Remarks"
            Height          =   255
            Left            =   480
            TabIndex        =   42
            Top             =   1560
            Width           =   735
         End
      End
      Begin VB.Frame Frame4 
         Height          =   2655
         Left            =   -74640
         TabIndex        =   31
         Top             =   600
         Width           =   11055
         Begin VB.TextBox TxtcrtCartonNo2 
            Height          =   315
            Left            =   3720
            MaxLength       =   20
            TabIndex        =   2
            Top             =   360
            Width           =   1575
         End
         Begin VB.TextBox TxtcrtRemarks 
            Height          =   855
            Left            =   1800
            MaxLength       =   3998
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   9
            Top             =   1560
            Width           =   8415
         End
         Begin VB.TextBox TxtcrtCartonNo 
            Height          =   315
            Left            =   1800
            MaxLength       =   20
            TabIndex        =   1
            Top             =   360
            Width           =   1575
         End
         Begin MSMask.MaskEdBox MaskcrtDORqst 
            Height          =   285
            Left            =   7560
            TabIndex        =   5
            Top             =   360
            Width           =   2655
            _ExtentX        =   4683
            _ExtentY        =   503
            _Version        =   393216
            MaxLength       =   10
            Mask            =   "##/##/####"
            PromptChar      =   "_"
         End
         Begin MSMask.MaskEdBox MaskcrtDODelivery 
            Height          =   285
            Left            =   7560
            TabIndex        =   6
            Top             =   600
            Width           =   2655
            _ExtentX        =   4683
            _ExtentY        =   503
            _Version        =   393216
            MaxLength       =   10
            Mask            =   "##/##/####"
            PromptChar      =   "_"
         End
         Begin MSMask.MaskEdBox MaskcrtDOReturn 
            Height          =   285
            Left            =   7560
            TabIndex        =   7
            Top             =   840
            Width           =   2655
            _ExtentX        =   4683
            _ExtentY        =   503
            _Version        =   393216
            MaxLength       =   10
            Mask            =   "##/##/####"
            PromptChar      =   "_"
         End
         Begin VB.ComboBox CboCartonStatus 
            Height          =   315
            ItemData        =   "FrmStorage.frx":003C
            Left            =   1800
            List            =   "FrmStorage.frx":003E
            TabIndex        =   3
            Text            =   "M - Medix"
            Top             =   640
            Width           =   3495
         End
         Begin MSMask.MaskEdBox MaskcrtDODestroy 
            Height          =   315
            Left            =   7560
            TabIndex        =   8
            Top             =   1080
            Width           =   2655
            _ExtentX        =   4683
            _ExtentY        =   556
            _Version        =   393216
            MaxLength       =   10
            Mask            =   "##/##/####"
            PromptChar      =   "_"
         End
         Begin MSMask.MaskEdBox MaskcrtDOSend 
            Height          =   285
            Left            =   1800
            TabIndex        =   4
            Top             =   915
            Width           =   3495
            _ExtentX        =   6165
            _ExtentY        =   503
            _Version        =   393216
            MaxLength       =   10
            Mask            =   "##/##/####"
            PromptChar      =   "_"
         End
         Begin VB.Label Label16 
            Caption         =   "To"
            Height          =   255
            Left            =   3450
            TabIndex        =   55
            Top             =   360
            Width           =   375
         End
         Begin VB.Label Label18 
            Caption         =   "Carton No"
            Height          =   255
            Left            =   480
            TabIndex        =   39
            Top             =   375
            Width           =   735
         End
         Begin VB.Label Label19 
            Caption         =   "Date Of Request"
            Height          =   255
            Left            =   6000
            TabIndex        =   38
            Top             =   390
            Width           =   1455
         End
         Begin VB.Label Label21 
            Caption         =   "Date Of Delivery"
            Height          =   255
            Left            =   6000
            TabIndex        =   37
            Top             =   645
            Width           =   1455
         End
         Begin VB.Label Label22 
            Caption         =   "Carton Status"
            Height          =   255
            Left            =   480
            TabIndex        =   36
            Top             =   680
            Width           =   975
         End
         Begin VB.Label Label23 
            Caption         =   "Date Of Return"
            Height          =   255
            Left            =   6000
            TabIndex        =   35
            Top             =   900
            Width           =   1335
         End
         Begin VB.Label Label24 
            Caption         =   "Remarks"
            Height          =   255
            Left            =   480
            TabIndex        =   34
            Top             =   1560
            Width           =   975
         End
         Begin VB.Label Label9 
            Caption         =   "Date Of Destroy"
            Height          =   255
            Left            =   6000
            TabIndex        =   33
            Top             =   1140
            Width           =   1215
         End
         Begin VB.Label Label10 
            Caption         =   "Date Of Sending"
            Height          =   255
            Left            =   480
            TabIndex        =   32
            Top             =   975
            Width           =   1455
         End
      End
      Begin MSComctlLib.ListView Lstcrt 
         Height          =   3495
         Left            =   -74640
         TabIndex        =   40
         Top             =   3360
         Width           =   11055
         _ExtentX        =   19500
         _ExtentY        =   6165
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         FullRowSelect   =   -1  'True
         GridLines       =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         NumItems        =   9
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "CartonNo"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Status"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   2
            SubItemIndex    =   2
            Text            =   "Date Register"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   3
            Text            =   "Date Send"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   4
            Text            =   "Date Request"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   5
            Text            =   "Date Delivery"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   6
            Text            =   "Date Return"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   7
            Text            =   "Date Destroy"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(9) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   8
            Text            =   "Remarks"
            Object.Width           =   2540
         EndProperty
      End
      Begin MSComctlLib.ListView Lststr 
         Height          =   3495
         Left            =   360
         TabIndex        =   51
         Top             =   3360
         Width           =   11055
         _ExtentX        =   19500
         _ExtentY        =   6165
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         FullRowSelect   =   -1  'True
         GridLines       =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         NumItems        =   12
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "CaseNo"
            Object.Width           =   2293
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Version"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "CartonNo"
            Object.Width           =   2117
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   2
            SubItemIndex    =   3
            Text            =   "Carton Status"
            Object.Width           =   2117
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   4
            Text            =   "File Status"
            Object.Width           =   1764
         EndProperty
         BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   2
            SubItemIndex    =   5
            Text            =   "Date Register"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   6
            Text            =   "RequestBy"
            Object.Width           =   2293
         EndProperty
         BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   7
            Text            =   "Date Request"
            Object.Width           =   2118
         EndProperty
         BeginProperty ColumnHeader(9) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   8
            Text            =   "Date Delivery"
            Object.Width           =   2118
         EndProperty
         BeginProperty ColumnHeader(10) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   9
            Text            =   "Date Return"
            Object.Width           =   2117
         EndProperty
         BeginProperty ColumnHeader(11) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   10
            Text            =   "Duration"
            Object.Width           =   1411
         EndProperty
         BeginProperty ColumnHeader(12) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   11
            Text            =   "Remarks"
            Object.Width           =   2540
         EndProperty
      End
   End
   Begin VB.Frame Frame5 
      Height          =   735
      Left            =   120
      TabIndex        =   30
      Top             =   7200
      Width           =   11715
      Begin VB.TextBox lblTotal 
         Height          =   285
         Left            =   1440
         TabIndex        =   60
         Top             =   240
         Width           =   615
      End
      Begin VB.TextBox lblCurrent 
         Height          =   285
         Left            =   240
         TabIndex        =   59
         Top             =   240
         Width           =   735
      End
      Begin VB.CommandButton CmdUpload 
         Caption         =   "&Upload"
         Default         =   -1  'True
         Height          =   350
         Left            =   3600
         TabIndex        =   57
         Top             =   240
         Width           =   675
      End
      Begin VB.CommandButton CmdListing 
         Caption         =   "&Listing"
         Height          =   350
         Left            =   8520
         TabIndex        =   26
         Top             =   240
         Width           =   675
      End
      Begin VB.CommandButton CmdPrintFile 
         Caption         =   "&Print File"
         Height          =   350
         Left            =   9160
         TabIndex        =   27
         Top             =   240
         Width           =   800
      End
      Begin VB.CommandButton CmdAdd 
         Caption         =   "&Add"
         Height          =   350
         Left            =   4280
         TabIndex        =   20
         Top             =   240
         Width           =   675
      End
      Begin VB.CommandButton CmdEdit 
         Caption         =   "&Edit"
         Height          =   350
         Left            =   4920
         TabIndex        =   21
         Top             =   240
         Width           =   705
      End
      Begin VB.CommandButton CmdDelete 
         Caption         =   "&Delete"
         Height          =   350
         Left            =   5640
         TabIndex        =   22
         Top             =   240
         Width           =   735
      End
      Begin VB.CommandButton CmdSearch 
         Caption         =   "&Search"
         Height          =   350
         Left            =   6360
         TabIndex        =   23
         Top             =   240
         Width           =   735
      End
      Begin VB.CommandButton CmdSave 
         Caption         =   "Sa&ve"
         Height          =   350
         Left            =   9960
         TabIndex        =   28
         Top             =   240
         Width           =   735
      End
      Begin VB.CommandButton CmdExit 
         Cancel          =   -1  'True
         Caption         =   "E&xit"
         Height          =   350
         Left            =   10680
         TabIndex        =   29
         Top             =   240
         Width           =   735
      End
      Begin VB.CommandButton CmdClear 
         Caption         =   "&Clear"
         Height          =   350
         Left            =   7800
         TabIndex        =   25
         Top             =   240
         Width           =   735
      End
      Begin VB.CommandButton CmdStop 
         Caption         =   "St&op"
         Height          =   350
         Left            =   7080
         TabIndex        =   24
         Top             =   240
         Width           =   735
      End
      Begin VB.Label Label14 
         Caption         =   "Records"
         Height          =   255
         Left            =   2280
         TabIndex        =   53
         Top             =   285
         Width           =   735
      End
      Begin VB.Label Label13 
         Caption         =   "of"
         Height          =   255
         Left            =   1200
         TabIndex        =   52
         Top             =   285
         Width           =   375
      End
   End
   Begin VB.Label Label12 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00404040&
      Caption         =   "Storage / Cartons"
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
      TabIndex        =   58
      Top             =   0
      Width           =   11805
   End
End
Attribute VB_Name = "FrmStorage"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public intExit As Integer
Dim bExit As Boolean
Dim EditFlag As Boolean
Dim AddFlag As Boolean
Dim listloop1 As Integer
Dim listloop2 As Integer
Dim Check As Boolean
Dim Check2 As Boolean

Public Function RemStr(str As String) As String
    RemStr = Replace(Trim(str), "'", "''")
End Function

Private Sub CboCarton_KeyPress(KeyAscii As Integer)
Dim NewText As String
Dim ValidCount As Integer
Dim ValidValue As String
Dim i As Integer
    If KeyAscii >= 32 And KeyAscii <> 127 Then
    NewText = LCase(Left(CboCarton.Text, CboCarton.SelStart) + Chr(KeyAscii) + Mid(CboCarton.Text, CboCarton.SelStart + CboCarton.SelLength + 1))
 
    ValidCount = 0
    ValidValue = ""
 
    For i = 0 To CboCarton.ListCount - 1
 
        If NewText = LCase(Left(CboCarton.List(i), Len(NewText))) Then
                ValidCount = ValidCount + 1
                ValidValue = CboCarton.List(i)
        End If
        Next
        If ValidCount <= 1 Then KeyAscii = 0
            If ValidCount = 1 Then
                CboCarton.Text = ValidValue
                CboCarton.SelStart = 0
                CboCarton.SelLength = Len(ValidValue)
            End If
        End If
End Sub

Private Sub CboCartonStatus_KeyPress(KeyAscii As Integer)
Dim NewText As String
Dim ValidCount As Integer
Dim ValidValue As String
Dim i As Integer
    If KeyAscii >= 32 And KeyAscii <> 127 Then
    NewText = LCase(Left(CboCartonStatus.Text, CboCartonStatus.SelStart) + Chr(KeyAscii) + Mid(CboCartonStatus.Text, CboCartonStatus.SelStart + CboCartonStatus.SelLength + 1))
 
    ValidCount = 0
    ValidValue = ""
 
    For i = 0 To CboCartonStatus.ListCount - 1
 
        If NewText = LCase(Left(CboCartonStatus.List(i), Len(NewText))) Then
                ValidCount = ValidCount + 1
                ValidValue = CboCartonStatus.List(i)
        End If
        Next
        If ValidCount <= 1 Then KeyAscii = 0
            If ValidCount = 1 Then
                CboCartonStatus.Text = ValidValue
                CboCartonStatus.SelStart = 0
                CboCartonStatus.SelLength = Len(ValidValue)
            End If
        End If
End Sub

Private Sub CboFileStatus_KeyPress(KeyAscii As Integer)
Dim NewText As String
Dim ValidCount As Integer
Dim ValidValue As String
Dim i As Integer
    If KeyAscii >= 32 And KeyAscii <> 127 Then
    NewText = LCase(Left(CboFileStatus.Text, CboFileStatus.SelStart) + Chr(KeyAscii) + Mid(CboFileStatus.Text, CboFileStatus.SelStart + CboFileStatus.SelLength + 1))
 
    ValidCount = 0
    ValidValue = ""
 
    For i = 0 To CboFileStatus.ListCount - 1
 
        If NewText = LCase(Left(CboFileStatus.List(i), Len(NewText))) Then
                ValidCount = ValidCount + 1
                ValidValue = CboFileStatus.List(i)
        End If
        Next
        If ValidCount <= 1 Then KeyAscii = 0
            If ValidCount = 1 Then
                CboFileStatus.Text = ValidValue
                CboFileStatus.SelStart = 0
                CboFileStatus.SelLength = Len(ValidValue)
            End If
        End If

End Sub


Private Sub CboUser_KeyPress(KeyAscii As Integer)
Dim NewText As String
Dim ValidCount As Integer
Dim ValidValue As String
Dim i As Integer
    If KeyAscii >= 32 And KeyAscii <> 127 Then
    NewText = LCase(Left(cboUser.Text, cboUser.SelStart) + Chr(KeyAscii) + Mid(cboUser.Text, cboUser.SelStart + cboUser.SelLength + 1))
 
    ValidCount = 0
    ValidValue = ""
 
    For i = 0 To cboUser.ListCount - 1
 
        If NewText = LCase(Left(cboUser.List(i), Len(NewText))) Then
                ValidCount = ValidCount + 1
                ValidValue = cboUser.List(i)
        End If
        Next
        If ValidCount <= 1 Then KeyAscii = 0
            If ValidCount = 1 Then
                cboUser.Text = ValidValue
                cboUser.SelStart = 0
                cboUser.SelLength = Len(ValidValue)
            End If
        End If
End Sub

Private Sub ChkScan_Click()
If ChkScan.Value = 0 Then
    CmdUpload.Visible = False
    Label15.Visible = True
    TxtstrCasNo2.Visible = True
    'Lststr.Visible = True
    'Lststr2.Visible = False
    ChkScan.Left = 5760
    ChkScan.Top = 380
    TxtstrCasNo.SetFocus
Else
    Call EntriesEnable(True)
    TxtstrCasNo.Text = ""
    TxtstrCasNo2.Text = ""
    CboCarton.Text = ""
    CboFileStatus.Text = ""
    cboUser.Text = ""
    TxtstrRemarks.Text = ""
    MaskstrDORqst.Text = "__/__/____"
    MaskstrDODelivery.Text = "__/__/____"
    MaskstrDOReturn.Text = "__/__/____"
    CmdUpload.Visible = True
    CmdUpload.Default = True
    Label15.Visible = False
    TxtstrCasNo2.Visible = False
    Lststr.ListItems.Clear
    'Lststr2.Visible = True
    ChkScan.Left = 3480
    ChkScan.Top = 380
    TxtstrCasNo.SetFocus
End If
End Sub

Private Sub cmdAdd_Click()
EditFlag = False
AddFlag = True

Select Case SSTab1.Tab
    Case 0
        Call ClearForm(Me)
    Case 1
        TxtstrCasNo.Text = ""
        TxtstrCasNo2.Text = ""
        CboCarton.Text = ""
        CboFileStatus.Text = ""
        cboUser.Text = ""
        TxtstrRemarks.Text = ""
        MaskstrDORqst.Text = "__/__/____"
        MaskstrDODelivery.Text = "__/__/____"
        MaskstrDOReturn.Text = "__/__/____"
        Lststr.ListItems.Clear
End Select
'Call ClearForm(Me)
Call DisableCmdButton(Me)
Call EntriesEnable(True)
CmdSave.Enabled = True
CmdAdd.Enabled = True
CmdDelete.Enabled = False
    Select Case SSTab1.Tab
        Case 0
            TxtcrtCartonNo.SetFocus
        Case 1
            TxtstrCasNo.SetFocus
    End Select
End Sub

Private Sub CmdClear_Click()
    
Check = False
Check2 = False

    Select Case SSTab1.Tab
        Case 0
            Call ClearForm(Me)
        Case 1
            TxtstrCasNo.Text = ""
            TxtstrCasNo2.Text = ""
            CboCarton.Text = ""
            CboFileStatus.Text = ""
            cboUser.Text = ""
            TxtstrRemarks.Text = ""
            MaskstrDORqst.Text = "__/__/____"
            MaskstrDODelivery.Text = "__/__/____"
            MaskstrDOReturn.Text = "__/__/____"
            'Lststr.listitems.Clear
            'Lststr2.listitems.Clear
            If ChkScan.Value = 0 Then
                Call EntriesEnable(True)
                TxtstrCasNo.SetFocus
            Else
                TxtstrCasNo.SetFocus
            End If
    End Select
    'Call ClearForm(Me)
    Call EntriesEnable(True)
    CmdAdd.Enabled = True
    CmdDelete.Enabled = False
    CmdEdit.Enabled = False
    Lstcrt.ListItems.Clear
    Lststr.ListItems.Clear
    'Lststr2.listitems.Clear
    txtTotalRecord = 0
End Sub

Private Sub CmdDelete_Click()
Dim rst2 As New ADODB.Recordset
Dim sql2 As String
Dim RecordSaved
    
    Screen.MousePointer = vbHourglass
    CmdDelete.Enabled = False
    'medixmasterconnWS.Open "File Name=" & BOSPath & "Conn\mediconnectWS.UDL"
    RecordSaved = MsgBox("Do you want to delete the record?", vbYesNo + vbExclamation)
    If RecordSaved = vbYes Then
        sql2 = ""
        If SSTab1.Tab = 0 Then
            sql2 = "Delete From CARTON Where CrtCartonNo = '" & TxtcrtCartonNo & "'"
        Else
    '       sql2 = "Delete From STORAGE Where StrCaseNo = '" & TxtstrCasNo & "' AND StrClmversion = '" & TxtstrClmNo & "'"
            sql2 = "Delete From STORAGE Where StrCaseNo = '" & TxtstrCasNo & "'"
        End If
        rst2.Open sql2, medixmasterconnWS, adOpenKeyset, adLockOptimistic
        MsgBox "Record Deleted... ", vbOKOnly + vbInformation, DialogBoxTitle
        CmdSave.Enabled = True
        Call ClearForm(Me)
        Call DisableCmdButton(Me)
        Call EntriesEnable(True)
        If SSTab1.Tab = 0 Then
            Call CartonListing
            TxtcrtCartonNo.SetFocus
        Else
            Call StorageListing
            TxtstrCasNo.SetFocus
        End If
    End If
   ' medixmasterconnWS.Close
   ' Set medixmasterconnWS = Nothing
    CmdDelete.Enabled = True
    Screen.MousePointer = vbDefault
End Sub

Private Sub cmdEdit_Click()
    EditFlag = True
    Call EntriesEnable(True)
    CmdAdd.Enabled = False
    CmdDelete.Enabled = False
    CmdSave.Enabled = True
    If TxtcrtCartonNo <> "" Then
       TxtcrtCartonNo.Enabled = False
       CboCartonStatus.SetFocus
    End If
    If TxtstrCasNo <> "" Then
       TxtstrCasNo.Enabled = False
'       TxtstrClmNo.Enabled = False
       CboCarton.SetFocus
    End If
End Sub

Private Sub cmdExit_Click()
'Dim RecordExit
'    RecordExit = MsgBox("Quit this form?", vbYesNo + vbExclamation)
'    If RecordExit = vbYes Then
'        bExit = True
        Unload Me
'    End If
End Sub
Private Function StorageSearchRecord()
Dim strsql As String
Dim strAppend As String
Dim sql As String
Dim storage As New ADODB.Recordset

    strAppend = ""
    
    If ChkScan.Value = 1 Then
        If Len(Trim(TxtstrCasNo)) Then
            strAppend = strAppend + " AND STRCaseNo = '" & RemStr(Trim(TxtstrCasNo)) & "'"
        End If
    Else
        If Len(Trim(TxtstrCasNo)) Then
            strAppend = strAppend + " AND STRCaseNo >= '" & RemStr(Trim(TxtstrCasNo)) & "'"
        End If
        If Len(Trim(TxtstrCasNo2)) Then
            strAppend = strAppend + " AND STRCaseNo <= '" & RemStr(Trim(TxtstrCasNo2)) & "'"
        End If
    End If
    
'    If Len(Trim(TxtstrClmNo)) Then
'        strappend = strappend + " AND strclmversion = '" & RemStr(Trim(TxtstrClmNo)) & "'"
'    End If
    
    If Len(Trim(CboCarton)) Then
        strAppend = strAppend + " AND STRCartonNo = '" & RemStr(Trim(CboCarton)) & "'"
    End If
    
    If Len(Trim(Mid(CboFileStatus, 1, 1))) Then
        strAppend = strAppend + " AND STRFileStatus = '" & RemStr(Trim(Mid(CboFileStatus, 1, 1))) & "'"
    End If
    
    If Len(Trim(cboUser)) Then
        strAppend = strAppend + " AND STRRequestBy = '" & RemStr(Trim(cboUser)) & "'"
    End If
    
    If IsDate(MaskstrDORqst) = True Then
        strAppend = strAppend + " And DORequest = '" & Format(MaskstrDORqst, "mm/dd/yyyy") & "'"
    End If
    
    If IsDate(MaskstrDODelivery) = True Then
        strAppend = strAppend + " And DODelivery = '" & Format(MaskstrDODelivery, "mm/dd/yyyy") & "'"
    End If
    
    If IsDate(MaskstrDOReturn) = True Then
        strAppend = strAppend + " And DOReturn = '" & Format(MaskstrDOReturn, "mm/dd/yyyy") & "'"
    End If
    
    Call StorageListing(strAppend)
    
End Function

Private Sub StorageListing(Optional strAppend As String)
Dim rst2 As New ADODB.Recordset
Dim rstCount As New ADODB.Recordset
Dim Record As ListItem
Dim sql As String
Dim sqlCount As String
Dim total As String
Dim Current As String
    
total = 0
Current = 0
    
    Lststr.ListItems.Clear

    'Sql = " Select STRCaseNo, STRCLMVersion, STRCartonNo, STRRequestBy, DORequest,  " & _
          " DODelivery, DOReturn, STRRemarks, STRFileStatus from STORAGE where 0 = 0 "
          
    sql = " Select * from View_Storage where 0 = 0 "
        
    If Len(Trim(strAppend)) Then
        sql = sql + strAppend
    End If
    
    
    rst2.Open sql, medixmasterconnWS, adOpenKeyset, adLockOptimistic
    
    If rst2.recordCount = 0 Then
        MsgBox "No record found!", vbOKOnly + vbCritical, DialogBoxTitle
        CmdAdd.Enabled = True
    Else
    
    Do
        Do Until rst2.EOF
        total = rst2.recordCount
        lblTotal = total
            With rst2
                Set Record = Lststr.ListItems.Add(, , rst2!StrCaseNo)
                
                If Len(rst2!StrCLMVersion) Then
                    Record.SubItems(1) = rst2!StrCLMVersion
                End If
                If Len(rst2!STRCartonNo) Then
                    Record.SubItems(2) = rst2!STRCartonNo
                End If
                If Len(rst2!CRTCartonStatus) Then
                    Record.SubItems(3) = rst2!CRTCartonStatus
                End If
                If Len(rst2!STRFileStatus) Then
                    Record.SubItems(4) = rst2!STRFileStatus
                End If
                If Len(rst2!LastUpdateDate) Then
                    Record.SubItems(5) = Format(rst2!LastUpdateDate, "dd/mm/yyyy")
                End If
                If Len(rst2!STRRequestBy) Then
                    Record.SubItems(6) = rst2!STRRequestBy
                End If
                If Len(rst2!DORequest) Then
                    Record.SubItems(7) = Format(rst2!DORequest, "dd/mm/yyyy")
                End If
                If Len(rst2!DODelivery) Then
                    Record.SubItems(8) = Format(rst2!DODelivery, "dd/mm/yyyy")
                End If
                If Len(rst2!DOReturn) Then
                    Record.SubItems(9) = Format(rst2!DOReturn, "dd/mm/yyyy")
                End If
                If Len(rst2!DOReturn) And Len(rst2!DODelivery) Then
                    Record.SubItems(10) = DateValue(rst2!DOReturn) - DateValue(rst2!DODelivery)
                End If
                If Len(rst2!STRRemarks) Then
                    Record.SubItems(11) = rst2!STRRemarks
                End If
                
                Current = Current + 1
                lblCurrent = Current
            End With
            rst2.MoveNext
            DoEvents
                If intExit = 1 Then
                   Check = False
                   Exit Do
                End If
        Loop
    Loop Until Check = False
End If
intExit = 0
rst2.Close
Set rst2 = Nothing
CmdExit.Enabled = True
End Sub

Private Sub CmdListing_Click()
If SSTab1.Tab = 1 Then
    If Trim(CboCarton) = "" Then
        MsgBox "Please select Carton No. ", vbOKOnly + vbCritical, DialogBoxTitle
        CboCarton.SetFocus
    Else
'On Error Resume Next
    Dim objExcel As Excel.Application
    Dim objWorkbook As Excel.Workbook
    Dim objWorksheet As Excel.WORKSHEET
    Dim AA, BB As Integer
    Dim yy, Cnt As Integer
    Dim strsql As String
    Dim Sto As New ADODB.Recordset
    Dim ECol1 As String
    Dim ECol2 As String
    Dim ECol3 As String
       
    Screen.MousePointer = vbHourglass

        If objExcel Is Nothing Then
            Set objExcel = CreateObject("Excel.application")
                objExcel.Interactive = False
            If objExcel.Visible = False Then
                objExcel.Visible = True
            End If
            objExcel.WindowState = xlMaximized
            Set objWorkbook = objExcel.Workbooks.Add(LocCardPath & "Storage\StorageTracking.xlt")
            Set objWorksheet = objWorkbook.Sheets(1)
            objExcel.DisplayAlerts = False
            Set objWorksheet = objWorkbook.Worksheets.Item(1)
        End If
        
        If Err.Number > 0 Then
            MsgBox "Microsoft Excel is Not loaded On this machine.", vbOKOnly + vbCritical, "Error Loading Excel"
            GoTo ExitFunction
        End If
        
        On Error GoTo HANDLE_ERROR
    
        strsql = "select STRCaseNo,STRFileStatus from Storage WHERE STRCartonNO  = '" & CboCarton & "'"
        Sto.Open strsql, medixmasterconnWS, adOpenKeyset, adLockOptimistic
        'Set objExcel = New excel.Application
        
        If Sto.recordCount <> 0 Then
            RecordSelected = Sto.recordCount
            objWorksheet.Cells(2, "F") = Format(Date, "MM/DD/YYYY")
            objWorksheet.Cells(2, "H") = CboCarton
            BB = 0
            yy = 0
            Cnt = 0
            ECol1 = Chr(65)
            ECol2 = Chr(66)
            ECol3 = Chr(67)
            
            'Do
            Do Until Sto.EOF
            'For i = 1 To RecordSelected
                BB = BB + 1
                If BB > 50 Then
                    yy = yy + 4
                    BB = 1
                    ECol1 = Chr(65 + yy)
                    ECol2 = Chr(66 + yy)
                    ECol3 = Chr(67 + yy)
                    objWorksheet.Cells(4, ECol1) = "No"
                    objWorksheet.Cells(4, ECol2) = "Case No"
                    objWorksheet.Cells(4, ECol3) = "Checking"
                End If
                
                Cnt = Cnt + 1
                objWorksheet.Cells(4 + CDbl(BB), ECol1) = Cnt
                objWorksheet.Cells(4 + CDbl(BB), ECol2) = Sto!StrCaseNo
                objWorksheet.Cells(4 + CDbl(BB), ECol3) = Sto!STRFileStatus
                
                Sto.MoveNext
            
            'Next i
        Loop
            
        End If
            'cmdCredit.Enabled = False
            objExcel.DisplayAlerts = True
            'objWorksheet.Columns.AutoFit
            objExcel.Interactive = True
            Screen.MousePointer = vbDefault
            Exit Sub
ExitFunction:
            objExcel.DisplayAlerts = True
            objExcel.Interactive = True
            objExcel.Quit
            Screen.MousePointer = vbDefault
            Exit Sub
HANDLE_ERROR:
            MsgBox "Export To Excel failed. Encountered the following Error" & vbCrLf & vbCrLf & _
            Err.Number & ": " & Err.Description, vbOKOnly + vbCritical, "Error Exporting To Excel"
            Set objWorksheet = Nothing
            Set objWorkbook = Nothing
            GoTo ExitFunction
    End If
End If
End Sub


Private Sub CmdPrintFile_Click()
Select Case SSTab1.Tab
    Case 0
        Screen.MousePointer = vbHourglass
        
            If Lstcrt.ListItems.Count <> 0 Then
                Call ExportListViewtoExcel(Lstcrt)
            Else
                MsgBox "Report cannot be printed without any record.", vbInformation
            End If
        
        Screen.MousePointer = vbDefault
         
         'Call ExportToExcelWorksheet(Lstcrt)
    Case 1
         'Call ExportToExcelWorksheet(Lststr)
        Screen.MousePointer = vbHourglass
        
        If Lststr.ListItems.Count <> 0 Then
            Call ExportListViewtoExcel(Lststr)
        Else
            MsgBox "Report cannot be printed without any record.", vbInformation
        End If
        
        Screen.MousePointer = vbDefault
         
    End Select
End Sub

Private Sub cmdSave_Click()
    Screen.MousePointer = vbHourglass
    CmdSave.Enabled = False
    'medixmasterconnWS.Open "File Name=" & BOSPath & "Conn\mediconnectWS.UDL"
    Select Case SSTab1.Tab
        Case 0
            Call SaveCarton
        Case 1
            If ChkScan.Value = 0 Then
                Call SaveStore
            Else
                Call SaveStore2
            End If
    End Select
    CmdAdd.Enabled = True
    CmdEdit.Enabled = False
   ' medixmasterconnWS.Close
   ' Set medixmasterconnWS = Nothing
    CmdSave.Enabled = True
    Screen.MousePointer = vbDefault
End Sub

Private Sub SaveCarton()
Dim RecordSaved
Dim CARTON As New ADODB.Recordset
Dim SaveCarton As New ADODB.Recordset
Dim strsql As String
Dim strAppend As String
    
    If TxtcrtCartonNo = "" Then
        MsgBox "Please Enter Carton Number", vbOKOnly + vbCritical, DialogBoxTitle
        TxtcrtCartonNo.SetFocus
    ElseIf Trim(Mid(CboCartonStatus, 1, 1)) = "" Then
        MsgBox "Please Enter Carton Status", vbOKOnly + vbCritical, DialogBoxTitle
        CboCartonStatus.SetFocus
     ElseIf Format(MaskcrtDORqst, "YYYYMMDD") < Format(MaskcrtDOSend, "YYYYMMDD") Then
        MsgBox "Invalid Date ! Date of Request must greater than Date of Send ", vbOKOnly + vbCritical, DialogBoxTitle
        MaskcrtDORqst = "__/__/____"
        MaskcrtDORqst.SetFocus
    ElseIf IsDate(MaskcrtDOReturn) < IsDate(MaskcrtDODelivery) Then
        MsgBox "Invalid Date ! Date of Return must greater than Date of Delivery ", vbOKOnly + vbCritical, DialogBoxTitle
        MaskcrtDOReturn = "__/__/____"
        MaskcrtDOReturn.SetFocus
    ElseIf IsDate(MaskcrtDODestroy) < IsDate(MaskcrtDOReturn) Then
        MsgBox "Invalid Date ! Date of Destroy must greater than Date of Return ", vbOKOnly + vbCritical, DialogBoxTitle
        MaskcrtDODestroy = "__/__/____"
        MaskcrtDODestroy.SetFocus
    ElseIf Trim(Mid(CboCartonStatus, 1, 1)) = "R" And Trim(MaskcrtDOReturn) = "__/__/____" Then
        MsgBox "Date Return cannot be null when the status is 'R' !", vbOKOnly + vbCritical, DialogBoxTitle
        MaskcrtDOReturn.SetFocus
    Else
          
        strsql = "Select * From CARTON Where CRTCartonNo = '" & RemStr(TxtcrtCartonNo) & "'"
        CARTON.Open strsql, medixmasterconnWS, adOpenKeyset, adLockOptimistic
    
        If CARTON.recordCount > 0 And EditFlag = False Then
            MsgBox "Record Duplicated! Please Change The Corresponding Field", vbCritical, "Error"
            CARTON.Close
            TxtcrtCartonNo.SetFocus
            TxtcrtCartonNo.SelStart = 0
            TxtcrtCartonNo.SelLength = Len(TxtcrtCartonNo)
            Set CARTON = Nothing
            Exit Sub
        Else
            RecordSaved = MsgBox("Do you want to save the record?", vbYesNo + vbExclamation)
            If RecordSaved = vbYes Then
                If EditFlag = False Then
               '     CARTON.Open "CARTON", medixmasterconnWS, adOpenKeyset, adLockOptimistic
                    CARTON.AddNew
                End If
                With CARTON
                    !CRTCartonNo = ChkStr(TxtcrtCartonNo)
                    !CRTCartonStatus = Trim(Mid(CboCartonStatus, 1, 1))
                    If MaskcrtDOSend <> "__/__/____" Then
                        !DoSend = MaskcrtDOSend
                    Else
                        !DoSend = Null
                    End If
                    If MaskcrtDORqst <> "__/__/____" Then
                        !DORequest = MaskcrtDORqst
                    Else
                        !DORequest = Null
                    End If
                    If MaskcrtDODelivery <> "__/__/____" Then
                        !DODelivery = MaskcrtDODelivery
                    Else
                        !DODelivery = Null
                    End If
                    If MaskcrtDOReturn <> "__/__/____" Then
                        !DOReturn = MaskcrtDOReturn
                    Else
                        !DOReturn = Null
                    End If
                    If MaskcrtDODestroy <> "__/__/____" Then
                        !dodestroy = MaskcrtDODestroy
                    Else
                        !dodestroy = Null
                    End If
                    !CRTRemarks = ChkStr(TxtcrtRemarks)
                    !LastUpdateBy = Logon
                    !LastUpdateDate = Date
                End With
                CARTON.Update
                If EditFlag = False Then
                    MsgBox "Record Saved...", vbOKOnly + vbInformation, DialogBoxTitle
                Else
                    MsgBox "Record Updated...", vbOKOnly + vbInformation, DialogBoxTitle
                End If
                CARTON.Close
                Set CARTON = Nothing
                strAppend = ""
                If Len(Trim(TxtcrtCartonNo)) Then
                    strAppend = " AND CRTCartonNo = '" & RemStr(Trim(TxtcrtCartonNo)) & "'"
                End If

                Call ClearForm(Me)
                Call CartonListing(strAppend)
            End If
        End If
        EditFlag = False
    End If
End Sub

Private Sub cmdSearch_Click()
    CmdAdd.Enabled = True
    CmdEdit.Enabled = False
    CmdDelete.Enabled = False
    CmdSave.Enabled = False
    CmdExit.Enabled = False
    
    Screen.MousePointer = vbHourglass
    CmdSearch.Enabled = False
    'medixmasterconnWS.Open "File Name=" & BOSPath & "Conn\mediconnectWS.UDL"
    Select Case SSTab1.Tab
        Case 0
            Call CartonSearchRecord
        Case 1
            Call StorageSearchRecord
    End Select
    'medixmasterconnWS.Close
    'Set medixmasterconnWS = Nothing
    CmdSearch.Enabled = True
    Screen.MousePointer = Default
End Sub

Private Function CartonSearchRecord()
Dim strsql As String
Dim strAppend As String
Dim sql As String
Dim storage As New ADODB.Recordset

    strAppend = ""
    If Len(Trim(TxtcrtCartonNo)) Then
        strAppend = strAppend + " AND CRTCartonNo >= '" & RemStr(Trim(TxtcrtCartonNo)) & "'"
    End If
    
    If Len(Trim(TxtcrtCartonNo2)) Then
        strAppend = strAppend + " AND CRTCartonNo <= '" & RemStr(Trim(TxtcrtCartonNo2)) & "'"
    End If
    
    If Len(Trim(Mid(CboCartonStatus, 1, 1))) Then
        strAppend = strAppend + " AND CRTCartonStatus = '" & RemStr(Trim(Mid(CboCartonStatus, 1, 1))) & "'"
    End If
    
    If IsDate(MaskcrtDOSend) = True Then
        strAppend = strAppend + " And DOSend = '" & Format(MaskcrtDOSend, "mm/dd/yyyy") & "'"
    End If
    
    If IsDate(MaskcrtDORqst) = True Then
        strAppend = strAppend + " And DORequest = '" & Format(MaskcrtDORqst, "mm/dd/yyyy") & "'"
    End If
    
    If IsDate(MaskcrtDODelivery) = True Then
        strAppend = strAppend + " And DODelivery = '" & Format(MaskcrtDODelivery, "mm/dd/yyyy") & "'"
    End If
    
    If IsDate(MaskcrtDOReturn) = True Then
        strAppend = strAppend + " And DOReturn = '" & Format(MaskcrtDOReturn, "mm/dd/yyyy") & "'"
    End If
    
    If IsDate(MaskcrtDODestroy) = True Then
        strAppend = strAppend + " And DODestroy = '" & Format(MaskcrtDODestroy, "mm/dd/yyyy") & "'"
    End If
    
    Call CartonListing(strAppend)
End Function

Private Sub CartonListing(Optional strAppend As String)
Dim rst2 As New ADODB.Recordset
Dim rstCount As New ADODB.Recordset
Dim Record As ListItem
Dim sql As String
Dim sqlCount As String
Dim total As String
Dim Current As String
    
total = 0
Current = 0
    
    Lstcrt.ListItems.Clear

    sql = " Select crtcartonno, crtcartonstatus, dosend, dorequest, " & _
          " dodelivery, doreturn, dodestroy, crtremarks, lastupdatedate from carton where 0 = 0 "
        
    If Len(Trim(strAppend)) Then
        sql = sql + strAppend
    End If
    
    rst2.Open sql, medixmasterconnWS, adOpenKeyset, adLockOptimistic
    
    If rst2.recordCount = 0 Then
        MsgBox "No record found!", vbOKOnly + vbCritical, DialogBoxTitle
        CmdAdd.Enabled = True
    Else
    total = rst2.recordCount
    lblTotal = total
    Do
        Do Until rst2.EOF
        
           With rst2
                Set Record = Lstcrt.ListItems.Add(, , rst2!CRTCartonNo)
                
                If Len(rst2!CRTCartonStatus) Then
                    Record.SubItems(1) = rst2!CRTCartonStatus
                End If
                If Len(rst2!LastUpdateDate) Then
                    Record.SubItems(2) = Format(rst2!LastUpdateDate, "dd/mm/yyyy")
                End If
                If Len(rst2!DoSend) Then
                    Record.SubItems(3) = Format(rst2!DoSend, "dd/mm/yyyy")
                End If
                If Len(rst2!DORequest) Then
                    Record.SubItems(4) = Format(rst2!DORequest, "dd/mm/yyyy")
                End If
                If Len(rst2!DODelivery) Then
                    Record.SubItems(5) = Format(rst2!DODelivery, "dd/mm/yyyy")
                End If
                If Len(rst2!DOReturn) Then
                    Record.SubItems(6) = Format(rst2!DOReturn, "dd/mm/yyyy")
                End If
                If Len(rst2!dodestroy) Then
                    Record.SubItems(7) = Format(rst2!dodestroy, "dd/mm/yyyy")
                End If
                If Len(rst2!CRTRemarks) Then
                    Record.SubItems(8) = rst2!CRTRemarks
                End If
                Current = Current + 1
                lblCurrent = Current
            End With
            rst2.MoveNext
            DoEvents
            If intExit = 1 Then
                Check = False
                Exit Do
            End If
          
        Loop
    Loop Until Check = False
    End If
    intExit = 0
    rst2.Close
    Set rst2 = Nothing
    CmdExit.Enabled = True
End Sub

Private Sub CmdStop_Click()
intExit = 1
End Sub

Private Sub SaveStore()
Dim RecordSaved
Dim SaveSTORAGE As New ADODB.Recordset
Dim strsql As String
Dim sql2 As String
Dim Store As New ADODB.Recordset
Dim CLM As New ADODB.Recordset
Dim CARTON As New ADODB.Recordset
Dim strAppend As String

    If Trim(TxtstrCasNo) = "" Then
        MsgBox "Please Enter Case Number", vbOKOnly + vbCritical, DialogBoxTitle
        TxtstrCasNo.SetFocus
'    ElseIf Trim(TxtstrClmNo) = "" Then
'        MsgBox "Please Enter Claim Number", vbOKOnly + vbCritical, DialogBoxTitle
'        TxtstrClmNo.SetFocus
    ElseIf Trim(CboCarton) = "" Then
        MsgBox "Please Enter Carton Number", vbOKOnly + vbCritical, DialogBoxTitle
        CboCarton.SetFocus
    ElseIf Trim(Mid(CboFileStatus, 1, 1)) = "" Then
        MsgBox "Please Enter File Status", vbOKOnly + vbCritical, DialogBoxTitle
        CboFileStatus.SetFocus
    ElseIf cboUser <> "" And MaskstrDORqst = "__/__/____" Then
        MsgBox "Please Enter Date Of Request", vbOKOnly + vbCritical, DialogBoxTitle
        MaskstrDORqst.SetFocus
    ElseIf cboUser <> "" And Format(MaskstrDODelivery, "YYYYMMDD") < Format(MaskstrDORqst, "YYYYMMDD") Then
        MsgBox "Invalid Date ! Date of Delivery must greater than Date of Request ", vbOKOnly + vbCritical, DialogBoxTitle
        MaskstrDODelivery = "__/__/____"
        MaskstrDODelivery.SetFocus
    ElseIf cboUser <> "" And Format(MaskstrDOReturn, "YYYYMMDD") < Format(MaskstrDODelivery, "YYYYMMDD") Then
        MsgBox "Invalid Date ! Date of Return must greater than Date of Delivery ", vbOKOnly + vbCritical, DialogBoxTitle
        MaskstrDOReturn = "__/__/____"
        MaskstrDOReturn.SetFocus
    Else
    
'        If Trim(TxtstrCasNo) <> "" Or Trim(TxtstrClmNo) <> "" Then
         If Trim(TxtstrCasNo) <> "" Then
'           sql2 = " Select * from CLAIM where CASNumber = '" & RemStr(TxtstrCasNo) & "' AND ClaimVersion = '" & RemStr(TxtstrClmNo) & "'"
'           sql2 = " Select * from CLAIM where CASNumber = '" & RemStr(TxtstrCasNo) & "'"
            sql2 = " Select * from CAS where CASNumber = '" & RemStr(TxtstrCasNo) & "'"
            
            CLM.Open sql2, medixmasterconn, adOpenKeyset, adLockOptimistic
            If CLM.recordCount = 0 Then
                MsgBox "No Such Case ! ", vbOKOnly + vbInformation, DialogBoxTitle
                CLM.Close
                Set CLM = Nothing
                TxtstrCasNo.SetFocus
                Exit Sub
            End If
        End If

'       strsql = "Select * From STORAGE Where STRCaseNo = '" & RemStr(TxtstrCasNo) & "' AND STRClmVersion = '" & RemStr(TxtstrClmNo) & "'"
        strsql = "Select * From STORAGE Where STRCaseNo = '" & RemStr(TxtstrCasNo) & "'"
        
        SaveSTORAGE.Open strsql, medixmasterconnWS, adOpenKeyset, adLockOptimistic
    
        If SaveSTORAGE.recordCount > 0 And EditFlag = False Then
            MsgBox "Record Duplicated! Please Change The Corresponding Field", vbCritical, "Error"
            SaveSTORAGE.Close
            TxtcrtCartonNo.SetFocus
            TxtcrtCartonNo.SelStart = 0
            TxtcrtCartonNo.SelLength = Len(TxtcrtCartonNo)
            Set SaveSTORAGE = Nothing
            Exit Sub
        Else
            RecordSaved = MsgBox("Do you want to save the record?", vbYesNo + vbExclamation)
            If RecordSaved = vbYes Then
                If EditFlag = False Then
                    SaveSTORAGE.Close
                    SaveSTORAGE.Open "STORAGE", medixmasterconnWS, adOpenKeyset, adLockOptimistic
                    SaveSTORAGE.AddNew
                End If
                With SaveSTORAGE
                    !StrCaseNo = ChkStr(TxtstrCasNo)
'                    !STRClmVersion = ChkStr(TxtstrClmNo)
                    !STRCartonNo = ChkStr(CboCarton)
                    !STRFileStatus = Trim(Mid(CboFileStatus, 1, 1))
                    !STRRequestBy = ChkStr(cboUser)
                    If cboUser <> "" Then
                        !DORequest = MaskstrDORqst
                        If MaskstrDODelivery <> "__/__/____" Then
                            !DODelivery = MaskstrDODelivery
                        Else
                            !DODelivery = Null
                        End If
                        If MaskstrDOReturn <> "__/__/____" Then
                            !DOReturn = MaskstrDOReturn
                        Else
                            !DOReturn = Null
                        End If
                    Else
                        !DODelivery = Null
                        !DOReturn = Null
                        !DORequest = Null
                    End If
                    !STRRemarks = ChkStr(TxtstrRemarks)
                    !LastUpdateBy = Logon
                    !LastUpdateDate = Date
                End With
                SaveSTORAGE.Update
                If EditFlag = False Then
                    MsgBox "Record Saved...", vbOKOnly + vbInformation, DialogBoxTitle
                Else
                    MsgBox "Record Updated...", vbOKOnly + vbInformation, DialogBoxTitle
                End If
                SaveSTORAGE.Close
                Set SaveSTORAGE = Nothing
                
                strAppend = ""
                If Len(Trim(TxtstrCasNo)) Then
                    strAppend = " AND STRCaseNo = '" & RemStr(Trim(TxtstrCasNo)) & "'"
                End If

                Call ClearForm(Me)
                Call StorageListing(strAppend)
            End If
            EditFlag = False
        End If
    End If
End Sub

Private Sub SaveStore2()
Dim RecordSaved
Dim SaveSTORAGE As New ADODB.Recordset
Dim strsql As String
Dim sql2 As String
Dim Store As New ADODB.Recordset
Dim CLM As New ADODB.Recordset
Dim CARTON As New ADODB.Recordset
Dim strAppend As String
Dim listrecord As Integer
Dim listloop As Integer

    If Lststr.ListItems.Count = 0 Then
        MsgBox "Please Enter Case Number", vbOKOnly + vbCritical, DialogBoxTitle
        TxtstrCasNo.SetFocus
    ElseIf Trim(CboCarton) = "" Then
        MsgBox "Please Enter Carton Number", vbOKOnly + vbCritical, DialogBoxTitle
        CboCarton.SetFocus
    ElseIf Trim(Mid(CboFileStatus, 1, 1)) = "" Then
        MsgBox "Please Enter File Status", vbOKOnly + vbCritical, DialogBoxTitle
        CboFileStatus.SetFocus
    ElseIf cboUser <> "" And MaskstrDORqst = "__/__/____" Then
        MsgBox "Please Enter Date Of Request", vbOKOnly + vbCritical, DialogBoxTitle
        MaskstrDORqst.SetFocus
    ElseIf cboUser <> "" And Format(MaskstrDODelivery, "YYYYMMDD") < Format(MaskstrDORqst, "YYYYMMDD") Then
        MsgBox "Invalid Date ! Date of Delivery must greater than Date of Request ", vbOKOnly + vbCritical, DialogBoxTitle
        MaskstrDODelivery = "__/__/____"
        MaskstrDODelivery.SetFocus
    ElseIf cboUser <> "" And Format(MaskstrDOReturn, "YYYYMMDD") < Format(MaskstrDODelivery, "YYYYMMDD") Then
        MsgBox "Invalid Date ! Date of Return must greater than Date of Delivery ", vbOKOnly + vbCritical, DialogBoxTitle
        MaskstrDOReturn = "__/__/____"
        MaskstrDOReturn.SetFocus
    Else
    
        RecordSaved = MsgBox("Do you want to save the record?", vbYesNo + vbExclamation)
            If RecordSaved = vbYes Then
            
                Screen.MousePointer = vbHourglass
                  Call UpdateRecord
                Screen.MousePointer = vbDefault
                TxtstrCasNo.SetFocus
                'Command2.Default = True
                Lststr.ListItems.Clear
            End If
    End If
End Sub

Private Sub UpdateRecord()
Dim SaveSTORAGE As New ADODB.Recordset
Dim listrecord As Integer
Dim listloop As Integer
Dim RecordUpdate
    
    listrecord = Lststr.ListItems.Count
                        
    SaveSTORAGE.Open "STORAGE", medixmasterconnWS, adOpenKeyset, adLockOptimistic
    
    For listloop = 1 To listrecord
    
    SaveSTORAGE.AddNew
    
    With SaveSTORAGE
        !StrCaseNo = ChkStr(Lststr.ListItems(listloop).Text)
        !STRCartonNo = ChkStr(CboCarton)
        !STRFileStatus = Trim(Mid(CboFileStatus, 1, 1))
        !STRRequestBy = ChkStr(cboUser)
        If cboUser <> "" Then
            !DORequest = MaskstrDORqst
            If MaskstrDODelivery <> "__/__/____" Then
                !DODelivery = MaskstrDODelivery
            Else
                !DODelivery = Null
            End If
            If MaskstrDOReturn <> "__/__/____" Then
                !DOReturn = MaskstrDOReturn
            Else
                !DOReturn = Null
            End If
        Else
            !DODelivery = Null
            !DOReturn = Null
            !DORequest = Null
        End If
        !STRRemarks = ChkStr(TxtstrRemarks)
        !LastUpdateBy = Logon
        !LastUpdateDate = Date
        .Update
    End With
    Next listloop
    SaveSTORAGE.Close
    Set SaveSTORAGE = Nothing
    RecordUpdate = MsgBox("Record updated...", vbOKOnly + vbInformation)
        
End Sub

Private Sub CmdUpload_Click()
Dim listrecord As Integer
    
    
    Screen.MousePointer = vbHourglass
    CmdUpload.Enabled = False
    TxtstrCasNo.SetFocus
    'medixmasterconn.Open "File Name=" & BOSPath & "Conn\mediconnect.UDL"
    'medixmasterconnWS.Open "File Name=" & BOSPath & "Conn\mediconnectWS.UDL"

    listrecord = Lststr.ListItems.Count
    If listrecord <> 0 Then
        
        For listloop1 = 1 To listrecord
            Call CaseNoCheck
        Next listloop1
    End If
        If Check = True Then
            MsgBox "Data has been uploaded!", vbOKOnly + vbCritical
        ElseIf Check = False Then
            Call SearchRecord
        End If
        Check = False
    'medixmasterconn.Close
    'Set medixmasterconn = Nothing
    'medixmasterconnWS.Close
    'Set medixmasterconnWS = Nothing
    CmdUpload.Enabled = True
    TxtstrCasNo = ""
    Screen.MousePointer = Default
    
End Sub

Private Function SearchRecord()
Dim strAppend As String
    
    strAppend = ""
    
    If TxtstrCasNo = "" Then
        MsgBox "Please Enter Case No", vbOKOnly + vbCritical
        TxtstrCasNo.SetFocus
    Else
        If Len(Trim(TxtstrCasNo)) Then
            strAppend = strAppend + " And CASNumber = '" & Trim(TxtstrCasNo) & "'"
        End If
        
        Call CaseListing(strAppend)
    End If

End Function

Private Sub CaseListing(Optional strAppend As String)
Dim rst2 As New ADODB.Recordset
Dim rst3 As New ADODB.Recordset
Dim Record As ListItem
Dim sql As String
Dim sqlStorage As String
    
    sql = " Select CASNumber From CAS WHERE 0 = 0 "
    
    If Len(Trim(strAppend)) Then
        sql = sql + strAppend
    End If

    rst2.Open sql, medixmasterconn, adOpenKeyset, adLockOptimistic
    
    If rst2.recordCount = 0 Then
        MsgBox "No record found!", vbOKOnly + vbCritical
    Else
        
        sqlStorage = " Select STRCaseNo From Storage WHERE STRCaseNo = '" & Trim(TxtstrCasNo) & "'"
        
        rst3.Open sqlStorage, medixmasterconnWS, adOpenKeyset, adLockOptimistic
        
        If rst3.recordCount > 0 Then
            MsgBox "Data has been saved!", vbOKOnly + vbCritical
        Else
            Do Until rst2.EOF
                     
                With rst2
                    If Len(!CasNumber) Then
                        Set Record = Lststr.ListItems.Add(, , !CasNumber)
                    End If
                End With
                rst2.MoveNext
            Loop
        rst3.Close
        Set rst3 = Nothing
        End If
    End If
    intExit = 0
    rst2.Close
    Set rst2 = Nothing
    'rst3.Close
    'Set rst3 = Nothing
End Sub

Private Function CaseNoCheck()
    If Trim(ChkStr(TxtstrCasNo)) = ChkStr(Lststr.ListItems(listloop1).Text) Then
        Check = True
    End If
End Function

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If Shift = 0 Then
        Select Case KeyCode
            Case vbKeyF6
                KeyCode = 0
                PrintScreenSnapShot
        End Select
    End If
End Sub

Private Sub Form_Load()
    intExit = 0
    Check = False
    Label15.Visible = False
    TxtstrCasNo2.Visible = False
    ChkScan.Left = 3480
    ChkScan.Top = 380
    
    Select Case SSTab1.Tab
        Case 0
            CmdUpload.Default = False
        Case 1
            CmdUpload.Default = True
    End Select
    
    CmdAdd.Enabled = True
    CmdEdit.Enabled = False
    CmdDelete.Enabled = False
    CmdSave.Enabled = False
    AddFlag = False
   ' medixmasterconn.Open "File Name=" & BOSPath & "Conn\mediconnect.UDL"
   ' medixmasterconnWS.Open "File Name=" & BOSPath & "Conn\mediconnectWS.UDL"
    Call ComboListing
   ' medixmasterconn.Close
   ' Set medixmasterconn = Nothing
   ' medixmasterconnWS.Close
    'Set medixmasterconnWS = Nothing
End Sub

Private Sub Lstcrt_Click()
Dim rst3 As New ADODB.Recordset
Dim strsql As String

    If Lstcrt.ListItems.Count <> 0 Then
        Call EnableCmdButton(Me)
        Call ClearForm(Me)
        strAppend = ""
        If Len(Trim(Lstcrt.SelectedItem.Text)) Then
            strAppend = strAppend + " where CRTCartonNO = '" & RemStr(Trim(Lstcrt.SelectedItem.Text)) & "'"
        End If
   
        strsql = " Select CRTCartonNO, DOSend, DORequest, DODelivery, DOReturn,  " & _
              " DODestroy, CRTCartonStatus, CRTRemarks from Carton " & strAppend
        rst3.Open strsql, medixmasterconnWS, adOpenKeyset, adLockOptimistic
    
        TxtcrtCartonNo = rst3!CRTCartonNo
        If Len(rst3!CRTCartonStatus) Then
            CboCartonStatus = rst3!CRTCartonStatus
        End If
        If Len(rst3!dodestroy) Then
            MaskcrtDODestroy = Format(rst3!dodestroy, "dd/mm/yyyy")
        End If
        If Len(rst3!DORequest) Then
            MaskcrtDORqst = Format(rst3!DORequest, "dd/mm/yyyy")
        End If
        If Len(rst3!DoSend) Then
            MaskcrtDOSend = Format(rst3!DoSend, "dd/mm/yyyy")
        End If
        If Len(rst3!DODelivery) Then
            MaskcrtDODelivery = Format(rst3!DODelivery, "dd/mm/yyyy")
        End If
        If Len(rst3!DOReturn) Then
            MaskcrtDOReturn = Format(rst3!DOReturn, "dd/mm/yyyy")
        End If
        If Len(rst3!CRTRemarks) Then
            TxtcrtRemarks = rst3!CRTRemarks
        End If
        
        Call EntriesEnable(False)
        CmdSave.Enabled = False
        rst3.Close
        Set rst3 = Nothing
    End If
End Sub

Private Sub Lstcrt_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
Lstcrt.SortKey = ColumnHeader.Index - 1
Lstcrt.Sorted = True
If Lstcrt.SortOrder = lvwAscending Then
    Lstcrt.SortOrder = lvwDescending
Else
    Lstcrt.SortOrder = lvwAscending
End If
End Sub

Private Sub Lstcrt_KeyDown(KeyCode As Integer, Shift As Integer)
Dim rst3 As New ADODB.Recordset
Dim strsql As String

    If Lstcrt.ListItems.Count <> 0 Then
        Call EnableCmdButton(Me)
        Call ClearForm(Me)
    strAppend = ""
    If Len(Trim(Lstcrt.SelectedItem.Text)) Then
        strAppend = strAppend + " where CRTCartonNO = '" & RemStr(Trim(Lstcrt.SelectedItem.Text)) & "'"
    End If
    strsql = " Select CRTCartonNO, DOSend, DORequest, DODelivery, DOReturn  " & _
          " DODestroy, CRTCartonStatus, CRTRemarks from Carton " & strAppend
    rst3.Open strsql, medixmasterconnWS, adOpenKeyset, adLockOptimistic
    
       TxtcrtCartonNo = rst3!CRTCartonNo
        If Len(rst3!CRTCartonStatus) Then
            CboCartonStatus = rst3!CRTCartonStatus
        End If
        If Len(rst3!dodestroy) Then
            MaskcrtDODestroy = rst3!dodestroy
        End If
        If Len(rst3!DORequest) Then
            MaskcrtDORqst = rst3!DORequest
        End If
        If Len(rst3!DoSend) Then
            MaskcrtDOSend = rst3!DoSend
        End If
        If Len(rst3!DODelivery) Then
            MaskcrtDODelivery = rst3!DODelivery
        End If
        If Len(rst3!DOReturn) Then
            MaskcrtDOReturn = rst3!DOReturn
        End If
        If Len(rst3!CRTRemarks) Then
            txtctrRemarks = rst3!CRTRemarks
        End If
        Call EntriesEnable(False)
        CmdSave.Enabled = False
        rst3.Close
        Set rst3 = Nothing
    End If
End Sub

Private Sub Lstcrt_KeyUp(KeyCode As Integer, Shift As Integer)
Dim rst3 As New ADODB.Recordset
Dim strsql As String

    If Lstcrt.ListItems.Count <> 0 Then
        Call EnableCmdButton(Me)
        Call ClearForm(Me)
    strAppend = ""
    If Len(Trim(Lstcrt.SelectedItem.Text)) Then
        strAppend = strAppend + " where CRTCartonNO = '" & RemStr(Trim(Lstcrt.SelectedItem.Text)) & "'"
    End If
   
    strsql = " Select CRTCartonNO, DOSend, DORequest, DODelivery, DOReturn  " & _
          " DODestroy, CRTCartonStatus, CRTRemarks from Carton " & strAppend
    rst3.Open strsql, medixmasterconnWS, adOpenKeyset, adLockOptimistic
   
       TxtcrtCartonNo = rst3!CRTCartonNo
        If Len(rst3!CRTCartonStatus) Then
            CboCartonStatus = rst3!CRTCartonStatus
        End If
        If Len(rst3!dodestroy) Then
            MaskcrtDODestroy = rst3!dodestroy
        End If
        If Len(rst3!DORequest) Then
            MaskcrtDORqst = rst3!DORequest
        End If
        If Len(rst3!DoSend) Then
            MaskcrtDOSend = rst3!DoSend
        End If
        If Len(rst3!DODelivery) Then
            MaskcrtDODelivery = rst3!DODelivery
        End If
        If Len(rst3!DOReturn) Then
            MaskcrtDOReturn = rst3!DOReturn
        End If
        If Len(rst3!CRTRemarks) Then
            txtctrRemarks = rst3!CRTRemarks
        End If
        Call EntriesEnable(False)
        CmdSave.Enabled = False
        rst3.Close
        Set rst3 = Nothing
    End If
End Sub

Private Sub Lststr_Click()
Dim rst3 As New ADODB.Recordset
Dim strsql As String
Dim CaseNo As String

    If ChkScan.Value = 0 Then
    
    If Lststr.ListItems.Count <> 0 Then
        Call EnableCmdButton(Me)
        Call ClearForm(Me)
    If Len(Trim(Lststr.SelectedItem.Text)) Then
        strAppend = strAppend + " where strcaseno = '" & RemStr(Trim(Lststr.SelectedItem.Text)) & "'"
    End If
    
'    If Len(Trim(TxtstrClmNo)) Then
'        strappend = strappend + " AND strclmversion = '" & RemStr(Trim(Lststr.SelectedItem.SubItems(2))) & "'"
'    End If
        
    strsql = " Select STRCaseNo, STRCLMVersion, STRCartonNo, STRRequestBy, DORequest,  " & _
          " DODelivery, DOReturn, STRRemarks, STRFileStatus from STORAGE " & strAppend
    rst3.Open strsql, medixmasterconnWS, adOpenKeyset, adLockOptimistic
    
        TxtstrCasNo = rst3!StrCaseNo
'        TxtstrClmNo = rst3!STRClmVersion
        If Len(rst3!STRCartonNo) Then
            CboCarton = rst3!STRCartonNo
        End If
        If Len(rst3!STRFileStatus) Then
            CboFileStatus = rst3!STRFileStatus
        End If
        If Len(rst3!STRRequestBy) Then
            cboUser = rst3!STRRequestBy
        End If
        If Len(rst3!DORequest) Then
            MaskstrDORqst = rst3!DORequest
        End If
        If Len(rst3!DODelivery) Then
            MaskstrDODelivery = rst3!DODelivery
        End If
        If Len(rst3!DOReturn) Then
            MaskstrDOReturn = rst3!DOReturn
        End If
        If Len(rst3!STRRemarks) Then
            TxtstrRemarks = rst3!STRRemarks
        End If
        Call EntriesEnable(False)
        CmdSave.Enabled = False
        rst3.Close
        Set rst3 = Nothing
    End If
    End If
End Sub

Private Sub Lststr_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
Lststr.SortKey = ColumnHeader.Index - 1
Lststr.Sorted = True
If Lststr.SortOrder = lvwAscending Then
    Lststr.SortOrder = lvwDescending
Else
    Lststr.SortOrder = lvwAscending
End If
End Sub

Private Sub MaskcrtDODelivery_LostFocus()
If IsDate(MaskcrtDODelivery) Then
    If EditFlag = True Or AddFlag = True Then
        If Format(MaskcrtDODelivery, "YYYYMMDD") < Format(MaskcrtDORqst, "YYYYMMDD") Then
            MsgBox "Delivery date must be GREATER than  Date Request ! ", vbCritical
'            MaskcrtDODelivery = "__/__/____"
            MaskcrtDODelivery.SetFocus
        End If
    End If
Else
    If MaskcrtDODelivery <> "__/__/____" Then
        MsgBox "Invalid Date Format! ", vbCritical
        MaskcrtDODelivery = "__/__/____"
        MaskcrtDODelivery.SetFocus
    End If
End If
End Sub

Private Sub MaskcrtDODestroy_LostFocus()
If IsDate(MaskcrtDODestroy) Then
    If EditFlag = True Or AddFlag = True Then
        If Format(MaskcrtDODestroy, "YYYYMMDD") < Format(MaskcrtDOReturn, "YYYYMMDD") Then
            MsgBox "Destroy Date must be GREATER than Return Date! ", vbCritical
            MaskcrtDODestroy.SetFocus
        End If
    End If
    Else
        If MaskcrtDODestroy <> "__/__/____" Then
            MsgBox "Invalid Date Format! ", vbCritical
            MaskcrtDODestroy = "__/__/____"
            MaskcrtDODestroy.SetFocus
        End If
End If
End Sub

Private Sub MaskcrtDOReturn_LostFocus()
If IsDate(MaskcrtDOReturn) Then
    If EditFlag = True Or AddFlag = True Then
        If Format(MaskcrtDOReturn, "YYYYMMDD") < Format(MaskcrtDODelivery, "YYYYMMDD") Then
            MsgBox "Return Date must be GREATER than Delivery Date! ", vbCritical
'            MaskcrtDOReturn = "__/__/____"
            MaskcrtDOReturn.SetFocus
        End If
    End If
Else
    If MaskcrtDOReturn <> "__/__/____" Then
        MsgBox "Invalid Date Format! ", vbCritical
        MaskcrtDOReturn = "__/__/____"
        MaskcrtDOReturn.SetFocus
    End If
End If
End Sub

Private Sub MaskcrtDORqst_LostFocus()
If IsDate(MaskcrtDORqst) Then
'        If MaskcrtDORqst <> "__/__/____" And (MaskcrtDOSend = "__/__/____" Or Format(MaskcrtDORqst, "DDMMYYYY") < Format(MaskcrtDOSend, "DDMMYYYY")) Then
    If EditFlag = True Or AddFlag = True Then
        If Format(MaskcrtDORqst, "YYYYMMDD") < Format(MaskcrtDOSend, "YYYYMMDD") Then
            MsgBox "Request Date must be GREATER than Date Send ! ", vbCritical
'            MaskcrtDORqst = "__/__/____"
            MaskcrtDORqst.SetFocus
        End If
    End If
Else
    If MaskcrtDORqst <> "__/__/____" Then
       MsgBox "Invalid Date Format! ", vbCritical
       MaskcrtDORqst = "__/__/____"
       MaskcrtDORqst.SetFocus
    End If
End If


'If IsDate(MaskcrtDORqst) Then
'    Else
'        If MaskcrtDORqst <> "__/__/____" Then
'            MsgBox "Invalid Date Format! ", vbCritical
'            MaskcrtDORqst = "__/__/____"
'            MaskcrtDORqst.SetFocus
'        End If
'End If
End Sub

Private Sub MaskcrtDOSend_LostFocus()
If IsDate(MaskcrtDOSend) Then
    Else
        If MaskcrtDOSend <> "__/__/____" Then
            MsgBox "Invalid Date Format! ", vbCritical
            MaskcrtDOSend = "__/__/____"
            MaskcrtDOSend.SetFocus
        End If
End If
End Sub

Private Sub MaskstrDODelivery_LostFocus()
If IsDate(MaskstrDODelivery) Then
    If EditFlag = True Or AddFlag = True Then
        If Format(MaskstrDODelivery, "YYYYMMDD") < Format(MaskstrDORqst, "YYYYMMDD") Then
            MsgBox "Delivery Date must Greater than Request Date ! ", vbCritical
'           MaskstrDODelivery = "__/__/____"
            MaskstrDODelivery.SetFocus
        End If
    End If
Else
    If MaskstrDODelivery <> "__/__/____" Then
        MsgBox "Invalid Date Format! ", vbCritical
        MaskstrDODelivery = "__/__/____"
        MaskstrDODelivery.SetFocus
    End If
End If
End Sub

Private Sub MaskstrDOReturn_LostFocus()
If IsDate(MaskstrDOReturn) Then
    If EditFlag = True Or AddFlag = True Then
        If Format(MaskstrDOReturn, "YYYYMMDD") < Format(MaskstrDODelivery, "YYYYMMDD") Then
            MsgBox "Return Date must Greater than Delviery Date ! ", vbCritical
'            MaskstrDOReturn = "__/__/____"
            MaskstrDOReturn.SetFocus
        End If
    End If
Else
    If MaskstrDOReturn <> "__/__/____" Then
        MsgBox "Invalid Date Format! ", vbCritical
        MaskstrDOReturn = "__/__/____"
        MaskstrDOReturn.SetFocus
    End If
End If
End Sub

Private Sub MaskstrDORqst_LostFocus()
If IsDate(MaskstrDORqst) Then
    Else
        If MaskstrDORqst <> "__/__/____" Then
            MsgBox "Invalid Date Format! ", vbCritical
            MaskstrDORqst = "__/__/____"
            MaskstrDORqst.SetFocus
        End If
End If
End Sub


Private Sub SSTab1_Click(PreviousTab As Integer)
 Screen.MousePointer = vbHourglass
    Select Case SSTab1.Tab
        Case 0
            Call ClearForm(Me)
            CmdSave.Default = False
            Lstcrt.ListItems.Clear
        Case 1
            ChkScan.Value = 1
            'Lststr2.Visible = True
            TxtstrCasNo.SetFocus
            CmdUpload.Default = True
            TxtstrCasNo.Text = ""
            TxtstrCasNo2.Text = ""
            CboCarton.Text = ""
            CboFileStatus.Text = ""
            cboUser.Text = ""
            TxtstrRemarks.Text = ""
            MaskstrDORqst.Text = "__/__/____"
            MaskstrDODelivery.Text = "__/__/____"
            MaskstrDOReturn.Text = "__/__/____"
            Lststr.ListItems.Clear
            'Lststr2.listitems.Clear
    End Select
    'Call ClearForm(Me)
    Call ComboListing
    Screen.MousePointer = Default
End Sub

Private Sub TxtcrtCartonNo_LostFocus()
    TxtcrtCartonNo = ChkStr(TxtcrtCartonNo)
End Sub

Private Sub TxtcrtRemarks_LostFocus()
    TxtcrtRemarks = ChkStr(TxtcrtRemarks)
End Sub

Private Sub TxtstrCasNo_LostFocus()
    TxtstrCasNo = ChkStr(TxtstrCasNo)
End Sub

Private Sub TxtstrClmNo_KeyPress(KeyAscii As Integer)
    KeyAscii = KeyNumber(KeyAscii)
End Sub

Private Sub EntriesEnable(status As Boolean)
    TxtstrCasNo.Enabled = status
'    TxtstrClmNo.Enabled = status
    CboCarton.Enabled = status
    CboFileStatus.Enabled = status
    cboUser.Enabled = status
    MaskstrDORqst.Enabled = status
    MaskstrDODelivery.Enabled = status
    MaskstrDOReturn.Enabled = status
    TxtstrRemarks.Enabled = status
    TxtcrtCartonNo.Enabled = status
    CboCartonStatus.Enabled = status
    MaskcrtDODestroy.Enabled = status
    MaskcrtDOSend.Enabled = status
    MaskcrtDORqst.Enabled = status
    MaskcrtDODelivery.Enabled = status
    MaskcrtDOReturn.Enabled = status
    TxtcrtRemarks.Enabled = status
End Sub

Private Sub ComboListing()
Dim USRRec As New ADODB.Recordset
Dim CRTRec As New ADODB.Recordset
    
    CboFileStatus.Clear
    CboFileStatus.AddItem "I - In"
    CboFileStatus.AddItem "O - Out"
    CboCartonStatus.Clear
    CboCartonStatus.AddItem "M - Medix"
    CboCartonStatus.AddItem "S - Store"
    CboCartonStatus.AddItem "R - Return"
    
    CRTRec.Open "CARTON", medixmasterconnWS, adOpenKeyset, adLockOptimistic
    CboCarton.Clear
    Do Until CRTRec.EOF
          CboCarton.AddItem CRTRec!CRTCartonNo
          CRTRec.MoveNext
    Loop
    
    USRRec.Open "USR", medixmasterconn, adOpenKeyset, adLockOptimistic
    cboUser.Clear
    Do Until USRRec.EOF
          cboUser.AddItem USRRec!USRName
          USRRec.MoveNext
    Loop
End Sub

Private Sub TxtstrRemarks_LostFocus()
    TxtstrRemarks = ChkStr(TxtstrRemarks)
End Sub

