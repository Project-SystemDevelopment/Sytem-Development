VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmMBMListBranch 
   Caption         =   "List of Branch"
   ClientHeight    =   5115
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   3885
   LinkTopic       =   "Form1"
   ScaleHeight     =   5115
   ScaleWidth      =   3885
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdExit 
      Cancel          =   -1  'True
      Caption         =   "Exit"
      Height          =   315
      Left            =   2760
      TabIndex        =   1
      Top             =   4800
      Width           =   855
   End
   Begin MSComctlLib.ListView lstBranchList 
      Height          =   4800
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   3840
      _ExtentX        =   6773
      _ExtentY        =   8467
      View            =   3
      LabelEdit       =   1
      Sorted          =   -1  'True
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      FullRowSelect   =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      NumItems        =   2
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "State Code"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Descriptions"
         Object.Width           =   8819
      EndProperty
   End
End
Attribute VB_Name = "frmMBMListBranch"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' To determine what is the source
Public Source As Integer
Public intMBMRegistration As Integer
Public intMBMAdjustment As Integer
Public ctl As Form


Private Sub Form_Load()
    intMBMRegistration = 1
    intMBMAdjustment = 2
     
    Call CenterForm(Me)
    
    If Source = intMBMRegistration Then
       Set ctl = frmMemberRegistration
    Else
       Set ctl = frmAdjustment
    End If
    
    If Len(Trim(ctl.cboBRNCode)) Then
        Call branchListing(ctl.cboBRNCode)
    Else
        Call branchListing
    End If
End Sub


Private Sub branchListing(Optional strbranch As String)
    Dim rst2 As New ADODB.Recordset
    Dim branchMaster As ListItem
    
    Dim lenStrbranch As Integer
    
    If Len(Trim(strbranch)) Then
        lenStrbranch = Len(Trim(strbranch))
    End If
    SQLbranch = "select BRNCOde, BRNName from  BRN  "
   
    lstBranchList.ListItems.Clear
    rst2.Open "BRN", medixmasterconn, adOpenKeyset, adLockOptimistic
        Do Until rst2.EOF
            Set branchMaster = lstBranchList.ListItems.Add(, , rst2!BRNCode)
                branchMaster.SubItems(1) = rst2!BRNName
            If Mid(ChkStr(rst2!BRNCode), 1, lenStrbranch) = ChkStr(strbranch) Then
              branchMaster.Selected = True
            End If
            rst2.MoveNext
        Loop
    rst2.Close
    Set rst2 = Nothing
End Sub

Private Sub lstBranchList_DblClick()
    ctl.cboBRNCode = lstBranchList.SelectedItem.SubItems(1)
    Unload Me
End Sub

Private Sub cmdExit_Click()
    Unload Me
End Sub

