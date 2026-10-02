VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmMBMListState 
   Caption         =   "Form1"
   ClientHeight    =   5070
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   3900
   LinkTopic       =   "Form1"
   ScaleHeight     =   5070
   ScaleWidth      =   3900
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdExit 
      Cancel          =   -1  'True
      Caption         =   "Exit"
      Height          =   255
      Left            =   2640
      TabIndex        =   1
      Top             =   4800
      Width           =   1215
   End
   Begin MSComctlLib.ListView lstStateList 
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
Attribute VB_Name = "frmMBMListState"
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
    
    If Source = intMBMRegistration Then
       Set ctl = frmMemberRegistration
    Else
       Set ctl = frmAdjustment
    End If
    
    Call CenterForm(Me)
    If Len(Trim(ctl.cboSTACode)) Then
        Call StateListing(ctl.cboSTACode)
    Else
        Call StateListing
    End If
End Sub


Private Sub StateListing(Optional strState As String)
    Dim rst2 As New ADODB.Recordset
    Dim StateMaster As ListItem
    
    Dim lenStrState As Integer
    
    If Len(Trim(strState)) Then
        lenStrState = Len(Trim(strState))
    End If
    SQLState = "select STACOde, STADescription  from  STA  "
   
    lstStateList.ListItems.Clear
    rst2.Open "STA", medixmasterconn, adOpenKeyset, adLockOptimistic
        Do Until rst2.EOF
            Set StateMaster = lstStateList.ListItems.Add(, , rst2!STACode)
                StateMaster.SubItems(1) = rst2!STADescription
            If Mid(ChkStr(rst2!STACode), 1, lenStrState) = ChkStr(strState) Then
              StateMaster.Selected = True
            End If
            rst2.MoveNext
        Loop
    rst2.Close
    Set rst2 = Nothing
End Sub

Private Sub lstStateList_DblClick()
    ctl.cboSTACode = lstStateList.SelectedItem.SubItems(1)
    Unload Me
End Sub


Private Sub cmdExit_Click()
    Unload Me
End Sub
