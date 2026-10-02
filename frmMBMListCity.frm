VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmMBMListCity 
   Caption         =   "List of City"
   ClientHeight    =   5415
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4785
   LinkTopic       =   "Form1"
   ScaleHeight     =   5415
   ScaleWidth      =   4785
   Begin VB.CommandButton cmdExit 
      Cancel          =   -1  'True
      Caption         =   "Exit"
      Height          =   255
      Left            =   3480
      TabIndex        =   1
      Top             =   5160
      Width           =   1095
   End
   Begin MSComctlLib.ListView lstCityList 
      Height          =   5160
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   4680
      _ExtentX        =   8255
      _ExtentY        =   9102
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
         Text            =   "City Code"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Descriptions"
         Object.Width           =   3528
      EndProperty
   End
End
Attribute VB_Name = "frmMBMListCity"
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
    If Len(Trim(ctl.cboCTYCode)) Then
        Call CityListing(ctl.cboCTYCode)
    Else
        Call CityListing
    End If
End Sub


Private Sub CityListing(Optional strCity As String)
    Dim rst2 As New ADODB.Recordset
    Dim CityMaster As ListItem
    Dim sqlCity As String
    Dim lenStrCity As Integer
    
    If Len(Trim(strCity)) Then
        lenStrCity = Len(Trim(strCity))
    End If
    
    sqlCity = "select CTYCOde, CTYDescription  from  CTY  "
   
    lstCityList.ListItems.Clear
        
    rst2.Open sqlCity, medixmasterconn, adOpenKeyset, adLockOptimistic
    
    Do Until rst2.EOF
        Set CityMaster = lstCityList.ListItems.Add(, , rst2!CTYCode)
            CityMaster.SubItems(1) = rst2!CTYDescription
        If Mid(ChkStr(rst2!CTYCode), 1, lenStrCity) = ChkStr(strCity) Then
            CityMaster.Selected = True
        End If
        rst2.MoveNext
    Loop
    rst2.Close
    Set rst2 = Nothing
End Sub


Private Sub lstCityList_DblClick()
    ctl.cboCTYCode = lstCityList.SelectedItem.SubItems(1)
    Unload Me
End Sub
Private Sub cmdExit_Click()
    Unload Me
End Sub

