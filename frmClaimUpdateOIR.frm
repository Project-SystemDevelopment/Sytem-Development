VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmClaimUpdateOIR 
   Caption         =   "Original Invoice Update"
   ClientHeight    =   4455
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   9165
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   4455
   ScaleWidth      =   9165
   WindowState     =   2  'Maximized
   Begin VB.Frame Frame2 
      Height          =   615
      Left            =   120
      TabIndex        =   8
      Top             =   3840
      Width           =   8895
      Begin VB.CommandButton cmpUpdate 
         Caption         =   "Update"
         Height          =   350
         Left            =   7080
         TabIndex        =   3
         Top             =   200
         Width           =   855
      End
      Begin VB.CommandButton cmdExit 
         Caption         =   "Exit"
         Height          =   350
         Left            =   7920
         TabIndex        =   4
         Top             =   200
         Width           =   855
      End
      Begin MSMask.MaskEdBox txtOIRDate 
         Height          =   315
         Left            =   960
         TabIndex        =   2
         Top             =   195
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   556
         _Version        =   393216
         MaxLength       =   10
         Mask            =   "##/##/####"
         PromptChar      =   "_"
      End
      Begin VB.Label Label17 
         Caption         =   "OIR Date"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   1
         Left            =   120
         TabIndex        =   9
         Top             =   225
         Width           =   1455
      End
   End
   Begin VB.Frame Frame1 
      Height          =   540
      Left            =   120
      TabIndex        =   6
      Top             =   260
      Width           =   8895
      Begin VB.CommandButton cmdGO 
         Caption         =   "GO"
         Default         =   -1  'True
         Height          =   300
         Left            =   3360
         TabIndex        =   1
         Top             =   165
         Width           =   615
      End
      Begin VB.TextBox txtWSNoFr 
         Height          =   285
         Left            =   1320
         MaxLength       =   20
         TabIndex        =   0
         Top             =   165
         Width           =   1815
      End
      Begin VB.Label Label20 
         Caption         =   "Case Number"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   7
         Top             =   165
         Width           =   1455
      End
   End
   Begin MSComctlLib.ListView lstInvoiceListing 
      Height          =   2760
      Left            =   120
      TabIndex        =   5
      Top             =   1080
      Width           =   8895
      _ExtentX        =   15690
      _ExtentY        =   4868
      View            =   3
      LabelEdit       =   1
      Sorted          =   -1  'True
      MultiSelect     =   -1  'True
      LabelWrap       =   0   'False
      HideSelection   =   0   'False
      AllowReorder    =   -1  'True
      Checkboxes      =   -1  'True
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   6
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Claim No"
         Object.Width           =   2469
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Invoice Number"
         Object.Width           =   2469
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "Invoice Date"
         Object.Width           =   2469
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   3
         Text            =   "Invoice Amount"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   4
         Text            =   "OIR Status"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   5
         Text            =   "OIR Date"
         Object.Width           =   2540
      EndProperty
   End
   Begin VB.Label Label28 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00404040&
      Caption         =   "Original Invoice Update"
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
      TabIndex        =   11
      Top             =   0
      Width           =   9045
   End
   Begin VB.Label Label20 
      Caption         =   "Invoice Listing"
      Height          =   255
      Index           =   1
      Left            =   240
      TabIndex        =   10
      Top             =   840
      Width           =   1455
   End
End
Attribute VB_Name = "frmClaimUpdateOIR"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim one As Boolean
Dim ItemChecked As Boolean

Private Sub cmdExit_Click()
    Unload Me
End Sub

Public Sub cmdGO_Click()
On Error GoTo errRun
Dim strsql As String
Dim rst As New ADODB.Recordset
Dim Record As ListItem

    Screen.MousePointer = vbHourglass
    lstInvoiceListing.ListItems.Clear
  '  medixmasterconnWS.Open "File Name=" & BOSPath & "Conn\mediconnectWS.UDL"
    
    strsql = "Select CASNumber, INVWSversion, InvNo,InvDate, INVAmount, " & _
            "INVOrigRecStatus, INVOrigRecDate from ClaimInvoice where CASNumber = '" & Trim(txtWSNoFr) & "'"
    
    rst.Open strsql, medixmasterconnWS, adOpenForwardOnly, adLockOptimistic
    
    Do While Not rst.EOF
        With rst
            Set Record = lstInvoiceListing.ListItems.Add(, , !CasNumber & "-" & !INVWSversion)
                Record.SubItems(1) = !INVNo
                Record.SubItems(2) = !InvDate
                Record.SubItems(3) = Format(!INVAmount, "#,##0.00")
                If Len(!INVOrigRecStatus) Then
                    Record.SubItems(4) = !INVOrigRecStatus
                End If
                If Len(!INVOrigRecDate) Then
                    Record.SubItems(5) = !INVOrigRecDate
                End If
                
        End With
    rst.MoveNext
    Loop
 '   medixmasterconnWS.Close
 '   Set medixmasterconnWS = Nothing
    Screen.MousePointer = vbDefault
    Exit Sub

errRun:
    MsgBox Err.Description
    Screen.MousePointer = vbDefault
    lstInvoiceListing.ListItems.Clear
  '  medixmasterconnWS.Close
  '  Set medixmasterconnWS = Nothing
End Sub

Private Sub cmpUpdate_Click()
Dim i As Integer
Dim strsql As String
Dim CompTrans As Boolean
Dim CaseNo, caseversion, tmpInvNo As String

    
    If txtOIRDate = "__/__/____" Then
        MsgBox "Please Enter Original Invoice Receipt Date! "
    Else
        For i = 1 To lstInvoiceListing.ListItems.Count
            If lstInvoiceListing.ListItems(i).Checked = True Then
                ItemChecked = True
                Exit For
            End If
        Next i
        
      '  medixmasterconnWS.Open "File Name=" & BOSPath & "Conn\mediconnectWS.UDL"
    
        medixmasterconnWS.beginTrans
        CompTrans = False
        
        For i = 1 To lstInvoiceListing.ListItems.Count
            If lstInvoiceListing.ListItems(i).Checked = True Then
                
                CaseNo = Trim(Mid(lstInvoiceListing.ListItems(i).Text, 1, IIf(InStr(1, lstInvoiceListing.ListItems(i).Text, "-") = 0, 10, InStr(1, lstInvoiceListing.ListItems(i).Text, "-") - 1)))
                caseversion = Trim(Mid(lstInvoiceListing.ListItems(i).Text, InStr(1, lstInvoiceListing.ListItems(i).Text, "-") + 1, Len(lstInvoiceListing.ListItems(i).Text) - InStr(1, lstInvoiceListing.ListItems(i).Text, "-")))
                tmpInvNo = Trim(lstInvoiceListing.ListItems(i).SubItems(1))
                strsql = ""
                strsql = " update ClaimInvoice set " & _
                         " INVOrigRecStatus = 'Y', " & _
                         " INVOrigRecDate = '" & Format(txtOIRDate, "mm/dd/yyyy") & "'" & _
                         " Where CASNumber = '" & Trim(CaseNo) & "' And INVWSversion = '" & Trim(caseversion) & "' And INVNo = '" & Trim(tmpInvNo) & "'"
                medixmasterconnWS.Execute strsql
            End If
            
        Next i
            
            medixmasterconnWS.commitTrans
            CompTrans = True
    
         '   medixmasterconnWS.Close
       '     Set medixmasterconnWS = Nothing
            MsgBox "Record is Updated! "
            Call cmdGO_Click
    End If
    Exit Sub
ErrorSave:
        Screen.MousePointer = vbDefault
        If CompTrans = False Then
            medixmasterconnWS.RollbackTrans
        End If
     ''   medixmasterconnWS.Close
    '    Set medixmasterconnWS = Nothing
        If CompTrans = False Then
            MsgBox Err.Description
        End If
End Sub

Private Sub lstInvoiceListing_Click()
    'ItemCheck = False
    'CmdSave.Enabled = False
    'EditFlag = False
    'Call LvDisplayData(, False)
End Sub
