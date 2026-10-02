VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form FrmDate 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Please Enter Date"
   ClientHeight    =   1800
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   2760
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1800
   ScaleWidth      =   2760
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame Frame2 
      Height          =   735
      Left            =   120
      TabIndex        =   3
      Top             =   960
      Width           =   2535
      Begin VB.CommandButton CmdClear 
         Caption         =   "&Update"
         Height          =   375
         Left            =   360
         TabIndex        =   5
         Top             =   240
         Width           =   855
      End
      Begin VB.CommandButton CmdExit 
         Cancel          =   -1  'True
         Caption         =   "&Cancel"
         Height          =   375
         Left            =   1320
         TabIndex        =   4
         Top             =   240
         Width           =   855
      End
   End
   Begin VB.Frame Frame1 
      Height          =   855
      Left            =   120
      TabIndex        =   0
      Top             =   0
      Width           =   2535
      Begin MSMask.MaskEdBox txtDate 
         Height          =   255
         Left            =   960
         TabIndex        =   1
         Top             =   360
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   450
         _Version        =   393216
         MaxLength       =   10
         Mask            =   "##/##/####"
         PromptChar      =   "_"
      End
      Begin VB.Label Label1 
         Caption         =   "Date"
         Height          =   255
         Left            =   240
         TabIndex        =   2
         Top             =   360
         Width           =   615
      End
   End
End
Attribute VB_Name = "FrmDate"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub CmdExit_Click()
    Dim RecordExit
    RecordExit = MsgBox("Are you sure?", vbYesNo + vbExclamation)
    If RecordExit = vbYes Then
        bExit = True
        Unload Me
    End If
End Sub
