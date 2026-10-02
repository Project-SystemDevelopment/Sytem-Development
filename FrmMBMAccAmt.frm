VERSION 5.00
Begin VB.Form FrmPaymentSubAmtPiad 
   Caption         =   "Form1"
   ClientHeight    =   2115
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   3735
   LinkTopic       =   "Form1"
   ScaleHeight     =   2115
   ScaleWidth      =   3735
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox txtRecAmt 
      Height          =   285
      Left            =   1560
      MaxLength       =   8
      TabIndex        =   6
      Top             =   480
      Width           =   2175
   End
   Begin VB.Frame Frame1 
      Height          =   975
      Left            =   0
      TabIndex        =   2
      Top             =   1080
      Width           =   3735
      Begin VB.CommandButton CmdUpdate 
         Caption         =   "&Update"
         Default         =   -1  'True
         Height          =   375
         Left            =   840
         TabIndex        =   5
         Top             =   360
         Width           =   975
      End
      Begin VB.CommandButton CmdClear 
         Caption         =   "&Clear"
         Height          =   375
         Left            =   1800
         TabIndex        =   4
         Top             =   360
         Width           =   855
      End
      Begin VB.CommandButton CmdOK 
         Cancel          =   -1  'True
         Caption         =   "&Cancel"
         Height          =   375
         Left            =   2640
         TabIndex        =   3
         Top             =   360
         Width           =   855
      End
   End
   Begin VB.TextBox txtCLMNo 
      Enabled         =   0   'False
      Height          =   285
      Left            =   1560
      TabIndex        =   1
      Top             =   120
      Width           =   2175
   End
   Begin VB.TextBox txtOSAmt 
      Enabled         =   0   'False
      Height          =   285
      Left            =   1560
      TabIndex        =   0
      Top             =   840
      Width           =   2175
   End
   Begin VB.Label Label4 
      Caption         =   "Receipt/Payment"
      Height          =   255
      Left            =   240
      TabIndex        =   9
      Top             =   480
      Width           =   1335
   End
   Begin VB.Label Label5 
      Caption         =   "Amount O/S"
      Height          =   255
      Left            =   240
      TabIndex        =   8
      Top             =   840
      Width           =   1215
   End
   Begin VB.Label Label1 
      Caption         =   "Claim No"
      Height          =   255
      Left            =   240
      TabIndex        =   7
      Top             =   120
      Width           =   1335
   End
End
Attribute VB_Name = "FrmPaymentSubAmtPiad"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
