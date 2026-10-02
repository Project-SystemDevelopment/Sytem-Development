VERSION 5.00
Begin VB.Form frmReport 
   Caption         =   "Report"
   ClientHeight    =   5325
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   7470
   LinkTopic       =   "Form1"
   ScaleHeight     =   5325
   ScaleWidth      =   7470
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.Frame Frame3 
      Caption         =   "Period Interval"
      Height          =   1335
      Left            =   360
      TabIndex        =   6
      Top             =   840
      Width           =   5775
      Begin VB.CommandButton Command1 
         Caption         =   "Generate"
         Height          =   375
         Left            =   4560
         TabIndex        =   15
         Top             =   840
         Width           =   1095
      End
      Begin VB.TextBox txtFromYear 
         Height          =   285
         Left            =   2760
         MaxLength       =   4
         TabIndex        =   14
         Top             =   585
         Width           =   975
      End
      Begin VB.TextBox txtFromMonth 
         Height          =   285
         Left            =   1440
         MaxLength       =   2
         TabIndex        =   13
         Top             =   585
         Width           =   975
      End
      Begin VB.TextBox txtToMonth 
         Height          =   285
         Left            =   1440
         MaxLength       =   2
         TabIndex        =   10
         Top             =   825
         Width           =   975
      End
      Begin VB.TextBox txtToYear 
         Height          =   285
         Left            =   2760
         MaxLength       =   4
         TabIndex        =   9
         Top             =   825
         Width           =   975
      End
      Begin VB.Label Label4 
         Caption         =   "To"
         Height          =   255
         Left            =   480
         TabIndex        =   12
         Top             =   825
         Width           =   1095
      End
      Begin VB.Label Label3 
         Caption         =   "From"
         Height          =   255
         Left            =   480
         TabIndex        =   11
         Top             =   585
         Width           =   1095
      End
      Begin VB.Label Label2 
         Caption         =   "Month"
         Height          =   255
         Left            =   1680
         TabIndex        =   8
         Top             =   240
         Width           =   735
      End
      Begin VB.Label Label1 
         Caption         =   "Year"
         Height          =   255
         Left            =   3000
         TabIndex        =   7
         Top             =   240
         Width           =   615
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Type Of Report"
      Height          =   2295
      Left            =   360
      TabIndex        =   4
      Top             =   2280
      Width           =   5775
      Begin VB.OptionButton OptProduction1 
         Caption         =   "Membership - Detail Report"
         Height          =   375
         Left            =   240
         TabIndex        =   5
         Top             =   240
         Value           =   -1  'True
         Width           =   4215
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Date Type"
      Height          =   735
      Left            =   360
      TabIndex        =   0
      Top             =   0
      Width           =   5775
      Begin VB.OptionButton Option8 
         Caption         =   "By Transaction Date"
         Height          =   255
         Left            =   3840
         TabIndex        =   3
         Top             =   240
         Width           =   1815
      End
      Begin VB.OptionButton Option7 
         Caption         =   "By Effective Date"
         Height          =   255
         Left            =   1920
         TabIndex        =   2
         Top             =   240
         Width           =   1575
      End
      Begin VB.OptionButton Option6 
         Caption         =   "By Bordx Date"
         Height          =   255
         Left            =   240
         TabIndex        =   1
         Top             =   240
         Value           =   -1  'True
         Width           =   1695
      End
   End
End
Attribute VB_Name = "frmReport"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

