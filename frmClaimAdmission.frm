VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmClaimAdmission 
   Caption         =   "Open New Worksheet"
   ClientHeight    =   7995
   ClientLeft      =   -1065
   ClientTop       =   450
   ClientWidth     =   8880
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   7995
   ScaleWidth      =   8880
   WindowState     =   2  'Maximized
   Begin VB.Frame Frame3 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   405
      Left            =   0
      TabIndex        =   8
      Top             =   0
      Width           =   11895
      Begin VB.Frame Frame8 
         Caption         =   "Frame8"
         Height          =   15
         Left            =   240
         TabIndex        =   9
         Top             =   480
         Width           =   7215
      End
      Begin VB.Label Label3 
         Appearance      =   0  'Flat
         BackColor       =   &H00000000&
         Caption         =   "Worksheet"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000005&
         Height          =   315
         Left            =   180
         TabIndex        =   10
         Top             =   60
         Width           =   6615
      End
   End
   Begin VB.Frame Frame2 
      Height          =   615
      Left            =   120
      TabIndex        =   7
      Top             =   360
      Width           =   11775
      Begin VB.ComboBox cboVersion 
         Height          =   315
         Left            =   5880
         Style           =   2  'Dropdown List
         TabIndex        =   32
         Top             =   240
         Width           =   735
      End
      Begin VB.CommandButton cmdSearchCase 
         Caption         =   "&Search"
         Height          =   255
         Left            =   3840
         TabIndex        =   2
         Top             =   240
         Width           =   735
      End
      Begin VB.CommandButton cmdDisplay 
         Caption         =   "&Go"
         Default         =   -1  'True
         Height          =   255
         Left            =   3120
         TabIndex        =   1
         Top             =   240
         Width           =   735
      End
      Begin VB.TextBox txtCASNumber 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   1320
         TabIndex        =   0
         Top             =   240
         Width           =   1740
      End
      Begin VB.Label Label25 
         Caption         =   "WorkSheet No"
         Height          =   255
         Left            =   4680
         TabIndex        =   33
         Top             =   270
         Width           =   1095
      End
      Begin VB.Label Label48 
         Caption         =   "Pol Status"
         Height          =   255
         Left            =   6720
         TabIndex        =   31
         Top             =   240
         Width           =   795
      End
      Begin VB.Label Label9 
         Caption         =   "Case Status"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   9000
         TabIndex        =   30
         Top             =   240
         Width           =   900
      End
      Begin VB.Label lblCASStatus 
         BackColor       =   &H00C0FFFF&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   255
         Left            =   9960
         TabIndex        =   29
         Top             =   240
         Width           =   1575
      End
      Begin VB.Label txtPolStatus 
         BackColor       =   &H00C0FFFF&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   255
         Left            =   7440
         TabIndex        =   28
         Top             =   240
         Width           =   1455
      End
      Begin VB.Label Label46 
         Caption         =   "Case No"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   11
         Top             =   240
         Width           =   1275
      End
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   7575
      Left            =   120
      TabIndex        =   3
      Top             =   960
      Width           =   11655
      _ExtentX        =   20558
      _ExtentY        =   13361
      _Version        =   393216
      Tabs            =   2
      Tab             =   1
      TabsPerRow      =   2
      TabHeight       =   520
      OLEDropMode     =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      TabCaption(0)   =   "&Invoices && Diagnosis"
      TabPicture(0)   =   "frmClaimAdmission.frx":0000
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "CmdInvoice"
      Tab(0).Control(1)=   "Frame16"
      Tab(0).Control(2)=   "cmdExit2"
      Tab(0).Control(3)=   "GeneralFrame"
      Tab(0).ControlCount=   4
      TabCaption(1)   =   "&WorkSheet"
      TabPicture(1)   =   "frmClaimAdmission.frx":001C
      Tab(1).ControlEnabled=   -1  'True
      Tab(1).Control(0)=   "Frame4"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).Control(1)=   "Frame9"
      Tab(1).Control(1).Enabled=   0   'False
      Tab(1).ControlCount=   2
      Begin VB.Frame Frame9 
         Height          =   615
         Left            =   120
         TabIndex        =   131
         Top             =   360
         Width           =   11415
         Begin VB.CommandButton cmdApprove 
            Caption         =   "Approve"
            Height          =   255
            Left            =   240
            TabIndex        =   134
            Top             =   240
            Width           =   855
         End
         Begin VB.CommandButton cmdSave 
            Caption         =   "Save"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   1080
            TabIndex        =   133
            Top             =   240
            Width           =   855
         End
         Begin VB.CommandButton cmdExit 
            Cancel          =   -1  'True
            Caption         =   "Exit"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   1920
            TabIndex        =   132
            Top             =   240
            Width           =   855
         End
         Begin VB.Label Label39 
            Caption         =   "Limit"
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
            Left            =   6480
            TabIndex        =   139
            Top             =   240
            Width           =   555
         End
         Begin VB.Label Label40 
            Caption         =   "Actual"
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
            Left            =   7320
            TabIndex        =   138
            Top             =   240
            Width           =   690
         End
         Begin VB.Label Label41 
            Caption         =   "Approved"
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
            Left            =   8040
            TabIndex        =   137
            Top             =   240
            Width           =   885
         End
         Begin VB.Label Label42 
            Caption         =   "Payable to Hospital"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   6.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   8880
            TabIndex        =   136
            Top             =   120
            Width           =   810
         End
         Begin VB.Label Label43 
            Caption         =   "Due To/(From) Member"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   6.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   480
            Left            =   9840
            TabIndex        =   135
            Top             =   120
            Width           =   1200
         End
      End
      Begin VB.Frame Frame4 
         Height          =   7095
         Left            =   120
         TabIndex        =   56
         Top             =   900
         Width           =   11415
         Begin VB.VScrollBar VScroll2 
            Height          =   6015
            Left            =   10920
            TabIndex        =   57
            Top             =   240
            Width           =   255
         End
         Begin VB.Frame Frame11 
            Height          =   7335
            Left            =   0
            TabIndex        =   58
            Top             =   0
            Width           =   10725
            Begin VB.TextBox P22 
               Alignment       =   1  'Right Justify
               Enabled         =   0   'False
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   285
               Left            =   5620
               TabIndex        =   244
               Top             =   5385
               Width           =   1020
            End
            Begin VB.TextBox Text4 
               BackColor       =   &H80000004&
               BorderStyle     =   0  'None
               Height          =   195
               Left            =   900
               TabIndex        =   227
               Text            =   "Post Hospitalisaion Treatment "
               Top             =   2400
               Width           =   3135
            End
            Begin VB.TextBox P1 
               Alignment       =   1  'Right Justify
               Enabled         =   0   'False
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   285
               Left            =   5620
               TabIndex        =   215
               Top             =   180
               Width           =   1020
            End
            Begin VB.TextBox P2 
               Alignment       =   1  'Right Justify
               Enabled         =   0   'False
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   285
               Left            =   5620
               TabIndex        =   223
               Top             =   405
               Width           =   1020
            End
            Begin VB.TextBox P3 
               Alignment       =   1  'Right Justify
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   285
               Left            =   5620
               TabIndex        =   224
               Top             =   600
               Width           =   1020
            End
            Begin VB.TextBox P4 
               Alignment       =   1  'Right Justify
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   285
               Left            =   5620
               TabIndex        =   225
               Top             =   840
               Width           =   1020
            End
            Begin VB.TextBox Text3 
               BackColor       =   &H80000004&
               BorderStyle     =   0  'None
               Height          =   165
               Left            =   900
               TabIndex        =   226
               Text            =   "Operating Theatre"
               Top             =   840
               Width           =   1815
            End
            Begin VB.TextBox A1 
               Alignment       =   1  'Right Justify
               BeginProperty DataFormat 
                  Type            =   1
                  Format          =   "0.00"
                  HaveTrueFalseNull=   0
                  FirstDayOfWeek  =   0
                  FirstWeekOfYear =   0
                  LCID            =   1033
                  SubFormatType   =   1
               EndProperty
               Enabled         =   0   'False
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   285
               Left            =   6630
               TabIndex        =   218
               Top             =   180
               Width           =   1020
            End
            Begin VB.TextBox A2 
               Alignment       =   1  'Right Justify
               Enabled         =   0   'False
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   285
               Left            =   6630
               TabIndex        =   219
               Top             =   405
               Width           =   1020
            End
            Begin VB.TextBox A3 
               Alignment       =   1  'Right Justify
               Enabled         =   0   'False
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   285
               Left            =   6630
               TabIndex        =   220
               Top             =   630
               Width           =   1020
            End
            Begin VB.TextBox AP1 
               Alignment       =   1  'Right Justify
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   285
               Left            =   7630
               TabIndex        =   70
               Top             =   180
               Width           =   1020
            End
            Begin VB.TextBox AP2 
               Alignment       =   1  'Right Justify
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   285
               Left            =   7630
               TabIndex        =   216
               Top             =   405
               Width           =   1020
            End
            Begin VB.TextBox A4 
               Alignment       =   1  'Right Justify
               Enabled         =   0   'False
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   285
               Left            =   6630
               TabIndex        =   221
               Top             =   840
               Width           =   1020
            End
            Begin VB.TextBox txtAnnualLimit 
               Alignment       =   1  'Right Justify
               Enabled         =   0   'False
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   285
               Left            =   5620
               TabIndex        =   151
               Text            =   "0.00"
               Top             =   6120
               Width           =   1020
            End
            Begin VB.TextBox txtOthers2 
               BeginProperty Font 
                  Name            =   "Times New Roman"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   315
               Left            =   840
               TabIndex        =   144
               Top             =   5700
               Width           =   4695
            End
            Begin VB.TextBox ICU1 
               Alignment       =   1  'Right Justify
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   285
               Left            =   2760
               MaxLength       =   3
               TabIndex        =   106
               Text            =   "0"
               Top             =   405
               Width           =   555
            End
            Begin VB.TextBox LG1 
               Alignment       =   1  'Right Justify
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   285
               Left            =   2760
               MaxLength       =   3
               TabIndex        =   105
               Text            =   "0"
               Top             =   4305
               Width           =   555
            End
            Begin VB.TextBox DB1 
               Alignment       =   1  'Right Justify
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   285
               Left            =   2760
               MaxLength       =   3
               TabIndex        =   104
               Text            =   "0"
               Top             =   180
               Width           =   555
            End
            Begin VB.TextBox LodgerAmount 
               Alignment       =   1  'Right Justify
               Enabled         =   0   'False
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   285
               Left            =   4320
               MaxLength       =   3
               TabIndex        =   65
               Text            =   "0"
               Top             =   4305
               Width           =   675
            End
            Begin VB.TextBox txtOthers1 
               BeginProperty Font 
                  Name            =   "Times New Roman"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   315
               Left            =   840
               TabIndex        =   64
               Top             =   5370
               Width           =   4695
            End
            Begin VB.TextBox ICUAmount 
               Alignment       =   1  'Right Justify
               Enabled         =   0   'False
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   285
               Left            =   4320
               MaxLength       =   3
               TabIndex        =   63
               Text            =   "0"
               Top             =   405
               Width           =   675
            End
            Begin VB.TextBox RoomAmount 
               Alignment       =   1  'Right Justify
               Enabled         =   0   'False
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   285
               Left            =   4320
               MaxLength       =   3
               TabIndex        =   62
               Text            =   "0"
               Top             =   180
               Width           =   675
            End
            Begin VB.TextBox txtDayCash 
               Alignment       =   1  'Right Justify
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   285
               Left            =   2760
               MaxLength       =   3
               TabIndex        =   60
               Text            =   "0"
               Top             =   3840
               Width           =   555
            End
            Begin VB.TextBox txtRMCash 
               Alignment       =   1  'Right Justify
               Enabled         =   0   'False
               Height          =   285
               Left            =   4320
               MaxLength       =   3
               TabIndex        =   59
               Text            =   "0"
               Top             =   3840
               Width           =   675
            End
            Begin VB.TextBox AP3 
               Alignment       =   1  'Right Justify
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   285
               Left            =   7630
               TabIndex        =   217
               Top             =   630
               Width           =   1020
            End
            Begin VB.TextBox AP4 
               Alignment       =   1  'Right Justify
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   285
               Left            =   7630
               TabIndex        =   222
               Top             =   840
               Width           =   1020
            End
            Begin VB.TextBox H1 
               Alignment       =   1  'Right Justify
               Enabled         =   0   'False
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   285
               Left            =   8640
               TabIndex        =   79
               Top             =   180
               Width           =   1020
            End
            Begin VB.TextBox H2 
               Alignment       =   1  'Right Justify
               Enabled         =   0   'False
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   285
               Left            =   8640
               TabIndex        =   78
               Top             =   405
               Width           =   1020
            End
            Begin VB.TextBox H3 
               Alignment       =   1  'Right Justify
               BeginProperty DataFormat 
                  Type            =   1
                  Format          =   "0.00"
                  HaveTrueFalseNull=   0
                  FirstDayOfWeek  =   0
                  FirstWeekOfYear =   0
                  LCID            =   1033
                  SubFormatType   =   1
               EndProperty
               Enabled         =   0   'False
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   285
               Left            =   8640
               TabIndex        =   77
               Top             =   630
               Width           =   1020
            End
            Begin VB.TextBox H4 
               Alignment       =   1  'Right Justify
               Enabled         =   0   'False
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   285
               Left            =   8640
               TabIndex        =   76
               Top             =   840
               Width           =   1020
            End
            Begin VB.TextBox M1 
               Alignment       =   1  'Right Justify
               BeginProperty DataFormat 
                  Type            =   1
                  Format          =   "#,##0.00"
                  HaveTrueFalseNull=   0
                  FirstDayOfWeek  =   0
                  FirstWeekOfYear =   0
                  LCID            =   1033
                  SubFormatType   =   1
               EndProperty
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   285
               Left            =   9640
               TabIndex        =   83
               Top             =   180
               Width           =   1020
            End
            Begin VB.TextBox M2 
               Alignment       =   1  'Right Justify
               BeginProperty DataFormat 
                  Type            =   1
                  Format          =   "#,##0.00"
                  HaveTrueFalseNull=   0
                  FirstDayOfWeek  =   0
                  FirstWeekOfYear =   0
                  LCID            =   1033
                  SubFormatType   =   1
               EndProperty
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   285
               Left            =   9640
               TabIndex        =   82
               Top             =   405
               Width           =   1020
            End
            Begin VB.TextBox M3 
               Alignment       =   1  'Right Justify
               BeginProperty DataFormat 
                  Type            =   1
                  Format          =   "#,##0.00"
                  HaveTrueFalseNull=   0
                  FirstDayOfWeek  =   0
                  FirstWeekOfYear =   0
                  LCID            =   1033
                  SubFormatType   =   1
               EndProperty
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   285
               Left            =   9640
               TabIndex        =   81
               Top             =   630
               Width           =   1020
            End
            Begin VB.TextBox M4 
               Alignment       =   1  'Right Justify
               BeginProperty DataFormat 
                  Type            =   1
                  Format          =   "#,##0.00"
                  HaveTrueFalseNull=   0
                  FirstDayOfWeek  =   0
                  FirstWeekOfYear =   0
                  LCID            =   1033
                  SubFormatType   =   1
               EndProperty
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   285
               Left            =   9640
               TabIndex        =   80
               Top             =   840
               Width           =   1020
            End
            Begin VB.Frame Frame5 
               Height          =   900
               Left            =   0
               TabIndex        =   152
               Top             =   960
               Width           =   10695
               Begin VB.TextBox P5 
                  Alignment       =   1  'Right Justify
                  BeginProperty Font 
                     Name            =   "Tahoma"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   285
                  Left            =   5620
                  TabIndex        =   159
                  Top             =   120
                  Width           =   1020
               End
               Begin VB.TextBox P6 
                  Alignment       =   1  'Right Justify
                  BeginProperty Font 
                     Name            =   "Tahoma"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   285
                  Left            =   5620
                  TabIndex        =   172
                  Top             =   360
                  Width           =   1020
               End
               Begin VB.TextBox P7 
                  Alignment       =   1  'Right Justify
                  BeginProperty Font 
                     Name            =   "Tahoma"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   285
                  Left            =   5620
                  TabIndex        =   162
                  Top             =   600
                  Width           =   1020
               End
               Begin VB.TextBox A5 
                  Alignment       =   1  'Right Justify
                  Enabled         =   0   'False
                  BeginProperty Font 
                     Name            =   "Tahoma"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   285
                  Left            =   6630
                  TabIndex        =   160
                  Top             =   120
                  Width           =   1020
               End
               Begin VB.TextBox A6 
                  Alignment       =   1  'Right Justify
                  Enabled         =   0   'False
                  BeginProperty Font 
                     Name            =   "Tahoma"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   285
                  Left            =   6630
                  TabIndex        =   161
                  Top             =   360
                  Width           =   1020
               End
               Begin VB.TextBox A7 
                  Alignment       =   1  'Right Justify
                  Enabled         =   0   'False
                  BeginProperty Font 
                     Name            =   "Tahoma"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   285
                  Left            =   6630
                  TabIndex        =   163
                  Top             =   600
                  Width           =   1020
               End
               Begin VB.TextBox AP5 
                  Alignment       =   1  'Right Justify
                  BeginProperty Font 
                     Name            =   "Tahoma"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   285
                  Left            =   7630
                  TabIndex        =   164
                  Top             =   120
                  Width           =   1020
               End
               Begin VB.TextBox AP6 
                  Alignment       =   1  'Right Justify
                  BeginProperty Font 
                     Name            =   "Tahoma"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   285
                  Left            =   7630
                  TabIndex        =   165
                  Top             =   360
                  Width           =   1020
               End
               Begin VB.TextBox AP7 
                  Alignment       =   1  'Right Justify
                  BeginProperty Font 
                     Name            =   "Tahoma"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   285
                  Left            =   7630
                  TabIndex        =   166
                  Top             =   600
                  Width           =   1020
               End
               Begin VB.TextBox H5 
                  Alignment       =   1  'Right Justify
                  Enabled         =   0   'False
                  BeginProperty Font 
                     Name            =   "Tahoma"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   285
                  Left            =   8640
                  TabIndex        =   156
                  Top             =   120
                  Width           =   1020
               End
               Begin VB.TextBox H6 
                  Alignment       =   1  'Right Justify
                  Enabled         =   0   'False
                  BeginProperty Font 
                     Name            =   "Tahoma"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   285
                  Left            =   8640
                  TabIndex        =   157
                  Top             =   360
                  Width           =   1020
               End
               Begin VB.TextBox H7 
                  Alignment       =   1  'Right Justify
                  Enabled         =   0   'False
                  BeginProperty Font 
                     Name            =   "Tahoma"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   285
                  Left            =   8640
                  TabIndex        =   158
                  Top             =   600
                  Width           =   1020
               End
               Begin VB.TextBox M5 
                  Alignment       =   1  'Right Justify
                  BeginProperty DataFormat 
                     Type            =   1
                     Format          =   "#,##0.00"
                     HaveTrueFalseNull=   0
                     FirstDayOfWeek  =   0
                     FirstWeekOfYear =   0
                     LCID            =   1033
                     SubFormatType   =   1
                  EndProperty
                  BeginProperty Font 
                     Name            =   "Tahoma"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   285
                  Left            =   9640
                  TabIndex        =   153
                  Top             =   120
                  Width           =   1020
               End
               Begin VB.TextBox M6 
                  Alignment       =   1  'Right Justify
                  BeginProperty DataFormat 
                     Type            =   1
                     Format          =   "#,##0.00"
                     HaveTrueFalseNull=   0
                     FirstDayOfWeek  =   0
                     FirstWeekOfYear =   0
                     LCID            =   1033
                     SubFormatType   =   1
                  EndProperty
                  BeginProperty Font 
                     Name            =   "Tahoma"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   285
                  Left            =   9640
                  TabIndex        =   154
                  Top             =   360
                  Width           =   1020
               End
               Begin VB.TextBox M7 
                  Alignment       =   1  'Right Justify
                  BeginProperty DataFormat 
                     Type            =   1
                     Format          =   "#,##0.00"
                     HaveTrueFalseNull=   0
                     FirstDayOfWeek  =   0
                     FirstWeekOfYear =   0
                     LCID            =   1033
                     SubFormatType   =   1
                  EndProperty
                  BeginProperty Font 
                     Name            =   "Tahoma"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   285
                  Left            =   9640
                  TabIndex        =   155
                  Top             =   600
                  Width           =   1020
               End
               Begin VB.Label Label15 
                  Caption         =   "Surgical Fees"
                  BeginProperty Font 
                     Name            =   "Tahoma"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   255
                  Left            =   900
                  TabIndex        =   169
                  Top             =   360
                  Width           =   2085
               End
               Begin VB.Label Label16 
                  Caption         =   "Pre-Hospital (X-ray && Lab)"
                  BeginProperty Font 
                     Name            =   "Tahoma"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   255
                  Left            =   900
                  TabIndex        =   168
                  Top             =   120
                  Width           =   2655
               End
               Begin VB.Label Label62 
                  Caption         =   "Surgical"
                  Height          =   495
                  Left            =   120
                  TabIndex        =   167
                  Top             =   120
                  Width           =   615
               End
               Begin VB.Label Label14 
                  Caption         =   "Anaesthetic's Fees"
                  BeginProperty Font 
                     Name            =   "Tahoma"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   255
                  Left            =   900
                  TabIndex        =   170
                  Top             =   600
                  Width           =   1635
               End
            End
            Begin VB.TextBox P8 
               Alignment       =   1  'Right Justify
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   285
               Left            =   5620
               TabIndex        =   178
               Top             =   1880
               Width           =   1020
            End
            Begin VB.TextBox P9 
               Alignment       =   1  'Right Justify
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   285
               Left            =   5620
               TabIndex        =   176
               Top             =   2100
               Width           =   1020
            End
            Begin VB.TextBox P10 
               Alignment       =   1  'Right Justify
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   285
               Left            =   5620
               TabIndex        =   174
               Top             =   2330
               Width           =   1020
            End
            Begin VB.TextBox A8 
               Alignment       =   1  'Right Justify
               Enabled         =   0   'False
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   285
               Left            =   6630
               TabIndex        =   177
               Top             =   1880
               Width           =   1020
            End
            Begin VB.TextBox A9 
               Alignment       =   1  'Right Justify
               Enabled         =   0   'False
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   285
               Left            =   6630
               TabIndex        =   175
               Top             =   2100
               Width           =   1020
            End
            Begin VB.TextBox A10 
               Alignment       =   1  'Right Justify
               Enabled         =   0   'False
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   285
               Left            =   6630
               TabIndex        =   173
               Top             =   2330
               Width           =   1020
            End
            Begin VB.TextBox AP8 
               Alignment       =   1  'Right Justify
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   285
               Left            =   7630
               TabIndex        =   187
               Top             =   1880
               Width           =   1020
            End
            Begin VB.TextBox AP9 
               Alignment       =   1  'Right Justify
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   285
               Left            =   7630
               TabIndex        =   186
               Top             =   2100
               Width           =   1020
            End
            Begin VB.TextBox AP10 
               Alignment       =   1  'Right Justify
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   285
               Left            =   7630
               TabIndex        =   185
               Top             =   2330
               Width           =   1020
            End
            Begin VB.TextBox H8 
               Alignment       =   1  'Right Justify
               Enabled         =   0   'False
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   285
               Left            =   8640
               TabIndex        =   181
               Top             =   1880
               Width           =   1020
            End
            Begin VB.TextBox H9 
               Alignment       =   1  'Right Justify
               Enabled         =   0   'False
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   285
               Left            =   8640
               TabIndex        =   180
               Top             =   2100
               Width           =   1020
            End
            Begin VB.TextBox H10 
               Alignment       =   1  'Right Justify
               Enabled         =   0   'False
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   285
               Left            =   8640
               TabIndex        =   179
               Top             =   2330
               Width           =   1020
            End
            Begin VB.TextBox M8 
               Alignment       =   1  'Right Justify
               BeginProperty DataFormat 
                  Type            =   1
                  Format          =   "#,##0.00"
                  HaveTrueFalseNull=   0
                  FirstDayOfWeek  =   0
                  FirstWeekOfYear =   0
                  LCID            =   1033
                  SubFormatType   =   1
               EndProperty
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   285
               Left            =   9640
               TabIndex        =   184
               Top             =   1880
               Width           =   1020
            End
            Begin VB.TextBox M9 
               Alignment       =   1  'Right Justify
               BeginProperty DataFormat 
                  Type            =   1
                  Format          =   "#,##0.00"
                  HaveTrueFalseNull=   0
                  FirstDayOfWeek  =   0
                  FirstWeekOfYear =   0
                  LCID            =   1033
                  SubFormatType   =   1
               EndProperty
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   285
               Left            =   9640
               TabIndex        =   183
               Top             =   2100
               Width           =   1020
            End
            Begin VB.TextBox M10 
               Alignment       =   1  'Right Justify
               BeginProperty DataFormat 
                  Type            =   1
                  Format          =   "#,##0.00"
                  HaveTrueFalseNull=   0
                  FirstDayOfWeek  =   0
                  FirstWeekOfYear =   0
                  LCID            =   1033
                  SubFormatType   =   1
               EndProperty
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   285
               Left            =   9640
               TabIndex        =   182
               Top             =   2330
               Width           =   1020
            End
            Begin VB.Frame Frame6 
               BackColor       =   &H80000004&
               Height          =   1095
               Left            =   0
               TabIndex        =   171
               Top             =   2530
               Width           =   10695
               Begin VB.TextBox txtKidneyRM 
                  Enabled         =   0   'False
                  Height          =   285
                  Left            =   4320
                  TabIndex        =   208
                  Top             =   765
                  Width           =   675
               End
               Begin VB.TextBox txtKidneyMth 
                  Height          =   285
                  Left            =   2760
                  TabIndex        =   207
                  Top             =   765
                  Width           =   555
               End
               Begin VB.TextBox P11 
                  Alignment       =   1  'Right Justify
                  BeginProperty Font 
                     Name            =   "Tahoma"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   285
                  Left            =   5620
                  TabIndex        =   235
                  Top             =   60
                  Width           =   1020
               End
               Begin VB.TextBox P12 
                  Alignment       =   1  'Right Justify
                  BeginProperty Font 
                     Name            =   "Tahoma"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   285
                  Left            =   5620
                  TabIndex        =   236
                  Top             =   320
                  Width           =   1020
               End
               Begin VB.TextBox P13 
                  Alignment       =   1  'Right Justify
                  BeginProperty Font 
                     Name            =   "Tahoma"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   285
                  Left            =   5620
                  TabIndex        =   237
                  Top             =   570
                  Width           =   1020
               End
               Begin VB.TextBox A11 
                  Alignment       =   1  'Right Justify
                  Enabled         =   0   'False
                  BeginProperty Font 
                     Name            =   "Tahoma"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   285
                  Left            =   6630
                  TabIndex        =   196
                  Top             =   60
                  Width           =   1020
               End
               Begin VB.TextBox A12 
                  Alignment       =   1  'Right Justify
                  Enabled         =   0   'False
                  BeginProperty Font 
                     Name            =   "Tahoma"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   285
                  Left            =   6630
                  TabIndex        =   195
                  Top             =   320
                  Width           =   1020
               End
               Begin VB.TextBox A13 
                  Alignment       =   1  'Right Justify
                  Enabled         =   0   'False
                  BeginProperty Font 
                     Name            =   "Tahoma"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   285
                  Left            =   6630
                  TabIndex        =   194
                  Top             =   570
                  Width           =   1020
               End
               Begin VB.TextBox AP11 
                  Alignment       =   1  'Right Justify
                  BeginProperty Font 
                     Name            =   "Tahoma"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   285
                  Left            =   7630
                  TabIndex        =   206
                  Top             =   60
                  Width           =   1020
               End
               Begin VB.TextBox AP12 
                  Alignment       =   1  'Right Justify
                  BeginProperty Font 
                     Name            =   "Tahoma"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   285
                  Left            =   7630
                  TabIndex        =   205
                  Top             =   320
                  Width           =   1020
               End
               Begin VB.TextBox AP13 
                  Alignment       =   1  'Right Justify
                  BeginProperty Font 
                     Name            =   "Tahoma"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   285
                  Left            =   7630
                  TabIndex        =   204
                  Top             =   570
                  Width           =   1020
               End
               Begin VB.TextBox H11 
                  Alignment       =   1  'Right Justify
                  Enabled         =   0   'False
                  BeginProperty Font 
                     Name            =   "Tahoma"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   285
                  Left            =   8640
                  TabIndex        =   199
                  Top             =   60
                  Width           =   1020
               End
               Begin VB.TextBox H12 
                  Alignment       =   1  'Right Justify
                  Enabled         =   0   'False
                  BeginProperty Font 
                     Name            =   "Tahoma"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   285
                  Left            =   8640
                  TabIndex        =   198
                  Top             =   320
                  Width           =   1020
               End
               Begin VB.TextBox H13 
                  Alignment       =   1  'Right Justify
                  Enabled         =   0   'False
                  BeginProperty Font 
                     Name            =   "Tahoma"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   285
                  Left            =   8640
                  TabIndex        =   197
                  Top             =   570
                  Width           =   1020
               End
               Begin VB.TextBox M11 
                  Alignment       =   1  'Right Justify
                  BeginProperty DataFormat 
                     Type            =   1
                     Format          =   "#,##0.00"
                     HaveTrueFalseNull=   0
                     FirstDayOfWeek  =   0
                     FirstWeekOfYear =   0
                     LCID            =   1033
                     SubFormatType   =   1
                  EndProperty
                  BeginProperty Font 
                     Name            =   "Tahoma"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   285
                  Left            =   9640
                  TabIndex        =   202
                  Top             =   60
                  Width           =   1020
               End
               Begin VB.TextBox M12 
                  Alignment       =   1  'Right Justify
                  BeginProperty DataFormat 
                     Type            =   1
                     Format          =   "#,##0.00"
                     HaveTrueFalseNull=   0
                     FirstDayOfWeek  =   0
                     FirstWeekOfYear =   0
                     LCID            =   1033
                     SubFormatType   =   1
                  EndProperty
                  BeginProperty Font 
                     Name            =   "Tahoma"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   285
                  Left            =   9640
                  TabIndex        =   201
                  Top             =   320
                  Width           =   1020
               End
               Begin VB.TextBox M13 
                  Alignment       =   1  'Right Justify
                  BeginProperty DataFormat 
                     Type            =   1
                     Format          =   "#,##0.00"
                     HaveTrueFalseNull=   0
                     FirstDayOfWeek  =   0
                     FirstWeekOfYear =   0
                     LCID            =   1033
                     SubFormatType   =   1
                  EndProperty
                  BeginProperty Font 
                     Name            =   "Tahoma"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   285
                  Left            =   9640
                  TabIndex        =   200
                  Top             =   570
                  Width           =   1020
               End
               Begin VB.TextBox P14 
                  Alignment       =   1  'Right Justify
                  BeginProperty Font 
                     Name            =   "Tahoma"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   285
                  Left            =   5620
                  TabIndex        =   238
                  Top             =   765
                  Width           =   1020
               End
               Begin VB.TextBox A14 
                  Alignment       =   1  'Right Justify
                  Enabled         =   0   'False
                  BeginProperty Font 
                     Name            =   "Tahoma"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   285
                  Left            =   6630
                  TabIndex        =   191
                  Top             =   765
                  Width           =   1020
               End
               Begin VB.TextBox AP14 
                  Alignment       =   1  'Right Justify
                  BeginProperty Font 
                     Name            =   "Tahoma"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   285
                  Left            =   7630
                  TabIndex        =   203
                  Top             =   765
                  Width           =   1020
               End
               Begin VB.TextBox H14 
                  Alignment       =   1  'Right Justify
                  Enabled         =   0   'False
                  BeginProperty Font 
                     Name            =   "Tahoma"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   285
                  Left            =   8640
                  TabIndex        =   192
                  Top             =   765
                  Width           =   1020
               End
               Begin VB.TextBox M14 
                  Alignment       =   1  'Right Justify
                  BeginProperty DataFormat 
                     Type            =   1
                     Format          =   "#,##0.00"
                     HaveTrueFalseNull=   0
                     FirstDayOfWeek  =   0
                     FirstWeekOfYear =   0
                     LCID            =   1033
                     SubFormatType   =   1
                  EndProperty
                  BeginProperty Font 
                     Name            =   "Tahoma"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   285
                  Left            =   9640
                  TabIndex        =   193
                  Top             =   765
                  Width           =   1020
               End
               Begin VB.Label Label36 
                  Caption         =   "Mths @ RM"
                  Height          =   255
                  Left            =   3360
                  TabIndex        =   214
                  Top             =   765
                  Width           =   1095
               End
               Begin VB.Label Label19 
                  Caption         =   "Emergency Out-Patient (Accident)"
                  BeginProperty Font 
                     Name            =   "Tahoma"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   255
                  Left            =   840
                  TabIndex        =   213
                  Top             =   120
                  Width           =   2670
               End
               Begin VB.Label Label20 
                  Caption         =   "Ambulance Fee"
                  BeginProperty Font 
                     Name            =   "Tahoma"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   255
                  Left            =   840
                  TabIndex        =   212
                  Top             =   345
                  Width           =   2085
               End
               Begin VB.Label Label21 
                  Caption         =   "Physiotherapy Treatment"
                  BeginProperty Font 
                     Name            =   "Tahoma"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   255
                  Left            =   840
                  TabIndex        =   211
                  Top             =   570
                  Width           =   2085
               End
               Begin VB.Label Label22 
                  Caption         =   "Kidney Dialysis && Cancer"
                  BeginProperty Font 
                     Name            =   "Tahoma"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   255
                  Left            =   840
                  TabIndex        =   210
                  Top             =   795
                  Width           =   1815
               End
               Begin VB.Label Label64 
                  Caption         =   "Out- Patient Benefits"
                  Height          =   735
                  Left            =   120
                  TabIndex        =   209
                  Top             =   165
                  Width           =   615
               End
            End
            Begin VB.TextBox P15 
               Alignment       =   1  'Right Justify
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   285
               Left            =   5620
               TabIndex        =   143
               Top             =   3580
               Width           =   1020
            End
            Begin VB.TextBox P16 
               Alignment       =   1  'Right Justify
               Enabled         =   0   'False
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   285
               Left            =   5620
               TabIndex        =   243
               Top             =   3840
               Width           =   1020
            End
            Begin VB.TextBox P17 
               Alignment       =   1  'Right Justify
               Enabled         =   0   'False
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   285
               Left            =   5620
               TabIndex        =   86
               Top             =   4080
               Width           =   1020
            End
            Begin VB.TextBox P18 
               Alignment       =   1  'Right Justify
               Enabled         =   0   'False
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   285
               Left            =   5620
               TabIndex        =   89
               Top             =   4305
               Width           =   1020
            End
            Begin VB.TextBox P19 
               Alignment       =   1  'Right Justify
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   285
               Left            =   5620
               TabIndex        =   92
               Top             =   4530
               Width           =   1020
            End
            Begin VB.TextBox P20 
               Alignment       =   1  'Right Justify
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   285
               Left            =   5620
               TabIndex        =   101
               Top             =   4785
               Width           =   1020
            End
            Begin VB.TextBox P21 
               Alignment       =   1  'Right Justify
               Enabled         =   0   'False
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   285
               Left            =   5620
               TabIndex        =   68
               Top             =   5040
               Width           =   1020
            End
            Begin VB.TextBox A15 
               Alignment       =   1  'Right Justify
               Enabled         =   0   'False
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   285
               Left            =   6630
               TabIndex        =   228
               Top             =   3580
               Width           =   1020
            End
            Begin VB.TextBox A16 
               Alignment       =   1  'Right Justify
               Enabled         =   0   'False
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   285
               Left            =   6630
               TabIndex        =   239
               Top             =   3840
               Width           =   1020
            End
            Begin VB.TextBox A17 
               Alignment       =   1  'Right Justify
               Enabled         =   0   'False
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   285
               Left            =   6630
               TabIndex        =   229
               Top             =   4080
               Width           =   1020
            End
            Begin VB.TextBox A18 
               Alignment       =   1  'Right Justify
               Enabled         =   0   'False
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   285
               Left            =   6630
               TabIndex        =   230
               Top             =   4320
               Width           =   1020
            End
            Begin VB.TextBox A19 
               Alignment       =   1  'Right Justify
               Enabled         =   0   'False
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   285
               Left            =   6630
               TabIndex        =   231
               Top             =   4530
               Width           =   1020
            End
            Begin VB.TextBox A20 
               Alignment       =   1  'Right Justify
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   285
               Left            =   6630
               TabIndex        =   233
               Top             =   4785
               Width           =   1020
            End
            Begin VB.TextBox A21 
               Alignment       =   1  'Right Justify
               Enabled         =   0   'False
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   285
               Left            =   6630
               TabIndex        =   234
               Top             =   5040
               Width           =   1020
            End
            Begin VB.TextBox AP15 
               Alignment       =   1  'Right Justify
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   285
               Left            =   7630
               TabIndex        =   140
               Top             =   3580
               Width           =   1020
            End
            Begin VB.TextBox AP16 
               Alignment       =   1  'Right Justify
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   285
               Left            =   7630
               TabIndex        =   242
               Top             =   3840
               Width           =   1020
            End
            Begin VB.TextBox AP17 
               Alignment       =   1  'Right Justify
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   285
               Left            =   7630
               TabIndex        =   71
               Top             =   4080
               Width           =   1020
            End
            Begin VB.TextBox AP18 
               Alignment       =   1  'Right Justify
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   285
               Left            =   7630
               TabIndex        =   72
               Top             =   4305
               Width           =   1020
            End
            Begin VB.TextBox AP19 
               Alignment       =   1  'Right Justify
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   285
               Left            =   7630
               TabIndex        =   73
               Top             =   4530
               Width           =   1020
            End
            Begin VB.TextBox AP20 
               Alignment       =   1  'Right Justify
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   285
               Left            =   7630
               TabIndex        =   75
               Top             =   4785
               Width           =   1020
            End
            Begin VB.TextBox AP21 
               Alignment       =   1  'Right Justify
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   285
               Left            =   7630
               TabIndex        =   69
               Top             =   5040
               Width           =   1020
            End
            Begin VB.TextBox H15 
               Alignment       =   1  'Right Justify
               Enabled         =   0   'False
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   285
               Left            =   8640
               TabIndex        =   142
               Top             =   3580
               Width           =   1020
            End
            Begin VB.TextBox H16 
               Alignment       =   1  'Right Justify
               Enabled         =   0   'False
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   285
               Left            =   8640
               TabIndex        =   241
               Top             =   3840
               Width           =   1020
            End
            Begin VB.TextBox H17 
               Alignment       =   1  'Right Justify
               Enabled         =   0   'False
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   285
               Left            =   8640
               TabIndex        =   85
               Top             =   4080
               Width           =   1020
            End
            Begin VB.TextBox H18 
               Alignment       =   1  'Right Justify
               Enabled         =   0   'False
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   285
               Left            =   8640
               TabIndex        =   88
               Top             =   4305
               Width           =   1020
            End
            Begin VB.TextBox H19 
               Alignment       =   1  'Right Justify
               Enabled         =   0   'False
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   285
               Left            =   8640
               TabIndex        =   91
               Top             =   4530
               Width           =   1020
            End
            Begin VB.TextBox H20 
               Alignment       =   1  'Right Justify
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   285
               Left            =   8640
               TabIndex        =   102
               Top             =   4785
               Width           =   1020
            End
            Begin VB.TextBox H21 
               Alignment       =   1  'Right Justify
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   285
               Left            =   8640
               TabIndex        =   67
               Top             =   5040
               Width           =   1020
            End
            Begin VB.TextBox M15 
               Alignment       =   1  'Right Justify
               BeginProperty DataFormat 
                  Type            =   1
                  Format          =   "#,##0.00"
                  HaveTrueFalseNull=   0
                  FirstDayOfWeek  =   0
                  FirstWeekOfYear =   0
                  LCID            =   1033
                  SubFormatType   =   1
               EndProperty
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   285
               Left            =   9640
               TabIndex        =   141
               Top             =   3580
               Width           =   1020
            End
            Begin VB.TextBox M16 
               Alignment       =   1  'Right Justify
               BeginProperty DataFormat 
                  Type            =   1
                  Format          =   "0.00"
                  HaveTrueFalseNull=   0
                  FirstDayOfWeek  =   0
                  FirstWeekOfYear =   0
                  LCID            =   1033
                  SubFormatType   =   1
               EndProperty
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   285
               Left            =   9640
               TabIndex        =   240
               Top             =   3840
               Width           =   1020
            End
            Begin VB.TextBox M17 
               Alignment       =   1  'Right Justify
               BeginProperty DataFormat 
                  Type            =   1
                  Format          =   "#,##0.00"
                  HaveTrueFalseNull=   0
                  FirstDayOfWeek  =   0
                  FirstWeekOfYear =   0
                  LCID            =   1033
                  SubFormatType   =   1
               EndProperty
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   285
               Left            =   9640
               TabIndex        =   84
               Top             =   4080
               Width           =   1020
            End
            Begin VB.TextBox M18 
               Alignment       =   1  'Right Justify
               BeginProperty DataFormat 
                  Type            =   1
                  Format          =   "#,##0.00"
                  HaveTrueFalseNull=   0
                  FirstDayOfWeek  =   0
                  FirstWeekOfYear =   0
                  LCID            =   1033
                  SubFormatType   =   1
               EndProperty
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   285
               Left            =   9640
               TabIndex        =   87
               Top             =   4305
               Width           =   1020
            End
            Begin VB.TextBox M19 
               Alignment       =   1  'Right Justify
               BeginProperty DataFormat 
                  Type            =   1
                  Format          =   "#,##0.00"
                  HaveTrueFalseNull=   0
                  FirstDayOfWeek  =   0
                  FirstWeekOfYear =   0
                  LCID            =   1033
                  SubFormatType   =   1
               EndProperty
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   285
               Left            =   9640
               TabIndex        =   90
               Top             =   4530
               Width           =   1020
            End
            Begin VB.TextBox M20 
               Alignment       =   1  'Right Justify
               BeginProperty DataFormat 
                  Type            =   1
                  Format          =   "#,##0.00"
                  HaveTrueFalseNull=   0
                  FirstDayOfWeek  =   0
                  FirstWeekOfYear =   0
                  LCID            =   1033
                  SubFormatType   =   1
               EndProperty
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   285
               Left            =   9640
               TabIndex        =   103
               Top             =   4785
               Width           =   1020
            End
            Begin VB.TextBox M21 
               Alignment       =   1  'Right Justify
               BeginProperty DataFormat 
                  Type            =   1
                  Format          =   "#,##0.00"
                  HaveTrueFalseNull=   0
                  FirstDayOfWeek  =   0
                  FirstWeekOfYear =   0
                  LCID            =   1033
                  SubFormatType   =   1
               EndProperty
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   285
               Left            =   9640
               TabIndex        =   66
               Top             =   5040
               Width           =   1020
            End
            Begin VB.TextBox A22 
               Alignment       =   1  'Right Justify
               Enabled         =   0   'False
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   285
               Left            =   6630
               TabIndex        =   232
               Top             =   5385
               Width           =   1020
            End
            Begin VB.TextBox AP22 
               Alignment       =   1  'Right Justify
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   285
               Left            =   7630
               TabIndex        =   147
               Top             =   5385
               Width           =   1020
            End
            Begin VB.TextBox H22 
               Alignment       =   1  'Right Justify
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   285
               Left            =   8640
               TabIndex        =   146
               Top             =   5385
               Width           =   1020
            End
            Begin VB.TextBox M22 
               Alignment       =   1  'Right Justify
               BeginProperty DataFormat 
                  Type            =   1
                  Format          =   "#,##0.00"
                  HaveTrueFalseNull=   0
                  FirstDayOfWeek  =   0
                  FirstWeekOfYear =   0
                  LCID            =   1033
                  SubFormatType   =   1
               EndProperty
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   285
               Left            =   9640
               TabIndex        =   145
               Top             =   5385
               Width           =   1020
            End
            Begin VB.TextBox P23 
               Alignment       =   1  'Right Justify
               Enabled         =   0   'False
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   285
               Left            =   5620
               TabIndex        =   245
               Top             =   5715
               Width           =   1020
            End
            Begin VB.TextBox A23 
               Alignment       =   1  'Right Justify
               Enabled         =   0   'False
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   285
               Left            =   6630
               TabIndex        =   248
               Top             =   5715
               Width           =   1020
            End
            Begin VB.TextBox AP23 
               Alignment       =   1  'Right Justify
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   285
               Left            =   7630
               TabIndex        =   150
               Top             =   5715
               Width           =   1020
            End
            Begin VB.TextBox H23 
               Alignment       =   1  'Right Justify
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   285
               Left            =   8640
               TabIndex        =   149
               Top             =   5715
               Width           =   1020
            End
            Begin VB.TextBox M23 
               Alignment       =   1  'Right Justify
               BeginProperty DataFormat 
                  Type            =   1
                  Format          =   "#,##0.00"
                  HaveTrueFalseNull=   0
                  FirstDayOfWeek  =   0
                  FirstWeekOfYear =   0
                  LCID            =   1033
                  SubFormatType   =   1
               EndProperty
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   285
               Left            =   9640
               TabIndex        =   148
               Top             =   5715
               Width           =   1020
            End
            Begin VB.TextBox TotalActual 
               Alignment       =   1  'Right Justify
               Enabled         =   0   'False
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   285
               Left            =   6630
               TabIndex        =   246
               Text            =   "0.00"
               Top             =   6120
               Width           =   1020
            End
            Begin VB.TextBox TotalApproved 
               Alignment       =   1  'Right Justify
               Enabled         =   0   'False
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   285
               Left            =   7630
               TabIndex        =   74
               Text            =   "0.00"
               Top             =   6120
               Width           =   1020
            End
            Begin VB.TextBox TotalPaidMedixHos 
               Alignment       =   1  'Right Justify
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   285
               Left            =   7635
               TabIndex        =   98
               Text            =   "0.00"
               Top             =   6630
               Width           =   1020
            End
            Begin VB.TextBox TotalGross 
               Alignment       =   1  'Right Justify
               Enabled         =   0   'False
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   285
               Left            =   6630
               TabIndex        =   247
               Text            =   "0.00"
               Top             =   6890
               Width           =   1020
            End
            Begin VB.TextBox TotalGrossApproved 
               Alignment       =   1  'Right Justify
               Enabled         =   0   'False
               Height          =   285
               Left            =   7630
               TabIndex        =   61
               Text            =   "0.00"
               Top             =   6890
               Width           =   1020
            End
            Begin VB.TextBox TotalGrossHos 
               Alignment       =   1  'Right Justify
               Enabled         =   0   'False
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   285
               Left            =   8640
               TabIndex        =   100
               Text            =   "0.00"
               Top             =   6890
               Width           =   1020
            End
            Begin VB.TextBox TotalHospital 
               Alignment       =   1  'Right Justify
               Enabled         =   0   'False
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   285
               Left            =   8640
               TabIndex        =   94
               Text            =   "0.00"
               Top             =   6120
               Width           =   1020
            End
            Begin VB.TextBox TotalPaidByMemberHos 
               Alignment       =   1  'Right Justify
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   285
               Left            =   8640
               TabIndex        =   96
               Text            =   "0.00"
               Top             =   6380
               Width           =   1020
            End
            Begin VB.TextBox TotalMember 
               Alignment       =   1  'Right Justify
               BeginProperty DataFormat 
                  Type            =   1
                  Format          =   "#,##0.00;(#,##0.00)"
                  HaveTrueFalseNull=   0
                  FirstDayOfWeek  =   0
                  FirstWeekOfYear =   0
                  LCID            =   1033
                  SubFormatType   =   1
               EndProperty
               Enabled         =   0   'False
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   285
               Left            =   9640
               TabIndex        =   93
               Text            =   "0.00"
               Top             =   6120
               Width           =   1020
            End
            Begin VB.TextBox TotalPaidByMemberMember 
               Alignment       =   1  'Right Justify
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   285
               Left            =   9640
               TabIndex        =   95
               Text            =   "0.00"
               Top             =   6380
               Width           =   1020
            End
            Begin VB.TextBox TotalPaidMedixMember 
               Alignment       =   1  'Right Justify
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   285
               Left            =   9645
               TabIndex        =   97
               Text            =   "0.00"
               Top             =   6630
               Visible         =   0   'False
               Width           =   1020
            End
            Begin VB.TextBox TotalGrossMember 
               Alignment       =   1  'Right Justify
               Enabled         =   0   'False
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   285
               Left            =   9640
               TabIndex        =   99
               Text            =   "0.00"
               Top             =   6890
               Width           =   1020
            End
            Begin VB.Label Label66 
               Caption         =   "Organ Transplant"
               Height          =   255
               Left            =   840
               TabIndex        =   111
               Top             =   3595
               Width           =   1935
            End
            Begin VB.Label Label13 
               Caption         =   "Pre-Hospital Specialist Consultation"
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   255
               Left            =   900
               TabIndex        =   190
               Top             =   1935
               Width           =   2580
            End
            Begin VB.Label Label17 
               Caption         =   "In-Hospital Physician Visit"
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   255
               Left            =   900
               TabIndex        =   189
               Top             =   2175
               Width           =   2085
            End
            Begin VB.Label Label63 
               Caption         =   "Non- Surgical"
               Height          =   495
               Left            =   120
               TabIndex        =   188
               Top             =   1920
               Width           =   615
            End
            Begin VB.Label Label61 
               Caption         =   "Hospital"
               Height          =   375
               Left            =   120
               TabIndex        =   130
               Top             =   240
               Width           =   615
            End
            Begin VB.Label Label65 
               Caption         =   "Others"
               Height          =   375
               Left            =   120
               TabIndex        =   129
               Top             =   3720
               Width           =   495
            End
            Begin VB.Label Label45 
               Caption         =   "Days @ RM"
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   255
               Left            =   3360
               TabIndex        =   128
               Top             =   4320
               Width           =   825
            End
            Begin VB.Label Label33 
               Caption         =   "Days @ RM"
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   255
               Left            =   3360
               TabIndex        =   127
               Top             =   405
               Width           =   870
            End
            Begin VB.Label Label28 
               Caption         =   "Days @ RM"
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   255
               Left            =   3360
               TabIndex        =   126
               Top             =   180
               Width           =   825
            End
            Begin VB.Label Label32 
               Caption         =   "Total"
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   255
               Left            =   3240
               TabIndex        =   125
               Top             =   6120
               Width           =   675
            End
            Begin VB.Line Line5 
               BorderWidth     =   2
               X1              =   -720
               X2              =   11520
               Y1              =   6050
               Y2              =   6050
            End
            Begin VB.Label Label30 
               Caption         =   "Registration Fee"
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   255
               Left            =   840
               TabIndex        =   124
               Top             =   5055
               Width           =   1635
            End
            Begin VB.Label Label29 
               Caption         =   "Guardian Allowance"
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   255
               Left            =   840
               TabIndex        =   123
               Top             =   4320
               Width           =   1635
            End
            Begin VB.Label Label24 
               Caption         =   "Telephone Charges"
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   255
               Left            =   840
               TabIndex        =   122
               Top             =   4800
               Width           =   2085
            End
            Begin VB.Label Label23 
               Caption         =   "Medical Report"
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   255
               Left            =   840
               TabIndex        =   121
               Top             =   4545
               Width           =   1770
            End
            Begin VB.Label Label11 
               Caption         =   "Supplies && Services"
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   255
               Left            =   900
               TabIndex        =   120
               Top             =   615
               Width           =   2445
            End
            Begin VB.Label Label10 
               Caption         =   "ICU"
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   255
               Left            =   900
               TabIndex        =   119
               Top             =   405
               Width           =   1485
            End
            Begin VB.Label Label1 
               Caption         =   "Room && Board"
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   255
               Left            =   900
               TabIndex        =   118
               Top             =   173
               Width           =   1515
            End
            Begin VB.Label Label34 
               Caption         =   "Paid By Mem"
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   255
               Left            =   3240
               TabIndex        =   117
               Top             =   6400
               Width           =   1230
            End
            Begin VB.Label Label35 
               Caption         =   "Payable by MEDIX"
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   255
               Left            =   3240
               TabIndex        =   116
               Top             =   6630
               Width           =   1380
            End
            Begin VB.Label Label37 
               Caption         =   "Final Total"
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   255
               Left            =   3240
               TabIndex        =   115
               Top             =   6890
               Width           =   1095
            End
            Begin VB.Label lblMaxRB 
               BackColor       =   &H00C0FFC0&
               Height          =   255
               Left            =   5040
               TabIndex        =   114
               Top             =   180
               Width           =   495
            End
            Begin VB.Label lblMaxICU 
               BackColor       =   &H00C0FFC0&
               Height          =   255
               Left            =   5040
               TabIndex        =   113
               Top             =   420
               Width           =   495
            End
            Begin VB.Label lblMaxLodger 
               BackColor       =   &H00C0FFC0&
               Height          =   255
               Left            =   5040
               TabIndex        =   112
               Top             =   4320
               Width           =   495
            End
            Begin VB.Label lblMaxCash 
               BackColor       =   &H00C0FFC0&
               Height          =   255
               Left            =   5040
               TabIndex        =   110
               Top             =   3840
               Width           =   495
            End
            Begin VB.Label Label57 
               Caption         =   "Days @RM"
               Height          =   255
               Left            =   3360
               TabIndex        =   109
               Top             =   3840
               Width           =   975
            End
            Begin VB.Label Label56 
               Caption         =   "Govt Hospital Benefits"
               Height          =   255
               Left            =   840
               TabIndex        =   108
               Top             =   3855
               Width           =   1695
            End
            Begin VB.Label Label31 
               Caption         =   "Service Tax"
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   255
               Left            =   840
               TabIndex        =   107
               Top             =   4095
               Width           =   1635
            End
         End
      End
      Begin VB.CommandButton CmdInvoice 
         Caption         =   "&Invoices"
         Height          =   375
         Left            =   -74760
         TabIndex        =   55
         Top             =   6660
         Width           =   975
      End
      Begin VB.Frame Frame16 
         Caption         =   " Diagnosis"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1680
         Left            =   -74760
         TabIndex        =   53
         Top             =   2760
         Width           =   11115
         Begin MSComctlLib.ListView lstFinalDiag 
            Height          =   1335
            Left            =   120
            TabIndex        =   54
            Top             =   240
            Width           =   10905
            _ExtentX        =   19235
            _ExtentY        =   2355
            View            =   3
            LabelEdit       =   1
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
            NumItems        =   4
            BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Text            =   "Index"
               Object.Width           =   0
            EndProperty
            BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   1
               Text            =   "ICD Code"
               Object.Width           =   2540
            EndProperty
            BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   2
               Text            =   "Diag Desc"
               Object.Width           =   2540
            EndProperty
            BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   3
               Text            =   "Date On-Set"
               Object.Width           =   2540
            EndProperty
         End
      End
      Begin VB.CommandButton cmdExit2 
         Caption         =   "E&xit"
         Height          =   375
         Left            =   -73560
         TabIndex        =   40
         Top             =   6660
         Width           =   855
      End
      Begin VB.Frame GeneralFrame 
         Height          =   6195
         Left            =   -74880
         TabIndex        =   12
         Top             =   360
         Width           =   11415
         Begin VB.Frame Frame1 
            Caption         =   "Invoices"
            Height          =   2055
            Left            =   120
            TabIndex        =   51
            Top             =   4080
            Width           =   11175
            Begin MSComctlLib.ListView lstInvoice 
               Height          =   1695
               Left            =   120
               TabIndex        =   52
               Top             =   240
               Width           =   10935
               _ExtentX        =   19288
               _ExtentY        =   2990
               View            =   3
               LabelEdit       =   1
               Sorted          =   -1  'True
               LabelWrap       =   -1  'True
               HideSelection   =   -1  'True
               AllowReorder    =   -1  'True
               FullRowSelect   =   -1  'True
               _Version        =   393217
               ForeColor       =   -2147483640
               BackColor       =   -2147483643
               BorderStyle     =   1
               Appearance      =   1
               NumItems        =   47
               BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
                  Text            =   "Inv No"
                  Object.Width           =   2540
               EndProperty
               BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
                  SubItemIndex    =   1
                  Text            =   "Inv Date"
                  Object.Width           =   2540
               EndProperty
               BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
                  SubItemIndex    =   2
                  Text            =   "INV Amount"
                  Object.Width           =   2540
               EndProperty
               BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
                  SubItemIndex    =   3
                  Text            =   "Pay To"
                  Object.Width           =   2540
               EndProperty
               BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
                  SubItemIndex    =   4
                  Text            =   "Hosp Type"
                  Object.Width           =   2540
               EndProperty
               BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
                  SubItemIndex    =   5
                  Text            =   "Hosp Type"
                  Object.Width           =   0
               EndProperty
               BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
                  SubItemIndex    =   6
                  Text            =   "Excess Paid By Member"
                  Object.Width           =   2540
               EndProperty
               BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
                  SubItemIndex    =   7
                  Text            =   "DailyRB"
                  Object.Width           =   2540
               EndProperty
               BeginProperty ColumnHeader(9) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
                  SubItemIndex    =   8
                  Text            =   "RBAmt"
                  Object.Width           =   2540
               EndProperty
               BeginProperty ColumnHeader(10) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
                  SubItemIndex    =   9
                  Text            =   "ICUDay"
                  Object.Width           =   2540
               EndProperty
               BeginProperty ColumnHeader(11) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
                  SubItemIndex    =   10
                  Text            =   "ICUAmt"
                  Object.Width           =   2540
               EndProperty
               BeginProperty ColumnHeader(12) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
                  SubItemIndex    =   11
                  Text            =   "HSS"
                  Object.Width           =   2540
               EndProperty
               BeginProperty ColumnHeader(13) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
                  SubItemIndex    =   12
                  Text            =   "OT"
                  Object.Width           =   2540
               EndProperty
               BeginProperty ColumnHeader(14) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
                  SubItemIndex    =   13
                  Text            =   "PHService"
                  Object.Width           =   2540
               EndProperty
               BeginProperty ColumnHeader(15) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
                  SubItemIndex    =   14
                  Text            =   "IHPysVisit"
                  Object.Width           =   2540
               EndProperty
               BeginProperty ColumnHeader(16) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
                  SubItemIndex    =   15
                  Text            =   "PHTreatment"
                  Object.Width           =   2540
               EndProperty
               BeginProperty ColumnHeader(17) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
                  SubItemIndex    =   16
                  Text            =   "MR"
                  Object.Width           =   2540
               EndProperty
               BeginProperty ColumnHeader(18) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
                  SubItemIndex    =   17
                  Text            =   "Phone Charges"
                  Object.Width           =   2540
               EndProperty
               BeginProperty ColumnHeader(19) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
                  SubItemIndex    =   18
                  Text            =   "LodgerDay"
                  Object.Width           =   2540
               EndProperty
               BeginProperty ColumnHeader(20) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
                  SubItemIndex    =   19
                  Text            =   "LodgerAmt"
                  Object.Width           =   2540
               EndProperty
               BeginProperty ColumnHeader(21) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
                  SubItemIndex    =   20
                  Text            =   "Admin"
                  Object.Width           =   2540
               EndProperty
               BeginProperty ColumnHeader(22) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
                  SubItemIndex    =   21
                  Text            =   "Tax"
                  Object.Width           =   2540
               EndProperty
               BeginProperty ColumnHeader(23) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
                  SubItemIndex    =   22
                  Text            =   "Others"
                  Object.Width           =   2540
               EndProperty
               BeginProperty ColumnHeader(24) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
                  SubItemIndex    =   23
                  Text            =   "CashDay"
                  Object.Width           =   2540
               EndProperty
               BeginProperty ColumnHeader(25) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
                  SubItemIndex    =   24
                  Text            =   "CashAmt"
                  Object.Width           =   2540
               EndProperty
               BeginProperty ColumnHeader(26) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
                  SubItemIndex    =   25
                  Text            =   "XRay"
                  Object.Width           =   2540
               EndProperty
               BeginProperty ColumnHeader(27) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
                  SubItemIndex    =   26
                  Text            =   "Lab"
                  Object.Width           =   2540
               EndProperty
               BeginProperty ColumnHeader(28) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
                  SubItemIndex    =   27
                  Text            =   "Spec"
                  Object.Width           =   2540
               EndProperty
               BeginProperty ColumnHeader(29) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
                  SubItemIndex    =   28
                  Text            =   "SFee"
                  Object.Width           =   2540
               EndProperty
               BeginProperty ColumnHeader(30) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
                  SubItemIndex    =   29
                  Text            =   "AFee"
                  Object.Width           =   2540
               EndProperty
               BeginProperty ColumnHeader(31) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
                  SubItemIndex    =   30
                  Text            =   "EOut-Patient"
                  Object.Width           =   2540
               EndProperty
               BeginProperty ColumnHeader(32) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
                  SubItemIndex    =   31
                  Text            =   "Ambfee"
                  Object.Width           =   2540
               EndProperty
               BeginProperty ColumnHeader(33) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
                  SubItemIndex    =   32
                  Text            =   "OPT"
                  Object.Width           =   2540
               EndProperty
               BeginProperty ColumnHeader(34) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
                  SubItemIndex    =   33
                  Text            =   "MonthlyKidney"
                  Object.Width           =   2540
               EndProperty
               BeginProperty ColumnHeader(35) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
                  SubItemIndex    =   34
                  Text            =   "Pre60"
                  Object.Width           =   2540
               EndProperty
               BeginProperty ColumnHeader(36) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
                  SubItemIndex    =   35
                  Text            =   "PreGP"
                  Object.Width           =   2540
               EndProperty
               BeginProperty ColumnHeader(37) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
                  SubItemIndex    =   36
                  Text            =   "PreMedic"
                  Object.Width           =   2540
               EndProperty
               BeginProperty ColumnHeader(38) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
                  SubItemIndex    =   37
                  Text            =   "PreMR "
                  Object.Width           =   2540
               EndProperty
               BeginProperty ColumnHeader(39) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
                  SubItemIndex    =   38
                  Text            =   "PreSpec"
                  Object.Width           =   2540
               EndProperty
               BeginProperty ColumnHeader(40) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
                  SubItemIndex    =   39
                  Text            =   "Post_Host"
                  Object.Width           =   2540
               EndProperty
               BeginProperty ColumnHeader(41) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
                  SubItemIndex    =   40
                  Text            =   "NotSame"
                  Object.Width           =   2540
               EndProperty
               BeginProperty ColumnHeader(42) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
                  SubItemIndex    =   41
                  Text            =   "Others"
                  Object.Width           =   2540
               EndProperty
               BeginProperty ColumnHeader(43) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
                  SubItemIndex    =   42
                  Text            =   "Status"
                  Object.Width           =   0
               EndProperty
               BeginProperty ColumnHeader(44) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
                  SubItemIndex    =   43
                  Text            =   "Index"
                  Object.Width           =   0
               EndProperty
               BeginProperty ColumnHeader(45) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
                  SubItemIndex    =   44
                  Text            =   "INVPayType"
                  Object.Width           =   2540
               EndProperty
               BeginProperty ColumnHeader(46) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
                  SubItemIndex    =   45
                  Text            =   "ClaimVerion"
                  Object.Width           =   176
               EndProperty
               BeginProperty ColumnHeader(47) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
                  SubItemIndex    =   46
                  Text            =   "Bordxed?"
                  Object.Width           =   176
               EndProperty
            End
         End
         Begin VB.TextBox txtCASRemark 
            Height          =   285
            Left            =   1200
            TabIndex        =   50
            Top             =   2040
            Width           =   4095
         End
         Begin VB.TextBox txtAdmissionDate 
            Enabled         =   0   'False
            Height          =   285
            Left            =   6600
            TabIndex        =   46
            Top             =   1680
            Width           =   1425
         End
         Begin VB.TextBox txtDisChargeDate 
            Enabled         =   0   'False
            Height          =   285
            Left            =   6600
            TabIndex        =   45
            Top             =   2040
            Width           =   1425
         End
         Begin VB.TextBox txtHLTCode 
            Height          =   285
            Left            =   480
            TabIndex        =   44
            Top             =   2640
            Visible         =   0   'False
            Width           =   495
         End
         Begin VB.TextBox txtINSCode 
            Height          =   285
            Left            =   480
            TabIndex        =   43
            Top             =   3120
            Visible         =   0   'False
            Width           =   615
         End
         Begin VB.TextBox txtHOS 
            Enabled         =   0   'False
            Height          =   285
            Left            =   1200
            TabIndex        =   41
            Top             =   960
            Width           =   4095
         End
         Begin VB.TextBox Text10 
            Height          =   645
            Left            =   9240
            ScrollBars      =   3  'Both
            TabIndex        =   39
            Top             =   1680
            Width           =   2055
         End
         Begin VB.TextBox Text9 
            Height          =   645
            Left            =   1200
            ScrollBars      =   3  'Both
            TabIndex        =   37
            Top             =   1320
            Width           =   4095
         End
         Begin VB.TextBox Text8 
            Height          =   285
            Left            =   6600
            TabIndex        =   35
            Top             =   1320
            Width           =   1425
         End
         Begin VB.TextBox txtMBMNumber 
            Height          =   285
            Left            =   9240
            TabIndex        =   26
            Top             =   240
            Width           =   2055
         End
         Begin VB.TextBox txtMBMPolNo 
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   9240
            TabIndex        =   24
            Top             =   600
            Width           =   2055
         End
         Begin VB.TextBox txtType 
            Enabled         =   0   'False
            Height          =   285
            Left            =   9240
            TabIndex        =   22
            Top             =   1320
            Width           =   2055
         End
         Begin VB.TextBox TxtPayor 
            Enabled         =   0   'False
            Height          =   285
            Left            =   6600
            TabIndex        =   20
            Top             =   600
            Width           =   1425
         End
         Begin VB.TextBox txtPolEffDate 
            Enabled         =   0   'False
            Height          =   285
            Left            =   6600
            TabIndex        =   19
            Top             =   960
            Width           =   1425
         End
         Begin VB.TextBox txtExpDate 
            Enabled         =   0   'False
            Height          =   285
            Left            =   9240
            TabIndex        =   17
            Top             =   960
            Width           =   2055
         End
         Begin VB.TextBox txtPLNCode 
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   6600
            TabIndex        =   5
            Top             =   240
            Width           =   1425
         End
         Begin VB.TextBox txtPatientName 
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   1200
            TabIndex        =   6
            Top             =   600
            Width           =   4095
         End
         Begin VB.TextBox txtMBMName 
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   1200
            TabIndex        =   4
            Top             =   240
            Width           =   4095
         End
         Begin VB.Label Label60 
            Caption         =   "Remark"
            Height          =   375
            Left            =   120
            TabIndex        =   49
            Top             =   2040
            Width           =   855
         End
         Begin VB.Label Label49 
            Caption         =   "DOA"
            Height          =   255
            Left            =   5640
            TabIndex        =   48
            Top             =   1680
            Width           =   1335
         End
         Begin VB.Label Label50 
            Caption         =   "DOD"
            Height          =   255
            Left            =   5640
            TabIndex        =   47
            Top             =   2040
            Width           =   375
         End
         Begin VB.Label Label51 
            Caption         =   "Hospital"
            Height          =   255
            Left            =   120
            TabIndex        =   42
            Top             =   960
            Width           =   855
         End
         Begin VB.Label Label7 
            Caption         =   "Fol-up Dates"
            Height          =   255
            Left            =   8280
            TabIndex        =   38
            Top             =   1680
            Width           =   1335
         End
         Begin VB.Label Label6 
            Caption         =   "Disability"
            Height          =   255
            Left            =   120
            TabIndex        =   36
            Top             =   1320
            Width           =   615
         End
         Begin VB.Label Label4 
            Caption         =   "ICD Code"
            Height          =   255
            Left            =   5640
            TabIndex        =   34
            Top             =   1320
            Width           =   735
         End
         Begin VB.Label Label26 
            Caption         =   "Mem No"
            Height          =   255
            Left            =   8280
            TabIndex        =   27
            Top             =   240
            Width           =   1215
         End
         Begin VB.Label Label27 
            Caption         =   "Pol No"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   8280
            TabIndex        =   25
            Top             =   600
            Width           =   795
         End
         Begin VB.Label Label44 
            Caption         =   "Claim Types"
            Height          =   255
            Left            =   8280
            TabIndex        =   23
            Top             =   1320
            Width           =   1095
         End
         Begin VB.Label Label55 
            Caption         =   "Payor"
            Height          =   255
            Left            =   5640
            TabIndex        =   21
            Top             =   600
            Width           =   615
         End
         Begin VB.Label Label54 
            Caption         =   "Exp Date"
            Height          =   375
            Left            =   8280
            TabIndex        =   18
            Top             =   960
            Width           =   1335
         End
         Begin VB.Label Label47 
            Caption         =   "Eff Date"
            Height          =   375
            Left            =   5640
            TabIndex        =   16
            Top             =   960
            Width           =   1335
         End
         Begin VB.Label Label5 
            Caption         =   "Plan Code"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   5640
            TabIndex        =   15
            Top             =   240
            Width           =   795
         End
         Begin VB.Label Label8 
            Caption         =   "Patient Name"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   120
            TabIndex        =   14
            Top             =   600
            Width           =   1275
         End
         Begin VB.Label Label2 
            Caption         =   "Mem Name"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   120
            TabIndex        =   13
            Top             =   240
            Width           =   1275
         End
      End
   End
End
Attribute VB_Name = "frmClaimAdmission"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim intLastPoz As Integer
Dim intLastPoz2 As Integer
Dim VersionNo As Integer

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    '--- Enable PrintScreen feature
    '--- IMPORTANT - Don't forget to set .KeyPreview = True for this form at design time!!
    If Shift = 0 Then
        Select Case KeyCode
            Case vbKeyF6
                KeyCode = 0
                PrintScreenSnapShot
        End Select
    End If
End Sub

Private Sub Form_Load()
    Dim intFullHeight As Integer
    Dim intDisplayHeight As Integer
    
    VersionNo = 1
    
    intFullHeight = 4000
    intDisplayHeight = 2535
    intLastPoz = 0
    intLastPoz2 = 0
    
    Me.Height = intDisplayHeight
    cboVersion.AddItem "1"
    cboVersion.ListIndex = 0
    With VScroll2
        '.Height = Me.ScaleHeight
        .Min = 0
        .Max = intFullHeight - intDisplayHeight
        .SmallChange = Screen.TwipsPerPixelX * 5
        .LargeChange = .SmallChange
    End With
End Sub

Private Sub cmdExit_Click()
    Dim RecordExit
    RecordExit = MsgBox("Are you sure?", vbYesNo + vbExclamation)
    If RecordExit = vbYes Then
        Frame11.Top = 120
        Unload Me
    End If
End Sub


Private Sub cmdExit2_Click()
 Dim RecordExit
    RecordExit = MsgBox("Are you sure?", vbYesNo + vbExclamation)
    If RecordExit = vbYes Then
        Frame11.Top = 120
        Unload Me
    End If
End Sub



' **********************88
Public Sub cmdDisplay_Click()
    Dim CAS As New ADODB.Recordset
    Dim strsql, sqlPolStatus   As String
    Dim PolStatus As New ADODB.Recordset
    
'    cboHOSCode.Clear
 '   lstInvoice.ListItems.Clear
    strsql = "SELECT CASNumber, MBMStatus,  MBMNumber, PAycode, " & _
        " MBMPatientCovID, MBMPolicyNo, MBMName, " & _
        " PLNCode, CASStatus, MBMPayorExpDate , " & _
        " CASRemark,  MBMPayorEffDate , HLTCode ," & _
        " CASHOSCode,    CASDateAdmitted, CASDischargeDate , MBMAvailableLimit,  " & _
        "  PayTypeCode,  INSCOde  "
    strsql = strsql & " from VIEW_Cas_MBM_Claims where 0=0"
    
    If Trim(cboVersion) = "" Then
        MsgBox "Please select a version nombor!", vbCritical
        cboVersion.SetFocus
    ElseIf Trim(txtCASNumber) = "" Then
        MsgBox "Please select a Case Number ! ", vbCritical
        txtCASNumber.SetFocus
    Else
        strsql = strsql & " And CASNumber = '" & Trim(txtCASNumber) & "'"
        CAS.Open strsql, medixmasterconn, adOpenKeyset, adLockOptimistic
            If CAS.recordCount Then
                Call Display_Diagnosis("F", txtCASNumber)
'                Call Display_Hospital(txtCASNumber)
                sqlPolStatus = " select pstatuscode ,pstatusDesc from PolicyStatus where PStatusCode ='" & Trim(CAS!MBMStatus) & "'"
                PolStatus.Open sqlPolStatus, medixmasterconn, adOpenKeyset, adLockOptimistic
                If PolStatus.recordCount Then
                    txtPolStatus = PolStatus!PStatusDesc
                End If
                
                PolStatus.Close
                Set PolStatus = Nothing

                If ChkStr(CAS!CASStatus) = "O" Then
                    lblCASStatus = "Open"
                ElseIf ChkStr(CAS!CASStatus) = "C" Then
                    lblCASStatus = "Close"
                ElseIf ChkStr(CAS!CASStatus) = "R" Then
                    lblCASStatus = "Repudiated"
                End If
                
                txtCASNumber = CAS!CASNumber
                txtMBMNumber = CAS!MBMNumber
                txtMBMName = CAS!MBMName
                txtMBMPolNo = CAS!MBMPolicyNo
                txtPLNCode = CAS!PLNCode
                txtINSCode = CAS!InsCode
                txtPolEffDate = CAS!MBMPayorEffDate
                txtExpDate = CAS!MBMPayorEXPDate
                TxtPayor = CAS!PAYCode
                
                If MBMPatientCovID <> 0 Then
                    txtPatientName = CAS!MBMCName
                    
                Else
                    txtPatientName = txtMBMName
                    
                End If

                If Trim(CAS!PayTypeCode) <> "" Then
                    txtType = CAS!PayTypeCode
                End If
                If Len(Trim(CAS!CASRemark)) Then
                    txtCASRemark = CAS!CASRemark
                End If
                txtHLTCode = CAS!HLTCode
                txtAdmissionDate = CAS!CASDateAdmitted
                If Len(Trim(CAS!CASDischargeDate)) <> 0 Then
                    txtDisChargeDate = CAS!CASDischargeDate
                End If
                If Len(Trim(CAS!CASHOSCode)) <> 0 Then
                    txtHOS = CAS!CASHOSCode
                End If
                txtAnnualLimit = CAS!MBMAvailableLimit
                PAnnualLimit = CAS!MBMAvailableLimit
               ' txtLGAmount = CAS!CASReserveAmount
                Call display_benefit
                Call getVersion
                Call InputFromInvoice
                Call getDatatoWS
                Call CalculateICULimit
                Call CalculateLodgerLimit
                Call CalculateRoomLimit
                Call CalculateTax
                ' No of days only entered after get date from Invoice
                 'RB
                Call A1_LostFocus
                'ICU
                Call A2_LostFocus
                
                ' Lodger
                Call A17_LostFocus
                'Tax
                Call A19_LostFocus
                'CAsh
                Call A21_LostFocus
                
                Call calcTotal
                Call txtRMCash_LostFocus
            Else
                MsgBox "Record not found! ", vbCritical
            End If
    End If
End Sub


Private Sub getVersion()
    Dim Version As New ADODB.Recordset
    Dim strsql As String
    
    strsql = "select casNumber, claimversion from claim where casNUmber ='" & Trim(txtCASNumber) & "'"
    Version.Open strsql, medixmasterconnWS, adOpenKeyset, adLockOptimistic
    If Not Version.EOF Then
        cboVersion.Clear
        Do While Not Version.EOF
            cboVersion.AddItem Version!claimversion
            Version.MoveNext
        Loop
        
    End If
    Version.Close
    Set Version = Nothing
End Sub

' ****************************************** Get Data from Invoice

Private Sub InputFromInvoice()
    Dim Invoice As New ADODB.Recordset
    Dim INVlist As ListItem
    Dim Version As Integer
    'cboVersion.ListIndex = 1
    If (cboVersion.Text) = "" Then
        Version = 1
    Else
        Version = cboVersion
    End If
    Dim strsql As String
        strsql = "SELECt INVIndex,  " & _
             " INVAmount, INVWSversion , " & _
            "INVPayee , ClaimType," & _
            " INVPayTo,  INVPaidBYMem, " & _
            " CLMBoardDay,     CLMRB,    CLMICUDay, " & _
            " CLMICU ,    CLMServices,    CLMTheatre, " & _
            " CLMPreXray,    CLMPreLab,    CLMPreOthers, " & _
            " CLMSurgicalAmt, CLMAna,    CLMPreHSAmt ," & _
            " CLMPhysician,    CLMPost,    CLMEmergency," & _
            " CLMAmbulance,    CLMOutpatient,    CLMMonthly," & _
            " CLMMedical,    CLMPhone,    CLMLodgerDay," & _
            " CLMLodger, CLMAdmin,     CLMTax, CLMOthers1," & _
            " CLMCashDay,    CLMCash, PreH60,    preMedic, PreGP," & _
            " PreSpec,    PostHosp,    PreMR," & _
            " PostNotSame,    Others  FROM View_CAS_Invoice " & _
            " where CASNUmber ='" & Trim(txtCASNumber) & "'" & _
            " and INVWSVersion= " & Trim(Version)
        Invoice.Open strsql, medixmasterconnWS, adOpenKeyset, adLockOptimistic
        
        If Invoice.recordCount Then
            lstInvoice.listitems.Clear
            Do Until Invoice.EOF
                Set INVlist = lstInvoice.listitems.Add(, , Invoice!InvINdex)
                With INVlist
                    If Trim(Invoice!INVAmount) <> "" Then
                        .SubItems(2) = Invoice!INVAmount
                    End If
                    If Trim(Invoice!INVPayee) <> "" Then
                        .SubItems(3) = Invoice!INVPayee
                    End If
                    If Trim(Invoice!ClaimType) <> "" Then
                        .SubItems(4) = Invoice!ClaimType
                    End If
                   
                    If Trim(Invoice!INVPaidBYMem) <> "" Then
                        .SubItems(6) = Invoice!INVPaidBYMem
                    End If
                    
                    If Trim(Invoice!CLMBoardDay) <> "" Then
                        .SubItems(7) = Invoice!CLMBoardDay
                    End If
                    If Trim(Invoice!CLMRB) <> "" Then
                        .SubItems(8) = Invoice!CLMRB
                    End If
                    If Trim(Invoice!CLMICUDay) <> "" Then
                        .SubItems(9) = Invoice!CLMICUDay
                    End If
                    If Trim(Invoice!CLMICU) <> "" Then
                        .SubItems(10) = Invoice!CLMICU
                    End If
                    If Trim(Invoice!CLMServices) <> "" Then
                        .SubItems(11) = Invoice!CLMServices
                    End If
                    If Trim(Invoice!CLMTheatre) <> "" Then
                        .SubItems(12) = Invoice!CLMTheatre
                    End If
                    If Trim(Invoice!CLMPreHSAmt) <> "" Then
                        .SubItems(13) = Invoice!CLMPreHSAmt
                    End If
                    If Trim(Invoice!CLMPhysician) <> "" Then
                        .SubItems(14) = Invoice!CLMPhysician
                    End If
                    If Trim(Invoice!CLMPost) <> "" Then
                        .SubItems(15) = Invoice!CLMPost
                    End If
                    If Trim(Invoice!CLMMedical) <> "" Then
                        .SubItems(16) = Invoice!CLMMedical
                    End If
                    If Trim(Invoice!CLMPhone) <> "" Then
                        .SubItems(17) = Invoice!CLMPhone
                    End If
                    If Trim(Invoice!CLMLodgerday) <> "" Then
                        .SubItems(18) = Invoice!CLMLodgerday
                    End If
                    If Trim(Invoice!CLMLodger) <> "" Then
                        .SubItems(19) = Invoice!CLMLodger
                    End If
                    If Trim(Invoice!CLMAdmin) <> "" Then
                        .SubItems(20) = Invoice!CLMAdmin
                    End If
                    If Trim(Invoice!CLMTax) <> "" Then
                        .SubItems(21) = Invoice!CLMTax
                    End If
                    If Trim(Invoice!CLMOthers1) <> "" Then
                        .SubItems(22) = Invoice!CLMOthers1
                    End If
                    If Trim(Invoice!CLMCashDay) <> "" Then
                    .SubItems(23) = Invoice!CLMCashDay
                    End If
                    If Trim(Invoice!CLMCash) <> "" Then
                        .SubItems(24) = Invoice!CLMCash
                    End If
                    If Trim(Invoice!CLMPreXRay) <> "" Then
                        .SubItems(25) = Invoice!CLMPreXRay
                    End If
                    If Trim(Invoice!CLMPreLab) <> "" Then
                        .SubItems(26) = Invoice!CLMPreLab
                    End If
                    If Trim(Invoice!CLMPreOthers) <> "" Then
                        .SubItems(27) = Invoice!CLMPreOthers
                    End If
                    If Trim(Invoice!CLMSurgicalAmt) <> "" Then
                        .SubItems(28) = Invoice!CLMSurgicalAmt
                    End If
                    If Trim(Invoice!CLMAna) <> "" Then
                        .SubItems(29) = Invoice!CLMAna
                    End If
                    If Trim(Invoice!CLMEmergency) <> "" Then
                        .SubItems(30) = Invoice!CLMEmergency
                    End If
                    If Trim(Invoice!CLMAmbulance) <> "" Then
                        .SubItems(31) = Invoice!CLMAmbulance
                    End If
                    If Trim(Invoice!CLMOutpatient) <> "" Then
                        .SubItems(32) = Invoice!CLMOutpatient
                    End If
                    If Trim(Invoice!CLMMonthly) <> "" Then
                        .SubItems(33) = Invoice!CLMMonthly
                    End If
                    If Trim(Invoice!PreH60) <> "" Then
                        .SubItems(34) = Invoice!PreH60
                    End If
                    If Trim(Invoice!PreMedic) <> "" Then
                        .SubItems(35) = Invoice!PreMedic
                    End If
                    If Trim(Invoice!PreGP) <> "" Then
                        .SubItems(36) = Invoice!PreGP
                    End If
                    If Trim(Invoice!PreSpec) <> "" Then
                        .SubItems(37) = Invoice!PreSpec
                    End If
                    If Trim(Invoice!PostHosp) <> "" Then
                        .SubItems(38) = Invoice!PostHosp
                    End If
                    If Trim(Invoice!PreMR) <> "" Then
                        .SubItems(39) = Invoice!PreMR
                    End If
                    If Trim(Invoice!PostNotSame) <> "" Then
                        .SubItems(40) = Invoice!PostNotSame
                    End If
                    If Trim(Invoice!Others) <> "" Then
                        .SubItems(41) = Invoice!Others
                    End If
                    .SubItems(43) = Invoice!InvINdex
                    
                    If Trim(Invoice!INVPayTo) <> "" Then
                        .SubItems(44) = Invoice!INVPayTo
                    End If
                    
                    If Trim(Invoice!INVWSVersion) <> "" Then
                        .SubItems(45) = Invoice!INVWSVersion
                    End If

                    
                End With
                Invoice.MoveNext
            Loop
        End If
        Invoice.Close
        Set Invoice = Nothing

End Sub

Private Sub getDatatoWS()
    Dim ForLoop, ListCount As Integer
    Dim tA1, tA2, tA3, tA4, tA5 As Double
    Dim tA6, tA7, tA8, tA9, tA10 As Double
    Dim tA11, tA12, tA13, tA14, tA15 As Double
    Dim tA16, tA17, tA18, tA19, tA20 As Double
    Dim tA21 As Double
    Dim tM1, tM2, tM3, tM4, tM5 As Double
    Dim tM6, tM7, tM8, tM9, tM10 As Double
    Dim tM11, tM12, tM13, tM14, tM15 As Double
    Dim tM16, tM17, tM18, tM19, tM20 As Double
    Dim tM21 As Double
    ListCount = lstInvoice.listitems.count
    
    tA1 = 0
    tA2 = 0
    tA3 = 0
    tA4 = 0
    tA5 = 0
    tA8 = 0
    tA9 = 0
    tA10 = 0
    tA15 = 0
    tA16 = 0
    tA17 = 0
    tA18 = 0
    tA19 = 0
    tA20 = 0
    tA21 = 0
    tA6 = 0
    tA7 = 0
    tA11 = 0
    tA12 = 0
    tA13 = 0
    tA14 = 0
    tM1 = 0
    tM2 = 0
    tM3 = 0
    tM4 = 0
    tM5 = 0
    tM6 = 0
    tM7 = 0
    tM8 = 0
    tM9 = 0
    tM10 = 0
    tM11 = 0
    tM12 = 0
    tM13 = 0
    tM14 = 0
    tM15 = 0
    tM16 = 0
    tM17 = 0
    tM18 = 0
    tM19 = 0
    tM20 = 0
    tM21 = 0

    
    For ForLoop = 1 To ListCount
        With lstInvoice.listitems(ForLoop)
           If ChkStr(.SubItems(44)) = "H" Then
                TotalPaidBytMembertMember = CDbl(TotalPaidBytMembertMember) + CDbl(.SubItems(6))
                If Trim(.SubItems(7)) = "" Then
                    .SubItems(7) = 0
                End If
                If Trim(.SubItems(8)) = "" Then
                    .SubItems(8) = 0
                End If
                
                DB1 = CDbl(DB1) + CDbl(.SubItems(7))
                tA1 = CDbl(tA1) + CDbl(.SubItems(7) * .SubItems(8))
                
                If Trim(.SubItems(9)) = "" Then
                    .SubItems(9) = 0
                End If
                If Trim(.SubItems(10)) = "" Then
                    .SubItems(10) = 0
                End If
                 ICU1 = CDbl(ICU1) + CDbl(.SubItems(9))
                'ICUtAmount = .SubItems(10)
                tA2 = CDbl(tA2) + CDbl(.SubItems(9)) * CDbl(.SubItems(10))
                
                If Trim(.SubItems(11)) = "" Then
                    .SubItems(11) = 0
                End If
                
                tA3 = CDbl(tA3) + .SubItems(11)
                If Trim(.SubItems(12)) = "" Then
                    .SubItems(12) = 0
                End If
    
                tA4 = CDbl(tA4) + .SubItems(12)
                If Trim(.SubItems(13)) = "" Then
                    .SubItems(13) = 0
                End If
                
                tA8 = CDbl(tA8) + .SubItems(13)
                If Trim(.SubItems(14)) = "" Then
                    .SubItems(14) = 0
                End If
                
                tA9 = CDbl(tA9) + .SubItems(14)
                
                If Trim(.SubItems(15)) = "" Then
                    .SubItems(15) = 0
                End If
                
                tA10 = CDbl(tA10) + .SubItems(15)
                
                If Trim(.SubItems(16)) = "" Then
                    .SubItems(16) = 0
                End If
                If Trim(.SubItems(17)) = "" Then
                    .SubItems(17) = 0
                End If
                
                tA15 = CDbl(tA15) + .SubItems(16)
                tA16 = CDbl(tA16) + .SubItems(17)
                If Trim(.SubItems(18)) = "" Then
                    .SubItems(18) = 0
                End If
                If Trim(.SubItems(19)) = "" Then
                    .SubItems(19) = 0
                End If
                txtLodgerDay = CDbl(txtLodgerDay) + .SubItems(18)
    
                tA17 = CDbl(tA17) + .SubItems(18) * .SubItems(19)
                
                If Trim(.SubItems(20)) = "" Then
                    .SubItems(20) = 0
                End If
                If Trim(.SubItems(21)) = "" Then
                    .SubItems(21) = 0
                End If
                tA18 = CDbl(tA18) + .SubItems(20)
                tA19 = CDbl(tA19) + .SubItems(21)
                If Trim(.SubItems(22)) = "" Then
                    .SubItems(22) = 0
                End If
                tA20 = CDbl(tA20) + .SubItems(22)
                
                If Trim(.SubItems(23)) = "" Then
                    .SubItems(23) = 0
                End If
                If Trim(.SubItems(24)) = "" Then
                    .SubItems(24) = 0
                End If
                txtDayCash = CDbl(txtDayCash) + .SubItems(23)
                tA21 = CDbl(tA21) + .SubItems(23) * .SubItems(24)
                
                
                If Trim(.SubItems(25)) = "" Then
                    .SubItems(25) = 0
                End If
                If Trim(.SubItems(26)) = "" Then
                    .SubItems(26) = 0
                End If
                If Trim(.SubItems(27)) = "" Then
                    .SubItems(27) = 0
                End If
                            
                tA5 = CLng(tA5) + CLng(.SubItems(25)) + CLng(.SubItems(26)) + CLng(.SubItems(27))
                
                If Trim(.SubItems(28)) = "" Then
                    .SubItems(28) = 0
                End If
                If Trim(.SubItems(29)) = "" Then
                    .SubItems(29) = 0
                End If
                tA6 = CDbl(tA6) + .SubItems(28)
                tA7 = CDbl(tA7) + .SubItems(29)
                If Trim(.SubItems(30)) = "" Then
                    .SubItems(30) = 0
                End If
                If Trim(.SubItems(31)) = "" Then
                    .SubItems(31) = 0
                End If
                tA11 = CDbl(tA11) + .SubItems(30)
                tA12 = CDbl(tA12) + .SubItems(31)
                If Trim(.SubItems(32)) = "" Then
                    .SubItems(32) = 0
                End If
                If Trim(.SubItems(33)) = "" Then
                    .SubItems(33) = 0
                End If
                tA13 = CDbl(tA13) + .SubItems(32)
                tA14 = CDbl(tA14) + .SubItems(33)
            
            ' ########################################## Pay To tMBtMber
            ElseIf ChkStr(.SubItems(44)) = "M" Then
                TotalPaidBytMembertMember = CDbl(TotalPaidBytMembertMember) + CDbl(.SubItems(6))
                If Trim(.SubItems(7)) = "" Then
                    .SubItems(7) = 0
                End If
                If Trim(.SubItems(8)) = "" Then
                    .SubItems(8) = 0
                End If
                
                DB1 = CDbl(DB1) + CDbl(.SubItems(7))
                tM1 = CDbl(tM1) + CDbl(.SubItems(7) * .SubItems(8))
                
                If Trim(.SubItems(9)) = "" Then
                    .SubItems(9) = 0
                End If
                If Trim(.SubItems(10)) = "" Then
                    .SubItems(10) = 0
                End If
                 ICU1 = CDbl(ICU1) + CDbl(.SubItems(9))
                'ICUtAmount = .SubItems(10)
                tM2 = CDbl(tM2) + CDbl(.SubItems(9)) * CDbl(.SubItems(10))
                
                If Trim(.SubItems(11)) = "" Then
                    .SubItems(11) = 0
                End If
                
                tM3 = CDbl(tM3) + .SubItems(11)
                If Trim(.SubItems(12)) = "" Then
                    .SubItems(12) = 0
                End If
    
                tM4 = CDbl(tM4) + .SubItems(12)
                If Trim(.SubItems(13)) = "" Then
                    .SubItems(13) = 0
                End If
                
                tM8 = CDbl(tM8) + .SubItems(13)
                If Trim(.SubItems(14)) = "" Then
                    .SubItems(14) = 0
                End If
                
                tM9 = CDbl(tM9) + .SubItems(14)
                
                If Trim(.SubItems(15)) = "" Then
                    .SubItems(15) = 0
                End If
                
                tM10 = CDbl(tM10) + .SubItems(15)
                
                If Trim(.SubItems(16)) = "" Then
                    .SubItems(16) = 0
                End If
                If Trim(.SubItems(17)) = "" Then
                    .SubItems(17) = 0
                End If
                
                tM15 = CDbl(tM15) + .SubItems(16)
                tM16 = CDbl(tM16) + .SubItems(17)
                If Trim(.SubItems(18)) = "" Then
                    .SubItems(18) = 0
                End If
                If Trim(.SubItems(19)) = "" Then
                    .SubItems(19) = 0
                End If
                txtLodgerDay = CDbl(txtLodgerDay) + .SubItems(18)
    
                tM17 = CDbl(tM17) + .SubItems(18) * .SubItems(19)
                
                If Trim(.SubItems(20)) = "" Then
                    .SubItems(20) = 0
                End If
                If Trim(.SubItems(21)) = "" Then
                    .SubItems(21) = 0
                End If
                tM18 = CDbl(tM18) + .SubItems(20)
                tM19 = CDbl(tM19) + .SubItems(21)
                If Trim(.SubItems(22)) = "" Then
                    .SubItems(22) = 0
                End If
                tM20 = CDbl(tM20) + .SubItems(22)
                
                If Trim(.SubItems(23)) = "" Then
                    .SubItems(23) = 0
                End If
                If Trim(.SubItems(24)) = "" Then
                    .SubItems(24) = 0
                End If
                txtDayCash = CDbl(txtDayCash) + .SubItems(23)
                tM21 = CDbl(tM21) + .SubItems(23) * .SubItems(24)
                
                
                If Trim(.SubItems(25)) = "" Then
                    .SubItems(25) = 0
                End If
                If Trim(.SubItems(26)) = "" Then
                    .SubItems(26) = 0
                End If
                If Trim(.SubItems(27)) = "" Then
                    .SubItems(27) = 0
                End If
                            
                tM5 = CLng(tM5) + CLng(.SubItems(25)) + CLng(.SubItems(26)) + CLng(.SubItems(27))
                
                If Trim(.SubItems(28)) = "" Then
                    .SubItems(28) = 0
                End If
                If Trim(.SubItems(29)) = "" Then
                    .SubItems(29) = 0
                End If
                tM6 = CDbl(tM6) + .SubItems(28)
                tM7 = CDbl(tM7) + .SubItems(29)
                If Trim(.SubItems(30)) = "" Then
                    .SubItems(30) = 0
                End If
                If Trim(.SubItems(31)) = "" Then
                    .SubItems(31) = 0
                End If
                tM11 = CDbl(tM11) + .SubItems(30)
                tM12 = CDbl(tM12) + .SubItems(31)
                If Trim(.SubItems(32)) = "" Then
                    .SubItems(32) = 0
                End If
                If Trim(.SubItems(33)) = "" Then
                    .SubItems(33) = 0
                End If
                tM13 = CDbl(tM13) + .SubItems(32)
                tM14 = CDbl(tM14) + .SubItems(33)
            
            End If
        End With
    Next
        
    A1 = Format(tA1, "#,#0.00")
    A2 = Format(tA2, "#,#0.00")
    A3 = Format(tA3, "#,#0.00")
    A4 = Format(tA4, "#,#0.00")
    A5 = Format(tA5, "#,#0.00")
    A6 = Format(tA6, "#,#0.00")
    A7 = Format(tA7, "#,#0.00")
    A8 = Format(tA8, "#,#0.00")
    A9 = Format(tA9, "#,#0.00")
    A10 = Format(tA10, "#,#0.00")
    A11 = Format(tA11, "#,#0.00")
    A12 = Format(tA12, "#,#0.00")
    A13 = Format(tA13, "#,#0.00")
    A14 = Format(tA14, "#,#0.00")
    A15 = Format(tA15, "#,#0.00")
    A16 = Format(tA16, "#,#0.00")
    A17 = Format(tA17, "#,#0.00")
    A18 = Format(tA18, "#,#0.00")
    A19 = Format(tA19, "#,#0.00")
    A20 = Format(tA20, "#,#0.00")
    A21 = Format(tA21, "#,#0.00")
    
    M1 = tM1
    M2 = tM2
    M3 = tM3
    M4 = tM4
    M5 = tM5
    M6 = tM6
    M7 = tM7
    M8 = tM8
    M9 = tM9
    M10 = tM10
    M11 = tM11
    M12 = tM12
    M13 = tM13
    M14 = tM14
    M15 = tM15
    M16 = tM16
    M17 = tM17
    M18 = tM18
    M19 = tM19
    M20 = tM20
    M21 = tM21
    
    Call A3_LostFocus
    Call A4_LostFocus
    Call A5_LostFocus
    Call A6_LostFocus
    Call A7_LostFocus
    Call A8_LostFocus
    Call A9_LostFocus
    Call A10_LostFocus
    Call A11_LostFocus
    Call A12_LostFocus
    Call A13_LostFocus
    Call A14_LostFocus
    Call A15_LostFocus
    Call A16_LostFocus

    Call A18_LostFocus
    
    Call A20_LostFocus
    
    Call A21_LostFocus
    Call A22_LostFocus
    Call A23_LostFocus
    
    
End Sub




' **************************************************

Private Sub CmdInvoice_Click()
    If FormLoadedByName("frmInvoice") = True Then
        frmInvoice.SetFocus
    Else
        frmInvoice.show
    End If
End Sub


Private Sub cmdShowHistory_Click()
    If Trim(txtCASNumber.Text) = "" And Trim(txtMBMNumber.Text) = "" Then
        MsgBox "Please key in Case Number or Member Number first!", vbCritical + vbOKOnly
    Else
        With frmClaimHistory
            .txtCASNumber = txtCASNumber.Text
            .txtMBMNumber = txtMBMNumber.Text
            Call .ListCases("O", txtMBMNumber, txtCASNumber)
            Call .ListCases("C", txtMBMNumber, txtCASNumber)
            Call .ListCases("R", txtMBMNumber, txtCASNumber)
            .show
        End With
    End If
End Sub

' ############ Scroll Bar ##################
Public Sub ScrollForm2()
    Frame11.Top = Frame11.Top + intLastPoz2 - VScroll2.value
    intLastPoz2 = VScroll2.value
End Sub

Private Sub VScroll2_Change()
    ScrollForm2
End Sub

Private Sub VScroll2_Scroll()
    ScrollForm2
End Sub

' ### calculation  functions
' Section  Limit
'Section  ON Enter ==== LIMIT
Private Sub DB1_lostfocus()
    If IsNumeric(RoomAmount) = True And IsNumeric(DB1) = True Then
'        MsgBox "Please Enter No of Days for Room And Board", vbOKOnly + vbCritical, DialogBoxTitle
 '       DB1 = 0
  '      DB1.SetFocus
   ' ElseIf Trim(RoomAmount) = "" Then
    '    MsgBox "Please Enter Room Amount for Room And Board", vbOKOnly + vbCritical, DialogBoxTitle
     '   RoomAmount = 0
      '  RoomAmount.SetFocus
    'Else
        Call CalculateRoomLimit
        Call CalculateTax
    End If
End Sub

Private Sub RoomAmount_LOSTFOCUS()
        If IsNumeric(RoomAmount) = True And IsNumeric(DB1) = True Then
'            MsgBox "Please Enter No of Days for Room And Board", vbOKOnly + vbCritical, DialogBoxTitle
 '           DB1 = 0
  '          DB1.SetFocus
   '     ElseIf Trim(RoomAmount) = "" Then
    '        MsgBox "Please Enter Room Amount for Room And Board", vbOKOnly + vbCritical, DialogBoxTitle
     '       RoomAmount = 0
      '      RoomAmount.SetFocus
       ' Else
            Call CalculateRoomLimit
            Call CalculateTax
        End If
End Sub

Private Sub ICU1_lostfocus()
    If IsNumeric(ICU1) = True And IsNumeric(ICUAmount) = True Then
'        MsgBox "Please Enter No of Days for Intensive Care Unit", vbOKOnly + vbCritical, DialogBoxTitle
 '       ICU1 = 0
  ''      ICU1.SetFocus
'    ElseIf Trim(ICUAmount) = "" Then
 '       MsgBox "Please Enter Room Amount for Intensive Care Unit", vbOKOnly + vbCritical, DialogBoxTitle
  '      ICUAmount = 0
   '     ICUAmount.SetFocus
'    Else
        Call CalculateICULimit
        Call CalculateTax
    End If
End Sub

Private Sub ICUAmount_lostfocus()
    If IsNumeric(ICU1) = True And IsNumeric(ICUAmount) = True Then
    '    MsgBox "Please Enter No of Days for Intensive Care Unit", vbOKOnly + vbCritical, DialogBoxTitle
     '   ICU1 = 0
      '  ICU1.SetFocus
    'ElseIf Trim(ICUAmount) = "" Then
     '   MsgBox "Please Enter Room Amount for Intensive Care Unit", vbOKOnly + vbCritical, DialogBoxTitle
      '  ICUAmount = 0
      ' ICUAmount.SetFocus
    'Else
        Call CalculateICULimit
        Call CalculateTax
    End If
End Sub

Private Sub txtKidneyMth_lostfocus()
    If IsNumeric(txtKidneyMth) = True And IsNumeric(txtKidneyRM) Then
        P14 = Format(CDbl(txtKidneyMth) * CDbl(txtKidneyRM), "#,#0.00")
    End If
End Sub

Private Sub txtKidneyRM_lostfocus()
    If IsNumeric(txtKidneyMth) = True And IsNumeric(txtKidneyRM) Then
        P14 = Format(CDbl(txtKidneyMth) * CDbl(txtKidneyRM), "#,#0.00")
    End If
End Sub

Private Sub txtDayCash_lostfocus()
 '   If Trim(txtDayCash) = "" Then
'        MsgBox "Please Enter No of Days for Cash Allowance", vbOKOnly + vbCritical, DialogBoxTitle
 '       txtDayCash = 0
  '      txtDayCash.SetFocus
   ' ElseIf Trim(txtRMCash) = "" Then
    '    MsgBox "Please Enter the Amount for Cash Allowance", vbOKOnly + vbCritical, DialogBoxTitle
     '   txtRMCash = 0
      '  txtRMCash.SetFocus
'    Else
    If IsNumeric(txtDayCash) = True And IsNumeric(txtRMCash) Then
        P16 = Format(CDbl(txtDayCash) * CDbl(txtRMCash), "#,#0.00")
'        txtDayCash.SetFocus
    End If
  '  End If
End Sub

Private Sub txtRMCash_LostFocus()
'        If Trim(txtDayCash) = "" Then
 '           MsgBox "Please Enter No of Days for Cash Allowance", vbOKOnly + vbCritical, DialogBoxTitle
  '          txtDayCash = 0
   '         txtDayCash.SetFocus
    '    ElseIf Trim(txtRMCash) = "" Then
     '       MsgBox "Please Enter the Amount for Cash Allowance", vbOKOnly + vbCritical, DialogBoxTitle
      '      txtRMCash = 0
       '     txtRMCash.SetFocus
        'Else
         '   If IsNumeric(txtDayCash) = True And IsNumeric(txtRMCash) Then
          '      P21 = Format(CDbl(txtDayCash) * CDbl(txtRMCash), "#,#0.00")
           '     txtDayCash.SetFocus
            'End If
        'End If
    If IsNumeric(txtDayCash) = True And IsNumeric(txtRMCash) Then
        P16 = Format(CDbl(txtDayCash) * CDbl(txtRMCash), "#,#0.00")
    End If
End Sub


Private Sub LG1_lostfocus()
    If IsNumeric(LG1) = True And IsNumeric(LodgetAmount) = True Then
    '    If Trim(LG1) = "" Then
     '       MsgBox "Please Enter No of Days for Lodger", vbOKOnly + vbCritical, DialogBoxTitle
      '      LG1 = 0
       '     LG1.SetFocus
'        ElseIf Trim(LodgerAmount) = "" Then
 '           MsgBox "Please Enter Lodger Amount for Lodger", vbOKOnly + vbCritical, DialogBoxTitle
  '          LodgerAmount = 0
   '         LodgerAmount.SetFocus
    '    Else
        Call CalculateLodgerLimit
     '   End If
    End If
End Sub

Private Sub LodgerAmount_lostfocus()
    If IsNumeric(LG1) = True And IsNumeric(LodgetAmount) = True Then
     '   MsgBox "Please Enter No of Days for Lodger", vbOKOnly + vbCritical, DialogBoxTitle
      '  LG1 = 0
       ' LG1.SetFocus
    'ElseIf Trim(LodgerAmount) = "" Then
     '   MsgBox "Please Enter Lodger Amount for Lodger", vbOKOnly + vbCritical, DialogBoxTitle
      '  LodgerAmount = 0
       ' LodgerAmount.SetFocus
    'Else
        Call CalculateLodgerLimit
    End If
End Sub



' $########### Calculation for limit

Private Sub CalculateRoomLimit()
    Dim LimitAmount As Double
    If IsNumeric(RoomAmount) Then
        LimitAmount = DB1 * RoomAmount
    '    If LimitAmount > CDbl(P1A) Then
     '       P1 = P1A
      '  Else
         P1 = Format(LimitAmount, "#,#0.00")
       ' End If
    End If
End Sub


Private Sub CalculateICULimit()
    Dim ICULimitAmount
    ICULimitAmount = ICU1 * ICUAmount
'    If ICULimitAmount > CDbl(P2A) Then
 '       P2 = P2A
  '  Else
        P2 = Format(ICULimitAmount, "#,#0.00")
   ' End If
End Sub

 
Private Sub CalculateLodgerLimit()
    Dim LodgerLimitAmount As Double
    
    LodgerLimitAmount = LG1 * LodgerAmount
   '     If LodgerLimitAmount > CDbl(P17A) Then
    '        P17 = P17A
     '   Else
    P18 = Format(LodgerLimitAmount, "#,#0.00")
      '  End If
End Sub

Private Sub CalculateTax()
'On Error Resume Next
    If IsNumeric(P1) = True And IsNumeric(P2) = True Then
        P19 = Format((CDbl(P1) + CDbl(P2)) * 0.05, "#,#0.00")
        Call SumActual
    End If

End Sub

' Section Approved Amount
' ON lost focus for Actual Value
Private Sub A1_LostFocus()
    If IsNumeric(A1) Then
        Call ApprovedRoom
    End If
    M1 = Format(CDbl(IIf(IsNumeric(M1), M1, 0)) + CDbl(IIf(IsNumeric(AP1), AP1, 0)) - CDbl(IIf(IsNumeric(H1), H1, 0)), "#,#0.00;(#,#0.00)")
    Call SumActual
End Sub

Private Sub A2_LostFocus()
    If IsNumeric(A2) Then
        Call ApprovedICU
    End If
    M2 = Format(CDbl(IIf(IsNumeric(M2), M2, 0)) + CDbl(IIf(IsNumeric(AP2), AP2, 0)) - CDbl(IIf(IsNumeric(H2), H2, 0)), "#,#0.00;(#,#0.00)")
    Call SumActual
End Sub

Private Sub A3_LostFocus()
    If IsNumeric(A3) Then
        Call ApproveHSS
    End If
    M3 = Format(CDbl(IIf(IsNumeric(M3), M3, 0)) + CDbl(IIf(IsNumeric(AP3), AP3, 0)) - CDbl(IIf(IsNumeric(H3), H3, 0)), "#,#0.00;(#,#0.00)")
    Call SumActual
End Sub

Private Sub A4_LostFocus()
    If IsNumeric(A4) Then
        Call ApproveOT
    End If
    M4 = Format(CDbl(IIf(IsNumeric(M4), M4, 0)) + CDbl(IIf(IsNumeric(AP4), AP4, 0)) - CDbl(IIf(IsNumeric(H4), H4, 0)), "#,#0.00;(#,#0.00)")
    Call SumActual
End Sub

Private Sub A5_LostFocus()
    If IsNumeric(A5) Then
        Call ApprovePHD
    End If
    M5 = Format(CDbl(IIf(IsNumeric(M5), M5, 0)) + CDbl(IIf(IsNumeric(AP5), AP5, 0)) - CDbl(IIf(IsNumeric(H5), H5, 0)), "#,#0.00;(#,#0.00)")
    Call SumActual
End Sub

Private Sub A6_LostFocus()
    If IsNumeric(A6) Then
        Call ApprovedSurgical
    End If
    M6 = Format(CDbl(IIf(IsNumeric(M6), M6, 0)) + CDbl(IIf(IsNumeric(AP6), AP6, 0)) - CDbl(IIf(IsNumeric(H6), H6, 0)), "#,#0.00;(#,#0.00)")
    Call SumActual
End Sub

Private Sub A7_LostFocus()
    If IsNumeric(A7) Then
        Call ApprovedAna
    End If
    M7 = Format(CDbl(IIf(IsNumeric(M7), M7, 0)) + CDbl(IIf(IsNumeric(AP7), AP7, 0)) - CDbl(IIf(IsNumeric(H7), H7, 0)), "#,#0.00;(#,#0.00)")
    Call SumActual
End Sub

Private Sub A8_LostFocus()
    If IsNumeric(A8) Then
        Call ApprovePSpec
    End If
    M8 = Format(CDbl(IIf(IsNumeric(M8), M8, 0)) + CDbl(IIf(IsNumeric(AP8), AP8, 0)) - CDbl(IIf(IsNumeric(H8), H8, 0)), "#,#0.00;(#,#0.00)")
    Call SumActual
End Sub

Private Sub A9_LostFocus()
    If IsNumeric(A9) Then
        Call ApproveInH
    End If
    M9 = Format(CDbl(IIf(IsNumeric(M9), M9, 0)) + CDbl(IIf(IsNumeric(AP9), AP9, 0)) - CDbl(IIf(IsNumeric(H9), H9, 0)), "#,#0.00;(#,#0.00)")
    Call SumActual
End Sub

Private Sub A10_LostFocus()
    If IsNumeric(A10) Then
        Call ApprovePostH
    End If
    M10 = Format(CDbl(IIf(IsNumeric(M10), M10, 0)) + CDbl(IIf(IsNumeric(AP10), AP10, 0)) - CDbl(IIf(IsNumeric(H10), H10, 0)), "#,#0.00;(#,#0.00)")
    Call SumActual
End Sub

Private Sub A11_LostFocus()
    If IsNumeric(A11) Then
         Call ApproveEOut
    End If
    M11 = Format(CDbl(IIf(IsNumeric(M11), M11, 0)) + CDbl(IIf(IsNumeric(AP11), AP11, 0)) - CDbl(IIf(IsNumeric(H11), H11, 0)), "#,#0.00;(#,#0.00)")
    Call SumActual
End Sub

Private Sub A12_LostFocus()
    If IsNumeric(A12) Then
       Call ApproveAF
    End If
    M12 = Format(CDbl(IIf(IsNumeric(M12), M12, 0)) + CDbl(IIf(IsNumeric(AP12), AP12, 0)) - CDbl(IIf(IsNumeric(H12), H12, 0)), "#,#0.00;(#,#0.00)")
    Call SumActual
End Sub

Private Sub A13_LostFocus()
    If IsNumeric(A13) Then
        Call ApprovePT
    End If
    M13 = Format(CDbl(IIf(IsNumeric(M13), M13, 0)) + CDbl(IIf(IsNumeric(AP13), AP13, 0)) - CDbl(IIf(IsNumeric(H13), H13, 0)), "#,#0.00;(#,#0.00)")
    Call SumActual
End Sub

Private Sub A14_LostFocus()
    If IsNumeric(A14) Then
        Call ApprovedKidney
    End If
    M14 = Format(CDbl(IIf(IsNumeric(M14), M14, 0)) + CDbl(IIf(IsNumeric(AP14), AP14, 0)) - CDbl(IIf(IsNumeric(H14), H14, 0)), "#,#0.00;(#,#0.00)")
    Call SumActual
End Sub

Private Sub A15_LostFocus()
    If IsNumeric(A15) Then
        Call ApprovedOrgan
        
    End If
    M15 = Format(CDbl(IIf(IsNumeric(M15), M15, 0)) + CDbl(IIf(IsNumeric(AP15), AP15, 0)) - CDbl(IIf(IsNumeric(H15), H15, 0)), "#,#0.00;(#,#0.00)")
    Call SumActual
End Sub

Private Sub A16_LostFocus()
    If IsNumeric(A16) Then
        Call ApproveCash
        
    End If
    M16 = Format(CDbl(IIf(IsNumeric(M16), M16, 0)) + CDbl(IIf(IsNumeric(AP16), AP16, 0)) - CDbl(IIf(IsNumeric(H16), H16, 0)), "#,#0.00;(#,#0.00)")
    Call SumActual
End Sub

Private Sub A17_LostFocus()
    If IsNumeric(A17) Then
        Call ApprovedTax
    End If
    M17 = Format(CDbl(IIf(IsNumeric(M17), M17, 0)) + CDbl(IIf(IsNumeric(AP17), AP17, 0)) - CDbl(IIf(IsNumeric(H17), H17, 0)), "#,#0.00;(#,#0.00)")
    Call SumActual
End Sub

Private Sub A18_LostFocus()
    If IsNumeric(A18) Then
        Call ApprovedLodger
    End If
    M18 = Format(CDbl(IIf(IsNumeric(M18), M18, 0)) + CDbl(IIf(IsNumeric(AP18), AP18, 0)) - CDbl(IIf(IsNumeric(H18), H18, 0)), "#,#0.00;(#,#0.00)")
    Call SumActual
End Sub

Private Sub A19_LostFocus()
    If IsNumeric(A19) Then
        Call ApproveMR
    End If
    M19 = Format(CDbl(IIf(IsNumeric(M19), M19, 0)) + CDbl(IIf(IsNumeric(AP19), AP19, 0)) - CDbl(IIf(IsNumeric(H19), H19, 0)), "#,#0.00;(#,#0.00)")
    Call SumActual
End Sub

Private Sub A20_LostFocus()
    If IsNumeric(A20) Then
        Call ApprovePhone
    End If
    M20 = Format(CDbl(IIf(IsNumeric(M20), M20, 0)) + CDbl(IIf(IsNumeric(AP20), AP20, 0)) - CDbl(IIf(IsNumeric(H20), H20, 0)), "#,#0.00;(#,#0.00)")
    Call SumActual
End Sub

Private Sub A21_LostFocus()
    If IsNumeric(A21) Then
        Call ApproveRegistration
    End If
    M21 = Format(CDbl(IIf(IsNumeric(M21), M21, 0)) + CDbl(IIf(IsNumeric(AP21), AP21, 0)) - CDbl(IIf(IsNumeric(H21), H21, 0)), "#,#0.00;(#,#0.00)")
    Call SumActual
End Sub

Private Sub A22_LostFocus()
    If IsNumeric(A22) Then
         Call ApprovedOthers
    End If
    M22 = Format(CDbl(IIf(IsNumeric(M22), M22, 0)) + CDbl(IIf(IsNumeric(AP22), AP22, 0)) - CDbl(IIf(IsNumeric(H22), H22, 0)), "#,#0.00;(#,#0.00)")
    Call SumActual
End Sub

Private Sub A23_LostFocus()
    If IsNumeric(A23) Then
        Call ApprovedOthers2
    End If
    M23 = Format(CDbl(IIf(IsNumeric(M23), M23, 0)) + CDbl(IIf(IsNumeric(AP23), AP23, 0)) - CDbl(IIf(IsNumeric(H23), H23, 0)), "#,#0.00;(#,#0.00)")
    Call SumActual
End Sub


' ON lost focus for Benefit Limit
Private Sub P1_LostFocus()
    If IsNumeric(P1) Then
        Call ApprovedRoom
    End If
    M1 = Format(CDbl(IIf(IsNumeric(PP1), PP1, 0)) - CDbl(IIf(IsNumeric(H1), H1, 0)), "#,#0.00;(#,#0.00)")
    Call SumActual
End Sub

Private Sub P2_LostFocus()
    If IsNumeric(P2) Then
        Call ApprovedICU
    End If
    M2 = Format(CDbl(IIf(IsNumeric(PP2), PP2, 0)) - CDbl(IIf(IsNumeric(H2), H2, 0)), "#,#0.00;(#,#0.00)")
    Call SumActual
End Sub

Private Sub P3_LostFocus()
    If IsNumeric(P3) Then
        Call ApproveHSS
    End If
    M3 = Format(CDbl(IIf(IsNumeric(PP3), PP3, 0)) - CDbl(IIf(IsNumeric(H3), H3, 0)), "#,#0.00;(#,#0.00)")
    Call SumActual
End Sub

Private Sub P4_LostFocus()
    If IsNumeric(P4) Then
        Call ApproveOT
    End If
    M4 = Format(CDbl(IIf(IsNumeric(PP4), PP4, 0)) - CDbl(IIf(IsNumeric(H4), H4, 0)), "#,#0.00;(#,#0.00)")
    Call SumActual
End Sub

Private Sub P5_LostFocus()
    If IsNumeric(P5) Then
        Call ApprovePHD
    End If
    M5 = Format(CDbl(IIf(IsNumeric(PP5), PP5, 0)) - CDbl(IIf(IsNumeric(H5), H5, 0)), "#,#0.00;(#,#0.00)")
    Call SumActual
End Sub

Private Sub P6_LostFocus()
    If IsNumeric(P6) Then
        Call ApprovedSurgical
    End If
    M6 = Format(CDbl(IIf(IsNumeric(PP6), PP6, 0)) - CDbl(IIf(IsNumeric(H6), H6, 0)), "#,#0.00;(#,#0.00)")
    Call SumActual
End Sub

Private Sub P7_LostFocus()
    'If IsNumeric(P7) Then
        'Call ApprovedPna
    'End If
    'M7 = Format(CDbl(IIf(IsNumeric(PP7), PP7, 0)) - CDbl(IIf(IsNumeric(H7), H7, 0)), "#,#0.00;(#,#0.00)")
    'Call SumActual
End Sub

Private Sub P8_LostFocus()
    If IsNumeric(P8) Then
        Call ApprovePSpec
    End If
    M8 = Format(CDbl(IIf(IsNumeric(PP8), PP8, 0)) - CDbl(IIf(IsNumeric(H8), H8, 0)), "#,#0.00;(#,#0.00)")
    Call SumActual
End Sub

Private Sub P9_LostFocus()
    If IsNumeric(P9) Then
        Call ApproveInH
    End If
    M9 = Format(CDbl(IIf(IsNumeric(PP9), PP9, 0)) - CDbl(IIf(IsNumeric(H9), H9, 0)), "#,#0.00;(#,#0.00)")
    Call SumActual
End Sub

Private Sub P10_LostFocus()
    If IsNumeric(P10) Then
        Call ApprovePostH
    End If
    M10 = Format(CDbl(IIf(IsNumeric(PP10), PP10, 0)) - CDbl(IIf(IsNumeric(H10), H10, 0)), "#,#0.00;(#,#0.00)")
    Call SumActual
End Sub

Private Sub P11_LostFocus()
    If IsNumeric(P11) Then
         Call ApproveEOut
    End If
    M11 = Format(CDbl(IIf(IsNumeric(PP11), PP11, 0)) - CDbl(IIf(IsNumeric(H11), H11, 0)), "#,#0.00;(#,#0.00)")
    Call SumActual
End Sub

Private Sub P12_LostFocus()
    'If IsNumeric(P12) Then
    '   Call ApprovePF
    'End If
    'M12 = Format(CDbl(IIf(IsNumeric(PP12), PP12, 0)) - CDbl(IIf(IsNumeric(H12), H12, 0)), "#,#0.00;(#,#0.00)")
    'Call SumActual
End Sub

Private Sub P13_LostFocus()
    If IsNumeric(P13) Then
        Call ApprovePT
    End If
    M13 = Format(CDbl(IIf(IsNumeric(PP13), PP13, 0)) - CDbl(IIf(IsNumeric(H13), H13, 0)), "#,#0.00;(#,#0.00)")
    Call SumActual
End Sub

Private Sub P14_LostFocus()
    If IsNumeric(P14) Then
        Call ApprovedKidney
    End If
    M14 = Format(CDbl(IIf(IsNumeric(PP14), PP14, 0)) - CDbl(IIf(IsNumeric(H14), H14, 0)), "#,#0.00;(#,#0.00)")
    Call SumActual
End Sub

Private Sub P15_LostFocus()
    If IsNumeric(P15) Then
        Call ApprovedOrgan
        
    End If
    M15 = Format(CDbl(IIf(IsNumeric(PP15), PP15, 0)) - CDbl(IIf(IsNumeric(H15), H15, 0)), "#,#0.00;(#,#0.00)")
    Call SumActual
End Sub

Private Sub P16_LostFocus()
    If IsNumeric(P16) Then
        Call ApproveCash
        
    End If
    M16 = Format(CDbl(IIf(IsNumeric(PP16), PP16, 0)) - CDbl(IIf(IsNumeric(H16), H16, 0)), "#,#0.00;(#,#0.00)")
    Call SumActual
End Sub

Private Sub P17_LostFocus()
    If IsNumeric(P17) Then
        Call ApprovedTax
    End If
    M17 = Format(CDbl(IIf(IsNumeric(PP17), PP17, 0)) - CDbl(IIf(IsNumeric(H17), H17, 0)), "#,#0.00;(#,#0.00)")
    Call SumActual
End Sub

Private Sub P18_LostFocus()
    If IsNumeric(P18) Then
        Call ApprovedLodger
    End If
    M18 = Format(CDbl(IIf(IsNumeric(PP18), PP18, 0)) - CDbl(IIf(IsNumeric(H18), H18, 0)), "#,#0.00;(#,#0.00)")
    Call SumActual
End Sub

Private Sub P19_LostFocus()
    If IsNumeric(P19) Then
        Call ApproveMR
    End If
    M19 = Format(CDbl(IIf(IsNumeric(PP19), PP19, 0)) - CDbl(IIf(IsNumeric(H19), H19, 0)), "#,#0.00;(#,#0.00)")
    Call SumActual
End Sub

Private Sub P20_LostFocus()
    If IsNumeric(P20) Then
        Call ApprovePhone
    End If
    M20 = Format(CDbl(IIf(IsNumeric(PP20), PP20, 0)) - CDbl(IIf(IsNumeric(H20), H20, 0)), "#,#0.00;(#,#0.00)")
    Call SumActual
End Sub

Private Sub P21_LostFocus()
    If IsNumeric(P21) Then
        Call ApproveRegistration
    End If
    M21 = Format(CDbl(IIf(IsNumeric(PP21), PP21, 0)) - CDbl(IIf(IsNumeric(H21), H21, 0)), "#,#0.00;(#,#0.00)")
    Call SumActual
End Sub

Private Sub P22_LostFocus()
    If IsNumeric(P22) Then
         Call ApprovedOthers
    End If
    M22 = Format(CDbl(IIf(IsNumeric(PP22), PP22, 0)) - CDbl(IIf(IsNumeric(H22), H22, 0)), "#,#0.00;(#,#0.00)")
    Call SumActual
End Sub

Private Sub P23_LostFocus()
    'If IsNumeric(P23) Then
    '    Call ApproveOther1
    'End If
    'M23 = Format(CDbl(IIf(IsNumeric(PP23), PP23, 0)) - CDbl(IIf(IsNumeric(H23), H23, 0)), "#,#0.00;(#,#0.00)")
    'Call SumActual
End Sub

' End Lost Focus FOR LIMIT SECTION
' ###################################################3
' --------------------- START APPROVED SECTION --------------

' Start Approved Functions
Private Sub ApprovedRoom()
    Dim tA1, tP1 As Double
    tA1 = IIf(IsNumeric(A1), A1, 0)
    tP1 = IIf(IsNumeric(P1), P1, 0)
    If ChkStr(P1) = "FULL" Then
        AP1 = Format(tA1, "#,#0.00")
    ElseIf CDbl(tA1) > CDbl(tP1) Then
        AP1 = Format(tP1, "#,#0.00")
    Else
        AP1 = Format(tA1, "#,#0.00")
    End If
    H1 = Format(tA1, "#,#0.00")
End Sub


Private Sub ApprovedICU()
'On Error Resume Next
    Dim tA2, tP2 As Double
    tA2 = IIf(IsNumeric(A2), A2, 0)
    tP2 = IIf(IsNumeric(P2), P2, 0)
    If ChkStr(P2) = "FULL" Then
        AP2 = Format(tA2, "#,#0.00")
    ElseIf CDbl(tA2) > CDbl(tP2) Then
        AP2 = Format(tP2, "#,#0.00")
    Else
        AP2 = Format(tA2, "#,#0.00")
    End If
    H2 = Format(tA2, "#,#0.00")
End Sub

Private Sub ApproveHSS()
'On Error Resume Next
    Dim tA3, tP3 As Double
    tA3 = IIf(IsNumeric(A3), A3, 0)
    tP3 = IIf(IsNumeric(P3), P3, 0)
    If ChkStr(P3) = "FULL" Then
        AP3 = Format(tA3, "#,#0.00")
    ElseIf CDbl(tA3) > CDbl(tP3) Then
        AP3 = Format(tP3, "#,#0.00")
    Else
        AP3 = Format(tA3, "#,#0.00")
    End If
    H3 = Format(tA3, "#,#0.00")
End Sub

Private Sub ApproveOT()
'On Error Resume Next
    Dim tA4, tP4 As Double
    tA4 = IIf(IsNumeric(A4), A4, 0)
    tP4 = IIf(IsNumeric(P4), P4, 0)
    If ChkStr(P4) = "FULL" Then
        AP4 = Format(tA4, "#,#0.00")
    ElseIf CDbl(tA4) > CDbl(tP4) Then
        AP4 = Format(tP4, "#,#0.00")
    Else
        AP4 = Format(tA4, "#,#0.00")
    End If
    H4 = Format(tA4, "#,#0.00")
End Sub

Private Sub ApprovePHD()
'On Error Resume Next
    Dim tA5, tP5 As Double
    tA5 = IIf(IsNumeric(A5), A5, 0)
    tP5 = IIf(IsNumeric(P5), P5, 0)
    If ChkStr(P5) = "FULL" Then
        AP5 = Format(tA5, "#,#0.00")
    ElseIf CDbl(tA5) > CDbl(tP5) Then
        AP5 = Format(tP5, "#,#0.00")
    Else
        AP5 = Format(tA5, "#,#0.00")
    End If
    H5 = Format(tA5, "#,#0.00")
End Sub


Private Sub ApprovedSurgical()
    Dim tA6, tP6 As Double
    tA6 = IIf(IsNumeric(A6), A6, 0)
    tP6 = IIf(IsNumeric(P6), P6, 0)
    If ChkStr(P6) = "FULL" Then
        AP6 = Format(tA6, "#,#0.00")
    ElseIf CDbl(tA6) > CDbl(tP6) Then
        AP6 = Format(tP6, "#,#0.00")
    Else
        AP6 = Format(tA6, "#,#0.00")
    End If
    H6 = Format(tA6, "#,#0.00")
End Sub


Private Sub ApprovedAna()
    Dim tA7, tP7 As Double
    tA7 = IIf(IsNumeric(A7), A7, 0)
    tP7 = IIf(IsNumeric(P7), P7, 0)
    If ChkStr(P7) = "FULL" Then
        AP7 = Format(tA7, "#,#0.00")
    ElseIf CDbl(tA7) > CDbl(tP7) Then
        AP7 = Format(tP7, "#,#0.00")
    Else
        AP7 = Format(tA7, "#,#0.00")
    End If
    H7 = Format(tA7, "#,#0.00")
End Sub

Private Sub ApprovePSpec()
    Dim tA8, tP8 As Double
    tA8 = IIf(IsNumeric(A8), A8, 0)
    tP8 = IIf(IsNumeric(P8), P8, 0)
    If ChkStr(P8) = "FULL" Then
        AP8 = Format(tA8, "#,#0.00")
    ElseIf CDbl(tA8) > CDbl(tP8) Then
        AP8 = Format(tP8, "#,#0.00")
    Else
        AP8 = Format(tA8, "#,#0.00")
    End If
    H8 = Format(tA8, "#,#0.00")
End Sub

Private Sub ApproveInH()
    Dim tA9, tP9 As Double
    tA9 = IIf(IsNumeric(A9), A9, 0)
    tP9 = IIf(IsNumeric(P9), P9, 0)
    If ChkStr(P9) = "FULL" Then
        AP9 = Format(tA9, "#,#0.00")
    ElseIf CDbl(tA9) > CDbl(tP9) Then
        AP9 = Format(tP9, "#,#0.00")
    Else
        AP9 = Format(tA9, "#,#0.00")
    End If
    H9 = Format(tA9, "#,#0.00")
End Sub


Private Sub ApprovePostH()
    Dim tA10, tP10 As Double
    tA10 = IIf(IsNumeric(A10), A10, 0)
    tP10 = IIf(IsNumeric(P10), P10, 0)
    If ChkStr(P10) = "FULL" Then
        AP10 = Format(tA10, "#,#0.00")
    ElseIf CDbl(tA10) > CDbl(tP10) Then
        AP10 = Format(tP10, "#,#0.00")
    Else
        AP10 = Format(tA10, "#,#0.00")
    End If
    H10 = Format(tA10, "#,#0.00")
End Sub

Private Sub ApproveEOut()
    Dim tA11, tP11 As Double
    tA11 = IIf(IsNumeric(A11), A11, 0)
    tP11 = IIf(IsNumeric(P11), P11, 0)
    If ChkStr(P11) = "FULL" Then
        AP11 = Format(tA11, "#,#0.00")
    ElseIf CDbl(tA11) > CDbl(tP11) Then
        AP11 = Format(tP11, "#,#0.00")
    Else
        AP11 = Format(tA11, "#,#0.00")
    End If
    H11 = Format(tA11, "#,#0.00")
End Sub

Private Sub ApproveAF()
    Dim tA12, tP12 As Double
    tA12 = IIf(IsNumeric(A12), A12, 0)
    tP12 = IIf(IsNumeric(P12), P12, 0)
    If ChkStr(P12) = "FULL" Then
        AP12 = Format(tA12, "#,#0.00")
    ElseIf CDbl(tA12) > CDbl(tP12) Then
        AP12 = Format(tP12, "#,#0.00")
    Else
        AP12 = Format(tA12, "#,#0.00")
    End If
    H12 = Format(tA12, "#,#0.00")
End Sub

Private Sub ApprovePT()
    Dim tA13, tP13 As Double
    tA13 = IIf(IsNumeric(A13), A13, 0)
    tP13 = IIf(IsNumeric(P13), P13, 0)
    If ChkStr(P13) = "FULL" Then
        AP13 = Format(tA13, "#,#0.00")
    ElseIf CDbl(tA13) > CDbl(tP13) Then
        AP13 = Format(tP13, "#,#0.00")
    Else
        AP13 = Format(tA13, "#,#0.00")
    End If
    H13 = Format(tA13, "#,#0.00")
End Sub


Private Sub ApprovedKidney()
'On Error Resume Next
    Dim tA14, tP14 As Double
    tA14 = IIf(IsNumeric(A14), A14, 0)
    tP14 = IIf(IsNumeric(P14), P14, 0)
    If ChkStr(P14) = "FULL" Then
        AP14 = Format(tA14, "#,#0.00")
    ElseIf CDbl(tA14) > CDbl(tP14) Then
        AP14 = Format(tP14, "#,#0.00")
    Else
        AP14 = Format(tA14, "#,#0.00")
    End If
    H14 = Format(tA14, "#,#0.00")
End Sub

Private Sub ApprovedOrgan()
'On Error Resume Next
    Dim tA15, tP15 As Double
    tA15 = IIf(IsNumeric(A15), A15, 0)
    tP15 = IIf(IsNumeric(P15), P15, 0)
    If ChkStr(P15) = "FULL" Then
        AP15 = Format(tA15, "#,#0.00")
    ElseIf CDbl(tA15) > CDbl(tP15) Then
        AP15 = Format(tP15, "#,#0.00")
    Else
        AP15 = Format(tA15, "#,#0.00")
    End If
    H15 = Format(tA15, "#,#0.00")
End Sub

Private Sub ApproveCash()
'On Error Resume Next
    Dim tA16, tP16 As Double
    tA16 = IIf(IsNumeric(A16), A16, 0)
    tP16 = IIf(IsNumeric(P16), P16, 0)
    If ChkStr(P16) = "FULL" Then
        AP16 = Format(tA16, "#,#0.00")
    ElseIf CDbl(tA16) > CDbl(tP16) Then
        AP16 = Format(tP16, "#,#0.00")
    Else
        AP16 = Format(tA16, "#,#0.00")
    End If
    H16 = Format(tA16, "#,#0.00")
End Sub

Private Sub ApprovedTax()
'On Error Resume Next
    Dim tA17, tP17 As Double
    tA17 = IIf(IsNumeric(A17), A17, 0)
    tP17 = IIf(IsNumeric(P17), P17, 0)
    If ChkStr(P17) = "FULL" Then
        AP17 = Format(tA17, "#,#0.00")
    ElseIf CDbl(tA17) > CDbl(tP17) Then
        AP17 = Format(tP17, "#,#0.00")
    Else
        AP17 = Format(tA17, "#,#0.00")
    End If
    H17 = Format(tA17, "#,#0.00")
End Sub

Private Sub ApprovedLodger()
'On Error Resume Next
    Dim tA18, tP18 As Double
    tA18 = IIf(IsNumeric(A18), A18, 0)
    tP18 = IIf(IsNumeric(P18), P18, 0)
    If ChkStr(P18) = "FULL" Then
        AP18 = Format(tA18, "#,#0.00")
    ElseIf CDbl(tA18) > CDbl(tP18) Then
        AP18 = Format(tP18, "#,#0.00")
    Else
        AP18 = Format(tA18, "#,#0.00")
    End If
    H18 = Format(tA18, "#,#0.00")
End Sub

Private Sub ApproveMR()
'On Error Resume Next
    Dim tA19, tP19 As Double
    tA19 = IIf(IsNumeric(A19), A19, 0)
    tP19 = IIf(IsNumeric(P19), P19, 0)
    If ChkStr(P19) = "FULL" Then
        AP19 = Format(tA19, "#,#0.00")
    ElseIf CDbl(tA19) > CDbl(tP19) Then
        AP19 = Format(tP19, "#,#0.00")
    Else
        AP19 = Format(tA19, "#,#0.00")
    End If
    H19 = Format(tA19, "#,#0.00")

End Sub

Private Sub ApprovePhone()
    Dim tA20, tP20 As Double
    tA20 = IIf(IsNumeric(A20), A20, 0)
    tP20 = IIf(IsNumeric(P20), P20, 0)
    If ChkStr(P20) = "FULL" Then
        AP20 = Format(tA20, "#,#0.00")
    ElseIf CDbl(tA20) > CDbl(tP20) Then
        AP20 = Format(tP20, "#,#0.00")
    Else
        AP20 = Format(tA20, "#,#0.00")
    End If
    H20 = Format(tA20, "#,#0.00")
End Sub


Private Sub ApproveRegistration()
'On Error Resume Next
    Dim tA21, tP21 As Double
    tA21 = IIf(IsNumeric(A21), A21, 0)
    tP21 = IIf(IsNumeric(P21), P21, 0)
    If ChkStr(P21) = "FULL" Then
        AP21 = Format(tA21, "#,#0.00")
    ElseIf CDbl(tA21) > CDbl(tP21) Then
        AP21 = Format(tP21, "#,#0.00")
    Else
        AP21 = Format(tA21, "#,#0.00")
    End If
    H21 = Format(tA21, "#,#0.00")
End Sub



Private Sub ApprovedOthers()
'On Error Resume Next
    Dim tA22, tP22 As Double
    tA22 = IIf(IsNumeric(A22), A22, 0)
    tP22 = IIf(IsNumeric(P22), P22, 0)
    If ChkStr(P22) = "FULL" Then
        AP22 = Format(tA22, "#,#0.00")
    ElseIf CDbl(tA22) > CDbl(tP22) Then
        AP22 = Format(tP22, "#,#0.00")
    Else
        AP22 = Format(tA22, "#,#0.00")
    End If
    H22 = Format(tA22, "#,#0.00")
End Sub


Private Sub ApprovedOthers2()
'On Error Resume Next
    Dim tA23, tP23 As Double
    tA23 = IIf(IsNumeric(A23), A23, 0)
    tP23 = IIf(IsNumeric(P23), P23, 0)
    If ChkStr(P23) = "FULL" Then
        AP23 = Format(tA23, "#,#0.00")
    ElseIf CDbl(tA23) > CDbl(tP23) Then
        AP23 = Format(tP23, "#,#0.00")
    Else
        AP23 = Format(tA23, "#,#0.00")
    End If
    H23 = Format(tA23, "#,#0.00")
End Sub

' ----------------- Start Lost focus to get the due to/from Member ------------
Private Sub AP1_lostFocus()
    M1 = Format(CDbl(AP1) - CDbl(H1), "#,#0.00;(#,#0.00)")
End Sub

Private Sub AP2_lostFocus()
    M2 = Format(CDbl(AP2) - CDbl(H2), "#,#0.00;(#,#0.00)")
End Sub

Private Sub AP3_lostFocus()
    M3 = Format(CDbl(AP3) - CDbl(H3), "#,#0.00;(#,#0.00)")
End Sub

Private Sub AP4_lostFocus()
    M4 = Format(CDbl(AP4) - CDbl(H4), "#,#0.00;(#,#0.00)")
End Sub

Private Sub AP5_lostFocus()
    M5 = Format(CDbl(AP5) - CDbl(H5), "#,#0.00;(#,#0.00)")
End Sub

Private Sub AP6_lostFocus()
    M6 = Format(CDbl(AP6) - CDbl(H6), "#,#0.00;(#,#0.00)")
End Sub

Private Sub AP7_lostFocus()
    M7 = Format(CDbl(AP7) - CDbl(H7), "#,#0.00;(#,#0.00)")
End Sub

Private Sub AP8_lostFocus()
    M8 = Format(CDbl(AP8) - CDbl(H8), "#,#0.00;(#,#0.00)")
End Sub

Private Sub AP9_lostFocus()
    M9 = Format(CDbl(AP9) - CDbl(H9), "#,#0.00;(#,#0.00)")
End Sub

Private Sub AP10_lostFocus()
    M10 = Format(CDbl(AP10) - CDbl(H10), "#,#0.00;(#,#0.00)")
End Sub

Private Sub AP11_lostFocus()
    M11 = Format(CDbl(AP11) - CDbl(H11), "#,#0.00;(#,#0.00)")
End Sub

Private Sub AP12_lostFocus()
    M12 = Format(CDbl(AP12) - CDbl(H12), "#,#0.00;(#,#0.00)")
End Sub

Private Sub AP13_lostFocus()
    M13 = Format(CDbl(AP13) - CDbl(H13), "#,#0.00;(#,#0.00)")
End Sub

Private Sub AP14_lostFocus()
    M14 = Format(CDbl(AP14) - CDbl(H14), "#,#0.00;(#,#0.00)")
End Sub

Private Sub AP15_lostFocus()
    M15 = Format(CDbl(AP15) - CDbl(H15), "#,#0.00;(#,#0.00)")
End Sub

Private Sub AP16_lostFocus()
    M16 = Format(CDbl(AP16) - CDbl(H16), "#,#0.00;(#,#0.00)")
End Sub

Private Sub AP17_lostFocus()
    M17 = Format(CDbl(AP17) - CDbl(H17), "#,#0.00;(#,#0.00)")
End Sub

Private Sub AP18_lostFocus()
    M18 = Format(CDbl(AP18) - CDbl(H18), "#,#0.00;(#,#0.00)")
End Sub

Private Sub AP19_lostFocus()
    M19 = Format(CDbl(AP19) - CDbl(H19), "#,#0.00;(#,#0.00)")
End Sub

Private Sub AP20_lostFocus()
    M20 = Format(CDbl(AP20) - CDbl(H20), "#,#0.00;(#,#0.00)")
End Sub

Private Sub AP21_lostFocus()
    M21 = Format(CDbl(AP21) - CDbl(H21), "#,#0.00;(#,#0.00)")
End Sub


'Section TOTAL
' '''''''''''''''''''''''''''''''''''

Private Sub calcTotal()
   
    M1 = Format(CDbl(AP1) - CDbl(H1), "#,#0.00;(#,#0.00)")
    M2 = Format(CDbl(AP2) - CDbl(H2), "#,#0.00;(#,#0.00)")
    M3 = Format(CDbl(AP3) - CDbl(H3), "#,#0.00;(#,#0.00)")
    M4 = Format(CDbl(AP4) - CDbl(H4), "#,#0.00;(#,#0.00)")
    M5 = Format(CDbl(AP5) - CDbl(H5), "#,#0.00;(#,#0.00)")
    M6 = Format(CDbl(AP6) - CDbl(H6), "#,#0.00;(#,#0.00)")
    M7 = Format(CDbl(AP7) - CDbl(H7), "#,#0.00;(#,#0.00)")
    M8 = Format(CDbl(AP8) - CDbl(H8), "#,#0.00;(#,#0.00)")
    M9 = Format(CDbl(AP9) - CDbl(H9), "#,#0.00;(#,#0.00)")
    M10 = Format(CDbl(AP10) - CDbl(H10), "#,#0.00;(#,#0.00)")
    M11 = Format(CDbl(AP11) - CDbl(H11), "#,#0.00;(#,#0.00)")
    M12 = Format(CDbl(AP12) - CDbl(H12), "#,#0.00;(#,#0.00)")
    M13 = Format(CDbl(AP13) - CDbl(H13), "#,#0.00;(#,#0.00)")
    M14 = Format(CDbl(AP14) - CDbl(H14), "#,#0.00;(#,#0.00)")
    M15 = Format(CDbl(AP15) - CDbl(H15), "#,#0.00;(#,#0.00)")
    M16 = Format(CDbl(AP16) - CDbl(H16), "#,#0.00;(#,#0.00)")
    M17 = Format(CDbl(AP17) - CDbl(H17), "#,#0.00;(#,#0.00)")
    M18 = Format(CDbl(AP18) - CDbl(H18), "#,#0.00;(#,#0.00)")
    M19 = Format(CDbl(AP19) - CDbl(H19), "#,#0.00;(#,#0.00)")
    M20 = Format(CDbl(AP20) - CDbl(H20), "#,#0.00;(#,#0.00)")
    M21 = Format(CDbl(AP21) - CDbl(H21), "#,#0.00;(#,#0.00)")
    
    Call SumActual
    Call CalculateTotalGross
End Sub

Private Sub TotalPaidByMemberHos_LostFocus()
    TotalPaidByMemberMember = Format(CDbl(TotalPaidByMemberHos), "#,#0.00;(#,#0.00)")
    TotalPaidMedixHos = IIf(CDbl(TotalHospital) - CDbl(TotalPaidByMemberHos) > CDbl(TotalPaidMedixHos) + CDbl(txtLGAmount), Format(txtLGAmount, "#,#0.00;(#,#0.00)"), Format(CDbl(TotalHospital) - CDbl(TotalPaidByMemberHos), "#,#0.00;(#,#0.00)"))
    'TotalGrossHos = Format(CDbl(TotalPaidMedixHos) + CDbl(TotalDiscountHos), "#,#0.00;(#,#0.00)")
    TotalGrossHos = Format(CDbl(TotalPaidMedixHos), "#,#0.00;(#,#0.00)")
    'TotalGrossMember = Format(CDbl(TotalMember) + CDbl(TotalPaidByMemberMember) + CDbl(TotalPaidMedixMember) + CDbl(TotalDiscountMember), "#,#0.00;(#,#0.00)")
    TotalGrossMember = Format(CDbl(TotalMember) + CDbl(TotalPaidByMemberMember) + CDbl(TotalPaidMedixMember), "#,#0.00;(#,#0.00)")
End Sub


Private Sub CalculateTotalGross()
'On Error Resume Next
    'TotalPaidByMemberHos = Format(CDbl(TotalPaidByMember), "#,#0.00;(#,#0.00)")
    'TotalPaidMedix = Format(IIf((CDbl(TotalActual) - CDbl(TotalPaidByMember)) > P21, P21, CDbl(TotalActual) - CDbl(TotalPaidByMember)), "#,#0.00;(#,#0.00)")
    'TotalPaidMedixMember = Format(CDbl(TotalPaidMedix), "#,#0.00;(#,#0.00)")
    TotalGrossApproved = Format(IIf(CDbl(TotalApproved) < CDbl(PAnnualLimit), TotalApproved, PAnnualLimit), "#,#0.00;(#,#0.00)")
    
'   TotalPaidMedixHos = IIf(CDbl(TotalHospital) - CDbl(TotalPaidByMemberHos) > CDbl(TotalPaidMedixHos), Format(txtLGAmount, "#,#0.00;(#,#0.00)"), Format(CDbl(TotalHospital) - CDbl(TotalPaidByMemberHos), "#,#0.00;(#,#0.00)"))
    TotalPaidByMemberMember = Format(CDbl(TotalPaidByMemberHos), "#,#0.00;(#,#0.00)")
'    TotalDiscountMember = IIf(CDbl(TotalGrossApproved) < CDbl(TotalApproved), Format(CDbl(TotalGrossApproved) - CDbl(TotalApproved), "#,#0.00;(#,#0.00)"), 0)
    
    'TotalGross = Format(CDbl(TotalActual) + CDbl(TotalDiscount), "#,#0.00;(#,#0.00)")
    TotalGross = Format(CDbl(TotalActual), "#,#0.00;(#,#0.00)")
   ' TotalGrossHos = Format(CDbl(TotalPaidMedixHos) + CDbl(TotalDiscountHos), "#,#0.00;(#,#0.00)")
'    TotalGrossHos = Format(CDbl(TotalPaidMedixHos), "#,#0.00;(#,#0.00)")
'    TotalGrossMember = Format(CDbl(TotalMember) + CDbl(TotalPaidByMemberMember) + CDbl(TotalPaidMedixMember) + CDbl(TotalDiscountMember), "#,#0.00;(#,#0.00)")
    TotalGrossMember = Format(CDbl(TotalMember) + CDbl(TotalPaidByMemberMember) + CDbl(TotalPaidMedixMember), "#,#0.00;(#,#0.00)")
End Sub

Private Sub SumActual()
'On Error Resume Next
    TotalActual = Format(CDbl(IIf(Trim(A1) = "", 0, A1)) + CDbl(IIf(Trim(A2) = "", 0, A2)) _
                + CDbl(IIf(Trim(A3) = "", 0, A3)) + CDbl(IIf(Trim(A4) = "", 0, A4)) _
                + CDbl(IIf(Trim(A5) = "", 0, A5)) + CDbl(IIf(Trim(A6) = "", 0, A6)) _
                + CDbl(IIf(Trim(A7) = "", 0, A7)) + CDbl(IIf(Trim(A8) = "", 0, A8)) _
                + CDbl(IIf(Trim(A9) = "", 0, A9)) + CDbl(IIf(Trim(A10) = "", 0, A10)) _
                + CDbl(IIf(Trim(A11) = "", 0, A11)) + CDbl(IIf(Trim(A12) = "", 0, A12)) + CDbl(IIf(Trim(A13) = "", 0, A13)) _
                + CDbl(IIf(Trim(A14) = "", 0, A14)) + CDbl(IIf(Trim(A15) = "", 0, A15)) + CDbl(IIf(Trim(A16) = "", 0, A16)) _
                + CDbl(IIf(Trim(A17) = "", 0, A17)) + CDbl(IIf(Trim(A18) = "", 0, A18)) + CDbl(IIf(Trim(A19) = "", 0, A19)) _
                + CDbl(IIf(Trim(A20) = "", 0, A20)) + CDbl(IIf(Trim(A21) = "", 0, A21)) + CDbl(IIf(Trim(A22) = "", 0, A22)) + CDbl(IIf(Trim(A23) = "", 0, A23)), "#,#0.00;(#,#0.00)")
    TotalAPpproved = Format(CDbl(IIf(Trim(AP1) = "", 0, AP1)) + CDbl(IIf(Trim(AP2) = "", 0, AP2)) _
                + CDbl(IIf(Trim(AP3) = "", 0, AP3)) + CDbl(IIf(Trim(AP4) = "", 0, AP4)) _
                + CDbl(IIf(Trim(AP5) = "", 0, AP5)) + CDbl(IIf(Trim(AP6) = "", 0, AP6)) _
                + CDbl(IIf(Trim(AP7) = "", 0, AP7)) + CDbl(IIf(Trim(AP8) = "", 0, AP8)) _
                + CDbl(IIf(Trim(AP9) = "", 0, AP9)) + CDbl(IIf(Trim(AP10) = "", 0, AP10)) _
                + CDbl(IIf(Trim(AP11) = "", 0, AP11)) + CDbl(IIf(Trim(AP12) = "", 0, AP12)) + CDbl(IIf(Trim(AP13) = "", 0, AP13)) _
                + CDbl(IIf(Trim(AP14) = "", 0, AP14)) + CDbl(IIf(Trim(AP15) = "", 0, AP15)) + CDbl(IIf(Trim(AP16) = "", 0, AP16)) _
                + CDbl(IIf(Trim(AP17) = "", 0, AP17)) + CDbl(IIf(Trim(AP18) = "", 0, AP18)) + CDbl(IIf(Trim(AP19) = "", 0, AP19)) _
                + CDbl(IIf(Trim(AP20) = "", 0, AP20)) + CDbl(IIf(Trim(AP21) = "", 0, AP21)) + CDbl(IIf(Trim(AP22) = "", 0, AP22)) + CDbl(IIf(Trim(AP23) = "", 0, AP23)), "#,#0.00;(#,#0.00)")
    TotalHospital = Format(CDbl(IIf(Trim(A1) = "", 0, A1)) + CDbl(IIf(Trim(A2) = "", 0, A2)) _
                + CDbl(IIf(Trim(A3) = "", 0, A3)) + CDbl(IIf(Trim(A4) = "", 0, A4)) _
                + CDbl(IIf(Trim(A5) = "", 0, A5)) + CDbl(IIf(Trim(A6) = "", 0, A6)) _
                + CDbl(IIf(Trim(A7) = "", 0, A7)) + CDbl(IIf(Trim(A8) = "", 0, A8)) _
                + CDbl(IIf(Trim(A9) = "", 0, A9)) + CDbl(IIf(Trim(A10) = "", 0, A10)) _
                + CDbl(IIf(Trim(A11) = "", 0, A11)) + CDbl(IIf(Trim(A12) = "", 0, A12)) + CDbl(IIf(Trim(A13) = "", 0, A13)) _
                + CDbl(IIf(Trim(A14) = "", 0, A14)) + CDbl(IIf(Trim(A15) = "", 0, A15)) + CDbl(IIf(Trim(A16) = "", 0, A16)) _
                + CDbl(IIf(Trim(A17) = "", 0, A17)) + CDbl(IIf(Trim(A18) = "", 0, A18)) + CDbl(IIf(Trim(A19) = "", 0, A19)) _
                + CDbl(IIf(Trim(A20) = "", 0, A20)) + CDbl(IIf(Trim(A21) = "", 0, A21)) + CDbl(IIf(Trim(A22) = "", 0, A22)) + CDbl(IIf(Trim(A23) = "", 0, A23)), "#,#0.00;(#,#0.00)")
    TotalMember = Format(CDbl(IIf(Trim(M1) = "", 0, M1)) + CDbl(IIf(Trim(M2) = "", 0, M2)) _
                + CDbl(IIf(Trim(M3) = "", 0, M3)) + CDbl(IIf(Trim(M4) = "", 0, M4)) _
                + CDbl(IIf(Trim(M5) = "", 0, M5)) + CDbl(IIf(Trim(M6) = "", 0, M6)) _
                + CDbl(IIf(Trim(M7) = "", 0, M7)) + CDbl(IIf(Trim(M8) = "", 0, M8)) _
                + CDbl(IIf(Trim(M9) = "", 0, M9)) + CDbl(IIf(Trim(M10) = "", 0, M10)) _
                + CDbl(IIf(Trim(M11) = "", 0, M11)) + CDbl(IIf(Trim(M12) = "", 0, M12)) + CDbl(IIf(Trim(M13) = "", 0, M13)) _
                + CDbl(IIf(Trim(M14) = "", 0, M14)) + CDbl(IIf(Trim(M15) = "", 0, M15)) + CDbl(IIf(Trim(M16) = "", 0, M16)) _
                + CDbl(IIf(Trim(M17) = "", 0, M17)) + CDbl(IIf(Trim(M18) = "", 0, M18)) + CDbl(IIf(Trim(M19) = "", 0, M19)) _
                + CDbl(IIf(Trim(M20) = "", 0, M20)) + CDbl(IIf(Trim(M21) = "", 0, M21)) + CDbl(IIf(Trim(M22) = "", 0, M22)) + CDbl(IIf(Trim(M23) = "", 0, M23)), "#,#0.00;(#,#0.00)")
End Sub

' #################### Search Funtion
Private Sub cmdSearchCase_Click()
    frmCASList.Source = 1
    frmCASList.show vbModal
End Sub


'  ############## Display data (TAB 2)
Public Sub Display_Diagnosis(DIAStatus As String, CASNumber As String)
    Dim DiagnosisSql, DiagnosisFirst, DiagnosisFinal As String
    Dim lstDiag As ListItem
    Dim rst3 As New ADODB.Recordset
    Dim rstFinal As New ADODB.Recordset
    
    DiagnosisSql = "select DIAIndex, CASNumber , ICDCode1 , ICDCOde2 , ICDCOde3, " & _
                        " DIADOS1, DIADOS2, DIADOS3 "
'    If ChkStr(DIAStatus) = "I" Then
    DiagnosisFirst = DiagnosisSql & " , DIAPrelimDiag1  as diag1 , DIAPrelimDiag2 as diag2 , DIAPrelimDiag3 as diag3 "
    DiagnosisFirst = DiagnosisFirst & " from DIAFirst  where CASNUMBer ='" & Trim(CASNumber) & "'"
 '   Else
    DiagnosisFinal = DiagnosisSql & " , DIAFinalDiag1 as diag1 , DIAFinalDiag2  as diag2  ,DIAFinalDiag3 as diag3  "
    DiagnosisFinal = DiagnosisFinal & " from DIAFinal where CASNUMBer ='" & Trim(CASNumber) & "'"
  '  End If
    
    rst3.Open DiagnosisFirst, medixmasterconn, adOpenKeyset, adLockOptimistic
    rstFinal.Open DiagnosisFinal, medixmasterconn, adOpenKeyset, adLockOptimistic
    ' For Final
    lstFinalDiag.listitems.Clear

     If rst3.recordCount Then
        Set lstDiag = lstFinalDiag.listitems.Add(, , rst3!DiaIndex)
        With lstDiag
            .SubItems(1) = Trim(rst3!diag1)
            .SubItems(2) = rst3!DIADOS1
            .SubItems(3) = rst3!ICDCode1
        End With
        Set lstDiag = lstFinalDiag.listitems.Add(, , rst3!DiaIndex)
        With lstDiag
            .SubItems(1) = Trim(rst3!diag2)
            .SubItems(2) = rst3!DIADOS2
            .SubItems(3) = rst3!ICDCode2
        End With
        Set lstDiag = lstFinalDiag.listitems.Add(, , rst3!DiaIndex)
        With lstDiag
            .SubItems(1) = Trim(rst3!diag3)
            .SubItems(2) = rst3!DIADOS3
            .SubItems(3) = rst3!ICDCode3
        End With
    End If
    
    
     If rstFinal.recordCount Then
        Set lstDiag = lstFinalDiag.listitems.Add(, , rstFinal!DiaIndex)
        With lstDiag
            .SubItems(1) = Trim(rstFinal!diag1)
            .SubItems(2) = rstFinal!DIADOS1
            .SubItems(3) = rstFinal!ICDCode1
        End With
        Set lstDiag = lstFinalDiag.listitems.Add(, , rstFinal!DiaIndex)
        With lstDiag
            .SubItems(1) = Trim(rstFinal!diag2)
            .SubItems(2) = rstFinal!DIADOS2
            .SubItems(3) = rstFinal!ICDCode2
        End With
        Set lstDiag = lstFinalDiag.listitems.Add(, , rstFinal!DiaIndex)
        With lstDiag
            .SubItems(1) = Trim(rstFinal!diag3)
            .SubItems(2) = rstFinal!DIADOS3
            .SubItems(3) = rstFinal!ICDCode3
        End With
    End If
    
    rst3.Close
    Set rst3 = Nothing
End Sub

Public Function FindHosName(HosCode As String) As String
    Dim HOSSql As String
    Dim rst3 As New ADODB.Recordset
    HOSSql = " select HOSCode, HOSNAme  from HOS where HOSCOde='" & Trim(HosCode) & "'"
    
    rst3.Open HOSSql, medixmasterconn, adOpenKeyset, adLockOptimistic
    If rst3.recordCount Then
        FindHosName = rst3!HOSName
    Else
        MsgBox "Hospital Name not found!", vbCritical + vbOKOnly
    End If
    rst3.Close
    Set rst3 = Nothing
End Function


'---------------------------------------------------------------------------------------------------------------------------------------------
' Benefits

Public Sub display_benefit()
    Dim benefitstr  As String
    Dim BNF As New ADODB.Recordset
     
    benefitstr = "Select BNFCode , BNFNoOfDays , BNFClaimableAmount , BNFDescription " & _
                " from View_BNF Where INSCode ='" & Trim(txtINSCode) & "' And PLNCode ='" & Trim(txtPLNCode) & "' And HLTCode ='" & Trim(txtHLTCode) & "'"
    BNF.Open benefitstr, medixmasterconn, adOpenKeyset, adLockOptimistic
    If BNF.recordCount Then
        Do Until BNF.EOF
            With BNF
                If InStr(1, !BNFCode, "-01") Then
                    P1A = !BNFClaimAbleAmount
                    lblMaxRB = !BNFNoOFDays
                ElseIf InStr(1, !BNFCode, "-02") Then
                    P2A = !BNFClaimAbleAmount
                    lblMaxICU = !BNFNoOFDays
                ElseIf InStr(1, !BNFCode, "-03") Then
                    P3 = !BNFClaimAbleAmount
                ElseIf InStr(1, !BNFCode, "-04") Then
                    P4 = !BNFClaimAbleAmount
                ElseIf InStr(1, !BNFCode, "-05") Then
                    P5 = !BNFClaimAbleAmount
                ElseIf InStr(1, !BNFCode, "-06") Then
                    P6 = !BNFClaimAbleAmount
                ElseIf InStr(1, !BNFCode, "-07") Then
                    P7 = !BNFClaimAbleAmount
                ElseIf InStr(1, !BNFCode, "-08") Then
                    P8 = !BNFClaimAbleAmount
                ElseIf InStr(1, !BNFCode, "-09") Then
                    P9 = !BNFClaimAbleAmount
                ElseIf InStr(1, !BNFCode, "-10") Then
                    P10 = !BNFClaimAbleAmount
                ElseIf InStr(1, !BNFCode, "-11") Then
                    P11 = !BNFClaimAbleAmount
                ElseIf InStr(1, !BNFCode, "-12") Then
                    P12 = !BNFClaimAbleAmount
                ElseIf InStr(1, !BNFCode, "-13") Then
                    P13 = !BNFClaimAbleAmount
                ElseIf InStr(1, !BNFCode, "-14") Then
                    P14 = !BNFClaimAbleAmount
                ElseIf InStr(1, !BNFCode, "-15") Then
                    P15 = !BNFClaimAbleAmount
                ElseIf InStr(1, !BNFCode, "-16") Then
                    txtRMCashA = !BNFClaimAbleAmount
                    lblMaxCash = !BNFNoOFDays
                ElseIf InStr(1, !BNFCode, "-17") Then
                    P17 = !BNFClaimAbleAmount
                ElseIf InStr(1, !BNFCode, "-18") Then
                    LodgerAmount = !BNFClaimAbleAmount
                    lblMaxLodger = !BNFNoOFDays
                ElseIf InStr(1, !BNFCode, "-19") Then
                    P19 = !BNFClaimAbleAmount
                ElseIf InStr(1, !BNFCode, "-20") Then
                    P20 = !BNFClaimAbleAmount
                ElseIf InStr(1, !BNFCode, "-21") Then
                    P21 = !BNFClaimAbleAmount
                ElseIf InStr(1, !BNFCode, "-22") Then
                    P22 = !BNFClaimAbleAmount
                    txtOthers1 = !BNFdescription
                ElseIf InStr(1, !BNFCode, "-23") Then
                    P23 = !BNFClaimAbleAmount
                    txtOthers1 = !BNFdescription
                End If
            End With
            BNF.MoveNext
        Loop
        
        If IsNumeric(P1A) Then
            RoomAmount = CInt(P1A)
        End If
        If IsNumeric(P2A) Then
            ICUAmount = CInt(P2A)
        End If
        If IsNumeric(P17A) = True Then
            LodgerAmount = CInt(P17A)
        End If
        If IsNumeric(txtRMCashA) Then
            txtRMCash = CInt(txtRMCashA)
        End If
        
    Else
        MsgBox "Error! Please check on your benefit table!", vbCritical + vbOKOnly
    End If
    BNF.Close
    Set BNF = Nothing
End Sub




' ########### SAVE TO DB #############################

Private Sub cmdSave_Click()
    On Error GoTo errSave
    Dim RecordSaved
    Dim startTrans As Boolean
    Dim strsql As String
    
    Dim CheckClaim As New ADODB.Recordset
    strsql = "select CASNumber from Claim where CAsNUmber ='" & Trim(txtCASNumber) & "'"
    CheckClaim.Open strsql, medixmasterconn, adOpenKeyset, adLockOptimistic
    
    If CheckClaim.recordCount Then
      MsgBox "Worksheet for this case has been opened before!" & vbNewLine & _
             " Please use Claim adjustment for further editting!", vbCritical + vbOKOnly
    Else
        If Len(Trim(txtCASNumber)) = 0 Then
            MsgBox "Please enter Case Number!", vbOKOnly + vbCritical, DialogBoxTitle
            txtCASNumber.SetFocus
        ElseIf Len(Trim(txtMBMNumber)) = 0 Then
            MsgBox "Please Member Number !", vbOKOnly + vbCritical, DialogBoxTitle
            txtMBMNumber.SetFocus
        ElseIf lstInvoice.listitems.count = 0 Then
            MsgBox "Worksheet cannot be saved without any invoice!", vbOKOnly + vbCritical, DialogBoxTitle
            SSTab1.Tab = 1
        Else
            RecordSaved = MsgBox("Are you sure?", vbYesNo + vbExclamation)
            If RecordSaved = vbYes Then
               Screen.MousePointer = vbHourglass
               startTrans = True
               medixmasterconnWS.beginTrans
               
               Call saveClaim
               Call saveLimit
               Call saveActual
               Call saveApprove
               Call saveDueMember
               Call saveHospital
               Call saveInvoice
                              
               medixmasterconnWS.CommitTrans
               Screen.MousePointer = Default
               MsgBox "Record Save", vbOKOnly + vbInformation, DialogBoxTitle
               cmdSave.Enabled = False
             End If
        End If
    End If
    CheckClaim.Close
    Set CheckClaim = Nothing
    Exit Sub
errSave:
    MsgBox ERR.Description
    If startTrans = True Then
        medixmasterconn.RollbackTrans
    End If
End Sub

Private Sub saveClaim()
    Dim Claim As New ADODB.Recordset
    Claim.Open "Claim", medixmasterconnWS, adOpenKeyset, adLockOptimistic
    
    Claim.AddNew
    With Claim
         !CASNumber = txtCASNumber
         !claimversion = VersionNo
         !MBMNumber = txtMBMNumber
         !MBMCoveredID = txtCoveredID
         !ClaimStatus = "O"
         !ClaimOpenDate = Date
         !ClaimOpenBy = ChkStr(Logon)
         !ClaimOthers = txtOthers
         !ClaimBoardDay = IIf(IsNumeric(DB1), DB1, 0)
         !ClaimICUDay = IIf(IsNumeric(ICU1), ICU1, 0)
         !ClaimLodgerDay = IIf(IsNumeric(LG1), LG1, 0)
         !ClaimCashAllowanceDay = IIf(IsNumeric(txtDayCash), txtDayCash, 0)
         !ClaimRoomAmount = IIf(IsNumeric(RoomAmount), RoomAmount, 0)
         !ClaimICUAmount = IIf(IsNumeric(ICUAmount), ICUAmount, 0)
         !ClaimLodgerAmount = IIf(IsNumeric(LodgerAmount), LodgerAmount, 0)
         !ClaimCashAllowanceAmount = IIf(IsNumeric(txtRMCash), txtRMCash, 0)
         !ClaimLastUpdateUser = ChkStr(Logon)
         !ClaimLastUpdateDate = Date
    End With
    Claim.Update
    
    Claim.Close
    Set Claim = Nothing
End Sub

Private Sub saveActual()
    Dim ClaimActual  As New ADODB.Recordset
    ClaimActual.Open "ClaimActual", medixmasterconnWS, adOpenKeyset, adLockOptimistic
    
    ClaimActual.AddNew
    With ClaimActual
         !CASNumber = txtCASNumber
         !CLMVersion = VersionNo
         !CLMBoard = A1
         !CLMIntensive = A2
         !CLMServices = A3
         !CLMTheatre = A4
         !CLMPre = A5
         !CLMSurgical = A6
         !CLMAna = A7
         !CLMHospitalisation = A8
         !CLMPhysician = A9
         !CLMPost = A10
         !CLMEmergency = A11
         !CLMAmbulance = A12
         !CLMOutpatient = A13
         !CLMMonthly = A14
         !CLMMedical = A15
         !CLMPhone = A16
         !CLMLodger = A17
         !CLMAdmin = A18
         !CLMTax = A19
         !CLMOthers = A20
         !CLMCash = A21
         '!CLMDiscount = TotalDiscount
         !CLMGross = TotalGross
         !CLMTotal = TotalActual
    End With
    ClaimActual.Update
    
    ClaimActual.Close
    Set ClaimActual = Nothing
End Sub

Private Sub saveLimit()
    Dim ClaimLimit  As New ADODB.Recordset
    ClaimLimit.Open "ClaimLimit", medixmasterconnWS, adOpenKeyset, adLockOptimistic
    
    ClaimLimit.AddNew
    With ClaimLimit
         !CASNumber = txtCASNumber
         !CLMLVersion = VersionNo
         !CLMLBoard = A1
         !CLMLIntensive = A2
         !CLMLServices = A3
         !CLMLTheatre = A4
         !CLMLPre = A5
         !CLMLSurgical = A6
         !CLMLAna = A7
         !CLMLHospitalisation = A8
         !CLMLPhysician = A9
         !CLMLPost = A10
         !CLMLEmergency = A11
         !CLMLAmbulance = A12
         !CLMLOutpatient = A13
         !CLMLMonthly = A14
         !CLMLMedical = A15
         !CLMLPhone = A16
         !CLMLLodger = A17
         !CLMLAdmin = A18
         !CLMLTax = A19
         !CLMLOthers = A20
    End With
    ClaimLimit.Update
    
    ClaimLimit.Close
    Set ClaimLimit = Nothing
End Sub


Private Sub saveApprove()
    Dim ClaimApprove As New ADODB.Recordset
    ClaimApprove.Open "ClaimApprove", medixmasterconnWS, adOpenKeyset, adLockOptimistic

    ClaimApprove.AddNew
    With ClaimApprove
         !CASNumber = txtCASNumber
         !CLAVersion = VersionNo
         !CLABoard = AP1
         !CLAIntensive = AP2
         !CLAServices = AP3
         !CLATheatre = AP4
         !CLAPre = AP5
         !CLASurgical = AP6
         !CLAAna = AP7
         !CLAHospitalisation = AP8
         !CLAPhysician = AP9
         !CLAPost = AP10
         !CLAEmergency = AP11
         !CLAAmbulance = AP12
         !CLAOutpatient = AP13
         !CLAMonthly = AP14
         !CLAMedical = AP15
         !CLAPhone = AP16
         !CLALodger = AP17
         !CLAAdmin = AP18
         !CLATax = AP19
         !CLAOthers = AP20
         !CLACash = AP21
         !CLATotal = TotalApproved
         !CLAGross = TotalGrossApproved
    End With
    ClaimApprove.Update
    
    ClaimApprove.Close
    Set ClaimApprove = Nothing

End Sub


Private Sub saveDueMember()
    Dim ClaimDueMember As New ADODB.Recordset

    ClaimDueMember.Open "ClaimDueMember", medixmasterconnWS, adOpenKeyset, adLockOptimistic
    ClaimDueMember.AddNew
    With ClaimDueMember
         !CASNumber = txtCASNumber
         !CLDVersion = VersionNo
         !CLDBoard = M1
         !CLDIntensive = M2
         !CLDServices = M3
         !CLDTheatre = M4
         !CLDPre = M5
         !CLDSurgical = M6
         !CLDAna = M7
         !CLDHospitalisation = M8
         !CLDPhysician = M9
         !CLDPost = M10
         !CLDEmergency = M11
         !CLDAmbulance = M12
         !CLDOutpatient = M13
         !CLDMonthly = M14
         !CLDMedical = M15
         !CLDPhone = M16
         !CLDLodger = M17
         !CLDAdmin = M18
         !CLDTax = M19
         !CLDOthers = M20
         !CLDCash = M21
         !CLDPaidByMember = TotalPaidByMemberMember
         !CLDPayableByMedix = TotalPaidMedixMember
         ''!CLDDiscount = TotalDiscountMember
         !CLDGross = TotalGrossMember
         !CLDTotal = TotalMember
    End With
    ClaimDueMember.Update

    ClaimDueMember.Close
    Set ClaimDueMember = Nothing
End Sub

Private Sub saveHospital()
    Dim ClaimHospital As New ADODB.Recordset
    ClaimHospital.Open "ClaimHospital", medixmasterconnWS, adOpenKeyset, adLockOptimistic
    
    ClaimHospital.AddNew
    With ClaimHospital
         !CASNumber = txtCASNumber
         !CLHVersion = VersionNo
         !CLHBoard = H1
         !CLHIntensive = H2
         !CLHServices = H3
         !CLHTheatre = H4
         !CLHPre = H5
         !CLHSurgical = H6
         !CLHAna = H7
         !CLHHospitalisation = H8
         !CLHPhysician = H9
         !CLHPost = H10
         !CLHEmergency = H11
         !CLHAmbulance = H12
         !CLHOutpatient = H13
         !CLHMonthly = H14
         !CLHMedical = H15
         !CLHPhone = H16
         !CLHLodger = H17
         !CLHAdmin = H18
         !CLHTax = H19
         !CLHOthers = H20
         !CLHCash = H21
         !CLHPaidByMember = TotalPaidByMemberHos
         !CLHPayableByMedix = TotalPaidMedixHos
         '!CLHDiscount = TotalDiscountHos
         !CLHGross = TotalGrossHos
         !CLHTotal = TotalHospital
    End With
    ClaimHospital.Update

    ClaimHospital.Close
    Set ClaimHospital = Nothing
End Sub

Private Sub saveInvoice()
    Dim ClaimInvoice As New ADODB.Recordset
    Dim ForLoop, ListCount As Integer
    
    ListCount = lstInvoice.listitems.count
    
    ClaimInvoice.Open "ClaimInvoice", medixmasterconnWS, adOpenKeyset, adLockOptimistic

    With ClaimInvoice
            For ForLoop = 1 To ListCount
                .AddNew
                !HosCode = lstInvoice.listitems.Item(ForLoop).Text
                !INVNO = lstInvoice.listitems.Item(ForLoop).SubItems(1)
                !InvDate = lstInvoice.listitems.Item(ForLoop).SubItems(2)
                !INVAmount = IIf(IsNumeric(lstInvoice.listitems.Item(ForLoop).SubItems(3)), lstInvoice.listitems.Item(ForLoop).SubItems(3), 0)
                !PaidToHos = IIf(IsNumeric(lstInvoice.listitems.Item(ForLoop).SubItems(4)), lstInvoice.listitems.Item(ForLoop).SubItems(4), 0)
                !PaidtoMBM = IIf(IsNumeric(lstInvoice.listitems.Item(ForLoop).SubItems(5)), lstInvoice.listitems.Item(ForLoop).SubItems(5), 0)
                !PaidByMBM = IIf(IsNumeric(lstInvoice.listitems.Item(ForLoop).SubItems(6)), lstInvoice.listitems.Item(ForLoop).SubItems(6), 0)
                !INVRemark = lstInvoice.listitems.Item(ForLoop).SubItems(7)
            Next
    End With
    ClaimInvoice.UpdateBatch
    
    ClaimInvoice.Close
    Set ClaimInvoice = Nothing
End Sub


'Public Sub Display_Hospital(CASNumber As String)
 '   Dim HOSSql As String
  '  Dim lstHos As ListItem
    'Dim rst3 As New ADODB.Recordset
    'hOSSql = " select HOS_CASEID , HOSCode, CaseCode, " & _
             " DateAdmit, DateDischarge ,HOSName"
'    HOSSql = HOSSql & " from VIEW_CAS_HOS_HOS where CASeCode ='" & Trim(CASNumber) & "'"
    
 '   rst3.Open HOSSql, medixmasterconn, adOpenKeyset, adLockOptimistic
    
'    lstHosList.listitems.Clear
 '   If rst3.recordCount Then
  '      Do Until rst3.EOF
   '         Set lstHos = lstHosList.listitems.Add(, , Trim(rst3!HosCode))
    '            With lstHos
     '               .SubItems(1) = Trim(rst3!HOSName)
      '              .SubItems(2) = CDate(rst3!DateAdmit)
       '             .SubItems(3) = CDate(rst3!DateDischarge)
        '            .SubItems(4) = rst3!HOS_CASEID
         '       End With
           ' cboHosCode.AddItem "Trim(rst3!HOSCode)" & "-" & Trim(rst3!HOSName)
          '  rst3.MoveNext
'        Loop
 '   End If
  '  rst3.Close
   ' Set rst3 = Nothing
'End Sub


