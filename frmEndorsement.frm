VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmEndorsement 
   Caption         =   "Upload Bordx"
   ClientHeight    =   7425
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   11235
   ClipControls    =   0   'False
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   7425
   ScaleWidth      =   11235
   WindowState     =   2  'Maximized
   Begin VB.Frame Frame4 
      Height          =   615
      Left            =   0
      TabIndex        =   20
      Top             =   6720
      Width           =   11175
      Begin VB.TextBox tempAnnLmt 
         Height          =   285
         Left            =   5400
         TabIndex        =   50
         Text            =   "Text1"
         Top             =   240
         Visible         =   0   'False
         Width           =   975
      End
      Begin VB.TextBox txtDateJoin 
         Height          =   375
         Left            =   0
         TabIndex        =   49
         Text            =   "Text1"
         Top             =   240
         Visible         =   0   'False
         Width           =   255
      End
      Begin VB.TextBox txtMBMPremium 
         Height          =   285
         Left            =   3240
         TabIndex        =   48
         Top             =   240
         Visible         =   0   'False
         Width           =   735
      End
      Begin VB.TextBox txtMBMMCOFees 
         Height          =   285
         Left            =   3960
         TabIndex        =   47
         Top             =   240
         Visible         =   0   'False
         Width           =   735
      End
      Begin VB.TextBox txtAccessFees 
         Height          =   285
         Left            =   4680
         TabIndex        =   46
         Top             =   240
         Visible         =   0   'False
         Width           =   735
      End
      Begin VB.TextBox txtMBMAdultChild 
         Height          =   285
         Left            =   5400
         TabIndex        =   45
         Top             =   240
         Visible         =   0   'False
         Width           =   615
      End
      Begin VB.TextBox txtImportFile 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   1320
         TabIndex        =   34
         Text            =   "txtImportFile"
         Top             =   240
         Visible         =   0   'False
         Width           =   345
      End
      Begin VB.TextBox txtTempINSCode 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   1680
         TabIndex        =   33
         Text            =   "txtTempINSCode"
         Top             =   240
         Visible         =   0   'False
         Width           =   345
      End
      Begin VB.TextBox txtTempPLNCode 
         Height          =   285
         Left            =   2040
         TabIndex        =   32
         Text            =   "txtTempPLNCode"
         Top             =   240
         Visible         =   0   'False
         Width           =   375
      End
      Begin VB.TextBox txtTempRecordLine 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   240
         TabIndex        =   31
         Text            =   "txtTempRecordLine"
         Top             =   240
         Visible         =   0   'False
         Width           =   345
      End
      Begin VB.TextBox txtTempERRMessage 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   600
         TabIndex        =   30
         Text            =   "txtTempERRMessage"
         Top             =   240
         Visible         =   0   'False
         Width           =   345
      End
      Begin VB.TextBox txtTempMemberNo 
         Height          =   285
         Left            =   960
         TabIndex        =   29
         Text            =   "txtTempMemberNo"
         Top             =   240
         Visible         =   0   'False
         Width           =   375
      End
      Begin VB.CommandButton cmdStop 
         Caption         =   "Stop"
         Height          =   315
         Left            =   8040
         TabIndex        =   28
         Top             =   240
         Width           =   960
      End
      Begin VB.CommandButton CmdPrint 
         Caption         =   "&Print Err List"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   9000
         TabIndex        =   25
         Top             =   240
         Width           =   1080
      End
      Begin VB.CommandButton cmdExit 
         Cancel          =   -1  'True
         Caption         =   "&Exit"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   10080
         TabIndex        =   23
         Top             =   240
         Width           =   960
      End
      Begin MSComDlg.CommonDialog CommonDialog1 
         Left            =   2520
         Top             =   120
         _ExtentX        =   847
         _ExtentY        =   847
         _Version        =   393216
      End
   End
   Begin VB.Frame Frame3 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   435
      Left            =   0
      TabIndex        =   18
      Top             =   0
      Width           =   11295
      Begin VB.Label Label3 
         BackColor       =   &H00000000&
         Caption         =   "Upload Bordx"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000005&
         Height          =   315
         Left            =   180
         TabIndex        =   19
         Top             =   60
         Width           =   6615
      End
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   6195
      Left            =   120
      TabIndex        =   13
      Top             =   480
      Width           =   11055
      _ExtentX        =   19500
      _ExtentY        =   10927
      _Version        =   393216
      TabHeight       =   741
      TabCaption(0)   =   "Data Verification"
      TabPicture(0)   =   "frmEndorsement.frx":0000
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Label8"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Frame1"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "Frame2"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "lstERRList"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).ControlCount=   4
      TabCaption(1)   =   "Generate ASCII File"
      TabPicture(1)   =   "frmEndorsement.frx":001C
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Frame8"
      Tab(1).ControlCount=   1
      TabCaption(2)   =   "ASCII Status"
      TabPicture(2)   =   "frmEndorsement.frx":0038
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "Frame11"
      Tab(2).Control(1)=   "Frame12"
      Tab(2).Control(2)=   "lstASCIIList"
      Tab(2).ControlCount=   3
      Begin VB.Frame Frame11 
         Height          =   1935
         Left            =   -74880
         TabIndex        =   90
         Top             =   480
         Width           =   10815
         Begin VB.Frame Frame15 
            Height          =   1095
            Left            =   120
            TabIndex        =   92
            Top             =   120
            Width           =   10575
            Begin VB.ComboBox cboHTL 
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   330
               Left            =   1200
               Sorted          =   -1  'True
               TabIndex        =   95
               Text            =   "S - SIHAT MALAYSIA"
               Top             =   165
               Width           =   2385
            End
            Begin VB.ComboBox cboPAY 
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   330
               Left            =   4440
               Sorted          =   -1  'True
               TabIndex        =   94
               Top             =   165
               Width           =   5865
            End
            Begin VB.ComboBox cboBordxOPT 
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   330
               ItemData        =   "frmEndorsement.frx":0054
               Left            =   1200
               List            =   "frmEndorsement.frx":0056
               Sorted          =   -1  'True
               TabIndex        =   93
               Text            =   "B - Batch Bordx"
               Top             =   570
               Width           =   2385
            End
            Begin VB.Label Label35 
               Caption         =   "Health Type"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   255
               Left            =   240
               TabIndex        =   98
               Top             =   165
               Width           =   1365
            End
            Begin VB.Label Label32 
               Caption         =   "Payor"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   255
               Left            =   3720
               TabIndex        =   97
               Top             =   165
               Width           =   915
            End
            Begin VB.Label Label31 
               Caption         =   "Bordx Opt"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   255
               Left            =   240
               TabIndex        =   96
               Top             =   600
               Width           =   765
            End
         End
         Begin VB.CommandButton cmdSearch 
            Caption         =   "Search"
            Height          =   375
            Left            =   9720
            TabIndex        =   91
            Top             =   1320
            Width           =   975
         End
         Begin VB.Frame Frame10 
            Height          =   600
            Left            =   120
            TabIndex        =   104
            Top             =   1200
            Width           =   9495
            Begin VB.TextBox txtBATNo 
               Height          =   285
               Left            =   5730
               TabIndex        =   105
               Top             =   120
               Width           =   2500
            End
            Begin MSMask.MaskEdBox txtBorDate 
               Height          =   285
               Left            =   1530
               TabIndex        =   106
               Top             =   120
               Width           =   2505
               _ExtentX        =   4419
               _ExtentY        =   503
               _Version        =   393216
               MaxLength       =   10
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "Tahoma"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Mask            =   "##/##/####"
               PromptChar      =   "_"
            End
            Begin VB.Label Label34 
               Caption         =   "Batch No"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   255
               Left            =   4290
               TabIndex        =   108
               Top             =   120
               Width           =   1275
            End
            Begin VB.Label Label33 
               Caption         =   "Receipt Date "
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   255
               Left            =   120
               TabIndex        =   107
               Top             =   120
               Width           =   1515
            End
         End
         Begin VB.Frame Frame9 
            Height          =   600
            Left            =   120
            TabIndex        =   99
            Top             =   1200
            Visible         =   0   'False
            Width           =   9495
            Begin MSMask.MaskEdBox txtBorDateFrom 
               Height          =   285
               Left            =   1530
               TabIndex        =   100
               Top             =   180
               Width           =   2505
               _ExtentX        =   4419
               _ExtentY        =   503
               _Version        =   393216
               MaxLength       =   10
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "Tahoma"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Mask            =   "##/##/####"
               PromptChar      =   "_"
            End
            Begin MSMask.MaskEdBox txtBorDateTo 
               Height          =   285
               Left            =   5610
               TabIndex        =   101
               Top             =   180
               Width           =   2505
               _ExtentX        =   4419
               _ExtentY        =   503
               _Version        =   393216
               MaxLength       =   10
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "Tahoma"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Mask            =   "##/##/####"
               PromptChar      =   "_"
            End
            Begin VB.Label Label30 
               Caption         =   "Bordx Date To"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   255
               Left            =   4440
               TabIndex        =   103
               Top             =   180
               Width           =   1275
            End
            Begin VB.Label Label29 
               Caption         =   "Bordx Date From"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   255
               Left            =   90
               TabIndex        =   102
               Top             =   225
               Width           =   1275
            End
         End
      End
      Begin VB.Frame Frame12 
         Caption         =   "Status"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   570
         Left            =   -74880
         TabIndex        =   84
         Top             =   5520
         Width           =   10740
         Begin VB.Label Label37 
            Alignment       =   2  'Center
            Caption         =   "Supplementary"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000FF&
            Height          =   225
            Left            =   2640
            TabIndex        =   88
            Top             =   240
            Width           =   1320
         End
         Begin VB.Label txtPrinNo 
            Alignment       =   2  'Center
            BackColor       =   &H00C0FFFF&
            Caption         =   "0"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   270
            Left            =   1320
            TabIndex        =   87
            Top             =   240
            Width           =   1080
         End
         Begin VB.Label txtSuppNo 
            Alignment       =   2  'Center
            BackColor       =   &H00C0FFFF&
            Caption         =   "0"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   270
            Left            =   3960
            TabIndex        =   86
            Top             =   240
            Width           =   1200
         End
         Begin VB.Label Label38 
            Alignment       =   2  'Center
            Caption         =   "Principal"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000FF&
            Height          =   225
            Left            =   600
            TabIndex        =   85
            Top             =   240
            Width           =   720
         End
      End
      Begin VB.Frame Frame8 
         Height          =   3375
         Left            =   -74880
         TabIndex        =   51
         Top             =   480
         Width           =   10815
         Begin VB.Frame Frame14 
            Height          =   2535
            Left            =   120
            TabIndex        =   52
            Top             =   120
            Width           =   10455
            Begin VB.Frame BordxFrame 
               Height          =   600
               Left            =   120
               TabIndex        =   74
               Top             =   1800
               Visible         =   0   'False
               Width           =   9015
               Begin MSMask.MaskEdBox txtBordxTo 
                  Height          =   285
                  Left            =   5730
                  TabIndex        =   75
                  Top             =   180
                  Width           =   2505
                  _ExtentX        =   4419
                  _ExtentY        =   503
                  _Version        =   393216
                  MaxLength       =   10
                  BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                     Name            =   "Tahoma"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Mask            =   "##/##/####"
                  PromptChar      =   "_"
               End
               Begin MSMask.MaskEdBox txtBordxFrom 
                  Height          =   285
                  Left            =   1530
                  TabIndex        =   76
                  Top             =   180
                  Width           =   2505
                  _ExtentX        =   4419
                  _ExtentY        =   503
                  _Version        =   393216
                  MaxLength       =   10
                  BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                     Name            =   "Tahoma"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Mask            =   "##/##/####"
                  PromptChar      =   "_"
               End
               Begin VB.Label Label23 
                  Caption         =   "Bordx Date To"
                  BeginProperty Font 
                     Name            =   "Arial"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   255
                  Left            =   4290
                  TabIndex        =   78
                  Top             =   180
                  Width           =   1035
               End
               Begin VB.Label Label22 
                  Caption         =   "Bordx Date From"
                  BeginProperty Font 
                     Name            =   "Arial"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   255
                  Left            =   90
                  TabIndex        =   77
                  Top             =   225
                  Width           =   1275
               End
            End
            Begin VB.Frame BatchFrame 
               Height          =   600
               Left            =   120
               TabIndex        =   69
               Top             =   1800
               Width           =   9015
               Begin VB.TextBox txtBatchNo1 
                  Height          =   285
                  Left            =   5640
                  TabIndex        =   70
                  Top             =   180
                  Width           =   2500
               End
               Begin MSMask.MaskEdBox txtReceiptDate 
                  Height          =   285
                  Left            =   1530
                  TabIndex        =   71
                  Top             =   180
                  Width           =   2505
                  _ExtentX        =   4419
                  _ExtentY        =   503
                  _Version        =   393216
                  MaxLength       =   10
                  BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                     Name            =   "Tahoma"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Mask            =   "##/##/####"
                  PromptChar      =   "_"
               End
               Begin MSComDlg.CommonDialog CommonDialog2 
                  Left            =   8400
                  Top             =   120
                  _ExtentX        =   847
                  _ExtentY        =   847
                  _Version        =   393216
               End
               Begin VB.Label Label26 
                  Caption         =   "Receipt Date From"
                  BeginProperty Font 
                     Name            =   "Arial"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   255
                  Left            =   90
                  TabIndex        =   73
                  Top             =   225
                  Width           =   1515
               End
               Begin VB.Label Label27 
                  Caption         =   "Batch No"
                  BeginProperty Font 
                     Name            =   "Arial"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   255
                  Left            =   4680
                  TabIndex        =   72
                  Top             =   180
                  Width           =   675
               End
            End
            Begin VB.Frame Frame16 
               Height          =   1695
               Left            =   120
               TabIndex        =   54
               Top             =   120
               Width           =   10215
               Begin VB.ComboBox cboPAYCode 
                  BeginProperty Font 
                     Name            =   "Arial"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   330
                  Left            =   4800
                  Sorted          =   -1  'True
                  TabIndex        =   58
                  Top             =   720
                  Width           =   5295
               End
               Begin VB.ComboBox cboHLTCode 
                  BeginProperty Font 
                     Name            =   "Arial"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   330
                  Left            =   1440
                  Sorted          =   -1  'True
                  TabIndex        =   57
                  Text            =   "S - SIHAT MALAYSIA"
                  Top             =   720
                  Width           =   2025
               End
               Begin VB.ComboBox cboExportOption 
                  BeginProperty Font 
                     Name            =   "Arial"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   330
                  ItemData        =   "frmEndorsement.frx":0058
                  Left            =   4800
                  List            =   "frmEndorsement.frx":005A
                  Sorted          =   -1  'True
                  TabIndex        =   56
                  Text            =   "B - Batch Bordx"
                  Top             =   1200
                  Width           =   2025
               End
               Begin VB.ComboBox cboASCIIFileVer 
                  BeginProperty Font 
                     Name            =   "Arial"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   330
                  Left            =   1440
                  Sorted          =   -1  'True
                  TabIndex        =   55
                  Top             =   1200
                  Width           =   2025
               End
               Begin VB.Label Label25 
                  Caption         =   "Export File Name"
                  BeginProperty Font 
                     Name            =   "Arial"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   255
                  Left            =   6840
                  TabIndex        =   68
                  Top             =   270
                  Width           =   1245
               End
               Begin VB.Label Label28 
                  Caption         =   "File on PC"
                  BeginProperty Font 
                     Name            =   "Arial"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   225
                  Left            =   3720
                  TabIndex        =   66
                  Top             =   240
                  Width           =   765
               End
               Begin VB.Label txtFilName 
                  BackColor       =   &H00FFFFFF&
                  Height          =   255
                  Left            =   8160
                  TabIndex        =   65
                  Top             =   240
                  Width           =   1905
               End
               Begin VB.Label txtLocalFileLocation 
                  BackColor       =   &H00FFFFFF&
                  BeginProperty Font 
                     Name            =   "Arial"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   285
                  Left            =   4800
                  TabIndex        =   64
                  Top             =   240
                  Width           =   1905
               End
               Begin VB.Label txtPath 
                  BackColor       =   &H00FFFFFF&
                  BeginProperty Font 
                     Name            =   "Arial"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   285
                  Left            =   1440
                  TabIndex        =   63
                  Top             =   240
                  Width           =   2025
               End
               Begin VB.Label Label18 
                  Caption         =   "Payor"
                  BeginProperty Font 
                     Name            =   "Arial"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   255
                  Left            =   3720
                  TabIndex        =   62
                  Top             =   720
                  Width           =   915
               End
               Begin VB.Label Label19 
                  Caption         =   "Health Type"
                  BeginProperty Font 
                     Name            =   "Arial"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   255
                  Left            =   120
                  TabIndex        =   61
                  Top             =   720
                  Width           =   1365
               End
               Begin VB.Label Label21 
                  Caption         =   "Bordx Option"
                  BeginProperty Font 
                     Name            =   "Arial"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   255
                  Left            =   3720
                  TabIndex        =   60
                  Top             =   1200
                  Width           =   1365
               End
               Begin VB.Label Label36 
                  Caption         =   "File Version"
                  BeginProperty Font 
                     Name            =   "Arial"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   255
                  Left            =   120
                  TabIndex        =   59
                  Top             =   1200
                  Width           =   1365
               End
               Begin VB.Label Label24 
                  Caption         =   "File on Server"
                  BeginProperty Font 
                     Name            =   "Arial"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   225
                  Left            =   120
                  TabIndex        =   67
                  Top             =   240
                  Width           =   1365
               End
            End
            Begin VB.CommandButton cmdGenerate 
               Caption         =   "&Ok"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   315
               Left            =   9360
               TabIndex        =   53
               Top             =   1920
               Width           =   960
            End
         End
         Begin VB.Frame Frame7 
            Caption         =   "Status"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   570
            Left            =   120
            TabIndex        =   79
            Top             =   2640
            Width           =   10500
            Begin VB.Label Label17 
               Alignment       =   2  'Center
               Caption         =   "Total Principal"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H000000FF&
               Height          =   225
               Left            =   120
               TabIndex        =   83
               Top             =   240
               Width           =   1080
            End
            Begin VB.Label txtTotalSupplementary1 
               Alignment       =   2  'Center
               BackColor       =   &H00C0FFFF&
               Caption         =   "0"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   270
               Left            =   4320
               TabIndex        =   81
               Top             =   240
               Width           =   1200
            End
            Begin VB.Label txtTotalPrincipals 
               Alignment       =   2  'Center
               BackColor       =   &H00C0FFFF&
               Caption         =   "0"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   270
               Left            =   1200
               TabIndex        =   80
               Top             =   240
               Width           =   1320
            End
            Begin VB.Label Label16 
               Alignment       =   2  'Center
               Caption         =   "Total Supplementary"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H000000FF&
               Height          =   225
               Left            =   2400
               TabIndex        =   82
               Top             =   240
               Width           =   2040
            End
         End
      End
      Begin MSComctlLib.ListView lstERRList 
         Height          =   2895
         Left            =   120
         TabIndex        =   21
         Top             =   2580
         Width           =   10815
         _ExtentX        =   19076
         _ExtentY        =   5106
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
         NumItems        =   5
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Rec Line"
            Object.Width           =   1411
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Policy No"
            Object.Width           =   2646
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "Name"
            Object.Width           =   3528
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   3
            Text            =   "Error Message"
            Object.Width           =   7056
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   4
            Text            =   "Filename"
            Object.Width           =   1764
         EndProperty
      End
      Begin VB.Frame Frame2 
         Caption         =   "Status"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   120
         TabIndex        =   35
         Top             =   5445
         Width           =   10785
         Begin VB.Label txtBatchNo 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00C0FFFF&
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   270
            Left            =   9720
            TabIndex        =   110
            Top             =   240
            Width           =   720
         End
         Begin VB.Label Label20 
            Caption         =   "Bat No"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000FF&
            Height          =   255
            Left            =   9120
            TabIndex        =   109
            Top             =   240
            Width           =   495
         End
         Begin VB.Label Label6 
            Alignment       =   2  'Center
            Caption         =   "Total Errors"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000FF&
            Height          =   225
            Left            =   2760
            TabIndex        =   43
            Top             =   240
            Width           =   960
         End
         Begin VB.Label txtTotalError 
            Alignment       =   2  'Center
            BackColor       =   &H00C0FFFF&
            Caption         =   "0"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   270
            Left            =   3720
            TabIndex        =   42
            Top             =   240
            Width           =   960
         End
         Begin VB.Label Label2 
            Alignment       =   2  'Center
            Caption         =   "Submission Date"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000FF&
            Height          =   225
            Left            =   120
            TabIndex        =   41
            Top             =   225
            Width           =   1425
         End
         Begin VB.Label txtSubmissionDate 
            Alignment       =   2  'Center
            BackColor       =   &H00C0FFFF&
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   270
            Left            =   1560
            TabIndex        =   40
            Top             =   240
            Width           =   1200
         End
         Begin VB.Label Label7 
            Alignment       =   2  'Center
            Caption         =   "Total Principal"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000FF&
            Height          =   225
            Left            =   4680
            TabIndex        =   39
            Top             =   240
            Width           =   1200
         End
         Begin VB.Label txtTotalPrincipal 
            Alignment       =   2  'Center
            BackColor       =   &H00C0FFFF&
            Caption         =   "0"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   270
            Left            =   5880
            TabIndex        =   37
            Top             =   240
            Width           =   945
         End
         Begin VB.Label txtTotalSupplementary 
            Alignment       =   2  'Center
            BackColor       =   &H00C0FFFF&
            Caption         =   "0"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   270
            Left            =   7800
            TabIndex        =   36
            Top             =   240
            Width           =   960
         End
         Begin VB.Label Label11 
            Alignment       =   2  'Center
            Caption         =   "Total Supp"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000FF&
            Height          =   225
            Left            =   6840
            TabIndex        =   38
            Top             =   240
            Width           =   960
         End
      End
      Begin VB.Frame Frame1 
         Height          =   2055
         Left            =   120
         TabIndex        =   14
         Top             =   480
         Width           =   10815
         Begin VB.ComboBox cboHLTCode1 
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   330
            Left            =   1515
            Sorted          =   -1  'True
            TabIndex        =   0
            Tag             =   "Thanks"
            Text            =   "S - Sihat Malaysia"
            Top             =   420
            Width           =   3030
         End
         Begin VB.ComboBox txtPAYCode 
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   330
            Left            =   1515
            Sorted          =   -1  'True
            Style           =   2  'Dropdown List
            TabIndex        =   1
            Top             =   720
            Width           =   3030
         End
         Begin VB.ComboBox txtFileName 
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   330
            Left            =   1515
            TabIndex        =   2
            Text            =   "N - New File"
            Top             =   1020
            Width           =   3030
         End
         Begin VB.CommandButton cmdFileLocation 
            Caption         =   " >"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   4240
            TabIndex        =   5
            Top             =   1635
            Width           =   320
         End
         Begin VB.ComboBox txtFileVersion 
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   330
            Left            =   1515
            Sorted          =   -1  'True
            TabIndex        =   3
            Top             =   1320
            Width           =   3030
         End
         Begin VB.Frame Frame13 
            Height          =   615
            Left            =   8280
            TabIndex        =   44
            Top             =   1320
            Width           =   2175
            Begin VB.CommandButton cmdUpload 
               Caption         =   "&Upload"
               Enabled         =   0   'False
               Height          =   315
               Left            =   1080
               TabIndex        =   12
               Top             =   200
               Width           =   960
            End
            Begin VB.CommandButton cmdVerify 
               Caption         =   "&Verify"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   315
               Left            =   120
               TabIndex        =   11
               Top             =   200
               Width           =   960
            End
         End
         Begin VB.TextBox txtImportFilename 
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   1515
            TabIndex        =   4
            Top             =   1620
            Width           =   2625
         End
         Begin VB.ComboBox cboSuppAnnualLimit 
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   330
            Left            =   6000
            Sorted          =   -1  'True
            TabIndex        =   6
            Text            =   "C - Combined"
            Top             =   420
            Width           =   1695
         End
         Begin VB.ComboBox CboVerifyIC 
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   330
            Left            =   6000
            Sorted          =   -1  'True
            Style           =   2  'Dropdown List
            TabIndex        =   7
            Top             =   720
            Width           =   1695
         End
         Begin VB.ComboBox CboBack2Back 
            Height          =   315
            Left            =   6000
            Sorted          =   -1  'True
            Style           =   2  'Dropdown List
            TabIndex        =   8
            Top             =   1020
            Width           =   1695
         End
         Begin VB.TextBox txtDiscount 
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   9000
            TabIndex        =   9
            Top             =   300
            Width           =   1425
         End
         Begin VB.ComboBox cboLBLStatus 
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   330
            Left            =   9000
            Sorted          =   -1  'True
            Style           =   2  'Dropdown List
            TabIndex        =   10
            Top             =   585
            Width           =   1425
         End
         Begin VB.Label Label15 
            Caption         =   "BTB Grantee"
            Height          =   255
            Left            =   4800
            TabIndex        =   114
            Top             =   1050
            Width           =   1215
         End
         Begin VB.Label Label14 
            Caption         =   "Label Status"
            Height          =   255
            Left            =   7920
            TabIndex        =   113
            Top             =   600
            Width           =   1215
         End
         Begin VB.Label Label13 
            Caption         =   "Verify I/C"
            Height          =   255
            Left            =   4800
            TabIndex        =   112
            Top             =   735
            Width           =   855
         End
         Begin VB.Label Label9 
            Caption         =   "Discount"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   7920
            TabIndex        =   111
            Top             =   360
            Width           =   855
         End
         Begin VB.Label Label12 
            Caption         =   "Supp Type"
            Height          =   255
            Left            =   4800
            TabIndex        =   27
            Top             =   435
            Width           =   975
         End
         Begin VB.Label Label10 
            Caption         =   "Health Type"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   240
            TabIndex        =   26
            Top             =   375
            Width           =   1770
         End
         Begin VB.Label Label5 
            Caption         =   "New/ Endt File"
            Height          =   255
            Left            =   240
            TabIndex        =   22
            Top             =   1065
            Width           =   1095
         End
         Begin VB.Label Label1 
            Caption         =   "Filename"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   240
            TabIndex        =   17
            Top             =   1680
            Width           =   975
         End
         Begin VB.Label Label55 
            Caption         =   "Payor"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   240
            TabIndex        =   16
            Top             =   720
            Width           =   1170
         End
         Begin VB.Label Label4 
            Caption         =   "File Version"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   240
            TabIndex        =   15
            Top             =   1380
            Width           =   960
         End
      End
      Begin MSComctlLib.ListView lstASCIIList 
         Height          =   3015
         Left            =   -74880
         TabIndex        =   89
         Top             =   2520
         Width           =   10815
         _ExtentX        =   19076
         _ExtentY        =   5318
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
         NumItems        =   7
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Date Generated"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "File Name"
            Object.Width           =   2646
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "Bordx Date"
            Object.Width           =   3528
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   3
            Text            =   "Batch No"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   4
            Text            =   "Payor"
            Object.Width           =   7056
         EndProperty
         BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   5
            Text            =   "Principal"
            Object.Width           =   1764
         EndProperty
         BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   6
            Text            =   "Supplementary"
            Object.Width           =   2540
         EndProperty
      End
      Begin VB.Label Label8 
         Alignment       =   2  'Center
         Caption         =   "Error Listing"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   225
         Left            =   120
         TabIndex        =   24
         Top             =   1740
         Width           =   1155
      End
   End
End
Attribute VB_Name = "frmEndorsement"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim MBMCover As New ADODB.Recordset
Dim MBM As New ADODB.Recordset
Dim MBMOthers As New ADODB.Recordset
Dim CLNOthers As New ADODB.Recordset
Dim ERR As New ADODB.Recordset
Dim MBMCoveredPersons As New ADODB.Recordset
Dim TextLine
Dim Message As String
Dim RECLine As String
Dim recordCount As Integer
Dim CoveredID As Integer
Dim CoveredIDCount As String
Dim Exportsql As String
Dim textField As String
Dim header As String
Dim DataString As String
Dim Trailer As String
Dim VNewPayor, MNumber As String
Dim INSCode As String
Dim MBMPolicyNo As String
Dim MBMEmployeeNo As String
Dim PayEffDate As String
Dim PayorEXPDate As String
Dim InsuredCode As String
Dim PlanCode As String
Dim TotalPrincipal As Integer
Dim TotalSupplementary As Integer
Dim ii As Integer
Dim Cnt As Integer
Dim BasicPrem As String
Dim BasicMCO As String
Dim BasicEA As String
Dim CisPlnCode As String
Dim CisInsType As String
Dim CisExpDate, CisEffDate As String
Dim MBMGroupCompany As String
Dim MBMPolSubNo1 As String
Dim MBMCRELCode As String
Dim Check As Boolean
Dim MBMDateJoin As String

Public intExit As Integer

Private Sub CboBack2Back_Change()
    If InStr(CboBack2Back.Text, "null") Then
       CboBack2Back.Enabled = False
    End If
End Sub

Private Sub CboBack2Back_KeyPress(KeyAscii As Integer)
Dim NewText As String
Dim ValidCount As Integer
Dim ValidValue As String
Dim i As Integer
    
    If KeyAscii >= 32 And KeyAscii <> 127 Then
    NewText = LCase(Left(CboBack2Back.Text, CboBack2Back.SelStart) + Chr(KeyAscii) + Mid(CboBack2Back.Text, CboBack2Back.SelStart + CboBack2Back.SelLength + 1))

    ValidCount = 0
    ValidValue = ""

    For i = 0 To CboBack2Back.ListCount - 1
        If NewText = LCase(Left(CboBack2Back.List(i), Len(NewText))) Then
                ValidCount = ValidCount + 1
                ValidValue = CboBack2Back.List(i)
        End If
    Next
    If ValidCount <= 1 Then KeyAscii = 0
         If ValidCount = 1 Then
           CboBack2Back.Text = ValidValue
           CboBack2Back.SelStart = 0
           CboBack2Back.SelLength = Len(ValidValue)
         End If
    End If
End Sub

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

Private Sub cboBordxOPT_Click()
    If Mid(cboBordxOPT, 1, 1) = "B" Then
        Frame10.Visible = True
        txtBorDate.SetFocus
        txtBorDate.TabIndex = 19
        txtBATNo.TabIndex = 20
        cmdSearch.TabIndex = 21
    ElseIf Mid(cboBordxOPT, 1, 1) = "D" Then
        Frame9.Visible = True
        txtBorDateFrom.SetFocus
        txtBorDateFrom.TabIndex = 19
        txtBorDateTo.TabIndex = 20
        cmdSearch.TabIndex = 21
    End If
End Sub

Private Sub cboBordxOPT_LostFocus()
    If Mid(cboBordxOPT, 1, 1) = "B" Then
        Frame10.Visible = True
        txtBorDate.SetFocus
        txtBorDate.TabIndex = 19
        txtBATNo.TabIndex = 20
        cmdSearch.TabIndex = 21
    ElseIf Mid(cboBordxOPT, 1, 1) = "D" Then
        Frame9.Visible = True
        txtBorDateFrom.SetFocus
        txtBorDateFrom.TabIndex = 19
        txtBorDateTo.TabIndex = 20
        cmdSearch.TabIndex = 21
    End If
End Sub

Private Sub cboExportOption_Click()
    If Mid(cboExportOption, 1, 1) = "B" Then
        BatchFrame.Visible = True
        BordxFrame.Visible = False
        txtReceiptDate.SetFocus
        txtReceiptDate.TabIndex = 13
        txtBatchNo1.TabIndex = 14
        cmdGenerate.TabIndex = 15
        'MethodFrame.Visible = True
        txtFileName = ""
    ElseIf Mid(cboExportOption, 1, 1) = "D" Then
        BatchFrame.Visible = False
        BordxFrame.Visible = True
        txtBordxFrom.SetFocus
        txtBordxFrom.TabIndex = 13
        txtBordxTo.TabIndex = 14
        cmdGenerate.TabIndex = 15
        'MethodFrame.Visible = True
        txtFileName = ""
    End If

End Sub

Private Sub cboExportOption_Change()
    If InStr(cboExportOption.Text, "null") Then
       cboExportOption.Enabled = False
    End If
End Sub

Private Sub cboExportOption_KeyPress(KeyAscii As Integer)
Dim NewText As String
Dim ValidCount As Integer
Dim ValidValue As String
Dim i As Integer
    
    If KeyAscii >= 32 And KeyAscii <> 127 Then
    NewText = LCase(Left(cboExportOption.Text, cboExportOption.SelStart) + Chr(KeyAscii) + Mid(cboExportOption.Text, cboExportOption.SelStart + cboExportOption.SelLength + 1))

    ValidCount = 0
    ValidValue = ""

    For i = 0 To cboExportOption.ListCount - 1
 
        If NewText = LCase(Left(cboExportOption.List(i), Len(NewText))) Then
            ValidCount = ValidCount + 1
            ValidValue = cboExportOption.List(i)
        End If
    Next
    If ValidCount <= 1 Then KeyAscii = 0
         If ValidCount = 1 Then
           cboExportOption.Text = ValidValue
           cboExportOption.SelStart = 0
           cboExportOption.SelLength = Len(ValidValue)
         End If
    End If
End Sub

Private Sub cboExportOption_LostFocus()
    If Mid(cboExportOption, 1, 1) = "B" Then
        BatchFrame.Visible = True
        BordxFrame.Visible = False
        txtReceiptDate.SetFocus
        txtReceiptDate.TabIndex = 13
        txtBatchNo1.TabIndex = 14
        cmdGenerate.TabIndex = 15
        'MethodFrame.Visible = True
        txtFileName = ""
    ElseIf Mid(cboExportOption, 1, 1) = "D" Then
        BatchFrame.Visible = False
        BordxFrame.Visible = True
        txtBordxFrom.SetFocus
        txtBordxFrom.TabIndex = 13
        txtBordxTo.TabIndex = 14
        cmdGenerate.TabIndex = 15
        'MethodFrame.Visible = True
        txtFileName = ""
    End If

End Sub

Private Sub cboHLTCode_Change()
Dim PAY As New ADODB.Recordset
Dim strsql As String
    
    strsql = "Select  PAYCode, PAYCompanyName From PAY where HLTCOde ='" & Mid(Trim(cboHLTCode), 1, 1) & "'"
    PAY.Open strsql, medixmasterconn, adOpenKeyset, adLockOptimistic
    cboPAYCode.Clear
    Do Until PAY.EOF
       cboPAYCode.AddItem ChkStr(PAY!PAYCode) & " - " & ChkStr(PAY!PAYCompanyName)
       PAY.MoveNext
    Loop
    PAY.Close
    Set PAY = Nothing

End Sub

Private Sub cboHLTCode_Click()
Dim PAY As New ADODB.Recordset
Dim strsql As String
    
    strsql = "Select  PAYCode, PAYCompanyName From PAY where HLTCOde ='" & Mid(Trim(cboHLTCode), 1, 1) & "'"
    PAY.Open strsql, medixmasterconn, adOpenKeyset, adLockOptimistic
    cboPAYCode.Clear
    Do Until PAY.EOF
       cboPAYCode.AddItem ChkStr(PAY!PAYCode) & " - " & ChkStr(PAY!PAYCompanyName)
       PAY.MoveNext
    Loop
    PAY.Close
    Set PAY = Nothing
End Sub

Private Sub cboHLTcode_KeyPress(KeyAscii As Integer)
Dim NewText As String
Dim ValidCount As Integer
Dim ValidValue As String
Dim i As Integer
    
    If KeyAscii >= 32 And KeyAscii <> 127 Then
    NewText = LCase(Left(cboHLTCode.Text, cboHLTCode.SelStart) + Chr(KeyAscii) + Mid(cboHLTCode.Text, cboHLTCode.SelStart + cboHLTCode.SelLength + 1))

    ValidCount = 0
    ValidValue = ""

    For i = 0 To cboHLTCode.ListCount - 1
        If NewText = LCase(Left(cboHLTCode.List(i), Len(NewText))) Then
            ValidCount = ValidCount + 1
            ValidValue = cboHLTCode.List(i)
        End If
    Next
    If ValidCount <= 1 Then KeyAscii = 0
         If ValidCount = 1 Then
           cboHLTCode.Text = ValidValue
           cboHLTCode.SelStart = 0
           cboHLTCode.SelLength = Len(ValidValue)
         End If
    End If
End Sub

Private Sub cboHLTCode1_Change()
Dim PAY As New ADODB.Recordset
Dim strsql As String
    
    strsql = "Select  PAYCode, PAYCompanyName From PAY where HLTCOde ='" & Mid(Trim(cboHLTCode1), 1, 1) & "'"
    PAY.Open strsql, medixmasterconn, adOpenKeyset, adLockOptimistic
    txtPAYCode.Clear
    Do Until PAY.EOF
       txtPAYCode.AddItem ChkStr(PAY!PAYCode) & " - " & ChkStr(PAY!PAYCompanyName)
       PAY.MoveNext
    Loop
    PAY.Close
    Set PAY = Nothing
End Sub

Private Sub cboHLTCode1_Click()
Dim PAY As New ADODB.Recordset
Dim strsql As String
    
    strsql = "Select  PAYCode, PAYCompanyName From PAY where HLTCOde ='" & Mid(Trim(cboHLTCode1), 1, 1) & "'"
    PAY.Open strsql, medixmasterconn, adOpenKeyset, adLockOptimistic
    txtPAYCode.Clear
    Do Until PAY.EOF
       txtPAYCode.AddItem ChkStr(PAY!PAYCode) & " - " & ChkStr(PAY!PAYCompanyName)
       PAY.MoveNext
    Loop
    PAY.Close
    Set PAY = Nothing
End Sub

Private Sub cboHLTCode1_KeyPress(KeyAscii As Integer)
Dim NewText As String
Dim ValidCount As Integer
Dim ValidValue As String
Dim i As Integer
    If KeyAscii >= 32 And KeyAscii <> 127 Then
    NewText = LCase(Left(cboHLTCode1.Text, cboHLTCode1.SelStart) + Chr(KeyAscii) + Mid(cboHLTCode1.Text, cboHLTCode1.SelStart + cboHLTCode1.SelLength + 1))
    
    ValidCount = 0
    ValidValue = ""
    
    For i = 0 To cboHLTCode1.ListCount - 1
        
        If NewText = LCase(Left(cboHLTCode1.List(i), Len(NewText))) Then
                ValidCount = ValidCount + 1
                ValidValue = cboHLTCode1.List(i)
        End If
    Next
    If ValidCount <= 1 Then KeyAscii = 0
        If ValidCount = 1 Then
           cboHLTCode1.Text = ValidValue
           cboHLTCode1.SelStart = 0
           cboHLTCode1.SelLength = Len(ValidValue)
        End If
    End If

End Sub

Private Sub cboHTL_Change()
Dim PAY As New ADODB.Recordset
Dim strsql As String
    
    strsql = "Select  PAYCode, PAYCompanyName From PAY where HLTCOde ='" & Mid(Trim(cboHTL), 1, 1) & "'"
    PAY.Open strsql, medixmasterconn, adOpenKeyset, adLockOptimistic
    cboPAY.Clear
    Do Until PAY.EOF
       cboPAY.AddItem ChkStr(PAY!PAYCode) & " - " & ChkStr(PAY!PAYCompanyName)
       PAY.MoveNext
    Loop
    PAY.Close
    Set PAY = Nothing
End Sub

Private Sub cboHTL_Click()
Dim PAY As New ADODB.Recordset
Dim strsql As String
    
    strsql = "Select  PAYCode, PAYCompanyName From PAY where HLTCOde ='" & Mid(Trim(cboHTL), 1, 1) & "'"
    PAY.Open strsql, medixmasterconn, adOpenKeyset, adLockOptimistic
    cboPAY.Clear
    Do Until PAY.EOF
       cboPAY.AddItem ChkStr(PAY!PAYCode) & " - " & ChkStr(PAY!PAYCompanyName)
       PAY.MoveNext
    Loop
    PAY.Close
    Set PAY = Nothing

End Sub

Private Sub cboPAYCode_Change()
    If InStr(cboPAYCode.Text, "null") Then
       cboPAYCode.Enabled = False
    End If
End Sub

Private Sub cboPAYCode_KeyPress(KeyAscii As Integer)
Dim NewText As String
Dim ValidCount As Integer
Dim ValidValue As String
Dim i As Integer
    
    If KeyAscii >= 32 And KeyAscii <> 127 Then
    NewText = LCase(Left(cboPAYCode.Text, cboPAYCode.SelStart) + Chr(KeyAscii) + Mid(cboPAYCode.Text, cboPAYCode.SelStart + cboPAYCode.SelLength + 1))

    ValidCount = 0
    ValidValue = ""

    For i = 0 To cboPAYCode.ListCount - 1
        If NewText = LCase(Left(cboPAYCode.List(i), Len(NewText))) Then
            ValidCount = ValidCount + 1
            ValidValue = cboPAYCode.List(i)
        End If
    Next
    If ValidCount <= 1 Then KeyAscii = 0
         If ValidCount = 1 Then
           cboPAYCode.Text = ValidValue
           cboPAYCode.SelStart = 0
           cboPAYCode.SelLength = Len(ValidValue)
         End If
    End If
End Sub

Private Sub CboVerifyIC_Change()
    If InStr(CboVerifyIC.Text, "null") Then
       CboVerifyIC.Enabled = False
    End If
End Sub

Private Sub CboVerifyIC_KeyPress(KeyAscii As Integer)
Dim NewText As String
Dim ValidCount As Integer
Dim ValidValue As String
Dim i As Integer

If KeyAscii >= 32 And KeyAscii <> 127 Then
    'NewText = LCase(Left(CboVerifyIC.Text, CboVerifyIC.SelStart) + Chr(KeyAscii) + Mid(CboVerifyIC.Text, CboVerifyIC.SelStart + CboVerifyIC.SelLength + 1))
    
    ValidCount = 0
    ValidValue = ""
    
    For i = 0 To CboVerifyIC.ListCount - 1
        
        If NewText = LCase(Left(CboVerifyIC.List(i), Len(NewText))) Then
                ValidCount = ValidCount + 1
                ValidValue = CboVerifyIC.List(i)
        End If
        Next
        If ValidCount <= 1 Then KeyAscii = 0
        
             If ValidCount = 1 Then
               CboVerifyIC.Text = ValidValue
               CboVerifyIC.SelStart = 0
               CboVerifyIC.SelLength = Len(ValidValue)
             End If
        End If
End Sub

Private Sub cboSuppAnnualLimit_Change()
    If InStr(cboSuppAnnualLimit.Text, "null") Then
       cboSuppAnnualLimit.Enabled = False
    End If
End Sub

Private Sub cboSuppAnnualLimit_KeyPress(KeyAscii As Integer)
Dim NewText As String
Dim ValidCount As Integer
Dim ValidValue As String
Dim i As Integer
    
    If KeyAscii >= 32 And KeyAscii <> 127 Then
    NewText = LCase(Left(cboSuppAnnualLimit.Text, cboSuppAnnualLimit.SelStart) + Chr(KeyAscii) + Mid(cboSuppAnnualLimit.Text, cboSuppAnnualLimit.SelStart + cboSuppAnnualLimit.SelLength + 1))

    ValidCount = 0
    ValidValue = ""

    For i = 0 To cboSuppAnnualLimit.ListCount - 1
 
        If NewText = LCase(Left(cboSuppAnnualLimit.List(i), Len(NewText))) Then
            ValidCount = ValidCount + 1
            ValidValue = cboSuppAnnualLimit.List(i)
        End If
    Next
    If ValidCount <= 1 Then KeyAscii = 0
         If ValidCount = 1 Then
           cboSuppAnnualLimit.Text = ValidValue
           cboSuppAnnualLimit.SelStart = 0
           cboSuppAnnualLimit.SelLength = Len(ValidValue)
         End If
    End If
End Sub


Private Sub cmdExit_Click()
'Dim RecordExit
'    RecordExit = MsgBox("Quit this form?", vbYesNo + vbExclamation)
'    If RecordExit = vbYes Then
        'bExit = True
        Unload Me
'    End If

End Sub

Private Sub cmdFileLocation_Click()
  CommonDialog1.CancelError = True
  On Error GoTo ErrHandler
  CommonDialog1.InitDir = "\\MARS\" & "Bor_In"
  CommonDialog1.Flags = cdlOFNHideReadOnly
  CommonDialog1.Filter = "All Files (*.*)|*.*|Text Files (*.txt)|*.txt|Dat Files (*.Dat)|*.Dat"
  'CommonDialog1.FilterIndex = 2
  CommonDialog1.ShowOpen
  txtImportFilename = CommonDialog1.FileTitle
  Exit Sub
  
ErrHandler:
  'User pressed the Cancel button
  Exit Sub
End Sub

Private Sub CmdGenerate_Click()
Dim strsql As String
Dim TotalPrincipal As Integer
Dim TotalSupplementary As Integer
Dim RecordSaved
Dim MBMOthers As New ADODB.Recordset
Dim MBMCoveredPersons As New ADODB.Recordset
Dim ASCII As New ADODB.Recordset

    If cboHLTCode = "" Then
        MsgBox "Please Enter Health Plan", vbOKOnly + vbCritical, DialogBoxTitle
    ElseIf cboPAYCode = "" Then
        MsgBox "Please Enter Payor Code", vbOKOnly + vbCritical, DialogBoxTitle
    ElseIf cboASCIIFileVer = "" Then
        MsgBox "Please Select ASCII Version", vbOKOnly + vbCritical, DialogBoxTitle
    ElseIf cboExportOption = "" Then
        MsgBox "Please Enter Export Option", vbOKOnly + vbCritical, DialogBoxTitle
    ElseIf (Mid(cboExportOption, 1, 1) = "B") And (txtReceiptDate = "__/__/____" Or txtBatchNo1 = "") Then
        MsgBox "Please Enter Receipt Date or Batch No", vbOKOnly + vbCritical, DialogBoxTitle
    ElseIf (Mid(cboExportOption, 1, 1) = "D") And (txtBordxFrom = "__/__/____" Or txtBordxTo = "__/__/____") Then
        MsgBox "Please Enter Bordx Date", vbOKOnly + vbCritical, DialogBoxTitle
    'ElseIf Option1 = False And Option2 = False Then
        MsgBox "Please Select Your Output Option", vbOKOnly + vbInformation, DialogBoxTitle
    Else
      RecordSaved = MsgBox("Are you sure?", vbYesNo + vbExclamation)
        If RecordSaved = vbYes Then
    
            Screen.MousePointer = vbHourglass
            If Mid(cboExportOption, 1, 1) = "B" Then
                textField = Trim(txtReceiptDate)
                Exportsql = " HLTCode = '" & Trim(Mid(cboHLTCode, 1, 1)) & _
                         "' And PAYCode = '" & Trim(Mid(cboPAYCode, 1, 2)) & "' " & _
                         " And MBMBordxDate = '" & Format(txtReceiptDate, "mm/dd/yyyy") & "'" & _
                         " And MBMStatus <> 'D' " & _
                         " And MBMDataStatus <> 'D' " & _
                          " And mbmbatchno = " & Trim(txtBatchNo1)
            ElseIf Mid(cboExportOption, 1, 1) = "D" Then
                textField = txtBordxTo
                Exportsql = "HLTCode = '" & Trim(Mid(cboHLTCode, 1, 1)) & _
                            "' And PAYCode = '" & Trim(Mid(cboPAYCode, 1, 2)) & "'" & _
                            " And MBMStatus <> 'D' " & _
                            " And MBMDataStatus <> 'D' " & _
                            " And MBMBordxDate >=  '" & Format(Trim(txtBordxFrom), "mm/dd/yyyy") & "'" & _
                            " And MBMBordxDate <= '" & Format(Trim(txtBordxTo), "mm/dd/yyyy") & "'"
            End If
                           
            strsql = "SELECT * From MBM Where " & Exportsql
            MBM.Open strsql, medixmasterconn, adOpenKeyset, adLockOptimistic

            If MBM.recordCount = 0 Then
                MsgBox "No Record Available to Export To ASCII File", vbOKOnly + vbCritical, DialogBoxTitle
                Screen.MousePointer = vbDefault
            MBM.Close
            Set MBM = Nothing
            
            Else
                If txtBatchNo1 <> "" Then
                    txtFilName = Trim(Mid(cboPAYCode, 1, 2) & Trim(Format(txtReceiptDate, "YYMMDD"))) & "(" & txtBatchNo1 & ")"
                Else
                    txtFilName = Trim(Mid(cboPAYCode, 1, 2) & Format(textField, "YYMMDD"))
                End If
                Open txtPath & "\" & txtFilName For Output As #1
                Open txtLocalFileLocation & txtFilName For Output As #2
            If Mid(cboExportOption, 1, 1) = "B" Then
                header = Trim("F" & Mid(cboPAYCode, 1, 2) & Mid(txtReceiptDate, 1, 2) & Mid(txtReceiptDate, 4, 2) & Mid(txtReceiptDate, 7, 4) & l("", 174))
            ElseIf Mid(cboExportOption, 1, 1) = "D" Then
                header = Trim("F" & Mid(cboPAYCode, 1, 2) & Mid(txtBordxFrom, 1, 2) & Mid(txtBordxFrom, 4, 2) & Mid(txtBordxFrom, 7, 4) & l("", 174))
            End If
                Print #1, Trim(header)
                Print #2, Trim(header)
            
                Call PrincipalASCII
                Trailer = "T" & l(txtTotalPrincipals, 10) & l(txtTotalSupplementary1, 10) & l("", 164)
                Print #1, Trailer
                Print #2, Trailer
                cmdGenerate.Enabled = False
                Close #1
                Close #2
                MsgBox "File Generated Sucessfully...", vbOKOnly + vbInformation, DialogBoxTitle
            End If
            Screen.MousePointer = vbDefault
        End If
    End If
    txtTotalPrincipals = l(txtTotalPrincipals, 10)
    txtTotalSupplementary1 = l(txtTotalSupplementary1, 10)
End Sub

Private Sub CmdPrint_Click()
    Call ExportToExcelError(lstERRList)
End Sub

Private Sub cmdSearch_Click()
Dim MBM As New ADODB.Recordset
Dim strsql As String
Dim TotalPrincipal As Integer
Dim TotalSupplementary As Integer
Dim RecordSaved
Dim ASCII As New ADODB.Recordset
Dim ASCIIListing As ListItem
Dim ASCIIsql As String

    If cboHTL = "" Then
        MsgBox "Please Enter Health Plan", vbOKOnly + vbCritical, DialogBoxTitle
    ElseIf cboPAY = "" Then
        MsgBox "Please Enter Payor Code", vbOKOnly + vbCritical, DialogBoxTitle
    ElseIf cboBordxOPT = "" Then
        MsgBox "Please Enter Export Option", vbOKOnly + vbCritical, DialogBoxTitle
    ElseIf (Mid(cboBordxOPT, 1, 1) = "B") And (txtBorDate = "__/__/____" Or txtBATNo = "") Then
        MsgBox "Please Enter Receipt Date or Batch No", vbOKOnly + vbCritical, DialogBoxTitle
    ElseIf (Mid(cboBordxOPT, 1, 1) = "D") And (txtBordxFrom = "__/__/____" Or txtBordxTo = "__/__/____") Then
        MsgBox "Please Enter Bordx Date", vbOKOnly + vbCritical, DialogBoxTitle
    Else
        Screen.MousePointer = vbHourglass
        If Mid(cboExportOption, 1, 1) = "B" Then
            textField = Trim(txtReceiptDate)
            ASCIIsql = " HLTCode = '" & Trim(Mid(cboHTL, 1, 1)) & _
                     "' And PAYCode = '" & Trim(Mid(cboPAY, 1, 2)) & "' " & _
                     " And MBMBordxDate = '" & Format(txtBorDate, "mm/dd/yyyy") & "'" & _
                      " And mbmbatchno = " & Trim(txtBATNo)
        ElseIf Mid(cboExportOption, 1, 1) = "D" Then
            textField = txtBordxTo
            ASCIIsql = "HLTCode = '" & Trim(Mid(cboHTL, 1, 1)) & _
                        "' And PAYCode = '" & Trim(Mid(cboPAY, 1, 2)) & "'" & _
                        " And MBMBordxDate >=  '" & Format(Trim(txtBorDateFrom), "mm/dd/yyyy") & "'" & _
                        " And MBMBordxDate <= '" & Format(Trim(txtBorDateTo), "mm/dd/yyyy") & "'"
        End If

        strsql = "SELECT * From ASCII Where " & ASCIIsql
        ASCII.Open strsql, medixmasterconn, adOpenKeyset, adLockOptimistic
        
        If ASCII.recordCount = 0 Then
            MsgBox "Sorry, No record Found!", vbOKOnly + vbCritical, DialogBoxTitle
        Else
            lstASCIIList.listitems.Clear
            Do Until ASCII.EOF
            Set ASCIIListing = lstASCIIList.listitems.Add(, , ASCII!DateGenerated)
                ASCIIListing.SubItems(1) = ASCII!filename
                ASCIIListing.SubItems(2) = ASCII!MBMBordxDate
                ASCIIListing.SubItems(3) = ASCII!MBMBatchNo
                ASCIIListing.SubItems(4) = ASCII!PAYCode
                ASCIIListing.SubItems(5) = ASCII!NoofPrinciple
                ASCIIListing.SubItems(6) = ASCII!NoofSupplementary
                ASCII.MoveNext
            Loop
        End If
        Screen.MousePointer = vbDefault
    End If

End Sub

Private Sub CmdStop_Click()
    ' stop
    intExit = 1
    'cmdExit.Enabled = True
End Sub

Private Sub CmdUpload_Click()
Dim RecordSaved
    
    If txtFileVersion = "" And Mid(txtFileName, 1, 1) = "N" Then
        MsgBox "Please Enter File Version", vbOKOnly + vbCritical, DialogBoxTitle
        txtFileVersion.SetFocus
    ElseIf txtPAYCode = "" Then
        MsgBox "Please Enter Payor Code", vbOKOnly + vbCritical, DialogBoxTitle
        txtPAYCode.SetFocus
    ElseIf txtFileName = "" Then
        MsgBox "Please Enter FileName", vbOKOnly + vbCritical, DialogBoxTitle
        txtPAYCode.SetFocus
    ElseIf txtImportFilename = "" Then
        MsgBox "Please Enter Filename", vbOKOnly + vbCritical, DialogBoxTitle
        'txtImportFilename.SetFocus
    Else
        RecordSaved = MsgBox("Do you want to upload this file?", vbYesNo + vbExclamation)
        If RecordSaved = vbYes Then
            Screen.MousePointer = vbHourglass
            cmdUpload.Enabled = False
            cmdExit.Enabled = False

            If Mid(txtFileName, 1, 1) = "N" Then
                Call GenerateMember
                Call UpdateMBMPriting
                SSTab1.Enabled = True
            ElseIf Mid(txtFileName, 1, 1) = "E" Then
                Call GenerateENDtMember
                Call UpdateMBMPriting
                
                SSTab1.Enabled = True
                cmdExit.Enabled = True
                
            End If
        End If
    End If
    Screen.MousePointer = vbDefault
    cmdExit.Enabled = True
End Sub

Private Sub cmdVerify_Click()
    txtSubmissionDate = ""
    txtTotalPrincipal = 0
    txtTotalSupplementary = 0
    txtTotalError = 0
    cmdFileLocation.Enabled = False
    cmdUpload.Enabled = False
    If Mid(txtFileName, 1, 1) = "N" And txtFileVersion = "" Then
        MsgBox "Please Enter File Version", vbOKOnly + vbCritical, DialogBoxTitle
        txtFileName.SetFocus
    ElseIf txtPAYCode = "" Then
        MsgBox "Please Enter Payor Code", vbOKOnly + vbCritical, DialogBoxTitle
        txtPAYCode.SetFocus
    ElseIf txtFileName = "" Then
        MsgBox "Please Enter File Type", vbOKOnly + vbCritical, DialogBoxTitle
        txtFileName.SetFocus
    ElseIf txtImportFilename = "" Then
        MsgBox "Please Enter Filename", vbOKOnly + vbCritical, DialogBoxTitle
        'txtImportFilename.SetFocus
    Else
        
        cmdVerify.Enabled = False
        cmdUpload.Enabled = False
        lstERRList.listitems.Clear
        If Mid(txtFileName, 1, 1) = "N" Then
            Call VerifyNewMember
            If intExit = 0 Then
                MsgBox "Verifed Sucessfully...", vbOKOnly + vbInformation, DialogBoxTitle
                Screen.MousePointer = vbDefault
            Else
                Screen.MousePointer = vbDefault
            End If
        ElseIf Mid(txtFileName, 1, 1) = "E" Then
            Call VerifyENDtMember
            If intExit = 0 Then
                MsgBox "Verifed Sucessfully...", vbOKOnly + vbInformation, DialogBoxTitle
                Screen.MousePointer = vbDefault
            Else
                Screen.MousePointer = vbDefault
            End If
        End If
    End If
    cmdVerify.Enabled = True
    cmdUpload.Enabled = True
    cmdFileLocation.Enabled = True
    intExit = 0
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    Call EnterToTab(KeyAscii)
End Sub

Private Sub Form_Load()
    Call ComboListing
    
    intExit = 0
    txtPath = SRVPath & "ASCI"
    txtLocalFileLocation = "C:\"
    Call ComboList
End Sub

'---------------------Start ComboLisiting--------------------------
Private Sub ComboListing()
    Dim HLT As New ADODB.Recordset
    Dim PAY As New ADODB.Recordset
    Dim strsql As String
    
    HLT.Open "HLT", medixmasterconn, adOpenKeyset, adLockOptimistic
    
    Do Until HLT.EOF
       cboHLTCode1.AddItem HLT!HLTCode & " - " & HLT!HLTDescription
       cboHTL.AddItem HLT!HLTCode & " - " & HLT!HLTDescription
       HLT.MoveNext
    Loop
        
    strsql = "Select  PAYCode, PAYCompanyName From PAY where HLTCOde ='" & Mid(Trim(cboHLTCode1), 1, 1) & "'"
    PAY.Open strsql, medixmasterconn, adOpenKeyset, adLockOptimistic
    txtPAYCode.Clear
    Do Until PAY.EOF
       txtPAYCode.AddItem ChkStr(PAY!PAYCode) & " - " & ChkStr(PAY!PAYCompanyName)
       cboPAY.AddItem ChkStr(PAY!PAYCode) & " - " & ChkStr(PAY!PAYCompanyName)
       PAY.MoveNext
    Loop
    PAY.Close
    Set PAY = Nothing

    txtFileName.AddItem "N" & " - " & "New File"
    txtFileName.AddItem "E" & " - " & "Endorsement File"
    
    CboVerifyIC.AddItem "Y" & " - " & "Yes"
    CboVerifyIC.AddItem "N" & " - " & "No"
    CboVerifyIC.Text = "Y" & " - " & "Yes"
        
    txtFileVersion.AddItem "01" & " - " & "Version 520"
    txtFileVersion.AddItem "02" & " - " & "Version 790"
    txtFileVersion.AddItem "03" & " - " & "Version 4800"
    txtFileVersion.AddItem "04" & " - " & "Version 4800(MAA Gen & TH)"
    
    cboASCIIFileVer.AddItem "01" & " - " & "Version 185"
    cboASCIIFileVer.AddItem "02" & " - " & "Version 195"
    cboASCIIFileVer.AddItem "03" & " - " & "Version 300"

    
    cboSuppAnnualLimit.AddItem "C" & " - " & "Combined Limit"
    cboSuppAnnualLimit.AddItem "S" & " - " & "Separate Limit"
    
    cboLBLStatus.AddItem "Y" & " - " & "YES"
    cboLBLStatus.AddItem "N" & " - " & "NO"
    
    CboBack2Back.AddItem "Y" & " - " & "Yes"
    CboBack2Back.AddItem "N" & " - " & "No"
    CboBack2Back.Text = "N" & " - " & "No"
    
    HLT.Close
    Set HLT = Nothing
End Sub

Private Sub SSTab1_Click(PreviousTab As Integer)
Dim PAY As New ADODB.Recordset
Dim strsql As String
    
    If SSTab1.Tab = 0 Then
        cmdStop.Visible = True
        CmdPrint.Visible = True
    Else
        cmdStop.Visible = False
        CmdPrint.Visible = False
    End If
End Sub

'---------------------End ComboLisiting--------------------------
Private Sub txtFileName_Change()
    
    If Mid(txtFileName, 1, 1) = "N" Then
        txtFileVersion.Clear
        txtFileVersion.Enabled = True
        txtFileVersion.Clear
        txtFileVersion.AddItem "01" & " - " & "Version 520"
        txtFileVersion.AddItem "02" & " - " & "Version 790"
        txtFileVersion.AddItem "03" & " - " & "Version 4800"
        txtFileVersion.AddItem "04" & " - " & "Version 4800(MAA Gen & TH)"
        CboBack2Back.Enabled = True
    ElseIf Mid(txtFileName, 1, 1) = "E" Then
        txtFileVersion.Clear
        txtFileVersion.AddItem "01" & " - " & "Version 810"
        'txtFileVersion.AddItem "02" & " - " & "Version 3410"
        txtFileVersion.AddItem "02" & " - " & "Version 4830"
        txtFileVersion.AddItem "03" & " - " & "Version 4830(MAA Gen & TH)"
        CboBack2Back.Enabled = False
    End If

End Sub

Private Sub txtFileName_Click()
    If Mid(txtFileName, 1, 1) = "N" Then
        txtFileVersion.Enabled = True
        txtFileVersion.Clear
        txtFileVersion.AddItem "01" & " - " & "Version 520"
        txtFileVersion.AddItem "02" & " - " & "Version 790"
        txtFileVersion.AddItem "03" & " - " & "Version 4800"
        txtFileVersion.AddItem "04" & " - " & "Version 4800(MAA Gen & TH)"
        CboBack2Back.Enabled = True
    ElseIf Mid(txtFileName, 1, 1) = "E" Then
        txtFileVersion.Clear
        txtFileVersion.AddItem "01" & " - " & "Version 810"
   '     txtFileVersion.AddItem "02" & " - " & "Version 3410"
        txtFileVersion.AddItem "02" & " - " & "Version 4830"
        txtFileVersion.AddItem "03" & " - " & "Version 4830(MAA Gen & TH)"
        CboBack2Back.Enabled = False
    End If
End Sub

'------------------------Start selection of related Info----------------
Private Sub txtFileName_KeyPress(KeyAscii As Integer)
Dim NewText As String
Dim ValidCount As Integer
Dim ValidValue As String
Dim i As Integer
    If KeyAscii >= 32 And KeyAscii <> 127 Then
    NewText = LCase(Left(txtFileName.Text, txtFileName.SelStart) + Chr(KeyAscii) + Mid(txtFileName.Text, txtFileName.SelStart + txtFileName.SelLength + 1))
 
    ValidCount = 0
    ValidValue = ""
    
    For i = 0 To txtFileName.ListCount - 1
        If NewText = LCase(Left(txtFileName.List(i), Len(NewText))) Then
            ValidCount = ValidCount + 1
            ValidValue = txtFileName.List(i)
        End If
    Next
    If ValidCount <= 1 Then KeyAscii = 0
    
         If ValidCount = 1 Then
           txtFileName.Text = ValidValue
           txtFileName.SelStart = 0
           txtFileName.SelLength = Len(ValidValue)
         End If
    End If
End Sub

Private Sub txtFileName_LostFocus()
    If Mid(txtFileName, 1, 1) = "N" Then
        CboBack2Back.Enabled = True
    ElseIf Mid(txtFileName, 1, 1) = "E" Then
        CboBack2Back.Enabled = False
    End If
End Sub

Private Sub txtfileversion_KeyPress(KeyAscii As Integer)
Dim NewText As String
Dim ValidCount As Integer
Dim ValidValue As String
Dim i As Integer

If KeyAscii >= 32 And KeyAscii <> 127 Then
    NewText = LCase(Left(txtFileVersion.Text, txtFileVersion.SelStart) + Chr(KeyAscii) + Mid(txtFileVersion.Text, txtFileVersion.SelStart + txtFileVersion.SelLength + 1))
    
    ValidCount = 0
    ValidValue = ""
    
    For i = 0 To txtFileVersion.ListCount - 1
        If NewText = LCase(Left(txtFileVersion.List(i), Len(NewText))) Then
            ValidCount = ValidCount + 1
            ValidValue = txtFileVersion.List(i)
        End If
    Next
    If ValidCount <= 1 Then KeyAscii = 0
         If ValidCount = 1 Then
           txtFileVersion.Text = ValidValue
           txtFileVersion.SelStart = 0
           txtFileVersion.SelLength = Len(ValidValue)
         End If
    End If
End Sub
'---------------------End selection of related Info-----------------------------

'---------------------Start Verification of New File--------------------------
Private Sub VerifyNewMember()
Dim PayorCode As String
Dim CoveredIDCount As Integer
Dim CheckICValue As String
Dim ErrorCount As Integer
Dim TotPrincipal As Integer
Dim TotalPrincipalCount As Integer
Dim totalSuppCount As Integer
Dim TotSupplementary As Integer
Dim ListError As ListItem
Dim Msg As VerifyError
Dim Check As Boolean
Dim strsqlPLN As String
Dim PLN As New ADODB.Recordset
Dim TempPLN As String
Dim MBM As New ADODB.Recordset
Dim strsql As String
Dim CoverDays As Integer
Dim NewPLNCode As String
Dim VNewIC, VNewDOB, VNewSex, VNewPostCode, VNewState, VNewCity As String
Dim VNewEffDate, VNewExpDate, VNewPlanCode, VNewInsured As String
Dim VNewDataType, VRenewal, VNewPolicyNo, vNewName As String
Dim VNewCLnPlan, VNewCLnIns, VNewClnEffDate, VNewClnExpDate As String
Dim MBMICArray() As String
Dim MBMICCnt As Integer
Dim ICCount As Integer
'Dim DiscSql As String
Dim tmpPolDisc As Double
Dim tmpRenewDisc As Double
'Dim DiscPlan As New ADODB.Recordset

    lstERRList.listitems.Clear
    
    Erase MBMICArray
    ReDim MBMICArray(0)
    MBMICCnt = 0
       
    Screen.MousePointer = vbHourglass
    Open txtImportFilename For Input As #1   ' Open file.
    Do
    Do While Not EOF(1)   ' Loop until end of file.
       Line Input #1, TextLine   ' Read line into variable.
        recordCount = recordCount + 1
            If recordCount = 1 Then
                txtSubmissionDate = Format(Mid(TextLine, 4, 2) & "/" & Mid(TextLine, 6, 2) & "/" & Mid(TextLine, 8, 4), "dd/mm/yyyy")
                If Mid(TextLine, 1, 1) <> "F" Then
                    Call CreateErrorList(recordCount, "", "", "1st Record must be header with Data Type 'F'")
                    ErrorCount = ErrorCount + 1
                    Call ErrorListing
                End If
                If Trim(Mid(TextLine, 2, 2)) = "" Or Trim(Mid(TextLine, 2, 2)) <> Mid(txtPAYCode, 1, 2) Then
                    Call CreateErrorList(recordCount, "", "", "Invalid Insurer's Code " & Trim(Mid(TextLine, 2, 2)))
                    ErrorCount = ErrorCount + 1
                    Call ErrorListing
                End If
                If IsDate(txtSubmissionDate) = False Then
                    Call CreateErrorList(recordCount, "", "", "Invalid Submission Date !")
                    ErrorCount = ErrorCount + 1
                    Call ErrorListing
                End If
                VNewPayor = Mid(TextLine, 2, 2)
            End If
                             
            If recordCount >= 2 And Mid(TextLine, 1, 1) <> "T" Then
                VNewPolicyNo = Mid(TextLine, 2, 25)
                vNewName = Mid(TextLine, 47, 60)
                VNewDataType = Mid(TextLine, 1, 1)
                If Trim(Mid(txtFileVersion, 1, 2)) = "04" Then
                    VNewCity = Mid(TextLine, 197, 50)
                    VNewPostCode = Mid(TextLine, 247, 5)
                    VNewState = Mid(TextLine, 252, 15)
                    VNewIC = Mid(TextLine, 303, 20)
                    VNewDOB = Mid(TextLine, 323, 2) & "/" & Mid(TextLine, 325, 2) & "/" & Mid(TextLine, 327, 4)
                    VNewSex = Mid(TextLine, 331, 1)
                    VNewEffDate = Mid(TextLine, 404, 2) & "/" & Mid(TextLine, 406, 2) & "/" & Mid(TextLine, 408, 4)
                    VNewExpDate = Mid(TextLine, 412, 2) & "/" & Mid(TextLine, 414, 2) & "/" & Mid(TextLine, 416, 4)
                    VNewPlanCode = Mid(TextLine, 420, 4)
                    VNewInsured = Mid(TextLine, 424, 1)
                    VRenewal = Mid(TextLine, 465, 1)
                Else
                    VNewCity = Mid(TextLine, 197, 20)
                    VNewPostCode = Mid(TextLine, 217, 5)
                    VNewState = Mid(TextLine, 222, 15)
                    VNewIC = Mid(TextLine, 273, 20)
                    VNewDOB = Mid(TextLine, 293, 2) & "/" & Mid(TextLine, 295, 2) & "/" & Mid(TextLine, 297, 4)
                    VNewSex = Mid(TextLine, 301, 1)
                    VNewEffDate = Mid(TextLine, 374, 2) & "/" & Mid(TextLine, 376, 2) & "/" & Mid(TextLine, 378, 4)
                    VNewExpDate = Mid(TextLine, 382, 2) & "/" & Mid(TextLine, 384, 2) & "/" & Mid(TextLine, 386, 4)
                    
                    If Trim(Mid(txtFileVersion, 1, 2)) = "01" Then
                        VNewPlanCode = Mid(TextLine, 390, 2)
                        VNewInsured = Mid(TextLine, 392, 1)
                        VRenewal = Mid(TextLine, 433, 1)
                    Else
                        VNewPlanCode = Mid(TextLine, 390, 4)
                        VNewInsured = Mid(TextLine, 394, 1)
                        VRenewal = Mid(TextLine, 435, 1)
                    End If
                End If
                
                If Mid(TextLine, 1, 1) <> "P" Then 'P-Principal, H, W, S, D
                    If Mid(TextLine, 1, 1) = "" Then
                        Call CreateErrorList(recordCount, VNewPolicyNo, vNewName, "Invalid Data Type, should be either 'H','W','S' or 'D' only")
                        ErrorCount = ErrorCount + 1
                        Call ErrorListing
                    End If
                End If
                
                If Trim(VNewPolicyNo) = "" Then ' Policy No.
                    Call CreateErrorList(recordCount, VNewPolicyNo, vNewName, "Policy No Is Blank")
                    ErrorCount = ErrorCount + 1
                    Call ErrorListing
                End If
                
                If Trim(vNewName) = "" Then
                    Call CreateErrorList(recordCount, VNewPolicyNo, vNewName, "Name Is Blank")
                    ErrorCount = ErrorCount + 1
                    Call ErrorListing
                End If
                
                If Trim(Mid(TextLine, 107, 30)) = "" Then ' Address
                    If Trim(Mid(TextLine, 137, 30)) = "" Then
                        If Trim(Mid(TextLine, 167, 30)) = "" Then
                            If Mid(TextLine, 1, 1) = "P" Then
                                Call CreateErrorList(recordCount, VNewPolicyNo, vNewName, "Address Is Blank")
                                ErrorCount = ErrorCount + 1
                                Call ErrorListing
                            End If
                        End If
                    End If
                End If
                
                If Trim(VNewIC) = "" Then
                    Call CreateErrorList(recordCount, VNewPolicyNo, vNewName, "IC,BC or PP Is Blank")
                    ErrorCount = ErrorCount + 1
                    Call ErrorListing
                End If
                
                If Mid(CboVerifyIC, 1, 1) = "Y" Then
                    'CheckICValue = CheckIC(Trim(VNewIC))
                    If VNewIC <> "" Then
                        strsql = "Select MBMNumber, MBMICBCPP From MBM Where MBMICBCPP = '" & Trim(VNewIC) & "' And MBMStatus = 'A'"
                        MBM.Open strsql, medixmasterconn, adOpenKeyset, adLockOptimistic
    
                        If MBM.recordCount > 0 Then
                            'CheckIC = MBM!MBMNumber
                            Do Until MBM.EOF
                                Call CreateErrorList(recordCount, VNewPolicyNo, vNewName, "Same IC " & Trim(VNewIC) & " Existed in Membership No. :" & MBM!MBMNumber)
                                ErrorCount = ErrorCount + 1
                                Call ErrorListing
                            MBM.MoveNext
                            Loop
                        End If
                        MBM.Close
                        Set MBM = Nothing
                    End If
                End If
                    
                If IsDate(Trim(VNewDOB)) = False Then
                    Call CreateErrorList(recordCount, VNewPolicyNo, vNewName, "Invalid Date format ~ Date Of Birth")
                    ErrorCount = ErrorCount + 1
                    Call ErrorListing
                End If
                                
                If Trim(VNewSex) = "" Then
                    Call CreateErrorList(recordCount, VNewPolicyNo, vNewName, "Sex Is Blank")
                    ErrorCount = ErrorCount + 1
                    Call ErrorListing
                End If
                
                If Trim(VNewPlanCode) = "" And VNewDataType = "P" Then
                    Call CreateErrorList(recordCount, VNewPolicyNo, vNewName, "Invalid Plan Code")
                    ErrorCount = ErrorCount + 1
                    Call ErrorListing
                End If
                
                If IsDate(Trim(VNewEffDate)) = False And VNewDataType = "P" Then
                    Call CreateErrorList(recordCount, VNewPolicyNo, vNewName, "Invalid Date format ~ Effective Date")
                    ErrorCount = ErrorCount + 1
                    Call ErrorListing
                End If
                                    
                If IsDate(Trim(VNewExpDate)) = False And VNewDataType = "P" Then
                    Call CreateErrorList(recordCount, VNewPolicyNo, vNewName, "Invalid Date format ~ Expiry Date ")
                    ErrorCount = ErrorCount + 1
                    Call ErrorListing
                End If
                If IsDate(Trim(VNewEffDate)) And IsDate(Trim(VNewExpDate)) And VNewDataType = "P" Then
                    CoverDays = CDate(VNewExpDate) - CDate(VNewEffDate)
                        If CoverDays < 30 Or CoverDays > 365 Then
                        Call CreateErrorList(recordCount, VNewPolicyNo, vNewName, "Policy Coverage must be > 30 and < 365 days ")
                        ErrorCount = ErrorCount + 1
                        Call ErrorListing
                    End If
                End If
                    
                
                If Len(Trim(VNewPlanCode)) <> "" And VNewDataType = "P" Then
        ' --- New Plan verify eg. plan 01 convert to plan 21
                    strsqlPLN = "select * from PLN where PLNCode = '" & Trim(VNewPlanCode) & "'"
                    PLN.Open strsqlPLN, medixmasterconn, adOpenKeyset, adLockOptimistic
                    If PLN.recordCount = 0 Then
                        Call CreateErrorList(recordCount, VNewPolicyNo, vNewName, "Invalid Plan Code " & " - " & Trim(VNewPlanCode))
                        ErrorCount = ErrorCount + 1
                        Call ErrorListing
                    Else
                        
                        If Trim(PLN!NewPLNCode) <> "" Then
                            NewPLNCode = Trim(PLN!NewPLNCode)
                            PLN.Close
                            Set PLN = Nothing
                            strsqlPLN = "select * from PLN where PLNCode = '" & NewPLNCode & "' AND HLTCode='S'"
                            PLN.Open strsqlPLN, medixmasterconn, adOpenKeyset, adLockOptimistic

                            If CDate(VNewEffDate) < CDate(PLN!PLNEffDate) Then
                                Call CreateErrorList(recordCount, VNewPolicyNo, vNewName, "Invalid Plan Code or Effective Date !")
                                ErrorCount = ErrorCount + 1
                                Call ErrorListing
                            End If
                        End If
                    End If
                    PLN.Close
                    Set PLN = Nothing
        ' --- End of New Plan verify eg. plan 01 convert to plan 21

        ' ----- Policy & Renewal Discount Verification
'                    tmpPolDisc = IIf(IsNumeric(Mid(TextLine, 3732, 5)), Mid(TextLine, 3732, 5), 0)
'                    tmpRenewDisc = IIf(IsNumeric(Mid(TextLine, 3737, 5)), Mid(TextLine, 3737, 5), 0)
'                    If tmpPolDisc <> 0 Or tmpRenewDisc <> 0 Then
'                        DiscSql = "SELECT * FROM DiscPlan WHERE PLNCode='" & VNewPlanCode & "'"
'                        DiscPlan.Open DiscSql, medixmasterconnMaint, adOpenKeyset, adLockOptimistic
'                        If DiscPlan.recordCount Then
'                            If CDbl(DiscPlan!PolicyDisc) <> CDbl(tmpPolDisc) Then
'                                Call CreateErrorList(recordCount, VNewPolicyNo, vNewName, "Wrong Policy Discount Value!")
'                                ErrorCount = ErrorCount + 1
'                                Call ErrorListing
'                            End If
'                            If CDbl(DiscPlan!RenewalDisc) <> CDbl(tmpRenewDisc) Then
'                                Call CreateErrorList(recordCount, VNewPolicyNo, vNewName, "Wrong Renewal Discount Value !")
'                                ErrorCount = ErrorCount + 1
'                                Call ErrorListing
'                            End If
'                        End If
'                        DiscPlan.Close
'                        Set DiscPlan = Nothing
'                    End If
       ' ----- End of Policy & Renewal Discount Verification
                    
                End If
                
                If Trim(VNewInsured) = "" And VNewDataType = "P" Then
                    Call CreateErrorList(recordCount, VNewPolicyNo, vNewName, "Invalid Insured Type")
                    ErrorCount = ErrorCount + 1
                    Call ErrorListing
                End If
                
                If Trim(VRenewal) <> "" Then
                    If VRenewal <> "R" And VRenewal <> "N" And VRenewal <> "C" And VRenewal <> "Y" And VRenewal <> "Y" And VRenewal <> "T" And VRenewal <> "P" Then
                        Call CreateErrorList(recordCount, VNewPolicyNo, vNewName, "Invalid Renewal Code !")
                        ErrorCount = ErrorCount + 1
                        Call ErrorListing
                    End If
                End If
' ***------- Check for Member Topup
                If Len(Trim(VNewPlanCode)) <> "" And VNewDataType = "P" Then
                    strsqlPLN = "select * from PLN where PLNCode = '" & Trim(VNewPlanCode) & "'"
                    PLN.Open strsqlPLN, medixmasterconn, adOpenKeyset, adLockOptimistic
                    If PLN.recordCount Then
                        If PLN!PLNTopUpStatus = "Y" Then
                            strsql = "Select Top 1 MBMNumber, INSCode, MBMTopupNumber From MBM Where MBMICBCPP = '" & Trim(VNewIC) & _
                                    "' And MBMStatus = 'A' AND PayCode ='" & VNewPayor & "' AND INSCode='" & VNewInsured & _
                                    "' ORDER BY MBMPayorEffDate DESC, MBMNumber "
                            MBM.Open strsql, medixmasterconn, adOpenKeyset, adLockOptimistic

                            If MBM.recordCount <= 0 Then
                                Call CreateErrorList(recordCount, VNewPolicyNo, vNewName, "NO Member Record found for Topup ! ")
                                ErrorCount = ErrorCount + 1
                                Call ErrorListing
                            Else
                                If Trim(MBM!MBMTopupNumber) <> "" Or IsNull(MBM!MBMTopupNumber) = False Then
                                    Call CreateErrorList(recordCount, VNewPolicyNo, vNewName, "NO Record found for Topup ! ")
                                    ErrorCount = ErrorCount + 1
                                    Call ErrorListing
                                End If
                                If Trim(MBM!INSCode) <> Trim(VNewInsured) Then
                                    Call CreateErrorList(recordCount, VNewPolicyNo, vNewName, "Topup Member Insured Type Is not SAME! ")
                                    ErrorCount = ErrorCount + 1
                                    Call ErrorListing
                                End If
                            End If
                            MBM.Close
                            Set MBM = Nothing
                        End If
                    End If
                    PLN.Close
                    Set PLN = Nothing
                End If
' ***------ End of Checking for Topup
' ************* CIS Varification
                If Trim(Mid(txtFileVersion, 1, 2)) = "03" Or Trim(Mid(txtFileVersion, 1, 2)) = "04" Then
                    VNewCLnPlan = Mid(TextLine, 6708, 4)
                    VNewCLnIns = Mid(TextLine, 6712, 1)
                    VNewClnEffDate = Mid(TextLine, 6713, 2) & "/" & Mid(TextLine, 6715, 2) & "/" & Mid(TextLine, 6717, 4)
                    VNewClnExpDate = Mid(TextLine, 6721, 2) & "/" & Mid(TextLine, 6723, 2) & "/" & Mid(TextLine, 6725, 4)
                    
                    If Trim(VNewCLnPlan) = "" And VNewDataType = "P" Then
                        Call CreateErrorList(recordCount, VNewPolicyNo, vNewName, "Clinical Plan Is NULL ! ")
                        ErrorCount = ErrorCount + 1
                        Call ErrorListing
                    End If
                    If Trim(VNewCLnPlan) <> "" Then
                        If Trim(VNewCLnIns) = "" And VNewDataType = "P" Then
                            Call CreateErrorList(recordCount, VNewPolicyNo, vNewName, "Clinical Insured Type Is NULL! ")
                            ErrorCount = ErrorCount + 1
                            Call ErrorListing
                        End If
                        If IsDate(Trim(VNewClnEffDate)) = False And VNewDataType = "P" Then
                            Call CreateErrorList(recordCount, VNewPolicyNo, vNewName, "Invalid Date Format ~ Clinical Eff Date ! ")
                            ErrorCount = ErrorCount + 1
                            Call ErrorListing
                        End If
                        If IsDate(Trim(VNewClnExpDate)) = False And VNewDataType = "P" Then
                            Call CreateErrorList(recordCount, VNewPolicyNo, vNewName, "Invalid Date Format ~ Clinical Exp Date ! ")
                            ErrorCount = ErrorCount + 1
                            Call ErrorListing
                        End If
                    End If
                End If
' ************* End of CIS verify
' ************* Chk 4 Duplicate Rec in ascii file
                If UBound(MBMICArray) > 0 Then
                For ICCount = 0 To UBound(MBMICArray)
                    If Trim(MBMICArray(ICCount)) = Trim(VNewIC) Then
                        Call CreateErrorList(recordCount, VNewPolicyNo, vNewName, "Same IC Number Found ! ")
                        ErrorCount = ErrorCount + 1
                        Call ErrorListing
                    End If
                Next ICCount
                End If
                MBMICCnt = MBMICCnt + 1
                ReDim Preserve MBMICArray(MBMICCnt)
                MBMICArray(MBMICCnt) = VNewIC
' ************* End of Duplicate Rec. checking
            End If
            If Mid(TextLine, 1, 1) = "P" Then
                TotalPrincipalCount = TotalPrincipalCount + 1
                txtTotalPrincipal = TotalPrincipalCount
            ElseIf Mid(TextLine, 1, 1) <> "P" And Mid(TextLine, 1, 1) <> "T" And Mid(TextLine, 1, 1) <> "F" Then
                totalSuppCount = totalSuppCount + 1
                txtTotalSupplementary = totalSuppCount
            End If
            
             
            DoEvents
                 

            If intExit = 1 Then
                Check = False
                Exit Do
            End If
        Loop
        Loop Until Check = False
    Close #1   ' Close file.
    txtTotalError = ErrorCount
    recordCount = 0
    
End Sub
'---------------------End Verification of New File--------------------------

'---------------------Start Verification of ENDt File-----------------------

Private Sub VerifyENDtMember()
    Dim PayorCode As String
    Dim CoveredIDCount As Integer
    Dim CheckICValue As String
    Dim ErrorCount As Integer
    Dim TotPrincipal As Integer
    Dim TotSupplementary As Integer
    Dim ListError As ListItem
    Dim Msg As VerifyError
    Dim strsql As String
    Dim VEndtIC, VEndtDOB, VEndtSuppID As String
    Dim VEndtPostCode, VEndtState, VEndtCity As String
    Dim VEndtEffDate, VEndtExpDate As String
    Dim VEndtPlanCode, VEndtInsured As String
    Dim VEndtDataType, VEndtPolicyNo, vEndtName As String
    Dim VEndtCLnPlan, VEndtCLnIns, VEndtClnEffDate, VEndtClnExpDate As String

    lstERRList.listitems.Clear
    
    Screen.MousePointer = vbHourglass
    ERR.Open "ERR", medixmasterconn, adOpenKeyset, adLockOptimistic
    Open txtImportFilename For Input As #1   ' Open file.
    Do While Not EOF(1)   ' Loop until end of file.
       Line Input #1, TextLine   ' Read line into variable.
        recordCount = recordCount + 1
            If recordCount = 1 Then
                txtSubmissionDate = Mid(TextLine, 4, 2) & "/" & Mid(TextLine, 6, 2) & "/" & Mid(TextLine, 8, 4)
                If Mid(TextLine, 1, 1) <> "E" Then
                    Call CreateErrorList(recordCount, "", "", "1st Record must be header with Data Type 'F'")
                    ErrorCount = ErrorCount + 1
                    Call ErrorListing
                End If
                If Trim(Mid(TextLine, 2, 2)) = "" Or Trim(Mid(TextLine, 2, 2)) <> Mid(txtPAYCode, 1, 2) Then
                    Call CreateErrorList(recordCount, "", "", "Invalid Insurer's Code " & Trim(Mid(TextLine, 2, 2)))
                    ErrorCount = ErrorCount + 1
                    Call ErrorListing
                End If
                If IsDate(txtSubmissionDate) = False Then
                    Call CreateErrorList(recordCount, "", "", "Invalid Date" & Trim(Mid(TextLine, 2, 2)))
                    ErrorCount = ErrorCount + 1
                    Call ErrorListing
                End If
            End If
                             
            If recordCount >= 2 Then
                If Mid(TextLine, 10, 1) = "P" Then
                    TotPrincipal = TotPrincipal + 1
                ElseIf Mid(TextLine, 10, 1) <> "P" And Mid(TextLine, 1, 1) <> "T" Then
                    TotSupplementary = TotSupplementary + 1
                End If
            End If
            
            If recordCount >= 2 And Mid(TextLine, 1, 1) <> "T" Then
                VEndtPolicyNo = Mid(TextLine, 29, 25)
                vEndtName = Mid(TextLine, 74, 60)
                VEndtDataType = Mid(TextLine, 10, 1)
'                If Trim(Mid(TextLine, 2, 8)) = "" Then
                If IsDate(Mid(TextLine, 2, 2) & "/" & Mid(TextLine, 4, 2) & "/" & Mid(TextLine, 6, 4)) = False Then
                    Call CreateErrorList(recordCount, VEndtPolicyNo, vEndtName, "Invalid date format ~ Endorsement Effective Date")
                    ErrorCount = ErrorCount + 1
                    Call ErrorListing
                End If
                
                If Trim(VEndtDataType) = "" Then
                    Call CreateErrorList(recordCount, VEndtPolicyNo, vEndtName, "Data Type is blank")
                    ErrorCount = ErrorCount + 1
                    Call ErrorListing
                End If
                
                If Trim(Mid(TextLine, 11, 18)) = "" Then
                    Call CreateErrorList(recordCount, VEndtPolicyNo, vEndtName, "Membership No Is Blank")
                    ErrorCount = ErrorCount + 1
                    Call ErrorListing
                End If
          
                If Trim(VEndtPolicyNo) = "" Then
                    Call CreateErrorList(recordCount, VEndtPolicyNo, vEndtName, "Policy No Is Blank")
                    ErrorCount = ErrorCount + 1
                    Call ErrorListing
                End If
                
                If Trim(vEndtName) = "" Then
                    Call CreateErrorList(recordCount, VEndtPolicyNo, vEndtName, "Name Is Blank")
                    ErrorCount = ErrorCount + 1
                    Call ErrorListing
                End If
                If Trim(Mid(TextLine, 1, 1)) <> "C" Then ' No verification for cancellation
                    If Trim(Mid(TextLine, 134, 30)) = "" Or Trim(Mid(TextLine, 164, 30)) = "" Or Trim(Mid(TextLine, 194, 30)) = "" And Mid(TextLine, 10, 1) = "P" Then
                        Call CreateErrorList(recordCount, VEndtPolicyNo, vEndtName, "Address Is Blank")
                        ErrorCount = ErrorCount + 1
                        Call ErrorListing
                    End If
                    
                    If Mid(txtFileVersion, 1, 2) = "03" Then
                        VEndtCity = Mid(TextLine, 224, 50)
                        VEndtPostCode = Mid(TextLine, 274, 5)
                        VEndtState = Mid(TextLine, 279, 15)
                        VEndtIC = Mid(TextLine, 330, 20)
                        VEndtDOB = Mid(TextLine, 350, 2) & "/" & Mid(TextLine, 352, 2) & "/" & Mid(TextLine, 354, 4)
                        VEndtEffDate = Mid(TextLine, 431, 2) & "/" & Mid(TextLine, 433, 2) & "/" & Mid(TextLine, 435, 4)
                        VEndtExpDate = Mid(TextLine, 439, 2) & "/" & Mid(TextLine, 441, 2) & "/" & Mid(TextLine, 443, 4)
                        VEndtPlanCode = Mid(TextLine, 447, 4)
                        VEndtInsured = Mid(TextLine, 451, 1)
                        
                    Else
                        VEndtCity = Mid(TextLine, 224, 20)
                        VEndtPostCode = Mid(TextLine, 244, 5)
                        VEndtState = Mid(TextLine, 249, 15)
                        VEndtIC = Mid(TextLine, 300, 20)
                        VEndtDOB = Mid(TextLine, 320, 2) & "/" & Mid(TextLine, 322, 2) & "/" & Mid(TextLine, 324, 4)
                        VEndtEffDate = Mid(TextLine, 401, 2) & "/" & Mid(TextLine, 403, 2) & "/" & Mid(TextLine, 405, 4)
                        VEndtExpDate = Mid(TextLine, 409, 2) & "/" & Mid(TextLine, 411, 2) & "/" & Mid(TextLine, 413, 4)
                        VEndtPlanCode = Mid(TextLine, 417, 4)
                        VEndtInsured = Mid(TextLine, 421, 1)
                    End If
                    If Mid(TextLine, 10, 1) <> "P" And Mid(TextLine, 1, 1) = "A" Then
                       If Mid(txtFileVersion, 1, 2) = "02" Or Mid(txtFileVersion, 1, 2) = "03" Then
                          If Mid(txtFileVersion, 1, 2) = "02" Then
                             VEndtSuppID = Trim(Mid(TextLine, 2963, 2))
                          Else
                             VEndtSuppID = Trim(Mid(TextLine, 2993, 2))
                          End If
                          If VEndtSuppID = "" Then
                             Call CreateErrorList(recordCount, VEndtPolicyNo, vEndtName, "Cover ID is Blank ")
                             ErrorCount = ErrorCount + 1
                             Call ErrorListing
                         End If
                       End If
                    End If
                    
                    If Mid(CboVerifyIC, 1, 1) = "Y" Then
                        'CheckICValue = CheckIC(VEndtIC)
                        If Trim(VEndtIC) <> "" Then
                            strsql = "Select MBMNumber, MBMICBCPP From MBM Where MBMICBCPP = '" & Trim(VEndtIC) & "' And MBMStatus = 'A'"
                            MBM.Open strsql, medixmasterconn, adOpenKeyset, adLockOptimistic
        
                            If MBM.recordCount > 0 Then
                                'CheckIC = MBM!MBMNumber
                                Do Until MBM.EOF
                                    Call CreateErrorList(recordCount, VEndtPolicyNo, vEndtName, "Same IC " & Trim(VEndtIC) & " Existed in Membership No. :" & MBM!MBMNumber)
                                    ErrorCount = ErrorCount + 1
                                    Call ErrorListing
                                    MBM.MoveNext
                                Loop
                            End If
                            MBM.Close
                            Set MBM = Nothing
                        End If
                    End If
                    
                    If Trim(VEndtCity) = "" And Mid(TextLine, 10, 1) = "P" Then
                        Call CreateErrorList(recordCount, VEndtPolicyNo, vEndtName, "City Is Blank")
                        ErrorCount = ErrorCount + 1
                        Call ErrorListing
                    End If
                    
                    If Trim(VEndtPostCode) = "" And Mid(TextLine, 10, 1) = "P" Then
                        Call CreateErrorList(recordCount, VEndtPolicyNo, vEndtName, "PostCode Is Blank")
                        ErrorCount = ErrorCount + 1
                        Call ErrorListing
                    End If
        
                    If Trim(VEndtState) = "" And Mid(TextLine, 10, 1) = "P" Then
                        Call CreateErrorList(recordCount, VEndtPolicyNo, vEndtName, "State Code Is Blank")
                        ErrorCount = ErrorCount + 1
                        Call ErrorListing
                    End If
    
                    If Trim(VEndtIC) = "" Then
                        Call CreateErrorList(recordCount, VEndtPolicyNo, vEndtName, "IC/BC/PP Is Blank")
                        ErrorCount = ErrorCount + 1
                        Call ErrorListing
                    End If
                    
                    If IsDate(Trim(VEndtDOB)) = False Then
                        Call CreateErrorList(recordCount, VEndtPolicyNo, vEndtName, "Invalid Date format ~ DOB ")
                        ErrorCount = ErrorCount + 1
                        Call ErrorListing
                    End If
                    If Trim(Mid(TextLine, 10, 1)) = "P" And IsDate(Trim(VEndtEffDate)) = False Then
                        Call CreateErrorList(recordCount, VEndtPolicyNo, vEndtName, "Invalid Date format ~ Effective Date")
                        ErrorCount = ErrorCount + 1
                        Call ErrorListing
                    End If
                    If Trim(Mid(TextLine, 10, 1)) = "P" And IsDate(Trim(VEndtExpDate)) = False Then
                        Call CreateErrorList(recordCount, VEndtPolicyNo, vEndtName, "Invalid Date format ~ Expiry Date")
                        ErrorCount = ErrorCount + 1
                        Call ErrorListing
                    End If
                    If Trim(VEndtInsured) = "" And Mid(TextLine, 10, 1) = "P" Then
                        Call CreateErrorList(recordCount, VEndtPolicyNo, vEndtName, "Invalid Insured Type")
                        ErrorCount = ErrorCount + 1
                        Call ErrorListing
                    End If
                        
                    If Trim(VEndtPlanCode) = "" And Mid(TextLine, 10, 1) = "P" Then
                        Call CreateErrorList(recordCount, VEndtPolicyNo, vEndtName, "Invalid Plan Code")
                        ErrorCount = ErrorCount + 1
                        Call ErrorListing
                    End If
                End If
' ************* CIS Verification
'loh                If Trim(Mid(txtFileVersion, 1, 2)) = "02" Or Trim(Mid(txtFileVersion, 1, 2)) = "03" Then
'loh                    VEndtCLnPlan = Mid(TextLine, 6737, 4)
'loh                    VEndtCLnIns = Mid(TextLine, 6741, 1)
'loh                    VEndtClnEffDate = Mid(TextLine, 6742, 2) & "/" & Mid(TextLine, 6744, 2) & "/" & Mid(TextLine, 6746, 4)
'loh                    VEndtClnExpDate = Mid(TextLine, 6750, 2) & "/" & Mid(TextLine, 6752, 2) & "/" & Mid(TextLine, 6754, 4)
'loh                    If Trim(VEndtCLnPlan) = "" And VendtDataType = "P" Then
'loh                        Call CreateErrorList(recordCount, VEndtPolicyNo, vEndtName, "Clinical Plan Is NULL ! ")
'loh                       ErrorCount = ErrorCount + 1
'loh                        Call ErrorListing
'loh                    End If
'loh                    If Trim(VEndtCLnIns) = "" And VendtDataType = "P" Then
'loh                        Call CreateErrorList(recordCount, VEndtPolicyNo, vEndtName, "Clinical Insured Type Is NULL! ")
'loh                        ErrorCount = ErrorCount + 1
'loh                        Call ErrorListing
'loh                    End If
'loh                    If IsDate(VEndtClnEffDate) = False And VendtDataType = "P" Then
'loh                        Call CreateErrorList(recordCount, VEndtPolicyNo, vEndtName, "Invalid Date Format ~ Clinical Eff Date ! ")
'loh                        ErrorCount = ErrorCount + 1
'loh                        Call ErrorListing
'loh                    End If
'loh                    If IsDate(VEndtClnExpDate) = False And VendtDataType = "P" Then
'loh                        Call CreateErrorList(recordCount, VEndtPolicyNo, vEndtName, "Invalid Date Format ~ Clinical Exp Date ! ")
'loh                        ErrorCount = ErrorCount + 1
'loh                        Call ErrorListing
'loh                    End If
'loh                End If
' ************* End of CIS verify
            End If
        DoEvents
    Loop
    If Mid(TextLine, 1, 1) <> "T" Then
        Call CreateErrorList(recordCount, "", "", "1st Record must be trailer with Data Type 'T'")
        ErrorCount = ErrorCount + 1
    ElseIf Trim(Mid(TextLine, 2, 10)) <> TotPrincipal Then
        Call CreateErrorList(recordCount, "", "", "Total Principal Record Not Match Trailer Total Principal")
    ElseIf Trim(Mid(TextLine, 12, 10)) <> TotSupplementary Then
        Call CreateErrorList(recordCount, "", "", "Total Supplementary Record Not Match Trailer Total Supplementary")
    End If
                            
    Close #1   ' Close file.

    ERR.Close
    txtTotalError = ErrorCount
    txtTotalPrincipal = TotPrincipal
    txtTotalSupplementary = TotSupplementary
    recordCount = 0
    
End Sub
'**************************End Verification of ENDt File*********************************

Private Sub CreateErrorList(RecordLine, PolicyNo, Name, Msg)
        
    Message = Msg
    RECLine = RecordLine
        
End Sub

Private Sub ErrorListing()
Dim ListError As ListItem
txtTempRecordLine = recordCount
txtTempERRMessage = Message

    Set ListError = lstERRList.listitems.Add(, , txtTempRecordLine)
    If Mid(txtFileName, 1, 1) = "N" Then
        ListError.SubItems(1) = Trim(Mid(TextLine, 2, 25))
        ListError.SubItems(2) = Trim(Mid(TextLine, 47, 60))
    ElseIf Mid(txtFileName, 1, 1) = "E" Then
        ListError.SubItems(1) = Trim(Mid(TextLine, 29, 25))
        ListError.SubItems(2) = Trim(Mid(TextLine, 74, 60))
    End If
    
    ListError.SubItems(3) = Trim(txtTempERRMessage)
    ListError.SubItems(4) = Trim(txtImportFilename)
        
    txtTempERRMessage = ""
    Message = ""
End Sub

Private Sub GenerateMember()
'On Error Resume Next
Dim strsql As String
Dim PayorCode As String
Dim recordCount As Integer
Dim uploadSupp As Integer
Dim uploadprin As Integer

    If txtTotalSupplementary <> "" Then
        uploadSupp = txtTotalSupplementary
    End If
    If txtTotalPrincipal <> "" Then
        uploadprin = txtTotalPrincipal
    End If
    
    If vbYes Then
                       
        Open txtImportFilename For Input As #1   ' Open file.
        Line Input #1, TextLine   ' Read line into variable.
        If Mid(TextLine, 2, 2) <> Mid(txtPAYCode, 1, 2) Then
            MsgBox "Payor Selected Is Not Match The File Payor", vbOKOnly + vbCritical, DialogBoxTitle
            txtPAYCode.SetFocus
        Else
            If Mid(TextLine, 1, 1) = "F" Then
                PayorCode = Mid(TextLine, 2, 2)
                txtSubmissionDate = Mid(TextLine, 4, 2) & "/" & Mid(TextLine, 6, 2) & "/" & Mid(TextLine, 8, 4)
            End If
            
            recordCount = 1
            MBM.Open "MBM", medixmasterconn, adOpenKeyset, adLockOptimistic
            MBMCover.Open "MBMCoveredPersons", medixmasterconn, adOpenKeyset, adLockOptimistic
            MBMOthers.Open "MBMOthers", medixmasterconn, adOpenKeyset, adLockOptimistic
            CLNOthers.Open "MBMCLNOthers", medixmasterconn, adOpenKeyset, adLockOptimistic

            txtBatchNo = CheckBatch(Mid(txtPAYCode, 1, 2), txtSubmissionDate, Mid(txtFileName, 1, 1))
            Do While Not EOF(1)   ' Loop until end of file.
               Line Input #1, TextLine   ' Read line into variable.
                recordCount = recordCount + 1
                If Mid(TextLine, 1, 1) = "P" And Mid(TextLine, 1, 1) <> "T" Then
                    CoveredID = "00"
                    Call PrincipalDataImport(Mid(txtPAYCode, 1, 2), CoveredID & 0)
                    Call OthersDataImport
                    If Trim(Mid(TextLine, 6708, 4)) <> "" Then
                        Call CISDataImport
                    End If
                ElseIf Mid(TextLine, 1, 1) <> "P" And Mid(TextLine, 1, 1) <> "T" Then
                    Call CoverPersonDataImport(Mid(txtPAYCode, 1, 2), CoveredID)
                End If
                DoEvents
            Loop
            MsgBox "Record Import Sucessfully...", vbOKOnly + vbInformation, DialogBoxTitle
            MBM.Close
            MBMOthers.Close
            MBMCover.Close
            CLNOthers.Close
            Set MBM = Nothing
            Set MBMOthers = Nothing
            Set MBMCover = Nothing
            Set CLNOthers = Nothing
        End If
    End If
    Close #1   ' Close file.

End Sub

Private Sub PrincipalDataImport(Payor As String, CoveredIDCount As Integer)
'On Error Resume Next
'On Error GoTo error1
Dim MemberDOB As String
Dim FindAge As String
Dim AgeValue As String
Dim AnnualLimitValue As String
Dim GetProString As String
Dim tmpPreMemberNo As String
Dim ListMember As ListItem
Dim INewCityCode, INewPostCode, INewSTACode As String
Dim INewTelNoH, INewTelNoO, INewTelNoM, INewIC, INewDOB, INewSex As String
Dim INewBRNCode As String
Dim INewTakeOver, INewRenewal, INewAppName As String
Dim TMPMBM As New ADODB.Recordset
Dim tmpPlan As New ADODB.Recordset
Dim TmpSql, TmpSql2 As String
Dim Check As Boolean
Dim HSClnInd As String
Dim TempPolDis As Double
Dim TempRenewDis As Double
'Dim DiscPlan As New ADODB.Recordset
    
    Do
        MBM.AddNew
        With MBM
            If Mid(TextLine, 1, 1) = "P" And Mid(TextLine, 1, 1) <> "T" Then
                If Mid(cboSuppAnnualLimit, 1, 1) = "C" Then
                    !MBMMemberType = "P"
                Else
                    !MBMMemberType = "Q"
                End If
                !MBMPolicyNo = Mid(TextLine, 2, 25)
                !MBMCOVID = "00"
                !MBMSalutation = Mid(TextLine, 42, 5)
                !MBMName = Mid(TextLine, 47, 60)
                !MBMAddress1 = Mid(TextLine, 107, 30)
                !MBMAddress2 = Mid(TextLine, 137, 30)
                !MBMAddress3 = Mid(TextLine, 167, 30)
                !MBMValueDate = Date
                !MBMDataStatus = "A"
                !MBMExpiredStatus = "N"
                !PAYCode = Mid(txtPAYCode, 1, 2)
                !HLTCode = Mid(cboHLTCode1, 1, 1)
                If Len(Trim(Mid(TextLine, 3624, 60))) Then
                    !MBMAppName = Mid(TextLine, 3624, 60)
                End If
            If Len(Trim(Mid(TextLine, 3742, 20))) Then
                !MBMIcBcPp2nd = Trim(Mid(TextLine, 3742, 20))
            End If

                If Mid(txtFileVersion, 1, 2) <> "04" Then
                    INewCityCode = Mid(TextLine, 197, 20)
                    INewPostCode = Mid(TextLine, 217, 5)
                    INewSTACode = Mid(TextLine, 222, 15)
                    INewTelNoH = Mid(TextLine, 237, 12)
                    INewTelNoM = Mid(TextLine, 249, 12)
                    INewTelNoO = Mid(TextLine, 261, 12)
                    INewIC = Mid(TextLine, 273, 20)
                    INewSex = Mid(TextLine, 301, 1)
                    '!NATCode = Mid(TextLine, 302, 10)
                    PayEffDate = Mid(TextLine, 374, 2) & "/" & Mid(TextLine, 376, 2) & "/" & Mid(TextLine, 378, 4)
                    PayorEXPDate = Mid(TextLine, 382, 2) & "/" & Mid(TextLine, 384, 2) & "/" & Mid(TextLine, 386, 4)
                    MemberDOB = Mid(TextLine, 293, 2) & "/" & Mid(TextLine, 295, 2) & "/" & Mid(TextLine, 297, 4)
              
                    PlanCode = Trim(Mid(TextLine, 390, 4))
                    InsuredCode = Mid(TextLine, 394, 1)
                    INewTakeOver = Mid(TextLine, 435, 1)
                    INewRenewal = Mid(TextLine, 435, 1)
                    
                    If Mid(txtFileVersion, 1, 2) = "01" Then
                        PlanCode = Trim(Mid(TextLine, 390, 2))
                        InsuredCode = Mid(TextLine, 392, 1)
                        INewTakeOver = Mid(TextLine, 433, 1)
                        INewRenewal = Mid(TextLine, 433, 1)
                        INewBRNCode = Mid(TextLine, 497, 15)
                    End If
                    If Mid(txtFileVersion, 1, 2) = "02" Then
                        INewBRNCode = Mid(TextLine, 499, 15)
                    ElseIf Mid(txtFileVersion, 1, 2) = "03" Then
                        INewBRNCode = Mid(TextLine, 2939, 15)
                    End If
                Else
                    INewCityCode = Mid(TextLine, 197, 50)
                    INewPostCode = Mid(TextLine, 247, 5)
                    INewSTACode = Mid(TextLine, 252, 15)
                    INewTelNoH = Mid(TextLine, 267, 12)
                    INewTelNoM = Mid(TextLine, 279, 12)
                    INewTelNoO = Mid(TextLine, 291, 12)
                    INewIC = Mid(TextLine, 303, 20)
                    INewSex = Mid(TextLine, 331, 1)
                    '!NATCode = Mid(TextLine, 302, 10)
                    PlanCode = Trim(Mid(TextLine, 420, 4))
                    InsuredCode = Mid(TextLine, 424, 1)
                    INewTakeOver = Mid(TextLine, 465, 1)
                    INewRenewal = Mid(TextLine, 465, 1)
                    INewBRNCode = Mid(TextLine, 2969, 15)
                    PayEffDate = Mid(TextLine, 404, 2) & "/" & Mid(TextLine, 406, 2) & "/" & Mid(TextLine, 408, 4)
                    PayorEXPDate = Mid(TextLine, 412, 2) & "/" & Mid(TextLine, 414, 2) & "/" & Mid(TextLine, 416, 4)
                    MemberDOB = Mid(TextLine, 323, 2) & "/" & Mid(TextLine, 325, 2) & "/" & Mid(TextLine, 327, 4)
                   
                End If
                
                !CTYCode = INewCityCode
                !MBMPostCode = INewPostCode
                !STACode = INewSTACode
                !MBMTelnoH = INewTelNoH
                !MBMTelNoM = INewTelNoM
                !MBMTelNoO = INewTelNoO
                !MBMIcBcPp = INewIC
                !MBMDOB = MemberDOB
                !MBMSex = INewSex
                '!NATCode = Mid(TextLine, 302, 10)
                !MBMPayorEffDate = PayEffDate
                !MBMPayorEXPDate = PayorEXPDate
                
                If INewTakeOver = "C" Or INewTakeOver = "Y" Then
                    !MBMTakeOver = INewTakeOver
                    !MBMRenewFlag = "N"
                Else
                    !MBMRenewFlag = INewRenewal
                End If
                
                !BRNCode = INewBRNCode
                
                !MBMBordxDate = txtSubmissionDate
                !MBMStatus = "A"
                !MBMBatchNo = txtBatchNo
                !SRVCodeList = "EA"
        '------------------------------------------------------------------------------------------------------------------------
                If Trim(PlanCode) <> "" Then 'HIS prem, mco, annLmt, Access Calc.
                    !PLNCode = PlanCode
                    !INSCode = InsuredCode
                    FindAge = GetAge(CDate(PayEffDate), CDate(MemberDOB))
                    AgeValue = Mid(GetAgeBand(Mid(cboHLTCode1, 1, 1), FindAge), 1, 2)
                    !AgeCode = AgeValue
                    If AgeValue = "01" Then
                        !MBMAdultChild = "C"
                    Else
                        !MBMAdultChild = "A"
                    End If
                    
                    txtMBMPremium = Format(GetPremium(InsuredCode, Payor, AgeValue, PlanCode, Mid(cboHLTCode1, 1, 1), CDate(PayEffDate)), "#,##0.00")
                    BasicPrem = Format(Round(txtMBMPremium, 0), "#,##0.00")
                    txtMBMPremium = InsertPremium(CDate(PayorEXPDate), CDate(PayEffDate), Format(txtMBMPremium, "#,##0.00"))

                    TempPolDis = IIf(IsNumeric(Mid(TextLine, 3732, 5)), Mid(TextLine, 3732, 5), 0)
                    TempRenewDis = IIf(IsNumeric(Mid(TextLine, 3737, 5)), Mid(TextLine, 3737, 5), 0)
                
'                    TmpSql = "select * from DiscPlan where PLNCode = '" & Trim(PlanCode) & "'"
'                    DiscPlan.Open TmpSql, medixmasterconnMaint, adOpenKeyset, adLockOptimistic
'                    If DiscPlan.recordCount Then
                        If TempPolDis > 0 Then
                            txtMBMPremium = txtMBMPremium * (100 - TempPolDis) / 100
                        End If
                        If TempRenewDis > 0 Then
                            txtMBMPremium = txtMBMPremium * (100 - TempRenewDis) / 100
                        End If
'                    End If
'                    DiscPlan.Close
'                    Set DiscPlan = Nothing
                    txtMBMPremium = Format(Round(txtMBMPremium, 0), "#,##0.00")
                    
                    If AgeValue = "01" Then
                        txtMBMMCOFees = Format(GetMCO(InsuredCode, Mid(cboHLTCode1, 1, 1), "C", CDate(PayEffDate), PlanCode), "#,##0.00")
                    Else
                        txtMBMMCOFees = Format(GetMCO(InsuredCode, Mid(cboHLTCode1, 1, 1), "A", CDate(PayEffDate), PlanCode), "#,##0.00")
                    End If
                    BasicMCO = txtMBMMCOFees
                    txtMBMMCOFees = InsertMCOFees(CDate(PayorEXPDate), CDate(PayEffDate), Format(txtMBMMCOFees, "#,##0.00"))
                    txtMBMMCOFees = Round(txtMBMMCOFees)
                    AnnualLimitValue = GetAnnualLimit(InsuredCode, PlanCode, Mid(cboHLTCode1, 1, 1), CDate(PayEffDate))
                    !MBMAvailableLimit = AnnualLimitValue
                    tempAnnLmt = AnnualLimitValue
                End If
                CisPlnCode = Trim(Mid(TextLine, 6708, 4))
                CisInsType = Mid(TextLine, 6712, 1)
                CisEffDate = Mid(TextLine, 6713, 2) & "/" & Mid(TextLine, 6715, 2) & "/" & Mid(TextLine, 6717, 4)
                CisExpDate = Mid(TextLine, 6721, 2) & "/" & Mid(TextLine, 6723, 2) & "/" & Mid(TextLine, 6725, 4)
                
                If Trim(PlanCode) = "" And Trim(CisPlnCode) <> "" Then
                    tmpPreMemberNo = Payor & CisPlnCode & CisInsType & GenerateMembershipNo(Payor, CisPlnCode, CisInsType)
                Else
                    tmpPreMemberNo = Payor & PlanCode & InsuredCode & GenerateMembershipNo(Payor, PlanCode, InsuredCode)
                End If
                
                If Trim(PlanCode) = "" Then
                    HSClnInd = "C"
                    !MBMClinical = "Y"
                    !MBMCLNStatus = "A"
                    !MBMCLNDataStatus = "A"
                ElseIf Trim(CisPlnCode) <> "" Then
                    HSClnInd = "M"
                    !MBMClinical = "Y"
                    !MBMCLNStatus = "A"
                    !MBMCLNDataStatus = "A"
                Else
                    HSClnInd = "H"
                End If
                !HSClnInd = HSClnInd
                !MBMNumber = tmpPreMemberNo
                txtTempMemberNo = tmpPreMemberNo
                !MBMRegUser = Logon
                !MBMLastUpdateUser = Logon
 
' ---------- write Previous MBM no. to new MBMTopupMember field
                TmpSql = "select * from PLN where PLNCode = '" & Trim(PlanCode) & "'"
                tmpPlan.Open TmpSql, medixmasterconn, adOpenKeyset, adLockOptimistic
                If tmpPlan.recordCount Then
                    If tmpPlan!PLNTopUpStatus = "Y" Then
                        TmpSql2 = "Select TOP 1 MBMNumber, MBMTopupNumber, MBMTopupIND From MBM Where MBMICBCPP = '" & Trim(INewIC) & _
                                    "' And MBMStatus = 'A' AND PayCode ='" & Mid(txtPAYCode, 1, 2) & "' AND INSCode='" & Trim(InsuredCode) & _
                                    "' ORDER BY MBMPayorEffDate DESC, MBMNumber "
                        TMPMBM.Open TmpSql2, medixmasterconn, adOpenKeyset, adLockOptimistic
 
                        If TMPMBM.recordCount Then
                           !MBMTopupNumber = TMPMBM!MBMNumber
                           !MBMTopupInd = "Y"
                        End If
                        TMPMBM.Close
                        Set TMPMBM = Nothing
                    End If
                End If
                tmpPlan.Close
                Set tmpPlan = Nothing
' ---------- topup Member info
' ----- Policy year calculation
                'TmpSql2 = "Select MBMNumber From MBM Where MBMICBCPP = '" & Trim(INewIC) & _
                '            "' And MBMStatus = 'A' AND PayCode ='" & Mid(txtPAYCode, 1, 2) & "'"
                'TMPMBM.Open TmpSql2, medixmasterconn, adOpenKeyset, adLockOptimistic'

                'If TMPMBM.recordCount Then
                '    !MBMPolicyYear = TMPMBM.recordCount + 1
                'Else
                '    !MBMPolicyYear = 1
                'End If
                'TMPMBM.Close
                'Set TMPMBM = Nothing
' ----- Policy year calculation
            End If
        End With
        MBM.Update

        If intExit = 1 Then
            Check = False
            Exit Do
        End If
        
        Loop Until Check = False

    If Mid(txtFileVersion, 1, 2) = "04" Then
        PayEffDate = Mid(TextLine, 404, 2) & "/" & Mid(TextLine, 406, 2) & "/" & Mid(TextLine, 408, 4)
    Else
        PayEffDate = Mid(TextLine, 374, 2) & "/" & Mid(TextLine, 376, 2) & "/" & Mid(TextLine, 378, 4)
    End If
End Sub

Private Sub CoverPersonDataImport(Payor As String, CoveredIDCount As Integer)
Dim MemberDOB As String
Dim FindAge As String
Dim AgeValue As String
Dim GetProString As String
Dim tmpPreMemberNo As String
Dim ListMember As ListItem
Dim MBMCOVDOB As String
Dim StrSqlAnnLmt As String
Dim AnnLmt As New ADODB.Recordset
Dim tempSAnnLmt As String
Dim tempSAvailLmt As String
Dim tempSPremium As String
Dim ClnAnnLmtSql As String
Dim CLnAnnLmt As New ADODB.Recordset
Dim CLNPlnSql As String
Dim CLNPln As New ADODB.Recordset
Dim CISCovSuppLMtStat As String
Dim CISSuppAnnLmt, CISSuppVisLmt As String
Dim CISSuppBPrem, CISSuppPrem As String
Dim CISSuppBMCO, CISSuppMCO As String
Dim TEMPGWP As String
Dim TmpClmNonPayfr As String
Dim TmpClmNonPayTo As String
Dim TempPolDis As Double
Dim TempRenewDis As Double
'Dim DiscPlan As New ADODB.Recordset
Dim TmpSql As String

        StrSqlAnnLmt = "Select SuppLimitStatus, PLNCode from AnnualLimit where PLNCode = '" & Trim(PlanCode) & "'"
        AnnLmt.Open StrSqlAnnLmt, medixmasterconn, adOpenKeyset, adLockOptimistic

       ' medixmasterconn.BeginTrans
        CoveredIDCount = CoveredIDCount + 1
        MBMCover.AddNew
        With MBMCover
            If Mid(TextLine, 1, 1) <> "P" And Mid(TextLine, 1, 1) <> "T" Then
                !MBMCRELCode = Mid(TextLine, 1, 1)
                If CoveredIDCount < 10 Then
                    !MBMCCoverID = 0 & CoveredIDCount
                Else
                    !MBMCCoverID = CoveredIDCount
                End If
                !MBMCMemberType = "C"
            End If
            !MBMCSalutation = Mid(TextLine, 42, 5)
            !MBMCName = Trim(Mid(TextLine, 47, 60))
            If Len(Trim(Mid(TextLine, 3684, 12))) Then
                !MBMCPayorPremNet = IIf(IsNumeric(Trim(Mid(TextLine, 3684, 12))), Trim(Mid(TextLine, 3684, 12)), 0)
            End If
            If Len(Trim(Mid(TextLine, 3696, 12))) Then
                !MBMCPayorPremGross = IIf(IsNumeric(Trim(Mid(TextLine, 3696, 12))), Trim(Mid(TextLine, 3696, 12)), 0)
            End If
            If Len(Trim(Mid(TextLine, 3708, 12))) Then
                !MBMCPayorMCONet = IIf(IsNumeric(Trim(Mid(TextLine, 3708, 12))), Trim(Mid(TextLine, 3708, 12)), 0)
            End If
            If Len(Trim(Mid(TextLine, 3720, 12))) Then
                !MBMCPayorMCOGross = IIf(IsNumeric(Trim(Mid(TextLine, 3720, 12))), Trim(Mid(TextLine, 3720, 12)), 0)
            End If
            If Len(Trim(Mid(TextLine, 3732, 5))) Then
                !MBMCPolicyDisc = IIf(IsNumeric(Trim(Mid(TextLine, 3732, 5))), Trim(Mid(TextLine, 3732, 5)), 0)
            End If
            If Len(Trim(Mid(TextLine, 3737, 5))) Then
                !MBMCRenewalDisc = IIf(IsNumeric(Trim(Mid(TextLine, 3737, 5))), Trim(Mid(TextLine, 3737, 5)), 0)
            End If
            If Len(Trim(Mid(TextLine, 3742, 20))) Then
                !MBMCIcBcPp2nd = Trim(Mid(TextLine, 3742, 20))
            End If

            If Mid(txtFileVersion, 1, 2) = "04" Then
                !MBMCIcBcPpNo = Trim(Mid(TextLine, 303, 20))
                MBMCOVDOB = Trim(Mid(TextLine, 323, 2) & "/" & Mid(TextLine, 325, 2) & "/" & Mid(TextLine, 327, 4))
                !MBMCSex = Mid(TextLine, 331, 1)
            Else
                !MBMCIcBcPpNo = Trim(Mid(TextLine, 273, 20))
                MBMCOVDOB = Trim(Mid(TextLine, 293, 2) & "/" & Mid(TextLine, 295, 2) & "/" & Mid(TextLine, 297, 4))
                !MBMCSex = Mid(TextLine, 301, 1)
            End If
            If Len(MBMCOVDOB) Then
                !MBMCDOB = Trim(MBMCOVDOB)
            End If
            
'------------------------------------------------------------------------------------------------------------
            If Mid(txtFileVersion, 1, 2) = "01" Then
                !MBMCExclusion = Mid(TextLine, 434, 60)
'                !MBMCOccupation = Mid(TextLine, 320, 30)
                !MBMCWorkNature = Mid(TextLine, 350, 1)
                !MBMCBloodType = Mid(TextLine, 351, 3)
                !MBMCAllergic = Mid(TextLine, 354, 20)
                !MBMCSmoking = Mid(TextLine, 494, 1)
                !MBMCAlcohol = Mid(TextLine, 495, 1)
                
            ElseIf Mid(txtFileVersion, 1, 2) = "02" Then
                !MBMCExclusion = Mid(TextLine, 436, 60)
'                !MBMCOccupation = Mid(TextLine, 320, 30)
                !MBMCWorkNature = Mid(TextLine, 350, 1)
                !MBMCBloodType = Mid(TextLine, 351, 3)
                !MBMCAllergic = Mid(TextLine, 354, 20)
                !MBMCSmoking = Mid(TextLine, 496, 1)
                !MBMCAlcohol = Mid(TextLine, 497, 1)
                
            ElseIf Mid(txtFileVersion, 1, 2) = "03" Then
                !MBMCExclusion = Mid(TextLine, 436, 2500)
'                !MBMCOccupation = Mid(TextLine, 320, 30)
                !MBMCWorkNature = Mid(TextLine, 350, 1)
                !MBMCBloodType = Mid(TextLine, 351, 3)
                !MBMCAllergic = Mid(TextLine, 354, 20)
                !MBMCSmoking = Mid(TextLine, 2936, 1)
                !MBMCAlcohol = Mid(TextLine, 2937, 1)
                
            ElseIf Mid(txtFileVersion, 1, 2) = "04" Then
                !MBMCExclusion = Mid(TextLine, 466, 2500)
'                !MBMCOccupation = Mid(TextLine, 350, 30)
                !MBMCWorkNature = Mid(TextLine, 380, 1)
                !MBMCBloodType = Mid(TextLine, 381, 3)
                !MBMCAllergic = Mid(TextLine, 384, 20)
                !MBMCSmoking = Mid(TextLine, 2966, 1)
                !MBMCAlcohol = Mid(TextLine, 2967, 1)
                
                If Len(Trim(Mid(TextLine, 3050, 1))) Then
                    !MBMCStaffDec = Trim(Mid(TextLine, 3050, 1))
                End If
            
                If Len(Trim(Mid(TextLine, 3051, 5))) Then
                    !MBMCGenWP = Mid(TextLine, 3051, 5)
                    !MBMCGenEndWP = Format(DateAdd("d", Trim(Mid(TextLine, 3051, 5)), PayEffDate), "dd/mm/yyyy")
                End If
            
                If Len(Trim(Mid(TextLine, 3056, 5))) Then
                    !MBMCPreExWP = Mid(TextLine, 3056, 5)
                    !MBMCPreExEndWP = Format(DateAdd("d", Trim(Mid(TextLine, 3056, 5)), PayEffDate), "dd/mm/yyyy")
                End If
            
                If Len(Trim(Mid(TextLine, 3061, 5))) Then
                    !MBMCSpeIllWP = Mid(TextLine, 3061, 5)
'                    !MBMCSpeIllEndWP = Format(DateAdd("d", Trim(Mid(TextLine, 3061, 5)), PayEffDate), "dd/mm/yyyy")
                    !MBMSpeIllEndWP = Format(DateAdd("d", Trim(Mid(TextLine, 3061, 5)), PayEffDate), "dd/mm/yyyy")
                End If
            End If
            
                If Len(Mid(TextLine, 3413, 8)) Then
                    !MBMCPAYNo = Mid(TextLine, 3413, 8)
                End If
                If Len(Mid(TextLine, 3421, 2)) Then
                    !MBMCPAYSeqNo = Mid(TextLine, 3421, 2)
                End If
                If Len(Mid(TextLine, 3423, 8)) Then
                    !MBMCClientNo = Mid(TextLine, 3423, 8)
                End If
                If Len(Mid(TextLine, 3431, 2)) Then
                    !MBMCClientID = Mid(TextLine, 3431, 2)
                End If
                
            'Cater for different Plan/Benefit
                If IsNumeric(Mid(TextLine, 3448, 8)) Then
                    !MBMCRBAmt = Mid(TextLine, 3448, 8)
                End If
                If Len(Mid(TextLine, 3433, 15)) Then
                    !MBMCPrevPLNCode = Mid(TextLine, 3433, 15)
                End If
                If Len(Mid(TextLine, 3456, 60)) Then
                    !MBMCPrevInsCom = Mid(TextLine, 3456, 60)
                End If
                If Len(Mid(TextLine, 3516, 60)) Then
                    !MBMCPAYRemarks = Mid(TextLine, 3516, 60)
                End If
            'End If
            !MBMCStatus = "A"
'------------------------------------------------------------------------------------------------------------------------
            !MBMCBordxDate = txtSubmissionDate
            !MBMCBatchNo = txtBatchNo
            
            If Mid(txtFileVersion, 1, 2) = "04" Then
                MemberDOB = Mid(TextLine, 323, 2) & "/" & Mid(TextLine, 325, 2) & "/" & Mid(TextLine, 327, 4)
            Else
                MemberDOB = Mid(TextLine, 293, 2) & "/" & Mid(TextLine, 295, 2) & "/" & Mid(TextLine, 297, 4)
            End If
            FindAge = GetAge(CDate(PayEffDate), CDate(MemberDOB))
            AgeValue = Mid(GetAgeBand(Mid(cboHLTCode1, 1, 1), FindAge), 1, 2)
            !MBMCAgeband = AgeValue
            If AgeValue = "01" Then
                !MBMCAdultChild = "C"
            Else
                !MBMCAdultChild = "A"
            End If
            !MBMCNumber = txtTempMemberNo
            TempPolDis = IIf(IsNumeric(Mid(TextLine, 3732, 5)), Mid(TextLine, 3732, 5), 0)
            TempRenewDis = IIf(IsNumeric(Mid(TextLine, 3737, 5)), Mid(TextLine, 3737, 5), 0)
            
            If Mid(cboSuppAnnualLimit, 1, 1) = "S" Then
'                TmpSql = "select * from DiscPlan where PLNCode = '" & Trim(PlanCode) & "'"
'                DiscPlan.Open TmpSql, medixmasterconnMaint, adOpenKeyset, adLockOptimistic
                
                If AnnLmt!SuppLimitStatus = "B" Then
                    If CoveredIDCount = "1" Then
                        tempSPremium = Format(GetSPremium(InsuredCode, Payor, AgeValue, PlanCode, Mid(cboHLTCode1, 1, 1), CDate(PayEffDate)), "#,##0.00")
                        tempSPremium = InsertPremium(CDate(PayorEXPDate), CDate(PayEffDate), Format(tempSPremium, "#,##0.00"))
                        tempSAnnLmt = Format(GetSuppLimit(InsuredCode, PlanCode, Mid(cboHLTCode1, 1, 1), CDate(PayEffDate)), "#,##0.00")
                        !MBMCAnnualLimit = tempSAnnLmt
                        !MBMCAvailableLimit = tempSAnnLmt
                        !MBMCBasicPremium = Format(Round(tempSPremium, 0), "#,##0.00")
'                        If DiscPlan.recordCount Then
                            If TempPolDis > 0 Then
                                tempSPremium = tempSPremium * (100 - TempPolDis) / 100
                            End If
                            If TempRenewDis > 0 Then
                                tempSPremium = tempSPremium * (100 - TempRenewDis) / 100
                            End If
'                        End If
                        tempSPremium = Format(Round(tempSPremium, 0), "#,##0.00")
                        !MBMCPremium = tempSPremium
                    End If
                ElseIf AnnLmt!SuppLimitStatus = "C" Then 'Cater for SuppLimitStatus is "C"

                    tempSPremium = Format(GetSPremium(InsuredCode, Payor, AgeValue, PlanCode, Mid(cboHLTCode1, 1, 1), CDate(PayEffDate)), "#,##0.00")
                    tempSPremium = InsertPremium(CDate(PayorEXPDate), CDate(PayEffDate), Format(tempSPremium, "#,##0.00"))

                    tempSAnnLmt = Format(GetSuppLimit(InsuredCode, PlanCode, Mid(cboHLTCode1, 1, 1), CDate(PayEffDate)), "#,##0.00")
                    !MBMCAnnualLimit = tempSAnnLmt
                    !MBMCAvailableLimit = tempSAnnLmt
                    !MBMCBasicPremium = Format(Round(tempSPremium, 0), "#,##0.00")
'                    If DiscPlan.recordCount Then
                        If TempPolDis > 0 Then
                            tempSPremium = tempSPremium * (100 - TempPolDis) / 100
                        End If
                        If TempRenewDis > 0 Then
                            tempSPremium = tempSPremium * (100 - TempRenewDis) / 100
                        End If
'                    End If
                    tempSPremium = Format(Round(tempSPremium, 0), "#,##0.00")
                    !MBMCPremium = tempSPremium
                    TmpClmNonPayfr = Mid(TextLine, 3576, 2) & "/" & Mid(TextLine, 3578, 2) & "/" & Mid(TextLine, 3580, 4)
                    If IsDate(TmpClmNonPayfr) Then
                        !MBMCCLMNonPayFr = TmpClmNonPayfr
                    End If
                    TmpClmNonPayTo = Mid(TextLine, 3584, 2) & "/" & Mid(TextLine, 3586, 2) & "/" & Mid(TextLine, 3588, 4)
                    If IsDate(TmpClmNonPayTo) Then
                        !MBMCCLMNonPayTo = TmpClmNonPayTo
                    End If
                Else
                    !MBMCAnnualLimit = tempAnnLmt
                    !MBMCAvailableLimit = tempAnnLmt
                End If
'                DiscPlan.Close
'                Set DiscPlan = Nothing
            End If
' ******* Clinical Supp Limit
            CLNPlnSql = "SELECT AnnLmtIND, VisLmtIND from PLN where PLNCode='" & Trim(CisPlnCode) & _
                        "' AND HLTCode='" & Mid(cboHLTCode1, 1, 1) & "'"
            CLNPln.Open CLNPlnSql, MediConnCISDB, adOpenKeyset, adLockOptimistic
            
            CISCovSuppLMtStat = ""
            If CLNPln.recordCount Then
                If Trim(CLNPln!ANNLmtInd) = "Y" Then
                    ClnAnnLmtSql = "Select SuppLimitStatus from AnnualLimit where PLNCode = '" & Trim(CisPlnCode) & _
                                   "' AND HLTCode='" & Mid(cboHLTCode1, 1, 1) & "' AND INSCode='" & Trim(CisInsType) & "'"
                    CLnAnnLmt.Open ClnAnnLmtSql, MediConnCISDB, adOpenKeyset, adLockOptimistic
                    CISCovSuppLMtStat = CLnAnnLmt!SuppLimitStatus
                    CLnAnnLmt.Close
                    Set CLnAnnLmt = Nothing
                End If
            End If
            CLNPln.Close
            Set CLNPln = Nothing
            If Trim(CISCovSuppLMtStat) = "B" Then
                If CoveredIDCount = "1" Then
                    CISSuppPrem = Format(GetCLNSPremium(CisInsType, Payor, AgeValue, CisPlnCode, Mid(cboHLTCode1, 1, 1), CDate(CisEffDate)), "#,##0.00")
                    CISSuppBPrem = CISSuppPrem
                    CISSuppPrem = InsertPremium(CDate(CisExpDate), CDate(CisEffDate), Format(CISSuppPrem, "#,##0.00"))
                    CISSuppPrem = Format(Round(CISSuppPrem, 0), "#,##0.00")
                    
                    If AgeValue = "01" Then
                        CISSuppMCO = Format(GetCLNSMCO(CisInsType, Mid(cboHLTCode1, 1, 1), "C", CDate(CisEffDate), CisPlnCode), "#,##0.00")
                    Else
                        CISSuppMCO = Format(GetCLNSMCO(CisInsType, Mid(cboHLTCode1, 1, 1), "A", CDate(CisEffDate), CisPlnCode), "#,##0.00")
                    End If
                    CISSuppBMCO = CISSuppMCO
                    CISSuppMCO = InsertMCOFees(CDate(CisExpDate), CDate(CisEffDate), Format(CISSuppMCO, "#,##0.00"))
                    CISSuppMCO = Round(CISSuppMCO)

                    CISSuppAnnLmt = Format(GetCLNSuppLimit(CisInsType, CisPlnCode, Mid(cboHLTCode1, 1, 1), CDate(CisEffDate)), "#,##0.00")
                    CISSuppVisLmt = GetCLNSVisitLimit(CisInsType, CisPlnCode, Mid(cboHLTCode1, 1, 1), CDate(CisEffDate))
                    !MBMCCLNMCO = CISSuppMCO
                    !MBMCCLNBasicMCO = CISSuppBMCO
                    !MBMCCLNPrem = CISSuppPrem
                    !MBMCCLNBasicPrem = CISSuppBPrem
                    !MBMCCLNBalLmt = CISSuppAnnLmt
                    !MBMCCLNVisitLmt = CISSuppAnnLmt
                End If
            ElseIf Trim(CISCovSuppLMtStat) = "C" Then  'Cater for SuppLimitStatus is "C"
                    CISSuppPrem = Format(GetCLNSPremium(CisInsType, Payor, AgeValue, CisPlnCode, Mid(cboHLTCode1, 1, 1), CDate(CisEffDate)), "#,##0.00")
                    CISSuppBPrem = CISSuppPrem
                    CISSuppPrem = InsertPremium(CDate(CisExpDate), CDate(CisEffDate), Format(CISSuppPrem, "#,##0.00"))
                    CISSuppPrem = Format(Round(CISSuppPrem, 0), "#,##0.00")
                    
                    If AgeValue = "01" Then
                        CISSuppMCO = Format(GetCLNSMCO(CisInsType, Mid(cboHLTCode1, 1, 1), "C", CDate(CisEffDate), CisPlnCode), "#,##0.00")
                    Else
                        CISSuppMCO = Format(GetCLNSMCO(CisInsType, Mid(cboHLTCode1, 1, 1), "A", CDate(CisEffDate), CisPlnCode), "#,##0.00")
                    End If
                    CISSuppBMCO = CISSuppMCO
                    CISSuppMCO = InsertMCOFees(CDate(CisExpDate), CDate(CisEffDate), Format(CISSuppMCO, "#,##0.00"))
                    CISSuppMCO = Round(CISSuppMCO)

                    CISSuppAnnLmt = Format(GetCLNSuppLimit(CisInsType, CisPlnCode, Mid(cboHLTCode1, 1, 1), CDate(CisEffDate)), "#,##0.00")
                    CISSuppVisLmt = GetCLNSVisitLimit(CisInsType, CisPlnCode, Mid(cboHLTCode1, 1, 1), CDate(CisEffDate))
                    !MBMCCLNMCO = CISSuppMCO
                    !MBMCCLNBasicMCO = CISSuppBMCO
                    !MBMCCLNPrem = CISSuppPrem
                    !MBMCCLNBasicPrem = CISSuppBPrem
                    !MBMCCLNBalLmt = CISSuppAnnLmt
                    !MBMCCLNVisitLmt = CISSuppAnnLmt
            End If
' ****** End of Clinical Supp Limit
    End With
    MBMCover.Update
End Sub

Private Sub OthersDataImport()
On Error Resume Next
'On Error GoTo error1

Dim tmpPreMemberNo As String
Dim TempMBM As New ADODB.Recordset
Dim CheckICValue, AA, QQ As String
Dim TMP1stJoinedDate, TMPIC, TempSql As String
Dim TmpClmNonPayfr As String
Dim TmpClmNonPayTo As String

       ' medixmasterconn.BeginTrans
        
        MBMOthers.AddNew
        With MBMOthers
            !MBMNumber = txtTempMemberNo
            !MBMAgentCode = Trim(Mid(TextLine, 27, 15))
            
            If Mid(cboLBLStatus, 1, 1) = "Y" Then
                !LabelStatus = Mid(cboLBLStatus, 1, 1)
            End If
            If Mid(CboBack2Back, 1, 1) = "Y" Then
                !MBMBTBG = Mid(CboBack2Back, 1, 1)
            End If
            If Len(Trim(Mid(TextLine, 3684, 12))) Then
                !MBMPayorPremNet = IIf(IsNumeric(Trim(Mid(TextLine, 3684, 12))), Trim(Mid(TextLine, 3684, 12)), 0)
            End If
            If Len(Trim(Mid(TextLine, 3696, 12))) Then
                !MBMPayorPremGross = IIf(IsNumeric(Trim(Mid(TextLine, 3696, 12))), Trim(Mid(TextLine, 3696, 12)), 0)
            End If
            If Len(Trim(Mid(TextLine, 3708, 12))) Then
                !MBMPayorMCONet = IIf(IsNumeric(Trim(Mid(TextLine, 3708, 12))), Trim(Mid(TextLine, 3708, 12)), 0)
            End If
            If Len(Trim(Mid(TextLine, 3720, 12))) Then
                !MBMPayorMCOGross = IIf(IsNumeric(Trim(Mid(TextLine, 3720, 12))), Trim(Mid(TextLine, 3720, 12)), 0)
            End If
            If Len(Trim(Mid(TextLine, 3732, 5))) Then
                !MBMPolicyDisc = IIf(IsNumeric(Trim(Mid(TextLine, 3732, 5))), Trim(Mid(TextLine, 3732, 5)), 0)
            End If
            If Len(Trim(Mid(TextLine, 3737, 5))) Then
                !MBMRenewalDisc = IIf(IsNumeric(Trim(Mid(TextLine, 3737, 5))), Trim(Mid(TextLine, 3737, 5)), 0)
            End If

            If Mid(txtFileVersion, 1, 2) = "04" Then
                !MBMWeight = Trim(Mid(TextLine, 342, 3))
                !MBMHeight = Trim(Mid(TextLine, 345, 3))
                !MBMWorkNature = Mid(TextLine, 380, 1)
                !MBMBlood = Mid(TextLine, 381, 3)
                !MBMAllergic = Mid(TextLine, 384, 20)
                TMPIC = Mid(TextLine, 303, 20)
                TMP1stJoinedDate = Mid(TextLine, 404, 2) & "/" & Mid(TextLine, 406, 2) & "/" & Mid(TextLine, 408, 4)
            Else
                !MBMWeight = Trim(Mid(TextLine, 312, 3))
                !MBMHeight = Trim(Mid(TextLine, 315, 3))
                !MBMWorkNature = Mid(TextLine, 350, 1)
                !MBMBlood = Mid(TextLine, 351, 3)
                !MBMAllergic = Mid(TextLine, 354, 20)
                TMPIC = Mid(TextLine, 273, 20)
                TMP1stJoinedDate = Mid(TextLine, 374, 2) & "/" & Mid(TextLine, 376, 2) & "/" & Mid(TextLine, 378, 4)
            End If
            !MBMFirstMEMNo = txtTempMemberNo
            !MBMAccessFee = txtAccessFees
            !MBMPremium = txtMBMPremium
            !MBMBasicPrem = BasicPrem
            !MBMBasicMCO = BasicMCO
            !MBMBasicEA = BasicEA
            If txtDiscount <> "" Then
                !MBMDiscount = txtDiscount
                '!MBMDISStatus = True
                !MBMMCOFee = CDbl(txtMBMMCOFees) - CDbl(txtDiscount)
            Else
                !MBMMCOFee = txtMBMMCOFees
            End If
            !RELCODE = "P"
            '!MBMFirstJoined = Mid(TextLine, 572, 2) & "/" & Mid(TextLine, 574, 2) & "/" & Mid(TextLine, 576, 4)
        
        If Mid(txtFileVersion, 1, 2) = "01" Then
            If Len(Mid(TextLine, 393, 40)) Then
                !MBMGroupCompany = Mid(TextLine, 393, 40)
            End If
            If Len(Mid(TextLine, 434, 60)) Then
                !MBMExclusion = Mid(TextLine, 434, 60)
            End If
            If Len(Mid(TextLine, 494, 1)) Then
                !MBMSmoking = Mid(TextLine, 494, 1)
            End If
            If Len(Mid(TextLine, 495, 1)) Then
                !MBMAlcohol = Mid(TextLine, 495, 1)
            End If
            If Len(Mid(TextLine, 496, 1)) Then
                !MBMMaritalStatus = Mid(TextLine, 496, 1)
            End If
        
        ElseIf Mid(txtFileVersion, 1, 2) = "02" Then
            !MBMGroupCompany = Mid(TextLine, 395, 40)
            !MBMExclusion = Mid(TextLine, 436, 60)
            !MBMSmoking = Mid(TextLine, 496, 1)
            !MBMAlcohol = Mid(TextLine, 497, 1)
            !MBMMaritalStatus = Mid(TextLine, 498, 1)
            !MBMPrevPolNo = Mid(TextLine, 514, 25)
            If Len(Mid(TextLine, 539, 15)) Then
                !MBMPrevMemNo = Mid(TextLine, 539, 18)
            End If
            !MBMEmployeeNo = Mid(TextLine, 557, 15)
        ElseIf Mid(txtFileVersion, 1, 2) = "03" Then
            !MBMGroupCompany = Mid(TextLine, 395, 40)
            !MBMExclusion = Mid(TextLine, 436, 2500)
            !MBMSmoking = Mid(TextLine, 2936, 1)
            !MBMAlcohol = Mid(TextLine, 2937, 1)
            !MBMMaritalStatus = Mid(TextLine, 2938, 1)
            !MBMPrevPolNo = Mid(TextLine, 2954, 25)
            If Len(Mid(TextLine, 2979, 18)) Then
                !MBMPrevMemNo = Mid(TextLine, 2979, 18)
            End If
            !MBMEmployeeNo = Mid(TextLine, 2997, 15)
            If Len(Mid(TextLine, 3020, 1)) Then
                !MBMStaffDec = Trim(Mid(TextLine, 3020, 1))
            End If
        ElseIf Mid(txtFileVersion, 1, 2) = "04" Then
            !MBMGroupCompany = Mid(TextLine, 425, 40)
            !MBMExclusion = Mid(TextLine, 466, 2500)
            !MBMSmoking = Mid(TextLine, 2966, 1)
            !MBMAlcohol = Mid(TextLine, 2967, 1)
            !MBMMaritalStatus = Mid(TextLine, 2968, 1)
            !MBMPrevPolNo = Mid(TextLine, 2984, 25)
            If Len(Mid(TextLine, 3009, 18)) Then
                !MBMPrevMemNo = Mid(TextLine, 3009, 18)
            End If
            !MBMEmployeeNo = Mid(TextLine, 3027, 15)

            If Len(Mid(TextLine, 3050, 1)) Then
                !MBMStaffDec = Trim(Mid(TextLine, 3050, 1))
            End If
            
            If Len(Mid(TextLine, 3051, 5)) Then
                !MBMGenWP = Mid(TextLine, 3051, 5)
                '!MBMGenEndWP = CDate(Trim(Mid(TextLine, 3021, 5))) + CDbl(PayEffDate)
                !MBMGenEndWP = Format(DateAdd("d", Trim(Mid(TextLine, 3051, 5)), PayEffDate), "dd/mm/yyyy")
                '!MBMGenEndWP = !MBMGenWP + CDate(PayEffDate)
            End If
            
            If Len(Mid(TextLine, 3056, 5)) Then
                !MBMPreExWP = Mid(TextLine, 3056, 5)
                !MBMPreExEndWP = Format(DateAdd("d", Trim(Mid(TextLine, 3056, 5)), PayEffDate), "dd/mm/yyyy")
            End If
            
            If Len(Mid(TextLine, 3061, 5)) Then
                !MBMSpeIllWP = Mid(TextLine, 3061, 5)
                !MBMSpeIllEndWP = Format(DateAdd("d", Trim(Mid(TextLine, 3061, 5)), PayEffDate), "dd/mm/yyyy")
            End If
        End If
'DAte JOined Asigning
        
        'TempSql = "SELECT TOP 2 MBMFirstJoinedDate, MBMNumber FROM VIEW_MBM_MBMOthers where MBMIcBcPp='" & Trim(TMPIC) & "' AND PAYCode= '" & VNewPayor & _
        '        "' AND MBMStatus ='A' and MBMDataStatus='A' ORDER BY MBMPayorEffDate"
        'TempMBM.Open TempSql, medixmasterconn, adOpenKeyset, adLockOptimistic
        'AA = 0
        'If TempMBM.recordCount > 1 Then
        '    Do Until TempMBM.EOF
        '    QQ = TempMBM!MBMNumber
        '        AA = AA + 1
        '        If AA = 1 Then
        '            !MBMFirstJoinedDate = TempMBM!MBMFirstJoinedDate
        '        End If
        '    TempMBM.MoveNext
        '    Loop
        'Else
        '    !MBMFirstJoinedDate = TMP1stJoinedDate
        'End If
        'TempMBM.Close
        'Set TempMBM = Nothing
' DAte JOined
'------
        If Mid(txtFileVersion, 1, 2) = "03" Or Mid(txtFileVersion, 1, 2) = "04" Then
            !MBMDepartment = Mid(TextLine, 3221, 30)
            !MBMGRPCompanyAdd1 = Mid(TextLine, 3251, 30)
            !MBMGRPCompanyAdd2 = Mid(TextLine, 3281, 30)
            !MBMGRPCompanyAdd3 = Mid(TextLine, 3311, 30)
            !MBMGRPCompanyAdd4 = Mid(TextLine, 3341, 30)
            !MBMInstallment = Mid(TextLine, 3375, 1)
            !MBMPolSubNo1 = Mid(TextLine, 3376, 10)
            !MBMPOLSubNo2 = Mid(TextLine, 3386, 10)
            If Len(Mid(TextLine, 3396, 8)) Then
                !MBMPOLIssuedDate = Mid(TextLine, 3396, 2) & "/" & Mid(TextLine, 3398, 2) & "/" & Mid(TextLine, 3400, 4)
            End If
            !MBMRiskNo = Mid(TextLine, 3404, 4)
            !MBMTransNo = Mid(TextLine, 3408, 5)
            !MBMPayNo = Mid(TextLine, 3413, 8)
            !MBMPAYSeqNo = Mid(TextLine, 3421, 2)
            !MBMClientNo = Mid(TextLine, 3423, 8)
            !MBMClientID = Mid(TextLine, 3431, 2)
            !MBMPrevPLNCode = Mid(TextLine, 3433, 15)
            If Len(Mid(TextLine, 3448, 8)) Then
                !MBMRBAmt = Mid(TextLine, 3448, 8)
            End If
            !MBMPrevInsCom = Mid(TextLine, 3456, 60)
            !MBMPAYRemarks = Mid(TextLine, 3516, 60)
            TmpClmNonPayfr = Mid(TextLine, 3576, 2) & "/" & Mid(TextLine, 3578, 2) & "/" & Mid(TextLine, 3580, 4)
            If IsDate(TmpClmNonPayfr) Then
                !MBMCLMNonPayFr = TmpClmNonPayfr
            End If
            TmpClmNonPayTo = Mid(TextLine, 3584, 2) & "/" & Mid(TextLine, 3586, 2) & "/" & Mid(TextLine, 3588, 4)
            If IsDate(TmpClmNonPayTo) Then
                !MBMCLMNonPayTo = TmpClmNonPayTo
            End If
            If Trim(Mid(TextLine, 3592, 4)) <> "" And Trim(Mid(TextLine, 3596, 8)) <> "" Then
                !MBMNewPlanCode = Mid(TextLine, 3592, 4)
                !MBMNewPolEffDate = Mid(TextLine, 3596, 2) & "/" & Mid(TextLine, 3598, 2) & "/" & Mid(TextLine, 3600, 4)
            End If
            If Trim(Mid(TextLine, 3604, 20)) <> "" Then
                !MBMENDtRefNo = Mid(TextLine, 3604, 20)
            End If
        End If
        End With
        MBMOthers.Update
End Sub

Private Sub GenerateENDtMember()
    Dim PayorCode As String
    Dim CoveredIDCount As Integer
    Dim recordCount As Integer
    Dim RecCount As Integer
    Dim EndtInsType As String
    
    Open txtImportFilename For Input As #1   ' Open file.
       Line Input #1, TextLine   ' Read line into variable.
    If Mid(TextLine, 2, 2) <> Mid(txtPAYCode, 1, 2) Then
        MsgBox "Payor Selected Is Not Match The File Payor", vbOKOnly + vbCritical, DialogBoxTitle
        txtPAYCode.SetFocus
    Else
        If Mid(TextLine, 1, 1) = "E" Then
            PayorCode = Mid(TextLine, 2, 2)
        End If
        recordCount = 1
        MBM.Open "MBM", medixmasterconn, adOpenKeyset, adLockOptimistic
        Do While Not EOF(1)   ' Loop until end of file.
            Line Input #1, TextLine   ' Read line into variable.
            recordCount = recordCount + 1
            If Mid(txtFileVersion, 1, 2) = "04" Then
                EndtInsType = Mid(TextLine, 451, 1)
            Else
                EndtInsType = Mid(TextLine, 421, 1)
            End If
            If Mid(TextLine, 1, 1) = "I" Then
                If EndtInsType = "G" Or EndtInsType = "H" Then
                                 
                    If Mid(TextLine, 10, 1) = "P" Then
                        CoveredID = "00"
                        Call EndtPrincipalDataImport(Mid(txtPAYCode, 1, 2), CoveredID)
                        Call EndtOthersDataImport
                        RecCount = RecCount + 1
                    Else
                        Call EndtCoverPersonDataImport(Mid(txtPAYCode, 1, 2), CoveredID)
                        RecCount = RecCount + 1
                    End If
                End If
                            
            ElseIf Mid(TextLine, 1, 1) = "C" Then
                Call CancelENDtData
                Call ListEndtMembers
                                
                RecCount = RecCount + 1
            ElseIf Mid(TextLine, 1, 1) = "A" Then
                If Mid(TextLine, 10, 1) = "P" Then
                    Call EndtPrincipalAdjustment
                Else
                    Call EndtSupplementaryAdjustment
                End If
                Call ListEndtMembers
                RecCount = RecCount + 1
            ElseIf Mid(TextLine, 1, 1) = "S" Then
                Call EndtAddSupplementary
                Call ListEndtMembers
                RecCount = RecCount + 1
                RecCount = RecCount + 1
            End If
            DoEvents
        Loop
        MsgBox "Record Import Sucessfully...", vbOKOnly + vbInformation, DialogBoxTitle
        MBM.Close
    End If
    Close #1   ' Close file.
End Sub

Public Sub CancelENDtData()
Dim RecordSaved
Dim strsql As String
Dim CAS As New ADODB.Recordset
Dim MBM As New ADODB.Recordset
Dim strsql1 As String
Dim Cancel As New ADODB.Recordset
Dim strsql2 As String
Dim MBMCoveredPersons As New ADODB.Recordset
Dim strsql3 As String
Dim MemberAmendHis As New ADODB.Recordset
    Call EndtCancelPrincipalData
    Call EndtMBMRefund
    Call EndtCancelCoveredPersonsData
End Sub

Public Sub EndtCancelCoveredPersonsData()
Dim strsql As String
Dim MBMCoveredPersons As New ADODB.Recordset
Dim LastCoveredID  As Integer

    strsql = "Select * from MBMCoveredPersons Where MBMCNumber ='" & Mid(TextLine, 11, 18) & "'"
    MBMCoveredPersons.Open strsql, medixmasterconn, adOpenKeyset, adLockOptimistic
    
     If MBMCoveredPersons.recordCount Then
       Do Until MBMCoveredPersons.EOF
           With MBMCoveredPersons
             !MBMCStatus = "C"
             MBMCoveredPersons.Update
           End With
           MBMCoveredPersons.MoveNext
       Loop
    End If
    MBMCoveredPersons.Close
    Set MBMCoveredPersons = Nothing
End Sub

Public Sub EndtCancelPrincipalData()
Dim strsql As String
Dim MBM As New ADODB.Recordset
Dim MBMRefund As New ADODB.Recordset
Dim MemberAmendHis As New ADODB.Recordset

        strsql = "Select * from MBM Where MBMNumber ='" & Trim(Mid(TextLine, 11, 18)) & "'"
        MBM.Open strsql, medixmasterconn, adOpenKeyset, adLockOptimistic
    
        MemberAmendHis.Open "MBMAmendHistory", medixmasterconn, adOpenKeyset, adLockOptimistic
        If MemberAmendHis.recordCount Then
            MemberAmendHis.AddNew
            With MemberAmendHis
                !MBMNumber = MBM!MBMNumber
                !MBMPolicyNo = MBM!MBMPolicyNo
                !MBMIcBcPp = MBM!MBMIcBcPp
                !MBMName = MBM!MBMName
                !MBMExpiryDate = MBM!MBMPayorEXPDate
                !MBMPlaNCode = MBM!PLNCode
                !MBMAmendEdtType = "C"
                !MBMAEndtEffDate = Mid(TextLine, 2, 2) & "/" & Mid(TextLine, 4, 2) & "/" & Mid(TextLine, 6, 4)
                If IsDate(txtSubmissionDate) Then
                    !MBMAENDtBordxDate = CDate(txtSubmissionDate)
                End If
                !MBMAmendDate = Date
                !MBMAmendUser = Logon
                MemberAmendHis.Update
            End With
        End If
    If MBM.recordCount Then
        With MBM
            If Len(MBM!MBMStatus) Then
                !MBMStatus = "C"
                !MBMLastEndorseStatus = "C"
                !MBMENDtBordxDate = CDate(txtSubmissionDate)
                !MBMENDtEffDate = Mid(TextLine, 2, 2) & "/" & Mid(TextLine, 4, 2) & "/" & Mid(TextLine, 6, 4)
                !MBMLastAdjustmentDate = Date

                MBM.Update
            End If
        End With
        txtTempMemberNo = Trim(Mid(TextLine, 11, 18))
    End If
    
    MBM.Close
    Set MBM = Nothing
    MemberAmendHis.Close
    Set MemberAmendHis = Nothing
End Sub

Public Sub EndtMBMRefund()
Dim MBMRefund As New ADODB.Recordset
Dim strsql As String
Dim MCORefundValue As String
Dim MBMMCOFees As Double
Dim PremiumRefundValue As String
Dim MBMAccessFees As Double
Dim MBMAccessFeesEff As String
Dim MBMAccessFeesRef As String
Dim MCOEffective As String
Dim PremiumEffective As String
Dim FindAge As String
Dim MBMPremium As Double
Dim PlanCode As String
Dim InsuredCode As String
Dim HLTCode As String
Dim Payor As String
Dim AgeValue As String
Dim PayEffDate As String
Dim PAYExpDate As String
Dim MBMDOB As String
Dim strsqlCAS As String
Dim MBMOthers As New ADODB.Recordset

    HLTCode = Trim(Mid(cboHLTCode1, 1, 1))
    Payor = Trim(Mid(txtPAYCode, 1, 2))
            
    PlanCode = Trim(Mid(TextLine, 417, 4))
    InsuredCode = Trim(Mid(TextLine, 421, 1))
    MBMDOB = Trim(Mid(TextLine, 320, 2) & "/" & Mid(TextLine, 322, 2) & "/" & Mid(TextLine, 324, 4))
    PayEffDate = Trim(Mid(TextLine, 401, 2) & "/" & Mid(TextLine, 403, 2) & "/" & Mid(TextLine, 405, 4))
    PAYExpDate = Trim(Mid(TextLine, 409, 2) & "/" & Mid(TextLine, 411, 2) & "/" & Mid(TextLine, 413, 4))

    If Mid(txtFileVersion, 1, 2) = "03" Then
        PlanCode = Trim(Mid(TextLine, 447, 4))
        InsuredCode = Trim(Mid(TextLine, 451, 1))
        MBMDOB = Trim(Mid(TextLine, 350, 2) & "/" & Mid(TextLine, 352, 2) & "/" & Mid(TextLine, 354, 4))
        PayEffDate = Trim(Mid(TextLine, 431, 2) & "/" & Mid(TextLine, 433, 2) & "/" & Mid(TextLine, 435, 4))
        PAYExpDate = Trim(Mid(TextLine, 439, 2) & "/" & Mid(TextLine, 441, 2) & "/" & Mid(TextLine, 443, 4))
    End If
    
    strsql = "Select * from MBMRefund Where MBMNumber = '" & Trim(Mid(TextLine, 11, 18)) & "'"
    MBMRefund.Open strsql, medixmasterconn, adOpenKeyset, adLockOptimistic
    
    strsqlCAS = "Select * from MBMOthers where MBMNumber = '" & Trim(Mid(TextLine, 11, 18)) & "'"
    MBMOthers.Open strsqlCAS, medixmasterconn, adOpenKeyset, adLockOptimistic
    
    If MBMOthers.recordCount Then
        MBMPremium = MBMOthers!MBMPremium
        MBMMCOFees = MBMOthers!MBMMCOFee
        If Len(MBMOthers!MBMAccessFee) Then
            MBMAccessFees = MBMOthers!MBMAccessFee
        Else
            MBMAccessFees = 0
        End If
 
        MCORefundValue = Round(MCORefund(CDate(PAYExpDate), Mid(TextLine, 2, 2) & "/" & Mid(TextLine, 4, 2) & "/" & Mid(TextLine, 6, 4), CDate(PayEffDate), MBMMCOFees), 0)
        
        PremiumRefundValue = Round(MCOPremiumReduction(CDate(PAYExpDate), Mid(TextLine, 2, 2) & "/" & Mid(TextLine, 4, 2) & "/" & Mid(TextLine, 6, 4), CDate(PayEffDate), MBMPremium), 0)
        
        MBMAccessFeesRef = Round(AccessFeeRefund(CDate(PAYExpDate), Mid(TextLine, 2, 2) & "/" & Mid(TextLine, 4, 2) & "/" & Mid(TextLine, 6, 4), CDate(PayEffDate), MBMAccessFees), 2)
        
         PremiumEffective = CDbl(MBMPremium) - CDbl(PremiumRefundValue)
         MCOEffective = CDbl(MBMMCOFees) - CDbl(MCORefundValue)
         MBMAccessFeesEff = CDbl(MBMAccessFees) - CDbl(MBMAccessFeesRef)
         
         If MBMRefund.recordCount = 0 Then
            MBMRefund.AddNew
         End If
         With MBMRefund
            !MBMNumber = Trim(Mid(TextLine, 11, 18))
            !MBMPREMEffective = PremiumEffective
            !MBMPremReduction = PremiumRefundValue
            !MBMMCOEffective = MCOEffective
            !MBMMCORefund = MCORefundValue
            !MBMProviderRef = MBMAccessFeesRef
            !MBMProviderEff = MBMAccessFeesEff
         End With
         MBMRefund.Update
    End If
End Sub

Public Sub EndtDeleteSupplementary()
On Error Resume Next

Dim MBM As New ADODB.Recordset
Dim strsql As String
Dim listrecord As Integer
Dim listloop As Integer
Dim ListCount As Integer
Dim strsql2 As String
Dim LastCoveredID As Integer
Dim SelectAge As String
Dim strdelcov As String
Dim MBMCov As New ADODB.Recordset
Dim ForLoop As Integer

    strsql = "Select * from MBMCov Where MBMNumber = '" & Mid(TextLine, 11, 18) & "'"
    MBMCov.Open strsql, medixmasterconn, adOpenKeyset, adLockOptimistic
    
    strdelcov = "Delete from MBMCoveredPersons Where MBMCNumber ='" & Mid(TextLine, 11, 18) & "'" 'And MBMCoveredID ='" & lstCovAdd.ListItems(listloop).SubItems(10) & "'"
    MBMCov.Open strdelcov, medixmasterconn, adOpenKeyset, adLockOptimistic

    MBMCov.Close

End Sub
Private Sub EndtAddSupplementary()
    
    Dim MBM As New ADODB.Recordset
    Dim strsql As String
    Dim listrecord As Integer
    Dim listloop As Integer
    Dim ListCount As Integer
    Dim strsql2 As String
    Dim LastCoveredID As Integer
    Dim SelectAge As String
    Dim MBMCov As New ADODB.Recordset
    Dim ForLoop As Integer
    Dim BordxDate As String
    Dim PayEffDate As String
    Dim AgeValue As String
    
    strsql = "Select * from MBMCoveredPersons Where MBMCNumber = '" & Trim(Mid(TextLine, 11, 18)) & "'"
    MBMCov.Open strsql, medixmasterconn, adOpenKeyset, adLockOptimistic
    
    LastCoveredID = GetLastCoveredID(Trim(Mid(TextLine, 11, 18)))
              MBMCov.AddNew
              With MBMCov
                  !MBMCNumber = Trim(Mid(TextLine, 11, 18))
                  !MBMCSalutation = Mid(TextLine, 69, 5)
                  If Mid(txtFileVersion, 1, 2) = "03" Then ' 4830v MAA GEN
                    SelectAge = GetAge(Mid(TextLine, 431, 2) & "/" & Mid(TextLine, 433, 2) & "/" & Mid(TextLine, 435, 4), Mid(TextLine, 350, 2) & "/" & Mid(TextLine, 352, 2) & "/" & Mid(TextLine, 354, 4))
                  Else
                    SelectAge = GetAge(Mid(TextLine, 401, 2) & "/" & Mid(TextLine, 403, 2) & "/" & Mid(TextLine, 405, 4), Mid(TextLine, 320, 2) & "/" & Mid(TextLine, 322, 2) & "/" & Mid(TextLine, 324, 4))
                  End If
                  AgeValue = Mid(GetAgeBand(Trim(Mid(txtFileVersion, 1, InStr(1, txtFileVersion, "-") - 1)), SelectAge), 1, 2)
                  !MBMCAgeband = AgeValue
                    If AgeValue = "01" Then
                        !MBMCAdultChild = "C"
                    Else
                        !MBMCAdultChild = "A"
                    End If
                  !MBMCStatus = "A"
                  !MBMCName = Trim(Mid(TextLine, 74, 60))
                  If Mid(txtFileVersion, 1, 2) = "03" Then
                    !MBMCIcBcPpNo = Trim(Mid(TextLine, 330, 20))
                    !MBMCDOB = Mid(TextLine, 350, 2) & "/" & Mid(TextLine, 352, 2) & "/" & Mid(TextLine, 354, 4)
                    !MBMCSex = Mid(TextLine, 358, 1)
                    !MBMCRELCode = Mid(TextLine, 10, 1)
                    !MBMCBloodType = Mid(TextLine, 408, 3)
                    !MBMCAllergic = Mid(TextLine, 411, 20)
                  Else
                    !MBMCIcBcPpNo = Trim(Mid(TextLine, 300, 20))
                    !MBMCDOB = Mid(TextLine, 320, 2) & "/" & Mid(TextLine, 322, 2) & "/" & Mid(TextLine, 324, 4)
                    !MBMCSex = Mid(TextLine, 328, 1)
                    !MBMCRELCode = Mid(TextLine, 10, 1)
                    !MBMCBloodType = Mid(TextLine, 378, 3)
                    !MBMCAllergic = Mid(TextLine, 381, 20)
                  End If
                  
                  If Mid(txtFileVersion, 1, 2) = "01" Then
                    !MBMCExclusion = Mid(TextLine, 463, 60)
                  ElseIf Mid(txtFileVersion, 1, 2) = "02" Then
                    !MBMCExclusion = Mid(TextLine, 463, 2500)
                  ElseIf Mid(txtFileVersion, 1, 2) = "03" Then
                    !MBMCExclusion = Mid(TextLine, 493, 2500)
                  End If
                  If Len(Mid(TextLine, 3442, 8)) Then
                    !MBMCPAYNo = Mid(TextLine, 3442, 8)
                  End If
                  If Len(Mid(TextLine, 3450, 2)) Then
                    !MBMCPAYSeqNo = Mid(TextLine, 3450, 2)
                  End If
                  If Len(Mid(TextLine, 3452, 8)) Then
                    !MBMCClientNo = Mid(TextLine, 3452, 8)
                  End If
                  If Len(Mid(TextLine, 3460, 2)) Then
                    !MBMCClientID = Mid(TextLine, 3460, 2)
                  End If
                  !MBMCMemberType = "C"
                  !MBMCCoverID = LastCoveredID
             MBMCov.Update
             End With
      txtTempMemberNo = Trim(Mid(TextLine, 11, 18))
    
    MBMCov.Close
End Sub
Public Sub EndtPrincipalAdjustment()
'On Error Resume Next
    Dim MemberHis As New ADODB.Recordset
    Dim MBM As New ADODB.Recordset
    Dim MBMCov As New ADODB.Recordset
    Dim strsql As String
    Dim listrecord As Integer
    Dim listloop As Integer
    Dim ForLoop As Integer
    Dim ListCount As Integer
    Dim strsql2 As String
    Dim LastCoveredID As Integer
    Dim SelectAge As String
    Dim MemberAmendHis As New ADODB.Recordset
    Dim EndtExpDate As String
    
    If Mid(TextLine, 10, 1) = "P" Then
        strsql = "Select * from MBM Where MBMNumber = '" & Trim(Mid(TextLine, 11, 18)) & "'"
        MBM.Open strsql, medixmasterconn, adOpenKeyset, adLockOptimistic
                
        MemberAmendHis.Open "MBMAmendHistory", medixmasterconn, adOpenKeyset, adLockOptimistic
        
    If MBM.recordCount Then
        MemberAmendHis.AddNew
        With MemberAmendHis
            !MBMNumber = MBM!MBMNumber
            !MBMPolicyNo = MBM!MBMPolicyNo
            !MBMIcBcPp = MBM!MBMIcBcPp
            !MBMName = MBM!MBMName
            !MBMExpiryDate = MBM!MBMPayorEXPDate
            !MBMPlaNCode = MBM!PLNCode
            !MBMAmendEdtType = "C"
            !MBMAEndtEffDate = Mid(TextLine, 2, 2) & "/" & Mid(TextLine, 4, 2) & "/" & Mid(TextLine, 6, 4)
            If IsDate(txtSubmissionDate) Then
                !MBMAENDtBordxDate = CDate(txtSubmissionDate)
            End If
            !MBMAmendUser = Logon
            MemberAmendHis.Update
        End With
        With MBM
            !MBMENDtBordxDate = CDate(txtSubmissionDate)
            !MBMENDtEffDate = Mid(TextLine, 2, 2) & "/" & Mid(TextLine, 4, 2) & "/" & Mid(TextLine, 6, 4)
            !MBMPolicyNo = Mid(TextLine, 29, 25)
            !MBMSalutation = Mid(TextLine, 69, 5)
            !MBMName = Mid(TextLine, 74, 60)
            !MBMPolicyNo = Mid(TextLine, 29, 25)
            !MBMAddress1 = Mid(TextLine, 134, 30)
            !MBMAddress2 = Mid(TextLine, 164, 30)
            !MBMAddress3 = Mid(TextLine, 194, 30)
            !MBMStatus = "A"
            
            If Mid(txtFileVersion, 1, 2) = "01" Or Mid(txtFileVersion, 1, 2) = "02" Then
                !CTYCode = Mid(TextLine, 224, 20)
                !MBMPostCode = Mid(TextLine, 244, 5)
                !STACode = Mid(TextLine, 249, 15)
                !MBMTelnoH = Mid(TextLine, 264, 12)
                !MBMTelNoO = Mid(TextLine, 276, 12)
                !MBMTelNoM = Mid(TextLine, 288, 12)
                !MBMDOB = Mid(TextLine, 320, 2) & "/" & Mid(TextLine, 322, 2) & "/" & Mid(TextLine, 324, 4)
                !MBMSex = Mid(TextLine, 328, 1)
                !LGNCode = Mid(TextLine, 345, 1)
                !RACCode = Mid(TextLine, 346, 1)
                !MBMPayorEffDate = Mid(TextLine, 401, 2) & "/" & Mid(TextLine, 403, 2) & "/" & Mid(TextLine, 405, 4)
                !PLNCode = Mid(TextLine, 417, 4)
                !INSCode = Mid(TextLine, 421, 1)
                
                If Mid(txtFileVersion, 1, 2) = "01" Then
                    !BRNCode = Mid(TextLine, 526, 15)
                ElseIf Mid(txtFileVersion, 1, 2) = "02" Then
                    !BRNCode = Mid(TextLine, 2968, 15)
                End If
                If Mid(txtFileVersion, 1, 2) = "02" Then
                    !MBMAppName = Mid(TextLine, 3653, 60)
                End If
            ElseIf Mid(txtFileVersion, 1, 2) = "03" Then
                !CTYCode = Mid(TextLine, 224, 50)
                !MBMPostCode = Mid(TextLine, 274, 5)
                !STACode = Mid(TextLine, 279, 15)
                !MBMTelnoH = Mid(TextLine, 294, 12)
                !MBMTelNoO = Mid(TextLine, 306, 12)
                !MBMTelNoM = Mid(TextLine, 318, 12)
                !MBMDOB = Mid(TextLine, 350, 2) & "/" & Mid(TextLine, 352, 2) & "/" & Mid(TextLine, 354, 4)
                !MBMSex = Mid(TextLine, 358, 1)
                !LGNCode = Mid(TextLine, 375, 1)
                !RACCode = Mid(TextLine, 376, 1)
                !MBMPayorEffDate = Mid(TextLine, 431, 2) & "/" & Mid(TextLine, 433, 2) & "/" & Mid(TextLine, 435, 4)
                !BRNCode = Mid(TextLine, 2998, 15)
                !MBMAppName = Mid(TextLine, 3653, 60)
            End If
            !MBMLastAdjustmentDate = txtSubmissionDate
            MBM.Update
        End With
    End If
    MemberAmendHis.Close
    Set MemberAmendHis = Nothing
    MBM.Close
    Set MBM = Nothing
    End If
End Sub
Private Sub EndtSupplementaryAdjustment()
    
    Dim MBM As New ADODB.Recordset
    Dim strsql As String
    Dim strsql2 As String
    Dim listrecord As Integer
    Dim listloop As Integer
    Dim ListCount As Integer
    Dim LastCoveredID As Integer
    Dim SelectAge As String
    Dim MBMCov As New ADODB.Recordset
    Dim ForLoop As Integer
    Dim MemberDOB As String
    Dim PayEffDate As String
    Dim AgeValue As String
    Dim FindAge As String
    Dim SuppID As String
    Dim TmpClmNonPayfr As String
    Dim TmpClmNonPayTo As String

    strsql2 = "select * from MBM where MBMNumber = '" & Mid(TextLine, 11, 18) & "'"
    MBM.Open "MBM", medixmasterconn, adOpenKeyset, adLockOptimistic
    
    If Mid(txtFileVersion, 1, 2) = "01" Then
        strsql = "Select * from MBMCoveredPersons Where MBMCNumber = '" & Mid(TextLine, 11, 18) & "' AND MBMCICBCPPNo='" & Mid(TextLine, 74, 60) & "'"
        MBMCov.Open strsql, medixmasterconn, adOpenKeyset, adLockOptimistic
        If MBMCov.recordCount = 0 Then
            MBMCov.Close
            Set MBMCov = Nothing
            strsql = "Select * from MBMCoveredPersons Where MBMCNumber = '" & Mid(TextLine, 11, 18) & "' AND MBMCICBCPPNo='" & Mid(TextLine, 300, 20) & "'"
            MBMCov.Open strsql, medixmasterconn, adOpenKeyset, adLockOptimistic
        End If
    Else
        If Mid(txtFileVersion, 1, 2) = "02" Then 'Version 4830
            SuppID = Mid(TextLine, 2963, 2)
        ElseIf Mid(txtFileVersion, 1, 2) = "03" Then 'Version 4830 (MAA Gen)
            SuppID = Mid(TextLine, 2993, 2)
        End If
        strsql = "Select * from MBMCoveredPersons Where MBMCNumber = '" & Mid(TextLine, 11, 18) & "' and MBMCCoverID = '" & SuppID & "'"
        MBMCov.Open strsql, medixmasterconn, adOpenKeyset, adLockOptimistic
    End If
    PayEffDate = MBM!MBMPayorEffDate
    LastCoveredID = GetLastCoveredID(Mid(TextLine, 11, 18))
    If MBMCov.recordCount Then
        With MBMCov
            !MBMCSalutation = Mid(TextLine, 69, 5)
            !MBMCName = Mid(TextLine, 74, 60)
            
            If Mid(txtFileVersion, 1, 2) = "01" Or Mid(txtFileVersion, 1, 2) = "02" Then
                !MBMCSex = Mid(TextLine, 328, 1)
                !MBMCDOB = Mid(TextLine, 320, 2) & "/" & Mid(TextLine, 322, 2) & "/" & Mid(TextLine, 324, 4)
                !MBMCIcBcPpNo = Mid(TextLine, 300, 20)
'                !MBMCOccupation = Mid(TextLine, 347, 30)
                !MBMCWorkNature = Mid(TextLine, 377, 1)
                !MBMCBloodType = Mid(TextLine, 378, 3)
                !MBMCAllergic = Mid(TextLine, 381, 20)
                If Mid(txtFileVersion, 1, 2) = "01" Then
                    !MBMCExclusion = Mid(TextLine, 463, 60)
                    !MBMCSmoking = Mid(TextLine, 523, 1)
                    !MBMCAlcohol = Mid(TextLine, 524, 1)
                ElseIf Mid(txtFileVersion, 1, 2) = "02" Then
                    !MBMCExclusion = Mid(TextLine, 463, 2500)
                    !MBMCSmoking = Mid(TextLine, 2965, 1)
                    !MBMCAlcohol = Mid(TextLine, 2966, 1)
                End If
                    
            ElseIf Mid(txtFileVersion, 1, 2) = "03" Then
                !MBMCSex = Mid(TextLine, 358, 1)
                !MBMCDOB = Mid(TextLine, 350, 2) & "/" & Mid(TextLine, 352, 2) & "/" & Mid(TextLine, 354, 4)
                !MBMCIcBcPpNo = Mid(TextLine, 330, 20)
'                !MBMCOccupation = Mid(TextLine, 377, 30)
                !MBMCWorkNature = Mid(TextLine, 407, 1)
                !MBMCBloodType = Mid(TextLine, 408, 3)
                !MBMCAllergic = Mid(TextLine, 411, 20)
                !MBMCExclusion = Mid(TextLine, 493, 2500)
                !MBMCSmoking = Mid(TextLine, 2995, 1)
                !MBMCAlcohol = Mid(TextLine, 2996, 1)
                
                !MBMCPrevPLNCode = Mid(TextLine, 3462, 15)
                If IsNumeric(Mid(TextLine, 3477, 8)) Then
                    !MBMCRBAmt = Mid(TextLine, 3477, 8)
                End If
                !MBMCPrevInsCom = Mid(TextLine, 3485, 60)
                !MBMCPAYRemarks = Mid(TextLine, 3545, 60)
            End If
            If Mid(txtFileVersion, 1, 2) = "03" Then
                MemberDOB = Mid(TextLine, 350, 2) & "/" & Mid(TextLine, 352, 2) & "/" & Mid(TextLine, 354, 4)
            Else
                MemberDOB = Mid(TextLine, 320, 2) & "/" & Mid(TextLine, 322, 2) & "/" & Mid(TextLine, 324, 4)
            End If
            FindAge = GetAge(CDate(PayEffDate), CDate(MemberDOB))
            AgeValue = Mid(GetAgeBand(Mid(txtFileVersion, 1, 2), FindAge), 1, 2)
            !MBMCAgeband = AgeValue
            If AgeValue = "01" Then
                !MBMCAdultChild = "C"
            Else
                !MBMCAdultChild = "A"
            End If
            TmpClmNonPayfr = Mid(TextLine, 3605, 2) & "/" & Mid(TextLine, 3607, 2) & "/" & Mid(TextLine, 3609, 4)
            If IsDate(TmpClmNonPayfr) Then
                !MBMCCLMNonPayFr = TmpClmNonPayfr
            End If
            TmpClmNonPayTo = Mid(TextLine, 3613, 2) & "/" & Mid(TextLine, 3615, 2) & "/" & Mid(TextLine, 3617, 4)
            If IsDate(TmpClmNonPayTo) Then
                !MBMCCLMNonPayTo = TmpClmNonPayTo
            End If
            If Len(Mid(TextLine, 3442, 8)) Then
                !MBMCPAYNo = Mid(TextLine, 3442, 8)
            End If
            If Len(Mid(TextLine, 3450, 2)) Then
                !MBMCPAYSeqNo = Mid(TextLine, 3450, 2)
            End If
            If Len(Mid(TextLine, 3452, 8)) Then
                !MBMCClientNo = Mid(TextLine, 3452, 8)
            End If
            If Len(Mid(TextLine, 3460, 2)) Then
                !MBMCClientID = Mid(TextLine, 3460, 2)
            End If
        End With
        MBMCov.Update
        MBMCov.Close
    End If
End Sub

Private Sub EndtPrincipalDataImport(Payor As String, CoveredIDCount As Integer)
'On Error Resume Next
'On Error GoTo error1
Dim PayEffDate As String
Dim PAYExpDate As String
Dim MemberDOB As String
Dim FindAge As String
Dim InsuredCode As String
Dim PlanCode As String
Dim AgeValue As String
Dim AnnualLimitValue As String
Dim GetProString As String
Dim tmpPreMemberNo As String
Dim ListMember As ListItem
Dim MBM As New ADODB.Recordset
Dim MBMOthers As New ADODB.Recordset
Dim TempPolDis As Double
Dim TempRenewDis As Double
Dim TmpSql As String
'Dim DiscPlan As New ADODB.Recordset
        
        MBM.Open "MBM", medixmasterconn, adOpenKeyset, adLockOptimistic
        
        MBM.AddNew
        With MBM
            If Mid(TextLine, 10, 1) = "P" Then
                !MBMLastEndorseStatus = Mid(TextLine, 1, 1)
                !MBMENDtBordxDate = CDate(txtSubmissionDate)
                !MBMENDtEffDate = Mid(TextLine, 2, 2) & "/" & Mid(TextLine, 4, 2) & "/" & Mid(TextLine, 6, 4)
                !MBMMemberType = "P"
                !MBMPolicyNo = Mid(TextLine, 29, 25)
                !MBMSalutation = Mid(TextLine, 67, 5)
                !MBMName = Mid(TextLine, 74, 60)
                !MBMAddress1 = Mid(TextLine, 134, 30)
                !MBMAddress2 = Mid(TextLine, 164, 30)
                !MBMAddress3 = Mid(TextLine, 194, 30)
                !PAYCode = Mid(txtPAYCode, 1, 2)
                !HLTCode = Mid(cboHLTCode1, 1, 1)
                If Len(Trim(Mid(TextLine, 3771, 20))) Then
                    !MBMIcBcPp2nd = Trim(Mid(TextLine, 3771, 20))
                End If

                If Mid(txtFileVersion, 1, 2) = "01" Or Mid(txtFileVersion, 1, 2) = "02" Then
                    !CTYCode = Mid(TextLine, 224, 20)
                    !MBMPostCode = Mid(TextLine, 244, 5)
                    !STACode = Mid(TextLine, 249, 15)
                    !MBMTelnoH = Mid(TextLine, 264, 12)
                    !MBMTelNoM = Mid(TextLine, 276, 12)
                    !MBMTelNoO = Mid(TextLine, 288, 12)
                    !MBMIcBcPp = Mid(TextLine, 300, 20)
                    !MBMDOB = Mid(TextLine, 320, 2) & "/" & Mid(TextLine, 322, 2) & "/" & Mid(TextLine, 324, 4)
                    !MBMSex = Mid(TextLine, 328, 1)
                    !MBMPayorEffDate = Mid(TextLine, 401, 2) & "/" & Mid(TextLine, 403, 2) & "/" & Mid(TextLine, 405, 4)
                    !MBMPayorEXPDate = Mid(TextLine, 409, 2) & "/" & Mid(TextLine, 411, 2) & "/" & Mid(TextLine, 413, 4)
                    !PLNCode = Mid(TextLine, 417, 4)
                    !INSCode = Mid(TextLine, 421, 1)
                    If Mid(TextLine, 1, 2) = "02" Then
                        !MBMAppName = Mid(TextLine, 3653, 60)
                    End If
                    
                ElseIf Mid(txtFileVersion, 1, 2) = "03" Then
                    !CTYCode = Mid(TextLine, 224, 50)
                    !MBMPostCode = Mid(TextLine, 274, 5)
                    !STACode = Mid(TextLine, 279, 15)
                    !MBMTelnoH = Mid(TextLine, 294, 12)
                    !MBMTelNoM = Mid(TextLine, 306, 12)
                    !MBMTelNoO = Mid(TextLine, 318, 12)
                    !MBMIcBcPp = Mid(TextLine, 330, 20)
                    !MBMDOB = Mid(TextLine, 350, 2) & "/" & Mid(TextLine, 352, 2) & "/" & Mid(TextLine, 354, 4)
                    !MBMSex = Mid(TextLine, 358, 1)
                    !MBMPayorEffDate = Mid(TextLine, 431, 2) & "/" & Mid(TextLine, 433, 2) & "/" & Mid(TextLine, 435, 4)
                    !MBMPayorEXPDate = Mid(TextLine, 439, 2) & "/" & Mid(TextLine, 441, 2) & "/" & Mid(TextLine, 443, 4)
                    !PLNCode = Mid(TextLine, 447, 4)
                    !INSCode = Mid(TextLine, 451, 1)
                    !MBMAppName = Mid(TextLine, 3653, 60)
                End If
                InsuredCode = Trim(!INSCode)
                PlanCode = Trim(!PLNCode)
                !MBMBordxDate = CDate(txtSubmissionDate)
                !MBMDataStatus = "A"
                !MBMCOVID = "00"
                !MBMValueDate = Date
                !SRVCodeList = "EA"
                
        '------------------------------------------------------------------------------------------------------------------------
                !MBMStatus = "A"
                If Len(txtBatchNo) Then
                    !MBMBatchNo = txtBatchNo
                End If
                !MBMExpiredStatus = "N"
        '------------------------------------------------------------------------------------------------------------------------
                If Mid(txtFileVersion, 1, 2) = "01" Or Mid(txtFileVersion, 1, 2) = "02" Then
                    PayEffDate = Mid(TextLine, 401, 2) & "/" & Mid(TextLine, 403, 2) & "/" & Mid(TextLine, 405, 4)
                    PAYExpDate = Mid(TextLine, 409, 2) & "/" & Mid(TextLine, 411, 2) & "/" & Mid(TextLine, 413, 4)
                    MemberDOB = Mid(TextLine, 320, 2) & "/" & Mid(TextLine, 322, 2) & "/" & Mid(TextLine, 324, 4)
                ElseIf Mid(txtFileVersion, 1, 2) = "03" Then
                    PayEffDate = Mid(TextLine, 431, 2) & "/" & Mid(TextLine, 433, 2) & "/" & Mid(TextLine, 435, 4)
                    PAYExpDate = Mid(TextLine, 439, 2) & "/" & Mid(TextLine, 441, 2) & "/" & Mid(TextLine, 443, 4)
                    MemberDOB = Mid(TextLine, 350, 2) & "/" & Mid(TextLine, 352, 2) & "/" & Mid(TextLine, 354, 4)
                End If
                FindAge = GetAge(CDate(PayEffDate), CDate(MemberDOB))
                AgeValue = Mid(GetAgeBand(Mid(txtFileVersion, 1, 2), FindAge), 1, 2)
                !AgeCode = AgeValue
                If AgeValue = "01" Then
                    !MBMAdultChild = "C"
                Else
                    !MBMAdultChild = "A"
                End If
                
                    txtMBMPremium = Format(GetPremium(InsuredCode, Payor, AgeValue, PlanCode, Mid(cboHLTCode1, 1, 1), CDate(PayEffDate)), "#,##0.00")
                    txtMBMPremium = Round(InsertPremium(CDate(PAYExpDate), CDate(PayEffDate), Format(txtMBMPremium, "#,##0.00")), 0)
                    BasicPrem = txtMBMPremium
                    TempPolDis = IIf(IsNumeric(Mid(TextLine, 3761, 5)), Mid(TextLine, 3761, 5), 0)
                    TempRenewDis = IIf(IsNumeric(Mid(TextLine, 3766, 5)), Mid(TextLine, 3766, 5), 0)
                
'                    TmpSql = "select * from DiscPlan where PLNCode = '" & Trim(PlanCode) & "'"
'                    DiscPlan.Open TmpSql, medixmasterconnMaint, adOpenKeyset, adLockOptimistic
'                    If DiscPlan.recordCount Then
                        If TempPolDis > 0 Then
                            txtMBMPremium = txtMBMPremium * (100 - TempPolDis) / 100
                        End If
                        If TempRenewDis > 0 Then
                            txtMBMPremium = txtMBMPremium * (100 - TempRenewDis) / 100
                        End If
'                    End If
'                    DiscPlan.Close
'                    Set DiscPlan = Nothing

                    txtMBMPremium = Format(Round(txtMBMPremium, 0), "#,##0.00")

                    If AgeValue = "01" Then
                        txtMBMMCOFees = Format(GetMCO(InsuredCode, Mid(cboHLTCode1, 1, 1), "C", CDate(PayEffDate), PlanCode), "#,##0.00")
                    Else
                        txtMBMMCOFees = Format(GetMCO(InsuredCode, Mid(cboHLTCode1, 1, 1), "A", CDate(PayEffDate), PlanCode), "#,##0.00")
                    End If
                    txtMBMMCOFees = Round(InsertMCOFees(CDate(PAYExpDate), CDate(PayEffDate), Format(txtMBMMCOFees, "#,##0.00")), 0)

                    AnnualLimitValue = GetAnnualLimit(InsuredCode, PlanCode, Mid(cboHLTCode1, 1, 1), CDate(PayEffDate))
                    !MBMAvailableLimit = AnnualLimitValue
                                                    
                    If AgeValue = "01" Then
                        GetProString = GetProviderCharges("EA", InsuredCode, Mid(cboHLTCode1, 1, 1), "C")
                    Else
                        GetProString = GetProviderCharges("EA", InsuredCode, Mid(cboHLTCode1, 1, 1), "A")
                    End If
                    If Len(Trim(GetProString)) Then
                        txtAccessFees = Format(GetProString, "#,##0.00")
                    Else
                        txtAccessFees = 0
                    End If
                        txtAccessFees = Round(InsertProvider(CDate(PAYExpDate), CDate(PayEffDate), Format(txtAccessFees, "#,##0.00")), 2)
                                                         
                    tmpPreMemberNo = Payor & PlanCode & InsuredCode & GenerateMembershipNo(Payor, PlanCode, InsuredCode)
                    !MBMNumber = tmpPreMemberNo
                    txtTempMemberNo = tmpPreMemberNo
            End If
            End With
                MBM.Update
    MBM.Close
    Set MBM = Nothing
    
End Sub

Private Sub EndtOthersDataImport()
'On Error Resume Next
'On Error GoTo error1
    Dim tmpPreMemberNo As String
    Dim MBMOthers As New ADODB.Recordset
    Dim TmpClmNonPayfr As String
    Dim TmpClmNonPayTo As String
    
        MBMOthers.Open "MBMOthers", medixmasterconn, adOpenKeyset, adLockOptimistic
       'medixmasterconn.BeginTrans
        
        MBMOthers.AddNew
        With MBMOthers
            !MBMNumber = txtTempMemberNo
            !MBMAgentCode = Mid(TextLine, 54, 15)
            !MBMWeight = Mid(TextLine, 339, 3)
            !MBMHeight = Mid(TextLine, 342, 3)
            !MBMWorkNature = Mid(TextLine, 377, 1)
            !MBMOccupation = Mid(TextLine, 347, 30)
            !MBMBlood = Mid(TextLine, 351, 3)
            !MBMAllergic = Mid(TextLine, 381, 20)
            !MBMFirstMEMNo = txtTempMemberNo
            !MBMGroupCompany = Mid(TextLine, 422, 40)
            If Len(Trim(Mid(TextLine, 3713, 12))) Then
                !MBMPayorPremGross = IIf(IsNumeric(Trim(Mid(TextLine, 3713, 12))), Trim(Mid(TextLine, 3713, 12)), 0)
            End If
            If Len(Trim(Mid(TextLine, 3725, 12))) Then
                !MBMPayorPremNet = IIf(IsNumeric(Trim(Mid(TextLine, 3725, 12))), Trim(Mid(TextLine, 3725, 12)), 0)
            End If
            If Len(Trim(Mid(TextLine, 3737, 12))) Then
                !MBMPayorMCONet = IIf(IsNumeric(Trim(Mid(TextLine, 3737, 12))), Trim(Mid(TextLine, 3737, 12)), 0)
            End If
            
            If Len(Trim(Mid(TextLine, 3749, 12))) Then
                !MBMPayorMCOGross = IIf(IsNumeric(Trim(Mid(TextLine, 3749, 12))), Trim(Mid(TextLine, 3749, 12)), 0)
            End If
            If Len(Trim(Mid(TextLine, 3761, 5))) Then
                !MBMPolicyDisc = IIf(IsNumeric(Trim(Mid(TextLine, 3761, 5))), Trim(Mid(TextLine, 3761, 5)), 0)
            End If

            If Len(Trim(Mid(TextLine, 3766, 5))) Then
                !MBMRenewalDisc = IIf(IsNumeric(Trim(Mid(TextLine, 3766, 5))), Trim(Mid(TextLine, 3766, 5)), 0)
            End If

            If Mid(txtFileVersion, 1, 2) = "01" Then
                !MBMExclusion = Mid(TextLine, 463, 60)
                !MBMSmoking = Mid(TextLine, 523, 1)
                !MBMAlcohol = Mid(TextLine, 524, 1)
                !MBMMaritalStatus = Mid(TextLine, 525, 1)
                !MBMPrevPolNo = Mid(TextLine, 541, 25)
                !MBMPrevMemNo = Mid(TextLine, 566, 18)
                !MBMEmployeeNo = Mid(TextLine, 584, 15)
            ElseIf Mid(txtFileVersion, 1, 2) = "02" Or Mid(txtFileVersion, 1, 2) = "03" Then
                If Mid(txtFileVersion, 1, 2) = "02" Then
                    !MBMExclusion = Mid(TextLine, 463, 2500)
                    !MBMSmoking = Mid(TextLine, 2965, 1)
                    !MBMAlcohol = Mid(TextLine, 2966, 1)
                    !MBMMaritalStatus = Mid(TextLine, 2967, 1)
                    !MBMBranch = Mid(TextLine, 2968, 15)
                    !MBMPrevPolNo = Mid(TextLine, 2983, 25)
                    !MBMPrevMemNo = Mid(TextLine, 3008, 18)
                    !MBMEmployeeNo = Mid(TextLine, 3026, 15)
                ElseIf Mid(txtFileVersion, 1, 2) = "03" Then
                    !MBMExclusion = Mid(TextLine, 493, 2500)
                    !MBMSmoking = Mid(TextLine, 2995, 1)
                    !MBMAlcohol = Mid(TextLine, 2996, 1)
                    !MBMMaritalStatus = Mid(TextLine, 2997, 1)
                    !MBMBranch = Mid(TextLine, 2998, 15)
                    !MBMPrevPolNo = Mid(TextLine, 3013, 25)
                    !MBMPrevMemNo = Mid(TextLine, 3038, 18)
                    !MBMEmployeeNo = Mid(TextLine, 3056, 15)
                    !MBMStaffDec = Trim(Mid(TextLine, 3079, 1))
                    If Len(Mid(TextLine, 3080, 5)) Then
                        !MBMGenWP = Mid(TextLine, 3080, 5)
                        !MBMGenEndWP = Format(DateAdd("d", Trim(Mid(TextLine, 3080, 5)), PayEffDate), "dd/mm/yyyy")
                    Else
                        !MBMGenWP = ""
                        !MBMGenEndWP = ""
                    End If
                      
                    If Len(Mid(TextLine, 3085, 5)) Then
                        !MBMPreExWP = Trim(Mid(TextLine, 3085, 5))
                        !MBMPreExEndWP = Format(DateAdd("d", Trim(Mid(TextLine, 3085, 5)), PayEffDate), "dd/mm/yyyy")
                    Else
                        !MBMPreExWP = ""
                        !MBMPreExEndWP = ""
                    End If
                                
                    If Len(Mid(TextLine, 3090, 5)) Then
                        !MBMSpeIllWP = Mid(TextLine, 3090, 5)
                        !MBMSpeIllEndWP = Format(DateAdd("d", Trim(Mid(TextLine, 3090, 5)), PayEffDate), "dd/mm/yyyy")
                    Else
                        !MBMSpeIllWP = ""
                        !MBMSpeIllEndWP = ""
                    End If
                End If
                !MBMDepartment = Mid(TextLine, 3250, 30)
                !MBMGRPCompanyAdd1 = Mid(TextLine, 3280, 30)
                !MBMGRPCompanyAdd2 = Mid(TextLine, 3310, 30)
                !MBMGRPCompanyAdd3 = Mid(TextLine, 3340, 30)
                !MBMGRPCompanyAdd4 = Mid(TextLine, 3370, 30)
                !MBMInstallment = Mid(TextLine, 3404, 1)
                !MBMPolSubNo1 = Mid(TextLine, 3405, 10)
                !MBMPOLSubNo2 = Mid(TextLine, 3415, 10)
                !MBMPOLIssuedDate = Mid(TextLine, 3425, 8)
                !MBMRiskNo = Mid(TextLine, 3433, 4)
                !MBMTranNo = Mid(TextLine, 3437, 5)
                !MBMPayNo = Mid(TextLine, 3442, 8)
                !MBMPAYSeqNo = Mid(TextLine, 3450, 2)
                !MBMClientNo = Mid(TextLine, 3452, 8)
                !MBMClientID = Mid(TextLine, 3460, 2)
                !MBMCoPayRB = Mid(TextLine, 3605, 2)
                TmpClmNonPayfr = Mid(TextLine, 3605, 2) & "/" & Mid(TextLine, 3607, 2) & "/" & Mid(TextLine, 3609, 4)
                If IsDate(TmpClmNonPayfr) Then
                    !MBMCLMNonPayFr = TmpClmNonPayfr
                End If
                TmpClmNonPayTo = Mid(TextLine, 3613, 2) & "/" & Mid(TextLine, 3615, 2) & "/" & Mid(TextLine, 3617, 4)
                If IsDate(TmpClmNonPayTo) Then
                    !MBMCLMNonPayTo = TmpClmNonPayTo
                End If
                If Trim(Mid(TextLine, 3621, 4)) <> "" And Trim(Mid(TextLine, 3625, 8)) <> "" Then
                    !MBMNewPlanCode = Mid(TextLine, 3621, 4)
                    !MBMNewPolEffDate = Mid(TextLine, 3625, 2) & "/" & Mid(TextLine, 3627, 2) & "/" & Mid(TextLine, 3629, 4)
                End If
                If Trim(Mid(TextLine, 3633, 20)) <> "" Then
                    !MBMENDtRefNo = Mid(TextLine, 3633, 20)
                End If
                                
            End If
            !MBMAccessFee = txtAccessFees
            !MBMPremium = txtMBMPremium
            !MBMBasicPrem = BasicPrem
            !MBMMCOFee = txtMBMMCOFees
        End With
        MBMOthers.Update
    MBMOthers.Close
    Set MBMOthers = Nothing
End Sub
Private Sub EndtCoverPersonDataImport(Payor As String, CoveredIDCount As Integer)
'On Error GoTo error1
    Dim PayEffDate As String
    Dim MemberDOB As String
    Dim FindAge As String
    Dim AgeValue As String
    Dim GetProString As String
    Dim tmpPreMemberNo As String
    Dim ListMember As ListItem
    Dim MBMCover As New ADODB.Recordset
        
        MBMCover.Open "MBMCoveredPersons", medixmasterconn, adOpenKeyset, adLockOptimistic
       ' medixmasterconn.BeginTrans
        CoveredIDCount = CoveredIDCount + 1
        
        If CoveredIDCount < 10 Then
            CoveredIDCount = 0 & CoveredIDCount
        End If
        
        MBMCover.AddNew
        With MBMCover
            If Mid(TextLine, 1, 1) <> "P" Then
                !MBMCRELCode = Mid(TextLine, 10, 1)
                !MBMCCoverID = CoveredIDCount
                !MBMCMemberType = "C"
                !MBMCNumber = txtTempMemberNo
            End If
            !MBMCSalutation = Mid(TextLine, 69, 5)
            !MBMCName = Mid(TextLine, 74, 60)
            !MBMCIcBcPpNo = Mid(TextLine, 300, 20)
            If Len(Trim(Mid(TextLine, 3713, 12))) Then
                !MBMCPayorPremGross = IIf(IsNumeric(Trim(Mid(TextLine, 3713, 12))), Trim(Mid(TextLine, 3713, 12)), 0)
            End If
            If Len(Trim(Mid(TextLine, 3725, 12))) Then
                !MBMCPayorPremNet = IIf(IsNumeric(Trim(Mid(TextLine, 3725, 12))), Trim(Mid(TextLine, 3725, 12)), 0)
            End If
            If Len(Trim(Mid(TextLine, 3737, 12))) Then
                !MBMCPayorMCONet = IIf(IsNumeric(Trim(Mid(TextLine, 3737, 12))), Trim(Mid(TextLine, 3737, 12)), 0)
            End If
            
            If Len(Trim(Mid(TextLine, 3749, 12))) Then
                !MBMCPayorMCOGross = IIf(IsNumeric(Trim(Mid(TextLine, 3749, 12))), Trim(Mid(TextLine, 3749, 12)), 0)
            End If
            If Len(Trim(Mid(TextLine, 3761, 5))) Then
                !MBMCPolicyDisc = IIf(IsNumeric(Trim(Mid(TextLine, 3761, 5))), Trim(Mid(TextLine, 3761, 5)), 0)
            End If

            If Len(Trim(Mid(TextLine, 3766, 5))) Then
                !MBMCRenewalDisc = IIf(IsNumeric(Trim(Mid(TextLine, 3766, 5))), Trim(Mid(TextLine, 3766, 5)), 0)
            End If
            If Len(Trim(Mid(TextLine, 3771, 20))) Then
                !MBMCIcBcPp2nd = Trim(Mid(TextLine, 3771, 20))
            End If
            If Mid(txtFileVersion, 1, 1) = "01" Or Mid(txtFileVersion, 1, 1) = "02" Then
                !MBMCDOB = Mid(TextLine, 320, 2) & "/" & Mid(TextLine, 322, 2) & "/" & Mid(TextLine, 324, 4)
                !MBMCSex = Mid(TextLine, 328, 1)
                If Mid(txtFileVersion, 1, 1) = "01" Then
                    !MBMCExclusion = Mid(TextLine, 463, 60)
                ElseIf Mid(txtFileVersion, 1, 1) = "02" Then
                    !MBMCExclusion = Mid(TextLine, 463, 2500)
                End If
                
            ElseIf Mid(txtFileVersion, 1, 1) = "03" Then
                !MBMCDOB = Mid(TextLine, 350, 2) & "/" & Mid(TextLine, 352, 2) & "/" & Mid(TextLine, 354, 4)
                !MBMCSex = Mid(TextLine, 358, 1)
                !MBMCExclusion = Mid(TextLine, 493, 2500)
            'Cater for different Plan/Benefit
                !MBMCPrevPLNCode = Mid(TextLine, 3462, 15)
                !MBMCRBAmt = Mid(TextLine, 3477, 8)
                !MBMCPrevInsCom = Mid(TextLine, 3485, 60)
                !MBMCPAYRemarks = Mid(TextLine, 3545, 60)
            End If
            If Len(Mid(TextLine, 3442, 8)) Then
                !MBMCPAYNo = Mid(TextLine, 3442, 8)
            End If
            If Len(Mid(TextLine, 3450, 2)) Then
                !MBMCPAYSeqNo = Mid(TextLine, 3450, 2)
            End If
            
            If Len(Mid(TextLine, 3452, 8)) Then
                !MBMCClientNo = Mid(TextLine, 3452, 8)
            End If
            If Len(Mid(TextLine, 3460, 2)) Then
                !MBMCClientID = Mid(TextLine, 3460, 2)
            End If
            
            !MBMCStatus = "A"
            If Mid(txtFileVersion, 1, 1) = "01" Or Mid(txtFileVersion, 1, 1) = "02" Then
                PayEffDate = Mid(TextLine, 401, 2) & "/" & Mid(TextLine, 403, 2) & "/" & Mid(TextLine, 405, 4)
                MemberDOB = Mid(TextLine, 320, 2) & "/" & Mid(TextLine, 322, 2) & "/" & Mid(TextLine, 324, 4)
            ElseIf Mid(txtFileVersion, 1, 1) = "03" Then
                PayEffDate = Mid(TextLine, 431, 2) & "/" & Mid(TextLine, 433, 2) & "/" & Mid(TextLine, 435, 4)
                MemberDOB = Mid(TextLine, 350, 2) & "/" & Mid(TextLine, 352, 2) & "/" & Mid(TextLine, 354, 4)
            End If
            FindAge = GetAge(CDate(PayEffDate), CDate(MemberDOB))
            AgeValue = Mid(GetAgeBand(Mid(txtFileVersion, 1, 2), FindAge), 1, 2)
            !MBMCAgeband = AgeValue
            If AgeValue = "01" Then
                !MBMCAdultChild = "C"
            Else
                !MBMCAdultChild = "A"
            End If
                !MBMCNumber = txtTempMemberNo
        End With
            MBMCover.Update
        MBMCover.Close
End Sub

Public Sub ListMembers()
    
End Sub

Public Sub ListEndtMembers()
    
End Sub

Private Function CheckIC(IC As String)

End Function
Private Sub ComboList()
    Dim HLT As New ADODB.Recordset
    Dim PAY As New ADODB.Recordset
    Dim strsql As String
    
    HLT.Open "HLT", medixmasterconn, adOpenKeyset, adLockOptimistic
    
    Do Until HLT.EOF
        cboHLTCode.AddItem HLT!HLTCode & " - " & HLT!HLTDescription
        HLT.MoveNext
    Loop
    
    strsql = "Select  PAYCode, PAYCompanyName From PAY where HLTCOde ='" & Mid(Trim(cboHLTCode), 1, 1) & "'"
    PAY.Open strsql, medixmasterconn, adOpenKeyset, adLockOptimistic
    cboPAYCode.Clear
    Do Until PAY.EOF
       cboPAYCode.AddItem ChkStr(PAY!PAYCode) & " - " & ChkStr(PAY!PAYCompanyName)
       PAY.MoveNext
    Loop
    PAY.Close
    Set PAY = Nothing

    cboExportOption.AddItem "D" & " - " & "Bordx Range"
    cboExportOption.AddItem "B" & " - " & "Batch Bordx"
    
    cboBordxOPT.AddItem "D" & " - " & "Bordx Range"
    cboBordxOPT.AddItem "B" & " - " & "Batch Bordx"
    
    HLT.Close
    Set HLT = Nothing
End Sub

Private Sub PrincipalASCII()
Dim strsql As String
Dim MBM As New ADODB.Recordset
Dim TotalSupplementary As Integer
Dim strsplOthers As String
Dim MBMOthers As New ADODB.Recordset
    Cnt = 0
    
    strsql = "SELECT * From MBM Where " & Exportsql
    MBM.Open strsql, medixmasterconn, adOpenKeyset, adLockOptimistic
    
    
    Do
    Do Until MBM.EOF
        For ii = 1 To MBM.recordCount
            If MBM!MBMMemberType = "P" Then
                DataString = l(MBM!MBMMemberType, 1)
                TotalPrincipal = TotalPrincipal + 1
                txtTotalPrincipals = TotalPrincipal
                MNumber = Trim(MBM!MBMNumber)
                INSCode = Trim(MBM!INSCode)
                MBMPolicyNo = Trim(MBM!MBMPolicyNo)
            strsplOthers = "SELECT * From MBMOthers Where MBMNumber = '" & Trim(MBM!MBMNumber) & "'"
            MBMOthers.Open strsplOthers, medixmasterconn, adOpenKeyset, adLockOptimistic
                If MBMOthers.recordCount Then
                    If Len(MBMOthers!MBMFirstJoinedDate) Then
                        MBMDateJoin = Format(MBMOthers!MBMFirstJoinedDate, "yymmdd")
                    End If
                    If Len(MBMOthers!MBMEmployeeNo) Then
                        MBMEmployeeNo = MBMOthers!MBMEmployeeNo
                    End If
                    If Len(MBMOthers!MBMGroupCompany) Then
                        MBMGroupCompany = MBMOthers!MBMGroupCompany
                    End If
                    If Len(MBMOthers!MBMPolSubNo1) Then
                        MBMPolSubNo1 = MBMOthers!MBMPolSubNo1
                    Else
                        MBMPolSubNo1 = ""
                    End If
                End If
            MBMOthers.Close
            Set MBMOthers = Nothing
                DoEvents
            
            End If
             
            DataString = DataString & l(MBM!MBMPolicyNo, 25)
            DataString = DataString & l(MBMEmployeeNo, 15)
            DataString = DataString & l(MBM!MBMName, 60)
            DataString = DataString & l(MBM!MBMIcBcPp, 20)
            DataString = DataString & l(MBM!MBMNumber, 18)
            DataString = DataString & l("00", 2)
            
            If Mid(cboASCIIFileVer, 1, 2) = "01" Then
                DataString = DataString & l(MBMGroupCompany, 40)
                DataString = DataString & l("", 6) & Chr(13)
            ElseIf Mid(cboASCIIFileVer, 1, 2) = "02" Then
                DataString = DataString & l(MBMDateJoin, 8)
                DataString = DataString & l(MBMGroupCompany, 40)
                'DataString = DataString & l(MBMPolSubNo1, 10)
                DataString = DataString & l("", 6) & Chr(13)
            ElseIf Mid(cboASCIIFileVer, 1, 2) = "03" Then
                DataString = DataString & l(MBMDateJoin, 8)
                DataString = DataString & l(MBMGroupCompany, 40)
                DataString = DataString & l("", 8)
                DataString = DataString & l("", 4)
                DataString = DataString & l("", 5)
                DataString = DataString & l("", 8)
                DataString = DataString & l("", 2)
                DataString = DataString & l("", 8)
                DataString = DataString & l("", 2)
                DataString = DataString & l(MBMPolSubNo1, 10)
                DataString = DataString & l("", 65) & Chr(13)
            End If
            
            
            MBM.MoveNext
            
            'Debug.Print DataString
            Print #1, DataString
            Print #2, DataString
            If intExit = 1 Then
                Check = False
                Exit Do
            End If
             
            If INSCode = "F" Or INSCode = "H" Then
                Call SupplementaryASCII
            End If

        Next ii
     
    Loop
    Loop Until Check = False
End Sub

Private Sub SupplementaryASCII()
Dim strsql As String
Dim MBMCoveredPersons As New ADODB.Recordset

    strsql = "SELECT * From MBMCoveredPersons  Where MBMCNumber = '" & Trim(MNumber) & "'"
    MBMCoveredPersons.Open strsql, medixmasterconn, adOpenKeyset, adLockOptimistic
       
       
    Do Until MBMCoveredPersons.EOF
        DataString = l(MBMCoveredPersons!MBMCRELCode, 1)
        TotalSupplementary = TotalSupplementary + 1
        txtTotalSupplementary1 = TotalSupplementary
        DoEvents
        
        DataString = DataString & l(MBMPolicyNo, 25)
        DataString = DataString & l(MBMEmployeeNo, 15)
        DataString = DataString & l(MBMCoveredPersons!MBMCName, 60)
        DataString = DataString & l(MBMCoveredPersons!MBMCIcBcPpNo, 20)
        DataString = DataString & l(MBMCoveredPersons!MBMCNumber, 18)
        DataString = DataString & l(MBMCoveredPersons!MBMCCoverID, 2)
        DataString = DataString & l("", 160) & Chr(13)
        MBMCoveredPersons.MoveNext

        Print #1, DataString
        Print #2, DataString
    Loop
End Sub

Private Sub UpdateMBMPriting()
Dim BordxPrinting As New ADODB.Recordset
Dim Frm As Form
    BordxPrinting.Open "BordxPrinting", medixmasterconn, adOpenKeyset, adLockOptimistic
        
    BordxPrinting.AddNew
    With BordxPrinting
    
        !MBMBordxDate = txtSubmissionDate
        !PAYCode = Mid(txtPAYCode, 1, 2)
        If Len(txtBatchNo) Then
            !MBMBatchNo = txtBatchNo
        End If
        If Mid(txtFileName, 1, 1) = "N" Then
            !MBMBordxType = "N"
        ElseIf Mid(txtFileName, 1, 1) = "E" Then
            !MBMBordxType = "E"
        End If
    End With
    BordxPrinting.Update
    BordxPrinting.Close
    Set BordxPrinting = Nothing
End Sub

Private Sub CISDataImport()
Dim CisWarranty As String
Dim CisPrem, CISBasicPrem, CisAgeValue As String
Dim CisMCO, CisBasicMco, CisVisitLmt, CisBalLmt As String
Dim MemberDOB, CISPayor As String
Dim FindAge As String
Dim AgeValue As String
        
        CLNOthers.AddNew
        With CLNOthers
            
                CISPayor = Mid(txtPAYCode, 1, 2)
                If Mid(txtFileVersion, 1, 2) <> "04" Then
                    MemberDOB = Mid(TextLine, 293, 2) & "/" & Mid(TextLine, 295, 2) & "/" & Mid(TextLine, 297, 4)
                Else
                    MemberDOB = Mid(TextLine, 323, 2) & "/" & Mid(TextLine, 325, 2) & "/" & Mid(TextLine, 327, 4)
                End If
                CisWarranty = Mid(TextLine, 6729, 2500)
                
                FindAge = GetAge(CDate(CisEffDate), CDate(MemberDOB))
                AgeValue = Mid(GetAgeBand(Mid(cboHLTCode1, 1, 1), FindAge), 1, 2)
                               
                CisPrem = Format(GetCLNPremium(CisInsType, CISPayor, AgeValue, CisPlnCode, Mid(cboHLTCode1, 1, 1), CDate(CisEffDate)), "#,##0.00")
                CISBasicPrem = CisPrem
                'Format(Round(txtMBMPremium, 0), "#,##0.00")
                CisPrem = InsertPremium(CDate(CisExpDate), CDate(CisEffDate), Format(CisPrem, "#,##0.00"))
                CisPrem = Format(Round(CisPrem, 0), "#,##0.00")
                If AgeValue = "01" Then
                    CisMCO = Format(GetCLNMCO(CisInsType, Mid(cboHLTCode1, 1, 1), "C", CDate(CisEffDate), CisPlnCode), "#,##0.00")
                Else
                    CisMCO = Format(GetCLNMCO(CisInsType, Mid(cboHLTCode1, 1, 1), "A", CDate(CisEffDate), CisPlnCode), "#,##0.00")
                End If
                CisBasicMco = CisMCO
                CisMCO = InsertMCOFees(CDate(CisExpDate), CDate(CisEffDate), Format(CisMCO, "#,##0.00"))
                CisMCO = Round(CisMCO)
                CisBalLmt = GetCLNAnnualLimit(CisInsType, CisPlnCode, Mid(cboHLTCode1, 1, 1), CDate(CisEffDate))
                CisVisitLmt = GetVisitLimit(CisInsType, CisPlnCode, Mid(cboHLTCode1, 1, 1), CDate(CisEffDate))
                If Mid(cboSuppAnnualLimit, 1, 1) = "C" Then
                    !MBMCLNMemberType = "P"
                Else
                    !MBMCLNMemberType = "Q"
                End If
                !MBMNumber = txtTempMemberNo
                !MBMPolNumber = Mid(TextLine, 2, 25)
                !MBMCLNStatus = "A"
                !MBMCLNDataStatus = "A"
                !MBMCLNExpiryStatus = "N"
                !PLNType = CisPlnCode
                !InsType = CisInsType
                !MBMPolEffDate = CisEffDate
                !MBMPolExpDate = CisExpDate
                !MBMCLNWarranty = CisWarranty
                !MBMCLNBasicPrem = CISBasicPrem
                !MBMCLNPremium = CisPrem
                !MBMCLNBasicMCO = CisBasicMco
                !MBMCLNMCOFee = CisMCO
                !MBMCLNBalLmt = CisBalLmt
                !MBMCLNVisitLmt = CisVisitLmt
'
''***** end of CIs Update
            End With
            CLNOthers.Update
End Sub

