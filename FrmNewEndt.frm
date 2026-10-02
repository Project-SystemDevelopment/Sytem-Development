VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form FrmNewEndt 
   Caption         =   "Upload Member"
   ClientHeight    =   7905
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   11625
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   7905
   ScaleWidth      =   11625
   WindowState     =   2  'Maximized
   Begin VB.Frame Frame4 
      Height          =   615
      Left            =   120
      TabIndex        =   72
      Top             =   7200
      Width           =   11175
      Begin VB.TextBox tempAnnLmt 
         Height          =   285
         Left            =   5400
         TabIndex        =   79
         Text            =   "Text1"
         Top             =   240
         Visible         =   0   'False
         Width           =   975
      End
      Begin VB.TextBox txtDateJoin 
         Height          =   375
         Left            =   0
         TabIndex        =   78
         Text            =   "Text1"
         Top             =   240
         Visible         =   0   'False
         Width           =   255
      End
      Begin VB.TextBox txtMBMPremium 
         Height          =   285
         Left            =   3240
         TabIndex        =   77
         Top             =   240
         Visible         =   0   'False
         Width           =   735
      End
      Begin VB.TextBox txtMBMMCOFees 
         Height          =   285
         Left            =   3960
         TabIndex        =   76
         Top             =   240
         Visible         =   0   'False
         Width           =   735
      End
      Begin VB.TextBox txtAccessFees 
         Height          =   285
         Left            =   4680
         TabIndex        =   75
         Top             =   240
         Visible         =   0   'False
         Width           =   735
      End
      Begin VB.TextBox txtMBMAdultChild 
         Height          =   285
         Left            =   5400
         TabIndex        =   74
         Top             =   240
         Visible         =   0   'False
         Width           =   615
      End
      Begin VB.TextBox txtTempMemberNo 
         Height          =   285
         Left            =   960
         TabIndex        =   73
         Text            =   "txtTempMemberNo"
         Top             =   240
         Visible         =   0   'False
         Width           =   375
      End
      Begin VB.CommandButton cmdStop 
         Caption         =   "Stop"
         Height          =   315
         Left            =   7440
         TabIndex        =   11
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
         Left            =   8400
         TabIndex        =   12
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
         Left            =   9480
         TabIndex        =   13
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
   Begin TabDlg.SSTab SSTab1 
      Height          =   6975
      Left            =   120
      TabIndex        =   14
      Top             =   240
      Width           =   11175
      _ExtentX        =   19711
      _ExtentY        =   12303
      _Version        =   393216
      Tabs            =   2
      TabsPerRow      =   2
      TabHeight       =   520
      TabCaption(0)   =   "Upload Bordx"
      TabPicture(0)   =   "FrmNewEndt.frx":0000
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "lstERRList"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Frame2"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "Frame1"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).ControlCount=   3
      TabCaption(1)   =   "Generate ASCII File"
      TabPicture(1)   =   "FrmNewEndt.frx":001C
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Frame8"
      Tab(1).ControlCount=   1
      Begin VB.Frame Frame1 
         Height          =   1815
         Left            =   120
         TabIndex        =   59
         Top             =   360
         Width           =   10935
         Begin VB.ComboBox cboVerifyIC 
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
            TabIndex        =   6
            Text            =   "Y - Yes"
            Top             =   240
            Width           =   1695
         End
         Begin VB.Frame Frame13 
            Height          =   615
            Left            =   5520
            TabIndex        =   61
            Top             =   1120
            Width           =   2175
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
               TabIndex        =   9
               Top             =   200
               Width           =   960
            End
            Begin VB.CommandButton cmdUpload 
               Caption         =   "&Upload"
               Enabled         =   0   'False
               Height          =   315
               Left            =   1080
               TabIndex        =   10
               Top             =   200
               Width           =   960
            End
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
            TabIndex        =   4
            Top             =   1395
            Width           =   320
         End
         Begin VB.TextBox txtClnStatus 
            Height          =   285
            Left            =   8880
            TabIndex        =   60
            Text            =   "N"
            Top             =   480
            Visible         =   0   'False
            Width           =   975
         End
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
            Top             =   180
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
            Top             =   480
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
            Top             =   780
            Width           =   3030
         End
         Begin VB.ComboBox CboBack2Back 
            Height          =   315
            Left            =   6000
            Sorted          =   -1  'True
            TabIndex        =   7
            Text            =   "N - No"
            Top             =   540
            Width           =   1695
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
            Left            =   6000
            Sorted          =   -1  'True
            Style           =   2  'Dropdown List
            TabIndex        =   8
            Top             =   830
            Width           =   1695
         End
         Begin VB.TextBox txtFileVersion 
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
            TabIndex        =   3
            Text            =   "Version 4800"
            Top             =   1080
            Width           =   3030
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
            TabIndex        =   5
            Top             =   1370
            Width           =   2625
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
            TabIndex        =   70
            Top             =   1140
            Width           =   960
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
            TabIndex        =   69
            Top             =   480
            Width           =   1170
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
            TabIndex        =   68
            Top             =   1440
            Width           =   975
         End
         Begin VB.Label Label5 
            Caption         =   "New/ Endt File"
            Height          =   255
            Left            =   240
            TabIndex        =   67
            Top             =   825
            Width           =   1095
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
            TabIndex        =   66
            Top             =   135
            Width           =   1770
         End
         Begin VB.Label Label13 
            Caption         =   "Verify I/C"
            Height          =   255
            Left            =   4800
            TabIndex        =   65
            Top             =   255
            Width           =   855
         End
         Begin VB.Label Label14 
            Caption         =   "Label Status"
            Height          =   255
            Left            =   4800
            TabIndex        =   64
            Top             =   840
            Width           =   975
         End
         Begin VB.Label Label15 
            Caption         =   "BTB Grantee"
            Height          =   255
            Left            =   4800
            TabIndex        =   63
            Top             =   570
            Width           =   1215
         End
         Begin VB.Label Label8 
            Caption         =   "Clinical Status"
            Height          =   375
            Left            =   8040
            TabIndex        =   62
            Top             =   480
            Visible         =   0   'False
            Width           =   855
         End
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
         TabIndex        =   48
         Top             =   6240
         Width           =   10785
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
            TabIndex        =   58
            Top             =   240
            Width           =   960
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
            TabIndex        =   57
            Top             =   240
            Width           =   960
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
            TabIndex        =   56
            Top             =   240
            Width           =   945
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
            TabIndex        =   55
            Top             =   240
            Width           =   1200
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
            TabIndex        =   54
            Top             =   240
            Width           =   1200
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
            TabIndex        =   53
            Top             =   240
            Width           =   1425
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
            TabIndex        =   52
            Top             =   240
            Width           =   960
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
            TabIndex        =   51
            Top             =   240
            Width           =   960
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
            TabIndex        =   50
            Top             =   240
            Width           =   495
         End
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
            TabIndex        =   49
            Top             =   240
            Width           =   720
         End
      End
      Begin VB.Frame Frame8 
         Height          =   3375
         Left            =   -74760
         TabIndex        =   25
         Top             =   360
         Width           =   10815
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
            TabIndex        =   43
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
               TabIndex        =   47
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
               Left            =   4920
               TabIndex        =   46
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
               TabIndex        =   45
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
               Left            =   2640
               TabIndex        =   44
               Top             =   240
               Width           =   2040
            End
         End
         Begin VB.Frame Frame14 
            Height          =   2535
            Left            =   120
            TabIndex        =   26
            Top             =   120
            Width           =   10455
            Begin VB.Frame BordxFrame 
               Height          =   600
               Left            =   120
               TabIndex        =   40
               Top             =   1800
               Visible         =   0   'False
               Width           =   9015
               Begin MSMask.MaskEdBox txtBordxTo 
                  Height          =   285
                  Left            =   5730
                  TabIndex        =   23
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
                  TabIndex        =   22
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
                  TabIndex        =   42
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
                  TabIndex        =   41
                  Top             =   225
                  Width           =   1275
               End
            End
            Begin VB.Frame BatchFrame 
               Height          =   600
               Left            =   120
               TabIndex        =   35
               Top             =   1800
               Width           =   9015
               Begin VB.TextBox txtBatchNo1 
                  Height          =   285
                  Left            =   5640
                  TabIndex        =   36
                  Top             =   180
                  Width           =   2500
               End
               Begin MSMask.MaskEdBox txtReceiptDate 
                  Height          =   285
                  Left            =   1530
                  TabIndex        =   37
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
                  TabIndex        =   39
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
                  TabIndex        =   38
                  Top             =   180
                  Width           =   675
               End
            End
            Begin VB.Frame Frame16 
               Height          =   1695
               Left            =   120
               TabIndex        =   27
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
                  Style           =   2  'Dropdown List
                  TabIndex        =   19
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
                  TabIndex        =   18
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
                  ItemData        =   "FrmNewEndt.frx":0038
                  Left            =   4800
                  List            =   "FrmNewEndt.frx":003A
                  Sorted          =   -1  'True
                  TabIndex        =   21
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
                  TabIndex        =   20
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
                  TabIndex        =   34
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
                  TabIndex        =   33
                  Top             =   240
                  Width           =   765
               End
               Begin VB.Label txtFilName 
                  BackColor       =   &H00FFFFFF&
                  Height          =   255
                  Left            =   8160
                  TabIndex        =   17
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
                  TabIndex        =   16
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
                  TabIndex        =   15
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
                  TabIndex        =   32
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
                  TabIndex        =   31
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
                  TabIndex        =   30
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
                  TabIndex        =   29
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
                  TabIndex        =   28
                  Top             =   240
                  Width           =   1125
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
               TabIndex        =   24
               Top             =   1920
               Width           =   960
            End
         End
      End
      Begin MSComctlLib.ListView lstERRList 
         Height          =   3975
         Left            =   120
         TabIndex        =   71
         Top             =   2280
         Width           =   10935
         _ExtentX        =   19288
         _ExtentY        =   7011
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
            Text            =   "Line"
            Object.Width           =   970
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
            Object.Width           =   8819
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   4
            Text            =   "Filename"
            Object.Width           =   2646
         EndProperty
      End
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      BackColor       =   &H00404040&
      Caption         =   "    Upload MBM"
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
      Height          =   220
      Left            =   0
      TabIndex        =   80
      Top             =   0
      Width           =   11415
   End
End
Attribute VB_Name = "FrmNewEndt"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim MBM As New ADODB.Recordset
Dim TextLine
Dim CoveredID As Integer
Dim CoveredIDCount As String
Dim VNewPayor, MNumber As String
Dim INSCode As String
Dim MBMPolicyNo As String
Dim MBMEmployeeNo As String
Dim PayEffDate As String
Dim PayorEXPDate As String
Dim InsuredCode As String
Dim plancode As String
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
Dim Check As Boolean
Dim MBMPayEffDate As String
Dim Exportsql As String
Dim DataString As String
Dim AnnLmt As New ADODB.Recordset
Dim CISCovSuppLMtStat As String
Dim CISDCStatus As String
Dim CISDCPayor As String
Dim CISPDentStatus As String
Dim CISSDentStatus As String
Dim CISPSpecStatus As String
Dim CISSSpecStatus As String
Dim CISPMepStatus As String
Dim CISSMepStatus As String
Dim WLifeStatus As Boolean
Dim PrevLifeMBMNo As String
Dim TmpTopupNo As String
Dim MBMCancelStatus As String
Dim tmpGuaRenewal As String
Dim tmpPLNLifeTimeStatus As String
Dim tmpEndtPrincEffDate As String
Dim tmpEndtPrincExpDate As String
Public intExit As Integer

Private Sub cboASCIIFileVer_KeyPress(KeyAscii As Integer)
    Call ComboKeyPress(cboASCIIFileVer, KeyAscii)
End Sub

Private Sub CboHLTCode_KeyPress(KeyAscii As Integer)
    Call ComboKeyPress(cboHLTCode, KeyAscii)
End Sub

Private Sub cboHLTCode_LostFocus()
Dim strAppend As String
    strAppend = "And HLTCOde = '" & Mid(cboHLTCode, 1, 1) & "'"
    Call LoadComboListing(strAppend, "PAY", cboPAYCode)
End Sub

Private Sub cboHLTCode1_LostFocus()
Dim strAppend As String
    strAppend = "And HLTCOde = '" & Mid(cboHLTCode1, 1, 1) & "'"
    Call LoadComboListing(strAppend, "PAY", txtPAYCode)
End Sub

Private Sub cboExportOption_Change()
    If InStr(cboExportOption.Text, "null") Then
       cboExportOption.Enabled = False
    End If
End Sub

Private Sub cboExportOption_KeyPress(KeyAscii As Integer)
    Call ComboKeyPress(cboExportOption, KeyAscii)
End Sub

Private Sub cboExportOption_LostFocus()
    If Mid(cboExportOption, 1, 1) = "B" Then
        BatchFrame.Visible = True
        BordxFrame.Visible = False
        txtReceiptDate.SetFocus
        txtReceiptDate.TabIndex = 13
        txtBatchNo1.TabIndex = 14
        cmdGenerate.TabIndex = 15
        txtFileName = ""
    ElseIf Mid(cboExportOption, 1, 1) = "D" Then
        BatchFrame.Visible = False
        BordxFrame.Visible = True
        txtBordxFrom.SetFocus
        txtBordxFrom.TabIndex = 13
        txtBordxTo.TabIndex = 14
        cmdGenerate.TabIndex = 15
        txtFileName = ""
    End If
End Sub

Private Sub CmdGenerate_Click()
Dim strsql As String
Dim TotalPrincipal As Integer
Dim TotalSupplementary As Integer
Dim RecordSaved
Dim MBMOthers As New ADODB.Recordset
Dim MBMCoveredPersons As New ADODB.Recordset
Dim ASCII As New ADODB.Recordset
Dim textField As String
Dim header As String
Dim Trailer As String
    
    medixmasterconn.Open "File Name=" & BOSPath & "Conn\mediconnect.UDL"
    
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
                           
'            strsql = "SELECT * From MBM Where " & Exportsql
            strsql = "SELECT MBMNumber, MBMMemberType, INSCode, MBMPolicyNo, MBMPayorEffDate, MBMIcBcPp, " & _
                    " MBMName From MBM Where " & Exportsql

            MBM.Open strsql, medixmasterconn, adOpenKeyset, adLockOptimistic

            If MBM.recordCount = 0 Then
                MsgBox "No Record Available to Export To ASCII File", vbOKOnly + vbCritical, DialogBoxTitle
                Screen.MousePointer = vbDefault
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
            MBM.Close
            Set MBM = Nothing
            Screen.MousePointer = vbDefault
        End If
    End If
    
    txtTotalPrincipals = l(txtTotalPrincipals, 10)
    txtTotalSupplementary1 = l(txtTotalSupplementary1, 10)
    medixmasterconn.Close
    Set medixmasterconn = Nothing
End Sub

Private Sub PrincipalASCII()
Dim TotalSupplementary As Integer
Dim strsplOthers As String
Dim MBMOthers As New ADODB.Recordset
Dim TotalPrincipal As Integer
Dim MBMOther2 As New ADODB.Recordset
Dim MBMOthers2Sql As String
Dim XPRiskNo As String
Dim XPTransNo As String
Dim XPPayNo As String
Dim XPpaySeqNo As String
Dim XPClientNo As String
Dim XPClientID As String
Dim XPPolDate As String
    
    Cnt = 0
    
    Do Until MBM.EOF
        If Trim(MBM!MBMMemberType) = "Q" Then
            DataString = l("P", 1)
        Else
            DataString = l(MBM!MBMMemberType, 1)
        End If
        
        TotalPrincipal = TotalPrincipal + 1
        txtTotalPrincipals = TotalPrincipal
        MNumber = Trim(MBM!MBMNumber)
        INSCode = Trim(MBM!INSCode)
        MBMPolicyNo = Trim(MBM!MBMPolicyNo)
        
        strsplOthers = "SELECT MBMEmployeeNo, MBMGroupCompany, MBMPolSubNo1 From MBMOthers Where MBMNumber = '" & Trim(MBM!MBMNumber) & "'"
        MBMOthers.Open strsplOthers, medixmasterconn, adOpenKeyset, adLockOptimistic
        If MBMOthers.recordCount Then
            If Len(MBM!MBMPayorEffDate) Then
                MBMPayEffDate = Format(MBM!MBMPayorEffDate, "DDMMYYYY")
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
        
        MBMOthers2Sql = "SELECT MBMPOLIssuedDate, MBMRiskNo, MBMTransNo, MBMPayNo, MBMPaySeqNo, MBMClientNo, MBMClientID " & _
                        " From MBMOthersTwo Where MBMNumber = '" & Trim(MBM!MBMNumber) & "'"
        MBMOther2.Open MBMOthers2Sql, medixmasterconn, adOpenKeyset, adLockOptimistic
        If Len(MBMOther2!MBMPOLIssuedDate) Then
            XPPolDate = Format(MBMOther2!MBMPOLIssuedDate, "DDMMYYYY")
        End If
        XPRiskNo = IIf(Len(Trim(MBMOther2!MBMRiskNo)), MBMOther2!MBMRiskNo, "")
        XPTransNo = IIf(Len(Trim(MBMOther2!MBMTransNo)), MBMOther2!MBMTransNo, "")
        XPPayNo = IIf(Len(Trim(MBMOther2!MBMPayNo)), MBMOther2!MBMPayNo, "")
        XPpaySeqNo = IIf(Len(Trim(MBMOther2!MBMPaySeqNo)), MBMOther2!MBMPaySeqNo, "")
        XPClientNo = IIf(Len(Trim(MBMOther2!MBMClientNo)), MBMOther2!MBMClientNo, "")
        XPClientID = IIf(Len(Trim(MBMOther2!MBMClientID)), MBMOther2!MBMClientID, "")
        
        MBMOther2.Close
        Set MBMOther2 = Nothing
        DoEvents
        
        DataString = DataString & l(MBM!MBMPolicyNo, 25)
        DataString = DataString & l(MBMEmployeeNo, 15)
        DataString = DataString & l(MBM!MBMName, 60)
        DataString = DataString & l(MBM!MBMICBcPp, 20)
        DataString = DataString & l(MBM!MBMNumber, 18)
        
        DataString = DataString & l("00", 2)
        DataString = DataString & l(MBMPayEffDate, 8)
        DataString = DataString & l(MBMGroupCompany, 40)
        DataString = DataString & l(XPPolDate, 8)
        DataString = DataString & l(XPRiskNo, 4)
        DataString = DataString & l(XPTransNo, 5)
        DataString = DataString & l(XPPayNo, 8)
        DataString = DataString & l(XPpaySeqNo, 2)
        DataString = DataString & l(XPClientNo, 8)
        DataString = DataString & l(XPClientID, 2)
        DataString = DataString & l(MBMPolSubNo1, 10)
        DataString = DataString & l("", 64) & Chr(13)
            
        
        Print #1, DataString
        Print #2, DataString
        
        If intExit = 1 Then
            Check = False
            Exit Do
        End If
        If INSCode = "F" Or INSCode = "H" Then
            Call SupplementaryASCII
        End If
        
        MBM.MoveNext
    Loop
End Sub

Private Sub SupplementaryASCII()
Dim strsql As String
Dim MBMCoveredPersons As New ADODB.Recordset
Dim TotalSupplementary As Integer
Dim MBMCov2 As New ADODB.Recordset
Dim MBMCovSql As String
Dim XCPayNo As String
Dim XCpaySeqNo As String
Dim XCClientNo As String
Dim XCClientID As String

    strsql = "SELECT MBMCRELCode, MBMCName,MBMCIcBcPpNo,MBMCNumber,MBMCCoverID From MBMCoveredPersons  Where MBMCNumber = '" & Trim(MNumber) & "'"
    MBMCoveredPersons.Open strsql, medixmasterconn, adOpenKeyset, adLockOptimistic
    
    MBMCovSql = "SELECT  MBMCPayNo, MBMCPaySeqNo, MBMCClientNo, MBMCClientID " & _
              " From MBMCoveredPersonsTwo  Where MBMCNumber = '" & Trim(MNumber) & "'"
    MBMCov2.Open MBMCovSql, medixmasterconn, adOpenKeyset, adLockOptimistic
              
    Do Until MBMCoveredPersons.EOF
        DataString = l(MBMCoveredPersons!MBMCRELCode, 1)
        txtTotalSupplementary1 = txtTotalSupplementary1 + 1
        
        XCPayNo = IIf(Len(Trim(MBMCov2!MBMCPayNo)), MBMCov2!MBMCPayNo, "")
        XCpaySeqNo = IIf(Len(Trim(MBMCov2!MBMCPaySeqNo)), MBMCov2!MBMCPaySeqNo, "")
        XCClientNo = IIf(Len(Trim(MBMCov2!MBMCClientNo)), MBMCov2!MBMCClientNo, "")
        XCClientID = IIf(Len(Trim(MBMCov2!MBMCClientID)), MBMCov2!MBMCClientID, "")

        DoEvents
        
        DataString = DataString & l(MBMPolicyNo, 25)
        DataString = DataString & l(MBMEmployeeNo, 15)
        DataString = DataString & l(MBMCoveredPersons!MBMCName, 60)
        DataString = DataString & l(MBMCoveredPersons!MBMCICBCPPNo, 20)
        DataString = DataString & l(MBMCoveredPersons!MBMCNumber, 18)
        DataString = DataString & l(Format(MBMCoveredPersons!MBMCCoverID, "00"), 2)
        DataString = DataString & l(MBMPayEffDate, 8)
        DataString = DataString & l("", 40)
        DataString = DataString & l("", 8)
        DataString = DataString & l("", 4)
        DataString = DataString & l("", 5)
        DataString = DataString & l(XCPayNo, 8)
        DataString = DataString & l(XCpaySeqNo, 2)
        DataString = DataString & l(XCClientNo, 8)
        DataString = DataString & l(XCClientID, 2)
        DataString = DataString & l(MBMPolSubNo1, 10)
        DataString = DataString & l("", 64) & Chr(13)
        DataString = DataString & l("", 6) & Chr(13)
        MBMCoveredPersons.MoveNext

        Print #1, DataString
        Print #2, DataString
    Loop
    MBMCoveredPersons.Close
    Set MBMCoveredPersons = Nothing
    MBMCov2.Close
    Set MBMCov2 = Nothing
End Sub

Private Sub CboBack2Back_KeyPress(KeyAscii As Integer)
    Call ComboKeyPress(CboBack2Back, KeyAscii)
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

Private Sub cboHLTCode1_KeyPress(KeyAscii As Integer)
    Call ComboKeyPress(cboHLTCode1, KeyAscii)
End Sub

Private Sub CboVerifyIC_Change()
    If InStr(cboVerifyIC.Text, "null") Then
       cboVerifyIC.Enabled = False
    End If
End Sub

Private Sub CboVerifyIC_KeyPress(KeyAscii As Integer)
    Call ComboKeyPress(cboVerifyIC, KeyAscii)
End Sub

Private Sub CmdExit_Click()
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
  CommonDialog1.ShowOpen
  txtImportFilename = CommonDialog1.FileTitle
  Exit Sub
  
ErrHandler:
  Exit Sub
End Sub

Private Sub cmdPrint_Click()
    Screen.MousePointer = vbHourglass
        
    If lstERRList.listitems.count <> 0 Then
        Call ExportListViewtoExcel(lstERRList)
    Else
        MsgBox "Report cannot be printed without any record.", vbInformation
    End If
        
    Screen.MousePointer = vbDefault
End Sub

Private Sub cmdStop_Click()
    intExit = 1
End Sub

Private Sub CmdUpload_Click()
Dim RecordSaved
    
    If Trim(cboHLTCode1) = "" Then
        MsgBox "Please Select Health Type", vbOKOnly + vbCritical, DialogBoxTitle
        txtFileName.SetFocus
    ElseIf Trim(txtFileName) = "" Then
        MsgBox "Please Select New / Endt file", vbOKOnly + vbCritical, DialogBoxTitle
        txtFileName.SetFocus
    ElseIf Trim(txtPAYCode) = "" Then
        MsgBox "Please Enter Payor Code", vbOKOnly + vbCritical, DialogBoxTitle
        txtPAYCode.SetFocus
    ElseIf Trim(txtImportFilename) = "" Then
        MsgBox "Please Enter Filename", vbOKOnly + vbCritical, DialogBoxTitle
    Else
        RecordSaved = MsgBox("Do you want to upload this file?", vbYesNo + vbExclamation)
        If RecordSaved = vbYes Then
            Screen.MousePointer = vbHourglass
            cmdUpload.Enabled = False
            cmdExit.Enabled = False
            medixmasterconn.Open "File Name=" & BOSPath & "Conn\mediconnect.UDL"
            medixmasterconnMaint.Open "File Name=" & BOSPath & "Conn\MediConnHISMaint.UDL"
            MediConnCISDB.Open "File Name=" & BOSPath & "Conn\MediConnCISDB.UDL"

            If Mid(txtFileName, 1, 1) = "N" Then
                Call GenerateMember
            ElseIf Mid(txtFileName, 1, 1) = "E" Then
                Call GenerateENDtMember
                cmdExit.Enabled = True
            End If
            
            medixmasterconn.Close
            Set medixmasterconn = Nothing
            medixmasterconnMaint.Close
            Set medixmasterconnMaint = Nothing
            MediConnCISDB.Close
            Set MediConnCISDB = Nothing
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
    If Mid(txtFileName, 1, 1) = "" Then
        MsgBox "Please Select Bordx Type", vbOKOnly + vbCritical, DialogBoxTitle
        txtFileName.SetFocus
    ElseIf txtPAYCode = "" Then
        MsgBox "Please Enter Payor Code", vbOKOnly + vbCritical, DialogBoxTitle
        txtPAYCode.SetFocus
    ElseIf txtFileName = "" Then
        MsgBox "Please Enter File Type", vbOKOnly + vbCritical, DialogBoxTitle
        txtFileName.SetFocus
    ElseIf txtImportFilename = "" Then
        MsgBox "Please Enter Filename", vbOKOnly + vbCritical, DialogBoxTitle
    Else
        
        cmdVerify.Enabled = False
        cmdUpload.Enabled = False
        lstERRList.listitems.Clear

        If Mid(txtFileName, 1, 1) = "N" Then
            Call VerifyNewMember
        ElseIf Mid(txtFileName, 1, 1) = "E" Then
            Call VerifyENDtMember
        End If
        
        If intExit = 0 Then
            MsgBox "Verifed Sucessfully...", vbOKOnly + vbInformation, DialogBoxTitle
            Screen.MousePointer = vbDefault
        Else
            Screen.MousePointer = vbDefault
        End If
    End If
    cmdVerify.Enabled = True
    cmdUpload.Enabled = True
    cmdFileLocation.Enabled = True
    intExit = 0
End Sub

Private Sub Form_Load()
    Call ComboListing
    Call LoadComboListing("", "HLT", cboHLTCode)
    Call LoadComboListing("", "HLT", cboHLTCode1)
    Call LoadComboListing("", "PAY", cboPAYCode)
    Call LoadComboListing("", "PAY", txtPAYCode)
    cboHLTCode1.ListIndex = 1
    intExit = 0
    txtPath = SRVPath & "ASCI"
    txtLocalFileLocation = "C:\"
End Sub

'---------------------Start ComboLisiting--------------------------
Private Sub ComboListing()
    txtFileName.AddItem "N" & " - " & "New File"
    txtFileName.AddItem "E" & " - " & "Endorsement File"
    
    cboASCIIFileVer.AddItem "01" & " - " & "Version 185"
    cboASCIIFileVer.AddItem "02" & " - " & "Version 195"
    cboASCIIFileVer.AddItem "03" & " - " & "Version 300"

    cboVerifyIC.AddItem "Y" & " - " & "Yes"
    cboVerifyIC.AddItem "N" & " - " & "No"
    cboVerifyIC.Text = "Y" & " - " & "Yes"
        
    txtFileVersion = "Version 4800"
    
    cboLBLStatus.AddItem "Y" & " - " & "YES"
    cboLBLStatus.AddItem "N" & " - " & "NO"
    
    CboBack2Back.AddItem "Y" & " - " & "Yes"
    CboBack2Back.AddItem "N" & " - " & "No"
    CboBack2Back.Text = "N" & " - " & "No"
    
    cboExportOption.AddItem "D" & " - " & "Bordx Range"
    cboExportOption.AddItem "B" & " - " & "Batch Bordx"

End Sub

Private Sub SSTab1_Click(PreviousTab As Integer)
    If SSTab1.Tab = 0 Then
        cmdStop.Visible = True
        CmdPrint.Visible = True
    Else
        cmdStop.Visible = False
        CmdPrint.Visible = False
        cmdExit.Visible = True
    End If
End Sub

'------------------------Start selection of related Info----------------
Private Sub txtFileName_KeyPress(KeyAscii As Integer)
    Call ComboKeyPress(txtFileName, KeyAscii)
End Sub

Private Sub txtFileName_LostFocus()
    If Mid(txtFileName, 1, 1) = "N" Then
        txtFileVersion = "Version 4800"
        CboBack2Back.Enabled = True
    ElseIf Mid(txtFileName, 1, 1) = "E" Then
        txtFileVersion = "Version 4830"
        CboBack2Back.Enabled = False
    End If
End Sub
'---------------------Start Verification of New File--------------------------
Private Sub VerifyNewMember()
Dim TotalPrincipalCount As Integer
Dim totalSuppCount As Integer
Dim Check As Boolean
Dim VNewIC As String
Dim VNewDOB As String
Dim VNewSex As String
Dim VNewPostCode As String
Dim VNewState As String
Dim VNewCity As String
Dim VNewEffDate As String
Dim VNewExpDate As String
Dim VNewPlanCode As String
Dim VNewInsured As String
Dim VNewDataType As String
Dim VRenewal As String
Dim vNewName As String
Dim VNewPolicyNo As String
Dim VNewCLnPlan As String
Dim VNewCLnIns As String
Dim VNewClnEffDate As String
Dim VNewClnExpDate As String
Dim MBMICArray() As String
Dim MBMICCnt As Integer
Dim ICCount As Integer
Dim XrefSql1 As String
Dim XrefSql2 As String
Dim XrefDB As New ADODB.Recordset
Dim PrincPayEffdate, PrincPayExpDate As String
Dim PrincPlanCode, tmpPlanCode As String
Dim tmpSuppEffDate As String
Dim tmpSuppExpDate As String
Dim NoOfRecLine As Integer
Dim tmpTableName As String
Dim TmpRecFound As Boolean
Dim SelFields As String
Dim SelCriteria As String
Dim tmpCode As String
Dim tmpDesc As String

    lstERRList.listitems.Clear
    Erase MBMICArray
    ReDim MBMICArray(0)
    MBMICCnt = 0
       
    Screen.MousePointer = vbHourglass
    Open txtImportFilename For Input As #1   ' Open file.
    Do
    Do While Not EOF(1)   ' Loop until end of file.
       Line Input #1, TextLine   ' Read line into variable.
        NoOfRecLine = NoOfRecLine + 1
            If NoOfRecLine = 1 Then
                txtSubmissionDate = Format(Mid(TextLine, 4, 2) & "/" & Mid(TextLine, 6, 2) & "/" & Mid(TextLine, 8, 4), "dd/mm/yyyy")
                If Mid(TextLine, 1, 1) <> "F" Then
                    Call ErrorListing(NoOfRecLine, "", "", "1st Record must be header with Data Type 'F'")
                End If
                If Trim(Mid(TextLine, 2, 2)) = "" Or Trim(Mid(TextLine, 2, 2)) <> Mid(txtPAYCode, 1, 2) Then
                    Call ErrorListing(NoOfRecLine, "", "", "Invalid Insurer's Code " & Trim(Mid(TextLine, 2, 2)))
                End If
                If IsDate(txtSubmissionDate) = False Then
                    Call ErrorListing(NoOfRecLine, "", "", "Invalid Submission Date !")
                End If
                VNewPayor = Mid(TextLine, 2, 2)
            End If
                             
            If NoOfRecLine >= 2 And Mid(TextLine, 1, 1) <> "T" Then
                VNewDataType = Trim(Mid(TextLine, 1, 1))
                VNewPolicyNo = Trim(Mid(TextLine, 2, 25))
                vNewName = Trim(Mid(TextLine, 47, 60))
                VNewCity = Trim(Mid(TextLine, 197, 20))
                VNewPostCode = Trim(Mid(TextLine, 217, 5))
                VNewState = Trim(Mid(TextLine, 222, 15))
                VNewIC = Trim(Mid(TextLine, 273, 20))
                VNewDOB = Mid(TextLine, 293, 2) & "/" & Mid(TextLine, 295, 2) & "/" & Mid(TextLine, 297, 4)
                VNewSex = Trim(Mid(TextLine, 301, 1))
                VNewEffDate = Mid(TextLine, 374, 2) & "/" & Mid(TextLine, 376, 2) & "/" & Mid(TextLine, 378, 4)
                VNewExpDate = Mid(TextLine, 382, 2) & "/" & Mid(TextLine, 384, 2) & "/" & Mid(TextLine, 386, 4)
                VNewPlanCode = Trim(Mid(TextLine, 390, 4))
                VNewInsured = Trim(Mid(TextLine, 394, 1))
                VRenewal = Trim(Mid(TextLine, 435, 1))
                If VNewDataType = "P" Then
                    PrincPayEffdate = Mid(TextLine, 374, 2) & "/" & Mid(TextLine, 376, 2) & "/" & Mid(TextLine, 378, 4)
                    PrincPayExpDate = Mid(TextLine, 382, 2) & "/" & Mid(TextLine, 384, 2) & "/" & Mid(TextLine, 386, 4)
                    PrincPlanCode = Trim(Mid(TextLine, 390, 4))
                End If
                If Trim(txtClnStatus) = "Y" And Trim(VNewPlanCode) <> "" Then ' Current payor is for Clinical and NO HIS
                    Call ErrorListing(NoOfRecLine, VNewPolicyNo, vNewName, "Invalid Plan Code for 'DC' !! ")
                End If
                If Mid(cboVerifyIC, 1, 1) = "Y" Then
                    If VNewIC <> "" Then
                        XrefSql1 = "Select MBMNumber, MBMICBCPP "
                        XrefSql2 = " And MBMICBCPP = '" & Trim(VNewIC) & "' And MBMStatus = 'A'"
                        Call SPMBMXref(NoOfRecLine, VNewPolicyNo, vNewName, VNewInsured, XrefSql1, XrefSql2, "1")
                    End If
                End If
' ************* Chk 4 Duplicate Rec in ascii file
                If UBound(MBMICArray) > 0 Then
                    For ICCount = 0 To UBound(MBMICArray)
                        If Trim(MBMICArray(ICCount)) = Trim(VNewIC) Then
                            Call ErrorListing(NoOfRecLine, VNewPolicyNo, vNewName, "Same IC Number Found in upload File! ")
                        End If
                    Next ICCount
                End If
                MBMICCnt = MBMICCnt + 1
                ReDim Preserve MBMICArray(MBMICCnt)
                MBMICArray(MBMICCnt) = VNewIC
' ************* End of Duplicate Rec. checking

' ------------- Checking for Principal & Supplementary, Name, IC, DOB, Sex
                If vNewName = "" Then
                    Call ErrorListing(NoOfRecLine, VNewPolicyNo, vNewName, "Name Is Blank")
                End If
                If VNewIC = "" Then
                    Call ErrorListing(NoOfRecLine, VNewPolicyNo, vNewName, "IC,BC or PP Is Blank")
                End If
                If IsDate(VNewDOB) = False Then
                    Call ErrorListing(NoOfRecLine, VNewPolicyNo, vNewName, "Invalid Date format ~ Date Of Birth")
                End If
                If VNewSex = "" Then
                    Call ErrorListing(NoOfRecLine, VNewPolicyNo, vNewName, "Sex Is Blank")
                End If
' ------------- End of Principal & Supplementary, Name, IC, DOB, Sex Checking
                If VNewDataType = "P" Then '---------- P - Principal
                    If Trim(Mid(TextLine, 107, 90)) = "" Then ' Add1+Add2+Add3 for Princ only
                        Call ErrorListing(NoOfRecLine, VNewPolicyNo, vNewName, "Address Is Blank")
                    End If
                    If VNewPostCode = "" Then
                        Call ErrorListing(NoOfRecLine, VNewPolicyNo, vNewName, "Post Code Is Blank")
                    End If
                    If VNewState = "" Then
                        Call ErrorListing(NoOfRecLine, VNewPolicyNo, vNewName, "State Is Blank")
                    End If
                    If VNewCity = "" Then
                        Call ErrorListing(NoOfRecLine, VNewPolicyNo, vNewName, "City Is Blank")
                    End If
                    If Trim(VNewPlanCode) = "" Then
                        Call ErrorListing(NoOfRecLine, VNewPolicyNo, vNewName, "HIS Plan Code is null")
                    Else
                        If Trim(VNewPolicyNo) = "" Then  ' HIS cannot be blank in Policy No.
                            Call ErrorListing(NoOfRecLine, VNewPolicyNo, vNewName, "Policy No Is Blank")
                        End If
                    End If
                        
                    SelFields = "SELECT RenewalType as tmpCode, RenewalDesc as tmpDesc "
                    SelCriteria = " And RenewalType ='" & Trim(VRenewal) & "'"
                    tmpTableName = "Renewal"
                    tmpCode = ""
                    tmpDesc = ""
                    TmpRecFound = False
                    Call SPMaintTableRecVerify(SelFields, tmpTableName, SelCriteria, TmpRecFound, tmpCode, tmpDesc)
                    If TmpRecFound = False Then
                        Call ErrorListing(NoOfRecLine, VNewPolicyNo, vNewName, "No such Renewal Code - '" & Trim(VRenewal) & "'")
                    End If
                    
                    SelFields = "SELECT PLNCode as tmpCode, PLNTopUpStatus as tmpDesc "
                    SelCriteria = " And PLNCode ='" & Trim(VNewPlanCode) & "'"
                    tmpTableName = "PLN"
                    tmpCode = ""
                    tmpDesc = ""
                    TmpRecFound = False
                    Call SPHISTableRecVerify(SelFields, tmpTableName, SelCriteria, TmpRecFound, tmpCode, tmpDesc)
                    If TmpRecFound = False Then
                        Call ErrorListing(NoOfRecLine, VNewPolicyNo, vNewName, "No such Plan Code - ' " & Trim(VNewPlanCode) & "' !")
                    Else
                        '-------- check for Ins Type, Eff Date, ExpDAte
                        Call VerifyMBMByPassingPA(NoOfRecLine, "HIS", VNewPolicyNo, vNewName, VNewDataType, VNewPlanCode, VNewInsured, VNewEffDate, VNewExpDate)
                        If Trim(tmpDesc) = "Y" Then  ' ***------- Check for Member Topup
                            XrefSql1 = "Select Top 1 MBMNumber, INSCode, MBMTopupNumber "
                            XrefSql2 = " AND MBMICBCPP = '" & Trim(VNewIC) & _
                                       "' And MBMStatus = 'A' AND PayCode ='" & VNewPayor & "' AND INSCode='" & VNewInsured & _
                                       "' ORDER BY MBMPayorEffDate DESC, MBMNumber "
                            Call SPMBMXref(NoOfRecLine, VNewPolicyNo, vNewName, VNewInsured, XrefSql1, XrefSql2, "2")
                            ' ------------------------ ***------ End of Checking for Topup
                        End If
                    End If
                Else '-----------------VNewDataType <> "P", Supplementary
                    If Trim(VNewInsured) = "I" Or Trim(VNewInsured) = "G" Then
                        Call ErrorListing(NoOfRecLine, VNewPolicyNo, vNewName, "Wrong Isured Type fpr Supplementary - '" & Trim(VNewInsured) & "'")
                    End If

                    SelFields = "SELECT RELCode as tmpCode, RELDescription as tmpDesc "
                    SelCriteria = " And RELCode ='" & Trim(VNewDataType) & "'"
                    tmpTableName = "REL"
                    tmpCode = ""
                    tmpDesc = ""
                    TmpRecFound = False

                    Call SPHISTableRecVerify(SelFields, tmpTableName, SelCriteria, TmpRecFound, tmpCode, tmpDesc)
                    If TmpRecFound = False Then
                        Call ErrorListing(NoOfRecLine, VNewPolicyNo, vNewName, "No Such Data Type - ' " & Trim(VNewDataType) & "' !")
                    Else
                        If Trim(VNewPlanCode) <> "" Then   ' HIS Supplementary Checking
                            tmpPlanCode = Trim(VNewPlanCode)
                            tmpSuppEffDate = VNewEffDate
                            tmpSuppExpDate = VNewExpDate
                            If Trim(tmpPlanCode) = "" Then
                                tmpPlanCode = Trim(PrincPlanCode)
                            End If
                            If IsDate(tmpSuppEffDate) = False Then
                                tmpSuppEffDate = PrincPayEffdate
                            End If
                            If IsDate(tmpSuppExpDate) = False Then
                                tmpSuppExpDate = PrincPayExpDate
                            End If
                            If IsDate(VNewDOB) Then
                                Call GetSuppAgeLmt(NoOfRecLine, VNewDOB, "HIS", VNewPolicyNo, vNewName, VNewDataType, tmpPlanCode, VNewInsured, CDate(tmpSuppEffDate), tmpSuppExpDate)
                            End If
                        End If
                    End If
                End If

' ************* CIS Varification
                VNewCLnPlan = Trim(Mid(TextLine, 6708, 4))
                VNewCLnIns = Trim(Mid(TextLine, 6712, 1))
                VNewClnEffDate = Mid(TextLine, 6713, 2) & "/" & Mid(TextLine, 6715, 2) & "/" & Mid(TextLine, 6717, 4)
                VNewClnExpDate = Mid(TextLine, 6721, 2) & "/" & Mid(TextLine, 6723, 2) & "/" & Mid(TextLine, 6725, 4)
                If Trim(VNewCLnPlan) = "" And VNewDataType = "P" Then
                    Call ErrorListing(NoOfRecLine, VNewPolicyNo, vNewName, "Clinical Plan Is NULL ! ")
                End If
                If Trim(txtClnStatus) = "Y" And VNewDataType = "P" And Trim(VNewCLnPlan) = "" Then
                    Call ErrorListing(NoOfRecLine, VNewPolicyNo, vNewName, "Clinical Plan Code cannot be null !! ")
                End If
                If Trim(VNewCLnPlan) <> "" Then
                    Call VerifyMBMByPassingPA(NoOfRecLine, "CIS", VNewDataType, " ", vNewName, VNewCLnPlan, VNewCLnIns, VNewClnEffDate, VNewClnExpDate)
                    If Trim(Mid(TextLine, 9229, 20)) = "" Then
                        Call ErrorListing(NoOfRecLine, VNewPolicyNo, vNewName, "Clinical Card Type is blank !! ")
                    End If
                End If
' ************* End of CIS verify
            End If ' NoOfRecLine >= 2 And Mid(TextLine, 1, 1) <> "T"
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
    NoOfRecLine = 0
End Sub
'---------------------End Verification of New File--------------------------

'---------------------Start Verification of ENDt File-----------------------
Private Sub VerifyENDtMember()
Dim TotPrincipal As Integer
Dim TotSupplementary As Integer
Dim VEndtIC, VEndtDOB, VEndtSuppID As String
Dim VEndtPostCode, VEndtState, VEndtCity As String
Dim VEndtEffDate As String
Dim VEndtExpDate As String
Dim VEndtPlanCode As String
Dim VEndtInsured As String
Dim VEndtDataType As String
Dim VEndtPolicyNo As String
Dim vEndtName As String
Dim VEndtCLnPlan, VEndtCLnIns, VEndtClnEffDate, VEndtClnExpDate As String
Dim NoOfERecLine As Integer
Dim tmpTableName As String
Dim TmpRecFound As Boolean
Dim SelFields As String
Dim SelCriteria As String
Dim tmpCode As String
Dim tmpDesc As String
Dim XrefSql1 As String
Dim XrefSql2 As String

    lstERRList.listitems.Clear
    Screen.MousePointer = vbHourglass
    Open txtImportFilename For Input As #1   ' Open file.
    Do While Not EOF(1)   ' Loop until end of file.
       Line Input #1, TextLine   ' Read line into variable.
        NoOfERecLine = NoOfERecLine + 1
            If NoOfERecLine = 1 Then
                txtSubmissionDate = Mid(TextLine, 4, 2) & "/" & Mid(TextLine, 6, 2) & "/" & Mid(TextLine, 8, 4)
                If Mid(TextLine, 1, 1) <> "E" Then
                    Call ErrorListing(NoOfERecLine, "", "", "1st Record must be header with Data Type 'F'")
                End If
                If Trim(Mid(TextLine, 2, 2)) = "" Or Trim(Mid(TextLine, 2, 2)) <> Mid(txtPAYCode, 1, 2) Then
                    Call ErrorListing(NoOfERecLine, "", "", "Invalid Insurer's Code " & Trim(Mid(TextLine, 2, 2)))
                End If
                If IsDate(txtSubmissionDate) = False Then
                    Call ErrorListing(NoOfERecLine, "", "", "Invalid Date" & Trim(Mid(TextLine, 2, 2)))
                End If
            End If
                             
            If NoOfERecLine >= 2 Then
                If Mid(TextLine, 10, 1) = "P" Then
                    TotPrincipal = TotPrincipal + 1
                ElseIf Mid(TextLine, 10, 1) <> "P" And Mid(TextLine, 1, 1) <> "T" Then
                    TotSupplementary = TotSupplementary + 1
                End If
            End If
            
            If NoOfERecLine >= 2 And Mid(TextLine, 1, 1) <> "T" Then
                VEndtPolicyNo = Mid(TextLine, 29, 25)
                vEndtName = Mid(TextLine, 74, 60)
                VEndtDataType = Mid(TextLine, 10, 1)
                If IsDate(Mid(TextLine, 2, 2) & "/" & Mid(TextLine, 4, 2) & "/" & Mid(TextLine, 6, 4)) = False Then
                    Call ErrorListing(NoOfERecLine, VEndtPolicyNo, vEndtName, "Invalid date format ~ Endorsement Effective Date")
                End If
                
                If Trim(VEndtDataType) = "" Then
                    Call ErrorListing(NoOfERecLine, VEndtPolicyNo, vEndtName, "Data Type is blank")
                End If
                
                If Trim(Mid(TextLine, 11, 18)) = "" Then
                    Call ErrorListing(NoOfERecLine, VEndtPolicyNo, vEndtName, "Membership No Is Blank")
                End If
                If Trim(vEndtName) = "" Then
                    Call ErrorListing(NoOfERecLine, VEndtPolicyNo, vEndtName, "Name Is Blank")
                End If
                If Trim(Mid(TextLine, 1, 1)) = "C" And Trim(VEndtDataType) = "P" Then
                    SelFields = "Select CASNumber as tmpCode, MBMNumber as tmpDesc "
                    SelCriteria = " and CASStatus='O' AND MBMNumber='" & Trim(Mid(TextLine, 11, 18)) & "'"
                    tmpTableName = "CASCrossReference"
                    tmpCode = ""
                    tmpDesc = ""
                    TmpRecFound = False
                    Call SPMaintTableRecVerify(SelFields, tmpTableName, SelCriteria, TmpRecFound, tmpCode, tmpDesc)
                    If TmpRecFound = True Then
                         Call ErrorListing(NoOfERecLine, VEndtPolicyNo, vEndtName, "There is at least an outstanding case under this member!")
                    End If
                End If
                If Trim(Mid(TextLine, 1, 1)) <> "C" Then ' No verification for cancellation
                    VEndtCity = Mid(TextLine, 224, 20)
                    VEndtPostCode = Mid(TextLine, 244, 5)
                    VEndtState = Mid(TextLine, 249, 15)
                    VEndtIC = Mid(TextLine, 300, 20)
                    VEndtDOB = Mid(TextLine, 320, 2) & "/" & Mid(TextLine, 322, 2) & "/" & Mid(TextLine, 324, 4)
                    VEndtEffDate = Mid(TextLine, 401, 2) & "/" & Mid(TextLine, 403, 2) & "/" & Mid(TextLine, 405, 4)
                    VEndtExpDate = Mid(TextLine, 409, 2) & "/" & Mid(TextLine, 411, 2) & "/" & Mid(TextLine, 413, 4)
                    VEndtPlanCode = Mid(TextLine, 417, 4)
                    VEndtInsured = Mid(TextLine, 421, 1)
                    If Trim(VEndtPlanCode) <> "" Then
                        If Trim(VEndtPolicyNo) = "" Then
                            Call ErrorListing(NoOfERecLine, VEndtPolicyNo, vEndtName, "Policy No Is Blank")
                        End If
                    End If
                    If Mid(cboVerifyIC, 1, 1) = "Y" Then
                        If Trim(VEndtIC) <> "" Then
                            XrefSql1 = "Select MBMNumber, MBMICBCPP "
                            XrefSql2 = " And MBMICBCPP = '" & Trim(VEndtIC) & "' And MBMStatus = 'A'"
                            Call SPMBMXref(NoOfERecLine, VEndtPolicyNo, vEndtName, VEndtInsured, XrefSql1, XrefSql2, "1")
                        End If
                    End If
                    
                    If Trim(VEndtIC) = "" Then
                        Call ErrorListing(NoOfERecLine, VEndtPolicyNo, vEndtName, "IC/BC/PP Is Blank")
                    End If
                    
                    If IsDate(Trim(VEndtDOB)) = False Then
                        Call ErrorListing(NoOfERecLine, VEndtPolicyNo, vEndtName, "Invalid Date format ~ DOB ")
                    End If
                    If Trim(Mid(TextLine, 10, 1)) = "P" Then
                        If Trim(Mid(TextLine, 134, 90)) = "" Then
                            Call ErrorListing(NoOfERecLine, VEndtPolicyNo, vEndtName, "Address Is Blank")
                        End If
                        SelFields = "SELECT PLNCode as tmpCode, PLNTopUpStatus as tmpDesc "
                        SelCriteria = " And PLNCode ='" & Trim(VEndtPlanCode) & "'"
                        tmpTableName = "PLN"
                        tmpCode = ""
                        tmpDesc = ""
                        TmpRecFound = False
                        Call SPHISTableRecVerify(SelFields, tmpTableName, SelCriteria, TmpRecFound, tmpCode, tmpDesc)
                        If TmpRecFound = False Then
                            Call ErrorListing(NoOfERecLine, VEndtPolicyNo, vEndtName, "No such Plan Code - ' " & Trim(VEndtPlanCode) & "' !")
                        Else
                            Call VerifyMBMByPassingPA(NoOfERecLine, "HIS", VEndtPolicyNo, vEndtName, VEndtDataType, VEndtPlanCode, VEndtInsured, VEndtEffDate, VEndtExpDate)
                        End If
                    End If
                End If
' ************* CIS Verification
'loh                    VEndtCLnPlan = Mid(TextLine, 6737, 4)
'loh                    VEndtCLnIns = Mid(TextLine, 6741, 1)
'loh                    VEndtClnEffDate = Mid(TextLine, 6742, 2) & "/" & Mid(TextLine, 6744, 2) & "/" & Mid(TextLine, 6746, 4)
'loh                    VEndtClnExpDate = Mid(TextLine, 6750, 2) & "/" & Mid(TextLine, 6752, 2) & "/" & Mid(TextLine, 6754, 4)
'loh                    If Trim(VEndtCLnPlan) = "" And VendtDataType = "P" Then
'loh                        Call ErrorListing(NoOfERecLine, VEndtPolicyNo, vEndtName, "Clinical Plan Is NULL ! ")
'loh                        Call ErrorListing
'loh                    End If
'loh                    If Trim(VEndtCLnIns) = "" And VendtDataType = "P" Then
'loh                        Call ErrorListing(NoOfERecLine, VEndtPolicyNo, vEndtName, "Clinical Insured Type Is NULL! ")
'loh                        Call ErrorListing
'loh                    End If
'loh                    If IsDate(VEndtClnEffDate) = False And VendtDataType = "P" Then
'loh                        Call ErrorListing(NoOfERecLine, VEndtPolicyNo, vEndtName, "Invalid Date Format ~ Clinical Eff Date ! ")
'loh                        Call ErrorListing
'loh                    End If
'loh                    If IsDate(VEndtClnExpDate) = False And VendtDataType = "P" Then
'loh                        Call ErrorListing(NoOfERecLine, VEndtPolicyNo, vEndtName, "Invalid Date Format ~ Clinical Exp Date ! ")
'loh                        Call ErrorListing
'loh                    End If
' ************* End of CIS verify
            End If
        DoEvents
    Loop
    If Mid(TextLine, 1, 1) <> "T" Then
        Call ErrorListing(NoOfERecLine, "", "", "1st Record must be trailer with Data Type 'T'")
    ElseIf Trim(Mid(TextLine, 2, 10)) <> TotPrincipal Then
        Call ErrorListing(NoOfERecLine, "", "", "Total Principal Record Not Match Trailer Total Principal")
    ElseIf Trim(Mid(TextLine, 12, 10)) <> TotSupplementary Then
        Call ErrorListing(NoOfERecLine, "", "", "Total Supplementary Record Not Match Trailer Total Supplementary")
    End If
                            
    Close #1   ' Close file.
    txtTotalPrincipal = TotPrincipal
    txtTotalSupplementary = TotSupplementary
    NoOfERecLine = 0
    
End Sub
'**************************End Verification of ENDt File*********************************


Private Sub ErrorListing(tmpLineNo As Integer, tmpPol As String, tmpMBMName As String, tmpMSG As String)
Dim ListError As ListItem
    txtTotalError = txtTotalError + 1
    Set ListError = lstERRList.listitems.Add(, , tmpLineNo)
        ListError.SubItems(1) = Trim(tmpPol)
        ListError.SubItems(2) = Trim(tmpMBMName)
        ListError.SubItems(3) = Trim(tmpMSG)
        ListError.SubItems(4) = Trim(txtImportFilename)
End Sub

Private Sub GenerateMember()
On Err GoTo ErrorSave
Dim startTrans As Boolean
Dim strsql As String
Dim PayorCode As String
Dim recordCount As Integer
Dim CovCnt As Integer

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
            
            medixmasterconn.beginTrans
            medixmasterconnMaint.beginTrans
            startTrans = True
            txtBatchNo = CheckBatch(Mid(txtPAYCode, 1, 2), txtSubmissionDate, Mid(txtFileName, 1, 1))
            Do While Not EOF(1)   ' Loop until end of file.
               Line Input #1, TextLine   ' Read line into variable.
                recordCount = recordCount + 1
            
                If Mid(TextLine, 1, 1) = "P" Then
                    
                    CoveredID = "00"
                    CovCnt = "00"
                    Call SavePrincipalDataImport(Mid(txtPAYCode, 1, 2), CoveredID & 0)
                    Call SaveMBMTwo("N")
                    Call SaveOthersDataImport
                    Call SaveMBMOthersTwo("N")
                    If Trim(Mid(TextLine, 6708, 4)) <> "" Then
                        Call SaveCISDataImport("N")
                    End If
                ElseIf Mid(TextLine, 1, 1) <> "P" And Mid(TextLine, 1, 1) <> "T" Then
                    
                    Call SaveCovPersonDataImport(Mid(txtPAYCode, 1, 2), CoveredID)
                    CovCnt = CovCnt + 1
                    Call SaveCovPersonsTwo(CovCnt, "N")
                End If
                DoEvents
            Loop
            
            medixmasterconn.CommitTrans
            medixmasterconnMaint.CommitTrans
            MsgBox "Record Import Sucessfully...", vbOKOnly + vbInformation, DialogBoxTitle
            
            startTrans = False
            Screen.MousePointer = vbDefault
            Close #1
            Exit Sub
        End If
    Close #1   ' Close file.
ErrorSave:
        Screen.MousePointer = vbDefault
        If startTrans = True Then
            medixmasterconn.RollbackTrans
            medixmasterconnMaint.RollbackTrans
        End If
        MsgBox Err.Description
End Sub

Private Sub SavePrincipalDataImport(Payor As String, CoveredIDCount As Integer)
Dim MemberDOB As String
Dim FindAge As String
Dim AgeValue As String
Dim AnnualLimitValue As String
Dim GetProString As String
Dim tmpPreMemberNo As String
Dim TmpCityCode, tmpPostCode, TmpSTACode As String
Dim tmpIC, TmpDOB, TmpSex As String
Dim TmpBRNCode, TmpRaceCode As String
Dim TmpTakeOver, TmpRenewal, TmpAppName As String
Dim TmpLMBM As New ADODB.Recordset
Dim tmpPlan As New ADODB.Recordset
Dim TmpSql, TmpSql2 As String
Dim Check As Boolean
Dim tmpPolDis As Double
Dim tmpRenewDis As Double
Dim DiscPlan As New ADODB.Recordset
Dim TmpTopupInd As String
Dim XrefSql As String
Dim XrefDB As New ADODB.Recordset
Dim strsqlannlmt As String
Dim CLNPlnSql  As String
Dim CLNPln As New ADODB.Recordset
Dim ClnAnnLmtSql As String
Dim CLnAnnLmt As New ADODB.Recordset
Dim MBMWSql As String
Dim TmpCovID, TmpAdd1, TmpAdd2, TmpAdd3 As String
Dim TmpValueDate, TmpMemberType, TmpExpStatus As String
Dim tmpBatchNo As String
Dim TmpAdultChild, TmpAvLmt, tmpClnStatus As String
Dim tmpClnDAtaStatus, TmpHSClnInd, tmpPayor, tmpHLTCode As String
Dim TmpPremium, TmpBasicPrem As String
Dim tmpINSCode, tmpPlanCode As String
Dim TmpEffDate, TmpExpDate, TmpBordxDate As String
Dim TmpMBMStatus, TmpDataStatus As String
Dim TmpPolicy, TmpAgeCode As String
Dim TmpMBMNumber, tmpMBMName As String
Dim SuppLmtStatus As String
Dim HisCisEffDate, HISCISExpDate As String
Dim HISCisPLN As String
Dim HISCiSIns As String
Dim tmpPLNTopUpStatus As String
Dim tmpKidneyAnnLmt As Double
Dim tmpCancerAnnLmt As Double
Dim tmpHomeNursAnnLmt As Double
Dim tmpSuppBNFStatus As String
Dim RoundUpStatus As String
Dim tempMCOFee As Double
    WLifeStatus = False
    TmpPolicy = Mid(TextLine, 2, 25)
    TmpCovID = "00"
    tmpMBMName = Mid(TextLine, 47, 60)
    TmpAdd1 = Mid(TextLine, 107, 30)
    TmpAdd2 = Mid(TextLine, 137, 30)
    TmpAdd3 = Mid(TextLine, 167, 30)
    TmpValueDate = Format(Date, "yyyy/mm/dd")
    TmpDataStatus = "A"
    TmpExpStatus = "N"
    tmpPayor = Mid(txtPAYCode, 1, 2)
    tmpHLTCode = Mid(cboHLTCode1, 1, 1)
    TmpBordxDate = Format(txtSubmissionDate, "yyyy/mm/dd")
    TmpMBMStatus = "A"
    tmpBatchNo = txtBatchNo
    TmpTopupNo = " "
    TmpTopupInd = " "
            
    TmpCityCode = Mid(TextLine, 197, 20)
    tmpPostCode = Mid(TextLine, 217, 5)
    TmpSTACode = Mid(TextLine, 222, 15)
    tmpIC = Mid(TextLine, 273, 20)
    TmpSex = Mid(TextLine, 301, 1)
    TmpRaceCode = Mid(TextLine, 319, 1)
    TmpEffDate = Format(Mid(TextLine, 374, 2) & "/" & Mid(TextLine, 376, 2) & "/" & Mid(TextLine, 378, 4), "yyyy/mm/dd")
    TmpExpDate = Format(Mid(TextLine, 382, 2) & "/" & Mid(TextLine, 384, 2) & "/" & Mid(TextLine, 386, 4), "yyyy/mm/dd")
    TmpDOB = Format(Mid(TextLine, 293, 2) & "/" & Mid(TextLine, 295, 2) & "/" & Mid(TextLine, 297, 4), "yyyy/mm/dd")
    tmpPlanCode = Trim(Mid(TextLine, 390, 4))
    tmpINSCode = Mid(TextLine, 394, 1)
    TmpTakeOver = Mid(TextLine, 435, 1)
    TmpRenewal = Mid(TextLine, 435, 1)
    TmpBRNCode = ""
    
    PayEffDate = Mid(TextLine, 374, 2) & "/" & Mid(TextLine, 376, 2) & "/" & Mid(TextLine, 378, 4)
    PayorEXPDate = Mid(TextLine, 382, 2) & "/" & Mid(TextLine, 384, 2) & "/" & Mid(TextLine, 386, 4)
    MemberDOB = Mid(TextLine, 293, 2) & "/" & Mid(TextLine, 295, 2) & "/" & Mid(TextLine, 297, 4)
    plancode = Trim(Mid(TextLine, 390, 4))
    InsuredCode = Mid(TextLine, 394, 1)
        
    If TmpTakeOver = "C" Or TmpTakeOver = "Y" Then
        TmpRenewal = "N"
    End If
    
    HISCisPLN = plancode
    HISCiSIns = InsuredCode
    HisCisEffDate = PayEffDate
    HISCISExpDate = PayorEXPDate
    
    CisPlnCode = Trim(Mid(TextLine, 6708, 4))
    CisInsType = Mid(TextLine, 6712, 1)
    CisEffDate = Mid(TextLine, 6713, 2) & "/" & Mid(TextLine, 6715, 2) & "/" & Mid(TextLine, 6717, 4)
    CisExpDate = Mid(TextLine, 6721, 2) & "/" & Mid(TextLine, 6723, 2) & "/" & Mid(TextLine, 6725, 4)
    
    If Trim(plancode) = "" Then
        HisCisEffDate = CisEffDate
        HISCISExpDate = CisExpDate
        HISCisPLN = CisPlnCode
        HISCiSIns = CisInsType
    End If

    SuppLmtStatus = "A"
    TmpMemberType = "P"
    If Trim(plancode) <> "" Then
        strsqlannlmt = "Select SuppLimitStatus, PLNCode from AnnualLimit where PLNCode = '" & Trim(plancode) & _
                        "' AND HLTCode='" & Mid(cboHLTCode1, 1, 1) & "' AND INSCode='" & Trim(InsuredCode) & "'"
        AnnLmt.Open strsqlannlmt, medixmasterconn, adOpenKeyset, adLockOptimistic
    Else
        strsqlannlmt = "Select SuppLimitStatus, PLNCode from AnnualLimit where PLNCode = '" & Trim(CisPlnCode) & _
                       "' AND HLTCode='" & Mid(cboHLTCode1, 1, 1) & "' AND INSCode='" & Trim(CisInsType) & "'"
        AnnLmt.Open strsqlannlmt, MediConnCISDB, adOpenKeyset, adLockOptimistic
    End If
    If AnnLmt.recordCount > 0 Then
        If Trim(AnnLmt!SuppLimitStatus) <> "" Then
            SuppLmtStatus = Trim(AnnLmt!SuppLimitStatus)
        End If
    End If
    AnnLmt.Close
    Set AnnLmt = Nothing
    
    If Trim(SuppLmtStatus) <> "A" Then
        TmpMemberType = "Q"
    End If
        
    FindAge = GetAge(CDate(HisCisEffDate), CDate(MemberDOB))
    AgeValue = Mid(GetAgeBand(Mid(cboHLTCode1, 1, 1), FindAge, HISCisPLN, HISCiSIns), 1, 2)
    TmpAgeCode = AgeValue
    TmpAdultChild = "C"
    If AgeValue <> "01" Then
        TmpAdultChild = "A"
    End If
'------------------------------------------------------------------------------------------------------------------------
    TmpTopupNo = " "
    TmpTopupInd = " "
    tmpPLNTopUpStatus = ""
    tmpPLNLifeTimeStatus = ""
    tmpSuppBNFStatus = "A"
    RoundUpStatus = "Y"
    If Trim(plancode) <> "" Then 'only for HIS prem, mco, annLmt, Access Calc.
        TmpSql = "SELECT PLNTopUpStatus, PLNLifeTimeStatus,PLNRoundUpStatus from PLN where PlnCode='" & Trim(plancode) & _
                "' and HLTCode = '" & Mid(cboHLTCode1, 1, 1) & _
                "' AND PLNEffDate <='" & Format(PayEffDate, "mm/dd/yyyy") & "'"
        If Mid(cboHLTCode1, 1, 1) = "N" Then
            TmpSql = TmpSql & " and PAYCode='" & Trim(Payor) & "'"
        End If

        tmpPlan.Open TmpSql, medixmasterconn, adOpenKeyset, adLockOptimistic
        If tmpPlan.recordCount Then
            tmpPLNTopUpStatus = IIf(Len(Trim(tmpPlan!PLNTopUpStatus)), Trim(tmpPlan!PLNTopUpStatus), "")
            tmpPLNLifeTimeStatus = IIf(Len(Trim(tmpPlan!PLNLifeTimeStatus)), Trim(tmpPlan!PLNLifeTimeStatus), "")
            RoundUpStatus = IIf(Len(Trim(tmpPlan!PLNRoundUpStatus)), Trim(tmpPlan!PLNRoundUpStatus), "Y")
        End If
        tmpPlan.Close
        Set tmpPlan = Nothing
        txtMBMPremium = Format(GetPremium(InsuredCode, Payor, AgeValue, plancode, Mid(cboHLTCode1, 1, 1), CDate(PayEffDate)), "#,##0.00")
        BasicPrem = Format(Round(txtMBMPremium, 0), "#,##0.00")
        txtMBMPremium = InsertPremium(CDate(PayorEXPDate), CDate(PayEffDate), Format(txtMBMPremium, "#,##0.00"))

        tmpPolDis = IIf(IsNumeric(Mid(TextLine, 3732, 5)), Mid(TextLine, 3732, 5), 0)
        tmpRenewDis = IIf(IsNumeric(Mid(TextLine, 3737, 5)), Mid(TextLine, 3737, 5), 0)
        If tmpPolDis > 0 Then
            txtMBMPremium = txtMBMPremium * (100 - tmpPolDis) / 100
        End If
        If tmpRenewDis > 0 Then
            txtMBMPremium = txtMBMPremium * (100 - tmpRenewDis) / 100
        End If
        txtMBMPremium = Format(Round(txtMBMPremium, 0), "#,##0.00")
        
        If AgeValue = "01" Then
            txtMBMMCOFees = Format(GetMCO(InsuredCode, Mid(cboHLTCode1, 1, 1), "C", CDate(PayEffDate), plancode), "#,##0.00")
        Else
            txtMBMMCOFees = Format(GetMCO(InsuredCode, Mid(cboHLTCode1, 1, 1), "A", CDate(PayEffDate), plancode), "#,##0.00")
        End If
        BasicMCO = txtMBMMCOFees
        txtMBMMCOFees = InsertMCOFees(CDate(PayorEXPDate), CDate(PayEffDate), Format(txtMBMMCOFees, "#,##0.00"))
'        txtMBMMCOFees = Round(txtMBMMCOFees)
        tempMCOFee = txtMBMMCOFees
        If RoundUpStatus <> "N" Then
            tempMCOFee = Round(tempMCOFee)
        End If
        txtMBMMCOFees = tempMCOFee
        AnnualLimitValue = GetAnnualLimit(InsuredCode, plancode, Mid(cboHLTCode1, 1, 1), CDate(PayEffDate))
        TmpAvLmt = CDbl(AnnualLimitValue)
        
' ---------- write Previous MBM no. to new MBMTopupMember field (only for HIS)
        If Trim(tmpPLNTopUpStatus) = "Y" Then
            XrefSql = "Select TOP 1 MBMNumber, MBMTopupNumber, MBMTopupIND From MBMCrossReference Where MBMICBCPP = '" & Trim(tmpIC) & _
                        "' And MBMStatus = 'A' AND PayCode ='" & Mid(txtPAYCode, 1, 2) & "' AND INSCode='" & Trim(InsuredCode) & _
                        "' ORDER BY MBMPayorEffDate DESC, MBMNumber "
            XrefDB.Open XrefSql, medixmasterconnMaint, adOpenKeyset, adLockOptimistic
            
            If XrefDB.recordCount Then
                TmpTopupNo = XrefDB!MBMNumber
                TmpTopupInd = "Y"
            End If
            XrefDB.Close
            Set XrefDB = Nothing
        End If
'-------- end of Topup
    End If   'if trim(planCode)<> ""
    
    If Trim(plancode) = "" And Trim(CisPlnCode) <> "" Then
        tmpPreMemberNo = Payor & CisPlnCode & CisInsType & GenerateMembershipNo(Payor, CisPlnCode, CisInsType)
    Else
        tmpPreMemberNo = Payor & plancode & InsuredCode & GenerateMembershipNo(Payor, plancode, InsuredCode)
    End If
    TmpMBMNumber = tmpPreMemberNo
    CLNPlnSql = "SELECT AnnLmtIND, VisLmtIND, DCStatus, PAYCode, PDentalStatus, SDentalStatus, " & _
                "PSpecStatus, SSpecStatus, PMEPStatus, SMEPStatus from PLN where PLNCode='" & Trim(CisPlnCode) & _
            "' AND HLTCode='" & Mid(cboHLTCode1, 1, 1) & "'"
    CLNPln.Open CLNPlnSql, MediConnCISDB, adOpenKeyset, adLockOptimistic
    
    CISCovSuppLMtStat = "A"
    CISDCStatus = ""
    CISDCPayor = ""
    CISPDentStatus = ""
    CISSDentStatus = ""
    CISPSpecStatus = ""
    CISSSpecStatus = ""
    CISPMepStatus = ""
    CISSMepStatus = ""
    If CLNPln.recordCount > 0 Then
        CISDCStatus = IIf(Len(Trim(CLNPln!DCStatus)), Trim(CLNPln!DCStatus), "")
        CISDCPayor = IIf(Len(Trim(CLNPln!PAYCode)), Trim(CLNPln!PAYCode), "")
        CISPDentStatus = IIf(Len(Trim(CLNPln!PDentalStatus)), Trim(CLNPln!PDentalStatus), "")
        CISSDentStatus = IIf(Len(Trim(CLNPln!SDentalStatus)), Trim(CLNPln!SDentalStatus), "")
        CISPSpecStatus = IIf(Len(Trim(CLNPln!PSpecStatus)), Trim(CLNPln!PSpecStatus), "")
        CISSSpecStatus = IIf(Len(Trim(CLNPln!SSpecStatus)), Trim(CLNPln!SSpecStatus), "")
        CISPMepStatus = IIf(Len(Trim(CLNPln!PMepStatus)), Trim(CLNPln!PMepStatus), "")
        CISSMepStatus = IIf(Len(Trim(CLNPln!SMepStatus)), Trim(CLNPln!SMepStatus), "")

        If Trim(CLNPln!ANNLmtInd) = "Y" Then
            ClnAnnLmtSql = "Select SuppLimitStatus from AnnualLimit where PLNCode = '" & Trim(CisPlnCode) & _
                           "' AND HLTCode='" & Mid(cboHLTCode1, 1, 1) & "' AND INSCode='" & Trim(CisInsType) & "'"
            CLnAnnLmt.Open ClnAnnLmtSql, MediConnCISDB, adOpenKeyset, adLockOptimistic
            If CLnAnnLmt.recordCount > 0 Then
                CISCovSuppLMtStat = CLnAnnLmt!SuppLimitStatus
            End If
            CLnAnnLmt.Close
            Set CLnAnnLmt = Nothing
        End If
    End If
    CLNPln.Close
    Set CLNPln = Nothing
    TmpHSClnInd = "H"
    If Trim(plancode) = "" And Trim(CisPlnCode) <> "" Then
        TmpHSClnInd = "C"
    End If
    If Trim(plancode) <> "" And Trim(CisPlnCode) = "" Then
        TmpHSClnInd = "H"
    End If
    If Trim(CisPlnCode) <> "" And Trim(plancode) <> "" Then
        If Trim(CISDCStatus) = "Y" Then
            TmpHSClnInd = "M"
        Else
            TmpHSClnInd = "B"
        End If
    End If
    TmpMBMNumber = tmpPreMemberNo
    txtTempMemberNo = tmpPreMemberNo
             
    TmpAdd1 = Trim(Replace(TmpAdd1, "'", "''"))
    TmpAdd2 = Trim(Replace(TmpAdd2, "'", "''"))
    TmpAdd3 = Trim(Replace(TmpAdd3, "'", "''"))
    TmpCityCode = Trim(Replace(TmpCityCode, "'", "''"))
    TmpSTACode = Trim(Replace(TmpSTACode, "'", "''"))
    tmpMBMName = Trim(Replace(tmpMBMName, "'", "''"))
    TmpPolicy = Trim(Replace(TmpPolicy, "'", "''"))
' ----- LifeTimeLimit & BenefitLimit
    PrevLifeMBMNo = ""
    tmpGuaRenewal = Trim(Mid(TextLine, 3842, 5))
    If Trim(tmpGuaRenewal) <> "" Then
        If Trim(tmpPLNLifeTimeStatus) = "Y" Then
            tmpKidneyAnnLmt = 0
            tmpCancerAnnLmt = 0
            tmpHomeNursAnnLmt = 0
            Call GetBenefitLimit(InsuredCode, plancode, Mid(cboHLTCode1, 1, 1), CDate(PayEffDate), tmpKidneyAnnLmt, tmpCancerAnnLmt, tmpHomeNursAnnLmt, tmpSuppBNFStatus)
            Call SaveMBMBnfAnnLmt(TmpMBMNumber, "00", "A", tmpKidneyAnnLmt, tmpCancerAnnLmt, tmpHomeNursAnnLmt)
            TmpSql2 = "Select top 1 MBMNumber, MBMTopupNumber, MBMAvailableLimit From MBM where MBMPolicyNo='" & Trim(TmpPolicy) & _
                      "' And MBMStatus = 'A' AND PayCode ='" & Mid(txtPAYCode, 1, 2) & "' AND INSCode='" & Trim(InsuredCode) & _
                      "' ORDER BY MBMPayorEffDate , MBMNumber "
            TmpLMBM.Open TmpSql2, medixmasterconn, adOpenForwardOnly, adLockReadOnly
            TmpTopupNo = TmpMBMNumber
            PrevLifeMBMNo = TmpMBMNumber

            If TmpLMBM.EOF = False Then
                TmpTopupNo = TmpLMBM!MBMTopupNumber
                PrevLifeMBMNo = TmpLMBM!MBMTopupNumber
                If Trim(tmpGuaRenewal) <> "3" Then
                    TmpAvLmt = IIf(IsNumeric(TmpLMBM!MBMAvailableLimit), TmpLMBM!MBMAvailableLimit, 0)
                End If
            End If
            TmpLMBM.Close
            Set TmpLMBM = Nothing
            
'            TmpSql2 = "Select MBMNumber, MBMLifeTimeLmt From VIEW_MBMLifetime where MBMPolicyNo='" & Trim(TmpPolicy) & _
'                      "' AND MBMCovID='00'"
            TmpSql2 = "Select MBMNumber, MBMLifeTimeLmt From MBMLifeTime where MBMNumber='" & Trim(PrevLifeMBMNo) & _
                      "' AND MBMCovID='00'"
            TmpLMBM.Open TmpSql2, medixmasterconn, adOpenForwardOnly, adLockReadOnly
            If TmpLMBM.EOF Then
                If Trim(tmpGuaRenewal) = "3" Then
                    WLifeStatus = True
                End If
            End If
            TmpLMBM.Close
            Set TmpLMBM = Nothing
            
            If WLifeStatus = True Then
                Call SaveMBMLifeTimeLmt(TmpTopupNo, "00", "A")
            End If
            
        End If
    End If
' ----- end of LifeTimeLimit & BenefitLimit
    If Trim(tmpPlanCode) <> "" Then
        MBMWSql = "EXEC Upload_MBM_HIS '" & tmpPayor & "', '" & tmpHLTCode & "', '" & TmpMBMNumber & "', '" & TmpPolicy & "', '" & TmpCovID & "', " & _
                  "'" & Trim(tmpIC) & "', '" & tmpMBMName & "', '" & TmpDOB & "','" & TmpSex & "','" & TmpRaceCode & "', " & _
                  "'" & TmpMBMStatus & "','" & TmpDataStatus & "', '" & TmpMemberType & "', " & _
                  "'" & TmpAdd1 & "', '" & TmpAdd2 & "', '" & TmpAdd3 & "', " & _
                  "'" & TmpCityCode & "', '" & tmpPostCode & "', '" & TmpSTACode & "', " & _
                  "'" & TmpExpStatus & "', '" & TmpValueDate & "', '" & TmpEffDate & "', '" & TmpExpDate & "', " & _
                  "'" & TmpTakeOver & "', '" & TmpRenewal & "', '" & TmpBRNCode & "', '" & TmpBordxDate & "', '" & Trim(tmpBatchNo) & "', " & _
                  "'" & tmpPlanCode & "', '" & tmpINSCode & "', '" & TmpAgeCode & "', '" & TmpAdultChild & "', " & TmpAvLmt & ", " & _
                  "'" & TmpHSClnInd & "', '" & Trim(TmpTopupNo) & "', '" & Trim(TmpTopupInd) & "'"
    Else
        MBMWSql = "EXEC Upload_MBM_CIS '" & tmpPayor & "', '" & tmpHLTCode & "', '" & TmpMBMNumber & "', '" & TmpCovID & "', " & _
                  "'" & tmpIC & "', '" & tmpMBMName & "', '" & TmpDOB & "','" & TmpSex & "','" & TmpRaceCode & "', " & _
                  "'" & TmpMBMStatus & "','" & TmpDataStatus & "', '" & TmpMemberType & "', " & _
                  "'" & TmpAdd1 & "', '" & TmpAdd2 & "', '" & TmpAdd3 & "', " & _
                  "'" & TmpCityCode & "', '" & tmpPostCode & "', '" & TmpSTACode & "', " & _
                  "'" & TmpExpStatus & "', '" & TmpValueDate & "', '" & TmpBRNCode & "', " & _
                  "'" & TmpBordxDate & "', '" & Trim(tmpBatchNo) & "', '" & TmpAgeCode & "', " & _
                  "'" & TmpAdultChild & "', '" & TmpHSClnInd & "'"
                  
    End If
    
    medixmasterconn.Execute MBMWSql
    Call SaveMBMXRef(TmpMBMNumber, "00", Trim(TmpPolicy), tmpIC, tmpMBMName, TmpEffDate, TmpExpDate, TmpAgeCode, tmpINSCode, tmpPlanCode, TmpBordxDate, TmpTopupNo, TmpTopupInd, TmpDataStatus, TmpMBMStatus, TmpDOB, tmpHLTCode, tmpPayor, tmpBatchNo)
    
    If intExit = 1 Then
        Check = False
    End If
End Sub

Private Sub SaveCovPersonDataImport(Payor As String, CoveredIDCount As Integer)
Dim MemberDOB As String
Dim FindAge As String
Dim AgeValue As String
Dim tempSAnnLmt As String
Dim tempSAvailLmt As String
Dim tempSPremium As String
Dim tmpClmNonPayFr As String
Dim tmpClmNonPayTo As String
Dim SuppPlanCode As String
Dim strSptLmt As String
Dim SptLmtInd As String
Dim SptPlnLmt As New ADODB.Recordset
Dim strSptPrem As String
Dim SptPremInd As String
Dim SptPlnPrem As New ADODB.Recordset
Dim strSptMCOSql As String
Dim SptMCOInd As String
Dim McoCovIDChrg As String
Dim SptPlnMCO As New ADODB.Recordset
Dim tmpIC, TmpSex As String
Dim TmpDOB As String
Dim tmpPolDis As Double
Dim tmpRenewDis As Double
Dim TmpRelCode, TmpCovID As String
Dim TmpMemberType As String
Dim TmpAdultChild As String
Dim TmpAvLmt As String
Dim TmpPremium, TmpBasicPrem, TmpBasicMCO, tmpMCO As String
Dim TmpBordxDate, tmpBatchNo As String
Dim tmpStatus As String
Dim TmpMBMNumber, tmpMBMName As String
Dim SuppWSql As String
Dim StrDateSql, StrWDateSql As String
Dim MBMCovRec As New ADODB.Recordset
Dim HisCisEffDate As String
Dim HIsExpDate As String
Dim tmpSuppExpDate As String
Dim tmpPOlNo As String
Dim tmpCisPlnCOde As String
Dim TmpMBM As New ADODB.Recordset
Dim TmpSql2 As String
Dim tmpPlncode As String
Dim tmpInsType As String
Dim tmpSuppGuaRenewal As String
Dim tmpKidneyAnnLmt As Double
Dim tmpCancerAnnLmt As Double
Dim tmpHomeNursAnnLmt As Double
Dim tmpSuppBNFStatus As String
Dim PLNRec As New ADODB.Recordset
Dim SqlPln, RoundUpStatus As String

    TmpAvLmt = 0
    TmpBasicPrem = 0
    TmpPremium = 0
    TmpBasicMCO = 0
    tmpMCO = 0

    CoveredIDCount = Format(CoveredIDCount + 1, "00")
    TmpRelCode = Trim(Mid(TextLine, 1, 1))
    TmpMBMNumber = Trim(txtTempMemberNo)
    TmpCovID = Format(CoveredIDCount, "00")
    TmpMemberType = "C"
    tmpMBMName = Trim(Mid(TextLine, 47, 60))
    tmpStatus = "A"
    TmpBordxDate = Format(txtSubmissionDate, "yyyy/mm/dd")
    tmpBatchNo = Trim(txtBatchNo)
    tmpPOlNo = Trim(Mid(TextLine, 2, 25))
    tmpIC = Trim(Mid(TextLine, 273, 20))
    TmpSex = Mid(TextLine, 301, 1)
    TmpDOB = Trim(Mid(TextLine, 293, 2) & "/" & Mid(TextLine, 295, 2) & "/" & Mid(TextLine, 297, 4))
    MemberDOB = Trim(Mid(TextLine, 293, 2) & "/" & Mid(TextLine, 295, 2) & "/" & Mid(TextLine, 297, 4))
    tmpPolDis = IIf(IsNumeric(Trim(Mid(TextLine, 3732, 5))), Trim(Mid(TextLine, 3732, 5)), 0)
    tmpRenewDis = IIf(IsNumeric(Trim(Mid(TextLine, 3737, 5))), Trim(Mid(TextLine, 3737, 5)), 0)
    SuppPlanCode = Trim(Mid(TextLine, 390, 4))
    
    HisCisEffDate = PayEffDate
    HIsExpDate = PayorEXPDate
    If Trim(SuppPlanCode) = "" Then
        SuppPlanCode = plancode
    End If
    tmpPlncode = SuppPlanCode
    tmpInsType = INSCode
    If Trim(SuppPlanCode) = "" Then
        tmpPlncode = CisPlnCode
        tmpInsType = CisInsType
    End If
    If Trim(plancode) = "" Then
        HisCisEffDate = Mid(TextLine, 6713, 2) & "/" & Mid(TextLine, 6715, 2) & "/" & Mid(TextLine, 6717, 4)
    End If
    FindAge = GetAge(CDate(HisCisEffDate), CDate(MemberDOB))
    AgeValue = Mid(GetAgeBand(Mid(cboHLTCode1, 1, 1), FindAge, tmpPlncode, CisInsType), 1, 2)
    TmpAdultChild = "A"
    If AgeValue = "01" Then
        TmpAdultChild = "C"
    End If
    
    TmpDOB = Format(TmpDOB, "yyyy/mm/dd")
    If Trim(SuppPlanCode) <> "" Then
        tmpSuppExpDate = GetSuppExpDate(MemberDOB, "HIS", tmpPOlNo, tmpMBMName, SuppPlanCode, PayEffDate, PayorEXPDate, Mid(cboHLTCode1, 1, 1))
        HIsExpDate = Format(tmpSuppExpDate, "dd/mm/yyyy")
        tmpSuppExpDate = Format(tmpSuppExpDate, "yyyy/mm/dd")
    End If
'    StrDateSql = ""
'    tmpClmNonPayFr = Mid(TextLine, 3576, 2) & "/" & Mid(TextLine, 3578, 2) & "/" & Mid(TextLine, 3580, 4)
'    If IsDate(tmpClmNonPayFr) Then
'        tmpClmNonPayFr = Format(tmpClmNonPayFr, "yyyy/mm/dd")
'        StrDateSql = StrDateSql & ", MBMCCLMNonPayFr "
'        StrWDateSql = StrWDateSql & ",'" & tmpClmNonPayFr & "'"
'    End If
'    tmpClmNonPayTo = Mid(TextLine, 3584, 2) & "/" & Mid(TextLine, 3586, 2) & "/" & Mid(TextLine, 3588, 4)
'    If IsDate(tmpClmNonPayTo) Then
'        tmpClmNonPayTo = Format(tmpClmNonPayTo, "yyyy/mm/dd")
'        StrDateSql = StrDateSql & ", MBMCCLMNonPayTo "
'        StrWDateSql = StrWDateSql & ",'" & tmpClmNonPayTo & "'"
'    End If
    SptLmtInd = "A"
    If Trim(SuppPlanCode) <> "" Then
        strSptLmt = "Select SuppLimitStatus, PLNCode from AnnualLimit where PLNCode = '" & Trim(SuppPlanCode) & _
                    "' AND HLTCode='" & Mid(cboHLTCode1, 1, 1) & "' AND INSCode='" & Trim(InsuredCode) & _
                    "' AND AnnualEffDate <= '" & Format(HisCisEffDate, "mm/dd/yyyy") & "' order by AnnualEffDate desc"
        SptPlnLmt.Open strSptLmt, medixmasterconn, adOpenKeyset, adLockOptimistic
        
        If SptPlnLmt.recordCount > 0 Then
            If Trim(SptPlnLmt!SuppLimitStatus) <> "" Then
                SptLmtInd = Trim(SptPlnLmt!SuppLimitStatus)
            End If
        End If
        SptPlnLmt.Close
        Set SptPlnLmt = Nothing
        
        If Mid(cboHLTCode1, 1, 1) = "S" Then
            strSptPrem = "Select SuppPrmStatus, PLNCode from PRM where PLNCode = '" & Trim(SuppPlanCode) & _
                         "' AND HLTCode='" & Mid(cboHLTCode1, 1, 1) & "' AND INSCode='" & Trim(InsuredCode) & _
                         "' AND PRMEffDate <= '" & Format(HisCisEffDate, "mm/dd/yyyy") & "' order by PRMEffDate desc"
        Else
            strSptPrem = "Select SuppPrmStatus, PLNCode from PRM where PLNCode = '" & Trim(SuppPlanCode) & _
                         "' AND HLTCode='" & Mid(cboHLTCode1, 1, 1) & "' AND INSCode='" & Trim(InsuredCode) & _
                         "' AND PAYCode='" & Mid(txtPAYCode, 1, 2) & _
                         "' AND PRMEffDate <= '" & Format(HisCisEffDate, "mm/dd/yyyy") & "' order by PRMEffDate desc"
        End If
        SptPlnPrem.Open strSptPrem, medixmasterconn, adOpenKeyset, adLockOptimistic
        SptPremInd = "A"
        If SptPlnPrem.recordCount > 0 Then
            If Trim(SptPlnPrem!SuppPrmStatus) <> "" Then
                SptPremInd = Trim(SptPlnPrem!SuppPrmStatus)
            End If
        End If
        SptPlnPrem.Close
        Set SptPlnPrem = Nothing
        
        strSptMCOSql = "Select SuppMCOStatus, CovIDChrg from MCO where PLNCode = '" & Trim(SuppPlanCode) & _
                       "' AND HLTCode='" & Mid(cboHLTCode1, 1, 1) & "' AND INSCode='" & Trim(InsuredCode) & _
                       "' AND MCOEffdate <= '" & Format(HisCisEffDate, "mm/dd/yyyy") & "' order by MCOEffDate desc"

        SptPlnMCO.Open strSptMCOSql, medixmasterconn, adOpenKeyset, adLockOptimistic
        SptMCOInd = "A"
        McoCovIDChrg = ""
        If SptPlnMCO.recordCount > 0 Then
            If Trim(SptPlnMCO!SuppMCOStatus) <> "" Then
                SptMCOInd = Trim(SptPlnMCO!SuppMCOStatus)
            End If
            If Trim(SptPlnMCO!CovIDChrg) <> "" Then
                McoCovIDChrg = Trim(SptPlnMCO!CovIDChrg)
            End If
        Else
            SptPlnMCO.Close
            strSptMCOSql = "Select SuppMCOStatus, CovIDChrg from MCO where PLNCode is null " & _
                           " AND HLTCode='" & Mid(cboHLTCode1, 1, 1) & "' AND INSCode='" & Trim(InsuredCode) & _
                           "' AND MCOEffdate <= '" & Format(HisCisEffDate, "mm/dd/yyyy") & "' order by MCOEffDate desc"
            SptPlnMCO.Open strSptMCOSql, medixmasterconn, adOpenKeyset, adLockOptimistic
            If Trim(SptPlnMCO!SuppMCOStatus) <> "" Then
                SptMCOInd = Trim(SptPlnMCO!SuppMCOStatus)
            End If
            If Trim(SptPlnMCO!CovIDChrg) <> "" Then
                McoCovIDChrg = Trim(SptPlnMCO!CovIDChrg)
            End If
        End If
        SptPlnMCO.Close
        Set SptPlnMCO = Nothing
        tmpSuppGuaRenewal = Trim(Mid(TextLine, 3842, 5))
        If Trim(tmpSuppGuaRenewal) = "" Then
            tmpSuppGuaRenewal = Trim(tmpGuaRenewal)
        End If
        If Trim(SptLmtInd) = "B" And Format(CoveredIDCount, "00") = "01" Or Trim(SptLmtInd) = "C" Then
            tempSAnnLmt = Format(GetSuppLimit(InsuredCode, SuppPlanCode, Mid(cboHLTCode1, 1, 1), CDate(PayEffDate)), "#,##0.00")
            TmpAvLmt = tempSAnnLmt
        End If
' ----- Life Time LImit
        If Trim(tmpSuppGuaRenewal) <> "" Then
           If Trim(tmpPLNLifeTimeStatus) = "Y" Then
                tmpKidneyAnnLmt = 0
                tmpCancerAnnLmt = 0
                tmpHomeNursAnnLmt = 0
                Call GetBenefitLimit(InsuredCode, SuppPlanCode, Mid(cboHLTCode1, 1, 1), CDate(PayEffDate), tmpKidneyAnnLmt, tmpCancerAnnLmt, tmpHomeNursAnnLmt, tmpSuppBNFStatus)
                If (tmpSuppBNFStatus = "B" And TmpCovID = "01") Or tmpSuppBNFStatus = "C" Then
                    Call SaveMBMBnfAnnLmt(TmpMBMNumber, TmpCovID, tmpSuppBNFStatus, tmpKidneyAnnLmt, tmpCancerAnnLmt, tmpHomeNursAnnLmt)
                End If
                If WLifeStatus Then
                    Call SaveMBMLifeTimeLmt(PrevLifeMBMNo, TmpCovID, SptLmtInd)
                Else
                    TmpSql2 = "Select top 1 MBMCNumber,MBMCAvailableLimit From MBMCoveredPersons where MBMCNumber='" & Trim(PrevLifeMBMNo) & _
                             "' AND MBMCCoverID='" & TmpCovID & "'"
                    TmpMBM.Open TmpSql2, medixmasterconn, adOpenForwardOnly, adLockReadOnly
                    If TmpMBM.EOF = False Then
                        If Trim(tmpGuaRenewal) <> "3" Then
                            TmpAvLmt = IIf(IsNumeric(TmpMBM!MBMCAvailableLimit), TmpMBM!MBMCAvailableLimit, 0)
                        End If
                    End If
                    TmpMBM.Close
                    Set TmpMBM = Nothing
                End If
            End If
        End If
' ----- End of Life Time LImit
        
        If Trim(SptPremInd) = "B" And Format(CoveredIDCount, "00") = "01" Or Trim(SptPremInd) = "C" Then
            tempSPremium = Format(GetSPremium(InsuredCode, Payor, AgeValue, SuppPlanCode, Mid(cboHLTCode1, 1, 1), CDate(PayEffDate)), "#,##0.00")
            tempSPremium = InsertPremium(CDate(HIsExpDate), CDate(PayEffDate), Format(tempSPremium, "#,##0.00"))
            TmpBasicPrem = Format(Round(tempSPremium, 0), "#,##0.00")
            If tmpPolDis > 0 Then
                tempSPremium = tempSPremium * (100 - tmpPolDis) / 100
            End If
            If tmpRenewDis > 0 Then
                tempSPremium = tempSPremium * (100 - tmpRenewDis) / 100
            End If
            tempSPremium = Format(Round(tempSPremium, 0), "#,##0.00")
            TmpPremium = tempSPremium
        End If
        If (Trim(SptMCOInd) = "B" And TmpCovID = "01") _
            Or Trim(SptMCOInd) = "C" _
            Or (Trim(SptMCOInd) = "D" And TmpCovID >= Trim(McoCovIDChrg)) Then
            
            TmpBasicMCO = Format(GetSuppMCO(InsuredCode, Mid(cboHLTCode1, 1, 1), TmpAdultChild, CDate(PayEffDate), plancode), "#,##0.00")
            tmpMCO = InsertMCOFees(CDate(HIsExpDate), CDate(PayEffDate), Format(TmpBasicMCO, "#,##0.00"))
'            tmpMCO = Round(tmpMCO)
            SqlPln = "SELECT PLNRoundUpStatus from PLN where PlnCode='" & Trim(SuppPlanCode) & _
                    "' and HLTCode = '" & Mid(cboHLTCode1, 1, 1) & _
                    "' AND PLNEffDate <='" & Format(HisCisEffDate, "mm/dd/yyyy") & "'"
            If Mid(cboHLTCode1, 1, 1) = "N" Then
                SqlPln = SqlPln & " and PAYCode='" & Mid(txtPAYCode, 1, 2) & "'"
            End If
            PLNRec.Open SqlPln, medixmasterconn, adOpenKeyset, adLockOptimistic
            RoundUpStatus = "Y"
            If PLNRec.recordCount Then
                RoundUpStatus = IIf(Len(Trim(PLNRec!PLNRoundUpStatus)), Trim(PLNRec!PLNRoundUpStatus), "Y")
            End If
            PLNRec.Close
            Set PLNRec = Nothing
            If RoundUpStatus <> "N" Then
                tmpMCO = Round(tmpMCO)
            End If
        End If
   End If
' ******* Clinical Cov Persons Info Upload
    If CisInsType = "H" Or CisInsType = "F" Then
        tmpCisPlnCOde = Trim(Mid(TextLine, 6708, 4))
        If Trim(tmpCisPlnCOde) = "" Then
            tmpCisPlnCOde = CisPlnCode
        End If
        If Trim(tmpCisPlnCOde) <> "" Then
            Call SaveCISCovPerson(Payor, CoveredIDCount, "N")
        End If
    End If
    
' ****** End of Clinical Cov Persons Info Upload
    TmpPremium = CDbl(IIf(IsNumeric(TmpPremium) = True, TmpPremium, 0))
    TmpBasicPrem = CDbl(IIf(IsNumeric(TmpBasicPrem) = True, TmpBasicPrem, 0))
    tmpPolDis = CDbl(IIf(IsNumeric(tmpPolDis) = True, tmpPolDis, 0))
    tmpRenewDis = CDbl(IIf(IsNumeric(tmpRenewDis) = True, tmpRenewDis, 0))
    TmpAvLmt = CDbl(IIf(IsNumeric(TmpAvLmt) = True, TmpAvLmt, 0))
    tmpMBMName = Replace(tmpMBMName, "'", "''")
    tmpMCO = CDbl(IIf(IsNumeric(tmpMCO) = True, tmpMCO, 0))
    TmpBasicMCO = CDbl(IIf(IsNumeric(TmpBasicMCO) = True, TmpBasicMCO, 0))
    
'    SuppWSql = "Insert into MBMCoveredPersons(MBMCRELCode, MBMCNumber, MBMCCoverID, MBMCDOB," & _
'                "MBMCMemberType, MBMCName, MBMCStatus, MBMCBordxDate, MBMCBatchNo, " & _
'                "MBMCICBCPPNo, MBMCSex, MBMCPolicyDisc, MBMCRenewalDisc, MBMCPLNCode, " & _
'                "MBMCAgeband, MBMCAdultChild, MBMCAnnualLimit, MBMCAvailableLimit, " & _
'                "MBMCBasicPremium, MBMCPremium "
'    If Trim(StrDateSql) <> "" Then
'        SuppWSql = SuppWSql & StrDateSql
'    End If
'    SuppWSql = SuppWSql & ")"
    If Trim(plancode) <> "" Then
        SuppWSql = "EXEC Upload_MBMCovPersons '" & TmpRelCode & "','" & TmpMBMNumber & "','" & TmpCovID & "','" & TmpDOB & "'," & _
                    "'" & TmpMemberType & "','" & tmpMBMName & "','" & tmpStatus & "','" & TmpBordxDate & "','" & tmpBatchNo & "', " & _
                    "'" & tmpIC & "','" & TmpSex & "'," & tmpPolDis & "," & tmpRenewDis & ",'" & SuppPlanCode & "', " & _
                    "'" & AgeValue & "','" & TmpAdultChild & "'," & TmpAvLmt & "," & TmpAvLmt & "," & TmpBasicMCO & "," & tmpMCO & ", " & _
                    "" & TmpBasicPrem & "," & TmpPremium & ",'" & Format(Date, "yyyy/mm/dd") & "','" & tmpSuppExpDate & "'"
    Else
        SuppWSql = "EXEC Upload_MBMCovPersonsWOExp '" & TmpRelCode & "','" & TmpMBMNumber & "','" & TmpCovID & "','" & TmpDOB & "'," & _
                    "'" & TmpMemberType & "','" & tmpMBMName & "','" & tmpStatus & "','" & TmpBordxDate & "','" & tmpBatchNo & "', " & _
                    "'" & tmpIC & "','" & TmpSex & "'," & tmpPolDis & "," & tmpRenewDis & ",'" & SuppPlanCode & "', " & _
                    "'" & AgeValue & "','" & TmpAdultChild & "'," & TmpAvLmt & "," & TmpAvLmt & "," & TmpBasicMCO & "," & tmpMCO & "," & _
                    "" & TmpBasicPrem & "," & TmpPremium & ",'" & Format(Date, "yyyy/mm/dd") & "'"
    End If

'    If Trim(StrWDateSql) <> "" Then
'        SuppWSql = SuppWSql & StrWDateSql
'    End If
'    SuppWSql = SuppWSql & ")"
    medixmasterconn.Execute SuppWSql
    
End Sub

Private Sub SaveCISCovPerson(Payor As String, CoveredIDCount As Integer, BordxFmtType As String)
Dim FindAge As String
Dim AgeValue As String
Dim CISSuppAnnLmt, CISSuppVisLmt As String
Dim CISSuppBPrem, CISSuppPrem As String
Dim CISSuppBMCO, CISSuppMCO As String
Dim SuppPlanCode As String
Dim TmpMBMNumber As String
Dim CisDentalLmt As Double
Dim CisSpecLmt As Double
Dim CisMepLmt As Double
Dim CisDentalExtLmt As Double
Dim Cis1stSpecLmt As Double
Dim CisAnnLmtType As String
Dim CisSuppSql As String
Dim TmpCovID As String
Dim MemberDOB As String
Dim SysDate As String
Dim tmpMBMName As String
Dim tmpCisExpDate As String
Dim tmpSuppExpDate As String
Dim CISPayor As String
Dim CISstrSptLmt As String
Dim CISSptLmtInd As String
Dim CISSptPlnLmt As New ADODB.Recordset
Dim CISstrSptPrem As String
Dim CISSptPremInd As String
Dim CISSptPlnPrem As New ADODB.Recordset
Dim CISstrSptMCOSql As String
Dim CISSptMCOInd As String
Dim CISSptPlnMCO As New ADODB.Recordset
Dim tmpCISPln As String
Dim tmpCisInsType As String
Dim tmpCisEffDate As String
Dim tmpCardType As String
Dim CisWarranty As String
Dim CLNStatus As String

    CoveredIDCount = Format(CoveredIDCount, "00")
    CISPayor = IIf(Trim(CISDCStatus) = "Y", CISDCPayor, Mid(txtPAYCode, 1, 2))

    If Trim(BordxFmtType) = "N" Then
            MemberDOB = Mid(TextLine, 293, 2) & "/" & Mid(TextLine, 295, 2) & "/" & Mid(TextLine, 297, 4)
        tmpCISPln = Trim(Mid(TextLine, 6708, 4))
        tmpCisInsType = Mid(TextLine, 6712, 1)
        tmpCisEffDate = Mid(TextLine, 6713, 2) & "/" & Mid(TextLine, 6715, 2) & "/" & Mid(TextLine, 6717, 4)
        tmpCisExpDate = Mid(TextLine, 6721, 2) & "/" & Mid(TextLine, 6723, 2) & "/" & Mid(TextLine, 6725, 4)
        CisWarranty = Trim(Mid(TextLine, 6729, 2500))
        tmpCardType = Trim(Mid(TextLine, 9229, 20))
    ElseIf Trim(BordxFmtType) = "E" Then
            MemberDOB = Mid(TextLine, 320, 2) & "/" & Mid(TextLine, 322, 2) & "/" & Mid(TextLine, 324, 4)
        tmpCISPln = Trim(Mid(TextLine, 6737, 4))
        tmpCisInsType = Mid(TextLine, 6741, 1)
        tmpCisEffDate = Mid(TextLine, 6742, 2) & "/" & Mid(TextLine, 6744, 2) & "/" & Mid(TextLine, 6746, 4)
        tmpCisExpDate = Mid(TextLine, 6750, 2) & "/" & Mid(TextLine, 6752, 2) & "/" & Mid(TextLine, 6754, 4)
        CisWarranty = Trim(Mid(TextLine, 6758, 2500))
        tmpCardType = Trim(Mid(TextLine, 9258, 20))
   End If
    
    
    If IsDate(tmpCisEffDate) = False Then
        tmpCisEffDate = CisEffDate
    End If
    If IsDate(tmpCisExpDate) = False Then
        tmpCisExpDate = CisExpDate
    End If
    If Trim(tmpCISPln) = "" Then
        tmpCISPln = CisPlnCode
    End If
    If Trim(tmpCisInsType) = "" Then
        tmpCisInsType = CisInsType
    End If
    
    TmpMBMNumber = Trim(txtTempMemberNo)
    TmpCovID = Format(CoveredIDCount, "00")
    FindAge = GetAge(CDate(tmpCisEffDate), CDate(MemberDOB))
    AgeValue = Mid(GetAgeBand(Mid(cboHLTCode1, 1, 1), FindAge, tmpCISPln, tmpCisInsType), 1, 2)
    tmpMBMName = Trim(Mid(TextLine, 47, 60))
    
    If Trim(tmpCISPln) <> "" Then
        tmpSuppExpDate = GetSuppExpDate(MemberDOB, "CIS", "", tmpMBMName, tmpCISPln, tmpCisEffDate, tmpCisExpDate, CISPayor)
        tmpSuppExpDate = Format(tmpSuppExpDate, "yyyy/mm/dd")
    End If

' ******* Clinical Supp Limit
        CISstrSptLmt = "Select SuppLimitStatus, PLNCode from AnnualLimit where PLNCode = '" & Trim(tmpCISPln) & _
                       "' AND HLTCode='" & Mid(cboHLTCode1, 1, 1) & "' AND INSCode='" & Trim(tmpCisInsType) & _
                       "' AND AnnEffDate <= '" & Format(tmpCisEffDate, "mm/dd/yyyy") & "' order by AnnEffDate desc"
        CISSptPlnLmt.Open CISstrSptLmt, MediConnCISDB, adOpenKeyset, adLockOptimistic
        CISSptLmtInd = "A"
        If CISSptPlnLmt.recordCount > 0 Then
            If Trim(CISSptPlnLmt!SuppLimitStatus) <> "" Then
                CISSptLmtInd = Trim(CISSptPlnLmt!SuppLimitStatus)
            End If
        End If
        CISSptPlnLmt.Close
        Set CISSptPlnLmt = Nothing
        If Mid(txtPAYCode, 1, 1) = "S" Then
            CISstrSptPrem = "Select SuppPrmStatus, PLNCode from PRM where PLNCode = '" & Trim(tmpCISPln) & _
                             "' AND HLTCode='" & Mid(cboHLTCode1, 1, 1) & "' AND INSCode='" & Trim(tmpCisInsType) & _
                             "' AND PRMEffDate <= '" & Format(tmpCisEffDate, "mm/dd/yyyy") & "' order by PRMEffDate desc"
        Else
            CISstrSptPrem = "Select SuppPrmStatus, PLNCode from PRM where PLNCode = '" & Trim(tmpCISPln) & _
                             "' AND HLTCode='" & Mid(cboHLTCode1, 1, 1) & "' AND INSCode='" & Trim(tmpCisInsType) & _
                             "' AND PAYCode='" & Trim(CISPayor) & _
                             "' AND PRMEffDate <= '" & Format(tmpCisEffDate, "mm/dd/yyyy") & "' order by PRMEffDate desc"
        End If
        CISSptPlnPrem.Open CISstrSptPrem, MediConnCISDB, adOpenKeyset, adLockOptimistic
        CISSptPremInd = "A"
        If CISSptPlnPrem.recordCount > 0 Then
            If Trim(CISSptPlnPrem!SuppPrmStatus) <> "" Then
                CISSptPremInd = Trim(CISSptPlnPrem!SuppPrmStatus)
            End If
        End If
        CISSptPlnPrem.Close
        Set CISSptPlnPrem = Nothing
        
        CISstrSptMCOSql = "Select SuppMCOStatus from MCO where PLNCode = '" & Trim(tmpCISPln) & _
                         "' AND HLTCode='" & Mid(cboHLTCode1, 1, 1) & "' AND INSCode='" & Trim(tmpCisInsType) & _
                         "' AND MCOEffdate <= '" & Format(tmpCisEffDate, "mm/dd/yyyy") & "' order by MCOEffDate desc"

        CISSptPlnMCO.Open CISstrSptMCOSql, MediConnCISDB, adOpenKeyset, adLockOptimistic
        CISSptMCOInd = "A"
        If CISSptPlnMCO.recordCount > 0 Then
            If Trim(CISSptPlnMCO!SuppMCOStatus) <> "" Then
                CISSptMCOInd = Trim(CISSptPlnMCO!SuppMCOStatus)
            End If
        Else
            CISSptPlnMCO.Close
            CISstrSptMCOSql = "Select SuppMCOStatus from MCO where PLNCode is null " & _
                              " AND HLTCode='" & Mid(cboHLTCode1, 1, 1) & "' AND INSCode='" & Trim(tmpCisInsType) & _
                              "' AND MCOEffdate <= '" & Format(tmpCisEffDate, "mm/dd/yyyy") & "' order by MCOEffDate desc"

            CISSptPlnMCO.Open CISstrSptMCOSql, MediConnCISDB, adOpenKeyset, adLockOptimistic
            If CISSptPlnMCO.recordCount > 0 Then
                If Trim(CISSptPlnMCO!SuppMCOStatus) <> "" Then
                    CISSptMCOInd = Trim(CISSptPlnMCO!SuppMCOStatus)
                End If
            End If
        End If
        CISSptPlnMCO.Close
        Set CISSptPlnMCO = Nothing
        
'----
    If Trim(CISSptPremInd) = "B" And Format(CoveredIDCount, "00") = "01" Or Trim(CISSptPremInd) = "C" Then
        CISSuppPrem = Format(GetCLNSPremium(tmpCisInsType, CISPayor, AgeValue, tmpCISPln, Mid(cboHLTCode1, 1, 1), CDate(tmpCisEffDate)), "#,##0.00")
        CISSuppBPrem = CISSuppPrem
        CISSuppPrem = InsertPremium(CDate(tmpSuppExpDate), CDate(tmpCisEffDate), Format(CISSuppPrem, "#,##0.00"))
        CISSuppPrem = Format(Round(CISSuppPrem, 0), "#,##0.00")
    End If
    If Trim(CISSptMCOInd) = "B" And Format(CoveredIDCount, "00") = "01" Or Trim(CISSptMCOInd) = "C" Then
        If AgeValue = "01" Then
            CISSuppMCO = Format(GetCLNSMCO(tmpCisInsType, Mid(cboHLTCode1, 1, 1), "C", CDate(tmpCisEffDate), tmpCISPln), "#,##0.00")
        Else
            CISSuppMCO = Format(GetCLNSMCO(tmpCisInsType, Mid(cboHLTCode1, 1, 1), "A", CDate(tmpCisEffDate), tmpCISPln), "#,##0.00")
        End If
        CISSuppBMCO = CISSuppMCO
        CISSuppMCO = InsertMCOFees(CDate(tmpSuppExpDate), CDate(tmpCisEffDate), Format(CISSuppMCO, "#,##0.00"))
        CISSuppMCO = Round(CISSuppMCO)
    End If
    If Trim(CISSptLmtInd) = "B" And Format(CoveredIDCount, "00") = "01" Or Trim(CISSptLmtInd) = "C" Then
        CISSuppAnnLmt = Format(GetCLNSuppLimit(tmpCisInsType, tmpCISPln, Mid(cboHLTCode1, 1, 1), CDate(tmpCisEffDate)), "#,##0.00")
        CISSuppVisLmt = GetCLNSVisitLimit(tmpCisInsType, tmpCISPln, Mid(cboHLTCode1, 1, 1), CDate(tmpCisEffDate))
        If Trim(CISSDentStatus) = "S" Then 'stand alone dental limit
            CisAnnLmtType = "SDentalAnnLimit"
            CisDentalLmt = GetCISLimit(tmpCisInsType, tmpCISPln, Mid(cboHLTCode1, 1, 1), CDate(tmpCisEffDate), CisAnnLmtType)
        End If
        If Trim(CISSSpecStatus) = "S" Then 'stand alone Spec limit
            CisAnnLmtType = "SSpecAnnLimit"
            CisSpecLmt = GetCISLimit(tmpCisInsType, tmpCISPln, Mid(cboHLTCode1, 1, 1), CDate(tmpCisEffDate), CisAnnLmtType)
        End If
        If Trim(CISSMepStatus) = "S" Then 'stand alone MEP limit
            CisAnnLmtType = "SMepAnnLimit"
            CisMepLmt = GetCISLimit(tmpCisInsType, tmpCISPln, Mid(cboHLTCode1, 1, 1), CDate(tmpCisEffDate), CisAnnLmtType)
        End If
    End If
    
' ****** End of Clinical Supp Limit
    CISSuppBPrem = CDbl(IIf(IsNumeric(CISSuppBPrem) = True, CISSuppBPrem, 0))
    CISSuppPrem = CDbl(IIf(IsNumeric(CISSuppPrem) = True, CISSuppPrem, 0))
    CISSuppBMCO = CDbl(IIf(IsNumeric(CISSuppBMCO) = True, CISSuppBMCO, 0))
    CISSuppMCO = CDbl(IIf(IsNumeric(CISSuppMCO) = True, CISSuppMCO, 0))
    CISSuppAnnLmt = CDbl(IIf(IsNumeric(CISSuppAnnLmt) = True, CISSuppAnnLmt, 0))
    CISSuppVisLmt = CDbl(IIf(IsNumeric(CISSuppVisLmt) = True, CISSuppVisLmt, 0))
    SysDate = Format(Date, "yyyy/mm/dd")
    CisWarranty = Replace(CisWarranty, "'", "''")
    CLNStatus = "A"
'    CisSuppSql = "Insert into MBMCoveredPersonsCLN(MBMCNumber, MBMCCoverID, " & _
'               "MBMCCLNMCO, MBMCCLNBasicMCO, MBMCCLNPrem, MBMCCLNBasicPrem, " & _
'                "MBMCCLNBalLmt, MBMCCLNVisitLmt, MBMCCLNDentalLmt, MBMCCLNSpecLmt, " & _
'                "MBMCCLNMEPLmt, MBMCCLNRegDate, MBMCCLNRegUser, " & _
'                "MBMCCLNLastUpdateDate, MBMCCLNLastUpdateUser)"
    
    CisSuppSql = "EXEC Upload_MBMCovPersonsCis '" & TmpMBMNumber & "','" & TmpCovID & "', " & _
                "" & CISSuppMCO & "," & CISSuppBMCO & "," & CISSuppPrem & "," & CISSuppBPrem & ", " & _
                "" & CISSuppAnnLmt & "," & CISSuppVisLmt & "," & CisDentalLmt & "," & CisSpecLmt & ", " & _
                "" & CisMepLmt & ",'" & SysDate & "','" & Logon & "'," & _
                "'" & SysDate & "','" & Logon & "','" & tmpSuppExpDate & "'," & _
                "'" & tmpCardType & "','" & CisWarranty & "','" & CLNStatus & "','" & Trim(tmpCISPln) & "'"
    medixmasterconn.Execute CisSuppSql
End Sub

Private Sub SaveOthersDataImport()
Dim tmpPreMemberNo As String
Dim TempMBM As New ADODB.Recordset
Dim AA, QQ As String
Dim tmpIC, TempSql As String
Dim tmpClmNonPayFr As String
Dim tmpClmNonPayTo As String
Dim Tmp1stJoinedDate As String
Dim tmpAgent, tmplblStatus, TmpPrevTOpln As String
Dim tmp1stMBMNo, TmpPremium, TmpBasicPrem As String
Dim TmpBasicMCO, tmpBasicEa, tmpMcoFee As String
Dim TmpRelCode, tmpPolDis, tmpRenewDis As String
Dim tmpGrpComp, tmpPrevPolNo, tmpPrevMBMNo, tmpEmployeeNO As String
Dim tmpBranch, tmpDepart, tmpPolSubNo1, tmpPolSubNo2, tmpInstall As String
Dim tmpMBMNewPln, TmpMBMNewPolEffDate As String
Dim tmpEndtRefNo, TmpPolCondition As String
Dim MBMOWSql As String
'Dim StrDateSql, StrWDateSql As String
Dim tmpUnExcess, tmpGuaRenewal As String
Dim MBMOtherRec As New ADODB.Recordset

    TmpRelCode = "P"
    tmplblStatus = ""
    tmpAgent = Trim(Mid(TextLine, 27, 15))
    If Mid(cboLBLStatus, 1, 1) = "Y" Then
        tmplblStatus = Mid(cboLBLStatus, 1, 1)
    End If
    tmp1stMBMNo = Trim(txtTempMemberNo)
    TmpPrevTOpln = Trim(Mid(TextLine, 3371, 4))
    tmpPolDis = Trim(Mid(TextLine, 3732, 5))
    tmpRenewDis = Trim(Mid(TextLine, 3737, 5))
    TmpPolCondition = Trim(Mid(TextLine, 3800, 5))
    tmpGrpComp = Trim(Mid(TextLine, 395, 40))
    tmpBranch = Trim(Mid(TextLine, 2939, 15))
    tmpPrevPolNo = Trim(Mid(TextLine, 2954, 25))
    tmpPrevMBMNo = Trim(Mid(TextLine, 2979, 18))
    tmpEmployeeNO = Trim(Mid(TextLine, 2997, 15))
    Tmp1stJoinedDate = Mid(TextLine, 3012, 2) & "/" & Mid(TextLine, 3014, 2) & "/" & Mid(TextLine, 3016, 4)
    tmpDepart = Trim(Mid(TextLine, 3221, 30))
    tmpInstall = Mid(TextLine, 3375, 1)
    tmpPolSubNo1 = Trim(Mid(TextLine, 3376, 10))
    tmpPolSubNo2 = Trim(Mid(TextLine, 3386, 10))
    tmpClmNonPayFr = Mid(TextLine, 3576, 2) & "/" & Mid(TextLine, 3578, 2) & "/" & Mid(TextLine, 3580, 4)
    tmpClmNonPayTo = Mid(TextLine, 3584, 2) & "/" & Mid(TextLine, 3586, 2) & "/" & Mid(TextLine, 3588, 4)
    If Trim(Mid(TextLine, 3592, 4)) <> "" And Trim(Mid(TextLine, 3596, 8)) <> "" Then
       tmpMBMNewPln = Trim(Mid(TextLine, 3592, 4))
       TmpMBMNewPolEffDate = Mid(TextLine, 3596, 2) & "/" & Mid(TextLine, 3598, 2) & "/" & Mid(TextLine, 3600, 4)
    End If
    tmpEndtRefNo = Trim(Mid(TextLine, 3604, 20))
    tmpPolDis = Trim(Mid(TextLine, 3732, 5))
    tmpRenewDis = Trim(Mid(TextLine, 3737, 5))
    tmpUnExcess = Trim(Mid(TextLine, 3830, 12))
    tmpGuaRenewal = Trim(Mid(TextLine, 3842, 5))
     
'------------------
'    StrDateSql = ""
'    StrWDateSql = ""
'    If IsDate(Tmp1stJoinedDate) Then
'        Tmp1stJoinedDate = Format(Tmp1stJoinedDate, "YYYY/MM/DD")
'        StrDateSql = StrDateSql & ", MBMFirstJoinedDate "
'        StrWDateSql = StrWDateSql & ",'" & Tmp1stJoinedDate & "'"
'    End If
'    If IsDate(tmpClmNonPayFr) Then
'        tmpClmNonPayFr = Format(tmpClmNonPayFr, "YYYY/MM/DD")
'        StrDateSql = StrDateSql & ", MBMCLMNonPayFr "
'        StrWDateSql = StrWDateSql & ",'" & tmpClmNonPayFr & "'"
'    End If
'    If IsDate(tmpClmNonPayTo) Then
'        tmpClmNonPayTo = Format(tmpClmNonPayTo, "YYYY/MM/DD")
'        StrDateSql = StrDateSql & ", MBMCLMNonPayTo "
'        StrWDateSql = StrWDateSql & ",'" & tmpClmNonPayTo & "'"
'    End If
'    If IsDate(TmpMBMNewPolEffDate) Then
'        TmpMBMNewPolEffDate = Format(TmpMBMNewPolEffDate, "YYYY/MM/DD")
'        StrDateSql = StrDateSql & ", MBMNewPolEffDate "
'        StrWDateSql = StrWDateSql & ",'" & TmpMBMNewPolEffDate & "'"
'    End If
    
    TmpPremium = CDbl(IIf(IsNumeric(txtMBMPremium) = True, txtMBMPremium, 0))
    TmpBasicPrem = CDbl(IIf(IsNumeric(BasicPrem) = True, BasicPrem, 0))
    TmpBasicMCO = CDbl(IIf(IsNumeric(BasicMCO) = True, BasicMCO, 0))
    tmpMcoFee = CDbl(IIf(IsNumeric(txtMBMMCOFees) = True, txtMBMMCOFees, 0))
    tmpPolDis = CDbl(IIf(IsNumeric(tmpPolDis) = True, tmpPolDis, 0))
    tmpRenewDis = CDbl(IIf(IsNumeric(tmpRenewDis) = True, tmpRenewDis, 0))
    tmpUnExcess = CDbl(IIf(IsNumeric(tmpUnExcess) = True, tmpUnExcess, 0))
    tmpGrpComp = Replace(tmpGrpComp, "'", "''")
    tmpDepart = Replace(tmpDepart, "'", "''")
'    MBMOWSql = "Insert Into MBMOthers(MBMNumber, MBMAgentCode, LabelStatus, MBMAllergic, " & _
'                 "MBMFirstMEMNo, MBMPremium, MBMBasicPrem, MBMBasicMCO, MBMMCOFee, " & _
'                 "RELCODE, MBMPolicyDisc, MBMRenewalDisc, MBMGroupCompany, MBMPrevTakeOverPlnCode,  " & _
'                 "MBMPrevPolNo, MBMPrevMemNo, MBMEmployeeNo, " & _
'                 "MBMDepartment, MBMInstallment, MBMPolSubNo1, MBMPOLSubNo2, MBMBranch, " & _
'                 " MBMNewPlanCode, MBMENDtRefNo, MBMPolCondition, MBMUndExcess "
'  If Trim(StrDateSql) <> "" Then
'        MBMOWSql = MBMOWSql & StrDateSql
'  End If
'  MBMOWSql = MBMOWSql & ")"
'  MBMOWSql = MBMOWSql & "Values('" & Trim(txtTempMemberNo) & "','" & tmpAgent & "','" & tmplblStatus & "','" & TmpAllergic & "', " & _
'                 "'" & tmp1stMBMNo & "'," & TmpPremium & "," & TmpBasicPrem & "," & tmpBasicMco & "," & tmpMcoFee & ", " & _
'                 "'" & TmpRelCode & "'," & tmpPolDis & "," & tmpRenewDis & ",'" & tmpGrpComp & "', " & TmpPrevTOpln & "', " & _
'                 "'" & tmpPrevPolNo & "','" & tmpPrevMBMNo & "','" & tmpEmployeeNO & "', " & _
'                 "'" & tmpDepart & "','" & tmpInstall & "','" & tmpPolSubNo1 & "','" & tmpPolSubNo2 & "','" & tmpBranch & "', " & _
'                 "'" & tmpMBMNewPln & "','" & tmpEndtRefNo & "','" & TmpPolCondition & "'," & tmpUnExcess
'  If Trim(StrWDateSql) <> "" Then
'        MBMOWSql = MBMOWSql & StrWDateSql
'  End If
'  MBMOWSql = MBMOWSql & ")"

  MBMOWSql = "EXEC Upload_MBMOthers '" & Trim(txtTempMemberNo) & "','" & tmpAgent & "','" & tmplblStatus & "', " & _
                 "'" & tmp1stMBMNo & "'," & TmpPremium & "," & TmpBasicPrem & "," & TmpBasicMCO & "," & tmpMcoFee & ", " & _
                 "'" & TmpRelCode & "'," & tmpPolDis & "," & tmpRenewDis & ",'" & tmpGrpComp & "', '" & TmpPrevTOpln & "', " & _
                 "'" & tmpPrevPolNo & "','" & tmpPrevMBMNo & "','" & tmpEmployeeNO & "', " & _
                 "'" & tmpDepart & "','" & tmpInstall & "','" & tmpPolSubNo1 & "','" & tmpPolSubNo2 & "','" & tmpBranch & "', " & _
                 "'" & tmpMBMNewPln & "','" & tmpEndtRefNo & "','" & TmpPolCondition & "'," & tmpUnExcess & ",'" & tmpGuaRenewal & "'"
  medixmasterconn.Execute MBMOWSql

End Sub

Private Sub GenerateENDtMember()
    Dim PayorCode As String
    Dim recordCount As Integer
    Dim CovCnt As Integer
    Dim EndtInsType As String
    Dim CovTwoSql As String
    
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
'        MBM.Open "MBM", medixmasterconn, adOpenKeyset, adLockOptimistic
        Do While Not EOF(1)   ' Loop until end of file.
            Line Input #1, TextLine   ' Read line into variable.
            recordCount = recordCount + 1
                EndtInsType = Mid(TextLine, 421, 1)
'            If Mid(TextLine, 1, 1) = "I" Then
'                If EndtInsType = "G" Or EndtInsType = "H" Then
                                 
'                    If Mid(TextLine, 10, 1) = "P" Then
'                        CoveredID = "00"
'                        CovCnt = "00"
'                        Call EndtSavePrincipalDataImport(Mid(txtPAYCode, 1, 2), CoveredID)
'                        MBMTwo.Open "MBMTwo", medixmasterconn, adOpenKeyset, adLockOptimistic
'                        Call SaveMBMTwo("E")
'                        MBMTwo.Close
'                        Set MBMTwo = Nothing
'
'                        Call EndtSaveOthersDataImport
'                        MBMOthersTwo.Open "MBMOthersTwo", medixmasterconn, adOpenKeyset, adLockOptimistic
'                        Call SaveMBMOthersTwo("E")
'                        MBMOthersTwo.Close
'                        Set MBMOthersTwo = Nothing
'
'                        If Trim(Mid(TextLine, 6708, 4)) <> "" Then
'
'                            Call SaveCISDataImport("E")
'
'                        End If
'
'                    Else
'
'                        Call EndtSaveCovPersonDataImport(Mid(txtPAYCode, 1, 2), CoveredID)
'                        MBMCovTwo.Open "MBMCoveredPersonsTwo", medixmasterconn, adOpenKeyset, adLockOptimistic
'                        CovCnt = Format(CovCnt + 1, "00")
'                        Call SaveCovPersonsTwo(CovCnt, "E")
'                        MBMCovTwo.Close
'                        Set MBMCovTwo = Nothing
'                    End If
'                End If
                            
            If Mid(TextLine, 1, 1) = "C" Then
                If Mid(TextLine, 10, 1) = "P" Then
                    Call CancelENDtData
                    
                End If
                
            'ElseIf Mid(TextLine, 1, 1) = "A" Then
            '    If Mid(TextLine, 10, 1) = "P" Then
            '        Call EndtPrincipalAdjustment
            '    Else
            '        Call EndtSupplementaryAdjustment
            '    End If
                
 '           ElseIf Mid(TextLine, 1, 1) = "S" Then
 '               Call EndtAddSupplementary
 '               MBMCovTwo.Open "MBMCoveredPersonsTwo", medixmasterconn, adOpenKeyset, adLockOptimistic
 '               CovCnt = Format(GetLastCoveredID(Trim(Mid(TextLine, 11, 18))), "00")
 '
 '               Call SaveCovPersonsTwo(CovCnt, "E")
 '               MBMCovTwo.Close
 '               Set MBMCovTwo = Nothing
 '           'ElseIf Mid(TextLine, 1, 1) = "T" Then
 '           '    EndtDeleteSupplementary
 '           '
            End If
            DoEvents
        Loop
        MsgBox "Record Import Sucessfully...", vbOKOnly + vbInformation, DialogBoxTitle
'        MBM.Close
    End If
    Close #1   ' Close file.
End Sub

Private Sub CancelENDtData()
Dim tmpMBMNo As String
Dim TmpCovID As String
Dim CancelDate As String
    tmpMBMNo = Trim(Mid(TextLine, 11, 18))
    CancelDate = Mid(TextLine, 2, 2) & "/" & Mid(TextLine, 4, 2) & "/" & Mid(TextLine, 6, 4)
    Call EndtCancelPrincipalData(tmpMBMNo, CDate(CancelDate))
    Call EndtCancelCoveredPersonsData(tmpMBMNo, CDate(CancelDate))
End Sub

Private Sub EndtCancelCoveredPersonsData(tmpMBMNo As String, CancelDate As Date)
Dim strsql As String
Dim MBMCoveredPersons As New ADODB.Recordset
Dim tmpCPrem, tmpCMco As Double
Dim TmpExpDate As String
    
    strsql = "Select MBMCStatus, MBMCExpDate, MBMCPremium, MBMCMCO, MBMCCoverID " & _
             "from MBMCoveredPersons Where MBMCNumber ='" & Trim(tmpMBMNo) & "'"
    MBMCoveredPersons.Open strsql, medixmasterconn, adOpenKeyset, adLockOptimistic
    
    If MBMCoveredPersons.recordCount Then
       Do Until MBMCoveredPersons.EOF
           With MBMCoveredPersons
                TmpExpDate = IIf(Len(!MBMCExpDate), !MBMCExpDate, Null)
                !MBMCStatus = "C"
                tmpCPrem = IIf(IsNumeric(!MBMCPremium), !MBMCPremium, 0)
                tmpCMco = IIf(IsNumeric(!MBMCMCO), !MBMCMCO, 0)
                TmpExpDate = IIf(Len(TmpExpDate), TmpExpDate, tmpEndtPrincExpDate)
                If tmpCPrem > 0 Or tmpCMco > 0 Then
                     Call EndtMBMRefund(Trim(tmpMBMNo), CDate(tmpEndtPrincEffDate), CDate(TmpExpDate), CDate(CancelDate), !MBMCCoverID, CInt(tmpCPrem), CInt(tmpCMco), MBMCancelStatus)
                End If
                MBMCoveredPersons.Update
           End With
           MBMCoveredPersons.MoveNext
       Loop
    End If
    MBMCoveredPersons.Close
    Set MBMCoveredPersons = Nothing
End Sub

Private Sub EndtCancelPrincipalData(tmpMBMNo As String, CancelDate As Date)
Dim strsql As String
Dim MBM As New ADODB.Recordset
Dim StrAmendSql As String
Dim MBMMCOFees As Double
Dim MBMPremium As Double
Dim SqlMBMO As String
Dim MBMOthers As New ADODB.Recordset
Dim MBMTwoSql, XMBMSql As String

    strsql = "Select MBMNumber, MBMStatus, MBMLastEndorseStatus, MBMENDtBordxDate,MBMENDtEffDate, " & _
             "MBMPolicyNo, MBMICBcPp, MBMName, MBMPayorEXPDate, PLNCode, " & _
             "MBMPayorEffDate from MBM Where MBMNumber ='" & Trim(tmpMBMNo) & "'"
    MBM.Open strsql, medixmasterconn, adOpenKeyset, adLockOptimistic

    If MBM.recordCount Then
        StrAmendSql = "Insert into MBMAmendHistory(MBMNumber, MBMPolicyNo, MBMICBcPp, MBMName, " & _
            "MBMExpiryDate, MBMPlanCode, MBMAmendEdtType, " & _
            "MBMAEndtBordxDate, MBMAEndtEffDate, " & _
            "MBMAmendUser, MBMAmendDate, MBMAEndtBatchNo) " & _
            "Values ('" & MBM!MBMNumber & "','" & MBM!MBMPolicyNo & "','" & MBM!MBMICBcPp & "','" & MBM!MBMName & "'," & _
            "'" & Format(MBM!MBMPayorEXPDate, "YYYY/MM/DD") & "','" & MBM!PLNCode & "','C'," & _
            "'" & Format(txtSubmissionDate, "YYYY/MM/DD") & "','" & Format(CancelDate, "YYYY/MM/DD") & "'," & _
            "'" & Logon & "','" & Format(Date, "YYYY/MM/DD") & "','" & txtBatchNo & "')"
        medixmasterconn.Execute StrAmendSql
        
        With MBM
            tmpEndtPrincEffDate = MBM!MBMPayorEffDate
            tmpEndtPrincExpDate = MBM!MBMPayorEXPDate
            If Len(MBM!MBMStatus) Then
                !MBMStatus = "C"
                !MBMLastEndorseStatus = "C"
                !MBMENDtBordxDate = CDate(txtSubmissionDate)
                !MBMENDtEffDate = CancelDate
                MBM.Update
            End If
        End With
        
    End If
    MBM.Close
    Set MBM = Nothing
    
    txtTempMemberNo = Trim(tmpMBMNo)
    
    MBMTwoSql = "Update MBMTwo SET MBMLastAdjustmentDate = '" & Format(Date, "YYYY/MM/DD") & "' , MBMLastUpdateUser= '" & Logon & _
                "' WHERE MBMNumber = '" & Trim(tmpMBMNo) & "'"
    medixmasterconn.Execute MBMTwoSql

    XMBMSql = "Update MBMCrossReference SET MBMStatus = 'C' WHERE MBMNumber = '" & Trim(tmpMBMNo) & "'"
    medixmasterconnMaint.Execute XMBMSql
    
    SqlMBMO = "Select MBMPremium, MBMMCOFee, MBMAccessFee from MBMOthers where MBMNumber = '" & Trim(tmpMBMNo) & "'"
    MBMOthers.Open SqlMBMO, medixmasterconn, adOpenKeyset, adLockOptimistic
    
    If MBMOthers.recordCount Then
        MBMPremium = MBMOthers!MBMPremium
        MBMMCOFees = MBMOthers!MBMMCOFee
    End If
    MBMOthers.Close
    Set MBMOthers = Nothing
    Call EndtMBMRefund(Trim(tmpMBMNo), CDate(tmpEndtPrincEffDate), CDate(tmpEndtPrincExpDate), CDate(CancelDate), "00", CInt(MBMPremium), CInt(MBMMCOFees), MBMCancelStatus)
End Sub

Private Sub EndtMBMRefund(tmpMBMNo As String, PayEffDate As Date, PAYExpDate As Date, CancelDate As Date, TmpCovID As String, TmpPremium As Double, tmpMCO As Double, MBMCancelStatus As String)
Dim MBMMCOFees As Double
Dim HLTCode As String
Dim RefundSql As String
Dim MCORefundValue As Double
Dim PremiumRefundValue As Double
Dim PremiumEffective As Double
Dim MCOEffective As Double
    
    HLTCode = Mid(cboHLTCode1, 1, 1)
    
    MCORefundValue = Format(MCORefund(CDate(PAYExpDate), CDate(CancelDate), CDate(PayEffDate), tmpMCO, MBMCancelStatus), "#,##0.00")
    
    PremiumRefundValue = Format(MCOPremiumReduction(CDate(PAYExpDate), CDate(CancelDate), CDate(PayEffDate), TmpPremium), "#,##0.00")
    PremiumEffective = Format(CDbl(TmpPremium) - CDbl(PremiumRefundValue), "#,##0.00")
    
    MCOEffective = Format(CDbl(tmpMCO) - CDbl(MCORefundValue), "#,##0.00")
    RefundSql = "EXEC Upload_MBMMcoRefund '" & Trim(tmpMBMNo) & "','" & Trim(TmpCovID) & _
              "'," & PremiumEffective & "," & PremiumRefundValue & "," & MCOEffective & _
              "," & MCORefundValue & ",'" & Mid(txtPAYCode, 1, 2) & "'," & _
              "'" & Format(txtSubmissionDate, "yyyy/mm/dd") & "','" & txtBatchNo & "','" & HLTCode & "'"
    medixmasterconn.Execute RefundSql
End Sub

Private Sub EndtDeleteSupplementary()
'On Error Resume Next
Dim strdelcov As String
Dim MBMCov As New ADODB.Recordset
Dim VEndtSuppID As String
'    strsql = "Select * from MBMCov Where MBMNumber = '" & Mid(TextLine, 11, 18) & "'"
'    MBMCov.Open strsql, medixmasterconn, adOpenKeyset, adLockOptimistic
    VEndtSuppID = Trim(Mid(TextLine, 2963, 2))
    strdelcov = "Delete from MBMCoveredPersons Where MBMCNumber ='" & Mid(TextLine, 11, 18) & "'" 'And MBMCCoverID ='" & VEndtSuppID & "'"
    MBMCov.Open strdelcov, medixmasterconn, adOpenKeyset, adLockOptimistic

    MBMCov.Close
    Set MBMCov = Nothing
End Sub
Private Sub EndtAddSupplementary()
    
Dim strsql As String
Dim LastCoveredID As Integer
Dim SelectAge As String
Dim MBMCov As New ADODB.Recordset
Dim AgeValue As String

    strsql = "Select * from MBMCoveredPersons Where MBMCNumber = '" & Trim(Mid(TextLine, 11, 18)) & "'"
    MBMCov.Open strsql, medixmasterconn, adOpenKeyset, adLockOptimistic

    LastCoveredID = GetLastCoveredID(Trim(Mid(TextLine, 11, 18)))
        MBMCov.AddNew
        With MBMCov
            !MBMCNumber = Trim(Mid(TextLine, 11, 18))
            !MBMCRELCode = Mid(TextLine, 10, 1)
            SelectAge = GetAge(Mid(TextLine, 401, 2) & "/" & Mid(TextLine, 403, 2) & "/" & Mid(TextLine, 405, 4), Mid(TextLine, 320, 2) & "/" & Mid(TextLine, 322, 2) & "/" & Mid(TextLine, 324, 4))
            AgeValue = Mid(GetAgeBand(Mid(cboHLTCode1, 1, 1), SelectAge, plancode, INSCode), 1, 2)
            !MBMCAgeband = AgeValue
            If AgeValue = "01" Then
                !MBMCAdultChild = "C"
            Else
                !MBMCAdultChild = "A"
            End If
            !MBMCStatus = "A"
            !MBMCName = Trim(Mid(TextLine, 74, 60))
            !MBMCICBCPPNo = Trim(Mid(TextLine, 300, 20))
            !MBMCDOB = Mid(TextLine, 320, 2) & "/" & Mid(TextLine, 322, 2) & "/" & Mid(TextLine, 324, 4)
            !MBMCSex = Mid(TextLine, 328, 1)
            !MBMCAllergic = Mid(TextLine, 381, 20)
            !MBMCPolicyDisc = IIf(IsNumeric(Trim(Mid(TextLine, 3761, 5))), Trim(Mid(TextLine, 3761, 5)), 0)
            !MBMCRenewalDisc = IIf(IsNumeric(Trim(Mid(TextLine, 3766, 5))), Trim(Mid(TextLine, 3766, 5)), 0)

            !MBMCMemberType = "C"
            !MBMCCoverID = Format(LastCoveredID, "00")
            MBMCov.Update
        End With
        
        MBMCov.Close
        Set MBMCov = Nothing
End Sub

Private Sub EndtPrincipalAdjustment()
'On Error Resume Next
    Dim MemberHis As New ADODB.Recordset
    Dim MBM As New ADODB.Recordset
    Dim MBMCov As New ADODB.Recordset
    Dim strsql As String
    Dim listrecord As Integer
    Dim listloop As Integer
    Dim ListCount As Integer
    Dim StrSql2 As String
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
            !MBMICBcPp = MBM!MBMICBcPp
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
            !MBMName = Mid(TextLine, 74, 60)
            !MBMPolicyNo = Mid(TextLine, 29, 25)
            !MBMAddress1 = Mid(TextLine, 134, 30)
            !MBMAddress2 = Mid(TextLine, 164, 30)
            !MBMAddress3 = Mid(TextLine, 194, 30)
            !MBMStatus = "A"
            
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
            
            !BRNCode = Mid(TextLine, 2968, 15)
            !MBMAppName = Mid(TextLine, 3653, 60)
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
    Dim StrSql2 As String
    Dim listrecord As Integer
    Dim listloop As Integer
    Dim ListCount As Integer
    Dim LastCoveredID As Integer
    Dim SelectAge As String
    Dim MBMCov As New ADODB.Recordset
    Dim MemberDOB As String
    Dim PayEffDate As String
    Dim AgeValue As String
    Dim FindAge As String
    Dim SuppID As String
    Dim tmpClmNonPayFr As String
    Dim tmpClmNonPayTo As String

    StrSql2 = "select * from MBM where MBMNumber = '" & Mid(TextLine, 11, 18) & "'"
    MBM.Open "MBM", medixmasterconn, adOpenKeyset, adLockOptimistic
    SuppID = Mid(TextLine, 2963, 2)
    If Trim(SuppID) = "" Then
        strsql = "Select * from MBMCoveredPersons Where MBMCNumber = '" & Mid(TextLine, 11, 18) & "' AND MBMCICBCPPNo='" & Mid(TextLine, 74, 60) & "'"
        MBMCov.Open strsql, medixmasterconn, adOpenKeyset, adLockOptimistic
        If MBMCov.recordCount = 0 Then
            MBMCov.Close
            Set MBMCov = Nothing
            strsql = "Select * from MBMCoveredPersons Where MBMCNumber = '" & Mid(TextLine, 11, 18) & "' AND MBMCICBCPPNo='" & Mid(TextLine, 300, 20) & "'"
            MBMCov.Open strsql, medixmasterconn, adOpenKeyset, adLockOptimistic
        End If
    Else
        SuppID = Mid(TextLine, 2963, 2)
        strsql = "Select * from MBMCoveredPersons Where MBMCNumber = '" & Mid(TextLine, 11, 18) & "' and MBMCCoverID = '" & SuppID & "'"
        MBMCov.Open strsql, medixmasterconn, adOpenKeyset, adLockOptimistic
    End If
    PayEffDate = MBM!MBMPayorEffDate
    LastCoveredID = GetLastCoveredID(Mid(TextLine, 11, 18))
    If MBMCov.recordCount Then
        With MBMCov
            !MBMCName = Mid(TextLine, 74, 60)
        
            !MBMCSex = Mid(TextLine, 328, 1)
            !MBMCDOB = Mid(TextLine, 320, 2) & "/" & Mid(TextLine, 322, 2) & "/" & Mid(TextLine, 324, 4)
            !MBMCICBCPPNo = Mid(TextLine, 300, 20)
'                !MBMCOccupation = Mid(TextLine, 347, 30)
            !MBMCWorkNature = Mid(TextLine, 377, 1)
            !MBMCBloodType = Mid(TextLine, 378, 3)
            !MBMCAllergic = Mid(TextLine, 381, 20)
            !MBMCSmoking = Mid(TextLine, 2965, 1)
            !MBMCAlcohol = Mid(TextLine, 2966, 1)
                    
            MemberDOB = Mid(TextLine, 320, 2) & "/" & Mid(TextLine, 322, 2) & "/" & Mid(TextLine, 324, 4)
            FindAge = GetAge(CDate(PayEffDate), CDate(MemberDOB))
            AgeValue = Mid(GetAgeBand(Mid(cboHLTCode1, 1, 1), FindAge, plancode, INSCode), 1, 2)
            !MBMCAgeband = AgeValue
            If AgeValue = "01" Then
                !MBMCAdultChild = "C"
            Else
                !MBMCAdultChild = "A"
            End If
            tmpClmNonPayFr = Mid(TextLine, 3605, 2) & "/" & Mid(TextLine, 3607, 2) & "/" & Mid(TextLine, 3609, 4)
            If IsDate(tmpClmNonPayFr) Then
                !MBMCCLMNonPayFr = tmpClmNonPayFr
            End If
            tmpClmNonPayTo = Mid(TextLine, 3613, 2) & "/" & Mid(TextLine, 3615, 2) & "/" & Mid(TextLine, 3617, 4)
            If IsDate(tmpClmNonPayTo) Then
                !MBMCCLMNonPayTo = tmpClmNonPayTo
            End If
            If Len(Mid(TextLine, 3442, 8)) Then
                !MBMCPayNo = Mid(TextLine, 3442, 8)
            End If
            If Len(Mid(TextLine, 3450, 2)) Then
                !MBMCPaySeqNo = Mid(TextLine, 3450, 2)
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

Private Sub EndtSavePrincipalDataImport(Payor As String, CoveredIDCount As Integer)
'On Error Resume Next
'On Error GoTo error1
Dim PayEffDate As String
Dim PAYExpDate As String
Dim MemberDOB As String
Dim FindAge As String
Dim InsuredCode As String
Dim plancode As String
Dim AgeValue As String
Dim AnnualLimitValue As String
Dim GetProString As String
Dim tmpPreMemberNo As String
Dim MBM As New ADODB.Recordset
Dim MBMOthers As New ADODB.Recordset
Dim TempPolDis As Double
Dim TempRenewDis As Double
Dim TmpSql As String
Dim DiscPlan As New ADODB.Recordset
        
    MBM.Open "MBM", medixmasterconn, adOpenKeyset, adLockOptimistic
    
    MBM.AddNew
    With MBM
        If Mid(TextLine, 10, 1) = "P" Then
            !MBMLastEndorseStatus = Mid(TextLine, 1, 1)
            !MBMENDtBordxDate = CDate(txtSubmissionDate)
            !MBMENDtEffDate = Mid(TextLine, 2, 2) & "/" & Mid(TextLine, 4, 2) & "/" & Mid(TextLine, 6, 4)
            !MBMMemberType = "P"
            !MBMPolicyNo = Mid(TextLine, 29, 25)
            !MBMName = Mid(TextLine, 74, 60)
            !MBMAddress1 = Mid(TextLine, 134, 30)
            !MBMAddress2 = Mid(TextLine, 164, 30)
            !MBMAddress3 = Mid(TextLine, 194, 30)
            !PAYCode = Mid(txtPAYCode, 1, 2)
            !HLTCode = Mid(cboHLTCode1, 1, 1)
            !CTYCode = Mid(TextLine, 224, 20)
            !MBMPostCode = Mid(TextLine, 244, 5)
            !STACode = Mid(TextLine, 249, 15)
            !MBMICBcPp = Mid(TextLine, 300, 20)
            !MBMDOB = Mid(TextLine, 320, 2) & "/" & Mid(TextLine, 322, 2) & "/" & Mid(TextLine, 324, 4)
            !MBMSex = Mid(TextLine, 328, 1)
            !MBMPayorEffDate = Mid(TextLine, 401, 2) & "/" & Mid(TextLine, 403, 2) & "/" & Mid(TextLine, 405, 4)
            !MBMPayorEXPDate = Mid(TextLine, 409, 2) & "/" & Mid(TextLine, 411, 2) & "/" & Mid(TextLine, 413, 4)
            !PLNCode = Mid(TextLine, 417, 4)
            !INSCode = Mid(TextLine, 421, 1)
                
            InsuredCode = Trim(!INSCode)
            plancode = Trim(!PLNCode)
            !MBMBordxDate = CDate(txtSubmissionDate)
            !MBMDataStatus = "A"
            !MBMCOVID = "00"
            !MBMValueDate = Date
            
    '------------------------------------------------------------------------------------------------------------------------
            !MBMStatus = "A"
            If Len(txtBatchNo) Then
                !MBMBatchNO = txtBatchNo
            End If
            !MBMExpiredStatus = "N"
    '------------------------------------------------------------------------------------------------------------------------
            PayEffDate = Mid(TextLine, 401, 2) & "/" & Mid(TextLine, 403, 2) & "/" & Mid(TextLine, 405, 4)
            PAYExpDate = Mid(TextLine, 409, 2) & "/" & Mid(TextLine, 411, 2) & "/" & Mid(TextLine, 413, 4)
            MemberDOB = Mid(TextLine, 320, 2) & "/" & Mid(TextLine, 322, 2) & "/" & Mid(TextLine, 324, 4)
            FindAge = GetAge(CDate(PayEffDate), CDate(MemberDOB))
            AgeValue = Mid(GetAgeBand(Mid(cboHLTCode1, 1, 1), FindAge, plancode, INSCode), 1, 2)
            !AgeCode = AgeValue
            If AgeValue = "01" Then
                !MBMAdultChild = "C"
            Else
                !MBMAdultChild = "A"
            End If
            
            txtMBMPremium = Format(GetPremium(InsuredCode, Payor, AgeValue, plancode, Mid(cboHLTCode1, 1, 1), CDate(PayEffDate)), "#,##0.00")
            txtMBMPremium = Round(InsertPremium(CDate(PAYExpDate), CDate(PayEffDate), Format(txtMBMPremium, "#,##0.00")), 0)
            BasicPrem = txtMBMPremium
            TempPolDis = IIf(IsNumeric(Mid(TextLine, 3761, 5)), Mid(TextLine, 3761, 5), 0)
            TempRenewDis = IIf(IsNumeric(Mid(TextLine, 3766, 5)), Mid(TextLine, 3766, 5), 0)
            If TempPolDis > 0 Then
                txtMBMPremium = txtMBMPremium * (100 - TempPolDis) / 100
            End If
            If TempRenewDis > 0 Then
                txtMBMPremium = txtMBMPremium * (100 - TempRenewDis) / 100
            End If
            txtMBMPremium = Format(Round(txtMBMPremium, 0), "#,##0.00")

            If AgeValue = "01" Then
                txtMBMMCOFees = Format(GetMCO(InsuredCode, Mid(cboHLTCode1, 1, 1), "C", CDate(PayEffDate), plancode), "#,##0.00")
            Else
                txtMBMMCOFees = Format(GetMCO(InsuredCode, Mid(cboHLTCode1, 1, 1), "A", CDate(PayEffDate), plancode), "#,##0.00")
            End If
            txtMBMMCOFees = Round(InsertMCOFees(CDate(PAYExpDate), CDate(PayEffDate), Format(txtMBMMCOFees, "#,##0.00")), 0)

            AnnualLimitValue = GetAnnualLimit(InsuredCode, plancode, Mid(cboHLTCode1, 1, 1), CDate(PayEffDate))
            !MBMAvailableLimit = AnnualLimitValue
                                            
            If AgeValue = "01" Then
                GetProString = GetProviderCharges("ME", InsuredCode, Mid(cboHLTCode1, 1, 1), "C")
            Else
                GetProString = GetProviderCharges("ME", InsuredCode, Mid(cboHLTCode1, 1, 1), "A")
            End If
            If Len(Trim(GetProString)) Then
                txtAccessFees = Format(GetProString, "#,##0.00")
            Else
                txtAccessFees = 0
            End If
                txtAccessFees = Round(InsertProvider(CDate(PAYExpDate), CDate(PayEffDate), Format(txtAccessFees, "#,##0.00")), 2)
                                                 
            tmpPreMemberNo = Payor & plancode & InsuredCode & GenerateMembershipNo(Payor, plancode, InsuredCode)
            !MBMNumber = tmpPreMemberNo
            txtTempMemberNo = tmpPreMemberNo
        End If
    End With
    MBM.Update
    MBM.Close
    Set MBM = Nothing

End Sub

Private Sub EndtSaveOthersDataImport()
'On Error Resume Next
'On Error GoTo error1
Dim tmpPreMemberNo As String
Dim MBMOthers As New ADODB.Recordset
Dim tmpClmNonPayFr As String
Dim tmpClmNonPayTo As String
    
    MBMOthers.Open "MBMOthers", medixmasterconn, adOpenKeyset, adLockOptimistic
    
    MBMOthers.AddNew
    With MBMOthers
        !MBMNumber = txtTempMemberNo
        !MBMAgentCode = Mid(TextLine, 54, 15)
        !MBMAllergic = Mid(TextLine, 381, 20)
        !MBMFirstMEMNo = txtTempMemberNo
        !MBMGroupCompany = Mid(TextLine, 422, 40)
        If Len(Trim(Mid(TextLine, 3761, 5))) Then
            !MBMPolicyDisc = IIf(IsNumeric(Trim(Mid(TextLine, 3761, 5))), Trim(Mid(TextLine, 3761, 5)), 0)
        End If
        If Len(Trim(Mid(TextLine, 3766, 5))) Then
            !MBMRenewalDisc = IIf(IsNumeric(Trim(Mid(TextLine, 3766, 5))), Trim(Mid(TextLine, 3766, 5)), 0)
        End If
        !MBMBranch = Mid(TextLine, 2968, 15)
        !MBMPrevPolNo = Mid(TextLine, 2983, 25)
        !MBMPrevMemNo = Mid(TextLine, 3008, 18)
        !MBMEmployeeNo = Mid(TextLine, 3026, 15)
        !MBMDepartment = Mid(TextLine, 3250, 30)
        !MBMInstallment = Mid(TextLine, 3404, 1)
        !MBMPolSubNo1 = Mid(TextLine, 3405, 10)
        !MBMPOLSubNo2 = Mid(TextLine, 3415, 10)
    
        tmpClmNonPayFr = Mid(TextLine, 3605, 2) & "/" & Mid(TextLine, 3607, 2) & "/" & Mid(TextLine, 3609, 4)
        If IsDate(tmpClmNonPayFr) Then
            !MBMCLMNonPayFr = tmpClmNonPayFr
        End If
        tmpClmNonPayTo = Mid(TextLine, 3613, 2) & "/" & Mid(TextLine, 3615, 2) & "/" & Mid(TextLine, 3617, 4)
        If IsDate(tmpClmNonPayTo) Then
            !MBMCLMNonPayTo = tmpClmNonPayTo
        End If
        If Trim(Mid(TextLine, 3621, 4)) <> "" And Trim(Mid(TextLine, 3625, 8)) <> "" Then
            !MBMNewPlanCode = Mid(TextLine, 3621, 4)
            !MBMNewPolEffDate = Mid(TextLine, 3625, 2) & "/" & Mid(TextLine, 3627, 2) & "/" & Mid(TextLine, 3629, 4)
        End If
        If Trim(Mid(TextLine, 3633, 20)) <> "" Then
            !MBMENDtRefNo = Mid(TextLine, 3633, 20)
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
Private Sub EndtSaveCovPersonDataImport(Payor As String, CoveredIDCount As Integer)
'On Error GoTo error1
Dim PayEffDate As String
Dim MemberDOB As String
Dim FindAge As String
Dim AgeValue As String
Dim GetProString As String
Dim tmpPreMemberNo As String
Dim MBMCover As New ADODB.Recordset
        
     MBMCover.Open "MBMCoveredPersons", medixmasterconn, adOpenKeyset, adLockOptimistic
    ' medixmasterconn.BeginTrans
     CoveredIDCount = Format(CoveredIDCount + 1, "00")
     
     MBMCover.AddNew
     With MBMCover
        !MBMCRELCode = Mid(TextLine, 10, 1)
        !MBMCCoverID = CoveredIDCount
        !MBMCMemberType = "C"
        !MBMCNumber = Trim(txtTempMemberNo)
        !MBMCName = Mid(TextLine, 74, 60)
        !MBMCStatus = "A"
        !MBMCName = Trim(Mid(TextLine, 74, 60))
        !MBMCICBCPPNo = Trim(Mid(TextLine, 300, 20))
        !MBMCDOB = Mid(TextLine, 320, 2) & "/" & Mid(TextLine, 322, 2) & "/" & Mid(TextLine, 324, 4)
        !MBMCSex = Mid(TextLine, 328, 1)
        !MBMCAllergic = Mid(TextLine, 381, 20)
        !MBMCPolicyDisc = IIf(IsNumeric(Trim(Mid(TextLine, 3761, 5))), Trim(Mid(TextLine, 3761, 5)), 0)
        !MBMCRenewalDisc = IIf(IsNumeric(Trim(Mid(TextLine, 3766, 5))), Trim(Mid(TextLine, 3766, 5)), 0)
        !MBMCDOB = Mid(TextLine, 320, 2) & "/" & Mid(TextLine, 322, 2) & "/" & Mid(TextLine, 324, 4)
        !MBMCSex = Mid(TextLine, 328, 1)
        PayEffDate = Mid(TextLine, 401, 2) & "/" & Mid(TextLine, 403, 2) & "/" & Mid(TextLine, 405, 4)
        MemberDOB = Mid(TextLine, 320, 2) & "/" & Mid(TextLine, 322, 2) & "/" & Mid(TextLine, 324, 4)
        FindAge = GetAge(CDate(PayEffDate), CDate(MemberDOB))
        AgeValue = Mid(GetAgeBand(Mid(cboHLTCode1, 1, 1), FindAge, plancode, INSCode), 1, 2)
        !MBMCAgeband = AgeValue
        If AgeValue = "01" Then
            !MBMCAdultChild = "C"
        Else
            !MBMCAdultChild = "A"
        End If
        
'        If Mid(cboSuppAnnualLimit, 1, 1) = "S" Then
            'TmpSql = "select * from DiscPlan where PLNCode = '" & Trim(PlanCode) & "'"
            'DiscPlan.Open TmpSql, medixmasterconnMaint, adOpenKeyset, adLockOptimistic
            
'            If AnnLmt!SuppLimitStatus = "B" Then
'                If CoveredIDCount = "1" Then
'                    tempSPremium = Format(GetSPremium(InsuredCode, Payor, AgeValue, PlanCode, Mid(cboHLTCode1, 1, 1), CDate(PayEffDate)), "#,##0.00")
'                    tempSPremium = InsertPremium(CDate(PayorEXPDate), CDate(PayEffDate), Format(tempSPremium, "#,##0.00"))
'                    tempSAnnLmt = Format(GetSuppLimit(InsuredCode, PlanCode, Mid(cboHLTCode1, 1, 1), CDate(PayEffDate)), "#,##0.00")
'                    !MBMCAnnualLimit = tempSAnnLmt
'                    !MBMCAvailableLimit = tempSAnnLmt
'                    !MBMCBasicPremium = Format(Round(tempSPremium, 0), "#,##0.00")
'                    'If DiscPlan.recordCount Then
'                        If TempPolDis > 0 Then
'                            tempSPremium = tempSPremium * (100 - TempPolDis) / 100
'                        End If
'                        If TempRenewDis > 0 Then
'                            tempSPremium = tempSPremium * (100 - TempRenewDis) / 100
'                        End If
'                    'End If
'                    tempSPremium = Format(Round(tempSPremium, 0), "#,##0.00")
'                    !MBMCPremium = tempSPremium
'                End If
'            ElseIf AnnLmt!SuppLimitStatus = "C" Then 'Cater for SuppLimitStatus is "C"

'                tempSPremium = Format(GetSPremium(InsuredCode, Payor, AgeValue, PlanCode, Mid(cboHLTCode1, 1, 1), CDate(PayEffDate)), "#,##0.00")
'                tempSPremium = InsertPremium(CDate(PayorEXPDate), CDate(PayEffDate), Format(tempSPremium, "#,##0.00"))

'                tempSAnnLmt = Format(GetSuppLimit(InsuredCode, PlanCode, Mid(cboHLTCode1, 1, 1), CDate(PayEffDate)), "#,##0.00")
'                !MBMCAnnualLimit = tempSAnnLmt
'                !MBMCAvailableLimit = tempSAnnLmt
'                !MBMCBasicPremium = Format(Round(tempSPremium, 0), "#,##0.00")
'                'If DiscPlan.recordCount Then
'                    If TempPolDis > 0 Then
'                        tempSPremium = tempSPremium * (100 - TempPolDis) / 100
'                    End If
'                    If TempRenewDis > 0 Then
'                        tempSPremium = tempSPremium * (100 - TempRenewDis) / 100
'                    End If
'                'End If
'                tempSPremium = Format(Round(tempSPremium, 0), "#,##0.00")
'                !MBMCPremium = tempSPremium
'                TmpClmNonPayfr = Mid(TextLine, 3576, 2) & "/" & Mid(TextLine, 3578, 2) & "/" & Mid(TextLine, 3580, 4)
'                If IsDate(TmpClmNonPayfr) Then
'                    !MBMCCLMNonPayFr = TmpClmNonPayfr
'                End If
'                TmpClmNonPayTo = Mid(TextLine, 3584, 2) & "/" & Mid(TextLine, 3586, 2) & "/" & Mid(TextLine, 3588, 4)
'                If IsDate(TmpClmNonPayTo) Then
'                    !MBMCCLMNonPayTo = TmpClmNonPayTo
'                End If
'            Else
'                !MBMCAnnualLimit = tempAnnLmt
'                !MBMCAvailableLimit = tempAnnLmt
'            End If
'            'DiscPlan.Close
'            'Set DiscPlan = Nothing
'        End If
'' ******* Clinical Supp Limit
'        CLNPlnSql = "SELECT AnnLmtIND, VisLmtIND from PLN where PLNCode='" & Trim(CisPlnCode) & _
'                    "' AND HLTCode='" & Mid(cboHLTCode1, 1, 1) & "'"
'        CLNPln.Open CLNPlnSql, MediConnCISDB, adOpenKeyset, adLockOptimistic
'
'        CISCovSuppLMtStat = ""
'        If CLNPln.recordCount Then
'            If Trim(CLNPln!ANNLmtInd) = "Y" Then
'                ClnAnnLmtSql = "Select SuppLimitStatus from AnnualLimit where PLNCode = '" & Trim(CisPlnCode) & _
'                               "' AND HLTCode='" & Mid(cboHLTCode1, 1, 1) & "' AND INSCode='" & Trim(CisInsType) & "'"
'                CLnAnnLmt.Open ClnAnnLmtSql, MediConnCISDB, adOpenKeyset, adLockOptimistic
'                CISCovSuppLMtStat = CLnAnnLmt!SuppLimitStatus
'                CLnAnnLmt.Close
'                Set CLnAnnLmt = Nothing
'            End If
'        End If
'        CLNPln.Close
'        Set CLNPln = Nothing
'        If Trim(CISCovSuppLMtStat) = "B" Then
'            If CoveredIDCount = "1" Then
'                CISSuppPrem = Format(GetCLNSPremium(CisInsType, Payor, AgeValue, CisPlnCode, Mid(cboHLTCode1, 1, 1), CDate(CisEffDate)), "#,##0.00")
'                CISSuppBPrem = CISSuppPrem
'                CISSuppPrem = InsertPremium(CDate(CisExpDate), CDate(CisEffDate), Format(CISSuppPrem, "#,##0.00"))
'                CISSuppPrem = Format(Round(CISSuppPrem, 0), "#,##0.00")
'
'                If AgeValue = "01" Then
'                    CISSuppMCO = Format(GetCLNSMCO(CisInsType, Mid(cboHLTCode1, 1, 1), "C", CDate(CisEffDate), CisPlnCode), "#,##0.00")
'                Else
'                    CISSuppMCO = Format(GetCLNSMCO(CisInsType, Mid(cboHLTCode1, 1, 1), "A", CDate(CisEffDate), CisPlnCode), "#,##0.00")
'                End If
'                CISSuppBMCO = CISSuppMCO
'                CISSuppMCO = InsertMCOFees(CDate(CisExpDate), CDate(CisEffDate), Format(CISSuppMCO, "#,##0.00"))
'                CISSuppMCO = Round(CISSuppMCO)

'                CISSuppAnnLmt = Format(GetCLNSuppLimit(CisInsType, CisPlnCode, Mid(cboHLTCode1, 1, 1), CDate(CisEffDate)), "#,##0.00")
'                CISSuppVisLmt = GetCLNSVisitLimit(CisInsType, CisPlnCode, Mid(cboHLTCode1, 1, 1), CDate(CisEffDate))
'                !MBMCCLNMCO = CISSuppMCO
'                !MBMCCLNBasicMCO = CISSuppBMCO
'                !MBMCCLNPrem = CISSuppPrem
'                !MBMCCLNBasicPrem = CISSuppBPrem
'                !MBMCCLNBalLmt = CISSuppAnnLmt
'                !MBMCCLNVisitLmt = CISSuppAnnLmt
'            End If
'        ElseIf Trim(CISCovSuppLMtStat) = "C" Then  'Cater for SuppLimitStatus is "C"
'                CISSuppPrem = Format(GetCLNSPremium(CisInsType, Payor, AgeValue, CisPlnCode, Mid(cboHLTCode1, 1, 1), CDate(CisEffDate)), "#,##0.00")
'                CISSuppBPrem = CISSuppPrem
'                CISSuppPrem = InsertPremium(CDate(CisExpDate), CDate(CisEffDate), Format(CISSuppPrem, "#,##0.00"))
'                CISSuppPrem = Format(Round(CISSuppPrem, 0), "#,##0.00")
'
'                If AgeValue = "01" Then
'                    CISSuppMCO = Format(GetCLNSMCO(CisInsType, Mid(cboHLTCode1, 1, 1), "C", CDate(CisEffDate), CisPlnCode), "#,##0.00")
'                Else
'                    CISSuppMCO = Format(GetCLNSMCO(CisInsType, Mid(cboHLTCode1, 1, 1), "A", CDate(CisEffDate), CisPlnCode), "#,##0.00")
'                End If
'                CISSuppBMCO = CISSuppMCO
'                CISSuppMCO = InsertMCOFees(CDate(CisExpDate), CDate(CisEffDate), Format(CISSuppMCO, "#,##0.00"))
'                CISSuppMCO = Round(CISSuppMCO)

'                CISSuppAnnLmt = Format(GetCLNSuppLimit(CisInsType, CisPlnCode, Mid(cboHLTCode1, 1, 1), CDate(CisEffDate)), "#,##0.00")
'                CISSuppVisLmt = GetCLNSVisitLimit(CisInsType, CisPlnCode, Mid(cboHLTCode1, 1, 1), CDate(CisEffDate))
'                !MBMCCLNMCO = CISSuppMCO
'                !MBMCCLNBasicMCO = CISSuppBMCO
'                !MBMCCLNPrem = CISSuppPrem
'                !MBMCCLNBasicPrem = CISSuppBPrem
'                !MBMCCLNBalLmt = CISSuppAnnLmt
'                !MBMCCLNVisitLmt = CISSuppAnnLmt
'        End If
    End With
    MBMCover.Update
    MBMCover.Close
End Sub

Private Sub SaveCISDataImport(BordxFmtType As String)
Dim CisWarranty, PolNO As String
Dim CisPrem, CISBasicPrem, CisAgeValue As String
Dim CisMCO, CisBasicMco, CisVisitLmt, CisBalLmt As String
Dim MemberDOB, CISPayor As String
Dim FindAge As String
Dim AgeValue As String
Dim TmpMBMType, tmpCisEffDate, tmpCisExpDate As String
Dim tmpClnStatus, tmpClnDAtaStatus, tmpClnExpStatus As String
Dim StrCLNSql As String
Dim CisDentalLmt As Double
Dim CisSpecLmt As Double
Dim CisMepLmt As Double
Dim CisAnnLmtType As String
Dim MBMClnOthersREc As New ADODB.Recordset
Dim tmpCardType As String

    CISPayor = IIf(Trim(CISDCStatus) = "Y", CISDCPayor, Mid(txtPAYCode, 1, 2))
    If Trim(BordxFmtType) = "N" Then
        MemberDOB = Mid(TextLine, 293, 2) & "/" & Mid(TextLine, 295, 2) & "/" & Mid(TextLine, 297, 4)
        PolNO = Trim(Mid(TextLine, 2, 25))
        CisWarranty = Trim(Mid(TextLine, 6729, 2500))
        tmpCardType = Trim(Mid(TextLine, 9229, 20))
    ElseIf Trim(BordxFmtType) = "E" Then
        MemberDOB = Mid(TextLine, 320, 2) & "/" & Mid(TextLine, 322, 2) & "/" & Mid(TextLine, 324, 4)
        PolNO = Trim(Mid(TextLine, 29, 25))
        CisWarranty = Trim(Mid(TextLine, 6758, 2500))
        tmpCardType = Trim(Mid(TextLine, 9258, 20))
    End If
    tmpClnStatus = "A"
    tmpClnDAtaStatus = "A"
    tmpClnExpStatus = "N"
    
    FindAge = GetAge(CDate(CisEffDate), CDate(MemberDOB))
    AgeValue = Mid(GetAgeBand(Mid(cboHLTCode1, 1, 1), FindAge, CisPlnCode, CisInsType), 1, 2)
    
    CisPrem = Format(GetCLNPremium(CisInsType, CISPayor, AgeValue, CisPlnCode, Mid(cboHLTCode1, 1, 1), CDate(CisEffDate)), "#,##0.00")
    CISBasicPrem = CisPrem
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
    
    CisDentalLmt = 0
    CisSpecLmt = 0
    CisMepLmt = 0
    
    If Trim(CISPDentStatus) = "S" Then 'stand alone dental limit
        CisAnnLmtType = "PDentalAnnLimit"
        CisDentalLmt = GetCISLimit(CisInsType, CisPlnCode, Mid(cboHLTCode1, 1, 1), CDate(CisEffDate), CisAnnLmtType)
    End If
    If Trim(CISPSpecStatus) = "S" Then 'stand alone Spec limit
        CisAnnLmtType = "PSpecAnnLimit"
        CisSpecLmt = GetCISLimit(CisInsType, CisPlnCode, Mid(cboHLTCode1, 1, 1), CDate(CisEffDate), CisAnnLmtType)
    End If
    If Trim(CISPMepStatus) = "S" Then 'stand alone MEP limit
        CisAnnLmtType = "PMepAnnLimit"
        CisMepLmt = GetCISLimit(CisInsType, CisPlnCode, Mid(cboHLTCode1, 1, 1), CDate(CisEffDate), CisAnnLmtType)
    End If
    If Trim(CISCovSuppLMtStat) = "A" Then
        TmpMBMType = "P"
    Else
        TmpMBMType = "Q"
    End If
    
    CisPrem = CDbl(IIf(IsNumeric(CisPrem) = True, CisPrem, 0))
    CISBasicPrem = CDbl(IIf(IsNumeric(CISBasicPrem) = True, CISBasicPrem, 0))
    CisBasicMco = CDbl(IIf(IsNumeric(CisBasicMco) = True, CisBasicMco, 0))
    CisMCO = CDbl(IIf(IsNumeric(CisMCO) = True, CisMCO, 0))
    CisBalLmt = CDbl(IIf(IsNumeric(CisBalLmt) = True, CisBalLmt, 0))
    CisVisitLmt = CDbl(IIf(IsNumeric(CisVisitLmt) = True, CisVisitLmt, 0))
    CisDentalLmt = CDbl(IIf(IsNumeric(CisDentalLmt) = True, CisDentalLmt, 0))
    CisSpecLmt = CDbl(IIf(IsNumeric(CisSpecLmt) = True, CisSpecLmt, 0))
    CisMepLmt = CDbl(IIf(IsNumeric(CisMepLmt) = True, CisMepLmt, 0))
    
    tmpCisEffDate = Format(CisEffDate, "yyyy/mm/dd")
    tmpCisExpDate = Format(CisExpDate, "yyyy/mm/dd")
    PolNO = Replace(PolNO, "'", "''")
    CisWarranty = Replace(CisWarranty, "'", "''")
'    StrCLNSql = "Insert Into MBMCLNOthers(MBMCLNMemberType, PAYCode, MBMNumber," & _
'                "MBMPolNumber, MBMCLNStatus, MBMCLNDataStatus, MBMCLNExpiryStatus, " & _
'                "PLNType, InsType, MBMCLNWarranty, MBMCLNBasicPrem, MBMCLNPremium,  " & _
'                "MBMCLNBasicMCO, MBMCLNMCOFee, MBMCLNBalLmt, MBMCLNVisitLmt, " & _
'                "MBMDentalLmt, MBMSpecLmt, MBMMEPLmt, " & _
'                "MBMPolEffDate, MBMPolExpDate)" & _
'                "Values('" & TmpMBMType & "','" & Trim(CISPayor) & "','" & Trim(txtTempMemberNo) & "'," & _
'                "'" & PolNO & "','" & tmpClnStatus & "','" & tmpClnDAtaStatus & "','" & tmpClnExpStatus & "', " & _
'                "'" & Trim(CisPlnCode) & "','" & CisInsType & "','" & CisWarranty & "'," & CISBasicPrem & "," & CisPrem & ", " & _
'                "" & CisBasicMco & "," & CisMCO & "," & CisBalLmt & "," & CisVisitLmt & "," & _
'                "" & CisDentalLmt & "," & CisSpecLmt & "," & CisMepLmt & "," & _
'                "'" & tmpCisEffDate & "','" & tmpCisExpDate & "')"
'    medixmasterconn.Execute StrCLNSql
    StrCLNSql = "EXEC Upload_MBMClnOthers '" & TmpMBMType & "','" & Trim(CISPayor) & "','" & Trim(txtTempMemberNo) & "'," & _
                "'" & PolNO & "','" & tmpClnStatus & "','" & tmpClnDAtaStatus & "','" & tmpClnExpStatus & "', " & _
                "'" & Trim(CisPlnCode) & "','" & CisInsType & "','" & CisWarranty & "'," & CISBasicPrem & "," & CisPrem & ", " & _
                "" & CisBasicMco & "," & CisMCO & "," & CisBalLmt & "," & CisVisitLmt & "," & _
                "" & CisDentalLmt & "," & CisSpecLmt & "," & CisMepLmt & "," & _
                "'" & tmpCisEffDate & "','" & tmpCisExpDate & "','" & tmpCardType & "'"
    medixmasterconn.Execute StrCLNSql

End Sub

Private Sub SaveMBMTwo(BordxFmtType As String)
Dim TmpTelNoH, TmpTelNoO, TmpTelNoM, TmpSRVCodeList As String
Dim TmpNatCode, TmpLgnCode, TmpClinical As String
Dim TmpSalutation, TmpAppName, TmpIcBcPp2nd As String
Dim MBMTwoWSql As String
Dim MBMTwoREc As New ADODB.Recordset

    TmpAppName = " "
    TmpIcBcPp2nd = " "
       
    If Trim(BordxFmtType) = "N" Then
         TmpSalutation = Trim(Mid(TextLine, 42, 5))
         TmpTelNoH = Trim(Mid(TextLine, 237, 12))
         TmpTelNoM = Trim(Mid(TextLine, 249, 12))
         TmpTelNoO = Trim(Mid(TextLine, 261, 12))
         TmpNatCode = Trim(Mid(TextLine, 302, 10))
         TmpLgnCode = Trim(Mid(TextLine, 318, 1))
         TmpAppName = Trim(Mid(TextLine, 3624, 60))
         TmpIcBcPp2nd = Trim(Mid(TextLine, 3742, 20))
     ElseIf Trim(BordxFmtType) = "E" Then
         TmpSalutation = Trim(Mid(TextLine, 69, 5))
         TmpTelNoH = Trim(Mid(TextLine, 264, 12))
         TmpTelNoM = Trim(Mid(TextLine, 276, 12))
         TmpTelNoO = Trim(Mid(TextLine, 288, 12))
         TmpNatCode = Trim(Mid(TextLine, 329, 10))
         TmpLgnCode = Trim(Mid(TextLine, 345, 1))
         TmpAppName = Trim(Mid(TextLine, 3653, 60))
         TmpIcBcPp2nd = Trim(Mid(TextLine, 3771, 20))
     End If
     TmpClinical = ""
     If Trim(CisPlnCode) <> "" Then
         TmpClinical = "Y"
     End If
     TmpAppName = Replace(TmpAppName, "'", "''")
     TmpSalutation = Replace(TmpSalutation, "'", "''")
     TmpNatCode = Replace(TmpNatCode, "'", "''")
     MBMTwoWSql = "EXEC Upload_MBMTwo '" & Trim(txtTempMemberNo) & " ','" & TmpSalutation & "','" & TmpTelNoH & "','" & TmpTelNoM & "'," & _
                 "'" & TmpTelNoO & "','" & TmpNatCode & "','" & TmpLgnCode & "','" & TmpSRVCodeList & "', '" & Logon & "'," & _
                 "'" & Logon & "','" & TmpClinical & "','" & TmpAppName & "','" & TmpIcBcPp2nd & "','A'"
     medixmasterconn.Execute MBMTwoWSql
End Sub

Private Sub SaveMBMOthersTwo(BordxFmtType As String)
Dim TmpWeight, TmpHeight, TmpWorkNature As String
Dim TmpBlood, TmpOccupation, TmpSmoking, TmpAlcohol As String
Dim TmpMaritalStatus, TmpBTBG As String
Dim tmpGrpCompAdd1, TmpGrpCompAdd2, TmpGrpCompAdd3, TmpGrpCompAdd4 As String
Dim POLIssuedDate As String
Dim TmpRiskNo, TmpTransNo, TmpClientNo, TmpClientID As String
Dim TmpPayNo, TmpPaySeqNo As String
Dim tmpRBAmt As String
Dim TmpPrevInsCom, TmpPrevPlnCode, TmpPayRemarks As String
Dim TmpCoPayRB, TmpstaffDec, TmpMBMBankAccNo As String
Dim tmpGenWP, TmpPreExWP, TmpSpeIllWP As String
Dim tmpGenEndWP, TmpPreExEndWP, TmpSpeIllEndWP As String
Dim TmpPayPremNet As Double
Dim TmpPayMcoNet As Double
Dim TmpPayMCOGross As Double
Dim TmpPayPremGross As Double
Dim TmpLoadingPrem As Double
Dim TmpFamDiscAmt As Double
Dim TmpRenewDiscAmt As Double
Dim TmpExclusion As String
Dim MBMOTwoWSql As String
'Dim StrDateSql As String
'Dim StrWDateSql As String
Dim MBMOthers2Rec As New ADODB.Recordset
Dim tmpPolDate As String
Dim TmpAllergic As String

    tmpGenWP = ""
    TmpPreExWP = ""
    TmpSpeIllWP = ""
    TmpMBMBankAccNo = ""
    TmpstaffDec = ""
    If Trim(BordxFmtType) = "N" Then
        tmpGrpCompAdd1 = Trim(Mid(TextLine, 3251, 30))
        TmpGrpCompAdd2 = Trim(Mid(TextLine, 3281, 30))
        TmpGrpCompAdd3 = Trim(Mid(TextLine, 3311, 30))
        TmpGrpCompAdd4 = Trim(Mid(TextLine, 3341, 30))
        POLIssuedDate = Mid(TextLine, 3396, 2) & "/" & Mid(TextLine, 3398, 2) & "/" & Mid(TextLine, 3400, 4)
        TmpRiskNo = Trim(Mid(TextLine, 3404, 4))
        TmpTransNo = Trim(Mid(TextLine, 3408, 5))
        TmpPayNo = Trim(Mid(TextLine, 3413, 8))
        TmpPaySeqNo = Trim(Mid(TextLine, 3421, 2))
        TmpClientNo = Trim(Mid(TextLine, 3423, 8))
        TmpClientID = Trim(Mid(TextLine, 3431, 2))
        tmpRBAmt = Trim(Mid(TextLine, 3448, 8))
        TmpPrevInsCom = Trim(Mid(TextLine, 3456, 60))
        TmpPrevPlnCode = Trim(Mid(TextLine, 3433, 15))
        TmpPayRemarks = Trim(Mid(TextLine, 3516, 60))
        TmpCoPayRB = Trim(Mid(TextLine, 3762, 2))
        TmpPayPremNet = IIf(IsNumeric(Trim(Mid(TextLine, 3684, 12))), Trim(Mid(TextLine, 3684, 12)), 0)
        TmpPayPremGross = Format(IIf(IsNumeric(Trim(Mid(TextLine, 3696, 12))), Trim(Mid(TextLine, 3696, 12)), 0), "#,##0.00")
        TmpPayMcoNet = IIf(IsNumeric(Trim(Mid(TextLine, 3708, 12))), Trim(Mid(TextLine, 3708, 12)), 0)
        TmpPayMCOGross = IIf(IsNumeric(Trim(Mid(TextLine, 3720, 12))), Trim(Mid(TextLine, 3720, 12)), 0)
        TmpLoadingPrem = IIf(IsNumeric(Trim(Mid(TextLine, 3764, 12))), Trim(Mid(TextLine, 3764, 12)), 0)
        TmpFamDiscAmt = IIf(IsNumeric(Trim(Mid(TextLine, 3776, 12))), Trim(Mid(TextLine, 3776, 12)), 0)
        TmpRenewDiscAmt = IIf(IsNumeric(Trim(Mid(TextLine, 3788, 12))), Trim(Mid(TextLine, 3788, 12)), 0)
        TmpAllergic = Trim(Mid(TextLine, 354, 20))

        TmpWeight = Trim(Mid(TextLine, 312, 3))
        TmpHeight = Trim(Mid(TextLine, 315, 3))
        TmpWorkNature = Mid(TextLine, 350, 1)
        TmpBlood = Trim(Mid(TextLine, 351, 3))
        TmpOccupation = Trim(Mid(TextLine, 320, 30))
        TmpSmoking = Mid(TextLine, 2936, 1)
        TmpAlcohol = Mid(TextLine, 2937, 1)
        TmpMaritalStatus = Mid(TextLine, 2938, 1)
        TmpExclusion = Trim(Mid(TextLine, 436, 2500))
        TmpMBMBankAccNo = Trim(Mid(TextLine, 3805, 25))
    ElseIf Trim(BordxFmtType) = "E" Then
        TmpWeight = Trim(Mid(TextLine, 339, 3))
        TmpHeight = Trim(Mid(TextLine, 342, 3))
        TmpWorkNature = Mid(TextLine, 377, 1)
        TmpBlood = Trim(Mid(TextLine, 378, 3))
        TmpOccupation = Trim(Mid(TextLine, 347, 30))
        TmpSmoking = Mid(TextLine, 2965, 1)
        TmpAlcohol = Mid(TextLine, 2966, 1)
        TmpMaritalStatus = Mid(TextLine, 2967, 1)
        TmpExclusion = Trim(Mid(TextLine, 463, 2500))
        tmpGrpCompAdd1 = Trim(Mid(TextLine, 3280, 30))
        TmpGrpCompAdd2 = Trim(Mid(TextLine, 3310, 30))
        TmpGrpCompAdd3 = Trim(Mid(TextLine, 3340, 30))
        TmpGrpCompAdd4 = Trim(Mid(TextLine, 3370, 30))
        POLIssuedDate = Mid(TextLine, 3425, 2) & "/" & Mid(TextLine, 3427, 2) & "/" & Mid(TextLine, 3429, 4)
        TmpRiskNo = Trim(Mid(TextLine, 3433, 4))
        TmpTransNo = Trim(Mid(TextLine, 3437, 5))
        TmpPayNo = Trim(Mid(TextLine, 3442, 8))
        TmpPaySeqNo = Trim(Mid(TextLine, 3450, 2))
        TmpClientNo = Trim(Mid(TextLine, 3452, 8))
        TmpClientID = Trim(Mid(TextLine, 3460, 2))
        tmpRBAmt = Mid(TextLine, 3477, 8)
        TmpPrevInsCom = Trim(Mid(TextLine, 3585, 60))
        TmpPrevPlnCode = Trim(Mid(TextLine, 3462, 15))
        TmpPayRemarks = Trim(Mid(TextLine, 3545, 60))
        TmpAllergic = Trim(Mid(TextLine, 381, 20))
        TmpCoPayRB = Mid(TextLine, 3791, 2)
        TmpPayPremNet = IIf(IsNumeric(Trim(Mid(TextLine, 3713, 12))), Trim(Mid(TextLine, 3713, 12)), 0)
        TmpPayPremGross = IIf(IsNumeric(Trim(Mid(TextLine, 3725, 12))), Trim(Mid(TextLine, 3725, 12)), 0)
        TmpPayMcoNet = IIf(IsNumeric(Trim(Mid(TextLine, 3737, 12))), Trim(Mid(TextLine, 3737, 12)), 0)
        TmpPayMCOGross = IIf(IsNumeric(Trim(Mid(TextLine, 3749, 12))), Trim(Mid(TextLine, 3749, 12)), 0)
        TmpLoadingPrem = IIf(IsNumeric(Trim(Mid(TextLine, 3793, 12))), Trim(Mid(TextLine, 3793, 12)), 0)
        TmpFamDiscAmt = IIf(IsNumeric(Trim(Mid(TextLine, 3805, 12))), Trim(Mid(TextLine, 3805, 12)), 0)
        TmpRenewDiscAmt = IIf(IsNumeric(Trim(Mid(TextLine, 3817, 12))), Trim(Mid(TextLine, 3817, 12)), 0)
        TmpMBMBankAccNo = Trim(Mid(TextLine, 3834, 25))
    End If
    If Mid(CboBack2Back, 1, 1) = "Y" Then
        TmpBTBG = Mid(CboBack2Back, 1, 1)
    End If
        
    If IsDate(POLIssuedDate) Then
        tmpPolDate = Format(POLIssuedDate, "YYYY/MM/DD")
    End If
    tmpRBAmt = CDbl(IIf(IsNumeric(tmpRBAmt) = True, tmpRBAmt, 0))
    tmpGenWP = CDbl(IIf(IsNumeric(tmpGenWP) = True, tmpGenWP, 0))
    TmpPreExWP = CDbl(IIf(IsNumeric(TmpPreExWP) = True, TmpPreExWP, 0))
    TmpSpeIllWP = CDbl(IIf(IsNumeric(TmpSpeIllWP) = True, TmpSpeIllWP, 0))

    TmpOccupation = Replace(TmpOccupation, "'", "''")
    tmpGrpCompAdd1 = Replace(tmpGrpCompAdd1, "'", "''")
    TmpGrpCompAdd2 = Replace(TmpGrpCompAdd2, "'", "''")
    TmpGrpCompAdd3 = Replace(TmpGrpCompAdd3, "'", "''")
    TmpGrpCompAdd4 = Replace(TmpGrpCompAdd4, "'", "''")
    TmpPrevInsCom = Replace(TmpPrevInsCom, "'", "''")
    TmpExclusion = Replace(TmpExclusion, "'", "''")
    TmpMBMBankAccNo = Replace(TmpMBMBankAccNo, "'", "''")
    TmpAllergic = Replace(TmpAllergic, "'", "''")

    If IsDate(POLIssuedDate) Then
        MBMOTwoWSql = "EXEC Upload_MBMOthersTwo '" & Trim(txtTempMemberNo) & "', '" & TmpWeight & "', '" & TmpHeight & "','" & TmpWorkNature & "', " & _
                      "'" & TmpBlood & "','" & TmpOccupation & "','" & TmpSmoking & "','" & TmpAlcohol & "','" & TmpMaritalStatus & "', " & _
                      "'" & tmpGrpCompAdd1 & "','" & TmpGrpCompAdd2 & "','" & TmpGrpCompAdd3 & "','" & TmpGrpCompAdd4 & "', " & _
                      "'" & TmpRiskNo & "','" & TmpTransNo & "','" & TmpPayNo & "','" & TmpPaySeqNo & "','" & TmpMBMBankAccNo & "'," & _
                      "'" & TmpClientNo & "','" & TmpClientID & "'," & tmpRBAmt & "," & _
                      "'" & TmpPrevInsCom & "','" & TmpPrevPlnCode & "','" & TmpPayRemarks & "','" & TmpCoPayRB & "','" & TmpstaffDec & "','" & TmpExclusion & "', " & _
                      "'" & TmpBTBG & "'," & TmpPayPremNet & "," & TmpPayMcoNet & "," & TmpPayMCOGross & "," & TmpLoadingPrem & ", " & _
                      "" & TmpFamDiscAmt & "," & TmpRenewDiscAmt & ", " & _
                      "" & tmpGenWP & "," & TmpPreExWP & "," & TmpSpeIllWP & " ," & TmpPayPremGross & ",'" & tmpPolDate & "','" & TmpAllergic & "'"
    Else
        MBMOTwoWSql = "EXEC Upload_MBMOthersTwo_WOPol '" & Trim(txtTempMemberNo) & "', '" & TmpWeight & "', '" & TmpHeight & "','" & TmpWorkNature & "', " & _
                      "'" & TmpBlood & "','" & TmpOccupation & "','" & TmpSmoking & "','" & TmpAlcohol & "','" & TmpMaritalStatus & "', " & _
                      "'" & tmpGrpCompAdd1 & "','" & TmpGrpCompAdd2 & "','" & TmpGrpCompAdd3 & "','" & TmpGrpCompAdd4 & "', " & _
                      "'" & TmpRiskNo & "','" & TmpTransNo & "','" & TmpPayNo & "','" & TmpPaySeqNo & "','" & TmpMBMBankAccNo & "'," & _
                      "'" & TmpClientNo & "','" & TmpClientID & "'," & tmpRBAmt & "," & _
                      "'" & TmpPrevInsCom & "','" & TmpPrevPlnCode & "','" & TmpPayRemarks & "','" & TmpCoPayRB & "','" & TmpstaffDec & "','" & TmpExclusion & "', " & _
                      "'" & TmpBTBG & "'," & TmpPayPremNet & "," & TmpPayMcoNet & "," & TmpPayMCOGross & "," & TmpLoadingPrem & ", " & _
                      "" & TmpFamDiscAmt & "," & TmpRenewDiscAmt & ", " & _
                      "" & tmpGenWP & "," & TmpPreExWP & "," & TmpSpeIllWP & " ," & TmpPayPremGross & ",'" & TmpAllergic & "'"
    End If
    medixmasterconn.Execute MBMOTwoWSql

End Sub

Private Sub SaveCovPersonsTwo(CovCnt As Integer, BordxFmtType As String)
Dim TmpBlood, TmpOccupation, TmpSmoking, TmpAlcohol, TmpWorkNature  As String
Dim TmpSalutation, TmpMaritalStatus, TmpAllergic As String
Dim TmpRiskNo, TmpTransNo, TmpClientNo, TmpClientID As String
Dim TmpPayNo, TmpPaySeqNo As String
Dim tmpRBAmt As String
Dim TmpPrevInsCom, TmpPrevPlnCode, TmpPayRemarks As String
Dim TmpCoPayRB, TmpstaffDec As String
Dim tmpGenWP, TmpPreExWP, TmpSpeIllWP As String
Dim tmpGenEndWP, TmpPreExEndWP, TmpSpeIllEndWP As String
Dim TmpPayorPremNet As Double
Dim TmpPayorMcoNet As Double
Dim TmpPayorMCoGross As Double
Dim TmpPayorPremGross As Double
Dim LoadingPrem As Double
Dim FamDiscAmt As Double
Dim RenewDiscAmt As Double
Dim TmpIcBcPp2nd, TmpCovID As String
Dim TmpSuppExclusion As String
Dim Supp2WSql As String
'Dim StrDateSql, StrWDateSql As String
Dim MBMCov2Rec As New ADODB.Recordset
Dim TmpCUndExcess As String

    tmpGenWP = ""
    TmpPreExWP = ""
    TmpSpeIllWP = ""
    If Trim(BordxFmtType) = "N" Then
        TmpOccupation = Mid(TextLine, 320, 30)
        TmpWorkNature = Mid(TextLine, 350, 1)
        TmpBlood = Mid(TextLine, 351, 3)
        TmpAllergic = Mid(TextLine, 354, 20)
        TmpSalutation = Trim(Mid(TextLine, 42, 5))
        TmpSmoking = Mid(TextLine, 2936, 1)
        TmpAlcohol = Mid(TextLine, 2937, 1)
        TmpSuppExclusion = Trim(Mid(TextLine, 436, 2500))
        TmpPrevPlnCode = Trim(Mid(TextLine, 3433, 15))
        TmpPrevInsCom = Trim(Mid(TextLine, 3456, 60))
        TmpPayRemarks = Trim(Mid(TextLine, 3516, 60))
        TmpPayNo = Trim(Mid(TextLine, 3413, 8))
        TmpPaySeqNo = Trim(Mid(TextLine, 3421, 2))
        TmpClientNo = Trim(Mid(TextLine, 3423, 8))
        TmpClientID = Trim(Mid(TextLine, 3431, 2))
        tmpRBAmt = Trim(Mid(TextLine, 3448, 8))
        TmpCUndExcess = Trim(Mid(TextLine, 3830, 12))
        TmpstaffDec = Trim(Mid(TextLine, 3050, 1))
        tmpGenWP = Mid(TextLine, 3051, 5)
        TmpPreExWP = Mid(TextLine, 3056, 5)
        TmpSpeIllWP = Mid(TextLine, 3061, 5)
        TmpPayorPremNet = IIf(IsNumeric(Trim(Mid(TextLine, 3684, 12))), Trim(Mid(TextLine, 3684, 12)), 0)
        TmpPayorPremGross = IIf(IsNumeric(Trim(Mid(TextLine, 3696, 12))), Format(Trim(Mid(TextLine, 3696, 12)), "#,##0.00"), 0)
        TmpPayorMcoNet = IIf(IsNumeric(Trim(Mid(TextLine, 3708, 12))), Trim(Mid(TextLine, 3708, 12)), 0)
        TmpPayorMCoGross = IIf(IsNumeric(Trim(Mid(TextLine, 3720, 12))), Trim(Mid(TextLine, 3720, 12)), 0)
        TmpIcBcPp2nd = Trim(Mid(TextLine, 3742, 20))
        LoadingPrem = IIf(IsNumeric(Trim(Mid(TextLine, 3764, 12))), Trim(Mid(TextLine, 3764, 12)), 0)
        FamDiscAmt = IIf(IsNumeric(Trim(Mid(TextLine, 3776, 12))), Trim(Mid(TextLine, 3776, 12)), 0)
        RenewDiscAmt = IIf(IsNumeric(Trim(Mid(TextLine, 3788, 12))), Trim(Mid(TextLine, 3788, 12)), 0)

    ElseIf Trim(BordxFmtType) = "E" Then
        txtTempMemberNo = Trim(Mid(TextLine, 11, 18))
        TmpSalutation = Trim(Mid(TextLine, 69, 5))
        TmpWorkNature = Mid(TextLine, 377, 1)
        TmpBlood = Trim(Mid(TextLine, 378, 3))
        TmpOccupation = Trim(Mid(TextLine, 347, 30))
        TmpAllergic = Trim(Mid(TextLine, 381, 20))
        TmpSmoking = Mid(TextLine, 2965, 1)
        TmpAlcohol = Mid(TextLine, 2966, 1)
        TmpSuppExclusion = Trim(Mid(TextLine, 463, 2500))
        TmpPrevPlnCode = Trim(Mid(TextLine, 3462, 15))
        TmpPrevInsCom = Trim(Mid(TextLine, 3485, 60))
        TmpPayRemarks = Trim(Mid(TextLine, 3545, 60))
        TmpPayNo = Trim(Mid(TextLine, 3442, 8))
        TmpPaySeqNo = Trim(Mid(TextLine, 3450, 2))
        TmpClientNo = Trim(Mid(TextLine, 3452, 8))
        TmpClientID = Trim(Mid(TextLine, 3460, 2))
        tmpRBAmt = Mid(TextLine, 3477, 8)
        TmpCUndExcess = Trim(Mid(TextLine, 3859, 12))
        TmpstaffDec = Trim(Mid(TextLine, 3079, 1))
        TmpPayorPremNet = IIf(IsNumeric(Trim(Mid(TextLine, 3713, 12))), Trim(Mid(TextLine, 3713, 12)), 0)
        TmpPayorPremGross = IIf(IsNumeric(Trim(Mid(TextLine, 3725, 12))), Trim(Mid(TextLine, 3725, 12)), 0)
        TmpPayorMcoNet = IIf(IsNumeric(Trim(Mid(TextLine, 3737, 12))), Trim(Mid(TextLine, 3737, 12)), 0)
        TmpPayorMCoGross = IIf(IsNumeric(Trim(Mid(TextLine, 3749, 12))), Trim(Mid(TextLine, 3749, 12)), 0)
        LoadingPrem = IIf(IsNumeric(Trim(Mid(TextLine, 3793, 12))), Trim(Mid(TextLine, 3793, 12)), 0)
        FamDiscAmt = IIf(IsNumeric(Trim(Mid(TextLine, 3805, 12))), Trim(Mid(TextLine, 3805, 12)), 0)
        RenewDiscAmt = IIf(IsNumeric(Trim(Mid(TextLine, 3817, 12))), Trim(Mid(TextLine, 3817, 12)), 0)
    End If
    
    tmpGenWP = CDbl(IIf(IsNumeric(tmpGenWP) = True, tmpGenWP, 0))
    TmpPreExWP = CDbl(IIf(IsNumeric(TmpPreExWP) = True, TmpPreExWP, 0))
    TmpSpeIllWP = CDbl(IIf(IsNumeric(TmpSpeIllWP) = True, TmpSpeIllWP, 0))

    TmpSuppExclusion = Replace(TmpSuppExclusion, "'", "''")
    TmpOccupation = Replace(TmpOccupation, "'", "''")
    TmpPrevInsCom = Replace(TmpPrevInsCom, "'", "''")
    TmpSalutation = Replace(TmpSalutation, "'", "''")
    
'    StrDateSql = ""
'    StrWDateSql = ""
'    If IsNumeric(TmpGenWP) Then
'        TmpGenEndWP = Format(DateAdd("d", TmpGenWP, PayEffDate), "yyyy/mm/dd")
'        StrDateSql = StrDateSql & ", MBMCGenEndWP "
'        StrWDateSql = StrWDateSql & ", '" & TmpGenEndWP & "'"
'    End If
'    If IsNumeric(TmpPreExWP) Then
'        TmpPreExEndWP = Format(DateAdd("d", TmpPreExWP, PayEffDate), "yyyy/mm/dd")
'        StrDateSql = StrDateSql & ", MBMCPreExEndWP "
'        StrWDateSql = StrWDateSql & ", '" & TmpPreExEndWP & "'"
'    End If
'    If IsNumeric(TmpSpeIllWP) Then
'        TmpSpeIllEndWP = Format(DateAdd("d", TmpSpeIllWP, PayEffDate), "yyyy/mm/dd")
'        StrDateSql = StrDateSql & ", MBMSpeIllEndWP "
'        StrWDateSql = StrWDateSql & ", '" & TmpSpeIllEndWP & "'"
'    End If
    
    tmpRBAmt = CDbl(IIf(IsNumeric(tmpRBAmt) = True, tmpRBAmt, 0))
    TmpCUndExcess = CDbl(IIf(IsNumeric(TmpCUndExcess) = True, TmpCUndExcess, 0))
    TmpCovID = Format(CovCnt, "00")
'    Supp2WSql = "Insert Into MBMCoveredPersonsTwo(MBMCNumber, MBMCCoverID, MBMCSalutation, " & _
'                "MBMCOccupation, MBMCWorkNature, MBMCBloodType, MBMCAllergic, MBMCSmoking,  " & _
'                "MBMCAlcohol, MBMCPrevPlnCode, MBMCPrevInsCom, MBMCPayRemarks, MBMCExclusion, " & _
'                "MBMCPayNo, MBMCPaySeqNo, MBMCClientNo, MBMCClientID, MBMCRBAmt, MBMCstaffDec, " & _
'                "MBMCPayorPremNet,  MBMCPayorMcoNet, MBMCPayorMCoGross,  " & _
'                "MBMCIcBcPp2nd, MBMCLoadingPrem, MBMCFamDiscAmt, MBMCRenewDiscAmt,MBMCPayorPremGross, " & _
'                "MBMCGenWP , MBMCPreExWP, MBMCSpeIllWP"
'
'    If Trim(StrDateSql) <> "" Then
'       Supp2WSql = Supp2WSql & StrDateSql
'    End If
'        Supp2WSql = Supp2WSql & ")"
'        Supp2WSql = Supp2WSql & "Values('" & Trim(txtTempMemberNo) & "','" & TmpCovID & "','" & TmpSalutation & "', " & _
'                "'" & TmpOccupation & " ','" & TmpWorkNature & "','" & TmpBlood & "','" & TmpAllergic & "','" & TmpSmoking & "', " & _
'                "'" & TmpAlcohol & "','" & TmpPrevPlnCode & "','" & TmpPrevInsCom & "','" & TmpPayRemarks & "','" & TmpSuppExclusion & "', " & _
'                "'" & TmpPayNo & "','" & TmpPaySeqNo & "','" & TmpClientNo & "','" & TmpClientID & "'," & TmpRBAmt & ",'" & TmpstaffDec & "', " & _
'                "" & TmpPayorPremNet & ", " & TmpPayorMcoNet & ", " & TmpPayorMCoGross & ", " & _
'                "'" & TmpIcBcPp2nd & "', " & LoadingPrem & ", " & FamDiscAmt & ", " & RenewDiscAmt & ", " & TmpPayorPremGross & " "
'    If Trim(StrWDateSql) <> "" Then
'        Supp2WSql = Supp2WSql & StrWDateSql
'    End If
'        Supp2WSql = Supp2WSql & ")"
'    medixmasterconn.Execute Supp2WSql
    Supp2WSql = "EXEC Upload_MBMCoveredPersonsTwo '" & Trim(txtTempMemberNo) & "','" & TmpCovID & "','" & TmpSalutation & "', " & _
            "'" & TmpOccupation & " ','" & TmpWorkNature & "','" & TmpBlood & "','" & TmpAllergic & "','" & TmpSmoking & "', " & _
            "'" & TmpAlcohol & "','" & TmpPrevPlnCode & "','" & TmpPrevInsCom & "','" & TmpPayRemarks & "','" & TmpSuppExclusion & "', " & _
            "'" & TmpPayNo & "','" & TmpPaySeqNo & "','" & TmpClientNo & "','" & TmpClientID & "'," & tmpRBAmt & "," & _
            "" & TmpPayorPremNet & ", " & TmpPayorMcoNet & ", " & TmpPayorMCoGross & ", " & _
            "'" & TmpIcBcPp2nd & "', " & LoadingPrem & ", " & FamDiscAmt & ", " & RenewDiscAmt & ", " & TmpPayorPremGross & "," & _
            "" & tmpGenWP & ", " & TmpPreExWP & ", " & TmpSpeIllWP & ",'" & TmpstaffDec & "'," & TmpCUndExcess
    medixmasterconn.Execute Supp2WSql
End Sub

Private Sub GetSuppAgeLmt(tmpLine As Integer, DateOfBirth, tmpProdType, tmpPOlNo As String, tmpName As String, tmpDataType As String, tmpPlanCode, tmpInsType As String, TmpEffDate, TmpExpDate As String)
Dim CheckYr As Long
Dim PLNRec As New ADODB.Recordset
Dim PLNSql As String
Dim OverAge As Boolean
Dim PlnAgeLmt As String
Dim SuppDobEffDate As String

    SuppDobEffDate = ""
    PlnAgeLmt = ""
    OverAge = False
    SuppDobEffDate = Format(DateOfBirth, "dd/mm") & "/" & Year(TmpEffDate)
    CheckYr = Year(TmpEffDate) - Year(DateOfBirth)
    medixmasterconn.Open "File Name=" & BOSPath & "Conn\mediconnect.UDL"

    PLNSql = "select PlnAgeLimit from PLN where PLNCode = '" & Trim(tmpPlanCode) & "'"
    PLNRec.Open PLNSql, medixmasterconn, adOpenKeyset, adLockOptimistic
    If PLNRec.recordCount <= 0 Then
        Call ErrorListing(tmpLine, tmpPOlNo, tmpName, "Invalid " & tmpProdType & " Plan Code")
    Else
        If Trim(PLNRec!PlnAgeLimit) <> "" Then
             PlnAgeLmt = IIf(IsNumeric(PLNRec!PlnAgeLimit), PLNRec!PlnAgeLimit, 0)
        End If
    End If
    PLNRec.Close
    Set PLNRec = Nothing
    medixmasterconn.Close
    Set medixmasterconn = Nothing
    
    If Trim(PlnAgeLmt) <> "" Then
        If CheckYr > PlnAgeLmt Then
            OverAge = True
        ElseIf CheckYr = PlnAgeLmt Then
            If CDate(SuppDobEffDate) <= CDate(TmpEffDate) Then
                OverAge = True
            End If
        End If
    End If
   
    If OverAge = True Then
        Call ErrorListing(tmpLine, tmpPOlNo, tmpName, "Supplementary age is above " & PlnAgeLmt)
    End If
End Sub

Private Function GetSuppExpDate(DateOfBirth, tmpProdType, tmpPOlNo, tmpName, tmpPlanCode, TmpEffDate, TmpExpDate, tmpPayor As String)
Dim CheckYr As Long
Dim PLNRec As New ADODB.Recordset
Dim PLNSql As String
Dim PlnAgeLmt As String
Dim SuppDobExpDate As String

    PlnAgeLmt = ""
    SuppDobExpDate = Format(DateOfBirth, "dd/mm") & "/" & Year(TmpExpDate)
    CheckYr = Year(TmpExpDate) - Year(DateOfBirth)
    If Mid(cboHLTCode1, 1, 1) = "S" Then
        PLNSql = "select PlnAgeLimit from PLN where PLNCode = '" & Trim(tmpPlanCode) & _
                "' AND HLTCode='" & Mid(cboHLTCode1, 1, 1) & "'"
    Else
        PLNSql = "select PlnAgeLimit from PLN where PLNCode = '" & Trim(tmpPlanCode) & _
                "' AND PAYCode='" & Trim(tmpPayor) & "'"
    End If
    If Trim(tmpProdType) = "HIS" Then
'        medixmasterconn.Open "File Name=" & BOSPath & "Conn\mediconnect.UDL"
        PLNRec.Open PLNSql, medixmasterconn, adOpenForwardOnly, adLockReadOnly
    ElseIf Trim(tmpProdType) = "CIS" Then
'        MediConnCISDB.Open "File Name=" & BOSPath & "Conn\MediConnCISDB.UDL"
        PLNRec.Open PLNSql, MediConnCISDB, adOpenForwardOnly, adLockReadOnly
    End If
    If PLNRec.recordCount <= 0 Then
'        Call ErrorListing(NoOfERecLine, tmpPOlNo, tmpName, "Invalid " & tmpProdType & " Plan Code")
'        Call ErrorListing
    Else
        If Trim(PLNRec!PlnAgeLimit) <> "" Then
             PlnAgeLmt = IIf(IsNumeric(PLNRec!PlnAgeLimit), PLNRec!PlnAgeLimit, 0)
        End If
    End If
'    PLNRec.Close
'    Set PLNRec = Nothing
'    If Trim(tmpProdType) = "HIS" Then
'        medixmasterconn.Close
'        Set medixmasterconn = Nothing
'    ElseIf Trim(tmpProdType) = "CIS" Then
'        MediConnCISDB.Close
'        Set MediConnCISDB = Nothing
'    End If
    GetSuppExpDate = TmpExpDate
    If Trim(PlnAgeLmt) <> "" Then
        If CheckYr = PlnAgeLmt Then
            If CDate(SuppDobExpDate) <= CDate(TmpExpDate) Then
                GetSuppExpDate = CDate(SuppDobExpDate)
            End If
        End If
    End If

End Function

Private Sub SaveMBMXRef(XNumber, XCovID, XPolNo, XIC, XName, XEffdate, XExpDate, XAgeCode, XInsCode, XPlan, XBordxDate, XTopupNo, XTopupIND, XDataStatus, XStatus, XDOb, XHltCode, XPayor, XBatchNo)
Dim XrefWSql As String
Dim XrefRec As New ADODB.Recordset

    XPolNo = Replace(XPolNo, "'", "''")
    XName = Replace(XName, "'", "''")

    If Trim(XPlan) <> "" Then
        XrefWSql = "EXEC Upload_MBMXref_HIS '" & XNumber & "', '" & Format(XCovID, "00") & "', '" & XPolNo & "', '" & XIC & "'," & _
                "'" & Trim(XName) & "', '" & XEffdate & "','" & XExpDate & "', '" & XAgeCode & "','" & XInsCode & "','" & XPlan & "'," & _
                "'" & XBordxDate & "', '" & Trim(XTopupNo) & "', '" & Trim(XTopupIND) & "'," & _
                "'" & XDataStatus & "','" & XStatus & "', '" & XDOb & "','" & XHltCode & "', '" & XPayor & "', '" & Trim(XBatchNo) & "'"
    Else
        XrefWSql = "EXEC Upload_MBMXref_CIS '" & XNumber & "', '" & Format(XCovID, "00") & "', '" & XIC & "'," & _
                "'" & Trim(XName) & "', '" & XAgeCode & "'," & _
                "'" & XBordxDate & "'," & _
                "'" & XDataStatus & "','" & XStatus & "', '" & XDOb & "','" & XHltCode & "', '" & XPayor & "', '" & Trim(XBatchNo) & "'"
    End If
    medixmasterconn.Execute XrefWSql

End Sub

Private Sub txtPAYCode_LostFocus()
Dim strsql As String
Dim PayRec As New ADODB.Recordset
Dim cmd As New ADODB.Command
Dim prm1 As New ADODB.Parameter
Dim prm2 As New ADODB.Parameter
Dim tmpTableName As String

    txtClnStatus = "N"
    MBMCancelStatus = "N"
    strsql = " AND PAYCode ='" & Mid(Trim(txtPAYCode), 1, 2) & "'"
    tmpTableName = "PAY"
    
    medixmasterconn.Open "File Name=" & BOSPath & "Conn\mediconnect.UDL"

    Set cmd.ActiveConnection = medixmasterconn
    cmd.CommandText = "SearchHISDBComboListing"
    cmd.CommandType = adCmdStoredProc
    cmd.CommandTimeout = 300
    Set prm1 = cmd.CreateParameter("strAppend", adVarChar, adParamInput, 2000, strsql)
    cmd.Parameters.Append prm1
    Set prm2 = cmd.CreateParameter("TableName", adVarChar, adParamInput, 255, tmpTableName)
    cmd.Parameters.Append prm2
    PayRec.CursorLocation = adUseClient
    PayRec.CursorType = adOpenForwardOnly
    PayRec.CacheSize = 100
    Set PayRec = cmd.Execute
    
    Do Until PayRec.EOF
        With PayRec
           txtClnStatus = IIf(Len(Trim(PayRec!CLNStatus)), Trim(PayRec!CLNStatus), "N")
           MBMCancelStatus = IIf(Len(Trim(PayRec!MBMCancelStatus)), Trim(PayRec!MBMCancelStatus), "N")
           PayRec.MoveNext
        End With
    Loop
    
    PayRec.Close
    Set PayRec = Nothing
    medixmasterconn.Close
    Set medixmasterconn = Nothing
End Sub

Private Sub VerifyMBMByPassingPA(tmpLine As Integer, tmpProdType As String, tmpPOlNo As String, tmpName As String, tmpDataType As String, tmpPlanCode As String, tmpInsType As String, TmpEffDate As String, TmpExpDate As String)
Dim CoverDays As String
Dim tmpTableName As String
Dim TmpRecFound As Boolean
Dim SelFields As String
Dim SelCriteria As String
Dim tmpCode As String
Dim tmpDesc As String
    tmpTableName = "INS"
    TmpRecFound = False
    SelFields = "SELECT INSCode as tmpCode, INSDescription as tmpDesc "
    SelCriteria = " And INSCode ='" & Trim(tmpInsType) & "'"
    tmpTableName = "INS"
    tmpCode = ""
    tmpDesc = ""
    TmpRecFound = False

    Call SPHISTableRecVerify(SelFields, tmpTableName, SelCriteria, TmpRecFound, tmpCode, tmpDesc)

    If TmpRecFound = False Then
        Call ErrorListing(tmpLine, tmpPOlNo, tmpName, "No Such " & tmpProdType & " Insured Type - '" & Trim(tmpInsType) & "' !")
    End If
    If IsDate(Trim(TmpEffDate)) = False Then
        Call ErrorListing(tmpLine, tmpPOlNo, tmpName, "Invalid " & tmpProdType & " Date format ~ Effective Date")
    End If
    If IsDate(Trim(TmpExpDate)) = False Then
        Call ErrorListing(tmpLine, tmpPOlNo, tmpName, "Invalid " & tmpProdType & " Date format ~ Expiry Date ")
    End If
    If IsDate(Trim(TmpEffDate)) And IsDate(Trim(TmpExpDate)) Then
        CoverDays = CDate(TmpExpDate) - CDate(TmpEffDate)
        If CoverDays < 30 Or CoverDays > 365 Then
            Call ErrorListing(tmpLine, tmpPOlNo, tmpName, tmpProdType & " Policy Coverage must be > 30 and < 365 days ")
        End If
    End If
End Sub

Private Sub SaveMBMLifeTimeLmt(LMBMNumber, LMBMCovID, LmtType As String)
Dim TmpSql As String
Dim TmpLifeSql As String
Dim TmpLimit As String
Dim TmpLmtVar As String
Dim cmd As New ADODB.Command
Dim AccumAmtRec As New ADODB.Recordset
Dim prm1 As New ADODB.Parameter
Dim prm2 As New ADODB.Parameter
Dim TmpSql2 As Integer
Dim tmpAccAmt As Double
    
    tmpAccAmt = 0
    
    TmpSql2 = 0
    If Trim(LMBMCovID) <> "00" Then
        TmpSql2 = 1
    End If
    TmpSql = " And MBMTopupNumber ='" & Trim(LMBMNumber) & "' AND MBMPatientCovID='" & Trim(LMBMCovID) & "'"
    Set cmd.ActiveConnection = medixmasterconn
    cmd.CommandText = "SearchAccAppAmt"
    cmd.CommandType = adCmdStoredProc
    cmd.CommandTimeout = 15
    Set prm1 = cmd.CreateParameter("strAppend", adVarChar, adParamInput, 255, TmpSql)
    cmd.Parameters.Append prm1
    Set prm2 = cmd.CreateParameter("strAppendCase", adVarChar, adParamInput, 255, TmpSql2)
    cmd.Parameters.Append prm2
    AccumAmtRec.CursorLocation = adUseClient
    AccumAmtRec.CursorType = adOpenForwardOnly
    AccumAmtRec.CacheSize = 100
    Set AccumAmtRec = cmd.Execute

    Do While Not AccumAmtRec.EOF
        With AccumAmtRec
            tmpAccAmt = IIf(IsNumeric(!CLMTotalApproved), !CLMTotalApproved, 0)
            AccumAmtRec.MoveNext
        End With
    Loop
    AccumAmtRec.Close
    Set AccumAmtRec = Nothing
    
    TmpLmtVar = "SuppLifeTimeLimit"
    If LmtType = "A" Then
        TmpLmtVar = "LifeTimeLimit"
    End If
    TmpLimit = GetLifeTimeLimit(TmpLmtVar, InsuredCode, plancode, Mid(cboHLTCode1, 1, 1), CDate(PayEffDate))
    TmpLimit = CDbl(TmpLimit) - tmpAccAmt
    TmpLifeSql = "Insert Into MBMLifeTime(MBMNumber, MBMCOVID, MBMLifeTimeLmt) " & _
          "Values('" & LMBMNumber & "','" & LMBMCovID & "'," & TmpLimit & ")"
    medixmasterconn.Execute TmpLifeSql
End Sub

Private Sub SaveMBMBnfAnnLmt(LMBMNumber, LMBMCovID, LmtType, tmpKidneyAnnLmt, tmpCancerAnnLmt, tmpHomeNursAnnLmt)
Dim TmpBnfLmtSql As String
    
    TmpBnfLmtSql = "Insert Into MBMBnfLmt(MBMNumber, MBMCoveredID, MBMKidneyAnnLmt, MBMKidneyBalLmt, " & _
                 "MBMCancerAnnLmt, MBMCancerBalLmt, MBMHomeNCAnnLmt, MBMHomeNCBalLmt) " & _
                 "Values('" & LMBMNumber & "','" & LMBMCovID & "'," & tmpKidneyAnnLmt & "," & tmpKidneyAnnLmt & _
                 "," & tmpCancerAnnLmt & "," & tmpCancerAnnLmt & "," & tmpHomeNursAnnLmt & "," & tmpHomeNursAnnLmt & ")"
    medixmasterconn.Execute TmpBnfLmtSql

End Sub


Private Sub ComboKeyPress(tmpComboBox As ComboBox, KeyAscii As Integer)
Dim NewText As String
Dim ValidCount As Integer
Dim ValidValue As String
Dim i As Integer
    
    If KeyAscii >= 32 And KeyAscii <> 127 Then
    NewText = LCase(Left(tmpComboBox.Text, tmpComboBox.SelStart) + Chr(KeyAscii) + Mid(tmpComboBox.Text, tmpComboBox.SelStart + tmpComboBox.SelLength + 1))
    ValidCount = 0
    ValidValue = ""
    For i = 0 To tmpComboBox.ListCount - 1
        If NewText = LCase(Left(tmpComboBox.List(i), Len(NewText))) Then
            ValidCount = ValidCount + 1
            ValidValue = tmpComboBox.List(i)
        End If
    Next
    If ValidCount <= 1 Then KeyAscii = 0
         If ValidCount = 1 Then
           tmpComboBox.Text = ValidValue
           tmpComboBox.SelStart = 0
           tmpComboBox.SelLength = Len(ValidValue)
         End If
    End If
End Sub

Private Sub LoadComboListing(strsql As String, tmpTableName As String, tmpComboBox As ComboBox)
Dim TmpREc As New ADODB.Recordset
Dim cmd As New ADODB.Command
Dim prm1 As New ADODB.Parameter
Dim prm2 As New ADODB.Parameter

    tmpComboBox.Clear
    medixmasterconn.Open "File Name=" & BOSPath & "Conn\mediconnect.UDL"

    Set cmd.ActiveConnection = medixmasterconn
    cmd.CommandText = "SearchHISDBComboListing"
    cmd.CommandType = adCmdStoredProc
    cmd.CommandTimeout = 300
    Set prm1 = cmd.CreateParameter("strAppend", adVarChar, adParamInput, 2000, strsql)
    cmd.Parameters.Append prm1
    Set prm2 = cmd.CreateParameter("TableName", adVarChar, adParamInput, 255, tmpTableName)
    cmd.Parameters.Append prm2
    TmpREc.CursorLocation = adUseClient
    TmpREc.CursorType = adOpenForwardOnly
    TmpREc.CacheSize = 100
    Set TmpREc = cmd.Execute
    
    Do Until TmpREc.EOF
        With TmpREc
            If Trim(tmpTableName) = "HLT" Then
                tmpComboBox.AddItem !HLTCode & " - " & !HLTDescription
            ElseIf Trim(tmpTableName) = "PAY" Then
                tmpComboBox.AddItem !PAYCode & " - " & !PAYCompanyName
            End If
            TmpREc.MoveNext
        End With
    Loop
    
    TmpREc.Close
    Set TmpREc = Nothing
    medixmasterconn.Close
    Set medixmasterconn = Nothing
End Sub

Private Sub SPHISTableRecVerify(tmpSelFields As String, tmpTableName As String, tmpSelCriteria As String, TmpRecFound As Boolean, tmpCode As String, tmpDesc As String)
Dim TmpREc As New ADODB.Recordset
Dim cmd As New ADODB.Command
Dim prm1 As New ADODB.Parameter
Dim prm2 As New ADODB.Parameter
Dim Prm3 As New ADODB.Parameter

    TmpRecFound = False
    medixmasterconn.Open "File Name=" & BOSPath & "Conn\mediconnect.UDL"

    Set cmd.ActiveConnection = medixmasterconn
    cmd.CommandText = "SearchHISTableRec"
    cmd.CommandType = adCmdStoredProc
    cmd.CommandTimeout = 300
    Set prm1 = cmd.CreateParameter("SelFields", adVarChar, adParamInput, 2000, tmpSelFields)
    cmd.Parameters.Append prm1
    Set prm2 = cmd.CreateParameter("TableName", adVarChar, adParamInput, 255, tmpTableName)
    cmd.Parameters.Append prm2
    Set Prm3 = cmd.CreateParameter("SelCriteria", adVarChar, adParamInput, 2000, tmpSelCriteria)
    cmd.Parameters.Append Prm3

    TmpREc.CursorLocation = adUseClient
    TmpREc.CursorType = adOpenForwardOnly
    TmpREc.CacheSize = 100
    Set TmpREc = cmd.Execute
    
    Do Until TmpREc.EOF
        With TmpREc
            tmpCode = IIf(Len(TmpREc!tmpCode), TmpREc!tmpCode, "")
            tmpDesc = IIf(Len(TmpREc!tmpDesc), TmpREc!tmpDesc, "")
            TmpRecFound = True
            TmpREc.MoveNext
            Exit Do
        End With
    Loop
    
    TmpREc.Close
    Set TmpREc = Nothing
    medixmasterconn.Close
    Set medixmasterconn = Nothing
End Sub

Private Sub SPMaintTableRecVerify(tmpSelFields As String, tmpTableName As String, tmpSelCriteria As String, TmpRecFound As Boolean, tmpCode As String, tmpDesc As String)
Dim TmpREc1 As New ADODB.Recordset
Dim cmd As New ADODB.Command
Dim prm1 As New ADODB.Parameter
Dim prm2 As New ADODB.Parameter
Dim Prm3 As New ADODB.Parameter

    TmpRecFound = False
    medixmasterconnMaint.Open "File Name=" & BOSPath & "Conn\MediConnHISMaint.UDL"

    Set cmd.ActiveConnection = medixmasterconnMaint
    cmd.CommandText = "SearchMaintTableRec"
    cmd.CommandType = adCmdStoredProc
    cmd.CommandTimeout = 300
    Set prm1 = cmd.CreateParameter("SelFields", adVarChar, adParamInput, 2000, tmpSelFields)
    cmd.Parameters.Append prm1
    Set prm2 = cmd.CreateParameter("TableName", adVarChar, adParamInput, 255, tmpTableName)
    cmd.Parameters.Append prm2
    Set Prm3 = cmd.CreateParameter("SelCriteria", adVarChar, adParamInput, 255, tmpSelCriteria)
    cmd.Parameters.Append Prm3
    TmpREc1.CursorLocation = adUseClient
    TmpREc1.CursorType = adOpenForwardOnly
    TmpREc1.CacheSize = 100
    Set TmpREc1 = cmd.Execute
    
    Do Until TmpREc1.EOF
        With TmpREc1
            tmpCode = IIf(Len(TmpREc1!tmpCode), TmpREc1!tmpCode, "")
            tmpDesc = IIf(Len(TmpREc1!tmpDesc), TmpREc1!tmpDesc, "")
            TmpRecFound = True
            TmpREc1.MoveNext
            Exit Do
        End With
    Loop
    
    TmpREc1.Close
    Set TmpREc1 = Nothing
    medixmasterconnMaint.Close
    Set medixmasterconnMaint = Nothing
End Sub

Private Sub SPCISDBTableRecVerify(tmpSelFields As String, tmpTableName As String, tmpSelCriteria As String, TmpRecFound As Boolean, tmpCode As String, tmpDesc As String)
Dim TmpREc1 As New ADODB.Recordset
Dim cmd As New ADODB.Command
Dim prm1 As New ADODB.Parameter
Dim prm2 As New ADODB.Parameter
Dim Prm3 As New ADODB.Parameter

    TmpRecFound = False
    MediConnCISDB.Open "File Name=" & BOSPath & "Conn\MediConnCISDB.UDL"

    Set cmd.ActiveConnection = MediConnCISDB
    cmd.CommandText = "SearchCISDBTableRec"
    cmd.CommandType = adCmdStoredProc
    cmd.CommandTimeout = 300
    Set prm1 = cmd.CreateParameter("SelFields", adVarChar, adParamInput, 2000, tmpSelFields)
    cmd.Parameters.Append prm1
    Set prm2 = cmd.CreateParameter("TableName", adVarChar, adParamInput, 255, tmpTableName)
    cmd.Parameters.Append prm2
    Set Prm3 = cmd.CreateParameter("SelCriteria", adVarChar, adParamInput, 255, tmpSelCriteria)
    cmd.Parameters.Append Prm3
    TmpREc1.CursorLocation = adUseClient
    TmpREc1.CursorType = adOpenForwardOnly
    TmpREc1.CacheSize = 100
    Set TmpREc1 = cmd.Execute
    
    Do Until TmpREc1.EOF
        With TmpREc1
            tmpCode = IIf(Len(TmpREc1!tmpCode), TmpREc1!tmpCode, "")
            tmpDesc = IIf(Len(TmpREc1!tmpDesc), TmpREc1!tmpDesc, "")
            TmpRecFound = True
            TmpREc1.MoveNext
            Exit Do
        End With
    Loop
    
    TmpREc1.Close
    Set TmpREc1 = Nothing
    MediConnCISDB.Close
    Set MediConnCISDB = Nothing
End Sub

Private Sub SPMBMXref(NoOfRecLine As Integer, VNewPolicyNo As String, vNewName As String, VNewInsured As String, XrefSql1 As String, XrefSql2 As String, tmpVSql As String)
Dim XrefDB As New ADODB.Recordset
Dim cmd As New ADODB.Command
Dim prm1 As New ADODB.Parameter
Dim prm2 As New ADODB.Parameter
    
    medixmasterconnMaint.Open "File Name=" & BOSPath & "Conn\MediConnHISMaint.UDL"
    Set cmd.ActiveConnection = medixmasterconnMaint
    cmd.CommandText = "SearchMBMCrossRef"
    cmd.CommandType = adCmdStoredProc
    cmd.CommandTimeout = 300
    Set prm1 = cmd.CreateParameter("SelField", adVarChar, adParamInput, 2000, XrefSql1)
    cmd.Parameters.Append prm1
    Set prm2 = cmd.CreateParameter("SelCriteria", adVarChar, adParamInput, 255, XrefSql2)
    cmd.Parameters.Append prm2
    XrefDB.CursorLocation = adUseClient
    XrefDB.CursorType = adOpenForwardOnly
    XrefDB.CacheSize = 100
    Set XrefDB = cmd.Execute
    Do Until XrefDB.EOF
        If tmpVSql = "1" Then
            Call ErrorListing(NoOfRecLine, VNewPolicyNo, vNewName, "Same IC " & Trim(XrefDB!MBMICBcPp) & " Existed in Membership No. :" & XrefDB!MBMNumber)
        ElseIf tmpVSql = "2" Then
            If Trim(XrefDB!MBMTopupNumber) <> "" Or IsNull(XrefDB!MBMTopupNumber) = False Then
                Call ErrorListing(NoOfRecLine, VNewPolicyNo, vNewName, "NO Record found for Topup ! ")
            End If
            If Trim(XrefDB!INSCode) <> Trim(VNewInsured) Then
                Call ErrorListing(NoOfRecLine, VNewPolicyNo, vNewName, "Topup Member Insured Type Is not SAME! ")
            End If
        End If
        XrefDB.MoveNext
    Loop
    XrefDB.Close
    Set XrefDB = Nothing
    
    medixmasterconnMaint.Close
    Set medixmasterconnMaint = Nothing
End Sub



