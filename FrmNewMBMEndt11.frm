VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form FrmNewMBMEndt11 
   Caption         =   "Member Uploading Form"
   ClientHeight    =   7830
   ClientLeft      =   60
   ClientTop       =   510
   ClientWidth     =   11400
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   7830
   ScaleWidth      =   11400
   WindowState     =   2  'Maximized
   Begin VB.Frame Frame4 
      Height          =   615
      Left            =   120
      TabIndex        =   73
      Top             =   7200
      Width           =   11175
      Begin VB.TextBox txtBordxFormat 
         Height          =   285
         Left            =   240
         TabIndex        =   83
         Top             =   240
         Visible         =   0   'False
         Width           =   615
      End
      Begin VB.TextBox txtPayorMBMNoType 
         Height          =   285
         Left            =   6000
         TabIndex        =   82
         Text            =   "Text1"
         Top             =   360
         Visible         =   0   'False
         Width           =   735
      End
      Begin VB.TextBox txtMBMPremium 
         Height          =   285
         Left            =   3240
         TabIndex        =   81
         Top             =   240
         Visible         =   0   'False
         Width           =   735
      End
      Begin VB.TextBox txtMBMMCOFees 
         Height          =   285
         Left            =   3960
         TabIndex        =   80
         Top             =   240
         Visible         =   0   'False
         Width           =   735
      End
      Begin VB.TextBox txtAccessFees 
         Height          =   285
         Left            =   4680
         TabIndex        =   79
         Top             =   240
         Visible         =   0   'False
         Width           =   735
      End
      Begin VB.TextBox txtMBMAdultChild 
         Height          =   285
         Left            =   5400
         TabIndex        =   78
         Top             =   240
         Visible         =   0   'False
         Width           =   375
      End
      Begin VB.TextBox txtTempMemberNo 
         Height          =   285
         Left            =   960
         TabIndex        =   77
         Text            =   "txtTempMemberNo"
         Top             =   240
         Visible         =   0   'False
         Width           =   375
      End
      Begin VB.CommandButton cmdStop 
         Caption         =   "Stop"
         Height          =   315
         Left            =   7440
         TabIndex        =   76
         Top             =   200
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
         TabIndex        =   75
         Top             =   200
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
         TabIndex        =   74
         Top             =   200
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
      TabIndex        =   0
      Top             =   240
      Width           =   11175
      _ExtentX        =   19711
      _ExtentY        =   12303
      _Version        =   393216
      Tabs            =   2
      TabsPerRow      =   2
      TabHeight       =   520
      TabCaption(0)   =   "Upload Bordx"
      TabPicture(0)   =   "FrmNewMBMEndt11.frx":0000
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "lstERRList"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Frame2"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "Frame1"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).ControlCount=   3
      TabCaption(1)   =   "Generate ASCII File"
      TabPicture(1)   =   "FrmNewMBMEndt11.frx":001C
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Frame8"
      Tab(1).ControlCount=   1
      Begin VB.Frame Frame1 
         Height          =   1815
         Left            =   120
         TabIndex        =   47
         Top             =   360
         Width           =   10935
         Begin VB.Frame Frame13 
            Height          =   615
            Left            =   5520
            TabIndex        =   59
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
               TabIndex        =   61
               Top             =   200
               Width           =   960
            End
            Begin VB.CommandButton cmdUpload 
               Caption         =   "&Upload"
               Enabled         =   0   'False
               Height          =   315
               Left            =   1080
               TabIndex        =   60
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
            TabIndex        =   58
            Top             =   1395
            Width           =   320
         End
         Begin VB.TextBox txtClnStatus 
            Height          =   285
            Left            =   9600
            TabIndex        =   57
            Text            =   "N"
            Top             =   480
            Visible         =   0   'False
            Width           =   615
         End
         Begin VB.TextBox txtMBMCancelStatus 
            Height          =   285
            Left            =   9600
            TabIndex        =   56
            Text            =   "N"
            Top             =   840
            Visible         =   0   'False
            Width           =   615
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
            TabIndex        =   55
            Tag             =   "Thanks"
            Text            =   "S - Sihat Malaysia"
            Top             =   180
            Width           =   3030
         End
         Begin VB.ComboBox cboPayorCode 
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
            TabIndex        =   54
            Top             =   480
            Width           =   3030
         End
         Begin VB.ComboBox cboFileType 
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
            TabIndex        =   53
            Text            =   "N - New File"
            Top             =   780
            Width           =   3030
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
            TabIndex        =   52
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
            TabIndex        =   51
            Top             =   1365
            Width           =   2625
         End
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
            Left            =   5880
            TabIndex        =   50
            Text            =   "Y - Yes"
            Top             =   240
            Width           =   1695
         End
         Begin VB.ComboBox CboBack2Back 
            Height          =   315
            Left            =   5880
            Sorted          =   -1  'True
            TabIndex        =   49
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
            Left            =   5880
            Sorted          =   -1  'True
            Style           =   2  'Dropdown List
            TabIndex        =   48
            Top             =   830
            Width           =   1695
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
            TabIndex        =   71
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
            TabIndex        =   70
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
            TabIndex        =   69
            Top             =   1440
            Width           =   975
         End
         Begin VB.Label Label5 
            Caption         =   "New/ Endt File"
            Height          =   255
            Left            =   240
            TabIndex        =   68
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
            TabIndex        =   67
            Top             =   135
            Width           =   1770
         End
         Begin VB.Label Label13 
            Caption         =   "Verify I/C"
            Height          =   255
            Left            =   4800
            TabIndex        =   66
            Top             =   255
            Width           =   855
         End
         Begin VB.Label Label14 
            Caption         =   "Label Status"
            Height          =   255
            Left            =   4800
            TabIndex        =   65
            Top             =   840
            Width           =   975
         End
         Begin VB.Label Label15 
            Caption         =   "BTB Grantee"
            Height          =   255
            Left            =   4800
            TabIndex        =   64
            Top             =   570
            Width           =   1215
         End
         Begin VB.Label Label8 
            Caption         =   "Clinical Status"
            Height          =   375
            Index           =   0
            Left            =   8040
            TabIndex        =   63
            Top             =   480
            Visible         =   0   'False
            Width           =   1335
         End
         Begin VB.Label Label8 
            Caption         =   "MBM Cancel Status"
            Height          =   375
            Index           =   1
            Left            =   8040
            TabIndex        =   62
            Top             =   840
            Visible         =   0   'False
            Width           =   1575
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
         TabIndex        =   36
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
            TabIndex        =   46
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
            TabIndex        =   45
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
            TabIndex        =   44
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
            TabIndex        =   43
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
            TabIndex        =   42
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
            TabIndex        =   41
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
            TabIndex        =   40
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
            TabIndex        =   39
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
            TabIndex        =   38
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
            TabIndex        =   37
            Top             =   240
            Width           =   720
         End
      End
      Begin VB.Frame Frame8 
         Height          =   3975
         Left            =   -74760
         TabIndex        =   1
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
            TabIndex        =   31
            Top             =   3000
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
               TabIndex        =   35
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
               TabIndex        =   34
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
               TabIndex        =   33
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
               TabIndex        =   32
               Top             =   240
               Width           =   2040
            End
         End
         Begin VB.Frame Frame14 
            Height          =   2895
            Left            =   120
            TabIndex        =   2
            Top             =   120
            Width           =   10455
            Begin VB.Frame BordxFrame 
               Height          =   600
               Left            =   120
               TabIndex        =   26
               Top             =   1800
               Visible         =   0   'False
               Width           =   9015
               Begin MSMask.MaskEdBox txtBordxTo 
                  Height          =   285
                  Left            =   5730
                  TabIndex        =   27
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
                  Left            =   1440
                  TabIndex        =   28
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
                  TabIndex        =   30
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
                  TabIndex        =   29
                  Top             =   225
                  Width           =   1275
               End
            End
            Begin VB.Frame BatchFrame 
               Height          =   600
               Left            =   120
               TabIndex        =   21
               Top             =   1800
               Width           =   9015
               Begin VB.TextBox txtBatchNo1 
                  Height          =   285
                  Left            =   5640
                  TabIndex        =   22
                  Top             =   180
                  Width           =   2500
               End
               Begin MSMask.MaskEdBox txtReceiptDate 
                  Height          =   285
                  Left            =   1530
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
                  TabIndex        =   25
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
                  TabIndex        =   24
                  Top             =   180
                  Width           =   675
               End
            End
            Begin VB.Frame Frame16 
               Height          =   1695
               Left            =   120
               TabIndex        =   4
               Top             =   120
               Width           =   10215
               Begin VB.ComboBox cboPlan 
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
                  Left            =   8040
                  Sorted          =   -1  'True
                  TabIndex        =   9
                  Top             =   1200
                  Width           =   2025
               End
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
                  TabIndex        =   8
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
                  TabIndex        =   7
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
                  ItemData        =   "FrmNewMBMEndt11.frx":0038
                  Left            =   4800
                  List            =   "FrmNewMBMEndt11.frx":003A
                  Sorted          =   -1  'True
                  TabIndex        =   6
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
                  TabIndex        =   5
                  Top             =   1200
                  Width           =   2025
               End
               Begin VB.Label Label36 
                  Caption         =   "Plan Code"
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
                  Index           =   1
                  Left            =   7080
                  TabIndex        =   20
                  Top             =   1200
                  Width           =   885
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
                  TabIndex        =   19
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
                  TabIndex        =   18
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
                  TabIndex        =   14
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
                  TabIndex        =   13
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
                  TabIndex        =   12
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
                  Index           =   0
                  Left            =   120
                  TabIndex        =   11
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
                  TabIndex        =   10
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
               TabIndex        =   3
               Top             =   1920
               Width           =   960
            End
         End
      End
      Begin MSComctlLib.ListView lstERRList 
         Height          =   3975
         Left            =   120
         TabIndex        =   72
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
      TabIndex        =   84
      Top             =   0
      Width           =   11415
   End
End
Attribute VB_Name = "FrmNewMBMEndt11"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim TextLine
Dim WLifeStatus As Boolean
Dim tHLTCode As String
Dim tPayor As String
Dim tmpVPrinPlnCode As String '
Dim tmpVPrinInsCode As String '
Dim tmpVPrinEffDate As String '
Dim tmpVPrinExpDate As String '
Dim tmpVPrinGuaRenew As String
Dim tMBMNo As String
Dim tHncInsType As String
Dim tPLNTopupStatus As String
Dim tPLNPrmIND As String
Dim tPLNMcoIND As String
Dim tPLNAnnLmtIND As String
Dim tPLNAgeLimit As String
Dim tPLNLifeTimeStatus As String
Dim tAnnSuppLmtStatus As String
Dim tCPayor As String
Dim tCPlnCode As String
Dim tCInsType As String
Dim tmpCisVEffDate As String
Dim tmpCisVExpDate As String
Dim tCPLNPRMInd As String
Dim tCPLNMCOInd As String
Dim tCPLNAnnLmtInd As String
Dim tCPLNVsTLmtInd As String
Dim tCPLNAgeLimit As String
Dim tCPLNLifeTimeStatus As String
Dim tCPlnDCStatus As String
Dim tCPlnPDentStatus As String
Dim tCPlnSDentStatus As String
Dim tCPlnPSpecStatus As String
Dim tCPlnSSpecStatus As String
Dim tCPlnPMepStatus As String
Dim tCPlnSMepStatus As String
Dim tCRoundUpStatus As String
Dim tCMCOProcfeeStatus As String
Dim tCProRate As String
Dim tCPlnPDenatlLmt, tCPlnSDentalLmt As String
Dim tCPlnPSpecLmt, tCPlnSSpecLmt As String
Dim tCAnnSuppLmtStatus As String
Dim PrevLifeMBMNo As String
Dim TmpRoundUpStatus As String
Dim tmpProRate As String
Dim tmpEndtPrincEffDate As String
Dim tmpEndtPrincExpDate As String
Dim AscMBM As New ADODB.Recordset
Dim DataString As String
Dim BasicPrem As Double
Dim tmpVTopupNo As String
Dim tmpPLNPremSexStatus As String
Dim tmpPlnCode As String
Dim tmpPLNGuaRenewal As String
Dim tmpStatus As String
Dim PaySearch As String
Public intExit As Integer
Private tmpPAYFormat As String
Dim tmpPLNPayor As String
Dim tmpCancelstatus As Boolean
Dim CancelDate As String
Dim tmpFileStatus As String
Dim tmpReinstateDate As String
Dim tAdultChild As String
Dim TempPremium, tempMCOFee As Double
Dim tempOriMCO, tempOriPrem As Double
Dim RREffDate As String
Dim RRExpDate As String
Dim tmpPLNCancerLTInd As String
Dim tmpfilename As String
Dim RecordArray() As String


Private Sub cboASCIIFileVer_KeyPress(KeyAscii As Integer)
    Call ComboKeyPress(cboASCIIFileVer, KeyAscii)
End Sub

Private Sub CboHLTCode_KeyPress(KeyAscii As Integer)
    Call ComboKeyPress(cboHLTCode, KeyAscii)
End Sub

Private Sub CboHLTCode_LostFocus()
Dim SelCriteria As String
Dim SelFields As String
    SelFields = "SELECT PAYCode as tmpCode, PAYCompanyName as tmpName "
    SelCriteria = "And HLTCOde = '" & Mid(cboHLTCode, 1, 1) & "'"
    Call LoadComboListing(SelFields, "PAY", SelCriteria, cboPAYCode)
End Sub

Private Sub cboHLTCode1_LostFocus()
Dim SelCriteria As String
Dim SelFields As String
    SelFields = "SELECT PAYCode as tmpCode, PAYCompanyName as tmpName "
    SelCriteria = " And HLTCOde = '" & Mid(cboHLTCode1, 1, 1) & "'"
    Call LoadComboListing(SelFields, "PAY", SelCriteria, cboPayorCode)
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
        cboFileType = ""
    ElseIf Mid(cboExportOption, 1, 1) = "D" Then
        BatchFrame.Visible = False
        BordxFrame.Visible = True
        txtBordxFrom.SetFocus
        txtBordxFrom.TabIndex = 13
        txtBordxTo.TabIndex = 14
        cmdGenerate.TabIndex = 15
        cboFileType = ""
    End If
End Sub

Private Sub cboPlan_KeyPress(KeyAscii As Integer)
    Call ComboKeyPress(cboPlan, KeyAscii)
End Sub

Private Sub CmdGenerate_Click()
Dim StrSql As String
Dim TotalPrincipal As Integer
Dim TotalSupplementary As Integer
Dim RecordSaved
Dim MBMOthers As New ADODB.Recordset
Dim MBMCoveredPersons As New ADODB.Recordset
Dim ASCII As New ADODB.Recordset
Dim textField As String
Dim Header As String
Dim Trailer As String
Dim Exportsql As String
    
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
                          " And MBMBatchNo = " & Trim(txtBatchNo1)
            ElseIf Mid(cboExportOption, 1, 1) = "D" Then
                textField = txtBordxTo
                
                If cboPlan = "" Then
                
                    Exportsql = "HLTCode = '" & Trim(Mid(cboHLTCode, 1, 1)) & _
                                "' And PAYCode = '" & Trim(Mid(cboPAYCode, 1, 2)) & "'" & _
                                " And MBMStatus <> 'D' " & _
                                " And MBMDataStatus <> 'D' " & _
                                " And MBMBordxDate >=  '" & Format(Trim(txtBordxFrom), "mm/dd/yyyy") & "'" & _
                                " And MBMBordxDate <= '" & Format(Trim(txtBordxTo), "mm/dd/yyyy") & "'"
                Else
                    If cboPlan = "Medicalife200" Then
                        Exportsql = "HLTCode = '" & Trim(Mid(cboHLTCode, 1, 1)) & _
                                    "' And PAYCode = '" & Trim(Mid(cboPAYCode, 1, 2)) & "'" & _
                                    " And MBMStatus <> 'D' " & _
                                    " And MBMDataStatus <> 'D' " & _
                                    " And MBMBordxDate >=  '" & Format(Trim(txtBordxFrom), "mm/dd/yyyy") & "'" & _
                                    " And MBMBordxDate <= '" & Format(Trim(txtBordxTo), "mm/dd/yyyy") & "'" & _
                                    " And PLNCode >=  'ML11'" & _
                                    " And PLNCode <= 'ML14'"
                    ElseIf cboPlan = "Medicalife206" Then
                        Exportsql = "HLTCode = '" & Trim(Mid(cboHLTCode, 1, 1)) & _
                                    "' And PAYCode = '" & Trim(Mid(cboPAYCode, 1, 2)) & "'" & _
                                    " And MBMStatus <> 'D' " & _
                                    " And MBMDataStatus <> 'D' " & _
                                    " And MBMBordxDate >=  '" & Format(Trim(txtBordxFrom), "mm/dd/yyyy") & "'" & _
                                    " And MBMBordxDate <= '" & Format(Trim(txtBordxTo), "mm/dd/yyyy") & "'" & _
                                    " And PLNCode >=  'ML61'" & _
                                    " And PLNCode <= 'ML64'"
                    ElseIf cboPlan = "Medicalife207" Then
                        Exportsql = "HLTCode = '" & Trim(Mid(cboHLTCode, 1, 1)) & _
                                    "' And PAYCode = '" & Trim(Mid(cboPAYCode, 1, 2)) & "'" & _
                                    " And MBMStatus <> 'D' " & _
                                    " And MBMDataStatus <> 'D' " & _
                                    " And MBMBordxDate >=  '" & Format(Trim(txtBordxFrom), "mm/dd/yyyy") & "'" & _
                                    " And MBMBordxDate <= '" & Format(Trim(txtBordxTo), "mm/dd/yyyy") & "'" & _
                                    " And PLNCode >=  'ML71'" & _
                                    " And PLNCode <= 'ML74'"
                    ElseIf cboPlan = "Medicalife206 - Rider" Then
                        Exportsql = "HLTCode = '" & Trim(Mid(cboHLTCode, 1, 1)) & _
                                    "' And PAYCode = '" & Trim(Mid(cboPAYCode, 1, 2)) & "'" & _
                                    " And MBMStatus <> 'D' " & _
                                    " And MBMDataStatus <> 'D' " & _
                                    " And MBMBordxDate >=  '" & Format(Trim(txtBordxFrom), "mm/dd/yyyy") & "'" & _
                                    " And MBMBordxDate <= '" & Format(Trim(txtBordxTo), "mm/dd/yyyy") & "'" & _
                                    " And PLNCode >=  'MR61'" & _
                                    " And PLNCode <= 'MR64'"

                    ElseIf cboPlan = "Medicalife207 - Rider" Then
                        Exportsql = "HLTCode = '" & Trim(Mid(cboHLTCode, 1, 1)) & _
                                    "' And PAYCode = '" & Trim(Mid(cboPAYCode, 1, 2)) & "'" & _
                                    " And MBMStatus <> 'D' " & _
                                    " And MBMDataStatus <> 'D' " & _
                                    " And MBMBordxDate >=  '" & Format(Trim(txtBordxFrom), "mm/dd/yyyy") & "'" & _
                                    " And MBMBordxDate <= '" & Format(Trim(txtBordxTo), "mm/dd/yyyy") & "'" & _
                                    " And PLNCode >=  'MR71'" & _
                                    " And PLNCode <= 'MR74'"
                    End If
                End If
            End If
            StrSql = "SELECT MBMNumber, MBMMemberType, INSCode, MBMPolicyNo, MBMPayorEffDate, MBMIcBcPp, " & _
                    " MBMName, PLNCode From MBM Where " & Exportsql
            medixmasterconn.Open "File Name=" & BOSPath & "Conn\mediconnect.UDL"

            AscMBM.Open StrSql, medixmasterconn, adOpenKeyset, adLockOptimistic

            If AscMBM.recordCount = 0 Then
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
                    Header = Trim("F" & Mid(cboPAYCode, 1, 2) & Mid(txtReceiptDate, 1, 2) & Mid(txtReceiptDate, 4, 2) & Mid(txtReceiptDate, 7, 4) & l("", 174))
                ElseIf Mid(cboExportOption, 1, 1) = "D" Then
                    Header = Trim("F" & Mid(cboPAYCode, 1, 2) & Mid(txtBordxFrom, 1, 2) & Mid(txtBordxFrom, 4, 2) & Mid(txtBordxFrom, 7, 4) & l("", 174))
                End If
                Print #1, Trim(Header)
                Print #2, Trim(Header)
            
                Call PrincipalASCII
                Trailer = "T" & l(txtTotalPrincipals, 10) & l(txtTotalSupplementary1, 10) & l("", 164)
                Print #1, Trailer
                Print #2, Trailer
                cmdGenerate.Enabled = False
                Close #1
                Close #2
                MsgBox "File Generated Sucessfully...", vbOKOnly + vbInformation, DialogBoxTitle
            End If
            AscMBM.Close
            Set AscMBM = Nothing
            medixmasterconn.Close
            Set medixmasterconn = Nothing
            Screen.MousePointer = vbDefault
        End If
    End If
    
    txtTotalPrincipals = l(txtTotalPrincipals, 10)
    txtTotalSupplementary1 = l(txtTotalSupplementary1, 10)
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
Dim MNumber As String
Dim MBMPayEffDate As String
Dim MBMEmployeeNo As String
Dim MBMGroupCompany As String
Dim MBMPolSubNo1 As String
Dim Check As Boolean
Dim MBMPolicyNo As String

    Do Until AscMBM.EOF
        If Trim(AscMBM!MBMMemberType) = "Q" Then
            DataString = l("P", 1)
        Else
            DataString = l(AscMBM!MBMMemberType, 1)
        End If
        
        TotalPrincipal = TotalPrincipal + 1
        txtTotalPrincipals = TotalPrincipal
        MNumber = Trim(AscMBM!MBMNumber)
        If Len(AscMBM!MBMPayorEffDate) Then
            MBMPayEffDate = Format(AscMBM!MBMPayorEffDate, "DDMMYYYY")
        End If
        MBMPolicyNo = IIf(Len(AscMBM!MBMPolicyNo), AscMBM!MBMPolicyNo, "")
        tmpPlnCode = IIf(Len(AscMBM!PLNCode), AscMBM!PLNCode, "")
        strsplOthers = "SELECT MBMEmployeeNo, MBMGroupCompany, MBMPolSubNo1 From MBMOthers Where MBMNumber = '" & Trim(AscMBM!MBMNumber) & "'"
        MBMOthers.Open strsplOthers, medixmasterconn, adOpenKeyset, adLockOptimistic
        If MBMOthers.recordCount Then
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
                        " From MBMOthersTwo Where MBMNumber = '" & Trim(AscMBM!MBMNumber) & "'"
        MBMOther2.Open MBMOthers2Sql, medixmasterconn, adOpenKeyset, adLockOptimistic
        If Len(MBMOther2!MBMPOLIssuedDate) Then
            XPPolDate = Format(MBMOther2!MBMPOLIssuedDate, "DDMMYYYY")
        End If
        XPRiskNo = IIf(Len(Trim(MBMOther2!MBMRiskNo)), MBMOther2!MBMRiskNo, "")
        XPTransNo = IIf(Len(Trim(MBMOther2!MBMTransNo)), MBMOther2!MBMTransNo, "")
        XPPayNo = IIf(Len(Trim(MBMOther2!MBMPayNo)), MBMOther2!MBMPayNo, "")
        XPpaySeqNo = IIf(Len(Trim(MBMOther2!MBMPAYSeqNo)), MBMOther2!MBMPAYSeqNo, "")
        XPClientNo = IIf(Len(Trim(MBMOther2!MBMClientNo)), MBMOther2!MBMClientNo, "")
        XPClientID = IIf(Len(Trim(MBMOther2!MBMClientID)), MBMOther2!MBMClientID, "")
        
        MBMOther2.Close
        Set MBMOther2 = Nothing
        DoEvents
        
        DataString = DataString & l(MBMPolicyNo, 25)
        DataString = DataString & l(MBMEmployeeNo, 15)
        DataString = DataString & l(AscMBM!MBMName, 60)
        DataString = DataString & l(AscMBM!MBMIcBcPp, 20)
        DataString = DataString & l(AscMBM!MBMNumber, 18)
        
        DataString = DataString & l("00", 2)
        DataString = DataString & l(MBMPayEffDate, 8)
        DataString = DataString & l(MBMGroupCompany, 40)
        DataString = DataString & l(XPPolDate, 8)
        'DataString = DataString & l(XPRiskNo, 4)
        DataString = DataString & l(tmpPlnCode, 4)
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
        If Trim(AscMBM!INSCode) = "F" Or Trim(AscMBM!INSCode) = "H" Then
            Call SupplementaryASCII(MNumber, MBMPayEffDate, MBMPolSubNo1, MBMEmployeeNo, MBMPolicyNo)
        End If
        
        AscMBM.MoveNext
    Loop
End Sub

Private Sub SupplementaryASCII(MNumber As String, MBMPayEffDate As String, MBMPolSubNo1 As String, MBMEmployeeNo As String, MBMPolicyNo As String)
Dim StrSql As String
Dim MBMCoveredPersons As New ADODB.Recordset
Dim TotalSupplementary As Integer
Dim MBMCov2 As New ADODB.Recordset
Dim MBMCovSql As String
Dim XCPayNo As String
Dim XCpaySeqNo As String
Dim XCClientNo As String
Dim XCClientID As String

    StrSql = "SELECT MBMCRELCode, MBMCName,MBMCIcBcPpNo,MBMCNumber,MBMCCoverID From MBMCoveredPersons  Where MBMCNumber = '" & Trim(MNumber) & "'"
    MBMCoveredPersons.Open StrSql, medixmasterconn, adOpenKeyset, adLockOptimistic
    
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
        
        DataString = DataString & l(Trim(MBMPolicyNo), 25)
        DataString = DataString & l(Trim(MBMEmployeeNo), 15)
        DataString = DataString & l(MBMCoveredPersons!MBMCName, 60)
        DataString = DataString & l(MBMCoveredPersons!MBMCICBCPPNo, 20)
        DataString = DataString & l(MBMCoveredPersons!MBMCNumber, 18)
        DataString = DataString & l(Format(MBMCoveredPersons!MBMCCoverID, "00"), 2)
        DataString = DataString & l(MBMPayEffDate, 8)
        DataString = DataString & l("", 40)
        DataString = DataString & l("", 8)
        DataString = DataString & l(tmpPlnCode, 4)
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
  CommonDialog1.ShowOpen
  txtImportFilename = CommonDialog1.FileTitle
  Exit Sub
  
ErrHandler:
  Exit Sub
End Sub

Private Sub cmdPrint_Click()
    Screen.MousePointer = vbHourglass
        
    If lstERRList.ListItems.Count <> 0 Then
        Call ExportListViewtoExcel(lstERRList)
    Else
        MsgBox "Report cannot be printed without any record.", vbInformation
    End If
        
    Screen.MousePointer = vbDefault
End Sub

Private Sub cmdStop_Click()
    intExit = 1
End Sub

Private Sub cmdUpload_Click()
Dim RecordSaved
Dim PAY As New ADODB.Recordset
Dim StrSql As String
    PaySearch = ""
    tmpFileStatus = ""
    tmpCancelstatus = False
    Open txtImportFilename For Input As #1   ' Open file.
        Line Input #1, TextLine   ' Read line into variable.
        tmpPAYFormat = Mid(TextLine, 28, 3)
        tmpFileStatus = Mid(TextLine, 12, 1)
    Close #1
    
    If Trim(cboHLTCode1) = "" Then
        MsgBox "Please Select Health Type", vbOKOnly + vbCritical, DialogBoxTitle
        cboFileType.SetFocus
    ElseIf Trim(cboFileType) = "" Then
        MsgBox "Please Select New / Endt file", vbOKOnly + vbCritical, DialogBoxTitle
        cboFileType.SetFocus
    ElseIf Trim(cboPayorCode) = "" And (Trim(tmpPAYFormat) <> "MAA" And Trim(tmpPAYFormat) <> "TAK") Then
        MsgBox "Please Enter Payor Code", vbOKOnly + vbCritical, DialogBoxTitle
        cboPayorCode.SetFocus
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
            
            If Trim(tmpPAYFormat) <> "MAA" And Trim(tmpPAYFormat) <> "TAK" Then
                StrSql = "Select BordxFormatType from PAY where PayCode = '" & Mid(cboPayorCode, 1, 2) & "'"
                PAY.Open StrSql, medixmasterconn, adOpenKeyset, adLockOptimistic
                If PAY.recordCount Then
                    txtBordxFormat = IIf(Len(PAY!BordxFormatType), PAY!BordxFormatType, "")
                Else
                    txtBordxFormat = ""
                End If
                PAY.Close
                Set PAY = Nothing
            Else
                txtBordxFormat = "3"
                txtPayorMBMNoType = "P"
            End If
            
            If txtBordxFormat = "4" Then
                Call GenerateMemberAXA
            Else
                If tmpFileStatus = "N" Then
                    If Mid(cboFileType, 1, 1) = "N" Then
                        Call GenerateMember
                    ElseIf Mid(cboFileType, 1, 1) = "E" Then
                        Call GenerateENDtMember
                        cmdExit.Enabled = True
                    End If
                ElseIf tmpFileStatus = "R" Or tmpFileStatus = "I" Then
                    Call GenerateRRMember
                ElseIf tmpFileStatus = "C" Then
                    Call GenerateENDtMember
                Else
                    Call GenerateMember
                End If
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
Dim PAY As New ADODB.Recordset
Dim StrSql As String
    PaySearch = ""
    txtSubmissionDate = ""
    txtTotalPrincipal = 0
    txtTotalSupplementary = 0
    txtTotalError = 0
    cmdFileLocation.Enabled = False
    cmdUpload.Enabled = False
    
    Open txtImportFilename For Input As #1   ' Open file.
        Line Input #1, TextLine   ' Read line into variable.
        tmpPAYFormat = Mid(TextLine, 28, 3)
    Close #1
    
    If Mid(cboFileType, 1, 1) = "" Then
        MsgBox "Please Select Bordx Type", vbOKOnly + vbCritical, DialogBoxTitle
        cboFileType.SetFocus
    ElseIf cboPayorCode = "" And (tmpPAYFormat <> "MAA" And tmpPAYFormat <> "TAK") Then
        MsgBox "Please Enter Payor Code", vbOKOnly + vbCritical, DialogBoxTitle
        cboPayorCode.SetFocus
    ElseIf cboFileType = "" Then
        MsgBox "Please Enter File Type", vbOKOnly + vbCritical, DialogBoxTitle
        cboFileType.SetFocus
    ElseIf txtImportFilename = "" Then
        MsgBox "Please Enter Filename", vbOKOnly + vbCritical, DialogBoxTitle
    Else
        
        cmdVerify.Enabled = False
        cmdUpload.Enabled = False
        lstERRList.ListItems.Clear
        medixmasterconnMaint.Open "File Name=" & BOSPath & "Conn\MediConnHISMaint.UDL"
        'edit by shan 13012011
        If medixmasterconn = "" Then
        medixmasterconn.Open "File Name=" & BOSPath & "Conn\mediconnect.UDL"
        End If

        
        If Trim(tmpPAYFormat) <> "MAA" And Trim(tmpPAYFormat) <> "TAK" Then
            StrSql = "Select BordxFormatType from PAY where PayCode = '" & Mid(cboPayorCode, 1, 2) & "'"
            PAY.Open StrSql, medixmasterconn, adOpenKeyset, adLockOptimistic
                If PAY.recordCount Then
                    txtBordxFormat = IIf(Len(PAY!BordxFormatType), PAY!BordxFormatType, "")
                Else
                    txtBordxFormat = ""
                End If
                PAY.Close
                Set PAY = Nothing
        Else
            txtBordxFormat = "3"
        End If
        
        
        If txtBordxFormat = "4" Then
            Call VerifyNewMemberAXA
        Else
             If Mid(cboFileType, 1, 1) = "N" Then
                 'If txtBordxFormat = "2" Then
                 '    Call VerifyNewMemberTwo
                 'Else
            ' If tmpPAYFormat <> "MAA" Then
                 Call VerifyNewMember
            ' ElseIf tmpPAYFormat = "MAA" And Mid(TextLine, 1, 12) = "N" Then
            '     Call VerifyNewMember
            ' ElseIf tmpPAYFormat = "MAA" And (Mid(TextLine, 1, 12) = "R" Or Mid(TextLine, 1, 1) = "P") Then
            '     Call VerifyNewMember
            ' End If
                     
                 'End If
             ElseIf Mid(cboFileType, 1, 1) = "E" Then
                 Call VerifyENDtMember
             End If
        End If
        
        If intExit = 0 Then
            MsgBox "Verified Sucessfully...", vbOKOnly + vbInformation, DialogBoxTitle
            Screen.MousePointer = vbDefault
        Else
            Screen.MousePointer = vbDefault
        End If
        medixmasterconn.Close
        Set medixmasterconn = Nothing
        medixmasterconnMaint.Close
        Set medixmasterconnMaint = Nothing
    End If
    cmdVerify.Enabled = True
    cmdUpload.Enabled = True
    cmdFileLocation.Enabled = True
    intExit = 0
End Sub

Private Sub Form_Load()
Dim SelFields As String
Dim SelCriteria As String
    
    Call ComboListing
    SelFields = ""
    SelCriteria = ""
    SelFields = "SELECT HLTCode as tmpCode, HLTDescription as tmpName "
    Call LoadComboListing(SelFields, "HLT", SelCriteria, cboHLTCode)
    Call LoadComboListing(SelFields, "HLT", SelCriteria, cboHLTCode1)
    SelFields = "SELECT PAYCode as tmpCode, PAYCompanyName as tmpName "
    Call LoadComboListing(SelFields, "PAY", SelCriteria, cboPAYCode)
    Call LoadComboListing(SelFields, "PAY", SelCriteria, cboPayorCode)
    cboHLTCode1.ListIndex = 1
    intExit = 0
    txtPath = SRVPath & "ASCI"
    txtLocalFileLocation = "C:\"
End Sub

'---------------------Start ComboLisiting--------------------------
Private Sub ComboListing()
    cboFileType.AddItem "N" & " - " & "New File"
    cboFileType.AddItem "E" & " - " & "Endorsement File"
    
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

    cboPlan.AddItem "Medicalife200"
    cboPlan.AddItem "Medicalife206"
    cboPlan.AddItem "Medicalife207"
    cboPlan.AddItem "Medicalife206 - Rider"
    cboPlan.AddItem "Medicalife207 - Rider"
End Sub

Private Sub SSTab1_Click(PreviousTab As Integer)
    If SSTab1.Tab = 0 Then
        cmdStop.Visible = True
        cmdPrint.Visible = True
    Else
        cmdStop.Visible = False
        cmdPrint.Visible = False
        cmdExit.Visible = True
    End If
End Sub

'------------------------Start selection of related Info----------------
Private Sub cboFileType_KeyPress(KeyAscii As Integer)
    Call ComboKeyPress(cboFileType, KeyAscii)
End Sub

Private Sub cboFileType_LostFocus()
    If Mid(cboFileType, 1, 1) = "N" Then
        txtFileVersion = "Version 4800"
        CboBack2Back.Enabled = True
    ElseIf Mid(cboFileType, 1, 1) = "E" Then
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
Dim PrincPayEffdate As String
Dim PrincPayExpDate As String
Dim PrincPlanCode As String
Dim tmpPlanCode As String
Dim PrincInsType As String
Dim NoOfRecLine As Integer
Dim tmpTableName As String
Dim TmpRecFound As Boolean
Dim SelFields As String
Dim SelCriteria As String
Dim tmpCode As String
Dim tmpDesc As String
Dim VNewPayor As String
Dim CisprincEffDate As String
Dim CisPrincExpdate As String
Dim CisPrincPlnCode As String
Dim CisPrincInsType As String
Dim VAge As Integer
Dim vPrincipalEffDate As String
Dim tmpPayorPLN As String
    lstERRList.ListItems.Clear
    Erase MBMICArray
    ReDim MBMICArray(0)
    MBMICCnt = 0
       
    Screen.MousePointer = vbHourglass
    Open txtImportFilename For Input As #1   ' Open file.
    Do
    Do While Not EOF(1)   ' Loop until end of file.
       Line Input #1, TextLine   ' Read line into variable.
        NoOfRecLine = NoOfRecLine + 1
       'If tmpPAYFormat <> "MAA" Then
            If NoOfRecLine = 1 Then
                txtSubmissionDate = Format(Mid(TextLine, 4, 2) & "/" & Mid(TextLine, 6, 2) & "/" & Mid(TextLine, 8, 4), "dd/mm/yyyy")
                If Mid(TextLine, 1, 1) <> "F" Then
                    Call ErrorListing(NoOfRecLine, "", "", "1st Record must be header with Data Type 'F'")
                End If
                
                If Trim(tmpPAYFormat) <> "MAA" And Trim(tmpPAYFormat) <> "TAK" Then
                    If Trim(Mid(TextLine, 2, 2)) = "" Or Trim(Mid(TextLine, 2, 2)) <> Mid(cboPayorCode, 1, 2) Then
                        Call ErrorListing(NoOfRecLine, "", "", "Invalid Insurer's Code " & Trim(Mid(TextLine, 2, 2)))
                    End If
                End If
                
                If IsDate(txtSubmissionDate) = False Then
                    Call ErrorListing(NoOfRecLine, "", "", "Invalid Submission Date !")
                End If
                
                If Trim(tmpPAYFormat) <> "MAA" And Trim(tmpPAYFormat) <> "TAK" Then
                    VNewPayor = Mid(TextLine, 2, 2)
                Else
                    VNewPayor = ""
                End If
                
            End If
                             
            If NoOfRecLine >= 2 And Mid(TextLine, 1, 1) <> "T" Then
                VNewDataType = Trim(Mid(TextLine, 1, 1))
                If txtBordxFormat = "3" Then
                    VNewPolicyNo = Trim(Mid(TextLine, 2, 17))
                Else
                    VNewPolicyNo = Trim(Mid(TextLine, 2, 25))
                End If
                
                vNewName = Trim(Mid(TextLine, 47, 60))
                vNewName = Trim(Replace(vNewName, "'", "''"))
                VNewCity = Trim(Mid(TextLine, 197, 20))
                VNewPostCode = Trim(Mid(TextLine, 217, 5))
                VNewState = Trim(Mid(TextLine, 222, 15))
                VNewIC = Trim(Mid(TextLine, 273, 20))
                VNewDOB = Mid(TextLine, 293, 2) & "/" & Mid(TextLine, 295, 2) & "/" & Mid(TextLine, 297, 4)
                VNewSex = Trim(Mid(TextLine, 301, 1))
                VNewEffDate = Mid(TextLine, 374, 2) & "/" & Mid(TextLine, 376, 2) & "/" & Mid(TextLine, 378, 4)
                If Trim(VNewDataType) = "P" Or Trim(VNewDataType) = "R" Then
                    vPrincipalEffDate = Mid(TextLine, 374, 2) & "/" & Mid(TextLine, 376, 2) & "/" & Mid(TextLine, 378, 4)
                End If
                VNewExpDate = Mid(TextLine, 382, 2) & "/" & Mid(TextLine, 384, 2) & "/" & Mid(TextLine, 386, 4)
                
                If txtBordxFormat = "3" Then
                    PaySearch = 2
                    tmpPayorPLN = Trim(Mid(TextLine, 390, 7))
                    SelFields = "SELECT PLNCode as tmpCode, PLNLifeTimeStatus as tmpDesc, PLNPayorCode as tmpPLNPayor "
                    SelCriteria = " And PLNPayorPlan ='" & Trim(tmpPayorPLN) & "'"
                    tmpTableName = "PLN"
                    tmpCode = ""
                    tmpDesc = ""
                    TmpRecFound = False
                    Call SPHISTableRecVerify(SelFields, tmpTableName, SelCriteria, TmpRecFound, tmpCode, tmpDesc, tmpStatus, tmpPLNPayor)
                    If TmpRecFound = True Then
                        VNewPlanCode = tmpCode
                    Else
                        VNewPlanCode = ""
                    End If
                Else
                    VNewPlanCode = Trim(Mid(TextLine, 390, 4))
                End If
                
                If Trim(tmpPAYFormat) = "MAA" Or Trim(tmpPAYFormat) = "TAK" Then
                    VNewPayor = tmpPLNPayor
                End If
                
                If txtBordxFormat = "3" Then
                    VNewInsured = Trim(Mid(TextLine, 397, 1))
                Else
                    VNewInsured = Trim(Mid(TextLine, 394, 1))
                End If
                VRenewal = Trim(Mid(TextLine, 435, 1))
                If VNewEffDate <> "  /  /    " Then
                    VAge = Year(VNewEffDate) - Year(VNewDOB)
                Else
                    If vPrincipalEffDate <> "  /  /    " Then
                        VAge = Year(vPrincipalEffDate) - Year(VNewDOB)
                    End If
                End If
                If VNewDataType = "P" Or VNewDataType = "R" Then
                    PrincPayEffdate = Mid(TextLine, 374, 2) & "/" & Mid(TextLine, 376, 2) & "/" & Mid(TextLine, 378, 4)
                    PrincPayExpDate = Mid(TextLine, 382, 2) & "/" & Mid(TextLine, 384, 2) & "/" & Mid(TextLine, 386, 4)
                    If txtBordxFormat = "3" Then
                        PrincPlanCode = VNewPlanCode
                    Else
                        PrincPlanCode = Trim(Mid(TextLine, 390, 4))
                    End If
                    If txtBordxFormat = "3" Then
                        PrincInsType = Trim(Mid(TextLine, 397, 1))
                    Else
                        PrincInsType = Trim(Mid(TextLine, 394, 1))
                    End If
                    
                    'PrincInsType = Trim(Mid(TextLine, 397, 1))
                Else
                    VNewPlanCode = IIf(Trim(VNewPlanCode) = "", PrincPlanCode, VNewPlanCode)
                    VNewInsured = IIf(Trim(VNewInsured) = "", PrincInsType, VNewInsured)
                    If VNewPlanCode <> "" And PrincPayEffdate <> "" Then
                        VNewEffDate = IIf(IsDate(VNewEffDate) = True, VNewEffDate, CDate(PrincPayEffdate))
                        VNewExpDate = IIf(IsDate(VNewExpDate) = True, VNewExpDate, CDate(PrincPayExpDate))
                    End If
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
                If Trim(VNewPlanCode) <> "" Then
                    PaySearch = 2
                    SelFields = "SELECT PLNCode as tmpCode, PLNLifeTimeStatus as tmpDesc, PLNPayorCode as tmpPLNPayor "
                    SelCriteria = " And PLNCode ='" & Trim(VNewPlanCode) & "'"
                    tmpTableName = "PLN"
                    tmpCode = ""
                    tmpDesc = ""
                    TmpRecFound = False
                    Call SPHISTableRecVerify(SelFields, tmpTableName, SelCriteria, TmpRecFound, tmpCode, tmpDesc, tmpStatus, tmpPLNPayor)
                    If TmpRecFound = False Then
                        Call ErrorListing(NoOfRecLine, VNewPolicyNo, vNewName, "No such Plan Code - ' " & Trim(VNewPlanCode) & "' !")
                    Else
                        If Trim(tmpDesc) = "Y" Then
                            If Trim(Mid(TextLine, 3842, 5)) <> "1" And Trim(Mid(TextLine, 3842, 5)) <> "3" Then
                                Call ErrorListing(NoOfRecLine, VNewPolicyNo, vNewName, "Invalid Guarantee Renewal Code !")
                            End If
                        End If
                    End If
                    Call VerifyMBMByPassingPA(NoOfRecLine, "HIS", VNewPolicyNo, vNewName, VNewDataType, VNewPlanCode, VNewInsured, VNewEffDate, VNewExpDate)
                End If
                '-------- check for Ins Type, Eff Date, ExpDAte
                
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
                   
                    If Trim(tmpDesc) = "Y" Then  ' ***------- Check for Member Topup
                        XrefSql1 = "Select Top 1 MBMNumber, INSCode, MBMTopupNumber "
                        XrefSql2 = " AND MBMICBCPP = '" & Trim(VNewIC) & _
                                   "' And MBMStatus = 'A' AND PayCode ='" & VNewPayor & "' AND INSCode='" & VNewInsured & _
                                   "' ORDER BY MBMPayorEffDate DESC, MBMNumber "
                        Call SPMBMXref(NoOfRecLine, VNewPolicyNo, vNewName, VNewInsured, XrefSql1, XrefSql2, "2")
                        ' ------------------------ ***------ End of Checking for Topup
                    End If
                    If CDbl(VAge) > 60 Then
                        Call ErrorListing(NoOfRecLine, VNewPolicyNo, vNewName, "Please check Principal's Age")
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

                    Call SPHISTableRecVerify(SelFields, tmpTableName, SelCriteria, TmpRecFound, tmpCode, tmpDesc, tmpStatus)

                    If TmpRecFound = False Then
                        Call ErrorListing(NoOfRecLine, VNewPolicyNo, vNewName, "No Such Data Type - ' " & Trim(VNewDataType) & "' !")
                    Else
                        If Trim(VNewPlanCode) <> "" Then   ' HIS Supplementary Checking
                            If IsDate(VNewDOB) Then
                                Call GetSuppAgeLmt(NoOfRecLine, VNewDOB, "HIS", VNewPolicyNo, vNewName, VNewDataType, VNewPlanCode, VNewInsured, CDate(VNewEffDate), VNewExpDate)
                            End If
                        End If
                    End If
                    
                    If (Trim(VNewDataType) = "S" Or Trim(VNewDataType) = "D") And CDbl(VAge) > 22 Then
                        Call ErrorListing(NoOfRecLine, VNewPolicyNo, vNewName, "Please check Dependent's Age")
                    End If
                    
                    If (Trim(VNewDataType) = "H" Or Trim(VNewDataType) = "W") And CDbl(VAge) > 60 Then
                        Call ErrorListing(NoOfRecLine, VNewPolicyNo, vNewName, "Please check Dependent's Age")
                    End If
                    
                End If

' ************* CIS Varification
                If txtBordxFormat <> "3" Then
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
                    If VNewDataType = "P" Then
                        CisPrincPlnCode = VNewCLnPlan
                        CisPrincInsType = VNewCLnIns
                        CisprincEffDate = VNewClnEffDate
                        CisPrincExpdate = VNewClnExpDate
                    Else
                        VNewCLnPlan = IIf(Trim(VNewCLnPlan) = "", CisPrincPlnCode, VNewCLnPlan)
                        VNewCLnIns = IIf(Trim(VNewCLnIns) = "", CisPrincInsType, VNewCLnIns)
                        VNewClnEffDate = IIf(IsDate(VNewClnEffDate), CisprincEffDate, VNewClnEffDate)
                        VNewClnExpDate = IIf(IsDate(VNewClnExpDate), CisPrincExpdate, VNewClnExpDate)
                    End If
    
                    If Trim(VNewCLnPlan) <> "" Then
                        Call VerifyMBMByPassingPA(NoOfRecLine, "CIS", VNewDataType, " ", vNewName, VNewCLnPlan, VNewCLnIns, VNewClnEffDate, VNewClnExpDate)
                        If Trim(Mid(TextLine, 9229, 20)) = "" Then
                            Call ErrorListing(NoOfRecLine, VNewPolicyNo, vNewName, "Clinical Card Type is blank !! ")
                        End If
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
        'End If
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

    lstERRList.ListItems.Clear
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
                If Trim(Mid(TextLine, 2, 2)) = "" Or Trim(Mid(TextLine, 2, 2)) <> Mid(cboPayorCode, 1, 2) Then
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
                If Trim(Mid(TextLine, 1, 1)) = "C" And Trim(VEndtDataType) = "P" Then
                    SelFields = "Select MBMNumber as tmpCode, MBMStatus as tmpDesc "
                    SelCriteria = " and MBMStatus='C' AND MBMNumber='" & Trim(Mid(TextLine, 11, 18)) & "'"
                    tmpTableName = "MBMCrossReference"
                    tmpCode = ""
                    tmpDesc = ""
                    TmpRecFound = False
                    Call SPMaintTableRecVerify(SelFields, tmpTableName, SelCriteria, TmpRecFound, tmpCode, tmpDesc)
                    If TmpRecFound = True Then
                         Call ErrorListing(NoOfERecLine, VEndtPolicyNo, vEndtName, "Membership Number : " & Trim(Mid(TextLine, 11, 18)) & " has been cancelled !")
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
                    If Trim(VEndtInsured) <> "" Then
                        tmpTableName = "INS"
                        TmpRecFound = False
                        SelFields = "SELECT INSCode as tmpCode, INSDescription as tmpDesc "
                        SelCriteria = " And INSCode ='" & Trim(VEndtInsured) & "'"
                        tmpCode = ""
                        tmpDesc = ""
                        TmpRecFound = False
                        Call SPHISTableRecVerify(SelFields, tmpTableName, SelCriteria, TmpRecFound, tmpCode, tmpDesc, tmpStatus)
                        If TmpRecFound = False Then
                            Call ErrorListing(NoOfERecLine, VEndtPolicyNo, vEndtName, "No such Plan Code - ' " & Trim(VEndtPlanCode) & "' !")
                        End If
                    End If

                    If Trim(Mid(TextLine, 10, 1)) = "P" Then
                        If Trim(Mid(TextLine, 134, 90)) = "" Then
                            Call ErrorListing(NoOfERecLine, VEndtPolicyNo, vEndtName, "Address Is Blank")
                        End If
                        PaySearch = 2
                        SelFields = "SELECT PLNCode as tmpCode, PLNTopUpStatus as tmpDesc, PLNPayorCode as tmpPLNPayor "
                        SelCriteria = " And PLNCode ='" & Trim(VEndtPlanCode) & "'"
                        tmpTableName = "PLN"
                        tmpCode = ""
                        tmpDesc = ""
                        TmpRecFound = False
                        Call SPHISTableRecVerify(SelFields, tmpTableName, SelCriteria, TmpRecFound, tmpCode, tmpDesc, tmpStatus)
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
    Set ListError = lstERRList.ListItems.Add(, , tmpLineNo)
        ListError.SubItems(1) = Trim(tmpPol)
        ListError.SubItems(2) = Trim(tmpMBMName)
        ListError.SubItems(3) = Trim(tmpMSG)
        ListError.SubItems(4) = Trim(txtImportFilename)
End Sub

Private Sub GenerateMember()
On Err GoTo ErrorSave
Dim startTrans As Boolean
Dim StrSql As String
Dim PayorCode As String
Dim recordCount As Integer
Dim tmpVCovID As String
        tmpFileStatus = ""
        Open txtImportFilename For Input As #1   ' Open file.
        Line Input #1, TextLine   ' Read line into variable.
        
        tmpPAYFormat = Mid(TextLine, 28, 3)
        tmpfilename = Mid(TextLine, 12, 19)
        If Mid(TextLine, 2, 2) <> Mid(cboPayorCode, 1, 2) And (Trim(tmpPAYFormat) <> "MAA" And Trim(tmpPAYFormat) <> "TAK") Then
            MsgBox "Payor Selected Is Not Match The File Payor", vbOKOnly + vbCritical, DialogBoxTitle
            cboPayorCode.SetFocus
        Else
            If Mid(TextLine, 1, 1) = "F" Then
                PayorCode = Mid(TextLine, 2, 2)
                txtSubmissionDate = Mid(TextLine, 4, 2) & "/" & Mid(TextLine, 6, 2) & "/" & Mid(TextLine, 8, 4)
            ElseIf Mid(TextLine, 1, 1) = "E" Then
                txtSubmissionDate = Mid(TextLine, 4, 2) & "/" & Mid(TextLine, 6, 2) & "/" & Mid(TextLine, 8, 4)
            End If
            
            recordCount = 1
            
            medixmasterconn.beginTrans
            medixmasterconnMaint.beginTrans
            startTrans = True
            If tmpPAYFormat <> "MAA" And Trim(tmpPAYFormat) <> "TAK" Then
                txtBatchNo = CheckBatch(Mid(cboPayorCode, 1, 2), txtSubmissionDate, Mid(cboFileType, 1, 1))
            Else
                txtBatchNo = CheckBatch(tmpPAYFormat, txtSubmissionDate, Mid(cboFileType, 1, 1))
            End If
            Do While Not EOF(1)   ' Loop until end of file.
               Line Input #1, TextLine   ' Read line into variable.
                recordCount = recordCount + 1
            
                If Mid(TextLine, 1, 1) <> "F" And Mid(TextLine, 1, 1) <> "T" Then
                    
                    
                    If tmpPAYFormat <> "MAA" And Trim(tmpPAYFormat) <> "TAK" Then
                        If Mid(TextLine, 1, 1) = "P" Then
                            tmpVCovID = "00"
                        Else
                            tmpVCovID = tmpVCovID + 1
                        End If
                        tmpVCovID = Format(tmpVCovID, "00")
                    
                        Call NewMBMInfoProcessing(tmpVCovID)
                    
                    ElseIf (tmpPAYFormat = "MAA" Or tmpPAYFormat = "TAK") And Mid(TextLine, 1, 12) = "N" Then
                            If Mid(TextLine, 1, 1) = "P" Then
                                tmpVCovID = "00"
                            Else
                                tmpVCovID = tmpVCovID + 1
                            End If
                            tmpVCovID = Format(tmpVCovID, "00")
                        
                        Call NewMBMInfoProcessing(tmpVCovID)
                        
                    ElseIf (tmpPAYFormat = "MAA" Or tmpPAYFormat = "TAK") Then
                            If Mid(TextLine, 1, 1) = "P" Then
                                tmpVCovID = "00"
                            Else
                                tmpVCovID = tmpVCovID + 1
                            End If
                            tmpVCovID = Format(tmpVCovID, "00")
                        
                        Call NewMBMInfoProcessing(tmpVCovID)
                    'ElseIf tmpPAYFormat = "MAA" And (Mid(TextLine, 1, 12) = "R" Or Mid(TextLine, 1, 1) <> "P") Then
                    '    Call GenerateENDtMember
                    End If
                End If
                DoEvents
            Loop
            
            medixmasterconn.commitTrans
            medixmasterconnMaint.commitTrans
            MsgBox "Record Import Successfully...", vbOKOnly + vbInformation, DialogBoxTitle
            
            startTrans = False
            'Screen.MousePointer = vbDefault
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

Private Sub GenerateRRMember()
On Err GoTo ErrorSave
Dim startTrans As Boolean
Dim StrSql As String
Dim PayorCode As String
Dim recordCount As Integer
Dim tmpVCovID As String
        'tmpFileStatus = ""
        Open txtImportFilename For Input As #1   ' Open file.
        Line Input #1, TextLine   ' Read line into variable.
        
        tmpPAYFormat = Mid(TextLine, 28, 3)
        tmpfilename = Mid(TextLine, 12, 19)
        If Mid(TextLine, 2, 2) <> Mid(cboPayorCode, 1, 2) And (Trim(tmpPAYFormat) <> "MAA" And Trim(tmpPAYFormat) <> "TAK") Then
            MsgBox "Payor Selected Is Not Match The File Payor", vbOKOnly + vbCritical, DialogBoxTitle
            cboPayorCode.SetFocus
        Else
            If Mid(TextLine, 1, 1) = "F" Then
                PayorCode = Mid(TextLine, 2, 2)
                txtSubmissionDate = Mid(TextLine, 4, 2) & "/" & Mid(TextLine, 6, 2) & "/" & Mid(TextLine, 8, 4)
            ElseIf Mid(TextLine, 1, 1) = "E" Then
                txtSubmissionDate = Mid(TextLine, 4, 2) & "/" & Mid(TextLine, 6, 2) & "/" & Mid(TextLine, 8, 4)
            End If
            
            recordCount = 1
            
            medixmasterconn.beginTrans
            medixmasterconnMaint.beginTrans
            startTrans = True
            If tmpPAYFormat <> "MAA" And tmpPAYFormat <> "TAK" Then
                txtBatchNo = CheckBatch(Mid(cboPayorCode, 1, 2), txtSubmissionDate, Mid(cboFileType, 1, 1))
            Else
                txtBatchNo = CheckBatch(tmpPAYFormat, txtSubmissionDate, Mid(cboFileType, 1, 1))
            End If
            Do While Not EOF(1)   ' Loop until end of file.
               Line Input #1, TextLine   ' Read line into variable.
                recordCount = recordCount + 1
            
                If Mid(TextLine, 1, 1) <> "T" Then
                    
                    
                    'ElseIf tmpPAYFormat = "MAA" And Mid(TextLine, 1, 12) = "N" Then
                    '        If Mid(TextLine, 1, 1) = "P" Then
                    '            tmpVCovID = "00"
                    '        Else
                    '            tmpVCovID = tmpVCovID + 1
                    '        End If
                    '        tmpVCovID = Format(tmpVCovID, "00")
                        
                        'Call NewMBMInfoProcessing(tmpVCovID)
                    
                    If (tmpPAYFormat = "MAA" Or tmpPAYFormat = "TAK") And Mid(TextLine, 1, 1) = "P" Then
                            If Mid(TextLine, 1, 1) = "P" Then
                                tmpVCovID = "00"
                            Else
                                tmpVCovID = tmpVCovID + 1
                            End If
                            tmpVCovID = Format(tmpVCovID, "00")
                        Call NewMBMInfoProcessing(tmpVCovID)
                    ElseIf (tmpPAYFormat = "MAA" Or tmpPAYFormat = "TAK") And Mid(TextLine, 1, 1) = "R" Then
                        Call GenerateReinstateMember
                    ElseIf (tmpPAYFormat = "MAA" Or tmpPAYFormat = "TAK") And (Mid(TextLine, 1, 1) <> "P" Or Mid(TextLine, 1, 1) <> "R") Then
                            If Mid(TextLine, 1, 1) = "P" Then
                                tmpVCovID = "00"
                            Else
                                tmpVCovID = tmpVCovID + 1
                            End If
                            tmpVCovID = Format(tmpVCovID, "00")
                        Call NewMBMInfoProcessing(tmpVCovID)
                        
                    End If
                End If
                DoEvents
            Loop
            
            medixmasterconn.commitTrans
            medixmasterconnMaint.commitTrans
            MsgBox "Record Import Successfully...", vbOKOnly + vbInformation, DialogBoxTitle
            
            startTrans = False
            'Screen.MousePointer = vbDefault
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

Private Sub GenerateReinstateMember()
Dim tmpMBMNo As String
Dim tmpPolicyNo As String
Dim tmpPayCode As String
Dim tmpCancelMode As String
Dim strsqlPAY As String
Dim PAY As New ADODB.Recordset
Dim CancelExpDate As String
Dim RRMBMNumber As String
Dim tmpVDOB As String
Dim tDOB As String
Dim tmpVSex As String
Dim StrSql As String
Dim MBM As New ADODB.Recordset
    'tmpMBMNo = Trim(Mid(TextLine, 11, 18))
    'tmpVEffDate = Mid(TextLine, 374, 2) & "/" & Mid(TextLine, 376, 2) & "/" & Mid(TextLine, 378, 4)
    tmpPolicyNo = Trim(Mid(TextLine, 2, 25))
    RREffDate = Mid(TextLine, 374, 2) & "/" & Mid(TextLine, 376, 2) & "/" & Mid(TextLine, 378, 4)
    RRExpDate = Mid(TextLine, 382, 2) & "/" & Mid(TextLine, 384, 2) & "/" & Mid(TextLine, 386, 4)
    
    StrSql = ""
    StrSql = "Select MBMNumber, MBMStatus,  " & _
             "MBMPolicyNo, MBMICBcPp, MBMName, MBMPayorEXPDate, PLNCode, " & _
             "MBMPayorEffDate,MBMStatus from MBM Where MBMPolicyNo ='" & Trim(tmpPolicyNo) & "' and MBMPayorEXPDate = '" & Format(RRExpDate, "yyyy/mm/dd") & "'"
    MBM.Open StrSql, medixmasterconn, adOpenKeyset, adLockOptimistic

    
    If MBM.recordCount Then
        RRMBMNumber = MBM!MBMNumber
        
        Call NewUpdateReinstatement(RRMBMNumber, RREffDate, RRExpDate)
    Else
        Call NewMBMInfoProcessing("00")
    End If
    
    MBM.Close
    Set MBM = Nothing
End Sub

Private Sub NewUpdateReinstatement(RRMBMNumber As String, RREffDate As String, RRExpDate As String)
Dim tmpEndtBordxDate As String
Dim MBMCov As New ADODB.Recordset
Dim strsqlMBMCov As String
Dim StrSql As String
Dim tmpExpDate As String
Dim tmpINSType As String
Dim tmpPlnCode As String
Dim tmpAdultChild As String
Dim tmpPayCode As String
Dim tmpHLTType As String
Dim tmpLapsedDate As String
Dim tmpVEndtRefNo As String

        Screen.MousePointer = vbHourglass
            tmpVEndtRefNo = Trim(Mid(TextLine, 3604, 20))
            tmpEndtBordxDate = Format(txtSubmissionDate, "YYYY/MM/DD")
            tmpReinstateDate = Format(RREffDate, "YYYY/MM/DD")
            tmpLapsedDate = ""
            'tmpExpDate = Format(txtExpDate, "YYYY/MM/DD")
            Call MemberCallFees
                    'tmpMCO = IIf(IsNumeric(Trim(Mid(TextLine, 3696, 12))), Trim(Mid(TextLine, 3696, 12)), 0)
        'tmpBMCO = IIf(IsNumeric(Trim(Mid(TextLine, 3708, 12))), Trim(Mid(TextLine, 3708, 12)), 0)

            tempOriMCO = IIf(IsNumeric(Trim(Mid(TextLine, 3696, 12))), Trim(Mid(TextLine, 3696, 12)), 0)
            tempMCOFee = IIf(IsNumeric(Trim(Mid(TextLine, 3684, 12))), Trim(Mid(TextLine, 3684, 12)), 0)
            'IIf(IsNumeric(Trim(Mid(TextLine, 3708, 12))), Trim(Mid(TextLine, 3708, 12)), 0)
            'tempOriMCO = CDbl(tempOriMCO) + CDbl(tempMCOFee)
            'tempOriMCO = Round(tempOriMCO)
            tempOriPrem = CDbl(tempOriPrem) + CDbl(TempPremium)
            StrSql = ""
            
            StrSql = " insert into MBMReinstatement (MBMNumber, MBMCoverID, MBMEndtBordxDate, MBMMCOFee, MBMPremium, " & _
                    " MBMLapsedDate, MBMReinstateDate, LastUpdateUSR, LastUpdateDate, MBMEndtBordxBatch, MBMFileName,MBMENDtRefNo)" & _
                    " Values('" & RRMBMNumber & "', '00', '" & Format(tmpEndtBordxDate, "yyyy/mm/dd") & "', " & tempMCOFee & ", " & CDbl(TempPremium) & ", " & _
                    " '" & tmpLapsedDate & "', '" & Format(tmpReinstateDate, "yyyy/mm/dd") & "', '" & ChkStr(Logon) & "', '" & Format(Date, "yyyy/mm/dd") & "', " & txtBatchNo & " , '" & tmpfilename & "','" & tmpVEndtRefNo & "'  )"
            medixmasterconn.Execute StrSql
            
            StrSql = ""
            'strsql = "Update MBM set MBMStatus = 'A', MBMFinalExpiryDate = '" & tmpExpDate & "'  Where MBMNumber = '" & Trim(txtMBMNumber) & "'"
            StrSql = "Update MBM set MBMStatus = 'A'  Where MBMNumber = '" & Trim(RRMBMNumber) & "'"
            medixmasterconn.Execute StrSql
            
            'strsql = ""
            'strsql = "Update MBMOthers set MBMENDtRefNo = '" & tmpVEndtRefNo & "'  Where MBMNumber = '" & Trim(RRMBMNumber) & "'"
            'medixmasterconn.Execute strsql
            
            
            'strsql = ""
            'strsql = "Update MBMCrossReference set MBMStatus = 'A' , MBMFinalExpiryDate = '" & tmpExpDate & "'  Where MBMNumber = '" & Trim(txtMBMNumber) & "'"
            'strsql = "Update MBMCrossReference set MBMStatus = 'A'  Where MBMNumber = '" & Trim(RRMBMNumber) & "'"
            'medixmasterconnMaint.Execute strsql
            
            
        Screen.MousePointer = Default
    'End If

End Sub

Public Sub MemberCallFees()
Dim PLNRec As New ADODB.Recordset
Dim SqlPln As String
Dim tmpPremSexStatus As String
Dim RoundUpStatus, PLNProRate As String
Dim tmpAGECode As String
Dim tmpSex As String
Dim tmpICNo As String
Dim tFindAge As String
Dim tAgeCode As String
Dim tAdultChild As String
Dim tmpVDOB As String
Dim tDOB As String
Dim tmpVSex As String
Dim tmpPayCode As String
Dim tmpVPlanCode As String
Dim tmpPayorPLN As String
Dim tPayor As String
Dim SelFields As String
'Dim SelFields As String
Dim SelCriteria As String
'Dim tmpPayorPLN As String
Dim tmpTableName As String
Dim tmpCode As String
Dim tmpDesc As String
Dim TmpRecFound As Boolean
'Dim tmpVPlanCode As String
Dim StrSql As String
Dim MBM As New ADODB.Recordset
Dim tHncInsType As String
'Dim RoundUpStatus As String
'Dim PLNProRate As String
'Dim tmpPremSexStatus As String
    'If Trim(tmpLapsedDate) <> "__/__/____" _
    '    And Trim(tmpHLTType) <> "__/__/____" _
    '    And Trim(tmpPlnCode) <> "" And Trim(tmpINSType) <> "" Then
    
        tmpPayorPLN = Trim(Mid(TextLine, 390, 7))
        SelFields = "SELECT PLNCode as tmpCode, PLNLifeTimeStatus as tmpDesc, PLNPayorCode as tmpPLNPayor "
        SelCriteria = " And PLNPayorPlan ='" & Trim(tmpPayorPLN) & "'"
        tmpTableName = "PLN"
        tmpCode = ""
        tmpDesc = ""
        TmpRecFound = False
        Call SPHISTableRecVerify(SelFields, tmpTableName, SelCriteria, TmpRecFound, tmpCode, tmpDesc, tmpStatus, tmpPLNPayor)
            If TmpRecFound = True Then
                tmpVPlanCode = tmpCode
                tPayor = tmpPLNPayor
            Else
                tmpVPlanCode = ""
                tmpPLNPayor = ""
            End If
    
        tmpPayCode = tPayor
    
        tmpVDOB = Format(Mid(TextLine, 293, 2) & "/" & Mid(TextLine, 295, 2) & "/" & Mid(TextLine, 297, 4), "yyyy/mm/dd")
        tDOB = Format(tmpVDOB, "YYYY/MM/DD")
        tmpVSex = Trim(Mid(TextLine, 301, 1))
        tFindAge = GetAge(CDate(RREffDate), CDate(tmpVDOB))
        tAgeCode = Mid(GetAgeBand(tHLTCode, tFindAge, tmpVPlanCode, tHncInsType), 1, 2)
        tAdultChild = "A"
        
        If txtBordxFormat = "3" Then
            tHncInsType = Trim(Mid(TextLine, 397, 1))
        Else
            tHncInsType = Trim(Mid(TextLine, 394, 1))
        End If
        
        If tAgeCode = "01" Then
            tAdultChild = "C"
        End If
        

        SqlPln = "SELECT PLNRoundUpStatus, PLNProRate,PLNPremGenderStatus from PLN where PlnCode='" & Trim(tmpVPlanCode) & _
        "' and HLTCode = '" & Mid(cboHLTCode1, 1, 1) & _
        "' AND PLNEffDate <='" & Format(tmpReinstateDate, "mm/dd/yyyy") & "'"
        
        If Mid(cboHLTCode1, 1, 1) = "N" Then
            SqlPln = SqlPln & " and PAYCode='" & Trim(tmpPayCode) & "'"
        End If

        PLNRec.Open SqlPln, medixmasterconn, adOpenKeyset, adLockOptimistic
        RoundUpStatus = "Y"
        PLNProRate = "Y"
        If PLNRec.recordCount Then
            RoundUpStatus = IIf(Len(Trim(PLNRec!PLNRoundUpStatus)), Trim(PLNRec!PLNRoundUpStatus), "Y")
            PLNProRate = IIf(Len(Trim(PLNRec!PLNProRate)), Trim(PLNRec!PLNProRate), "Y")
            tmpPremSexStatus = IIf(Len(Trim(PLNRec!PLNPremGenderStatus)), Trim(PLNRec!PLNPremGenderStatus), "N")
        End If
        PLNRec.Close
        Set PLNRec = Nothing
        
        If Trim(tAdultChild) <> "" Then
            tempMCOFee = Format(GetMCO(Trim(Mid(tHncInsType, 1, 1)), Mid(cboHLTCode1, 1, 1), tAdultChild, CDate(tmpReinstateDate), tmpVPlanCode), "#,##0.00")
            If PLNProRate <> "N" Then
                tempMCOFee = InsertMCOFees(CDate(RRExpDate), CDate(tmpReinstateDate), Format(tempMCOFee, "#,##0.00"))
            End If
            tempMCOFee = Format(tempMCOFee, "#,##0.00")
        End If

        If RoundUpStatus <> "N" Then
            tempMCOFee = Round(tempMCOFee)
        End If
        
        TempPremium = Format(GetPremium(Trim(Mid(tHncInsType, 1, 1)), Trim(Mid(tmpPayCode, 1, 2)), Trim(Mid(tAgeCode, 1, 2)), Trim(tmpVPlanCode), Trim(Mid(cboHLTCode1, 1, 1)), CDate(tmpReinstateDate), Mid(tmpVSex, 1, 1)), "#,##0.00")
        
        TempPremium = Format(InsertPremium(CDate(RRExpDate), CDate(tmpReinstateDate), Format(TempPremium, "#,##0.00")), "#,##0.00")
    'End If
End Sub

Private Sub GenerateENDtMember()
Dim tmpVEndtType As String
    Open txtImportFilename For Input As #1   ' Open file.
        Line Input #1, TextLine   ' Read line into variable.
            tmpfilename = Mid(TextLine, 12, 19)
            Do While Not EOF(1)   ' Loop until end of file.
                Line Input #1, TextLine   ' Read line into variable.
                    If Mid(TextLine, 1, 1) <> "T" Then
                    Call CancelENDtData
                    End If
                DoEvents
            Loop
            MsgBox "Record Import Sucessfully...", vbOKOnly + vbInformation, DialogBoxTitle
    Close #1   ' Close file.
End Sub

Private Sub CancelENDtData()
Dim tmpMBMNo As String
Dim tmpPolicyNo As String
Dim tmpPayCode As String
Dim tmpCancelMode As String
Dim strsqlPAY As String
Dim PAY As New ADODB.Recordset
Dim CancelExpDate As String
Dim SelFields As String
Dim SelCriteria As String
Dim tmpPayorPLN As String
Dim tmpTableName As String
Dim tmpCode As String
Dim tmpDesc As String
Dim TmpRecFound As Boolean
Dim tmpVPlanCode As String

    'tmpMBMNo = Trim(Mid(TextLine, 11, 18))
    'tmpVEffDate = Mid(TextLine, 374, 2) & "/" & Mid(TextLine, 376, 2) & "/" & Mid(TextLine, 378, 4)
        tmpPayorPLN = Trim(Mid(TextLine, 390, 7))
        SelFields = "SELECT PLNCode as tmpCode, PLNLifeTimeStatus as tmpDesc, PLNPayorCode as tmpPLNPayor "
        SelCriteria = " And PLNPayorPlan ='" & Trim(tmpPayorPLN) & "'"
        tmpTableName = "PLN"
        tmpCode = ""
        tmpDesc = ""
        TmpRecFound = False
        Call SPHISTableRecVerify(SelFields, tmpTableName, SelCriteria, TmpRecFound, tmpCode, tmpDesc, tmpStatus, tmpPLNPayor)
            If TmpRecFound = True Then
                tmpVPlanCode = tmpCode
                tPayor = tmpPLNPayor
            Else
                tmpVPlanCode = ""
                tmpPLNPayor = ""
            End If
    
    CancelExpDate = Mid(TextLine, 382, 2) & "/" & Mid(TextLine, 384, 2) & "/" & Mid(TextLine, 386, 4)
    CancelDate = Mid(TextLine, 374, 2) & "/" & Mid(TextLine, 376, 2) & "/" & Mid(TextLine, 378, 4)
    tmpPayCode = tPayor
    tmpPolicyNo = Trim(Mid(TextLine, 2, 25))
        
    'strsqlPAY = "Select PAYCancelMode, PAYCode from PAY where PAYCode = '" & tmpPayCode & "' "
    'Pay.Open strsqlPAY, medixmasterconn, adOpenKeyset, adLockOptimistic
    
    'If Pay.recordCount Then
    '    If Trim(Pay!PAYCancelMode) = "P" Then
    '
    'tmpCancelMode = Pay!PAYCancelMode
    '    Else
    '        tmpCancelMode = ""
    '    End If
    'End If
    
    
    Call EndtCancelPrincipalData(tmpPolicyNo, CDate(CancelDate), CDate(CancelExpDate))
    'Call EndtCancelCoveredPersonsData(tmpMBMNo, CDate(CancelDate))
    
    
End Sub

Private Sub EndtCancelCoveredPersonsData(tmpMBMNo As String, CancelDate As Date)
Dim StrSql As String
Dim MBMCoveredPersons As New ADODB.Recordset
Dim tmpCPrem, tmpCMco As Double
Dim tmpExpDate As String
    
    StrSql = "Select MBMCStatus, MBMCExpDate, MBMCPremium, MBMCMCO, MBMCCoverID " & _
             "from MBMCoveredPersons Where MBMCNumber ='" & Trim(tmpMBMNo) & "'"
    MBMCoveredPersons.Open StrSql, medixmasterconn, adOpenKeyset, adLockOptimistic
    
    If MBMCoveredPersons.recordCount Then
       Do Until MBMCoveredPersons.EOF
           With MBMCoveredPersons
                tmpExpDate = IIf(Len(!MBMCExpDate), !MBMCExpDate, Null)
                !MBMCStatus = "C"
                tmpCPrem = IIf(IsNumeric(!MBMCPremium), !MBMCPremium, 0)
                tmpCMco = IIf(IsNumeric(!MBMCMCO), !MBMCMCO, 0)
                tmpExpDate = IIf(Len(tmpExpDate), tmpExpDate, tmpEndtPrincExpDate)
                If tmpCPrem > 0 Or tmpCMco > 0 Then
                     Call EndtMBMRefund(Trim(tmpMBMNo), CDate(tmpEndtPrincEffDate), CDate(tmpExpDate), CDate(CancelDate), !MBMCCoverID, CInt(tmpCPrem), CInt(tmpCMco), txtMBMCancelStatus)
                End If
                MBMCoveredPersons.Update
           End With
           MBMCoveredPersons.MoveNext
       Loop
    End If
    MBMCoveredPersons.Close
    Set MBMCoveredPersons = Nothing
End Sub

Private Sub EndtCancelPrincipalData(tmpMBMNo As String, CancelDate As Date, CancelExpDate As Date)
Dim StrSql As String
Dim MBM As New ADODB.Recordset
Dim StrAmendSql As String
Dim MBMMCOFees As Double
Dim MBMPremium As Double
Dim SqlMBMO As String
Dim MBMOthers As New ADODB.Recordset
Dim MBMTwoSql, XMBMSql As String
Dim tmpName As String

    tmpCancelstatus = False
    StrSql = ""
    StrSql = "Select MBMNumber, MBMStatus,  " & _
             "MBMPolicyNo, MBMICBcPp, MBMName, MBMPayorEXPDate, PLNCode, " & _
             "MBMPayorEffDate,MBMStatus from MBM Where MBMPolicyNo ='" & Trim(tmpMBMNo) & "' and MBMPayorEXPDate = '" & Format(CancelExpDate, "yyyy/mm/dd") & "' and MBMStatus <> 'D'"
    MBM.Open StrSql, medixmasterconn, adOpenKeyset, adLockOptimistic

    
    If MBM.recordCount Then
        
        If MBM!MBMStatus = "A" Then
            txtTempMemberNo = Trim(MBM!MBMNumber)
            MBMTwoSql = "Update MBMTwo SET MBMLastAdjustmentDate = '" & Format(Date, "YYYY/MM/DD") & "', MBMLastUpdateUser= '" & Logon & _
                        "', MBMLastEndorseStatus = 'C' , MBMENDtBordxDate = '" & Format(txtSubmissionDate, "YYYY/MM/DD") & _
                        "', MBMCancelDate ='" & Format(CancelDate, "yyyy/mm/dd") & _
                        "', MBMENDtEffDate = '" & Format(CancelDate, "YYYY/MM/DD") & _
                        "', MBMFilename = '" & tmpfilename & "' WHERE MBMNumber = '" & Trim(txtTempMemberNo) & "'"
            medixmasterconn.Execute MBMTwoSql
        
            XMBMSql = "Update MBMCrossReference SET MBMStatus = 'C' WHERE MBMNumber = '" & Trim(txtTempMemberNo) & "'"
            medixmasterconnMaint.Execute XMBMSql
            'MBMOthers.Close
            'Set MBMOthers = Nothing
            'SqlMBMO = ""
            SqlMBMO = "Select MBMPremium, MBMMCOFee, MBMAccessFee,MBMENDtRefNo from MBMOthers where MBMNumber = '" & Trim(txtTempMemberNo) & "'"
            MBMOthers.Open SqlMBMO, medixmasterconn, adOpenKeyset, adLockOptimistic
            
            If MBMOthers.recordCount Then
                MBMPremium = MBMOthers!MBMPremium
                MBMMCOFees = MBMOthers!MBMMCOFee
               
            End If
            'With MBMOthers
            '    !MBMENDtRefNo = tmpVEndtRefNo
            '    MBMOthers.Update
            'End With

            
            
            MBMOthers.Close
            Set MBMOthers = Nothing
    
            tmpName = Trim(Replace(tmpName, "'", "''"))
            StrAmendSql = "Insert into MBMAmendHistory(MBMNumber, MBMPolicyNo, MBMICBcPp, MBMName, " & _
                "MBMExpiryDate, MBMPlanCode, MBMAmendEdtType, " & _
                "MBMAEndtBordxDate, MBMAEndtEffDate, " & _
                "MBMAmendUser, MBMAmendDate, MBMAEndtBatchNo) " & _
                "Values ('" & MBM!MBMNumber & "','" & MBM!MBMPolicyNo & "','" & MBM!MBMIcBcPp & "','" & tmpName & "'," & _
                "'" & Format(MBM!MBMPayorEXPDate, "YYYY/MM/DD") & "','" & MBM!PLNCode & "','C'," & _
                "'" & Format(txtSubmissionDate, "YYYY/MM/DD") & "','" & Format(CancelDate, "YYYY/MM/DD") & "'," & _
                "'" & Logon & "','" & Format(Date, "YYYY/MM/DD") & "','" & txtBatchNo & "')"
            medixmasterconn.Execute StrAmendSql
            
            With MBM
                tmpEndtPrincEffDate = MBM!MBMPayorEffDate
                tmpEndtPrincExpDate = MBM!MBMPayorEXPDate
                If Len(MBM!MBMStatus) Then
                    !MBMStatus = "C"
                    MBM.Update
                End If
            End With
                Call EndtMBMRefund(Trim(txtTempMemberNo), CDate(tmpEndtPrincEffDate), CDate(tmpEndtPrincExpDate), CDate(CancelDate), "00", CInt(MBMPremium), CInt(MBMMCOFees), txtMBMCancelStatus)
            Else
            
            
            End If
       ' Else
       '     tmpCancelstatus = True
       '     Call NewMBMInfoProcessing("00")
        End If
        MBM.Close
        Set MBM = Nothing
    
End Sub

Private Sub EndtMBMRefund(tmpMBMNo As String, PayEffDate As Date, PAYExpDate As Date, CancelDate As Date, TmpCovID As String, tmpPremium As Double, tmpMCO As Double, MBMCancelStatus As String)
Dim MBMMCOFees As Double
Dim HLTCode As String
Dim RefundSql As String
Dim MCORefundValue As Double
Dim PremiumRefundValue As Double
Dim PremiumEffective As Double
Dim MCOEffective As Double
Dim tmpVEndtRefNo As String
    
    tmpVEndtRefNo = Trim(Mid(TextLine, 3604, 20))

    HLTCode = Mid(cboHLTCode1, 1, 1)
    
    MCORefundValue = Format(MCORefund(CDate(PAYExpDate), CDate(CancelDate), CDate(PayEffDate), tmpMCO, MBMCancelStatus), "#,##0.00")
    
    PremiumRefundValue = Format(MCOPremiumReduction(CDate(PAYExpDate), CDate(CancelDate), CDate(PayEffDate), tmpPremium), "#,##0.00")
    PremiumEffective = Format(CDbl(tmpPremium) - CDbl(PremiumRefundValue), "#,##0.00")
    
    MCORefundValue = IIf(IsNumeric(Trim(Mid(TextLine, 3684, 12))), CDbl(Trim(Mid(TextLine, 3684, 12)) * -1), 0)
    'tmpBMCO = IIf(IsNumeric(Trim(Mid(TextLine, 3696, 12))), Trim(Mid(TextLine, 3696, 12)), 0)
    
    MCOEffective = Format(CDbl(tmpMCO) - CDbl(MCORefundValue), "#,##0.00")
    RefundSql = "EXEC Upload_MBMMcoRefund_2010V3 '" & Trim(tmpMBMNo) & "','" & Trim(TmpCovID) & _
                "'," & PremiumEffective & "," & PremiumRefundValue & "," & MCOEffective & _
                "," & MCORefundValue & ",'" & Mid(cboPayorCode, 1, 2) & "'," & _
                "'" & Format(txtSubmissionDate, "yyyy/mm/dd") & "','" & txtBatchNo & "','" & HLTCode & "', '" & tmpfilename & "', '" & tmpVEndtRefNo & "','" & Format(CancelDate, "yyyy/mm/dd") & "'"
    medixmasterconn.Execute RefundSql
    
End Sub

Private Sub EndtDeleteSupplementary()
Dim strdelcov As String
Dim MBMCov As New ADODB.Recordset
Dim VEndtSuppID As String
    VEndtSuppID = Trim(Mid(TextLine, 2963, 2))
    strdelcov = "Delete from MBMCoveredPersons Where MBMCNumber ='" & Mid(TextLine, 11, 18) & "'" 'And MBMCCoverID ='" & VEndtSuppID & "'"
    MBMCov.Open strdelcov, medixmasterconn, adOpenKeyset, adLockOptimistic

    MBMCov.Close
    Set MBMCov = Nothing
End Sub

Private Sub EndtAddSupplementary()
Dim StrSql As String
Dim LastCoveredID As Integer
Dim SelectAge As String
Dim MBMCov As New ADODB.Recordset
Dim AgeValue As String
Dim tmpPlanCode As String
Dim tmpINSType As String

    tmpPlanCode = Trim(Mid(TextLine, 417, 4))
    tmpINSType = Trim(Mid(TextLine, 421, 1))
    StrSql = "Select * from MBMCoveredPersons Where MBMCNumber = '" & Trim(Mid(TextLine, 11, 18)) & "'"
    MBMCov.Open StrSql, medixmasterconn, adOpenKeyset, adLockOptimistic

    LastCoveredID = GetLastCoveredID(Trim(Mid(TextLine, 11, 18)))
        MBMCov.AddNew
        With MBMCov
            !MBMCNumber = Trim(Mid(TextLine, 11, 18))
            !MBMCRELCode = Mid(TextLine, 10, 1)
            SelectAge = GetAge(Mid(TextLine, 401, 2) & "/" & Mid(TextLine, 403, 2) & "/" & Mid(TextLine, 405, 4), Mid(TextLine, 320, 2) & "/" & Mid(TextLine, 322, 2) & "/" & Mid(TextLine, 324, 4))
            AgeValue = Mid(GetAgeBand(Mid(cboHLTCode1, 1, 1), SelectAge, tmpPlanCode, tmpINSType), 1, 2)
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
    Dim StrSql As String
    Dim listrecord As Integer
    Dim listloop As Integer
    Dim ListCount As Integer
    Dim strsql2 As String
    Dim LastCoveredID As Integer
    Dim SelectAge As String
    Dim MemberAmendHis As New ADODB.Recordset
    Dim EndtExpDate As String
    
    If Mid(TextLine, 10, 1) = "P" Then
        StrSql = "Select * from MBM Where MBMNumber = '" & Trim(Mid(TextLine, 11, 18)) & "'"
        MBM.Open StrSql, medixmasterconn, adOpenKeyset, adLockOptimistic
                
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
            !MBMEndtBordxDate = CDate(txtSubmissionDate)
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
Dim StrSql As String
Dim strsql2 As String
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
Dim tmpPlanCode As String
Dim tmpINSType As String

    tmpPlanCode = Trim(Mid(TextLine, 417, 4))
    tmpINSType = Trim(Mid(TextLine, 421, 1))

    strsql2 = "select * from MBM where MBMNumber = '" & Mid(TextLine, 11, 18) & "'"
    MBM.Open "MBM", medixmasterconn, adOpenKeyset, adLockOptimistic
    SuppID = Mid(TextLine, 2963, 2)
    If Trim(SuppID) = "" Then
        StrSql = "Select * from MBMCoveredPersons Where MBMCNumber = '" & Mid(TextLine, 11, 18) & "' AND MBMCICBCPPNo='" & Mid(TextLine, 74, 60) & "'"
        MBMCov.Open StrSql, medixmasterconn, adOpenKeyset, adLockOptimistic
        If MBMCov.recordCount = 0 Then
            MBMCov.Close
            Set MBMCov = Nothing
            StrSql = "Select * from MBMCoveredPersons Where MBMCNumber = '" & Mid(TextLine, 11, 18) & "' AND MBMCICBCPPNo='" & Mid(TextLine, 300, 20) & "'"
            MBMCov.Open StrSql, medixmasterconn, adOpenKeyset, adLockOptimistic
        End If
    Else
        SuppID = Mid(TextLine, 2963, 2)
        StrSql = "Select * from MBMCoveredPersons Where MBMCNumber = '" & Mid(TextLine, 11, 18) & "' and MBMCCoverID = '" & SuppID & "'"
        MBMCov.Open StrSql, medixmasterconn, adOpenKeyset, adLockOptimistic
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
            AgeValue = Mid(GetAgeBand(Mid(cboHLTCode1, 1, 1), FindAge, tmpPlanCode, tmpINSType), 1, 2)
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
Dim tmpSql As String
Dim DiscPlan As New ADODB.Recordset
Dim tmpSex As String
    MBM.Open "MBM", medixmasterconn, adOpenKeyset, adLockOptimistic
    
    MBM.AddNew
    With MBM
        If Mid(TextLine, 10, 1) = "P" Then
            !MBMLastEndorseStatus = Mid(TextLine, 1, 1)
            !MBMEndtBordxDate = CDate(txtSubmissionDate)
            !MBMENDtEffDate = Mid(TextLine, 2, 2) & "/" & Mid(TextLine, 4, 2) & "/" & Mid(TextLine, 6, 4)
            !MBMMemberType = "P"
            !MBMPolicyNo = Mid(TextLine, 29, 25)
            !MBMName = Mid(TextLine, 74, 60)
            !MBMAddress1 = Mid(TextLine, 134, 30)
            !MBMAddress2 = Mid(TextLine, 164, 30)
            !MBMAddress3 = Mid(TextLine, 194, 30)
            !PAYCode = Mid(cboPayorCode, 1, 2)
            !HLTCode = Mid(cboHLTCode1, 1, 1)
            !CTYCode = Mid(TextLine, 224, 20)
            !MBMPostCode = Mid(TextLine, 244, 5)
            !STACode = Mid(TextLine, 249, 15)
            !MBMIcBcPp = Mid(TextLine, 300, 20)
            !MBMDOB = Mid(TextLine, 320, 2) & "/" & Mid(TextLine, 322, 2) & "/" & Mid(TextLine, 324, 4)
            !MBMSex = Mid(TextLine, 328, 1)
            tmpSex = Mid(TextLine, 328, 1)
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
                !MBMBatchNo = txtBatchNo
            End If
            !MBMExpiredStatus = "N"
    '------------------------------------------------------------------------------------------------------------------------
            PayEffDate = Mid(TextLine, 401, 2) & "/" & Mid(TextLine, 403, 2) & "/" & Mid(TextLine, 405, 4)
            PAYExpDate = Mid(TextLine, 409, 2) & "/" & Mid(TextLine, 411, 2) & "/" & Mid(TextLine, 413, 4)
            MemberDOB = Mid(TextLine, 320, 2) & "/" & Mid(TextLine, 322, 2) & "/" & Mid(TextLine, 324, 4)
            FindAge = GetAge(CDate(PayEffDate), CDate(MemberDOB))
            AgeValue = Mid(GetAgeBand(Mid(cboHLTCode1, 1, 1), FindAge, plancode, InsuredCode), 1, 2)
            !AGECode = AgeValue
            If AgeValue = "01" Then
                !MBMAdultChild = "C"
            Else
                !MBMAdultChild = "A"
            End If
            
            txtMBMPremium = Format(GetPremium(InsuredCode, Payor, AgeValue, plancode, Mid(cboHLTCode1, 1, 1), CDate(PayEffDate), Trim(tmpSex)), "#,##0.00")
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
                                                 
            tmpPreMemberNo = Payor & plancode & InsuredCode & GenerateMembershipNo(Payor, InsuredCode)
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
        !MBMFirstMemNo = txtTempMemberNo
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
Dim tmpPlanCode As String
Dim tmpINSType As String

    tmpPlanCode = Trim(Mid(TextLine, 417, 4))
    tmpINSType = Trim(Mid(TextLine, 421, 1))
        
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
        AgeValue = Mid(GetAgeBand(Mid(cboHLTCode1, 1, 1), FindAge, tmpPlanCode, tmpINSType), 1, 2)
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

Private Sub GetSuppAgeLmt(tmpLine As Integer, DateOfBirth, tmpProdType, tmpPOlNo As String, tmpName As String, tmpDataType As String, tmpPlanCode, tmpINSType As String, TmpEffDate, tmpExpDate As String)
Dim CheckYr As Long
Dim PLNRec As New ADODB.Recordset
Dim PLNSql As String
Dim OverAge As Boolean
Dim PlnAgeLmt As Integer
Dim SuppDobEffDate As String

    SuppDobEffDate = ""
    PlnAgeLmt = 0
    OverAge = False
    SuppDobEffDate = Format(DateOfBirth, "dd/mm") & "/" & Year(TmpEffDate)
    CheckYr = Year(TmpEffDate) - Year(DateOfBirth)
    'medixmasterconn.Open "File Name=" & BOSPath & "Conn\mediconnect.UDL"

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
    'medixmasterconn.Close
    'Set medixmasterconn = Nothing
    
    If PlnAgeLmt <> 0 Then
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

Private Function GetSuppExpDate(DateOfBirth, tmpExpDate As String, PlnAgeLmt As String)
Dim CheckYr As Long
Dim SuppDobExpDate As String

    SuppDobExpDate = Format(DateOfBirth, "dd/mm") & "/" & Year(tmpExpDate)
    CheckYr = Year(tmpExpDate) - Year(DateOfBirth)
    GetSuppExpDate = tmpExpDate
    If Trim(PlnAgeLmt) <> "" Then
        If CheckYr = PlnAgeLmt Then
            If CDate(SuppDobExpDate) <= CDate(tmpExpDate) Then
                GetSuppExpDate = CDate(SuppDobExpDate)
            End If
        End If
    End If

End Function

Private Sub SaveMBMXRef(XNumber As String, XCovID As String, XPolNo As String, XIC As String, _
XName As String, XEffdate As String, XExpDate As String, XAgeCode As String, XInsCode As String, XPlan As String, _
XBordxDate As String, XTopupNo As String, XTopupIND As String, _
XDataStatus As String, XStatus As String, XDOb As String, XHltCode As String, XPayor As String, XBatchNo As String)
Dim XrefWSql As String
Dim XrefRec As New ADODB.Recordset
Dim tEffDate, tExpDate, tBordxDate As String
    tEffDate = IIf(IsDate(XEffdate), Format(XEffdate, "yyyy/mm/dd"), "")
    tExpDate = IIf(IsDate(XExpDate), Format(XExpDate, "yyyy/mm/dd"), "")
    tBordxDate = IIf(IsDate(XBordxDate), Format(XBordxDate, "yyyy/mm/dd"), "")

    XrefWSql = "EXEC Upload_MBMXref_HIS '" & XNumber & "', '" & Format(XCovID, "00") & "', '" & XPolNo & "', '" & XIC & "'," & _
            "'" & Trim(XName) & "', '" & tEffDate & "','" & tExpDate & "', '" & XAgeCode & "','" & XInsCode & "','" & XPlan & "'," & _
            "'" & XBordxDate & "', '" & Trim(XTopupNo) & "', '" & Trim(XTopupIND) & "'," & _
            "'" & XDataStatus & "','" & XStatus & "', '" & XDOb & "','" & XHltCode & "', '" & XPayor & "', '" & Trim(XBatchNo) & "'"
    medixmasterconn.Execute XrefWSql

End Sub

Private Sub SaveMBMXRefAXA(XNumber As String, XCovID As String, XPolNo As String, XIC As String, _
XName As String, XEffdate As String, XExpDate As String, XAgeCode As String, XInsCode As String, XPlan As String, _
XBordxDate As String, XTopupNo As String, XTopupIND As String, _
XDataStatus As String, XStatus As String, XDOb As String, XHltCode As String, XPayor As String, XBatchNo As String)
Dim XrefWSql As String
Dim XrefRec As New ADODB.Recordset
Dim tEffDate, tExpDate, tBordxDate As String
    tEffDate = IIf(IsDate(XEffdate), Format(XEffdate, "yyyy/mm/dd"), "")
    tExpDate = IIf(IsDate(XExpDate), Format(XExpDate, "yyyy/mm/dd"), "")
    tBordxDate = IIf(IsDate(XBordxDate), Format(XBordxDate, "yyyy/mm/dd"), "")

    XrefWSql = "EXEC Upload_MBMXref_HIS '" & XNumber & "', '" & Format(XCovID, "00") & "', '" & XPolNo & "', '" & XIC & "'," & _
            "'" & Trim(XName) & "', '" & tEffDate & "','" & tExpDate & "', '" & XAgeCode & "','" & XInsCode & "','" & XPlan & "'," & _
            "'" & XBordxDate & "', '" & Trim(XTopupNo) & "', '" & Trim(XTopupIND) & "'," & _
            "'" & XDataStatus & "','" & XStatus & "', '" & XDOb & "','" & XHltCode & "', '" & XPayor & "', '" & Trim(XBatchNo) & "'"
    medixmasterconn.Execute XrefWSql

End Sub

Private Sub cboPayorCode_LostFocus()
Dim tmpTableName As String
Dim SelFields As String
Dim SelCriteria As String
Dim tmpCode As String
Dim tmpDesc As String
Dim TmpRecFound  As Boolean
    PaySearch = 1
    SelFields = "SELECT CLNStatus as tmpCode, MBMCancelStatus as tmpDesc, PAYMBMRegStatus as tmpStatus "
    SelCriteria = " And PAYCode ='" & Mid(cboPayorCode, 1, 2) & "'"
    tmpTableName = "PAY"
    tmpCode = "N"
    tmpDesc = "N"
    tmpStatus = ""
    TmpRecFound = False
    Call SPHISTableRecVerify(SelFields, tmpTableName, SelCriteria, TmpRecFound, tmpCode, tmpDesc, tmpStatus)
    txtClnStatus = IIf(Len(tmpCode), tmpCode, "N")
    txtMBMCancelStatus = IIf(Len(tmpDesc), tmpDesc, "N")
    txtPayorMBMNoType = IIf(Len(tmpStatus), tmpStatus, "")
End Sub

Private Sub VerifyMBMByPassingPA(tmpLine As Integer, tmpProdType As String, tmpPOlNo As String, tmpName As String, tmpDataType As String, tmpPlanCode As String, tmpINSType As String, TmpEffDate As String, tmpExpDate As String)
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
    SelCriteria = " And INSCode ='" & Trim(tmpINSType) & "'"
    tmpTableName = "INS"
    tmpCode = ""
    tmpDesc = ""
    TmpRecFound = False

    Call SPHISTableRecVerify(SelFields, tmpTableName, SelCriteria, TmpRecFound, tmpCode, tmpDesc, tmpStatus)

    If TmpRecFound = False Then
        Call ErrorListing(tmpLine, tmpPOlNo, tmpName, "No Such " & tmpProdType & " Insured Type - '" & Trim(tmpINSType) & "' !")
    End If
    If IsDate(Trim(TmpEffDate)) = False Then
        Call ErrorListing(tmpLine, tmpPOlNo, tmpName, "Invalid " & tmpProdType & " Date format ~ Effective Date")
    End If
    If IsDate(Trim(tmpExpDate)) = False Then
        Call ErrorListing(tmpLine, tmpPOlNo, tmpName, "Invalid " & tmpProdType & " Date format ~ Expiry Date ")
    End If
    If IsDate(Trim(TmpEffDate)) And IsDate(Trim(tmpExpDate)) Then
        CoverDays = CDate(tmpExpDate) - CDate(TmpEffDate)
        If CoverDays < 30 Or CoverDays > 365 Then
            Call ErrorListing(tmpLine, tmpPOlNo, tmpName, tmpProdType & " Policy Coverage must be > 30 and < 365 days ")
        End If
    End If
End Sub

Private Sub SaveMBMLifeTimeLmt(LMBMNumber As String, LMBMCovID As String, LmtType As String, tmpPlnCode As String, tmpINSType As String, tmpEDate As Date, tmpVAnnLmt As Double, tmpCNextLmt As Double, tmpVGuaRenewal As String)
Dim tmpSql As String
Dim TmpLifeSql As String
Dim TmpLimit As String
Dim TmpLmtVar As String
Dim cmd As New ADODB.Command
Dim TmpLMBM As New ADODB.Recordset
Dim Prm1 As New ADODB.Parameter
Dim Prm2 As New ADODB.Parameter
Dim TmpSql2 As Integer
    
    TmpSql2 = 0
    If Trim(LMBMCovID) <> "00" Then
        TmpSql2 = 1
    End If
    tmpSql = "Select MBMNumber, MBMLifeTimeLmt From MBMLifeTime where MBMNumber='" & Trim(LMBMNumber) & _
              "' AND MBMCovID='" & LMBMCovID & "'"
    TmpLMBM.Open tmpSql, medixmasterconn, adOpenForwardOnly, adLockReadOnly
    If TmpLMBM.EOF Then
        If tmpVGuaRenewal = "3" Then
            WLifeStatus = True
        End If
    Else
        tmpCNextLmt = tmpVAnnLmt
        If CDbl(tmpVAnnLmt) > CDbl(TmpLMBM!MBMLifeTimeLmt) Then
            tmpVAnnLmt = CDbl(TmpLMBM!MBMLifeTimeLmt)
        End If
    End If
    TmpLMBM.Close
    Set TmpLMBM = Nothing
    If WLifeStatus = True Then
        TmpLmtVar = "SuppLifeTimeLimit"
        If LmtType = "P" Then
            TmpLmtVar = "LifeTimeLimit"
        End If
        TmpLimit = GetLifeTimeLimit(TmpLmtVar, tmpINSType, tmpPlnCode, tHLTCode, CDate(tmpEDate))
        TmpLimit = CDbl(TmpLimit) - (tmpVAnnLmt - tmpCNextLmt)
        TmpLifeSql = "Insert Into MBMLifeTime(MBMNumber, MBMCOVID, MBMLifeTimeLmt) " & _
              "Values('" & LMBMNumber & "','" & LMBMCovID & "'," & TmpLimit & ")"
        medixmasterconn.Execute TmpLifeSql
    End If
End Sub

Private Sub SaveMBMBnfAnnLmt(LMBMNumber, LMBMCovID, LmtType, tmpKidneyAnnLmt, tmpCancerAnnLmt, tmpHomeNursAnnLmt)
Dim TmpBnfLmtSql As String
Dim bnfAnn As New ADODB.Recordset
Dim sql As String
Dim RcdCnt As Boolean
    
    RcdCnt = True
    sql = ""
    sql = "Select MBMNumber, MBMCoveredID from MBMBnfLmt where MBMNumber = '" & LMBMNumber & "' and MBMCoveredID = '" & LMBMCovID & "'"
    bnfAnn.Open sql, medixmasterconn, adOpenForwardOnly, adLockReadOnly
        Do While Not bnfAnn.EOF
            RcdCnt = True
            bnfAnn.MoveNext
        Loop
    bnfAnn.Close
    Set bnfAnn = Nothing
    
    
    If RcdCnt = False Then
        TmpBnfLmtSql = "Insert Into MBMBnfLmt(MBMNumber, MBMCoveredID, MBMKidneyAnnLmt, MBMKidneyBalLmt, " & _
                     "MBMCancerAnnLmt, MBMCancerBalLmt, MBMHomeNCAnnLmt, MBMHomeNCBalLmt) " & _
                     "Values('" & LMBMNumber & "','" & LMBMCovID & "'," & tmpKidneyAnnLmt & "," & tmpKidneyAnnLmt & _
                     "," & tmpCancerAnnLmt & "," & tmpCancerAnnLmt & "," & tmpHomeNursAnnLmt & "," & tmpHomeNursAnnLmt & ")"
        medixmasterconn.Execute TmpBnfLmtSql
    End If
End Sub

Private Sub SaveMBMBnfLTLmt(LMBMNumber, LMBMCovID, tmpKidneyAnnLmt, tmpCancerAnnLmt)
Dim TmpBnfLmtSql As String
Dim StrSql As String
Dim rst As New ADODB.Recordset
Dim recfnd As Boolean

    recfnd = False
    StrSql = "select MBMNumber,MBMCoverID from MBMBnfLTLmt where MBMNumber = '" & LMBMNumber & "' and  MBMCoverID = '" & LMBMCovID & "'"
    rst.Open StrSql, medixmasterconn, adOpenForwardOnly, adLockBatchOptimistic
    
    Do While Not rst.EOF
        recfnd = True
    rst.MoveNext
    Loop
    rst.Close
    Set rst = Nothing
    
    If recfnd = False Then
        TmpBnfLmtSql = "Insert Into MBMBnfLTLmt(MBMNumber, MBMCoverID, MBMKidneyLTLmt, MBMCancerLTLmt) " & _
                     "Values('" & LMBMNumber & "','" & LMBMCovID & "'," & tmpKidneyLTLmt & "," & tmpCancerLTLmt & " )"
        medixmasterconn.Execute TmpBnfLmtSql
    End If

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

Private Sub LoadComboListing(tmpSelFields As String, tmpTableName As String, tmpSelCriteria As String, tmpComboBox As ComboBox)
Dim TmpRec As New ADODB.Recordset
Dim cmd As New ADODB.Command
Dim Prm1 As New ADODB.Parameter
Dim Prm2 As New ADODB.Parameter
Dim Prm3 As New ADODB.Parameter
    
    medixmasterconn.Open "File Name=" & BOSPath & "Conn\mediconnect.UDL"

    Set cmd.ActiveConnection = medixmasterconn
    cmd.CommandText = "SearchHISTableRec"
    cmd.CommandType = adCmdStoredProc
    cmd.CommandTimeout = 300
    Set Prm1 = cmd.CreateParameter("SelFields", adVarChar, adParamInput, 2000, tmpSelFields)
    cmd.Parameters.Append Prm1
    Set Prm2 = cmd.CreateParameter("TableName", adVarChar, adParamInput, 255, tmpTableName)
    cmd.Parameters.Append Prm2
    Set Prm3 = cmd.CreateParameter("SelCriteria", adVarChar, adParamInput, 2000, tmpSelCriteria)
    cmd.Parameters.Append Prm3

    TmpRec.CursorLocation = adUseClient
    TmpRec.CursorType = adOpenForwardOnly
    TmpRec.CacheSize = 100
    Set TmpRec = cmd.Execute
    tmpComboBox.Clear
    Do Until TmpRec.EOF
        With TmpRec
            tmpComboBox.AddItem !tmpCode & " - " & !tmpName
            TmpRec.MoveNext
        End With
    Loop
    
    TmpRec.Close
    Set TmpRec = Nothing
    medixmasterconn.Close
    Set medixmasterconn = Nothing
End Sub

Private Sub SPHISTableRecVerify(tmpSelFields As String, tmpTableName As String, tmpSelCriteria As String, TmpRecFound As Boolean, tmpCode As String, tmpDesc As String, tmpStatus As String, Optional tmpPLNPayor As String)
Dim TmpRec As New ADODB.Recordset
Dim cmd As New ADODB.Command
Dim Prm1 As New ADODB.Parameter
Dim Prm2 As New ADODB.Parameter
Dim Prm3 As New ADODB.Parameter

    TmpRecFound = False
    'If PaySearch <> "" Then
    '    medixmasterconn.Open "File Name=" & BOSPath & "Conn\mediconnect.UDL"
    'End If
    'edit by shan 13012011
    If medixmasterconn = "" Then
        medixmasterconn.Open "File Name=" & BOSPath & "Conn\mediconnect.UDL"
    End If
    
    Set cmd.ActiveConnection = medixmasterconn
    cmd.CommandText = "SearchHISTableRec"
    cmd.CommandType = adCmdStoredProc
    cmd.CommandTimeout = 300
    Set Prm1 = cmd.CreateParameter("SelFields", adVarChar, adParamInput, 2000, tmpSelFields)
    cmd.Parameters.Append Prm1
    Set Prm2 = cmd.CreateParameter("TableName", adVarChar, adParamInput, 255, tmpTableName)
    cmd.Parameters.Append Prm2
    Set Prm3 = cmd.CreateParameter("SelCriteria", adVarChar, adParamInput, 2000, tmpSelCriteria)
    cmd.Parameters.Append Prm3

    TmpRec.CursorLocation = adUseClient
    TmpRec.CursorType = adOpenForwardOnly
    TmpRec.CacheSize = 100
    Set TmpRec = cmd.Execute
    
    Do Until TmpRec.EOF
        With TmpRec
            tmpCode = IIf(Len(TmpRec!tmpCode), TmpRec!tmpCode, "")
            tmpDesc = IIf(Len(TmpRec!tmpDesc), TmpRec!tmpDesc, "")
            If Trim(tmpTableName) = "PAY" Then
                tmpStatus = IIf(Len(TmpRec!tmpStatus), TmpRec!tmpStatus, "")
            End If
            If Trim(tmpTableName) = "PLN" Then
                tmpPLNPayor = IIf(Len(TmpRec!tmpPLNPayor), TmpRec!tmpPLNPayor, "")
            End If
            TmpRecFound = True
            TmpRec.MoveNext
            Exit Do
        End With
    Loop
    
    TmpRec.Close
    Set TmpRec = Nothing
    'If PaySearch <> "" Then
    '    PaySearch = ""
    '    medixmasterconn.Close
    '    Set medixmasterconn = Nothing
    'End If
End Sub

Private Sub SPMaintTableRecVerify(tmpSelFields As String, tmpTableName As String, tmpSelCriteria As String, TmpRecFound As Boolean, tmpCode As String, tmpDesc As String)
Dim TmpREc1 As New ADODB.Recordset
Dim cmd As New ADODB.Command
Dim Prm1 As New ADODB.Parameter
Dim Prm2 As New ADODB.Parameter
Dim Prm3 As New ADODB.Parameter

    TmpRecFound = False
    'medixmasterconnMaint.Open "File Name=" & BOSPath & "Conn\MediConnHISMaint.UDL"

    Set cmd.ActiveConnection = medixmasterconnMaint
    cmd.CommandText = "SearchMaintTableRec"
    cmd.CommandType = adCmdStoredProc
    cmd.CommandTimeout = 300
    Set Prm1 = cmd.CreateParameter("SelFields", adVarChar, adParamInput, 2000, tmpSelFields)
    cmd.Parameters.Append Prm1
    Set Prm2 = cmd.CreateParameter("TableName", adVarChar, adParamInput, 255, tmpTableName)
    cmd.Parameters.Append Prm2
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
    'medixmasterconnMaint.Close
    'Set medixmasterconnMaint = Nothing
End Sub

Private Sub SPCISDBTableRecVerify(tmpSelFields As String, tmpTableName As String, tmpSelCriteria As String, TmpRecFound As Boolean, tmpCode As String, tmpDesc As String)
Dim TmpREc1 As New ADODB.Recordset
Dim cmd As New ADODB.Command
Dim Prm1 As New ADODB.Parameter
Dim Prm2 As New ADODB.Parameter
Dim Prm3 As New ADODB.Parameter

    TmpRecFound = False
    MediConnCISDB.Open "File Name=" & BOSPath & "Conn\MediConnCISDB.UDL"

    Set cmd.ActiveConnection = MediConnCISDB
    cmd.CommandText = "SearchCISDBTableRec"
    cmd.CommandType = adCmdStoredProc
    cmd.CommandTimeout = 300
    Set Prm1 = cmd.CreateParameter("SelFields", adVarChar, adParamInput, 2000, tmpSelFields)
    cmd.Parameters.Append Prm1
    Set Prm2 = cmd.CreateParameter("TableName", adVarChar, adParamInput, 255, tmpTableName)
    cmd.Parameters.Append Prm2
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
Dim Prm1 As New ADODB.Parameter
Dim Prm2 As New ADODB.Parameter
    
    'medixmasterconnMaint.Open "File Name=" & BOSPath & "Conn\MediConnHISMaint.UDL"
    Set cmd.ActiveConnection = medixmasterconnMaint
    cmd.CommandText = "SearchMBMCrossRef"
    cmd.CommandType = adCmdStoredProc
    cmd.CommandTimeout = 300
    Set Prm1 = cmd.CreateParameter("SelField", adVarChar, adParamInput, 2000, XrefSql1)
    cmd.Parameters.Append Prm1
    Set Prm2 = cmd.CreateParameter("SelCriteria", adVarChar, adParamInput, 255, XrefSql2)
    cmd.Parameters.Append Prm2
    XrefDB.CursorLocation = adUseClient
    XrefDB.CursorType = adOpenForwardOnly
    XrefDB.CacheSize = 100
    Set XrefDB = cmd.Execute
    Do Until XrefDB.EOF
        If tmpVSql = "1" Then
            Call ErrorListing(NoOfRecLine, VNewPolicyNo, vNewName, "Same IC " & Trim(XrefDB!MBMIcBcPp) & " Existed in Membership No. :" & XrefDB!MBMNumber)
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
    
    'medixmasterconnMaint.Close
    'Set medixmasterconnMaint = Nothing
End Sub

Private Sub NewMBMInfoProcessing(tmpVCovID As String)
'-----
Dim tmpVDataType As String
Dim tmpVPolicy As String
Dim tmpVMBMName As String
Dim tmpVIC As String
Dim tmpVDOB As String
Dim tmpVEffDate As String
Dim tmpVExpDate As String
Dim tmpVINSCode As String
Dim tmpVPlanCode As String
Dim tmpVPolDis As Double
Dim tmpVRenewDis As Double
Dim tmpVGuaRenewal As String
Dim tmpVClmNonPayFr As String
Dim tmpVClmNonPayTo As String

' ----------- cis
Dim tmpCisVPlnCode As String
Dim tmpCisVInsType As String
Dim tmpCisVExpDate As String
Dim tmpCisVEffDate As String
Dim tmpCisVCardType As String
Dim tmpCisVWarranty As String
Dim tmpCLNVStatus As String
Dim tmpVClinical As String

Dim TmpRecFound As Boolean
Dim SelFields As String
Dim SelCriteria As String
Dim tmpCode As String
Dim tmpDesc As String
Dim tmpTableName As String
Dim tmpVAnnLmt As Double
Dim tmpVLifeLmt As Double
Dim RecFound As Boolean
Dim tmpMBMDataType As String
Dim tmpVMemberType As String
Dim tmpVTopupInd As String
Dim tmpVHSClnInd As String
Dim tmpVPremium As Double
Dim tmpVBasicPrem As Double
Dim tmpVMcoAmt As Double
Dim tmpVBasicMco As Double
Dim tmpVSuppBNFStatus As String
Dim tmpVSex As String

Dim tHnCEffDate As String
Dim tHnCExpDate As String
Dim tHnCPlnCode As String
Dim tDOB As String
Dim tAgeCode As String
Dim tFindAge As String

Dim tmpVKidneyAnnLmt As Double
Dim tmpVCancerAnnLmt As Double
Dim tmpVHomeNursAnnLmt As Double
Dim tmpClnStatus As String
Dim tmpClnDAtaStatus As String
Dim tmpClnExpStatus As String
Dim CISBasicPrem As Double
Dim CisPrem As Double
Dim CisMCO As Double
Dim CisBasicMco As Double
Dim CisBalLmt  As Double
Dim CisVisitLmt As Double
Dim TmpCISMBMType As String
Dim CisAnnLmtType As String
Dim CisDentalLmt As Double
Dim CisSpecLmt As Double
Dim CisMepLmt As Double
Dim tmpSuppExpDate As String
Dim tmpSuppCisExpDate As String
Dim tmpVValueDate As String
Dim tmpVBatchNo As String
Dim tmpVExpStatus As String
Dim tmpVMBMStatus As String
Dim tmpVDataStatus As String
Dim tmpVBordxDate As String
Dim tmpSuppEffDate As String
Dim tmpVRenewal As String
Dim tmpCNextLmt As Double
Dim TmpLMBM As New ADODB.Recordset
Dim TmpChkDigits As String
Dim tmpPayorPLN As String

    tmpVTopupNo = ""
    PrevLifeMBMNo = ""
    TmpChkDigits = ""
    tmpPLNPremSexStatus = ""
    'tMBMNo = ""
    'tmpCode = ""
    
    tmpVValueDate = Format(Date, "yyyy/mm/dd")
    tmpVDataStatus = "A"
    tmpVExpStatus = "N"
    If tmpCancelstatus = False Then
        tmpVBordxDate = Format(txtSubmissionDate, "yyyy/mm/dd")
        tmpVMBMStatus = "A"
         tmpVBatchNo = txtBatchNo
    Else
        tmpVBordxDate = Format(txtSubmissionDate, "yyyy/mm/dd")
        tmpVMBMStatus = "C"
        tmpVBatchNo = "1"
    End If
       
    WLifeStatus = False
    
    tHLTCode = Mid(cboHLTCode1, 1, 1)
    
    If tmpFileStatus = "I" And Trim(Mid(TextLine, 1, 1)) = "R" Then
        tmpVDataType = "P"
    Else
        tmpVDataType = Mid(TextLine, 1, 1)
    End If
    
    If txtBordxFormat = "3" Then
        tmpVPolicy = Trim(Mid(TextLine, 2, 17))
    Else
        tmpVPolicy = Trim(Mid(TextLine, 2, 25))
    End If
    tmpVMBMName = Trim(Mid(TextLine, 47, 60))
    tmpVPolicy = Trim(Replace(tmpVPolicy, "'", "''"))
    tmpVMBMName = Trim(Replace(tmpVMBMName, "'", "''"))
    tmpVRenewal = Trim(Mid(TextLine, 435, 1))
    If tmpVRenewal = "C" Or tmpVRenewal = "Y" Or tmpVRenewal = "" Then
        tmpVRenewal = "N"
    End If
    If txtBordxFormat = "3" Then
        tmpPayorPLN = Trim(Mid(TextLine, 390, 7))
        SelFields = "SELECT PLNCode as tmpCode, PLNLifeTimeStatus as tmpDesc, PLNPayorCode as tmpPLNPayor "
        SelCriteria = " And PLNPayorPlan ='" & Trim(tmpPayorPLN) & "'"
        tmpTableName = "PLN"
        tmpCode = ""
        tmpDesc = ""
        TmpRecFound = False
        Call SPHISTableRecVerify(SelFields, tmpTableName, SelCriteria, TmpRecFound, tmpCode, tmpDesc, tmpStatus, tmpPLNPayor)
            If TmpRecFound = True Then
                tmpVPlanCode = tmpCode
                tPayor = tmpPLNPayor
            Else
                tmpVPlanCode = ""
                tmpPLNPayor = ""
            End If
            'tmpVEffDate = Mid(TextLine, 374, 2) & "/" & Mid(TextLine, 376, 2) & "/" & Mid(TextLine, 378, 4)
            'tmpVExpDate = Mid(TextLine, 382, 2) & "/" & Mid(TextLine, 384, 2) & "/" & Mid(TextLine, 386, 4)
            
    Else
        tmpVPlanCode = Trim(Mid(TextLine, 390, 4))
        tmpPLNPayor = ""
        tPayor = Mid(cboPayorCode, 1, 2)
    End If

    tmpVIC = Trim(Mid(TextLine, 273, 20))
    tmpVIC = Trim(tmpVIC)
    tmpVDOB = Format(Mid(TextLine, 293, 2) & "/" & Mid(TextLine, 295, 2) & "/" & Mid(TextLine, 297, 4), "yyyy/mm/dd")
    tDOB = Format(tmpVDOB, "YYYY/MM/DD")
    tmpVSex = Trim(Mid(TextLine, 301, 1))
    tmpVEffDate = Mid(TextLine, 374, 2) & "/" & Mid(TextLine, 376, 2) & "/" & Mid(TextLine, 378, 4)
    tmpVExpDate = Mid(TextLine, 382, 2) & "/" & Mid(TextLine, 384, 2) & "/" & Mid(TextLine, 386, 4)
        
        
    'tmpVPlanCode = Trim(Mid(TextLine, 390, 4))
    If txtBordxFormat = "3" Then
        tmpVINSCode = Trim(Mid(TextLine, 397, 1))
    Else
        tmpVINSCode = Trim(Mid(TextLine, 394, 1))
    End If
    'tmpVINSCode = Trim(Mid(TextLine, 397, 1))

    tmpVClmNonPayFr = Mid(TextLine, 3576, 2) & "/" & Mid(TextLine, 3578, 2) & "/" & Mid(TextLine, 3580, 4)
    tmpVClmNonPayFr = IIf(IsDate(tmpVClmNonPayFr), Format(tmpVClmNonPayFr, "YYYY/MM/DD"), "")
    tmpVClmNonPayTo = Mid(TextLine, 3584, 2) & "/" & Mid(TextLine, 3586, 2) & "/" & Mid(TextLine, 3588, 4)
    tmpVClmNonPayTo = IIf(IsDate(tmpVClmNonPayTo), Format(tmpVClmNonPayTo, "YYYY/MM/DD"), "")
    
    If txtBordxFormat = "3" Then
        tmpVPolDis = 0
        tmpVRenewDis = 0
        tmpVGuaRenewal = ""
    Else
        tmpVPolDis = IIf(IsNumeric(Trim(Mid(TextLine, 3732, 5))), Trim(Mid(TextLine, 3732, 5)), 0)
        tmpVRenewDis = IIf(IsNumeric(Trim(Mid(TextLine, 3737, 5))), Trim(Mid(TextLine, 3737, 5)), 0)
        tmpVGuaRenewal = Trim(Mid(TextLine, 3842, 5))
    End If
    tmpCisVPlnCode = Trim(Mid(TextLine, 6708, 4))
    tmpCisVInsType = Trim(Mid(TextLine, 6712, 1))
    tmpCisVEffDate = Mid(TextLine, 6713, 2) & "/" & Mid(TextLine, 6715, 2) & "/" & Mid(TextLine, 6717, 4)
    tmpCisVExpDate = Mid(TextLine, 6721, 2) & "/" & Mid(TextLine, 6723, 2) & "/" & Mid(TextLine, 6725, 4)
    tmpCisVWarranty = Trim(Mid(TextLine, 6729, 2500))
    tmpCisVCardType = Trim(Mid(TextLine, 9229, 20))

    tmpCisVWarranty = Trim(Replace(tmpCisVWarranty, "'", "''"))
    tmpCisVCardType = Trim(Replace(tmpCisVCardType, "'", "''"))
'-------

    If tmpVDataType = "P" Then
        tmpVPrinPlnCode = tmpVPlanCode
        tmpVPrinInsCode = tmpVINSCode
        tmpVPrinEffDate = tmpVEffDate
        tmpVPrinExpDate = tmpVExpDate
        tmpMBMDataType = "P"
        tmpVPrinGuaRenew = tmpVGuaRenewal
    Else
        tmpMBMDataType = "C"
    End If

' use principal's data when supplementary data is blank
    tmpVPlanCode = IIf(Trim(tmpVPlanCode) = "", tmpVPrinPlnCode, tmpVPlanCode)
    tmpVINSCode = IIf(Trim(tmpVINSCode) = "", tmpVPrinInsCode, tmpVINSCode)
    tmpVEffDate = IIf(IsDate(tmpVEffDate), tmpVEffDate, tmpVPrinEffDate)
    tmpVExpDate = IIf(IsDate(tmpVExpDate), tmpVExpDate, tmpVPrinExpDate)
    tmpSuppEffDate = IIf(IsDate(tmpVEffDate), Format(tmpVEffDate, "yyyy/mm/dd"), "")
    tHnCPlnCode = tmpVPlanCode
    tHncInsType = tmpVINSCode
    tHnCEffDate = tmpVEffDate
    tHnCExpDate = tmpVExpDate
    
    If Trim(tHnCPlnCode) = "" Then
        tHnCPlnCode = tmpCisVPlnCode
        tHncInsType = tmpCisVInsType
        tHnCEffDate = tmpCisVEffDate
        tHnCExpDate = tmpCisVExpDate
    End If

    tCAnnSuppLmtStatus = "A"
    
    If Trim(tmpVPlanCode) <> "" Then
        Call SPHISDBTableRecAssign("PLN", tmpVPlanCode, tmpVINSCode, CDate(tmpVEffDate))
        tAnnSuppLmtStatus = "A"
        If Trim(tmpPLNGuaRenewal) <> "" Then
            tmpVGuaRenewal = tmpPLNGuaRenewal
        Else
            tmpVGuaRenewal = IIf(Trim(tmpVGuaRenewal) = "", tmpVPrinGuaRenew, tmpVGuaRenewal)
        End If
        
        If tPLNAnnLmtIND = "Y" Then
            Call SPHISDBTableRecAssign("AnnualLimit", tmpVPlanCode, tmpVINSCode, CDate(tmpVEffDate))
        End If
    End If
    If tmpCisVPlnCode <> "" Then
        Call SPCisDBTableRecAssign("PLN", tmpCisVPlnCode, tmpCisVInsType, CDate(tmpCisVEffDate))
        If tCPLNAnnLmtInd = "Y" Then
            Call SPCisDBTableRecAssign("AnnualLimit", tmpCisVPlnCode, tmpCisVInsType, CDate(tmpCisVEffDate))
        End If
    End If
    tmpVMemberType = "P"
    If Trim(tAnnSuppLmtStatus) <> "A" Then
        tmpVMemberType = "Q"
    End If
    tFindAge = GetAge(CDate(tHnCEffDate), CDate(tmpVDOB))
    tAgeCode = Mid(GetAgeBand(tHLTCode, tFindAge, tHnCPlnCode, tHncInsType), 1, 2)
    tAdultChild = "A"
    If tAgeCode = "01" Then
        tAdultChild = "C"
    End If

'------------------------------------------------------------------------------------------------------------------------
    tmpVTopupInd = ""
    
    If Trim(tmpVPlanCode) <> "" Then 'only for HIS prem, mco, annLmt, Access Calc.
        tmpVBasicPrem = 0
        tmpVPremium = 0
        If tPLNPrmIND = "Y" Then
            tmpVPremium = Format(GetHISPremium(tmpVINSCode, tPayor, tAgeCode, tmpVPlanCode, tHLTCode, CDate(tmpVEffDate), tmpMBMDataType, tmpVCovID, tmpPLNPremSexStatus, tmpVSex), "#,##0.00")
            tmpVBasicPrem = tmpVPremium
            If tmpVPolDis > 0 Then
                tmpVPremium = tmpVPremium * (100 - tmpVPolDis) / 100
            End If
            If tmpVRenewDis > 0 Then
                tmpVPremium = tmpVPremium * (100 - tmpVRenewDis) / 100
            End If
            
            tmpVPremium = InsertPremium(CDate(tmpVExpDate), CDate(tmpVEffDate), Format(tmpVPremium, "#,##0.00"))
            tmpVPremium = Format(tmpVPremium, "#,##0.00")
            
        End If
        tmpVBasicMco = 0
        tmpVMcoAmt = 0
        If tPLNMcoIND = "Y" Then
            tmpVMcoAmt = Format(GetHISMCO(tmpVINSCode, tHLTCode, tAdultChild, CDate(tmpVEffDate), tmpVPlanCode, tmpMBMDataType, tmpVCovID), "#,##0.00")
            tmpVBasicMco = tmpVMcoAmt
            If tmpProRate <> "N" Then
                tmpVMcoAmt = InsertMCOFees(CDate(tmpVExpDate), CDate(tmpVEffDate), Format(tmpVBasicMco, "#,##0.00"))
            End If
            tmpVMcoAmt = Format(tmpVMcoAmt, "#,##0.00")
            If TmpRoundUpStatus <> "N" Then
                tmpVMcoAmt = Round(tmpVMcoAmt)
            End If
        End If
        tmpVAnnLmt = 0
        If tPLNAnnLmtIND = "Y" Then
            tmpVAnnLmt = 0
            tmpVAnnLmt = GetHISAnnLmt(tmpVINSCode, tmpVPlanCode, tHLTCode, CDate(tmpVEffDate), tmpMBMDataType, tmpVCovID)
            tmpVAnnLmt = CDbl(tmpVAnnLmt)
            tmpCNextLmt = tmpVAnnLmt
        End If
        tmpVLifeLmt = 0
        
        If Trim(tmpPLNCancerLTInd) = "Y" Then
            Call GetHISLTKidneyCancer(tmpVINSCode, tmpVPlanCode, tHLTCode, CDate(tmpVEffDate), tmpMBMDataType, tmpVCovID)
        End If
        
' ---------- write Previous MBM no. to new MBMTopupMember field (only for HIS)
        If Trim(tPLNTopupStatus) = "Y" Then
            SelFields = "SELECT TOP 1 MBMNumber as tmpCode, MBMTopupNumber as tmpDesc "
            SelCriteria = " AND MBMICBCPP = '" & Trim(tmpVIC) & _
                        "' And MBMStatus = 'A' AND PayCode ='" & tPayor & "' AND INSCode='" & Trim(tmpVINSCode) & _
                        "' ORDER BY MBMPayorEffDate DESC, MBMNumber "
            tmpTableName = "MBMCrossReference"
            tmpCode = ""
            tmpDesc = ""
            TmpRecFound = False
            Call SPMaintTableRecVerify(SelFields, tmpTableName, SelCriteria, TmpRecFound, tmpCode, tmpDesc)
            If TmpRecFound = True Then
                tmpVTopupNo = tmpCode
                tmpVTopupInd = "Y"
            End If
        End If
'-------- end of Topup
    End If   'if trim(tmpVPlanCode)<> ""
    If tmpVDataType = "P" Then
        If Trim(tmpVPlanCode) = "" And Trim(tmpCisVPlnCode) <> "" Then
            'tMBMNo = tPayor & tmpCisVPlnCode & tmpCisVInsType & GenerateMembershipNo(tPayor, tmpCisVPlnCode, tmpCisVInsType)
            tMBMNo = tPayor & tmpCisVInsType & GenerateMembershipNo(tPayor, tmpCisVInsType)
            TmpChkDigits = GenerateChkDigits
            tMBMNo = tMBMNo & TmpChkDigits
        ElseIf Trim(tmpVPlanCode) <> "" And Trim(tmpCisVPlnCode) <> "" Then
            If txtPayorMBMNoType = "P" Then
                tMBMNo = tPayor & GenerateMBMNoPol(tPayor, tmpVINSCode, tmpVPolicy, tmpVRenewal, tmpFileStatus)
            Else
                tMBMNo = tPayor & GenerateMBMNo05(tPayor, tmpVINSCode, tDOB, tmpVMBMName, tmpVRenewal)
            End If
            'TmpChkDigits = GenerateChkDigits
            tMBMNo = tMBMNo & TmpChkDigits
        Else
            If txtPayorMBMNoType = "P" And tPayor = "XL" Then
                tMBMNo = tPayor & GenerateMBMNoPol(tPayor, tmpVINSCode, tmpVPolicy, tmpVRenewal, tmpFileStatus)
            Else
                tMBMNo = tPayor & GenerateMBMNo05(tPayor, tmpVINSCode, tDOB, tmpVMBMName, tmpVRenewal)
            End If
            
            tMBMNo = tMBMNo & TmpChkDigits
        End If
    End If
    tmpVHSClnInd = "H"
    If Trim(tmpVPlanCode) = "" And Trim(tmpCisVPlnCode) <> "" Then
        tmpVHSClnInd = "C"
    End If
    If Trim(tmpVPlanCode) <> "" And Trim(tmpCisVPlnCode) = "" Then
        tmpVHSClnInd = "H"
    End If
    If Trim(tmpCisVPlnCode) <> "" And Trim(tmpVPlanCode) <> "" Then
        If Trim(tCPlnDCStatus) = "Y" Then
            tmpVHSClnInd = "M"
        Else
            tmpVHSClnInd = "B"
        End If
    End If
    tmpVClinical = "N"
    If Trim(tmpCisVPlnCode) <> "" Then
         tmpVClinical = "Y"
    End If

' ----- LifeTimeLimit & BenefitLimit
    If Trim(tmpVGuaRenewal) <> "" Then
        If Trim(tPLNLifeTimeStatus) <> "Y" Then
            tmpVTopupNo = ""
        Else
            tmpVKidneyAnnLmt = 0
            tmpVCancerAnnLmt = 0
            tmpVHomeNursAnnLmt = 0
            If Trim(tmpPLNCancerLTInd) <> "Y" Then
                Call GetBenefitLimit(tmpVINSCode, tmpVPlanCode, tHLTCode, CDate(tmpVEffDate), tmpVKidneyAnnLmt, tmpVCancerAnnLmt, tmpVHomeNursAnnLmt, tmpVSuppBNFStatus)
                If (tmpVSuppBNFStatus = "A" And tmpVCovID = "00") Or _
                   (tmpVSuppBNFStatus = "B" And tmpVCovID = "01") Or _
                   tmpVSuppBNFStatus = "C" Then
                    Call SaveMBMBnfAnnLmt(tMBMNo, tmpVCovID, tmpVSuppBNFStatus, tmpVKidneyAnnLmt, tmpVCancerAnnLmt, tmpVHomeNursAnnLmt)
                End If
            End If
            tmpCNextLmt = tmpVAnnLmt
            If tmpVDataType = "P" Then
            '    If Trim(tPayor) = "XG" Then
            '        SelCriteria = "Select top 1 MBMNumber, MBMTopupNumber, MBMAvailableLimit From MBM where Substring(MBMPolicyNo,1,12)='" & Mid(tmpVPolicy, 1, 12) & _
            '              "' And MBMStatus = 'A' AND PayCode ='" & Trim(tPayor) & "' AND INSCode='" & Trim(tmpVINSCode) & _
            '              "' ORDER BY MBMPayorEffDate DESC "
            '        TmpLMBM.Open SelCriteria, medixmasterconn, adOpenForwardOnly, adLockReadOnly
            '    Else
                    SelCriteria = "Select top 1 MBMNumber, MBMTopupNumber, MBMAvailableLimit From MBM where MBMPolicyNo='" & Trim(tmpVPolicy) & _
                          "' And MBMStatus = 'A' AND PayCode ='" & Trim(tPayor) & "' AND INSCode='" & Trim(tmpVINSCode) & _
                          "' ORDER BY MBMPayorEffDate DESC "
                    TmpLMBM.Open SelCriteria, medixmasterconn, adOpenForwardOnly, adLockReadOnly
            '    End If
                tmpVTopupNo = tMBMNo
                PrevLifeMBMNo = tMBMNo
                Do While Not TmpLMBM.EOF
                    tmpVTopupNo = TmpLMBM!MBMTopupNumber
                    PrevLifeMBMNo = TmpLMBM!MBMNumber
                    tmpCNextLmt = IIf(IsNumeric(TmpLMBM!MBMAvailableLimit), TmpLMBM!MBMAvailableLimit, 0)
                    TmpLMBM.MoveNext
                Loop
                TmpLMBM.Close
                Set TmpLMBM = Nothing
            Else
                SelFields = "SELECT TOP 1 MBMCNumber as tmpCode, MBMCAvailableLimit as tmpDesc "
                SelCriteria = " AND MBMCNumber='" & Trim(PrevLifeMBMNo) & _
                          "' And MBMCName = '" & tmpVMBMName & "'"
                tmpTableName = "MBMCoveredPersons"
                tmpCode = ""
                tmpDesc = ""
                TmpRecFound = False
                Call SPSearchMBM(SelFields, tmpTableName, SelCriteria, TmpRecFound, tmpCode, tmpDesc)
                If TmpRecFound = True Then
                    tmpCNextLmt = IIf(IsNumeric(tmpDesc), tmpDesc, 0)
                End If
            End If
            If Trim(tmpVGuaRenewal) <> "3" Then
                tmpVAnnLmt = tmpCNextLmt
            End If
                
            SelFields = "Select MBMNumber as tmpCode, MBMLifeTimeLmt as tmpDesc "
            SelCriteria = " AND MBMNumber='" & Trim(tmpVTopupNo) & _
                      "' AND MBMCovID='" & Trim(tmpVCovID) & "'"
            tmpTableName = "MBMLifeTime"
            tmpCode = ""
            tmpDesc = ""
            TmpRecFound = False
            Call SPSearchMBM(SelFields, tmpTableName, SelCriteria, TmpRecFound, tmpCode, tmpDesc)
            If TmpRecFound = False Then
                If Trim(tmpVGuaRenewal) = "3" Then
                    WLifeStatus = True
                End If
            End If
            If WLifeStatus = True Then
                If (tAnnSuppLmtStatus = "A" And tmpVCovID = "00") Or (tAnnSuppLmtStatus = "B" And tmpVCovID = "01") Or tAnnSuppLmtStatus = "C" Then
                    Call SaveMBMLifeTimeLmt(tmpVTopupNo, tmpVCovID, tmpVDataType, tmpVPlanCode, tmpVINSCode, CDate(tmpVEffDate), tmpVAnnLmt, tmpCNextLmt, tmpVGuaRenewal)
                End If
            End If
            
            If Trim(tmpPLNCancerLTInd) = "Y" Then
                SelFields = ""
                SelCriteria = ""
                
                SelFields = "Select MBMNumber as tmpCode, MBMCancerLTLmt as tmpDesc "
                SelCriteria = " AND MBMNumber='" & Trim(tmpVTopupNo) & _
                          "' AND MBMCoverID='" & Trim(tmpVCovID) & "'"
                tmpTableName = "MBMBnfLTLmt"
                tmpCode = ""
                tmpDesc = ""
                TmpRecFound = False
                
                Call SPSearchMBM(SelFields, tmpTableName, SelCriteria, TmpRecFound, tmpCode, tmpDesc)
                If TmpRecFound = True Then
                    Call SaveMBMBnfLTLmt(tmpVTopupNo, tmpVCovID, tmpKidneyLTLmt, tmpCancerLTLmt)
                End If
            End If

        End If
    End If
' ----- end of LifeTimeLimit & BenefitLimit
' ---- CLINICAL
    If Trim(tmpCisVPlnCode) <> "" Then
          
        tmpClnStatus = "A"
        tmpClnDAtaStatus = "A"
        tmpClnExpStatus = "N"
        CISBasicPrem = 0
        CisPrem = 0
        If tCPLNPRMInd = "Y" Then
              CisPrem = Format(GetCISPremium(tmpCisVInsType, tCPayor, tAgeCode, tmpCisVPlnCode, tHLTCode, CDate(tmpCisVExpDate), tmpMBMDataType, tmpVCovID), "#,##0.00")
              CISBasicPrem = CisPrem
              CisPrem = InsertPremium(CDate(tmpCisVExpDate), CDate(tmpCisVEffDate), Format(CisPrem, "#,##0.00"))
              CisPrem = Format(Round(CisPrem, 0), "#,##0.00")
        End If
        CisMCO = 0
        CisBasicMco = 0
        If tCPLNMCOInd = "Y" Then
            CisMCO = Format(GetCisMCO(tmpCisVInsType, tHLTCode, tAdultChild, CDate(tmpCisVEffDate), tmpCisVPlnCode, tmpMBMDataType, tmpVCovID), "#,##0.00")
            CisBasicMco = CisMCO
            'CisMCO = InsertMCOFees(CDate(tmpCisVExpDate), CDate(tmpCisVEffDate), Format(CisMCO, "#,##0.00"))
            'CisMCO = Round(CisMCO)
            If tCProRate <> "N" Then
                CisMCO = InsertMCOFees(CDate(tmpCisVExpDate), CDate(tmpCisVEffDate), Format(CisMCO, "#,##0.00"))
            End If
            CisMCO = Format(CisMCO, "#,##0.00")
            If tCRoundUpStatus <> "N" Then
                CisMCO = Round(CisMCO)
            End If
        End If
        
        CisBalLmt = 0
        If tCPLNAnnLmtInd = "Y" Then
            CisBalLmt = GetCISAnnLmt(tmpCisVInsType, tmpCisVPlnCode, tHLTCode, CDate(tmpCisVEffDate), tmpMBMDataType, tmpVCovID)
        End If
        CisVisitLmt = 0
        If tCPLNVsTLmtInd = "Y" Then
            CisVisitLmt = GetVisitLimit(tmpCisVInsType, tmpCisVPlnCode, tHLTCode, CDate(tmpCisVEffDate))
        End If
        If Trim(tCAnnSuppLmtStatus) = "A" Then
            TmpCISMBMType = "P"
        Else
            TmpCISMBMType = "Q"
        End If
        
        CisDentalLmt = 0
        CisSpecLmt = 0
        CisMepLmt = 0
        
        If Trim(tCPlnPDentStatus) = "S" Then 'stand alone dental limit
            CisAnnLmtType = "PDentalAnnLimit"
            CisDentalLmt = GetCISLimit(tmpCisVInsType, tmpCisVPlnCode, tHLTCode, CDate(tmpCisVEffDate), CisAnnLmtType)
        End If
        If Trim(tCPlnPSpecStatus) = "S" Then 'stand alone Spec limit
            CisAnnLmtType = "PSpecAnnLimit"
            CisSpecLmt = GetCISLimit(tmpCisVInsType, tmpCisVPlnCode, tHLTCode, CDate(tmpCisVEffDate), CisAnnLmtType)
        End If
        If Trim(tCPlnPMepStatus) = "S" Then 'stand alone MEP limit
            CisAnnLmtType = "PMepAnnLimit"
            CisMepLmt = GetCISLimit(tmpCisVInsType, tmpCisVPlnCode, tHLTCode, CDate(tmpCisVEffDate), CisAnnLmtType)
        End If
        
    End If
    tmpVPremium = CDbl(IIf(IsNumeric(tmpVPremium) = True, tmpVPremium, 0))
    tmpVBasicPrem = CDbl(IIf(IsNumeric(tmpVBasicPrem) = True, tmpVBasicPrem, 0))
    tmpVBasicMco = CDbl(IIf(IsNumeric(tmpVBasicMco) = True, tmpVBasicMco, 0))
    tmpVMcoAmt = CDbl(IIf(IsNumeric(tmpVMcoAmt) = True, tmpVMcoAmt, 0))
    tmpVAnnLmt = CDbl(IIf(IsNumeric(tmpVAnnLmt) = True, tmpVAnnLmt, 0))

    CisPrem = CDbl(IIf(IsNumeric(CisPrem) = True, CisPrem, 0))
    CISBasicPrem = CDbl(IIf(IsNumeric(CISBasicPrem) = True, CISBasicPrem, 0))
    CisBasicMco = CDbl(IIf(IsNumeric(CisBasicMco) = True, CisBasicMco, 0))
    CisMCO = CDbl(IIf(IsNumeric(CisMCO) = True, CisMCO, 0))
    CisBalLmt = CDbl(IIf(IsNumeric(CisBalLmt) = True, CisBalLmt, 0))
    CisVisitLmt = CDbl(IIf(IsNumeric(CisVisitLmt) = True, CisVisitLmt, 0))
    CisDentalLmt = CDbl(IIf(IsNumeric(CisDentalLmt) = True, CisDentalLmt, 0))
    CisSpecLmt = CDbl(IIf(IsNumeric(CisSpecLmt) = True, CisSpecLmt, 0))
    CisMepLmt = CDbl(IIf(IsNumeric(CisMepLmt) = True, CisMepLmt, 0))
    'tmpVMBMStatus = IIf(Trim(tmpVPlanCode) <> "", tmpVMBMStatus, "")
    tmpCisVWarranty = Replace(tmpCisVWarranty, "'", "''")

' ---- END OF CLINICAL
    If tmpVDataType = "P" Then
        Call SaveMBMPrincipal(tmpVMBMName, tDOB, tmpVPolicy, tmpVCovID, tmpVMBMStatus, tmpVDataStatus, tmpVExpStatus, tmpVIC, _
             tmpVMemberType, tmpVSex, tmpVPlanCode, tmpVINSCode, tmpVEffDate, tmpVExpDate, _
             tAgeCode, tAdultChild, tmpVAnnLmt, tmpVHSClnInd, tmpVTopupNo, tmpVTopupInd, tmpVBordxDate, tmpVBatchNo) ' done
        Call SaveMBMTwo(tmpVClinical)  'DONE
        Call SaveMBMXRef(tMBMNo, "00", Trim(tmpVPolicy), tmpVIC, tmpVMBMName, tmpVEffDate, _
             tmpVExpDate, tAgeCode, tmpVINSCode, tmpVPlanCode, tmpVBordxDate, tmpVTopupNo, _
             tmpVTopupInd, tmpVDataStatus, tmpVMBMStatus, tmpVDOB, tHLTCode, tPayor, tmpVBatchNo)  'done
        Call SaveMBMOthers(tmpVDataType, tmpVPremium, tmpVBasicPrem, tmpVBasicMco, tmpVMcoAmt, tmpVPolDis, tmpVRenewDis, tmpVClmNonPayFr, tmpVClmNonPayTo, tmpVEffDate, tmpVGuaRenewal) 'done
        Call SaveMBMOthersTwo
        If Trim(tmpCisVPlnCode) <> "" Then
            Call SaveMBMCLNOthers(TmpCISMBMType, tCPayor, tmpVPolicy, tmpClnStatus, _
                 tmpClnDAtaStatus, tmpClnExpStatus, tmpCisVPlnCode, tmpCisVInsType, tmpCisVWarranty, _
                 CISBasicPrem, CisPrem, CisBasicMco, CisMCO, CisBalLmt, CisVisitLmt, _
                 CisDentalLmt, CisSpecLmt, CisMepLmt, tmpCisVEffDate, tmpCisVExpDate, tmpCisVCardType) 'done
        End If
    ElseIf tmpVDataType <> "R" Then
        tmpSuppExpDate = ""
        If Trim(tmpVPlanCode) <> "" Then
            tmpSuppExpDate = Format(GetSuppExpDate(CDate(tmpVDOB), tmpVExpDate, tPLNAgeLimit), "YYYY/MM/DD")
        End If
        Call SaveMBMCovPerson(tmpVDataType, tmpVCovID, tDOB, tmpMBMDataType, tmpVMBMName, _
             tmpVMBMStatus, tmpVBordxDate, tmpVBatchNo, tmpVIC, tmpVSex, _
             tmpVPlanCode, tAgeCode, tAdultChild, tmpVAnnLmt, _
             tmpVBasicMco, tmpVMcoAmt, tmpVBasicPrem, tmpVPremium, tmpSuppEffDate, tmpSuppExpDate) 'done
        Call SaveMBMCovPersonTwo(tmpVCovID, tmpVClmNonPayFr, tmpVClmNonPayTo, tmpVPolDis, tmpVRenewDis) 'done
        If Trim(tmpCisVPlnCode) <> "" Then
            tmpSuppExpDate = ""
            tmpSuppCisExpDate = Format(GetSuppExpDate(CDate(tmpVDOB), tmpCisVExpDate, tCPLNAgeLimit), "YYYY/MM/DD")
            Call SaveMBMCovPersonCLN(TmpCISMBMType, tmpVCovID, tmpClnStatus, _
                 tmpCisVPlnCode, tmpCisVWarranty, CISBasicPrem, CisPrem, _
                 CisBasicMco, CisMCO, CisBalLmt, CisVisitLmt, _
                 CisDentalLmt, CisSpecLmt, CisMepLmt, tmpSuppCisExpDate, _
                 tmpCisVWarranty, tmpCisVCardType, tmpCisVPlnCode)
        End If
    End If
    
    If intExit = 1 Then
    '    Check = False
    End If

End Sub


Private Sub SPHISDBTableRecAssign(tmpTableName As String, tmpPlnCode As String, tmpINSType As String, tmpEDate As Date)
Dim TmpRec As New ADODB.Recordset
Dim cmd As New ADODB.Command
Dim Prm1 As New ADODB.Parameter
Dim Prm2 As New ADODB.Parameter
Dim Prm3 As New ADODB.Parameter
Dim tmpSelField As String
Dim tmpSelCriteria As String
    If tmpTableName = "PLN" Then
        TmpRoundUpStatus = "Y"
        tmpProRate = "Y"
        tmpSelField = "Select TOP 1 PLNTopUpStatus, PLNProRate, PRMInd, MCOInd, AnnLmtInd, PLNLifeTimeStatus, PLNAgeLimit,PLNRoundUpStatus,PLNPremGenderStatus,PLNGuaranteeStatus,PLNCancerInd "
        
        tmpSelCriteria = " AND PLNCode = '" & Trim(tmpPlnCode) & _
                         "' AND HLTCode='" & tHLTCode & _
                         "' order by PLNEffDate desc "
'                         "' AND PLNEffDate <='" & Format(TmpEDate, "mm/dd/yyyy") & _

        If tHLTCode = "N" Then
            tmpSelCriteria = " AND PLNCode = '" & Trim(tmpPlnCode) & _
                             "' AND HLTCode ='" & tHLTCode & _
                             "' AND PAYCode ='" & Trim(tPayor) & _
                             "' order by PLNEffDate desc "
'                             "' AND PLNEffDate <='" & Format(TmpEDate, "mm/dd/yyyy") & _

        End If
    ElseIf tmpTableName = "AnnualLimit" Then
        tmpSelField = "Select TOP 1 SuppLimitStatus "
    
        tmpSelCriteria = " AND PLNCode = '" & Trim(tmpPlnCode) & "' AND INSCode='" & Trim(tmpINSType) & _
                         "' AND HLTCode='" & tHLTCode & _
                         "' AND AnnualEffDate <='" & Format(tmpEDate, "mm/dd/yyyy") & _
                         "' order by AnnualEffDate desc "

    End If
    TmpRec.MaxRecords = 1
    Set cmd.ActiveConnection = medixmasterconn
    cmd.CommandText = "SearchHISTableRec"
    cmd.CommandType = adCmdStoredProc
    cmd.CommandTimeout = 300
    Set Prm1 = cmd.CreateParameter("SelFields", adVarChar, adParamInput, 2000, tmpSelField)
    cmd.Parameters.Append Prm1
    Set Prm2 = cmd.CreateParameter("TableName", adVarChar, adParamInput, 255, tmpTableName)
    cmd.Parameters.Append Prm2
    Set Prm3 = cmd.CreateParameter("SelCriteria", adVarChar, adParamInput, 2000, tmpSelCriteria)
    cmd.Parameters.Append Prm3

    TmpRec.CursorLocation = adUseClient
    TmpRec.CursorType = adOpenForwardOnly
    TmpRec.CacheSize = 100
    Set TmpRec = cmd.Execute
    
    If TmpRec.EOF Then
        MsgBox tmpTableName & " Information Is Not Correct", vbOKOnly + vbCritical, DialogBoxTitle
    Else
        Do Until TmpRec.EOF
            With TmpRec
                If tmpTableName = "PLN" Then
                    tPLNTopupStatus = IIf(Len(Trim(TmpRec!PLNTopUpStatus)), Trim(TmpRec!PLNTopUpStatus), "N")
                    tPLNPrmIND = IIf(Len(Trim(TmpRec!PRMInd)), Trim(TmpRec!PRMInd), "N")
                    tPLNMcoIND = IIf(Len(Trim(TmpRec!MCOInd)), Trim(TmpRec!MCOInd), "N")
                    tPLNAnnLmtIND = IIf(Len(Trim(TmpRec!ANNLmtInd)), Trim(TmpRec!ANNLmtInd), "N")
                    tPLNAgeLimit = IIf(Len(Trim(TmpRec!PlnAgeLimit)), Trim(TmpRec!PlnAgeLimit), "")
                    tPLNLifeTimeStatus = IIf(Len(Trim(TmpRec!PLNLifeTimeStatus)), Trim(TmpRec!PLNLifeTimeStatus), "N")
                    TmpRoundUpStatus = IIf(Len(Trim(TmpRec!PLNRoundUpStatus)), Trim(TmpRec!PLNRoundUpStatus), "Y")
                    tmpProRate = IIf(Len(Trim(TmpRec!PLNProRate)), Trim(TmpRec!PLNProRate), "Y")
                    tmpPLNPremSexStatus = IIf(Len(Trim(TmpRec!PLNPremGenderStatus)), Trim(TmpRec!PLNPremGenderStatus), "N")
                    tmpPLNGuaRenewal = IIf(Len(Trim(TmpRec!PLNGuaranteeStatus)), TmpRec!PLNGuaranteeStatus, "")
                    tmpPLNCancerLTInd = IIf(Len(Trim(TmpRec!PLNCancerInd)), TmpRec!PLNCancerInd, "")
                ElseIf tmpTableName = "AnnualLimit" Then
                    tAnnSuppLmtStatus = IIf(Len(Trim(TmpRec!supplimitstatus)), TmpRec!supplimitstatus, "A")
                End If
                TmpRec.MoveNext
            End With
        Loop
    End If
    TmpRec.Close
    Set TmpRec = Nothing
End Sub

Private Sub SPCisDBTableRecAssign(tmpTableName As String, tmpPlnCode As String, tmpINSType As String, tmpEDate As Date)
Dim TmpRec As New ADODB.Recordset
Dim cmd As New ADODB.Command
Dim Prm1 As New ADODB.Parameter
Dim Prm2 As New ADODB.Parameter
Dim Prm3 As New ADODB.Parameter
Dim tmpSelField As String
Dim tmpSelCriteria As String
    
    If tmpTableName = "PLN" Then
        tCRoundUpStatus = "Y"
        tCProRate = "Y"
        tmpSelField = "Select TOP 1 DCStatus, PRMInd, MCOInd, AnnLmtInd, VISlmtInd, " & _
                      "PAYCode, PDentalStatus, SDentalStatus, PSpecStatus, SSpecStatus, " & _
                      "PMepStatus, SMepStatus, PDentalLimit, SDentalLimit, PSpecLimit, " & _
                      "SSpecLimit, AnnLmtIndSupp, VISLmtIndSupp, PlnAgeLimit, PLNRoundUpStatus, " & _
                      "MCOProcfeeStatus, PLNProRate "
        tmpSelCriteria = " AND PLNCode = '" & Trim(tmpPlnCode) & "' order by PLNEffDate desc "
'                         "' AND PLNEffDate <='" & Format(TmpEDate, "mm/dd/yyyy") & _

    ElseIf tmpTableName = "AnnualLimit" Then
        tmpSelField = "Select TOP 1 SuppLimitStatus "
        tmpSelCriteria = " AND PLNCode = '" & Trim(tmpPlnCode) & "' AND INSCode='" & Trim(tmpINSType) & _
                         "' AND HLTCode='" & tHLTCode & _
                         "' AND AnnEffDate <='" & Format(tmpEDate, "mm/dd/yyyy") & _
                         "' order by AnnEffDate desc "

    End If
    Set cmd.ActiveConnection = MediConnCISDB
    cmd.CommandText = "SearchCISDBTableRec"
    cmd.CommandType = adCmdStoredProc
    cmd.CommandTimeout = 300
    Set Prm1 = cmd.CreateParameter("SelFields", adVarChar, adParamInput, 2000, tmpSelField)
    cmd.Parameters.Append Prm1
    Set Prm2 = cmd.CreateParameter("TableName", adVarChar, adParamInput, 255, tmpTableName)
    cmd.Parameters.Append Prm2
    Set Prm3 = cmd.CreateParameter("SelCriteria", adVarChar, adParamInput, 2000, tmpSelCriteria)
    cmd.Parameters.Append Prm3

    TmpRec.CursorLocation = adUseClient
    TmpRec.CursorType = adOpenForwardOnly
    TmpRec.CacheSize = 100
    Set TmpRec = cmd.Execute
    
    If TmpRec.EOF Then
        MsgBox tmpTableName & " (Clinical) Information Is Not Correct", vbOKOnly + vbCritical, DialogBoxTitle
    Else
        Do Until TmpRec.EOF
            With TmpRec
                If tmpTableName = "PLN" Then
                    tCPLNPRMInd = IIf(Len(Trim(TmpRec!PRMInd)), Trim(TmpRec!PRMInd), "N")
                    tCPLNMCOInd = IIf(Len(Trim(TmpRec!MCOInd)), Trim(TmpRec!MCOInd), "N")
                    tCPLNAnnLmtInd = IIf(Len(Trim(TmpRec!ANNLmtInd)), Trim(TmpRec!ANNLmtInd), "N")
                    tCPLNVsTLmtInd = IIf(Len(Trim(TmpRec!VISLmtInd)), Trim(TmpRec!VISLmtInd), "N")
                    tCPLNAgeLimit = IIf(Len(Trim(TmpRec!PlnAgeLimit)), Trim(TmpRec!PlnAgeLimit), "")
                    tCPlnDCStatus = IIf(Len(Trim(TmpRec!DCStatus)), Trim(TmpRec!DCStatus), "")
                    tCPayor = IIf(Len(Trim(TmpRec!PAYCode)), Trim(TmpRec!PAYCode), tPayor)
                    tCPlnPDentStatus = IIf(Len(Trim(TmpRec!PDentalStatus)), Trim(TmpRec!PDentalStatus), "")
                    tCPlnSDentStatus = IIf(Len(Trim(TmpRec!SDentalStatus)), Trim(TmpRec!SDentalStatus), "")
                    tCPlnPSpecStatus = IIf(Len(Trim(TmpRec!PSpecStatus)), Trim(TmpRec!PSpecStatus), "")
                    tCPlnSSpecStatus = IIf(Len(Trim(TmpRec!SSpecStatus)), Trim(TmpRec!SSpecStatus), "")
                    tCPlnPMepStatus = IIf(Len(Trim(TmpRec!PMepStatus)), Trim(TmpRec!PMepStatus), "")
                    tCPlnSMepStatus = IIf(Len(Trim(TmpRec!SMepStatus)), Trim(TmpRec!SMepStatus), "")
                    tCRoundUpStatus = IIf(Len(Trim(TmpRec!PLNRoundUpStatus)), Trim(TmpRec!PLNRoundUpStatus), "Y")
                    tCProRate = IIf(Len(Trim(TmpRec!PLNProRate)), Trim(TmpRec!PLNProRate), "Y")
                    tCMCOProcfeeStatus = IIf(Len(Trim(TmpRec!MCOProcfeeStatus)), Trim(TmpRec!MCOProcfeeStatus), "N")
                ElseIf tmpTableName = "AnnualLimit" Then
                    tCAnnSuppLmtStatus = IIf(Len(Trim(TmpRec!supplimitstatus)), TmpRec!supplimitstatus, "A")
                End If
                TmpRec.MoveNext
            End With
        Loop
    End If
    TmpRec.Close
    Set TmpRec = Nothing
End Sub

Private Sub SaveMBMAmendHistory(tmpMBMNo As String, tmpPOlNo As String, tmpICNo As String, tmpName As String _
            , tmpPolExpDate As String, tmpPlnCode As String, tmpEndtType As String _
            , tmpEndtEffDate As String, tmpSubmitDate As String)
Dim MBMHisSql As String
    
    tmpPolExpDate = IIf(IsDate(tmpPolExpDate), Format(tmpPolExpDate), "")
    tmpEndtEffDate = IIf(IsDate(tmpEndtEffDate), Format(tmpEndtEffDate), "")
    tmpSubmitDate = IIf(IsDate(tmpSubmitDate), Format(tmpSubmitDate), "")
    MBMHisSql = "EXEC Upload_MBMAmendHistory '" & tmpMBMNo & "', '" & tmpPOlNo & "', '" & tmpICNo & "', " & _
                "'" & tmpName & "','" & tmpPolExpDate & "', '" & tmpPlnCode & "','" & tmpEndtType & "', '" & _
                "'" & tmpEndtEffDate & "','" & tmpSubmitDate & "','" & Format(Date, "yyyy/mm/dd") & "','" & Logon & "'"
    medixmasterconn.Execute MBMHisSql
End Sub

Private Sub SaveMBMOthers(tmpDataType As String, tmpPrem As Double, tmpBPrem As Double, tmpBMCO As Double, tmpMCO As Double, tmpPolDis As Double, tmpRewDis As Double, tmpVClmNonPayFr As String, tmpVClmNonPayTo As String, tmpVEffDate As String, tmpVGuaRenewal As String)
Dim MBMOWSql As String
Dim tmpVGrpCompName, tmpVPrevTOPln As String
Dim tmpVDateJoin As String
Dim tmpVPrevPolNo, tmpVPrevMBMNo, tmpVEmpNo As String
Dim tmpVAgentCode, tmpVDepartment As String
Dim tmpVPolSubNo1, tmpVPolSubNo2 As String
Dim tmpVBRNCode, tmpVPaymentMode As String
Dim tmpVUnExcess As Double
Dim tmpVUpgradePln, tmpVlblStatus  As String
Dim tmpVMBMNewPolEffDate As String
Dim tmpVEndtRefNo, tmpVPolCondition As String
Dim tmpJoinDate As String
Dim tmpInvoice As String
    tmpInvoice = ""
    tmpVlblStatus = Mid(cboLBLStatus, 1, 1)
    tmpVlblStatus = IIf(Len(tmpVlblStatus), tmpVlblStatus, "N")
    tmpVAgentCode = Trim(Mid(TextLine, 27, 15))
    If Trim(txtBordxFormat) = "3" Then
        tmpVGrpCompName = Trim(Mid(TextLine, 398, 37))
    Else
        tmpVGrpCompName = Trim(Mid(TextLine, 395, 40))
    End If
    tmpVBRNCode = Trim(Mid(TextLine, 2939, 15))
    tmpVDateJoin = Mid(TextLine, 3012, 2) & "/" & Mid(TextLine, 3014, 2) & "/" & Mid(TextLine, 3016, 4)
    tmpJoinDate = IIf(IsDate(tmpVDateJoin), Format(tmpVDateJoin, "YYYY/MM/DD"), Format(tmpVEffDate, "YYYY/MM/DD"))
    tmpJoinDate = IIf(tmpJoinDate <> "  /  /    ", tmpJoinDate, "")
    tmpVPrevTOPln = Trim(Mid(TextLine, 3371, 4))
    tmpVPaymentMode = Trim(Mid(TextLine, 3375, 1))
    tmpVPolSubNo1 = Trim(Mid(TextLine, 3376, 10))
    tmpVPolSubNo2 = Trim(Mid(TextLine, 3386, 10))
    tmpVPrevPolNo = Trim(Mid(TextLine, 2954, 25))
    tmpVPrevMBMNo = Trim(Mid(TextLine, 2979, 18))
    tmpVEmpNo = Trim(Mid(TextLine, 2997, 15))
    tmpVDepartment = Trim(Mid(TextLine, 3221, 30))
    tmpVDepartment = Replace(tmpVDepartment, "'", "''")
    tmpVUpgradePln = ""
    tmpVMBMNewPolEffDate = ""
    If Trim(Mid(TextLine, 3592, 4)) <> "" And Trim(Mid(TextLine, 3596, 8)) <> "" Then
       tmpVUpgradePln = Trim(Mid(TextLine, 3592, 4))
       tmpVMBMNewPolEffDate = Mid(TextLine, 3596, 2) & "/" & Mid(TextLine, 3598, 2) & "/" & Mid(TextLine, 3600, 4)
    End If
    tmpVMBMNewPolEffDate = IIf(IsDate(tmpVMBMNewPolEffDate), Format(tmpVMBMNewPolEffDate, "YYYY/MM/DD"), "")
    tmpVEndtRefNo = Trim(Mid(TextLine, 3604, 20))
    
    If txtBordxFormat = "3" Then
        tmpVPolCondition = Trim(Mid(TextLine, 3800, 5))
        If tmpFileStatus = "I" Or tmpFileStatus = "R" Then
            tmpMCO = IIf(IsNumeric(Trim(Mid(TextLine, 3684, 12))), Trim(Mid(TextLine, 3684, 12)), 0)
            tmpBMCO = IIf(IsNumeric(Trim(Mid(TextLine, 3696, 12))), Trim(Mid(TextLine, 3696, 12)), 0)
        Else
            tmpMCO = IIf(IsNumeric(Trim(Mid(TextLine, 3696, 12))), Trim(Mid(TextLine, 3696, 12)), 0)
            tmpBMCO = IIf(IsNumeric(Trim(Mid(TextLine, 3608, 12))), Trim(Mid(TextLine, 3608, 12)), 0)
        End If
    Else
        tmpVPolCondition = ""
    End If
    
    If txtBordxFormat = "3" Then
        tmpVUnExcess = 0
    Else
        tmpVUnExcess = IIf(IsNumeric(Trim(Mid(TextLine, 3830, 12))), Trim(Mid(TextLine, 3830, 12)), 0)
    End If
    tmpVAgentCode = Trim(Replace(tmpVAgentCode, "'", "''"))
    tmpVGrpCompName = Replace(tmpVGrpCompName, "'", "''")
    If Trim(tmpFileStatus) = "C" Then
        tmpInvoice = "VOID"
    Else
        tmpInvoice = ""
    End If
    
    
    MBMOWSql = "EXEC Upload_MBMOthersNew '" & Trim(tMBMNo) & "','" & tmpVAgentCode & "','" & tmpVlblStatus & "', " & _
                "'" & tMBMNo & "'," & tmpPrem & "," & tmpBPrem & "," & tmpBMCO & "," & tmpMCO & ", " & _
                "'" & tmpDataType & "'," & tmpPolDis & "," & tmpRewDis & ",'" & tmpVGrpCompName & "', '" & tmpVPrevTOPln & "', " & _
                "'" & tmpJoinDate & "','" & tmpVPrevPolNo & "','" & tmpVPrevMBMNo & "','" & tmpVEmpNo & "', " & _
                "'" & tmpVDepartment & "','" & tmpVPaymentMode & "','" & tmpVPolSubNo1 & "','" & tmpVPolSubNo2 & "','" & tmpVBRNCode & "', " & _
                "'" & tmpVUpgradePln & "','" & tmpVMBMNewPolEffDate & "','" & tmpVEndtRefNo & "','" & tmpVPolCondition & "'," & tmpVUnExcess & ",'" & tmpVGuaRenewal & "'," & _
                "'" & tmpVClmNonPayFr & "','" & tmpVClmNonPayTo & "',  '" & tmpInvoice & "'"
                   
    medixmasterconn.Execute MBMOWSql

End Sub

Private Sub SaveMBMOthersAXA(tmpDataType As String, tmpPrem As Double, tmpBPrem As Double, tmpBMCO As Double, tmpMCO As Double, tmpPolDis As Double, tmpRewDis As Double, tmpVClmNonPayFr As String, tmpVClmNonPayTo As String, tmpVEffDate As String, tmpVGuaRenewal As String, tmpVAXAMBMNo As String)
Dim MBMOWSql As String
Dim tmpVGrpCompName, tmpVPrevTOPln As String
Dim tmpVDateJoin As String
Dim tmpVPrevPolNo, tmpVPrevMBMNo, tmpVEmpNo As String
Dim tmpVAgentCode, tmpVDepartment As String
Dim tmpVPolSubNo1, tmpVPolSubNo2 As String
Dim tmpVBRNCode, tmpVPaymentMode As String
Dim tmpVUnExcess As Double
Dim tmpVUpgradePln, tmpVlblStatus  As String
Dim tmpVMBMNewPolEffDate As String
Dim tmpVEndtRefNo, tmpVPolCondition As String
Dim tmpJoinDate As String
Dim tmpInvoice As String
    tmpInvoice = ""
    tmpVlblStatus = Mid(cboLBLStatus, 1, 1)
    tmpVlblStatus = IIf(Len(tmpVlblStatus), tmpVlblStatus, "N")
    tmpVAgentCode = ""
    'If Trim(txtBordxFormat) = "3" Then
    '    tmpVGrpCompName = Trim(Mid(TextLine, 398, 37))
    'Else
        tmpVGrpCompName = Trim(RecordArray(2))
    'End If
    tmpVBRNCode = Trim(RecordArray(19))
    tmpVDateJoin = ""
    tmpJoinDate = IIf(IsDate(tmpVDateJoin), Format(tmpVDateJoin, "YYYY/MM/DD"), Format(tmpVEffDate, "YYYY/MM/DD"))
    tmpJoinDate = IIf(tmpJoinDate <> "  /  /    ", tmpJoinDate, "")
    tmpVPrevTOPln = ""
    tmpVPaymentMode = ""
    tmpVPolSubNo1 = ""
    tmpVPolSubNo2 = ""
    tmpVPrevPolNo = ""
    'tmpVPrevMBMNo = ""
    tmpVEmpNo = ""
    tmpVDepartment = ""
    tmpVDepartment = Replace(tmpVDepartment, "'", "''")
    tmpVUpgradePln = ""
    tmpVMBMNewPolEffDate = ""
    'If Trim(Mid(TextLine, 3592, 4)) <> "" And Trim(Mid(TextLine, 3596, 8)) <> "" Then
       tmpVUpgradePln = ""
       tmpVMBMNewPolEffDate = ""
    'End If
    tmpVMBMNewPolEffDate = ""
    tmpVEndtRefNo = ""
    
    'If txtBordxFormat = "3" Then
    '    tmpVPolCondition = Trim(Mid(TextLine, 3800, 5))
    '    If tmpFileStatus = "I" Or tmpFileStatus = "R" Then
    '        tmpMCO = IIf(IsNumeric(Trim(Mid(TextLine, 3684, 12))), Trim(Mid(TextLine, 3684, 12)), 0)
    '        tmpBMCO = IIf(IsNumeric(Trim(Mid(TextLine, 3696, 12))), Trim(Mid(TextLine, 3696, 12)), 0)
    '    Else
    '        tmpMCO = IIf(IsNumeric(Trim(Mid(TextLine, 3696, 12))), Trim(Mid(TextLine, 3696, 12)), 0)
    '        tmpBMCO = IIf(IsNumeric(Trim(Mid(TextLine, 3608, 12))), Trim(Mid(TextLine, 3608, 12)), 0)
    '    End If
    'Else
        tmpVPolCondition = ""
    'End If
    
    'If txtBordxFormat = "3" Then
        tmpVUnExcess = 0
    'Else
    '    tmpVUnExcess = IIf(IsNumeric(Trim(Mid(TextLine, 3830, 12))), Trim(Mid(TextLine, 3830, 12)), 0)
    'End If
    tmpVAgentCode = Trim(Replace(tmpVAgentCode, "'", "''"))
    tmpVGrpCompName = Replace(tmpVGrpCompName, "'", "''")
    If Trim(tmpFileStatus) = "C" Then
        tmpInvoice = "VOID"
    Else
        tmpInvoice = ""
    End If
    
    
    MBMOWSql = "EXEC Upload_MBMOthersNew '" & Trim(tMBMNo) & "','" & tmpVAgentCode & "','" & tmpVlblStatus & "', " & _
                "'" & tMBMNo & "'," & tmpPrem & "," & tmpBPrem & "," & tmpBMCO & "," & tmpMCO & ", " & _
                "'" & tmpDataType & "'," & tmpPolDis & "," & tmpRewDis & ",'" & tmpVGrpCompName & "', '" & tmpVPrevTOPln & "', " & _
                "'" & tmpJoinDate & "','" & tmpVPrevPolNo & "','" & tmpVAXAMBMNo & "','" & tmpVEmpNo & "', " & _
                "'" & tmpVDepartment & "','" & tmpVPaymentMode & "','" & tmpVPolSubNo1 & "','" & tmpVPolSubNo2 & "','" & tmpVBRNCode & "', " & _
                "'" & tmpVUpgradePln & "','" & tmpVMBMNewPolEffDate & "','" & tmpVEndtRefNo & "','" & tmpVPolCondition & "'," & tmpVUnExcess & ",'" & tmpVGuaRenewal & "'," & _
                "'" & tmpVClmNonPayFr & "','" & tmpVClmNonPayTo & "',  '" & tmpInvoice & "'"
                   
    medixmasterconn.Execute MBMOWSql

End Sub
Private Sub SaveMBMOthersTwo()
Dim MBMOTwoWSql As String
Dim TmpVBlood, tmpVAllergies As String
Dim tmpVWeight, tmpVHeight, TmpVOccupation, TmpVWorkNature As String
Dim TmpVSmoking, TmpVAlcohol, TmpVMaritalStatus As String
Dim tmpVGrpCompAdd1, tmpVGrpCompAdd2, tmpVGrpCompAdd3, tmpVGrpCompAdd4 As String
Dim tmpVRiskNo, tmpVTransNo, tmpVPayNo, tmpVPaySeqNo As String
Dim tmpVClientNo, tmpVClientID As String
Dim tmpVRBAmt As Double
Dim tmpVPrevPlnCode, tmpVPrevInsCom, tmpVPayRemarks As String
Dim tmpVWarrExc, tmpVBTBG As String
Dim tmpVPayPremNet As Double
Dim tmpVPayPremGross As Double
Dim tmpVPayMcoNet As Double
Dim tmpVPayMCOGross As Double
Dim tmpVLoadingPrem As Double
Dim tmpVFamDiscAmt As Double
Dim tmpVRenewDiscAmt As Double
Dim tmpVRnBStatus As String
Dim tmpVPOLIssuedDate As String
Dim tmpVMBMBankAccNo As String
    
    tmpVWeight = Trim(Mid(TextLine, 312, 3))
    tmpVHeight = Trim(Mid(TextLine, 315, 3))
    TmpVOccupation = Trim(Mid(TextLine, 320, 30))
    TmpVWorkNature = Trim(Mid(TextLine, 350, 1))
    TmpVBlood = Trim(Mid(TextLine, 351, 3))
    tmpVAllergies = Trim(Mid(TextLine, 354, 20))
    tmpVWarrExc = Trim(Mid(TextLine, 436, 2500))
    TmpVSmoking = Trim(Mid(TextLine, 2936, 1))
    TmpVAlcohol = Trim(Mid(TextLine, 2937, 1))
    TmpVMaritalStatus = Trim(Mid(TextLine, 2938, 1))
    tmpVGrpCompAdd1 = Trim(Mid(TextLine, 3251, 30))
    tmpVGrpCompAdd2 = Trim(Mid(TextLine, 3281, 30))
    tmpVGrpCompAdd3 = Trim(Mid(TextLine, 3311, 30))
    tmpVGrpCompAdd4 = Trim(Mid(TextLine, 3341, 30))
    tmpVPOLIssuedDate = Mid(TextLine, 3396, 2) & "/" & Mid(TextLine, 3398, 2) & "/" & Mid(TextLine, 3400, 4)
    tmpVPOLIssuedDate = IIf(IsDate(tmpVPOLIssuedDate), Format(tmpVPOLIssuedDate, "YYYY/MM/DD"), "")
    tmpVRiskNo = Trim(Mid(TextLine, 3404, 4))
    tmpVTransNo = Trim(Mid(TextLine, 3408, 5))
    tmpVPayNo = Trim(Mid(TextLine, 3413, 8))
    tmpVPaySeqNo = Trim(Mid(TextLine, 3421, 2))
    tmpVClientNo = Trim(Mid(TextLine, 3423, 8))
    tmpVClientID = Trim(Mid(TextLine, 3431, 2))
    tmpVPrevPlnCode = Trim(Mid(TextLine, 3433, 15))
    tmpVRBAmt = IIf(IsNumeric(Trim(Mid(TextLine, 3448, 8))), Trim(Mid(TextLine, 3448, 8)), 0)
    tmpVPrevInsCom = Trim(Mid(TextLine, 3456, 60))
    tmpVPayRemarks = Trim(Mid(TextLine, 3516, 60))
    
    If Trim(txtBordxFormat) = "3" Then
        tmpVPayPremNet = 0
        tmpVPayPremGross = 0
        tmpVPayMcoNet = IIf(IsNumeric(Trim(Mid(TextLine, 3684, 12))), Trim(Mid(TextLine, 3684, 12)), 0)
        tmpVPayMCOGross = IIf(IsNumeric(Trim(Mid(TextLine, 3696, 12))), Trim(Mid(TextLine, 3696, 12)), 0)
        tmpVRnBStatus = ""
        tmpVLoadingPrem = 0
        tmpVFamDiscAmt = 0
        tmpVRenewDiscAmt = 0
        tmpVMBMBankAccNo = ""
    Else
        tmpVPayPremNet = IIf(IsNumeric(Trim(Mid(TextLine, 3684, 12))), Trim(Mid(TextLine, 3684, 12)), 0)
        tmpVPayPremGross = IIf(IsNumeric(Trim(Mid(TextLine, 3696, 12))), Trim(Mid(TextLine, 3696, 12)), 0)
        tmpVPayMcoNet = IIf(IsNumeric(Trim(Mid(TextLine, 3708, 12))), Trim(Mid(TextLine, 3708, 12)), 0)
        tmpVPayMCOGross = IIf(IsNumeric(Trim(Mid(TextLine, 3720, 12))), Trim(Mid(TextLine, 3720, 12)), 0)
        tmpVRnBStatus = Trim(Mid(TextLine, 3762, 2))
        tmpVLoadingPrem = IIf(IsNumeric(Trim(Mid(TextLine, 3764, 12))), Trim(Mid(TextLine, 3764, 12)), 0)
        tmpVFamDiscAmt = IIf(IsNumeric(Trim(Mid(TextLine, 3776, 12))), Trim(Mid(TextLine, 3776, 12)), 0)
        tmpVRenewDiscAmt = IIf(IsNumeric(Trim(Mid(TextLine, 3788, 12))), Trim(Mid(TextLine, 3788, 12)), 0)
        tmpVMBMBankAccNo = Trim(Mid(TextLine, 3805, 25))
    End If
    
    tmpVBTBG = Mid(CboBack2Back, 1, 1)
    tmpVBTBG = IIf(Len(tmpVBTBG), tmpVBTBG, "N")
    tmpVMBMBankAccNo = Trim(Replace(tmpVMBMBankAccNo, "'", "''"))
    tmpVWarrExc = Trim(Replace(tmpVWarrExc, "'", "''"))
    tmpVGrpCompAdd1 = Trim(Replace(tmpVGrpCompAdd1, "'", "''"))
    tmpVGrpCompAdd2 = Trim(Replace(tmpVGrpCompAdd2, "'", "''"))
    tmpVGrpCompAdd3 = Trim(Replace(tmpVGrpCompAdd3, "'", "''"))
    tmpVGrpCompAdd4 = Trim(Replace(tmpVGrpCompAdd4, "'", "''"))
    tmpVAllergies = Trim(Replace(tmpVAllergies, "'", "''"))
    TmpVOccupation = Trim(Replace(TmpVOccupation, "'", "''"))
    tmpVPrevInsCom = Trim(Replace(tmpVPrevInsCom, "'", "''"))
    tmpVPayRemarks = Trim(Replace(tmpVPayRemarks, "'", "''"))
    tmpVPrevPlnCode = Trim(Replace(tmpVPrevPlnCode, "'", "''"))
    MBMOTwoWSql = "EXEC Upload_MBMOthersTwoNew '" & Trim(tMBMNo) & "','" & tmpVWeight & "','" & tmpVHeight & "','" & TmpVWorkNature & "', " & _
          "'" & TmpVBlood & "','" & TmpVOccupation & "','" & TmpVSmoking & "','" & TmpVAlcohol & "','" & TmpVMaritalStatus & "', " & _
          "'" & tmpVGrpCompAdd1 & "','" & tmpVGrpCompAdd2 & "','" & tmpVGrpCompAdd3 & "','" & tmpVGrpCompAdd4 & "', " & _
          "'" & tmpVRiskNo & "','" & tmpVTransNo & "','" & tmpVPayNo & "','" & tmpVPaySeqNo & "','" & tmpVMBMBankAccNo & "'," & _
          "'" & tmpVClientNo & "','" & tmpVClientID & "'," & tmpVRBAmt & "," & _
          "'" & tmpVPrevInsCom & "','" & tmpVPrevPlnCode & "','" & tmpVPayRemarks & "','" & tmpVRnBStatus & "','" & tmpVWarrExc & "', " & _
          "'" & tmpVBTBG & "'," & tmpVPayPremNet & "," & tmpVPayMcoNet & "," & tmpVPayMCOGross & "," & tmpVLoadingPrem & ", " & _
          "" & tmpVFamDiscAmt & "," & tmpVRenewDiscAmt & ", " & _
          "" & tmpVPayPremGross & ",'" & tmpVPOLIssuedDate & "','" & tmpVAllergies & "'"
    medixmasterconn.Execute MBMOTwoWSql

End Sub

Private Sub SaveMBMOthersTwoAXA()
Dim MBMOTwoWSql As String
Dim TmpVBlood, tmpVAllergies As String
Dim tmpVWeight, tmpVHeight, TmpVOccupation, TmpVWorkNature As String
Dim TmpVSmoking, TmpVAlcohol, TmpVMaritalStatus As String
Dim tmpVGrpCompAdd1, tmpVGrpCompAdd2, tmpVGrpCompAdd3, tmpVGrpCompAdd4 As String
Dim tmpVRiskNo, tmpVTransNo, tmpVPayNo, tmpVPaySeqNo As String
Dim tmpVClientNo, tmpVClientID As String
Dim tmpVRBAmt As Double
Dim tmpVPrevPlnCode, tmpVPrevInsCom, tmpVPayRemarks As String
Dim tmpVWarrExc, tmpVBTBG As String
Dim tmpVPayPremNet As Double
Dim tmpVPayPremGross As Double
Dim tmpVPayMcoNet As Double
Dim tmpVPayMCOGross As Double
Dim tmpVLoadingPrem As Double
Dim tmpVFamDiscAmt As Double
Dim tmpVRenewDiscAmt As Double
Dim tmpVRnBStatus As String
Dim tmpVPOLIssuedDate As String
Dim tmpVMBMBankAccNo As String
    
    tmpVWeight = ""
    tmpVHeight = ""
    TmpVOccupation = ""
    TmpVWorkNature = ""
    TmpVBlood = ""
    tmpVAllergies = ""
    tmpVWarrExc = Trim(RecordArray(17)) & Trim(RecordArray(18))
    TmpVSmoking = ""
    TmpVAlcohol = ""
    TmpVMaritalStatus = ""
    tmpVGrpCompAdd1 = ""
    tmpVGrpCompAdd2 = ""
    tmpVGrpCompAdd3 = ""
    tmpVGrpCompAdd4 = ""
    tmpVPOLIssuedDate = ""
    tmpVPOLIssuedDate = IIf(IsDate(tmpVPOLIssuedDate), Format(tmpVPOLIssuedDate, "YYYY/MM/DD"), "")
    tmpVRiskNo = ""
    tmpVTransNo = ""
    tmpVPayNo = ""
    tmpVPaySeqNo = ""
    tmpVClientNo = ""
    tmpVClientID = ""
    tmpVPrevPlnCode = ""
    tmpVRBAmt = 0
    tmpVPrevInsCom = ""
    tmpVPayRemarks = ""
    
   ' If Trim(txtBordxFormat) = "3" Then
        tmpVPayPremNet = 0
        tmpVPayPremGross = 0
        tmpVPayMcoNet = 0
        tmpVPayMCOGross = 0
        tmpVRnBStatus = ""
        tmpVLoadingPrem = 0
        tmpVFamDiscAmt = 0
        tmpVRenewDiscAmt = 0
        tmpVMBMBankAccNo = ""
   ' Else
   '     tmpVPayPremNet = IIf(IsNumeric(Trim(Mid(TextLine, 3684, 12))), Trim(Mid(TextLine, 3684, 12)), 0)
   '     tmpVPayPremGross = IIf(IsNumeric(Trim(Mid(TextLine, 3696, 12))), Trim(Mid(TextLine, 3696, 12)), 0)
   '     tmpVPayMcoNet = IIf(IsNumeric(Trim(Mid(TextLine, 3708, 12))), Trim(Mid(TextLine, 3708, 12)), 0)
   '     tmpVPayMCOGross = IIf(IsNumeric(Trim(Mid(TextLine, 3720, 12))), Trim(Mid(TextLine, 3720, 12)), 0)
   '     tmpVRnBStatus = Trim(Mid(TextLine, 3762, 2))
   '     tmpVLoadingPrem = IIf(IsNumeric(Trim(Mid(TextLine, 3764, 12))), Trim(Mid(TextLine, 3764, 12)), 0)
   '     tmpVFamDiscAmt = IIf(IsNumeric(Trim(Mid(TextLine, 3776, 12))), Trim(Mid(TextLine, 3776, 12)), 0)
   '     tmpVRenewDiscAmt = IIf(IsNumeric(Trim(Mid(TextLine, 3788, 12))), Trim(Mid(TextLine, 3788, 12)), 0)
   '     tmpVMBMBankAccNo = Trim(Mid(TextLine, 3805, 25))
   ' End If
    
    tmpVBTBG = Mid(CboBack2Back, 1, 1)
    tmpVBTBG = IIf(Len(tmpVBTBG), tmpVBTBG, "N")
    tmpVMBMBankAccNo = ""
    tmpVWarrExc = Trim(Replace(tmpVWarrExc, "'", "''"))
    tmpVGrpCompAdd1 = Trim(Replace(tmpVGrpCompAdd1, "'", "''"))
    tmpVGrpCompAdd2 = Trim(Replace(tmpVGrpCompAdd2, "'", "''"))
    tmpVGrpCompAdd3 = Trim(Replace(tmpVGrpCompAdd3, "'", "''"))
    tmpVGrpCompAdd4 = Trim(Replace(tmpVGrpCompAdd4, "'", "''"))
    tmpVAllergies = Trim(Replace(tmpVAllergies, "'", "''"))
    TmpVOccupation = Trim(Replace(TmpVOccupation, "'", "''"))
    tmpVPrevInsCom = Trim(Replace(tmpVPrevInsCom, "'", "''"))
    tmpVPayRemarks = Trim(Replace(tmpVPayRemarks, "'", "''"))
    tmpVPrevPlnCode = Trim(Replace(tmpVPrevPlnCode, "'", "''"))
    MBMOTwoWSql = "EXEC Upload_MBMOthersTwoNew '" & Trim(tMBMNo) & "','" & tmpVWeight & "','" & tmpVHeight & "','" & TmpVWorkNature & "', " & _
          "'" & TmpVBlood & "','" & TmpVOccupation & "','" & TmpVSmoking & "','" & TmpVAlcohol & "','" & TmpVMaritalStatus & "', " & _
          "'" & tmpVGrpCompAdd1 & "','" & tmpVGrpCompAdd2 & "','" & tmpVGrpCompAdd3 & "','" & tmpVGrpCompAdd4 & "', " & _
          "'" & tmpVRiskNo & "','" & tmpVTransNo & "','" & tmpVPayNo & "','" & tmpVPaySeqNo & "','" & tmpVMBMBankAccNo & "'," & _
          "'" & tmpVClientNo & "','" & tmpVClientID & "'," & tmpVRBAmt & "," & _
          "'" & tmpVPrevInsCom & "','" & tmpVPrevPlnCode & "','" & tmpVPayRemarks & "','" & tmpVRnBStatus & "','" & tmpVWarrExc & "', " & _
          "'" & tmpVBTBG & "'," & tmpVPayPremNet & "," & tmpVPayMcoNet & "," & tmpVPayMCOGross & "," & tmpVLoadingPrem & ", " & _
          "" & tmpVFamDiscAmt & "," & tmpVRenewDiscAmt & ", " & _
          "" & tmpVPayPremGross & ",'" & tmpVPOLIssuedDate & "','" & tmpVAllergies & "'"
    medixmasterconn.Execute MBMOTwoWSql

End Sub

Private Sub SaveMBMTwo(tmpVClinical As String)
Dim MBMTwoWSql As String
Dim tmpVSalutation, tmpVLanguage As String
Dim tmpVAppName, tmpVIcBcPp2nd, TmpVSRVCodeList As String
Dim tmpVNat, tmpVHomeTel, tmpVOffTel, tmpVMobilePhone, tmpVFilename As String
    
    TmpVSRVCodeList = "ME"
    tmpVNat = Trim(Mid(TextLine, 302, 10))
    tmpVSalutation = Trim(Mid(TextLine, 42, 5))
    tmpVLanguage = Trim(Mid(TextLine, 318, 1))
    tmpVHomeTel = Trim(Mid(TextLine, 237, 12))
    tmpVOffTel = Trim(Mid(TextLine, 249, 12))
    tmpVMobilePhone = Trim(Mid(TextLine, 261, 12))
    If txtBordxFormat = "3" Then
        tmpVIcBcPp2nd = Trim(Mid(TextLine, 3742, 20))
    Else
        tmpVIcBcPp2nd = ""
    End If
    
    tmpVAppName = Trim(Mid(TextLine, 3624, 60))
    tmpVAppName = Trim(Replace(tmpVAppName, "'", "''"))
    tmpVSalutation = Trim(Replace(tmpVSalutation, "'", "''"))
    tmpVNat = Trim(Replace(tmpVNat, "'", "''"))
    'tmpVFilename = txtImportFilename
    MBMTwoWSql = "EXEC Upload_MBMTwo '" & Trim(tMBMNo) & " ','" & tmpVSalutation & "','" & tmpVHomeTel & "','" & tmpVMobilePhone & "'," & _
                 "'" & tmpVOffTel & "','" & tmpVNat & "','" & tmpVLanguage & "','" & TmpVSRVCodeList & "', '" & Logon & "'," & _
                 "'" & Logon & "','" & tmpVClinical & "','" & tmpVAppName & "','" & tmpVIcBcPp2nd & "','A','" & tmpfilename & "'"
    medixmasterconn.Execute MBMTwoWSql

End Sub

Private Sub SaveMBMTwoAXA(tmpVClinical As String)
Dim MBMTwoWSql As String
Dim tmpVSalutation, tmpVLanguage As String
Dim tmpVAppName, tmpVIcBcPp2nd, TmpVSRVCodeList As String
Dim tmpVNat, tmpVHomeTel, tmpVOffTel, tmpVMobilePhone, tmpVFilename As String
    
    TmpVSRVCodeList = "ME"
    tmpVNat = ""
    tmpVSalutation = ""
    tmpVLanguage = ""
    tmpVHomeTel = ""
    tmpVOffTel = ""
    tmpVMobilePhone = ""
    tmpVIcBcPp2nd = ""
    
    tmpVAppName = Trim(RecordArray(2))
    tmpVAppName = Trim(Replace(tmpVAppName, "'", "''"))
    tmpVSalutation = ""
    tmpVNat = ""
    'tmpVFilename = txtImportFilename
    MBMTwoWSql = "EXEC Upload_MBMTwo '" & Trim(tMBMNo) & " ','" & tmpVSalutation & "','" & tmpVHomeTel & "','" & tmpVMobilePhone & "'," & _
                 "'" & tmpVOffTel & "','" & tmpVNat & "','" & tmpVLanguage & "','" & TmpVSRVCodeList & "', '" & Logon & "'," & _
                 "'" & Logon & "','" & tmpVClinical & "','" & tmpVAppName & "','" & tmpVIcBcPp2nd & "','A','" & tmpfilename & "'"
    medixmasterconn.Execute MBMTwoWSql

End Sub
Private Sub SaveMBMCLNOthers(TmpCISMBMType As String, tCPayor As String, tmpVPolicy As String, tmpClnStatus As String, _
tmpClnDAtaStatus As String, tmpClnExpStatus As String, tmpCisPlnCOde As String, tmpCisInsType As String, tmpCisWarranty As String, _
CISBasicPrem As Double, CisPrem As Double, CisBasicMco As Double, CisMCO As Double, CisBalLmt As Double, CisVisitLmt As Double, _
CisDentalLmt As Double, CisSpecLmt As Double, CisMepLmt As Double, tmpCisVEffDate As String, tmpCisVExpDate As String, tmpCisVCardType As String)
Dim StrCLNSql As String
Dim tEffDate, tExpDate As String
Dim FFSInvNo As String
Dim FFSInvBy As String
Dim FFSInvDate As String
Dim FFSInved As String
    FFSInved = ""
    FFSInvNo = ""
    FFSInvBy = ""
    FFSInvDate = ""
    If Trim(tCMCOProcfeeStatus) = "Y" Then
        FFSInved = "Y"
        FFSInvNo = "FFS"
        FFSInvBy = Logon
        FFSInvDate = Format(Date, "YYYY/MM/DD")
    End If
    tEffDate = IIf(IsDate(tmpCisVEffDate), Format(tmpCisVEffDate, "YYYY/MM/DD"), "")
    tExpDate = IIf(IsDate(tmpCisVExpDate), Format(tmpCisVExpDate, "YYYY/MM/DD"), "")
    
    StrCLNSql = "EXEC Upload_MBMClnOthers '" & TmpCISMBMType & "','" & Trim(tCPayor) & "','" & Trim(tMBMNo) & "'," & _
                "'" & tmpVPolicy & "','" & tmpClnStatus & "','" & tmpClnDAtaStatus & "','" & tmpClnExpStatus & "', " & _
                "'" & tmpCisPlnCOde & "','" & tmpCisInsType & "','" & tmpCisWarranty & "'," & CISBasicPrem & "," & CisPrem & ", " & _
                "" & CisBasicMco & "," & CisMCO & "," & CisBalLmt & "," & CisVisitLmt & "," & _
                "" & CisDentalLmt & "," & CisSpecLmt & "," & CisMepLmt & "," & _
                "'" & tEffDate & "','" & tExpDate & "','" & tmpCisVCardType & "'," & _
                "'" & FFSInved & "','" & FFSInvNo & "','" & FFSInvDate & "','" & FFSInvBy & "'"
    medixmasterconn.Execute StrCLNSql
End Sub

Private Sub SaveMBMCLNOthersAXA(TmpCISMBMType As String, tCPayor As String, tmpVPolicy As String, tmpClnStatus As String, _
tmpClnDAtaStatus As String, tmpClnExpStatus As String, tmpCisPlnCOde As String, tmpCisInsType As String, tmpCisWarranty As String, _
CISBasicPrem As Double, CisPrem As Double, CisBasicMco As Double, CisMCO As Double, CisBalLmt As Double, CisVisitLmt As Double, _
CisDentalLmt As Double, CisSpecLmt As Double, CisMepLmt As Double, tmpCisVEffDate As String, tmpCisVExpDate As String, tmpCisVCardType As String)
Dim StrCLNSql As String
Dim tEffDate, tExpDate As String
Dim FFSInvNo As String
Dim FFSInvBy As String
Dim FFSInvDate As String
Dim FFSInved As String
    FFSInved = ""
    FFSInvNo = ""
    FFSInvBy = ""
    FFSInvDate = ""
    If Trim(tCMCOProcfeeStatus) = "Y" Then
        FFSInved = "Y"
        FFSInvNo = "FFS"
        FFSInvBy = Logon
        FFSInvDate = Format(Date, "YYYY/MM/DD")
    End If
    tEffDate = IIf(IsDate(tmpCisVEffDate), Format(tmpCisVEffDate, "YYYY/MM/DD"), "")
    tExpDate = IIf(IsDate(tmpCisVExpDate), Format(tmpCisVExpDate, "YYYY/MM/DD"), "")
    
    StrCLNSql = "EXEC Upload_MBMClnOthers '" & TmpCISMBMType & "','" & Trim(tCPayor) & "','" & Trim(tMBMNo) & "'," & _
                "'" & tmpVPolicy & "','" & tmpClnStatus & "','" & tmpClnDAtaStatus & "','" & tmpClnExpStatus & "', " & _
                "'" & tmpCisPlnCOde & "','" & tmpCisInsType & "','" & tmpCisWarranty & "'," & CISBasicPrem & "," & CisPrem & ", " & _
                "" & CisBasicMco & "," & CisMCO & "," & CisBalLmt & "," & CisVisitLmt & "," & _
                "" & CisDentalLmt & "," & CisSpecLmt & "," & CisMepLmt & "," & _
                "'" & tEffDate & "','" & tExpDate & "','" & tmpCisVCardType & "'," & _
                "'" & FFSInved & "','" & FFSInvNo & "','" & FFSInvDate & "','" & FFSInvBy & "'"
    medixmasterconn.Execute StrCLNSql
End Sub

Private Sub SaveMBMCovPerson(tmpVDataType As String, TmpCovID As String, tDOB As String, _
tmpVMemberType As String, tmpVMBMName As String, tmpVMBMStatus As String, tmpVBordxDate As String, tmpVBatchNo As String, _
tmpVIC As String, tmpVSex As String, tmpVPlanCode As String, _
tAgeCode As String, tAdultChild As String, tmpVAnnLmt As Double, tmpVBasicMco As Double, tmpVMcoAmt As Double, _
tmpVBasicPrem As Double, tmpVPremium As Double, tmpSuppEffDate As String, tmpSuppExpDate As String)

Dim SuppWSql As String
    tmpVMBMName = Trim(Replace(tmpVMBMName, "'", ""))
    'tmpVIC = Trim(Replace(tmpVMBMName, "'", ""))
    SuppWSql = ""
    SuppWSql = "EXEC Upload_MBMCovPersonsNew '" & tmpVDataType & "','" & tMBMNo & "','" & TmpCovID & "','" & tDOB & "'," & _
                "'" & tmpVMemberType & "','" & tmpVMBMName & "','" & tmpVMBMStatus & "','" & tmpVBordxDate & "','" & tmpVBatchNo & "', " & _
                "'" & tmpVIC & "','" & tmpVSex & "','" & tmpVPlanCode & "', " & _
                "'" & tAgeCode & "','" & tAdultChild & "'," & tmpVAnnLmt & "," & tmpVAnnLmt & "," & tmpVBasicMco & "," & tmpVMcoAmt & ", " & _
                "" & tmpVBasicPrem & "," & tmpVPremium & ",'" & Format(Date, "yyyy/mm/dd") & "','" & tmpSuppEffDate & "','" & tmpSuppExpDate & "'"
                
    medixmasterconn.Execute SuppWSql
End Sub


Private Sub SaveMBMCovPersonAXA(tmpVDataType As String, TmpCovID As String, tDOB As String, _
tmpVMemberType As String, tmpVMBMName As String, tmpVMBMStatus As String, tmpVBordxDate As String, tmpVBatchNo As String, _
tmpVIC As String, tmpVSex As String, tmpVPlanCode As String, _
tAgeCode As String, tAdultChild As String, tmpVAnnLmt As Double, tmpVBasicMco As Double, tmpVMcoAmt As Double, _
tmpVBasicPrem As Double, tmpVPremium As Double, tmpSuppEffDate As String, tmpSuppExpDate As String, tmpVAXAMBMNo As String)

Dim SuppWSql As String
    tmpVMBMName = Trim(Replace(tmpVMBMName, "'", ""))
    'tmpVIC = Trim(Replace(tmpVMBMName, "'", ""))
    SuppWSql = ""
    SuppWSql = "EXEC Upload_MBMCovPersonsAXA '" & tmpVDataType & "','" & tMBMNo & "','" & TmpCovID & "','" & tDOB & "'," & _
                "'" & tmpVMemberType & "','" & tmpVMBMName & "','" & tmpVMBMStatus & "','" & tmpVBordxDate & "','" & tmpVBatchNo & "', " & _
                "'" & tmpVIC & "','" & tmpVSex & "','" & tmpVPlanCode & "', " & _
                "'" & tAgeCode & "','" & tAdultChild & "'," & tmpVAnnLmt & "," & tmpVAnnLmt & "," & tmpVBasicMco & "," & tmpVMcoAmt & ", " & _
                "" & tmpVBasicPrem & "," & tmpVPremium & ",'" & Format(Date, "yyyy/mm/dd") & "','" & tmpSuppEffDate & "','" & tmpSuppExpDate & "', '" & tmpVAXAMBMNo & "'"
                
    medixmasterconn.Execute SuppWSql
End Sub

Private Sub SaveMBMCovPersonTwo(tmpVCovID As String, tmpVClmNonPayFr As String, _
tmpVClmNonPayTo As String, tmpVPolDis As Double, tmpVRenewDis As Double)
Dim Supp2WSql As String
Dim tmpVSalutation, TmpVBlood, tmpVAllergies As String
Dim TmpVOccupation, TmpVWorkNature As String
Dim TmpVSmoking, TmpVAlcohol As String
Dim tmpVPayNo, tmpVPaySeqNo As String
Dim tmpVClientNo, tmpVClientID As String
Dim tmpVRBAmt As Double
Dim tmpVPrevPlnCode, tmpVPrevInsCom, tmpVPayRemarks As String
Dim tmpVWarrExc, tmpVIcBcPp2nd As String
Dim tmpVPayPremNet As Double
Dim tmpVPayPremGross As Double
Dim tmpVPayMcoNet As Double
Dim tmpVPayMCOGross As Double
Dim tmpVLoadingPrem As Double
Dim tmpVFamDiscAmt As Double
Dim tmpVRenewDiscAmt As Double
Dim tmpVUnExcess As Double
    
    tmpVSalutation = Trim(Mid(TextLine, 42, 5))
    TmpVOccupation = Trim(Mid(TextLine, 320, 30))
    TmpVWorkNature = Trim(Mid(TextLine, 350, 1))
    TmpVBlood = Trim(Mid(TextLine, 351, 3))
    tmpVAllergies = Trim(Mid(TextLine, 354, 20))
    tmpVWarrExc = Trim(Mid(TextLine, 436, 2500))
    TmpVSmoking = Trim(Mid(TextLine, 2936, 1))
    TmpVAlcohol = Trim(Mid(TextLine, 2937, 1))
    tmpVPayNo = Trim(Mid(TextLine, 3413, 8))
    tmpVPaySeqNo = Trim(Mid(TextLine, 3421, 2))
    tmpVClientNo = Trim(Mid(TextLine, 3423, 8))
    tmpVClientID = Trim(Mid(TextLine, 3431, 2))
    tmpVPrevPlnCode = Trim(Mid(TextLine, 3433, 15))
    tmpVRBAmt = IIf(IsNumeric(Trim(Mid(TextLine, 3448, 8))), Trim(Mid(TextLine, 3448, 8)), 0)
    tmpVPrevInsCom = Trim(Mid(TextLine, 3456, 60))
    tmpVPayRemarks = Trim(Mid(TextLine, 3516, 60))
    tmpVSalutation = Trim(Replace(tmpVSalutation, "'", "''"))
    tmpVWarrExc = Trim(Replace(tmpVWarrExc, "'", "''"))
    
    If Trim(txtBordxFormat) = "3" Then
        tmpVPayPremNet = 0
        tmpVPayPremGross = 0
        tmpVPayMcoNet = IIf(IsNumeric(Trim(Mid(TextLine, 3684, 12))), Trim(Mid(TextLine, 3684, 12)), 0)
        tmpVPayMCOGross = IIf(IsNumeric(Trim(Mid(TextLine, 3696, 12))), Trim(Mid(TextLine, 3696, 12)), 0)
        tmpVLoadingPrem = 0
        tmpVFamDiscAmt = 0
        tmpVRenewDiscAmt = 0
        tmpVUnExcess = 0
        tmpVIcBcPp2nd = ""
    Else
        tmpVPayPremNet = IIf(IsNumeric(Trim(Mid(TextLine, 3684, 12))), Trim(Mid(TextLine, 3684, 12)), 0)
        tmpVPayPremGross = IIf(IsNumeric(Trim(Mid(TextLine, 3696, 12))), Trim(Mid(TextLine, 3696, 12)), 0)
        tmpVPayMcoNet = IIf(IsNumeric(Trim(Mid(TextLine, 3708, 12))), Trim(Mid(TextLine, 3708, 12)), 0)
        tmpVPayMCOGross = IIf(IsNumeric(Trim(Mid(TextLine, 3720, 12))), Trim(Mid(TextLine, 3720, 12)), 0)
        tmpVLoadingPrem = IIf(IsNumeric(Trim(Mid(TextLine, 3764, 12))), Trim(Mid(TextLine, 3764, 12)), 0)
        tmpVFamDiscAmt = IIf(IsNumeric(Trim(Mid(TextLine, 3776, 12))), Trim(Mid(TextLine, 3776, 12)), 0)
        tmpVRenewDiscAmt = IIf(IsNumeric(Trim(Mid(TextLine, 3788, 12))), Trim(Mid(TextLine, 3788, 12)), 0)
        tmpVUnExcess = IIf(IsNumeric(Trim(Mid(TextLine, 3830, 12))), Trim(Mid(TextLine, 3830, 12)), 0)
        tmpVSalutation = Trim(Replace(tmpVSalutation, "'", "''"))
        tmpVWarrExc = Trim(Replace(tmpVWarrExc, "'", "''"))
        tmpVIcBcPp2nd = Trim(Mid(TextLine, 3742, 20))
        
    End If
    
    
    Supp2WSql = "EXEC Upload_MBMCovPersonsTwo '" & Trim(tMBMNo) & "','" & tmpVCovID & "','" & tmpVSalutation & "', " & _
        "'" & TmpVOccupation & " ','" & TmpVWorkNature & "','" & TmpVBlood & "','" & tmpVAllergies & "','" & TmpVSmoking & "', " & _
        "'" & TmpVAlcohol & "','" & tmpVPrevPlnCode & "','" & tmpVPrevInsCom & "','" & tmpVPayRemarks & "','" & tmpVWarrExc & "', " & _
        "'" & tmpVPayNo & "','" & tmpVPaySeqNo & "','" & tmpVClientNo & "','" & tmpVClientID & "'," & tmpVRBAmt & "," & _
        "" & tmpVPayPremNet & ", " & tmpVPayMcoNet & ", " & tmpVPayMCOGross & ", " & _
        "'" & tmpVIcBcPp2nd & "', " & tmpVLoadingPrem & ", " & tmpVFamDiscAmt & ", " & tmpVRenewDiscAmt & ", " & tmpVPayPremGross & "," & _
        "" & tmpVUnExcess & ",'" & tmpVClmNonPayFr & "','" & tmpVClmNonPayTo & "'," & tmpVPolDis & "," & tmpVRenewDis
    medixmasterconn.Execute Supp2WSql
End Sub

Private Sub SaveMBMCovPersonTwoAXA(tmpVCovID As String, tmpVClmNonPayFr As String, _
tmpVClmNonPayTo As String, tmpVPolDis As Double, tmpVRenewDis As Double)
Dim Supp2WSql As String
Dim tmpVSalutation, TmpVBlood, tmpVAllergies As String
Dim TmpVOccupation, TmpVWorkNature As String
Dim TmpVSmoking, TmpVAlcohol As String
Dim tmpVPayNo, tmpVPaySeqNo As String
Dim tmpVClientNo, tmpVClientID As String
Dim tmpVRBAmt As Double
Dim tmpVPrevPlnCode, tmpVPrevInsCom, tmpVPayRemarks As String
Dim tmpVWarrExc, tmpVIcBcPp2nd As String
Dim tmpVPayPremNet As Double
Dim tmpVPayPremGross As Double
Dim tmpVPayMcoNet As Double
Dim tmpVPayMCOGross As Double
Dim tmpVLoadingPrem As Double
Dim tmpVFamDiscAmt As Double
Dim tmpVRenewDiscAmt As Double
Dim tmpVUnExcess As Double
    
    tmpVSalutation = ""
    TmpVOccupation = ""
    TmpVWorkNature = ""
    TmpVBlood = ""
    tmpVAllergies = ""
    tmpVWarrExc = Trim(RecordArray(18))
    TmpVSmoking = ""
    TmpVAlcohol = ""
    tmpVPayNo = ""
    tmpVPaySeqNo = ""
    tmpVClientNo = ""
    tmpVClientID = ""
    tmpVPrevPlnCode = ""
    tmpVRBAmt = 0
    tmpVPrevInsCom = ""
    tmpVPayRemarks = ""
    tmpVSalutation = Trim(Replace(tmpVSalutation, "'", "''"))
    tmpVWarrExc = Trim(Replace(tmpVWarrExc, "'", "''"))
    
    'If Trim(txtBordxFormat) = "3" Then
        tmpVPayPremNet = 0
        tmpVPayPremGross = 0
        tmpVPayMcoNet = 0
        tmpVPayMCOGross = 0
        tmpVLoadingPrem = 0
        tmpVFamDiscAmt = 0
        tmpVRenewDiscAmt = 0
        tmpVUnExcess = 0
        tmpVIcBcPp2nd = ""
    'Else
    '    tmpVPayPremNet = IIf(IsNumeric(Trim(Mid(TextLine, 3684, 12))), Trim(Mid(TextLine, 3684, 12)), 0)
    '    tmpVPayPremGross = IIf(IsNumeric(Trim(Mid(TextLine, 3696, 12))), Trim(Mid(TextLine, 3696, 12)), 0)
    '    tmpVPayMcoNet = IIf(IsNumeric(Trim(Mid(TextLine, 3708, 12))), Trim(Mid(TextLine, 3708, 12)), 0)
    '    tmpVPayMCOGross = IIf(IsNumeric(Trim(Mid(TextLine, 3720, 12))), Trim(Mid(TextLine, 3720, 12)), 0)
    '    tmpVLoadingPrem = IIf(IsNumeric(Trim(Mid(TextLine, 3764, 12))), Trim(Mid(TextLine, 3764, 12)), 0)
    '    tmpVFamDiscAmt = IIf(IsNumeric(Trim(Mid(TextLine, 3776, 12))), Trim(Mid(TextLine, 3776, 12)), 0)
    '    tmpVRenewDiscAmt = IIf(IsNumeric(Trim(Mid(TextLine, 3788, 12))), Trim(Mid(TextLine, 3788, 12)), 0)
    '    tmpVUnExcess = IIf(IsNumeric(Trim(Mid(TextLine, 3830, 12))), Trim(Mid(TextLine, 3830, 12)), 0)
    '    tmpVSalutation = Trim(Replace(tmpVSalutation, "'", "''"))
    '    tmpVWarrExc = Trim(Replace(tmpVWarrExc, "'", "''"))
    '    tmpVIcBcPp2nd = Trim(Mid(TextLine, 3742, 20))
    '
    'End If
    
    
    Supp2WSql = "EXEC Upload_MBMCovPersonsTwo '" & Trim(tMBMNo) & "','" & tmpVCovID & "','" & tmpVSalutation & "', " & _
        "'" & TmpVOccupation & " ','" & TmpVWorkNature & "','" & TmpVBlood & "','" & tmpVAllergies & "','" & TmpVSmoking & "', " & _
        "'" & TmpVAlcohol & "','" & tmpVPrevPlnCode & "','" & tmpVPrevInsCom & "','" & tmpVPayRemarks & "','" & tmpVWarrExc & "', " & _
        "'" & tmpVPayNo & "','" & tmpVPaySeqNo & "','" & tmpVClientNo & "','" & tmpVClientID & "'," & tmpVRBAmt & "," & _
        "" & tmpVPayPremNet & ", " & tmpVPayMcoNet & ", " & tmpVPayMCOGross & ", " & _
        "'" & tmpVIcBcPp2nd & "', " & tmpVLoadingPrem & ", " & tmpVFamDiscAmt & ", " & tmpVRenewDiscAmt & ", " & tmpVPayPremGross & "," & _
        "" & tmpVUnExcess & ",'" & tmpVClmNonPayFr & "','" & tmpVClmNonPayTo & "'," & tmpVPolDis & "," & tmpVRenewDis
    medixmasterconn.Execute Supp2WSql
End Sub

Private Sub SaveMBMCovPersonCLN(TmpCISMBMType As String, tmpVCovID As String, tmpClnStatus As String, _
tmpCisPlnCOde As String, tmpCisWarranty As String, CISBasicPrem As Double, CisPrem As Double, _
CisBasicMco As Double, CisMCO As Double, CisBalLmt As Double, CisVisitLmt As Double, _
CisDentalLmt As Double, CisSpecLmt As Double, CisMepLmt As Double, tmpSuppCisExpDate As String, _
tmpCisVWarranty As String, tmpCisVCardType As String, tmpCisVPlnCode As String)

Dim CisSuppSql As String
Dim tmpSExpDate As String
Dim FFSInvNo As String
Dim FFSInvBy As String
Dim FFSInvDate As String
    FFSInvNo = ""
    FFSInvBy = ""
    FFSInvDate = ""
    If Trim(tCMCOProcfeeStatus) = "Y" Then
        FFSInvNo = "FFS"
        FFSInvBy = Logon
        FFSInvDate = Format(Date, "YYYY/MM/DD")
    End If
    tmpSExpDate = IIf(IsDate(CDate(tmpSuppCisExpDate)), Format(tmpSuppCisExpDate, "yyyy/mm/dd"), "")

    CisSuppSql = "EXEC Upload_MBMCovPersonsCis '" & tMBMNo & "','" & tmpVCovID & "', " & _
        "" & CisMCO & "," & CisBasicMco & "," & CisPrem & "," & CISBasicPrem & ", " & _
        "" & CisBalLmt & "," & CisVisitLmt & "," & CisDentalLmt & "," & CisSpecLmt & ", " & _
        "" & CisMepLmt & ",'" & Format(Date, "yyyy/mm/dd") & "','" & Logon & "'," & _
        "'" & Format(Date, "yyyy/mm/dd") & "','" & Logon & "','" & tmpSuppCisExpDate & "'," & _
        "'" & tmpCisVCardType & "','" & tmpCisVWarranty & "','" & tmpClnStatus & "','" & tmpCisVPlnCode & "'," & _
        "'" & FFSInvNo & "','" & FFSInvDate & "','" & FFSInvBy & "'"
    medixmasterconn.Execute CisSuppSql
End Sub

Private Sub SaveMBMCovPersonCLNAXA(TmpCISMBMType As String, tmpVCovID As String, tmpClnStatus As String, _
tmpCisPlnCOde As String, tmpCisWarranty As String, CISBasicPrem As Double, CisPrem As Double, _
CisBasicMco As Double, CisMCO As Double, CisBalLmt As Double, CisVisitLmt As Double, _
CisDentalLmt As Double, CisSpecLmt As Double, CisMepLmt As Double, tmpSuppCisExpDate As String, _
tmpCisVWarranty As String, tmpCisVCardType As String, tmpCisVPlnCode As String)

Dim CisSuppSql As String
Dim tmpSExpDate As String
Dim FFSInvNo As String
Dim FFSInvBy As String
Dim FFSInvDate As String
    FFSInvNo = ""
    FFSInvBy = ""
    FFSInvDate = ""
    If Trim(tCMCOProcfeeStatus) = "Y" Then
        FFSInvNo = "FFS"
        FFSInvBy = Logon
        FFSInvDate = Format(Date, "YYYY/MM/DD")
    End If
    tmpSExpDate = IIf(IsDate(CDate(tmpSuppCisExpDate)), Format(tmpSuppCisExpDate, "yyyy/mm/dd"), "")

    CisSuppSql = "EXEC Upload_MBMCovPersonsCis '" & tMBMNo & "','" & tmpVCovID & "', " & _
        "" & CisMCO & "," & CisBasicMco & "," & CisPrem & "," & CISBasicPrem & ", " & _
        "" & CisBalLmt & "," & CisVisitLmt & "," & CisDentalLmt & "," & CisSpecLmt & ", " & _
        "" & CisMepLmt & ",'" & Format(Date, "yyyy/mm/dd") & "','" & Logon & "'," & _
        "'" & Format(Date, "yyyy/mm/dd") & "','" & Logon & "','" & tmpSuppCisExpDate & "'," & _
        "'" & tmpCisVCardType & "','" & tmpCisVWarranty & "','" & tmpClnStatus & "','" & tmpCisVPlnCode & "'," & _
        "'" & FFSInvNo & "','" & FFSInvDate & "','" & FFSInvBy & "'"
    medixmasterconn.Execute CisSuppSql
End Sub
Private Sub SaveMBMPrincipal(tmpVMBMName As String, tDOB As String, tmpVPolicy As String, _
tmpVCovID As String, tmpVMBMStatus As String, tmpVDataStatus As String, tmpVExpStatus As String, tmpVIC As String, _
tmpVMemberType As String, tmpVSex As String, tmpVPlanCode As String, tmpVINSCode As String, tmpVEffDate As String, tmpVExpDate As String, _
tAgeCode As String, tAdultChild As String, tmpVAnnLmt As Double, tmpVHSClnInd As String, _
tmpVTopupNo As String, tmpVTopupInd As String, tmpVBordxDate As String, tmpVBatchNo As String)

Dim MBMWSql As String
Dim tmpVAdd1, tmpVAdd2, tmpVAdd3 As String
Dim tmpVRaceCode, tmpVCityCode, tmpVPostCode, tmpVSTACode As String
Dim tmpVTakeOver, tmpVRenewal As String
Dim tmpVValueDate As String
Dim tEffDate, tExpDate As String
    tEffDate = IIf(IsDate(tmpVEffDate), Format(tmpVEffDate, "YYYY/MM/DD"), "")
    tExpDate = IIf(IsDate(tmpVExpDate), Format(tmpVExpDate, "YYYY/MM/DD"), "")

    tmpVAdd1 = Trim(Mid(TextLine, 107, 30))
    tmpVAdd2 = Trim(Mid(TextLine, 137, 30))
    tmpVAdd3 = Trim(Mid(TextLine, 167, 30))
    tmpVCityCode = Trim(Mid(TextLine, 197, 20))
    tmpVPostCode = Trim(Mid(TextLine, 217, 5))
    tmpVSTACode = Trim(Mid(TextLine, 222, 15))
    tmpVTakeOver = Trim(Mid(TextLine, 435, 1))
    tmpVRenewal = Trim(Mid(TextLine, 435, 1))
    If tmpVTakeOver = "C" Or tmpVTakeOver = "Y" Or tmpVRenewal = "" Then
        tmpVRenewal = "N"
    End If
    tmpVRaceCode = Trim(Mid(TextLine, 319, 1))
    tmpVAdd1 = Trim(Replace(tmpVAdd1, "'", "''"))
    tmpVAdd2 = Trim(Replace(tmpVAdd2, "'", "''"))
    tmpVAdd3 = Trim(Replace(tmpVAdd3, "'", "''"))
    tmpVCityCode = Trim(Replace(tmpVCityCode, "'", "''"))
    tmpVSTACode = Trim(Replace(tmpVSTACode, "'", "''"))
    tmpVValueDate = Format(Date, "yyyy/mm/dd")
    MBMWSql = "EXEC Upload_MBM_HIS '" & tPayor & "', '" & tHLTCode & "', '" & tMBMNo & "', '" & tmpVPolicy & "', '" & tmpVCovID & "', " & _
            "'" & Trim(tmpVIC) & "', '" & tmpVMBMName & "', '" & tDOB & "','" & tmpVSex & "','" & tmpVRaceCode & "', " & _
            "'" & tmpVMBMStatus & "','" & tmpVDataStatus & "', '" & tmpVMemberType & "', " & _
            "'" & tmpVAdd1 & "', '" & tmpVAdd2 & "', '" & tmpVAdd3 & "', " & _
            "'" & tmpVCityCode & "', '" & tmpVPostCode & "', '" & tmpVSTACode & "', " & _
            "'" & tmpVExpStatus & "', '" & tmpVValueDate & "', '" & tEffDate & "', '" & tExpDate & "', " & _
            "'" & tmpVTakeOver & "', '" & tmpVRenewal & "', '" & tmpVBordxDate & "', '" & Trim(tmpVBatchNo) & "', " & _
            "'" & tmpVPlanCode & "', '" & tmpVINSCode & "', '" & tAgeCode & "', '" & tAdultChild & "', " & tmpVAnnLmt & ", " & _
            "'" & tmpVHSClnInd & "', '" & Trim(tmpVTopupNo) & "', '" & Trim(tmpVTopupInd) & "'"
    
    medixmasterconn.Execute MBMWSql

End Sub
Private Sub SaveMBMPrincipalAXA(tmpVMBMName As String, tDOB As String, tmpVPolicy As String, _
tmpVCovID As String, tmpVMBMStatus As String, tmpVDataStatus As String, tmpVExpStatus As String, tmpVIC As String, _
tmpVMemberType As String, tmpVSex As String, tmpVPlanCode As String, tmpVINSCode As String, tmpVEffDate As String, tmpVExpDate As String, _
tAgeCode As String, tAdultChild As String, tmpVAnnLmt As Double, tmpVHSClnInd As String, _
tmpVTopupNo As String, tmpVTopupInd As String, tmpVBordxDate As String, tmpVBatchNo As String)

Dim MBMWSql As String
Dim tmpVAdd1, tmpVAdd2, tmpVAdd3 As String
Dim tmpVRaceCode, tmpVCityCode, tmpVPostCode, tmpVSTACode As String
Dim tmpVTakeOver, tmpVRenewal As String
Dim tmpVValueDate As String
Dim tEffDate, tExpDate As String
    tEffDate = IIf(IsDate(tmpVEffDate), Format(tmpVEffDate, "YYYY/MM/DD"), "")
    tExpDate = IIf(IsDate(tmpVExpDate), Format(tmpVExpDate, "YYYY/MM/DD"), "")

    tmpVAdd1 = ""
    tmpVAdd2 = ""
    tmpVAdd3 = ""
    tmpVCityCode = ""
    tmpVPostCode = ""
    tmpVSTACode = ""
    tmpVTakeOver = Trim(RecordArray(13))
    
    If Trim(RecordArray(14)) = "R" Then
        tmpVRenewal = "R"
    Else
        tmpVRenewal = "N"
    End If
    If tmpVTakeOver = "C" Or tmpVTakeOver = "Y" Or tmpVRenewal = "" Then
        tmpVRenewal = "N"
    End If
    tmpVRaceCode = ""
    tmpVAdd1 = Trim(Replace(tmpVAdd1, "'", "''"))
    tmpVAdd2 = Trim(Replace(tmpVAdd2, "'", "''"))
    tmpVAdd3 = Trim(Replace(tmpVAdd3, "'", "''"))
    tmpVCityCode = Trim(Replace(tmpVCityCode, "'", "''"))
    tmpVSTACode = Trim(Replace(tmpVSTACode, "'", "''"))
    tmpVValueDate = Format(Date, "yyyy/mm/dd")
    MBMWSql = "EXEC Upload_MBM_HIS '" & tPayor & "', '" & tHLTCode & "', '" & tMBMNo & "', '" & tmpVPolicy & "', '" & tmpVCovID & "', " & _
            "'" & Trim(tmpVIC) & "', '" & tmpVMBMName & "', '" & tDOB & "','" & tmpVSex & "','" & tmpVRaceCode & "', " & _
            "'" & tmpVMBMStatus & "','" & tmpVDataStatus & "', '" & tmpVMemberType & "', " & _
            "'" & tmpVAdd1 & "', '" & tmpVAdd2 & "', '" & tmpVAdd3 & "', " & _
            "'" & tmpVCityCode & "', '" & tmpVPostCode & "', '" & tmpVSTACode & "', " & _
            "'" & tmpVExpStatus & "', '" & tmpVValueDate & "', '" & tEffDate & "', '" & tExpDate & "', " & _
            "'" & tmpVTakeOver & "', '" & tmpVRenewal & "', '" & tmpVBordxDate & "', '" & Trim(tmpVBatchNo) & "', " & _
            "'" & tmpVPlanCode & "', '" & tmpVINSCode & "', '" & tAgeCode & "', '" & tAdultChild & "', " & tmpVAnnLmt & ", " & _
            "'" & tmpVHSClnInd & "', '" & Trim(tmpVTopupNo) & "', '" & Trim(tmpVTopupInd) & "'"
    
    medixmasterconn.Execute MBMWSql

End Sub
Private Sub SPSearchMBM(tmpSelFields As String, tmpTableName As String, tmpSelCriteria As String, TmpRecFound As Boolean, tmpCode As String, tmpDesc As String)
Dim TmpRec As New ADODB.Recordset
Dim cmd As New ADODB.Command
Dim Prm1 As New ADODB.Parameter
Dim Prm2 As New ADODB.Parameter
Dim Prm3 As New ADODB.Parameter

    TmpRecFound = False

    Set cmd.ActiveConnection = medixmasterconn
    cmd.CommandText = "SearchHISTableRec"
    cmd.CommandType = adCmdStoredProc
    cmd.CommandTimeout = 300
    Set Prm1 = cmd.CreateParameter("SelFields", adVarChar, adParamInput, 2000, tmpSelFields)
    cmd.Parameters.Append Prm1
    Set Prm2 = cmd.CreateParameter("TableName", adVarChar, adParamInput, 255, tmpTableName)
    cmd.Parameters.Append Prm2
    Set Prm3 = cmd.CreateParameter("SelCriteria", adVarChar, adParamInput, 2000, tmpSelCriteria)
    cmd.Parameters.Append Prm3

    TmpRec.CursorLocation = adUseClient
    TmpRec.CursorType = adOpenForwardOnly
    TmpRec.CacheSize = 100
    Set TmpRec = cmd.Execute
    
    Do Until TmpRec.EOF
        With TmpRec
            tmpCode = IIf(Len(TmpRec!tmpCode), TmpRec!tmpCode, "")
            tmpDesc = IIf(Len(TmpRec!tmpDesc), TmpRec!tmpDesc, "")
            TmpRecFound = True
            TmpRec.MoveNext
            Exit Do
        End With
    Loop
    
    TmpRec.Close
    Set TmpRec = Nothing
End Sub

Private Sub VerifyNewMemberAXA()
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
Dim VNewInsuredCode As String
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
Dim PrincPayEffdate As String
Dim PrincPayExpDate As String
Dim PrincPlanCode As String
Dim tmpPlanCode As String
Dim PrincInsType As String
Dim NoOfRecLine As Integer
Dim tmpTableName As String
Dim TmpRecFound As Boolean
Dim SelFields As String
Dim SelCriteria As String
Dim tmpCode As String
Dim tmpDesc As String
Dim VNewPayor As String
Dim CisprincEffDate As String
Dim CisPrincExpdate As String
Dim CisPrincPlnCode As String
Dim CisPrincInsType As String
Dim VAge As Integer
Dim vPrincipalEffDate As String
Dim VNewRecIndicator As String
Dim recordcnt As Integer
Dim tmpString As String
Dim tmpNewIDType As String
Dim tmpPayorPLN As String
Dim VInsurerCode As String
Dim VInsurerUniqueNo  As String
Dim vNewAdd1 As String
Dim VFWUniqueNo As String
Dim VNewFWPassportNo As String
Dim VEndorseDate As String
Dim VCancelEffDate As String
Dim strsqlFee As String
Dim rstFee As New ADODB.Recordset
Dim Dupfilestatus As Boolean
Dim VNewAXAMBMNo As String
Dim VNewPolOwnerName As String
Dim VInsurerPolicyNo As String
Dim VNewPlanNo As String

lstERRList.ListItems.Clear
    
       
    Screen.MousePointer = vbHourglass
    
    
    'Verify New Member Hash Control
    
    
    If txtImportFilename <> "" Then
        Erase MBMICArray
        ReDim MBMICArray(0)
        MBMICCnt = 0
           
        Erase RecordArray

        Open txtImportFilename For Input As #1   ' Open file.
        Do
        Do While Not EOF(1)  ' Loop until end of file.
           Line Input #1, TextLine   ' Read line into variable.
            NoOfRecLine = NoOfRecLine + 1
                RecordArray() = Split(TextLine, "|")
                    VInsurerCode = "AX"
                    VNewPolicyNo = Trim(RecordArray(0))
                    VNewAXAMBMNo = Trim(RecordArray(1))
                    vNewName = Trim(RecordArray(3))
                    VNewPlanNo = Trim(RecordArray(5))
                    
                    If vNewName = "" Then
                        Call ErrorListing(NoOfRecLine, VNewPolicyNo, vNewName, "Employer Name Is Blank")
                    End If
                    If Trim(VInsurerCode) = "" Then
                        Call ErrorListing(NoOfRecLine, VNewPolicyNo, vNewName, "Insurer Code Is Blank")
                    End If
                    If Trim(VNewAXAMBMNo) = "" Then
                        Call ErrorListing(NoOfRecLine, VNewPolicyNo, vNewName, "Insurer Membership Number Is Blank")
                    End If
                    If Trim(VNewPolicyNo) = "" Then
                        Call ErrorListing(NoOfRecLine, VNewPolicyNo, vNewName, "Policy Number Is Blank")
                    End If
                    If Trim(VNewPlanNo) = "" Then
                        Call ErrorListing(NoOfRecLine, VNewPolicyNo, vNewName, "Plan Number Is Blank")
                    End If
                                            
                DoEvents
                If intExit = 1 Then
                    Check = False
                    Exit Do
                End If
            Loop
            Loop Until Check = False
        Close #1   ' Close file.
        
        End If
End Sub

Private Sub GenerateMemberAXA()
On Err GoTo ErrorSave
Dim startTrans As Boolean
Dim StrSql As String
Dim PayorCode As String
Dim recordCount As Integer
Dim tmpVCovID As String
        tmpFileStatus = ""
        Open txtImportFilename For Input As #1   ' Open file.
        'Line Input #1, TextLine   ' Read line into variable.
        
        'tmpPAYFormat = Mid(TextLine, 28, 3)
        'tmpfilename = Mid(TextLine, 12, 19)
        'If Mid(TextLine, 2, 2) <> Mid(cboPayorCode, 1, 2) And (Trim(tmpPAYFormat) <> "MAA" And Trim(tmpPAYFormat) <> "TAK") Then
        '    MsgBox "Payor Selected Is Not Match The File Payor", vbOKOnly + vbCritical, DialogBoxTitle
        '    cboPayorCode.SetFocus
        'Else
            'If Mid(TextLine, 1, 1) = "F" Then
                PayorCode = "AX"
                txtSubmissionDate = Date
            'ElseIf Mid(TextLine, 1, 1) = "E" Then
            '    txtSubmissionDate = Mid(TextLine, 4, 2) & "/" & Mid(TextLine, 6, 2) & "/" & Mid(TextLine, 8, 4)
            'End If
            
            recordCount = 0
            
            medixmasterconn.beginTrans
            medixmasterconnMaint.beginTrans
            startTrans = True
            'If tmpPAYFormat <> "MAA" And Trim(tmpPAYFormat) <> "TAK" Then
            '    txtBatchNo = CheckBatch(Mid(cboPayorCode, 1, 2), txtSubmissionDate, Mid(cboFileType, 1, 1))
            'Else
                txtBatchNo = CheckBatch(tmpPAYFormat, txtSubmissionDate, Mid(cboFileType, 1, 1))
            'End If
            
        'Do While Not EOF(1)  ' Loop until end of file.
        '   Line Input #1, TextLine   ' Read line into variable.
        '    NoOfRecLine = NoOfRecLine + 1
        '        RecordArray() = Split(TextLine, "|")
                
            Do While Not EOF(1)   ' Loop until end of file.
               Line Input #1, TextLine   ' Read line into variable.
                recordCount = recordCount + 1
                    RecordArray() = Split(TextLine, "|")

                'If Mid(TextLine, 1, 1) <> "F" And Mid(TextLine, 1, 1) <> "T" Then
                    
                    
                    'If tmpPAYFormat <> "MAA" And Trim(tmpPAYFormat) <> "TAK" Then
                        'If Mid(TextLine, 1, 1) = "P" Then
                        '    tmpVCovID = "00"
                        'Else
                        '    tmpVCovID = tmpVCovID + 1
                        'End If
                        'tmpVCovID = Format(tmpVCovID, "00")
                        If Trim(RecordArray(15)) = "Ownself" Then
                            'tmpVDataType = "P"
                            tmpVCovID = "00"
                        Else
                            If IsNumeric(tmpVCovID) = False Then
                                tmpVCovID = "0"
                            End If
                            tmpVCovID = tmpVCovID + 1
                            tmpVCovID = Format(tmpVCovID, "00")
                            'tmpVDataType = "C"
                        End If

                        Call NewMBMInfoProcessingAXA(tmpVCovID)
                    
                    'ElseIf (tmpPAYFormat = "MAA" Or tmpPAYFormat = "TAK") And Mid(TextLine, 1, 12) = "N" Then
                    '        If Mid(TextLine, 1, 1) = "P" Then
                    '            tmpVCovID = "00"
                    '        Else
                    '            tmpVCovID = tmpVCovID + 1
                    '        End If
                    '        tmpVCovID = Format(tmpVCovID, "00")
                    '
                    '    Call NewMBMInfoProcessing(tmpVCovID)
                    '
                    'ElseIf (tmpPAYFormat = "MAA" Or tmpPAYFormat = "TAK") Then
                    '        If Mid(TextLine, 1, 1) = "P" Then
                    '            tmpVCovID = "00"
                    '        Else
                    '            tmpVCovID = tmpVCovID + 1
                    '        End If
                    '        tmpVCovID = Format(tmpVCovID, "00")
                    '
                    '    Call NewMBMInfoProcessing(tmpVCovID)
                    'ElseIf tmpPAYFormat = "MAA" And (Mid(TextLine, 1, 12) = "R" Or Mid(TextLine, 1, 1) <> "P") Then
                    '    Call GenerateENDtMember
                    'End If
                'End If
                DoEvents
            Loop
            
            medixmasterconn.commitTrans
            medixmasterconnMaint.commitTrans
            MsgBox "Record Import Successfully...", vbOKOnly + vbInformation, DialogBoxTitle
            
            startTrans = False
            'Screen.MousePointer = vbDefault
            Close #1
            Exit Sub
        'End If
    Close #1   ' Close file.
ErrorSave:
        Screen.MousePointer = vbDefault
        If startTrans = True Then
            medixmasterconn.RollbackTrans
            medixmasterconnMaint.RollbackTrans
        End If
        MsgBox Err.Description
End Sub

Private Sub NewMBMInfoProcessingAXA(tmpVCovID As String)
'-----
Dim tmpVDataType As String
Dim tmpVPolicy As String
Dim tmpVMBMName As String
Dim tmpVIC As String
Dim tmpVDOB As String
Dim tmpVEffDate As String
Dim tmpVExpDate As String
Dim tmpVINSCode As String
Dim tmpVPlanCode As String
Dim tmpVPolDis As Double
Dim tmpVRenewDis As Double
Dim tmpVGuaRenewal As String
Dim tmpVClmNonPayFr As String
Dim tmpVClmNonPayTo As String

' ----------- cis
Dim tmpCisVPlnCode As String
Dim tmpCisVInsType As String
Dim tmpCisVExpDate As String
Dim tmpCisVEffDate As String
Dim tmpCisVCardType As String
Dim tmpCisVWarranty As String
Dim tmpCLNVStatus As String
Dim tmpVClinical As String

Dim TmpRecFound As Boolean
Dim SelFields As String
Dim SelCriteria As String
Dim tmpCode As String
Dim tmpDesc As String
Dim tmpTableName As String
Dim tmpVAnnLmt As Double
Dim tmpVLifeLmt As Double
Dim RecFound As Boolean
Dim tmpMBMDataType As String
Dim tmpVMemberType As String
Dim tmpVTopupInd As String
Dim tmpVHSClnInd As String
Dim tmpVPremium As Double
Dim tmpVBasicPrem As Double
Dim tmpVMcoAmt As Double
Dim tmpVBasicMco As Double
Dim tmpVSuppBNFStatus As String
Dim tmpVSex As String

Dim tHnCEffDate As String
Dim tHnCExpDate As String
Dim tHnCPlnCode As String
Dim tDOB As String
Dim tAgeCode As String
Dim tFindAge As String

Dim tmpVKidneyAnnLmt As Double
Dim tmpVCancerAnnLmt As Double
Dim tmpVHomeNursAnnLmt As Double
Dim tmpClnStatus As String
Dim tmpClnDAtaStatus As String
Dim tmpClnExpStatus As String
Dim CISBasicPrem As Double
Dim CisPrem As Double
Dim CisMCO As Double
Dim CisBasicMco As Double
Dim CisBalLmt  As Double
Dim CisVisitLmt As Double
Dim TmpCISMBMType As String
Dim CisAnnLmtType As String
Dim CisDentalLmt As Double
Dim CisSpecLmt As Double
Dim CisMepLmt As Double
Dim tmpSuppExpDate As String
Dim tmpSuppCisExpDate As String
Dim tmpVValueDate As String
Dim tmpVBatchNo As String
Dim tmpVExpStatus As String
Dim tmpVMBMStatus As String
Dim tmpVDataStatus As String
Dim tmpVBordxDate As String
Dim tmpSuppEffDate As String
Dim tmpVRenewal As String
Dim tmpCNextLmt As Double
Dim TmpLMBM As New ADODB.Recordset
Dim TmpChkDigits As String
Dim tmpPayorPLN As String
Dim tmpVAXAMBMNo As String
Dim tmpPolicyOwner As String
Dim tmpTransType As String
'Dim tmpVCovID As String

    tmpVTopupNo = ""
    PrevLifeMBMNo = ""
    TmpChkDigits = ""
    tmpPLNPremSexStatus = ""
    'tMBMNo = ""
    'tmpCode = ""
    
    tmpVValueDate = Format(Date, "yyyy/mm/dd")
    tmpVDataStatus = "A"
    tmpVExpStatus = "N"
    If tmpCancelstatus = False Then
        tmpVBordxDate = Format(txtSubmissionDate, "yyyy/mm/dd")
        tmpVMBMStatus = "A"
         tmpVBatchNo = txtBatchNo
    Else
        tmpVBordxDate = Format(txtSubmissionDate, "yyyy/mm/dd")
        tmpVMBMStatus = "C"
        tmpVBatchNo = "1"
    End If
       
    WLifeStatus = False

    If Trim(RecordArray(15)) = "Ownself" Then
        tmpVDataType = "P"
        'tmpVCovID = "00"
    Else
        'If IsNumeric(tmpVCovID) = False Then
        '    tmpVCovID = "0"
        'End If
        'tmpVCovID = tmpVCovID + 1
        'tmpVCovID = Format(tmpVCovID, "00")
        tmpVDataType = "C"
    End If
    
    tHLTCode = Mid(cboHLTCode1, 1, 1)
    'VInsurerCode = "AX"
    '                VNewPolicyNo = Trim(RecordArray(0))
    '                VNewAXAMBMNo = Trim(RecordArray(1))
    '                vNewName = Trim(RecordArray(2))
    '                VNewPlanNo = Trim(RecordArray(5))

    
    tmpVPolicy = Trim(RecordArray(0))
    tmpVPolicy = Trim(Replace(tmpVPolicy, "'", "''"))
    tmpVAXAMBMNo = Trim(RecordArray(1))
    tmpPolicyOwner = Trim(RecordArray(2))
    tmpPolicyOwner = Trim(Replace(tmpPolicyOwner, "'", ""))
    tmpVMBMName = Trim(RecordArray(3))
    tmpVMBMName = Trim(Replace(tmpVMBMName, "'", ""))
    tmpTransType = Trim(RecordArray(14))
    If Trim(tmpTransType) = "N" Or Trim(tmpTransType) = "R" Or Trim(tmpTransType) = "" Then
        tmpVRenewal = tmpTransType
    Else
        tmpVRenewal = ""
    End If
    'If txtBordxFormat = "3" Then
        tmpPayorPLN = Trim(RecordArray(5))
        SelFields = "SELECT PLNCode as tmpCode, PLNLifeTimeStatus as tmpDesc, PLNPayorCode as tmpPLNPayor "
        SelCriteria = " And PLNPayorPlan ='" & Trim(tmpPayorPLN) & "' And PLNPolNo = '" & Trim(tmpVPolicy) & "'"
        tmpTableName = "PLN"
        tmpCode = ""
        tmpDesc = ""
        TmpRecFound = False
        
        Call SPHISTableRecVerify(SelFields, tmpTableName, SelCriteria, TmpRecFound, tmpCode, tmpDesc, tmpStatus, tmpPLNPayor)
            If TmpRecFound = True Then
                tmpVPlanCode = tmpCode
                tPayor = tmpPLNPayor
            Else
                tmpVPlanCode = ""
                tmpPLNPayor = ""
            End If
            
            'tmpVEffDate = Mid(TextLine, 374, 2) & "/" & Mid(TextLine, 376, 2) & "/" & Mid(TextLine, 378, 4)
            'tmpVExpDate = Mid(TextLine, 382, 2) & "/" & Mid(TextLine, 384, 2) & "/" & Mid(TextLine, 386, 4)
            
    'Else
    '    tmpVPlanCode = Trim(Mid(TextLine, 390, 4))
    '    tmpPLNPayor = ""
    '    tPayor = Mid(cboPayorCode, 1, 2)
    'End If

    tmpVIC = Trim(RecordArray(4))
    tmpVIC = Trim(tmpVIC)
    tmpVDOB = Format(Mid(RecordArray(8), 1, 2) & "/" & Mid(RecordArray(8), 4, 2) & "/" & Mid(RecordArray(8), 7, 4), "yyyy/mm/dd")
    tDOB = Format(tmpVDOB, "YYYY/MM/DD")
    tmpVSex = Trim(RecordArray(7))
    tmpVEffDate = Mid(RecordArray(11), 1, 2) & "/" & Mid(RecordArray(11), 4, 2) & "/" & Mid(RecordArray(11), 7, 4)
    tmpVExpDate = Mid(RecordArray(12), 1, 2) & "/" & Mid(RecordArray(12), 4, 2) & "/" & Mid(RecordArray(12), 7, 4)
        
        
    'tmpVPlanCode = Trim(Mid(TextLine, 390, 4))
    'If txtBordxFormat = "3" Then
    '    tmpVINSCode = Trim(Mid(TextLine, 397, 1))
    'Else
        tmpVINSCode = "H"
    'End If
    'tmpVINSCode = Trim(Mid(TextLine, 397, 1))

    tmpVClmNonPayFr = ""
    tmpVClmNonPayFr = ""
    tmpVClmNonPayTo = ""
    tmpVClmNonPayTo = ""
    
    'If txtBordxFormat = "3" Then
        tmpVPolDis = 0
        tmpVRenewDis = 0
        tmpVGuaRenewal = ""
    'Else
    '    tmpVPolDis = IIf(IsNumeric(Trim(Mid(TextLine, 3732, 5))), Trim(Mid(TextLine, 3732, 5)), 0)
    '    tmpVRenewDis = IIf(IsNumeric(Trim(Mid(TextLine, 3737, 5))), Trim(Mid(TextLine, 3737, 5)), 0)
    '    tmpVGuaRenewal = Trim(Mid(TextLine, 3842, 5))
    'End If
    tmpCisVPlnCode = ""
    tmpCisVInsType = ""
    tmpCisVEffDate = ""
    tmpCisVExpDate = ""
    tmpCisVWarranty = ""
    tmpCisVCardType = ""

    tmpCisVWarranty = Trim(Replace(tmpCisVWarranty, "'", "''"))
    tmpCisVCardType = Trim(Replace(tmpCisVCardType, "'", "''"))
'-------

    If tmpVDataType = "P" Then
        tmpVPrinPlnCode = tmpVPlanCode
        tmpVPrinInsCode = tmpVINSCode
        tmpVPrinEffDate = tmpVEffDate
        tmpVPrinExpDate = tmpVExpDate
        tmpMBMDataType = "P"
        tmpVPrinGuaRenew = tmpVGuaRenewal
    Else
        tmpMBMDataType = "C"
    End If

' use principal's data when supplementary data is blank
    tmpVPlanCode = IIf(Trim(tmpVPlanCode) = "", tmpVPrinPlnCode, tmpVPlanCode)
    tmpVINSCode = IIf(Trim(tmpVINSCode) = "", tmpVPrinInsCode, tmpVINSCode)
    tmpVEffDate = IIf(IsDate(tmpVEffDate), tmpVEffDate, tmpVPrinEffDate)
    tmpVExpDate = IIf(IsDate(tmpVExpDate), tmpVExpDate, tmpVPrinExpDate)
    tmpSuppEffDate = IIf(IsDate(tmpVEffDate), Format(tmpVEffDate, "yyyy/mm/dd"), "")
    tHnCPlnCode = tmpVPlanCode
    tHncInsType = tmpVINSCode
    tHnCEffDate = tmpVEffDate
    tHnCExpDate = tmpVExpDate
    
    If Trim(tHnCPlnCode) = "" Then
        tHnCPlnCode = tmpCisVPlnCode
        tHncInsType = tmpCisVInsType
        tHnCEffDate = tmpCisVEffDate
        tHnCExpDate = tmpCisVExpDate
    End If

    tCAnnSuppLmtStatus = "A"
    
    If Trim(tmpVPlanCode) <> "" Then
        Call SPHISDBTableRecAssign("PLN", tmpVPlanCode, tmpVINSCode, CDate(tmpVEffDate))
        tAnnSuppLmtStatus = "A"
        If Trim(tmpPLNGuaRenewal) <> "" Then
            tmpVGuaRenewal = tmpPLNGuaRenewal
        Else
            tmpVGuaRenewal = IIf(Trim(tmpVGuaRenewal) = "", tmpVPrinGuaRenew, tmpVGuaRenewal)
        End If
        
        If tPLNAnnLmtIND = "Y" Then
            Call SPHISDBTableRecAssign("AnnualLimit", tmpVPlanCode, tmpVINSCode, CDate(tmpVEffDate))
        End If
    End If
    If tmpCisVPlnCode <> "" Then
        Call SPCisDBTableRecAssign("PLN", tmpCisVPlnCode, tmpCisVInsType, CDate(tmpCisVEffDate))
        If tCPLNAnnLmtInd = "Y" Then
            Call SPCisDBTableRecAssign("AnnualLimit", tmpCisVPlnCode, tmpCisVInsType, CDate(tmpCisVEffDate))
        End If
    End If
    tmpVMemberType = "P"
    If Trim(tAnnSuppLmtStatus) <> "A" Then
        tmpVMemberType = "Q"
    End If
    tFindAge = GetAge(CDate(tHnCEffDate), CDate(tmpVDOB))
    tAgeCode = Mid(GetAgeBand(tHLTCode, tFindAge, tHnCPlnCode, tHncInsType), 1, 2)
    tAdultChild = "A"
    If tAgeCode = "01" Then
        tAdultChild = "C"
    End If

'------------------------------------------------------------------------------------------------------------------------
    tmpVTopupInd = ""
    
    If Trim(tmpVPlanCode) <> "" Then 'only for HIS prem, mco, annLmt, Access Calc.
        tmpVBasicPrem = 0
        tmpVPremium = 0
        If tPLNPrmIND = "Y" Then
            tmpVPremium = Format(GetHISPremium(tmpVINSCode, tPayor, tAgeCode, tmpVPlanCode, tHLTCode, CDate(tmpVEffDate), tmpMBMDataType, tmpVCovID, tmpPLNPremSexStatus, tmpVSex), "#,##0.00")
            tmpVBasicPrem = tmpVPremium
            If tmpVPolDis > 0 Then
                tmpVPremium = tmpVPremium * (100 - tmpVPolDis) / 100
            End If
            If tmpVRenewDis > 0 Then
                tmpVPremium = tmpVPremium * (100 - tmpVRenewDis) / 100
            End If
            
            tmpVPremium = InsertPremium(CDate(tmpVExpDate), CDate(tmpVEffDate), Format(tmpVPremium, "#,##0.00"))
            tmpVPremium = Format(tmpVPremium, "#,##0.00")
            
        End If
        tmpVBasicMco = 0
        tmpVMcoAmt = 0
        If tPLNMcoIND = "Y" Then
            tmpVMcoAmt = Format(GetHISMCO(tmpVINSCode, tHLTCode, tAdultChild, CDate(tmpVEffDate), tmpVPlanCode, tmpMBMDataType, tmpVCovID), "#,##0.00")
            tmpVBasicMco = tmpVMcoAmt
            If tmpProRate <> "N" Then
                tmpVMcoAmt = InsertMCOFees(CDate(tmpVExpDate), CDate(tmpVEffDate), Format(tmpVBasicMco, "#,##0.00"))
            End If
            tmpVMcoAmt = Format(tmpVMcoAmt, "#,##0.00")
            If TmpRoundUpStatus <> "N" Then
                tmpVMcoAmt = Round(tmpVMcoAmt)
            End If
        End If
        tmpVAnnLmt = 0
        If tPLNAnnLmtIND = "Y" Then
            tmpVAnnLmt = 0
            tmpVAnnLmt = GetHISAnnLmt(tmpVINSCode, tmpVPlanCode, tHLTCode, CDate(tmpVEffDate), tmpMBMDataType, tmpVCovID)
            tmpVAnnLmt = CDbl(tmpVAnnLmt)
            tmpCNextLmt = tmpVAnnLmt
        End If
        tmpVLifeLmt = 0
        
        If Trim(tmpPLNCancerLTInd) = "Y" Then
            Call GetHISLTKidneyCancer(tmpVINSCode, tmpVPlanCode, tHLTCode, CDate(tmpVEffDate), tmpMBMDataType, tmpVCovID)
        End If
        
' ---------- write Previous MBM no. to new MBMTopupMember field (only for HIS)
        If Trim(tPLNTopupStatus) = "Y" Then
            SelFields = "SELECT TOP 1 MBMNumber as tmpCode, MBMTopupNumber as tmpDesc "
            SelCriteria = " AND MBMICBCPP = '" & Trim(tmpVIC) & _
                        "' And MBMStatus = 'A' AND PayCode ='" & tPayor & "' AND INSCode='" & Trim(tmpVINSCode) & _
                        "' ORDER BY MBMPayorEffDate DESC, MBMNumber "
            tmpTableName = "MBMCrossReference"
            tmpCode = ""
            tmpDesc = ""
            TmpRecFound = False
            Call SPMaintTableRecVerify(SelFields, tmpTableName, SelCriteria, TmpRecFound, tmpCode, tmpDesc)
            If TmpRecFound = True Then
                tmpVTopupNo = tmpCode
                tmpVTopupInd = "Y"
            End If
        End If
'-------- end of Topup
    End If   'if trim(tmpVPlanCode)<> ""
    If tmpVDataType = "P" Then
        If Trim(tmpVPlanCode) = "" And Trim(tmpCisVPlnCode) <> "" Then
            'tMBMNo = tPayor & tmpCisVPlnCode & tmpCisVInsType & GenerateMembershipNo(tPayor, tmpCisVPlnCode, tmpCisVInsType)
            tMBMNo = tPayor & tmpCisVInsType & GenerateMembershipNo(tPayor, tmpCisVInsType)
            TmpChkDigits = GenerateChkDigits
            tMBMNo = tMBMNo & TmpChkDigits
        ElseIf Trim(tmpVPlanCode) <> "" And Trim(tmpCisVPlnCode) <> "" Then
            If txtPayorMBMNoType = "P" Then
                tMBMNo = tPayor & GenerateMBMNoPol(tPayor, tmpVINSCode, tmpVPolicy, tmpVRenewal, tmpFileStatus)
            Else
                tMBMNo = tPayor & GenerateMBMNo05(tPayor, tmpVINSCode, tDOB, tmpVMBMName, tmpVRenewal)
            End If
            'TmpChkDigits = GenerateChkDigits
            tMBMNo = tMBMNo & TmpChkDigits
        Else
            If txtPayorMBMNoType = "P" And tPayor = "XL" Then
                tMBMNo = tPayor & GenerateMBMNoPol(tPayor, tmpVINSCode, tmpVPolicy, tmpVRenewal, tmpFileStatus)
            Else
                tMBMNo = tPayor & GenerateMBMNo05(tPayor, tmpVINSCode, tDOB, tmpVMBMName, tmpVRenewal)
            End If
            
            tMBMNo = tMBMNo & TmpChkDigits
        End If
    End If
    tmpVHSClnInd = "H"
    If Trim(tmpVPlanCode) = "" And Trim(tmpCisVPlnCode) <> "" Then
        tmpVHSClnInd = "C"
    End If
    If Trim(tmpVPlanCode) <> "" And Trim(tmpCisVPlnCode) = "" Then
        tmpVHSClnInd = "H"
    End If
    If Trim(tmpCisVPlnCode) <> "" And Trim(tmpVPlanCode) <> "" Then
        If Trim(tCPlnDCStatus) = "Y" Then
            tmpVHSClnInd = "M"
        Else
            tmpVHSClnInd = "B"
        End If
    End If
    tmpVClinical = "N"
    If Trim(tmpCisVPlnCode) <> "" Then
         tmpVClinical = "Y"
    End If

' ----- LifeTimeLimit & BenefitLimit
    If Trim(tmpVGuaRenewal) <> "" Then
        If Trim(tPLNLifeTimeStatus) <> "Y" Then
            tmpVTopupNo = ""
        Else
            tmpVKidneyAnnLmt = 0
            tmpVCancerAnnLmt = 0
            tmpVHomeNursAnnLmt = 0
            If Trim(tmpPLNCancerLTInd) <> "Y" Then
                Call GetBenefitLimit(tmpVINSCode, tmpVPlanCode, tHLTCode, CDate(tmpVEffDate), tmpVKidneyAnnLmt, tmpVCancerAnnLmt, tmpVHomeNursAnnLmt, tmpVSuppBNFStatus)
                If (tmpVSuppBNFStatus = "A" And tmpVCovID = "00") Or _
                   (tmpVSuppBNFStatus = "B" And tmpVCovID = "01") Or _
                   tmpVSuppBNFStatus = "C" Then
                    Call SaveMBMBnfAnnLmt(tMBMNo, tmpVCovID, tmpVSuppBNFStatus, tmpVKidneyAnnLmt, tmpVCancerAnnLmt, tmpVHomeNursAnnLmt)
                End If
            End If
            tmpCNextLmt = tmpVAnnLmt
            If tmpVDataType = "P" Then
            '    If Trim(tPayor) = "XG" Then
            '        SelCriteria = "Select top 1 MBMNumber, MBMTopupNumber, MBMAvailableLimit From MBM where Substring(MBMPolicyNo,1,12)='" & Mid(tmpVPolicy, 1, 12) & _
            '              "' And MBMStatus = 'A' AND PayCode ='" & Trim(tPayor) & "' AND INSCode='" & Trim(tmpVINSCode) & _
            '              "' ORDER BY MBMPayorEffDate DESC "
            '        TmpLMBM.Open SelCriteria, medixmasterconn, adOpenForwardOnly, adLockReadOnly
            '    Else
                    SelCriteria = "Select top 1 MBMNumber, MBMTopupNumber, MBMAvailableLimit From MBM where MBMPolicyNo='" & Trim(tmpVPolicy) & _
                          "' And MBMStatus = 'A' AND PayCode ='" & Trim(tPayor) & "' AND INSCode='" & Trim(tmpVINSCode) & _
                          "' ORDER BY MBMPayorEffDate DESC "
                    TmpLMBM.Open SelCriteria, medixmasterconn, adOpenForwardOnly, adLockReadOnly
            '    End If
                tmpVTopupNo = tMBMNo
                PrevLifeMBMNo = tMBMNo
                Do While Not TmpLMBM.EOF
                    tmpVTopupNo = TmpLMBM!MBMTopupNumber
                    PrevLifeMBMNo = TmpLMBM!MBMNumber
                    tmpCNextLmt = IIf(IsNumeric(TmpLMBM!MBMAvailableLimit), TmpLMBM!MBMAvailableLimit, 0)
                    TmpLMBM.MoveNext
                Loop
                TmpLMBM.Close
                Set TmpLMBM = Nothing
            Else
                SelFields = "SELECT TOP 1 MBMCNumber as tmpCode, MBMCAvailableLimit as tmpDesc "
                SelCriteria = " AND MBMCNumber='" & Trim(PrevLifeMBMNo) & _
                          "' And MBMCName = '" & tmpVMBMName & "'"
                tmpTableName = "MBMCoveredPersons"
                tmpCode = ""
                tmpDesc = ""
                TmpRecFound = False
                Call SPSearchMBM(SelFields, tmpTableName, SelCriteria, TmpRecFound, tmpCode, tmpDesc)
                If TmpRecFound = True Then
                    tmpCNextLmt = IIf(IsNumeric(tmpDesc), tmpDesc, 0)
                End If
            End If
            If Trim(tmpVGuaRenewal) <> "3" Then
                tmpVAnnLmt = tmpCNextLmt
            End If
                
            SelFields = "Select MBMNumber as tmpCode, MBMLifeTimeLmt as tmpDesc "
            SelCriteria = " AND MBMNumber='" & Trim(tmpVTopupNo) & _
                      "' AND MBMCovID='" & Trim(tmpVCovID) & "'"
            tmpTableName = "MBMLifeTime"
            tmpCode = ""
            tmpDesc = ""
            TmpRecFound = False
            Call SPSearchMBM(SelFields, tmpTableName, SelCriteria, TmpRecFound, tmpCode, tmpDesc)
            If TmpRecFound = False Then
                If Trim(tmpVGuaRenewal) = "3" Then
                    WLifeStatus = True
                End If
            End If
            If WLifeStatus = True Then
                If (tAnnSuppLmtStatus = "A" And tmpVCovID = "00") Or (tAnnSuppLmtStatus = "B" And tmpVCovID = "01") Or tAnnSuppLmtStatus = "C" Then
                    Call SaveMBMLifeTimeLmt(tmpVTopupNo, tmpVCovID, tmpVDataType, tmpVPlanCode, tmpVINSCode, CDate(tmpVEffDate), tmpVAnnLmt, tmpCNextLmt, tmpVGuaRenewal)
                End If
            End If
            
            If Trim(tmpPLNCancerLTInd) = "Y" Then
                SelFields = ""
                SelCriteria = ""
                
                SelFields = "Select MBMNumber as tmpCode, MBMCancerLTLmt as tmpDesc "
                SelCriteria = " AND MBMNumber='" & Trim(tmpVTopupNo) & _
                          "' AND MBMCoverID='" & Trim(tmpVCovID) & "'"
                tmpTableName = "MBMBnfLTLmt"
                tmpCode = ""
                tmpDesc = ""
                TmpRecFound = False
                
                Call SPSearchMBM(SelFields, tmpTableName, SelCriteria, TmpRecFound, tmpCode, tmpDesc)
                If TmpRecFound = True Then
                    Call SaveMBMBnfLTLmt(tmpVTopupNo, tmpVCovID, tmpKidneyLTLmt, tmpCancerLTLmt)
                End If
            End If

        End If
    End If
' ----- end of LifeTimeLimit & BenefitLimit
' ---- CLINICAL
    If Trim(tmpCisVPlnCode) <> "" Then
          
        tmpClnStatus = "A"
        tmpClnDAtaStatus = "A"
        tmpClnExpStatus = "N"
        CISBasicPrem = 0
        CisPrem = 0
        If tCPLNPRMInd = "Y" Then
              CisPrem = Format(GetCISPremium(tmpCisVInsType, tCPayor, tAgeCode, tmpCisVPlnCode, tHLTCode, CDate(tmpCisVExpDate), tmpMBMDataType, tmpVCovID), "#,##0.00")
              CISBasicPrem = CisPrem
              CisPrem = InsertPremium(CDate(tmpCisVExpDate), CDate(tmpCisVEffDate), Format(CisPrem, "#,##0.00"))
              CisPrem = Format(Round(CisPrem, 0), "#,##0.00")
        End If
        CisMCO = 0
        CisBasicMco = 0
        If tCPLNMCOInd = "Y" Then
            CisMCO = Format(GetCisMCO(tmpCisVInsType, tHLTCode, tAdultChild, CDate(tmpCisVEffDate), tmpCisVPlnCode, tmpMBMDataType, tmpVCovID), "#,##0.00")
            CisBasicMco = CisMCO
            'CisMCO = InsertMCOFees(CDate(tmpCisVExpDate), CDate(tmpCisVEffDate), Format(CisMCO, "#,##0.00"))
            'CisMCO = Round(CisMCO)
            If tCProRate <> "N" Then
                CisMCO = InsertMCOFees(CDate(tmpCisVExpDate), CDate(tmpCisVEffDate), Format(CisMCO, "#,##0.00"))
            End If
            CisMCO = Format(CisMCO, "#,##0.00")
            If tCRoundUpStatus <> "N" Then
                CisMCO = Round(CisMCO)
            End If
        End If
        
        CisBalLmt = 0
        If tCPLNAnnLmtInd = "Y" Then
            CisBalLmt = GetCISAnnLmt(tmpCisVInsType, tmpCisVPlnCode, tHLTCode, CDate(tmpCisVEffDate), tmpMBMDataType, tmpVCovID)
        End If
        CisVisitLmt = 0
        If tCPLNVsTLmtInd = "Y" Then
            CisVisitLmt = GetVisitLimit(tmpCisVInsType, tmpCisVPlnCode, tHLTCode, CDate(tmpCisVEffDate))
        End If
        If Trim(tCAnnSuppLmtStatus) = "A" Then
            TmpCISMBMType = "P"
        Else
            TmpCISMBMType = "Q"
        End If
        
        CisDentalLmt = 0
        CisSpecLmt = 0
        CisMepLmt = 0
        
        If Trim(tCPlnPDentStatus) = "S" Then 'stand alone dental limit
            CisAnnLmtType = "PDentalAnnLimit"
            CisDentalLmt = GetCISLimit(tmpCisVInsType, tmpCisVPlnCode, tHLTCode, CDate(tmpCisVEffDate), CisAnnLmtType)
        End If
        If Trim(tCPlnPSpecStatus) = "S" Then 'stand alone Spec limit
            CisAnnLmtType = "PSpecAnnLimit"
            CisSpecLmt = GetCISLimit(tmpCisVInsType, tmpCisVPlnCode, tHLTCode, CDate(tmpCisVEffDate), CisAnnLmtType)
        End If
        If Trim(tCPlnPMepStatus) = "S" Then 'stand alone MEP limit
            CisAnnLmtType = "PMepAnnLimit"
            CisMepLmt = GetCISLimit(tmpCisVInsType, tmpCisVPlnCode, tHLTCode, CDate(tmpCisVEffDate), CisAnnLmtType)
        End If
        
    End If
    tmpVPremium = CDbl(IIf(IsNumeric(tmpVPremium) = True, tmpVPremium, 0))
    tmpVBasicPrem = CDbl(IIf(IsNumeric(tmpVBasicPrem) = True, tmpVBasicPrem, 0))
    tmpVBasicMco = CDbl(IIf(IsNumeric(tmpVBasicMco) = True, tmpVBasicMco, 0))
    tmpVMcoAmt = CDbl(IIf(IsNumeric(tmpVMcoAmt) = True, tmpVMcoAmt, 0))
    tmpVAnnLmt = CDbl(IIf(IsNumeric(tmpVAnnLmt) = True, tmpVAnnLmt, 0))

    CisPrem = CDbl(IIf(IsNumeric(CisPrem) = True, CisPrem, 0))
    CISBasicPrem = CDbl(IIf(IsNumeric(CISBasicPrem) = True, CISBasicPrem, 0))
    CisBasicMco = CDbl(IIf(IsNumeric(CisBasicMco) = True, CisBasicMco, 0))
    CisMCO = CDbl(IIf(IsNumeric(CisMCO) = True, CisMCO, 0))
    CisBalLmt = CDbl(IIf(IsNumeric(CisBalLmt) = True, CisBalLmt, 0))
    CisVisitLmt = CDbl(IIf(IsNumeric(CisVisitLmt) = True, CisVisitLmt, 0))
    CisDentalLmt = CDbl(IIf(IsNumeric(CisDentalLmt) = True, CisDentalLmt, 0))
    CisSpecLmt = CDbl(IIf(IsNumeric(CisSpecLmt) = True, CisSpecLmt, 0))
    CisMepLmt = CDbl(IIf(IsNumeric(CisMepLmt) = True, CisMepLmt, 0))
    'tmpVMBMStatus = IIf(Trim(tmpVPlanCode) <> "", tmpVMBMStatus, "")
    tmpCisVWarranty = Replace(tmpCisVWarranty, "'", "''")

' ---- END OF CLINICAL
    
    If tmpVDataType = "P" Then
        Call SaveMBMPrincipalAXA(tmpVMBMName, tDOB, tmpVPolicy, tmpVCovID, tmpVMBMStatus, tmpVDataStatus, tmpVExpStatus, tmpVIC, _
             tmpVMemberType, tmpVSex, tmpVPlanCode, tmpVINSCode, tmpVEffDate, tmpVExpDate, _
             tAgeCode, tAdultChild, tmpVAnnLmt, tmpVHSClnInd, tmpVTopupNo, tmpVTopupInd, tmpVBordxDate, tmpVBatchNo) ' done
        Call SaveMBMTwoAXA(tmpVClinical)  'DONE
        Call SaveMBMXRefAXA(tMBMNo, "00", Trim(tmpVPolicy), tmpVIC, tmpVMBMName, tmpVEffDate, _
             tmpVExpDate, tAgeCode, tmpVINSCode, tmpVPlanCode, tmpVBordxDate, tmpVTopupNo, _
             tmpVTopupInd, tmpVDataStatus, tmpVMBMStatus, tmpVDOB, tHLTCode, tPayor, tmpVBatchNo)  'done
        Call SaveMBMOthersAXA(tmpVDataType, tmpVPremium, tmpVBasicPrem, tmpVBasicMco, tmpVMcoAmt, tmpVPolDis, tmpVRenewDis, tmpVClmNonPayFr, tmpVClmNonPayTo, tmpVEffDate, tmpVGuaRenewal, tmpVAXAMBMNo) 'done
        Call SaveMBMOthersTwoAXA
        If Trim(tmpCisVPlnCode) <> "" Then
            Call SaveMBMCLNOthersAXA(TmpCISMBMType, tCPayor, tmpVPolicy, tmpClnStatus, _
                 tmpClnDAtaStatus, tmpClnExpStatus, tmpCisVPlnCode, tmpCisVInsType, tmpCisVWarranty, _
                 CISBasicPrem, CisPrem, CisBasicMco, CisMCO, CisBalLmt, CisVisitLmt, _
                 CisDentalLmt, CisSpecLmt, CisMepLmt, tmpCisVEffDate, tmpCisVExpDate, tmpCisVCardType) 'done
        End If
    ElseIf tmpVDataType <> "R" Then
        tmpSuppExpDate = ""
        If Trim(tmpVPlanCode) <> "" Then
            tmpSuppExpDate = Format(GetSuppExpDate(CDate(tmpVDOB), tmpVExpDate, tPLNAgeLimit), "YYYY/MM/DD")
        End If
        Call SaveMBMCovPersonAXA(tmpVDataType, tmpVCovID, tDOB, tmpMBMDataType, tmpVMBMName, _
             tmpVMBMStatus, tmpVBordxDate, tmpVBatchNo, tmpVIC, tmpVSex, _
             tmpVPlanCode, tAgeCode, tAdultChild, tmpVAnnLmt, _
             tmpVBasicMco, tmpVMcoAmt, tmpVBasicPrem, tmpVPremium, tmpSuppEffDate, tmpSuppExpDate, tmpVAXAMBMNo) 'done
        Call SaveMBMCovPersonTwoAXA(tmpVCovID, tmpVClmNonPayFr, tmpVClmNonPayTo, tmpVPolDis, tmpVRenewDis) 'done
        If Trim(tmpCisVPlnCode) <> "" Then
            tmpSuppExpDate = ""
            tmpSuppCisExpDate = Format(GetSuppExpDate(CDate(tmpVDOB), tmpCisVExpDate, tCPLNAgeLimit), "YYYY/MM/DD")
            Call SaveMBMCovPersonCLNAXA(TmpCISMBMType, tmpVCovID, tmpClnStatus, _
                 tmpCisVPlnCode, tmpCisVWarranty, CISBasicPrem, CisPrem, _
                 CisBasicMco, CisMCO, CisBalLmt, CisVisitLmt, _
                 CisDentalLmt, CisSpecLmt, CisMepLmt, tmpSuppCisExpDate, _
                 tmpCisVWarranty, tmpCisVCardType, tmpCisVPlnCode)
        End If
    End If
    
    If intExit = 1 Then
    '    Check = False
    End If

End Sub





