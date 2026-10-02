VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form FrmNewMBMEndt 
   Caption         =   "Upload Member"
   ClientHeight    =   7980
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   11370
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   7980
   ScaleWidth      =   11370
   WindowState     =   2  'Maximized
   Begin VB.Frame Frame4 
      Height          =   615
      Left            =   120
      TabIndex        =   11
      Top             =   7200
      Width           =   11175
      Begin VB.CommandButton cmdExit 
         Cancel          =   -1  'True
         Caption         =   "&Exit"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   204
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   9480
         TabIndex        =   19
         Top             =   240
         Width           =   960
      End
      Begin VB.CommandButton CmdPrint 
         Caption         =   "&Print Err List"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   204
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   8400
         TabIndex        =   18
         Top             =   240
         Width           =   1080
      End
      Begin VB.CommandButton cmdStop 
         Caption         =   "Stop"
         Height          =   315
         Left            =   7440
         TabIndex        =   17
         Top             =   240
         Width           =   960
      End
      Begin VB.TextBox txtTempMemberNo 
         Height          =   285
         Left            =   960
         TabIndex        =   16
         Text            =   "txtTempMemberNo"
         Top             =   240
         Visible         =   0   'False
         Width           =   375
      End
      Begin VB.TextBox txtMBMAdultChild 
         Height          =   285
         Left            =   5400
         TabIndex        =   15
         Top             =   240
         Visible         =   0   'False
         Width           =   375
      End
      Begin VB.TextBox txtAccessFees 
         Height          =   285
         Left            =   4680
         TabIndex        =   14
         Top             =   240
         Visible         =   0   'False
         Width           =   735
      End
      Begin VB.TextBox txtMBMMCOFees 
         Height          =   285
         Left            =   3960
         TabIndex        =   13
         Top             =   240
         Visible         =   0   'False
         Width           =   735
      End
      Begin VB.TextBox txtMBMPremium 
         Height          =   285
         Left            =   3240
         TabIndex        =   12
         Top             =   240
         Visible         =   0   'False
         Width           =   735
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
      TabIndex        =   20
      Top             =   240
      Width           =   11175
      _ExtentX        =   19711
      _ExtentY        =   12303
      _Version        =   393216
      Tabs            =   2
      TabsPerRow      =   2
      TabHeight       =   520
      TabCaption(0)   =   "Upload Bordx"
      TabPicture(0)   =   "FrmNewMBMEndt.frx":0000
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "lstERRList"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Frame1"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "Frame2"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).ControlCount=   3
      TabCaption(1)   =   "Generate ASCII File"
      TabPicture(1)   =   "FrmNewMBMEndt.frx":001C
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Frame8"
      Tab(1).ControlCount=   1
      Begin VB.Frame Frame8 
         Height          =   3375
         Left            =   -74760
         TabIndex        =   44
         Top             =   360
         Width           =   10815
         Begin VB.Frame Frame14 
            Height          =   2535
            Left            =   120
            TabIndex        =   50
            Top             =   120
            Width           =   10455
            Begin VB.CommandButton cmdGenerate 
               Caption         =   "&Ok"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   8.25
                  Charset         =   204
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   315
               Left            =   9360
               TabIndex        =   76
               Top             =   1920
               Width           =   960
            End
            Begin VB.Frame Frame16 
               Height          =   1695
               Left            =   120
               TabIndex        =   61
               Top             =   120
               Width           =   10215
               Begin VB.ComboBox cboASCIIFileVer 
                  BeginProperty Font 
                     Name            =   "Arial"
                     Size            =   8.25
                     Charset         =   204
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   330
                  Left            =   1440
                  Sorted          =   -1  'True
                  TabIndex        =   65
                  Top             =   1200
                  Width           =   2025
               End
               Begin VB.ComboBox cboExportOption 
                  BeginProperty Font 
                     Name            =   "Arial"
                     Size            =   8.25
                     Charset         =   204
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   330
                  ItemData        =   "FrmNewMBMEndt.frx":0038
                  Left            =   4800
                  List            =   "FrmNewMBMEndt.frx":003A
                  Sorted          =   -1  'True
                  TabIndex        =   64
                  Text            =   "B - Batch Bordx"
                  Top             =   1200
                  Width           =   2025
               End
               Begin VB.ComboBox cboHLTCode 
                  BeginProperty Font 
                     Name            =   "Arial"
                     Size            =   8.25
                     Charset         =   204
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   330
                  Left            =   1440
                  Sorted          =   -1  'True
                  TabIndex        =   63
                  Text            =   "S - SIHAT MALAYSIA"
                  Top             =   720
                  Width           =   2025
               End
               Begin VB.ComboBox cboPAYCode 
                  BeginProperty Font 
                     Name            =   "Arial"
                     Size            =   8.25
                     Charset         =   204
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   330
                  Left            =   4800
                  Sorted          =   -1  'True
                  Style           =   2  'Dropdown List
                  TabIndex        =   62
                  Top             =   720
                  Width           =   5295
               End
               Begin VB.Label Label24 
                  Caption         =   "File on Server"
                  BeginProperty Font 
                     Name            =   "Arial"
                     Size            =   8.25
                     Charset         =   204
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   225
                  Left            =   120
                  TabIndex        =   75
                  Top             =   240
                  Width           =   1125
               End
               Begin VB.Label Label36 
                  Caption         =   "File Version"
                  BeginProperty Font 
                     Name            =   "Arial"
                     Size            =   8.25
                     Charset         =   204
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   255
                  Left            =   120
                  TabIndex        =   74
                  Top             =   1200
                  Width           =   1365
               End
               Begin VB.Label Label21 
                  Caption         =   "Bordx Option"
                  BeginProperty Font 
                     Name            =   "Arial"
                     Size            =   8.25
                     Charset         =   204
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   255
                  Left            =   3720
                  TabIndex        =   73
                  Top             =   1200
                  Width           =   1365
               End
               Begin VB.Label Label19 
                  Caption         =   "Health Type"
                  BeginProperty Font 
                     Name            =   "Arial"
                     Size            =   8.25
                     Charset         =   204
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   255
                  Left            =   120
                  TabIndex        =   72
                  Top             =   720
                  Width           =   1365
               End
               Begin VB.Label Label18 
                  Caption         =   "Payor"
                  BeginProperty Font 
                     Name            =   "Arial"
                     Size            =   8.25
                     Charset         =   204
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   255
                  Left            =   3720
                  TabIndex        =   71
                  Top             =   720
                  Width           =   915
               End
               Begin VB.Label txtPath 
                  BackColor       =   &H00FFFFFF&
                  BeginProperty Font 
                     Name            =   "Arial"
                     Size            =   8.25
                     Charset         =   204
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   285
                  Left            =   1440
                  TabIndex        =   70
                  Top             =   240
                  Width           =   2025
               End
               Begin VB.Label txtLocalFileLocation 
                  BackColor       =   &H00FFFFFF&
                  BeginProperty Font 
                     Name            =   "Arial"
                     Size            =   8.25
                     Charset         =   204
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   285
                  Left            =   4800
                  TabIndex        =   69
                  Top             =   240
                  Width           =   1905
               End
               Begin VB.Label txtFilName 
                  BackColor       =   &H00FFFFFF&
                  Height          =   255
                  Left            =   8160
                  TabIndex        =   68
                  Top             =   240
                  Width           =   1905
               End
               Begin VB.Label Label28 
                  Caption         =   "File on PC"
                  BeginProperty Font 
                     Name            =   "Arial"
                     Size            =   8.25
                     Charset         =   204
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   225
                  Left            =   3720
                  TabIndex        =   67
                  Top             =   240
                  Width           =   765
               End
               Begin VB.Label Label25 
                  Caption         =   "Export File Name"
                  BeginProperty Font 
                     Name            =   "Arial"
                     Size            =   8.25
                     Charset         =   204
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   255
                  Left            =   6840
                  TabIndex        =   66
                  Top             =   270
                  Width           =   1245
               End
            End
            Begin VB.Frame BatchFrame 
               Height          =   600
               Left            =   120
               TabIndex        =   56
               Top             =   1800
               Width           =   9015
               Begin VB.TextBox txtBatchNo1 
                  Height          =   285
                  Left            =   5640
                  TabIndex        =   57
                  Top             =   180
                  Width           =   2500
               End
               Begin MSMask.MaskEdBox txtReceiptDate 
                  Height          =   285
                  Left            =   1530
                  TabIndex        =   58
                  Top             =   180
                  Width           =   2505
                  _ExtentX        =   4419
                  _ExtentY        =   503
                  _Version        =   393216
                  MaxLength       =   10
                  BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                     Name            =   "Tahoma"
                     Size            =   8.25
                     Charset         =   204
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
               Begin VB.Label Label27 
                  Caption         =   "Batch No"
                  BeginProperty Font 
                     Name            =   "Arial"
                     Size            =   8.25
                     Charset         =   204
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   255
                  Left            =   4680
                  TabIndex        =   60
                  Top             =   180
                  Width           =   675
               End
               Begin VB.Label Label26 
                  Caption         =   "Receipt Date From"
                  BeginProperty Font 
                     Name            =   "Arial"
                     Size            =   8.25
                     Charset         =   204
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   255
                  Left            =   90
                  TabIndex        =   59
                  Top             =   225
                  Width           =   1515
               End
            End
            Begin VB.Frame BordxFrame 
               Height          =   600
               Left            =   120
               TabIndex        =   51
               Top             =   1800
               Visible         =   0   'False
               Width           =   9015
               Begin MSMask.MaskEdBox txtBordxTo 
                  Height          =   285
                  Left            =   5730
                  TabIndex        =   52
                  Top             =   180
                  Width           =   2505
                  _ExtentX        =   4419
                  _ExtentY        =   503
                  _Version        =   393216
                  MaxLength       =   10
                  BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                     Name            =   "Tahoma"
                     Size            =   8.25
                     Charset         =   204
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
                  TabIndex        =   53
                  Top             =   180
                  Width           =   2505
                  _ExtentX        =   4419
                  _ExtentY        =   503
                  _Version        =   393216
                  MaxLength       =   10
                  BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                     Name            =   "Tahoma"
                     Size            =   8.25
                     Charset         =   204
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Mask            =   "##/##/####"
                  PromptChar      =   "_"
               End
               Begin VB.Label Label22 
                  Caption         =   "Bordx Date From"
                  BeginProperty Font 
                     Name            =   "Arial"
                     Size            =   8.25
                     Charset         =   204
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   255
                  Left            =   90
                  TabIndex        =   55
                  Top             =   225
                  Width           =   1275
               End
               Begin VB.Label Label23 
                  Caption         =   "Bordx Date To"
                  BeginProperty Font 
                     Name            =   "Arial"
                     Size            =   8.25
                     Charset         =   204
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   255
                  Left            =   4290
                  TabIndex        =   54
                  Top             =   180
                  Width           =   1035
               End
            End
         End
         Begin VB.Frame Frame7 
            Caption         =   "Status"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   204
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   570
            Left            =   120
            TabIndex        =   45
            Top             =   2640
            Width           =   10500
            Begin VB.Label Label16 
               Alignment       =   2  'Center
               Caption         =   "Total Supplementary"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   8.25
                  Charset         =   204
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H000000FF&
               Height          =   225
               Left            =   2640
               TabIndex        =   49
               Top             =   240
               Width           =   2040
            End
            Begin VB.Label txtTotalPrincipals 
               Alignment       =   2  'Center
               BackColor       =   &H00C0FFFF&
               Caption         =   "0"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   8.25
                  Charset         =   204
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   270
               Left            =   1200
               TabIndex        =   48
               Top             =   240
               Width           =   1320
            End
            Begin VB.Label txtTotalSupplementary1 
               Alignment       =   2  'Center
               BackColor       =   &H00C0FFFF&
               Caption         =   "0"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   8.25
                  Charset         =   204
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   270
               Left            =   4920
               TabIndex        =   47
               Top             =   240
               Width           =   1200
            End
            Begin VB.Label Label17 
               Alignment       =   2  'Center
               Caption         =   "Total Principal"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   8.25
                  Charset         =   204
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H000000FF&
               Height          =   225
               Left            =   120
               TabIndex        =   46
               Top             =   240
               Width           =   1080
            End
         End
      End
      Begin VB.Frame Frame2 
         Caption         =   "Status"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   204
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   120
         TabIndex        =   33
         Top             =   6240
         Width           =   10785
         Begin VB.Label txtBatchNo 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00C0FFFF&
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   204
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   270
            Left            =   9720
            TabIndex        =   43
            Top             =   240
            Width           =   720
         End
         Begin VB.Label Label20 
            Caption         =   "Bat No"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   204
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000FF&
            Height          =   255
            Left            =   9120
            TabIndex        =   42
            Top             =   240
            Width           =   495
         End
         Begin VB.Label Label6 
            Alignment       =   2  'Center
            Caption         =   "Total Errors"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   204
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000FF&
            Height          =   225
            Left            =   2760
            TabIndex        =   41
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
               Charset         =   204
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
         Begin VB.Label Label2 
            Alignment       =   2  'Center
            Caption         =   "Submission Date"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   204
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000FF&
            Height          =   225
            Left            =   120
            TabIndex        =   39
            Top             =   240
            Width           =   1425
         End
         Begin VB.Label txtSubmissionDate 
            Alignment       =   2  'Center
            BackColor       =   &H00C0FFFF&
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   204
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   270
            Left            =   1560
            TabIndex        =   38
            Top             =   240
            Width           =   1200
         End
         Begin VB.Label Label7 
            Alignment       =   2  'Center
            Caption         =   "Total Principal"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   204
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000FF&
            Height          =   225
            Left            =   4680
            TabIndex        =   37
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
               Charset         =   204
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   270
            Left            =   5880
            TabIndex        =   36
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
               Charset         =   204
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   270
            Left            =   7800
            TabIndex        =   35
            Top             =   240
            Width           =   960
         End
         Begin VB.Label Label11 
            Alignment       =   2  'Center
            Caption         =   "Total Supp"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   204
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000FF&
            Height          =   225
            Left            =   6840
            TabIndex        =   34
            Top             =   240
            Width           =   960
         End
      End
      Begin VB.Frame Frame1 
         Height          =   1815
         Left            =   120
         TabIndex        =   21
         Top             =   360
         Width           =   10935
         Begin VB.TextBox txtMBMCancelStatus 
            Height          =   285
            Left            =   9600
            TabIndex        =   80
            Text            =   "N"
            Top             =   840
            Visible         =   0   'False
            Width           =   615
         End
         Begin VB.TextBox txtClnStatus 
            Height          =   285
            Left            =   9600
            TabIndex        =   23
            Text            =   "N"
            Top             =   480
            Visible         =   0   'False
            Width           =   615
         End
         Begin VB.CommandButton cmdFileLocation 
            Caption         =   " >"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   204
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   4240
            TabIndex        =   5
            Top             =   1395
            Width           =   320
         End
         Begin VB.Frame Frame13 
            Height          =   615
            Left            =   5520
            TabIndex        =   22
            Top             =   1120
            Width           =   2175
            Begin VB.CommandButton cmdUpload 
               Caption         =   "&Upload"
               Enabled         =   0   'False
               Height          =   315
               Left            =   1080
               TabIndex        =   10
               Top             =   200
               Width           =   960
            End
            Begin VB.CommandButton cmdVerify 
               Caption         =   "&Verify"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   8.25
                  Charset         =   204
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
         End
         Begin VB.ComboBox cboHLTCode1 
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   204
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
         Begin VB.ComboBox cboPayorCode 
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   204
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
         Begin VB.ComboBox cboFileType 
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   204
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
         Begin VB.TextBox txtFileVersion 
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   204
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
               Charset         =   204
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   1515
            TabIndex        =   4
            Top             =   1365
            Width           =   2625
         End
         Begin VB.ComboBox cboVerifyIC 
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   204
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   330
            Left            =   5880
            TabIndex        =   6
            Text            =   "Y - Yes"
            Top             =   240
            Width           =   1695
         End
         Begin VB.ComboBox CboBack2Back 
            Height          =   315
            Left            =   5880
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
               Charset         =   204
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   330
            Left            =   5880
            Sorted          =   -1  'True
            Style           =   2  'Dropdown List
            TabIndex        =   8
            Top             =   830
            Width           =   1695
         End
         Begin VB.Label Label8 
            Caption         =   "MBM Cancel Status"
            Height          =   375
            Index           =   1
            Left            =   8040
            TabIndex        =   79
            Top             =   840
            Visible         =   0   'False
            Width           =   1575
         End
         Begin VB.Label Label8 
            Caption         =   "Clinical Status"
            Height          =   375
            Index           =   0
            Left            =   8040
            TabIndex        =   32
            Top             =   480
            Visible         =   0   'False
            Width           =   1335
         End
         Begin VB.Label Label15 
            Caption         =   "BTB Grantee"
            Height          =   255
            Left            =   4800
            TabIndex        =   31
            Top             =   570
            Width           =   1215
         End
         Begin VB.Label Label14 
            Caption         =   "Label Status"
            Height          =   255
            Left            =   4800
            TabIndex        =   30
            Top             =   840
            Width           =   975
         End
         Begin VB.Label Label13 
            Caption         =   "Verify I/C"
            Height          =   255
            Left            =   4800
            TabIndex        =   29
            Top             =   255
            Width           =   855
         End
         Begin VB.Label Label10 
            Caption         =   "Health Type"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   204
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   240
            TabIndex        =   28
            Top             =   135
            Width           =   1770
         End
         Begin VB.Label Label5 
            Caption         =   "New/ Endt File"
            Height          =   255
            Left            =   240
            TabIndex        =   27
            Top             =   825
            Width           =   1095
         End
         Begin VB.Label Label1 
            Caption         =   "Filename"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   204
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   240
            TabIndex        =   26
            Top             =   1440
            Width           =   975
         End
         Begin VB.Label Label55 
            Caption         =   "Payor"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   204
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   240
            TabIndex        =   25
            Top             =   480
            Width           =   1170
         End
         Begin VB.Label Label4 
            Caption         =   "File Version"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   204
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   240
            TabIndex        =   24
            Top             =   1140
            Width           =   960
         End
      End
      Begin MSComctlLib.ListView lstERRList 
         Height          =   3975
         Left            =   120
         TabIndex        =   77
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
            Charset         =   204
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
         Charset         =   204
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000005&
      Height          =   220
      Left            =   0
      TabIndex        =   78
      Top             =   0
      Width           =   11415
   End
End
Attribute VB_Name = "FrmNewMBMEndt"
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
Dim tMBMNo As String
Dim tHncInsType As String
Dim tPLNTopupStatus As String
Dim tPLNPRMInd As String
Dim tPLNMCOInd As String
Dim tPLNAnnLmtInd As String
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
Dim tCPlnPDenatlLmt, tCPlnSDentalLmt As String
Dim tCPlnPSpecLmt, tCPlnSSpecLmt As String
Dim tCAnnSuppLmtStatus As String
Dim PrevLifeMBMNo As String
Dim TmpRoundUpStatus As String
Dim tmpEndtPrincEffDate As String
Dim tmpEndtPrincExpDate As String
Dim AscMBM As New ADODB.Recordset
Dim DataString As String
Dim BasicPrem As Double

Public intExit As Integer

Private Sub cboASCIIFileVer_KeyPress(KeyAscii As Integer)
    Call ComboKeyPress(cboASCIIFileVer, KeyAscii)
End Sub

Private Sub CboHLTCode_KeyPress(KeyAscii As Integer)
    Call ComboKeyPress(cboHLTCode, KeyAscii)
End Sub

Private Sub cboHLTCode_LostFocus()
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
                Exportsql = "HLTCode = '" & Trim(Mid(cboHLTCode, 1, 1)) & _
                            "' And PAYCode = '" & Trim(Mid(cboPAYCode, 1, 2)) & "'" & _
                            " And MBMStatus <> 'D' " & _
                            " And MBMDataStatus <> 'D' " & _
                            " And MBMBordxDate >=  '" & Format(Trim(txtBordxFrom), "mm/dd/yyyy") & "'" & _
                            " And MBMBordxDate <= '" & Format(Trim(txtBordxTo), "mm/dd/yyyy") & "'"
            End If
            strsql = "SELECT MBMNumber, MBMMemberType, INSCode, MBMPolicyNo, MBMPayorEffDate, MBMIcBcPp, " & _
                    " MBMName From MBM Where " & Exportsql
            medixmasterconn.Open "File Name=" & BOSPath & "Conn\mediconnect.UDL"

            AscMBM.Open strsql, medixmasterconn, adOpenKeyset, adLockOptimistic

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
        XPpaySeqNo = IIf(Len(Trim(MBMOther2!MBMPaySeqNo)), MBMOther2!MBMPaySeqNo, "")
        XPClientNo = IIf(Len(Trim(MBMOther2!MBMClientNo)), MBMOther2!MBMClientNo, "")
        XPClientID = IIf(Len(Trim(MBMOther2!MBMClientID)), MBMOther2!MBMClientID, "")
        
        MBMOther2.Close
        Set MBMOther2 = Nothing
        DoEvents
        
        DataString = DataString & l(MBMPolicyNo, 25)
        DataString = DataString & l(MBMEmployeeNo, 15)
        DataString = DataString & l(AscMBM!MBMName, 60)
        DataString = DataString & l(AscMBM!MBMICBcPp, 20)
        DataString = DataString & l(AscMBM!MBMNumber, 18)
        
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
        If Trim(AscMBM!INSCode) = "F" Or Trim(AscMBM!INSCode) = "H" Then
            Call SupplementaryASCII(MNumber, MBMPayEffDate, MBMPolSubNo1, MBMEmployeeNo, MBMPolicyNo)
        End If
        
        AscMBM.MoveNext
    Loop
End Sub

Private Sub SupplementaryASCII(MNumber As String, MBMPayEffDate As String, MBMPolSubNo1 As String, MBMEmployeeNo As String, MBMPolicyNo As String)
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
        
        DataString = DataString & l(Trim(MBMPolicyNo), 25)
        DataString = DataString & l(Trim(MBMEmployeeNo), 15)
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
        cboFileType.SetFocus
    ElseIf Trim(cboFileType) = "" Then
        MsgBox "Please Select New / Endt file", vbOKOnly + vbCritical, DialogBoxTitle
        cboFileType.SetFocus
    ElseIf Trim(cboPayorCode) = "" Then
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

            If Mid(cboFileType, 1, 1) = "N" Then
                Call GenerateMember
            ElseIf Mid(cboFileType, 1, 1) = "E" Then
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
    If Mid(cboFileType, 1, 1) = "" Then
        MsgBox "Please Select Bordx Type", vbOKOnly + vbCritical, DialogBoxTitle
        cboFileType.SetFocus
    ElseIf cboPayorCode = "" Then
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
        lstERRList.listitems.Clear

        If Mid(cboFileType, 1, 1) = "N" Then
            Call VerifyNewMember
        ElseIf Mid(cboFileType, 1, 1) = "E" Then
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
                If Trim(Mid(TextLine, 2, 2)) = "" Or Trim(Mid(TextLine, 2, 2)) <> Mid(cboPayorCode, 1, 2) Then
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
                    PrincInsType = Trim(Mid(TextLine, 394, 1))
                Else
                    VNewPlanCode = IIf(Trim(VNewPlanCode) = "", PrincPlanCode, VNewPlanCode)
                    VNewInsured = IIf(Trim(VNewInsured) = "", PrincInsType, VNewInsured)
                    If VNewPlanCode <> "" Then
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
                    SelFields = "SELECT PLNCode as tmpCode, PLNTopUpStatus as tmpDesc "
                    SelCriteria = " And PLNCode ='" & Trim(VNewPlanCode) & "'"
                    tmpTableName = "PLN"
                    tmpCode = ""
                    tmpDesc = ""
                    TmpRecFound = False
                    Call SPHISTableRecVerify(SelFields, tmpTableName, SelCriteria, TmpRecFound, tmpCode, tmpDesc)
                    If TmpRecFound = False Then
                        Call ErrorListing(NoOfRecLine, VNewPolicyNo, vNewName, "No such Plan Code - ' " & Trim(VNewPlanCode) & "' !")
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
                            If IsDate(VNewDOB) Then
                                Call GetSuppAgeLmt(NoOfRecLine, VNewDOB, "HIS", VNewPolicyNo, vNewName, VNewDataType, VNewPlanCode, VNewInsured, CDate(VNewEffDate), VNewExpDate)
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
                        Call SPHISTableRecVerify(SelFields, tmpTableName, SelCriteria, TmpRecFound, tmpCode, tmpDesc)
                        If TmpRecFound = False Then
                            Call ErrorListing(NoOfERecLine, VEndtPolicyNo, vEndtName, "No such Plan Code - ' " & Trim(VEndtPlanCode) & "' !")
                        End If
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
Dim tmpVCovID As String

        Open txtImportFilename For Input As #1   ' Open file.
        Line Input #1, TextLine   ' Read line into variable.
        If Mid(TextLine, 2, 2) <> Mid(cboPayorCode, 1, 2) Then
            MsgBox "Payor Selected Is Not Match The File Payor", vbOKOnly + vbCritical, DialogBoxTitle
            cboPayorCode.SetFocus
        Else
            If Mid(TextLine, 1, 1) = "F" Then
                PayorCode = Mid(TextLine, 2, 2)
                txtSubmissionDate = Mid(TextLine, 4, 2) & "/" & Mid(TextLine, 6, 2) & "/" & Mid(TextLine, 8, 4)
            End If
            
            recordCount = 1
            
            medixmasterconn.beginTrans
            medixmasterconnMaint.beginTrans
            startTrans = True
            txtBatchNo = CheckBatch(Mid(cboPayorCode, 1, 2), txtSubmissionDate, Mid(cboFileType, 1, 1))
            Do While Not EOF(1)   ' Loop until end of file.
               Line Input #1, TextLine   ' Read line into variable.
                recordCount = recordCount + 1
            
                If Mid(TextLine, 1, 1) <> "F" And Mid(TextLine, 1, 1) <> "T" Then
                    
                    If Mid(TextLine, 1, 1) = "P" Then
                        tmpVCovID = "00"
                    Else
                        tmpVCovID = tmpVCovID + 1
                    End If
                    tmpVCovID = Format(tmpVCovID, "00")

                    Call NewMBMInfoProcessing(tmpVCovID)
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


Private Sub GenerateENDtMember()
Dim tmpVEndtType As String
    Open txtImportFilename For Input As #1   ' Open file.
        Line Input #1, TextLine   ' Read line into variable.
        If Mid(TextLine, 2, 2) <> Mid(cboPayorCode, 1, 2) Then
            MsgBox "Payor Selected Is Not Match The File Payor", vbOKOnly + vbCritical, DialogBoxTitle
            cboPayorCode.SetFocus
        Else
            Do While Not EOF(1)   ' Loop until end of file.
                Line Input #1, TextLine   ' Read line into variable.
    
                tmpVEndtType = Mid(TextLine, 1, 1)
                If tmpVEndtType = "C" Then
                    If Mid(TextLine, 10, 1) = "P" Then
                        Call CancelENDtData
                    End If
                End If
                DoEvents
            Loop
            MsgBox "Record Import Sucessfully...", vbOKOnly + vbInformation, DialogBoxTitle
        End If
    Close #1   ' Close file.
End Sub

Private Sub CancelENDtData()
Dim tmpMBMNo As String
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
                     Call EndtMBMRefund(Trim(tmpMBMNo), CDate(tmpEndtPrincEffDate), CDate(TmpExpDate), CDate(CancelDate), !MBMCCoverID, CInt(tmpCPrem), CInt(tmpCMco), txtMBMCancelStatus)
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
Dim tmpName As String

    strsql = "Select MBMNumber, MBMStatus, MBMLastEndorseStatus, MBMENDtBordxDate,MBMENDtEffDate, " & _
             "MBMPolicyNo, MBMICBcPp, MBMName, MBMPayorEXPDate, PLNCode, " & _
             "MBMPayorEffDate from MBM Where MBMNumber ='" & Trim(tmpMBMNo) & "'"
    MBM.Open strsql, medixmasterconn, adOpenKeyset, adLockOptimistic

    If MBM.recordCount Then
    
    tmpName = Trim(Replace(tmpName, "'", "''"))
        StrAmendSql = "Insert into MBMAmendHistory(MBMNumber, MBMPolicyNo, MBMICBcPp, MBMName, " & _
            "MBMExpiryDate, MBMPlanCode, MBMAmendEdtType, " & _
            "MBMAEndtBordxDate, MBMAEndtEffDate, " & _
            "MBMAmendUser, MBMAmendDate, MBMAEndtBatchNo) " & _
            "Values ('" & MBM!MBMNumber & "','" & MBM!MBMPolicyNo & "','" & MBM!MBMICBcPp & "','" & tmpName & "'," & _
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
    Call EndtMBMRefund(Trim(tmpMBMNo), CDate(tmpEndtPrincEffDate), CDate(tmpEndtPrincExpDate), CDate(CancelDate), "00", CInt(MBMPremium), CInt(MBMMCOFees), txtMBMCancelStatus)
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
              "," & MCORefundValue & ",'" & Mid(cboPayorCode, 1, 2) & "'," & _
              "'" & Format(txtSubmissionDate, "yyyy/mm/dd") & "','" & txtBatchNo & "','" & HLTCode & "'"
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
Dim strsql As String
Dim LastCoveredID As Integer
Dim SelectAge As String
Dim MBMCov As New ADODB.Recordset
Dim AgeValue As String
Dim tmpPlanCode As String
Dim tmpInsType As String

    tmpPlanCode = Trim(Mid(TextLine, 417, 4))
    tmpInsType = Trim(Mid(TextLine, 421, 1))
    strsql = "Select * from MBMCoveredPersons Where MBMCNumber = '" & Trim(Mid(TextLine, 11, 18)) & "'"
    MBMCov.Open strsql, medixmasterconn, adOpenKeyset, adLockOptimistic

    LastCoveredID = GetLastCoveredID(Trim(Mid(TextLine, 11, 18)))
        MBMCov.AddNew
        With MBMCov
            !MBMCNumber = Trim(Mid(TextLine, 11, 18))
            !MBMCRELCode = Mid(TextLine, 10, 1)
            SelectAge = GetAge(Mid(TextLine, 401, 2) & "/" & Mid(TextLine, 403, 2) & "/" & Mid(TextLine, 405, 4), Mid(TextLine, 320, 2) & "/" & Mid(TextLine, 322, 2) & "/" & Mid(TextLine, 324, 4))
            AgeValue = Mid(GetAgeBand(Mid(cboHLTCode1, 1, 1), SelectAge, tmpPlanCode, tmpInsType), 1, 2)
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
Dim tmpPlanCode As String
Dim tmpInsType As String

    tmpPlanCode = Trim(Mid(TextLine, 417, 4))
    tmpInsType = Trim(Mid(TextLine, 421, 1))

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
            AgeValue = Mid(GetAgeBand(Mid(cboHLTCode1, 1, 1), FindAge, tmpPlanCode, tmpInsType), 1, 2)
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
            !PAYCode = Mid(cboPayorCode, 1, 2)
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
            AgeValue = Mid(GetAgeBand(Mid(cboHLTCode1, 1, 1), FindAge, plancode, InsuredCode), 1, 2)
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
Dim tmpPlanCode As String
Dim tmpInsType As String

    tmpPlanCode = Trim(Mid(TextLine, 417, 4))
    tmpInsType = Trim(Mid(TextLine, 421, 1))
        
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
        AgeValue = Mid(GetAgeBand(Mid(cboHLTCode1, 1, 1), FindAge, tmpPlanCode, tmpInsType), 1, 2)
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

Private Function GetSuppExpDate(DateOfBirth, TmpExpDate As String, PlnAgeLmt As String)
Dim CheckYr As Long
Dim SuppDobExpDate As String

    SuppDobExpDate = Format(DateOfBirth, "dd/mm") & "/" & Year(TmpExpDate)
    CheckYr = Year(TmpExpDate) - Year(DateOfBirth)
    GetSuppExpDate = TmpExpDate
    If Trim(PlnAgeLmt) <> "" Then
        If CheckYr = PlnAgeLmt Then
            If CDate(SuppDobExpDate) <= CDate(TmpExpDate) Then
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

Private Sub cboPayorCode_LostFocus()
Dim tmpTableName As String
Dim SelFields As String
Dim SelCriteria As String
Dim tmpCode As String
Dim tmpDesc As String
Dim TmpRecFound  As Boolean
    SelFields = "SELECT CLNStatus as tmpCode, MBMCancelStatus as tmpDesc "
    SelCriteria = " And PAYCode ='" & Mid(cboPayorCode, 1, 2) & "'"
    tmpTableName = "PAY"
    tmpCode = "N"
    tmpDesc = "N"
    TmpRecFound = False
    Call SPHISTableRecVerify(SelFields, tmpTableName, SelCriteria, TmpRecFound, tmpCode, tmpDesc)
    txtClnStatus = IIf(Len(tmpCode), tmpCode, "N")
    txtMBMCancelStatus = IIf(Len(tmpDesc), tmpDesc, "N")
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

Private Sub SaveMBMLifeTimeLmt(LMBMNumber As String, LMBMCovID As String, LmtType As String, tmpPlncode As String, tmpInsType As String, TmpEDate As Date)
Dim tmpSql As String
Dim TmpLifeSql As String
Dim TmpLimit As String
Dim TmpLmtVar As String
Dim cmd As New ADODB.Command
Dim AccumAmtRec As New ADODB.Recordset
Dim Prm1 As New ADODB.Parameter
Dim Prm2 As New ADODB.Parameter
Dim TmpSql2 As Integer
Dim tmpAccAmt As Double
    
    tmpAccAmt = 0
    
    TmpSql2 = 0
    If Trim(LMBMCovID) <> "00" Then
        TmpSql2 = 1
    End If
    tmpSql = " And MBMTopupNumber ='" & Trim(LMBMNumber) & "' AND MBMPatientCovID='" & Trim(LMBMCovID) & "'"
    Set cmd.ActiveConnection = medixmasterconn
    cmd.CommandText = "SearchAccAppAmt"
    cmd.CommandType = adCmdStoredProc
    cmd.CommandTimeout = 15
    Set Prm1 = cmd.CreateParameter("strAppend", adVarChar, adParamInput, 255, tmpSql)
    cmd.Parameters.Append Prm1
    Set Prm2 = cmd.CreateParameter("strAppendCase", adVarChar, adParamInput, 255, TmpSql2)
    cmd.Parameters.Append Prm2
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
    If LmtType = "P" Then
        TmpLmtVar = "LifeTimeLimit"
    End If
    TmpLimit = GetHISLifeTimeLmt(tmpInsType, tmpPlncode, tHLTCode, CDate(TmpEDate), LmtType, LMBMCovID)
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

Private Sub LoadComboListing(tmpSelFields As String, tmpTableName As String, tmpSelCriteria As String, tmpComboBox As ComboBox)
Dim TmpREc As New ADODB.Recordset
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

    TmpREc.CursorLocation = adUseClient
    TmpREc.CursorType = adOpenForwardOnly
    TmpREc.CacheSize = 100
    Set TmpREc = cmd.Execute
    tmpComboBox.Clear
    Do Until TmpREc.EOF
        With TmpREc
            tmpComboBox.AddItem !tmpCode & " - " & !tmpName
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
Dim Prm1 As New ADODB.Parameter
Dim Prm2 As New ADODB.Parameter
Dim Prm3 As New ADODB.Parameter

    TmpRecFound = False
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
Dim Prm1 As New ADODB.Parameter
Dim Prm2 As New ADODB.Parameter
Dim Prm3 As New ADODB.Parameter

    TmpRecFound = False
    medixmasterconnMaint.Open "File Name=" & BOSPath & "Conn\MediConnHISMaint.UDL"

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
    medixmasterconnMaint.Close
    Set medixmasterconnMaint = Nothing
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
    
    medixmasterconnMaint.Open "File Name=" & BOSPath & "Conn\MediConnHISMaint.UDL"
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

' -----------  cis
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
Dim tmpVTopupNo As String
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
Dim tAdultChild As String
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
    
    tmpVValueDate = Format(Date, "yyyy/mm/dd")
    tmpVDataStatus = "A"
    tmpVExpStatus = "N"
    tmpVBordxDate = Format(txtSubmissionDate, "yyyy/mm/dd")
    tmpVMBMStatus = "A"
    tmpVBatchNo = txtBatchNo

    WLifeStatus = False
    tPayor = Mid(cboPayorCode, 1, 2)
    tHLTCode = Mid(cboHLTCode1, 1, 1)
    tmpVDataType = Mid(TextLine, 1, 1)
    tmpVPolicy = Trim(Mid(TextLine, 2, 25))
    tmpVMBMName = Trim(Mid(TextLine, 47, 60))
    tmpVIC = Trim(Mid(TextLine, 273, 20))
    tmpVDOB = Format(Mid(TextLine, 293, 2) & "/" & Mid(TextLine, 295, 2) & "/" & Mid(TextLine, 297, 4), "yyyy/mm/dd")
    tDOB = Format(tmpVDOB, "YYYY/MM/DD")
    tmpVSex = Trim(Mid(TextLine, 301, 1))
    tmpVEffDate = Mid(TextLine, 374, 2) & "/" & Mid(TextLine, 376, 2) & "/" & Mid(TextLine, 378, 4)
    tmpVExpDate = Mid(TextLine, 382, 2) & "/" & Mid(TextLine, 384, 2) & "/" & Mid(TextLine, 386, 4)
        
    tmpVPlanCode = Trim(Mid(TextLine, 390, 4))
    tmpVINSCode = Trim(Mid(TextLine, 394, 1))

    tmpVClmNonPayFr = Mid(TextLine, 3576, 2) & "/" & Mid(TextLine, 3578, 2) & "/" & Mid(TextLine, 3580, 4)
    tmpVClmNonPayFr = IIf(IsDate(tmpVClmNonPayFr), Format(tmpVClmNonPayFr, "YYYY/MM/DD"), "")
    tmpVClmNonPayTo = Mid(TextLine, 3584, 2) & "/" & Mid(TextLine, 3586, 2) & "/" & Mid(TextLine, 3588, 4)
    tmpVClmNonPayTo = IIf(IsDate(tmpVClmNonPayTo), Format(tmpVClmNonPayTo, "YYYY/MM/DD"), "")
    tmpVPolDis = IIf(IsNumeric(Trim(Mid(TextLine, 3732, 5))), Trim(Mid(TextLine, 3732, 5)), 0)
    tmpVRenewDis = IIf(IsNumeric(Trim(Mid(TextLine, 3737, 5))), Trim(Mid(TextLine, 3737, 5)), 0)
    tmpVGuaRenewal = Trim(Mid(TextLine, 3842, 5))
    
    tmpCisVPlnCode = Trim(Mid(TextLine, 6708, 4))
    tmpCisVInsType = Trim(Mid(TextLine, 6712, 1))
    tmpCisVEffDate = Mid(TextLine, 6713, 2) & "/" & Mid(TextLine, 6715, 2) & "/" & Mid(TextLine, 6717, 4)
    tmpCisVExpDate = Mid(TextLine, 6721, 2) & "/" & Mid(TextLine, 6723, 2) & "/" & Mid(TextLine, 6725, 4)
    tmpCisVWarranty = Trim(Mid(TextLine, 6729, 2500))
    tmpCisVCardType = Trim(Mid(TextLine, 9229, 20))

    tmpVPolicy = Trim(Replace(tmpVPolicy, "'", "''"))
    tmpVMBMName = Trim(Replace(tmpVMBMName, "'", "''"))
    tmpCisVWarranty = Trim(Replace(tmpCisVWarranty, "'", "''"))
    tmpCisVCardType = Trim(Replace(tmpCisVCardType, "'", "''"))
'-------

    If tmpVDataType = "P" Then
        tmpVPrinPlnCode = tmpVPlanCode
        tmpVPrinInsCode = tmpVINSCode
        tmpVPrinEffDate = tmpVEffDate
        tmpVPrinExpDate = tmpVExpDate
        tmpMBMDataType = "P"
    Else
        tmpMBMDataType = "C"
    End If

' use principal's data when supplementary data is blank
    tmpVPlanCode = IIf(Trim(tmpVPlanCode) = "", tmpVPrinPlnCode, tmpVPlanCode)
    tmpVINSCode = IIf(Trim(tmpVINSCode) = "", tmpVPrinInsCode, tmpVINSCode)
    tmpVEffDate = IIf(IsDate(tmpVEffDate), tmpVEffDate, tmpVPrinEffDate)
    tmpVExpDate = IIf(IsDate(tmpVExpDate), tmpVExpDate, tmpVPrinExpDate)
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

    tmpVTopupNo = " "
    tCAnnSuppLmtStatus = "A"
    
    If Trim(tmpVPlanCode) <> "" Then
        Call SPHISDBTableRecAssign("PLN", tmpVPlanCode, tmpVINSCode, CDate(tmpVEffDate))
        Call SPHISDBTableRecAssign("AnnualLimit", tmpVPlanCode, tmpVINSCode, CDate(tmpVEffDate))
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
    tmpVTopupNo = " "
    tmpVTopupInd = ""
    
    If Trim(tmpVPlanCode) <> "" Then 'only for HIS prem, mco, annLmt, Access Calc.
        tmpVBasicPrem = 0
        tmpVPremium = 0
        If tPLNPRMInd = "Y" Then
            tmpVPremium = Format(GetHISPremium(tmpVINSCode, tPayor, tAgeCode, tmpVPlanCode, tHLTCode, CDate(tmpVEffDate), tmpMBMDataType, tmpVCovID), "#,##0.00")
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
        If tPLNMCOInd = "Y" Then
            tmpVMcoAmt = Format(GetHISMCO(tmpVINSCode, tHLTCode, tAdultChild, CDate(tmpVEffDate), tmpVPlanCode, tmpMBMDataType, tmpVCovID), "#,##0.00")
            tmpVBasicMco = tmpVMcoAmt
            tmpVMcoAmt = InsertMCOFees(CDate(tmpVExpDate), CDate(tmpVEffDate), Format(tmpVBasicMco, "#,##0.00"))
            tmpVMcoAmt = Format(tmpVMcoAmt, "#,##0.00")
            If TmpRoundUpStatus <> "N" Then
                tmpVMcoAmt = Round(tmpVMcoAmt)
            End If
        End If
        tmpVAnnLmt = 0
        If tPLNAnnLmtInd = "Y" Then
            tmpVAnnLmt = 0
            tmpVAnnLmt = GetHISAnnLmt(tmpVINSCode, tmpVPlanCode, tHLTCode, CDate(tmpVEffDate), tmpMBMDataType, tmpVCovID)
            tmpVAnnLmt = CDbl(tmpVAnnLmt)
        End If
        tmpVLifeLmt = 0
        
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
            tMBMNo = tPayor & tmpCisVPlnCode & tmpCisVInsType & GenerateMembershipNo(tPayor, tmpCisVPlnCode, tmpCisVInsType)
        Else
            tMBMNo = tPayor & tmpVPlanCode & tmpVINSCode & GenerateMembershipNo(tPayor, tmpVPlanCode, tmpVINSCode)
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
    PrevLifeMBMNo = ""
    If Trim(tmpVGuaRenewal) <> "" Then
        If Trim(tPLNLifeTimeStatus) = "Y" Then
            tmpVKidneyAnnLmt = 0
            tmpVCancerAnnLmt = 0
            tmpVHomeNursAnnLmt = 0
            Call GetBenefitLimit(tmpVINSCode, tmpVPlanCode, tHLTCode, CDate(tmpVEffDate), tmpVKidneyAnnLmt, tmpVCancerAnnLmt, tmpVHomeNursAnnLmt, tmpVSuppBNFStatus)
            If (tmpVSuppBNFStatus = "A" And tmpVCovID = "00") Or _
               (tmpVSuppBNFStatus = "B" And tmpVCovID = "01") Or _
               tmpVSuppBNFStatus = "C" Then
                Call SaveMBMBnfAnnLmt(tMBMNo, tmpVCovID, tmpVSuppBNFStatus, tmpVKidneyAnnLmt, tmpVCancerAnnLmt, tmpVHomeNursAnnLmt)
            End If
            SelFields = "SELECT TOP 1 MBMNumber as tmpCode, MBMAvailableLimit as tmpDesc "
            SelCriteria = " AND MBMPolicyNo='" & Trim(tmpVPolicy) & _
                      "' And MBMStatus = 'A' AND PayCode ='" & tPayor & "' AND INSCode='" & Trim(tmpVINSCode) & _
                      "' ORDER BY MBMPayorEffDate , MBMNumber "
            tmpTableName = "MBM"
            tmpCode = ""
            tmpDesc = ""
            TmpRecFound = False
            Call SPSearchMBM(SelFields, tmpTableName, SelCriteria, TmpRecFound, tmpCode, tmpDesc)
            tmpVTopupNo = tMBMNo
            If TmpRecFound = True Then
                tmpVTopupNo = tmpCode
                PrevLifeMBMNo = tmpVTopupNo
                If Trim(tmpVGuaRenewal) <> "3" Then
                    tmpVAnnLmt = IIf(IsNumeric(tmpDesc), tmpDesc, 0)
                End If
            End If
            SelFields = "Select MBMNumber as tmpCode, MBMLifeTimeLmt as tmpDesc "
            SelCriteria = " AND MBMNumber='" & Trim(PrevLifeMBMNo) & _
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
                Call SaveMBMLifeTimeLmt(tmpVTopupNo, tmpVCovID, tmpMBMDataType, tmpVPlanCode, tmpVINSCode, CDate(tmpVEffDate))
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
            CisMCO = InsertMCOFees(CDate(tmpCisVExpDate), CDate(tmpCisVEffDate), Format(CisMCO, "#,##0.00"))
            CisMCO = Round(CisMCO)
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
    Else
        tmpSuppExpDate = ""
        If Trim(tmpVPlanCode) <> "" Then
            tmpSuppExpDate = Format(GetSuppExpDate(CDate(tmpVDOB), tmpVExpDate, tPLNAgeLimit), "YYYY/MM/DD")
        End If
        Call SaveMBMCovPerson(tmpVDataType, tmpVCovID, tDOB, tmpMBMDataType, tmpVMBMName, _
             tmpVMBMStatus, tmpVBordxDate, tmpVBatchNo, tmpVIC, tmpVSex, tmpVPolDis, _
             tmpVRenewDis, tmpVPlanCode, tAgeCode, tAdultChild, tmpVAnnLmt, _
             tmpVBasicMco, tmpVMcoAmt, tmpVBasicPrem, tmpVPremium, tmpSuppExpDate, _
             tmpVClmNonPayFr, tmpVClmNonPayTo) 'done
        Call SaveMBMCovPersonTwo(tmpVCovID) 'done
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



Private Sub SPHISDBTableRecAssign(tmpTableName As String, tmpPlncode As String, tmpInsType As String, TmpEDate As Date)
Dim TmpREc As New ADODB.Recordset
Dim cmd As New ADODB.Command
Dim Prm1 As New ADODB.Parameter
Dim Prm2 As New ADODB.Parameter
Dim Prm3 As New ADODB.Parameter
Dim tmpSelField As String
Dim tmpSelCriteria As String
    
    If tmpTableName = "PLN" Then
        TmpRoundUpStatus = "Y"
        tmpSelField = "Select TOP 1 PLNTopUpStatus, PRMInd, MCOInd, AnnLmtInd, PLNLifeTimeStatus, PLNAgeLimit,PLNRoundUpStatus "
        
        tmpSelCriteria = " AND PLNCode = '" & Trim(tmpPlncode) & _
                         "' AND HLTCode='" & tHLTCode & _
                         "' order by PLNEffDate desc "
'                         "' AND PLNEffDate <='" & Format(TmpEDate, "mm/dd/yyyy") & _

        If tHLTCode = "N" Then
            tmpSelCriteria = " AND PLNCode = '" & Trim(tmpPlncode) & _
                             "' AND HLTCode='" & tHLTCode & _
                             "' AND PAYCode='" & Trim(tPayor) & _
                             "' order by PLNEffDate desc "
'                             "' AND PLNEffDate <='" & Format(TmpEDate, "mm/dd/yyyy") & _

        End If
    ElseIf tmpTableName = "AnnualLimit" Then
        tmpSelField = "Select TOP 1 SuppLimitStatus "
    
        tmpSelCriteria = " AND PLNCode = '" & Trim(tmpPlncode) & "' AND INSCode='" & Trim(tmpInsType) & _
                         "' AND HLTCode='" & tHLTCode & _
                         "' AND AnnualEffDate <='" & Format(TmpEDate, "mm/dd/yyyy") & _
                         "' order by AnnualEffDate desc "

    End If
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

    TmpREc.CursorLocation = adUseClient
    TmpREc.CursorType = adOpenForwardOnly
    TmpREc.CacheSize = 100
    Set TmpREc = cmd.Execute
    
    If TmpREc.EOF Then
        MsgBox tmpTableName & " Information Is Not Correct", vbOKOnly + vbCritical, DialogBoxTitle
    Else
        Do Until TmpREc.EOF
            With TmpREc
                If tmpTableName = "PLN" Then
                    tPLNTopupStatus = IIf(Len(Trim(TmpREc!PLNTopUpStatus)), Trim(TmpREc!PLNTopUpStatus), "N")
                    tPLNPRMInd = IIf(Len(Trim(TmpREc!PRMInd)), Trim(TmpREc!PRMInd), "N")
                    tPLNMCOInd = IIf(Len(Trim(TmpREc!MCOInd)), Trim(TmpREc!MCOInd), "N")
                    tPLNAnnLmtInd = IIf(Len(Trim(TmpREc!ANNLmtInd)), Trim(TmpREc!ANNLmtInd), "N")
                    tPLNAgeLimit = IIf(Len(Trim(TmpREc!PlnAgeLimit)), Trim(TmpREc!PlnAgeLimit), "")
                    tPLNLifeTimeStatus = IIf(Len(Trim(TmpREc!PLNLifeTimeStatus)), Trim(TmpREc!PLNLifeTimeStatus), "N")
                    TmpRoundUpStatus = IIf(Len(Trim(TmpREc!PLNRoundUpStatus)), Trim(TmpREc!PLNRoundUpStatus), "Y")
                ElseIf tmpTableName = "AnnualLimit" Then
                    tAnnSuppLmtStatus = IIf(Len(Trim(TmpREc!SuppLimitStatus)), TmpREc!SuppLimitStatus, "A")
                End If
                TmpREc.MoveNext
            End With
        Loop
    End If
    TmpREc.Close
    Set TmpREc = Nothing
End Sub

Private Sub SPCisDBTableRecAssign(tmpTableName As String, tmpPlncode As String, tmpInsType As String, TmpEDate As Date)
Dim TmpREc As New ADODB.Recordset
Dim cmd As New ADODB.Command
Dim Prm1 As New ADODB.Parameter
Dim Prm2 As New ADODB.Parameter
Dim Prm3 As New ADODB.Parameter
Dim tmpSelField As String
Dim tmpSelCriteria As String
    
    If tmpTableName = "PLN" Then
        tmpSelField = "Select TOP 1 DCStatus, PRMInd, MCOInd, AnnLmtInd, VISlmtInd, " & _
                      "PAYCode, PDentalStatus, SDentalStatus, PSpecStatus, SSpecStatus, " & _
                      "PMepStatus, SMepStatus, PDentalLimit, SDentalLimit, PSpecLimit, " & _
                      "SSpecLimit, AnnLmtIndSupp, VISLmtIndSupp, PlnAgeLimit "
        tmpSelCriteria = " AND PLNCode = '" & Trim(tmpPlncode) & "' AND HLTCode='" & tHLTCode & _
                         "' order by PLNEffDate desc "
'                         "' AND PLNEffDate <='" & Format(TmpEDate, "mm/dd/yyyy") & _

        'If tHLTCode = "N" Then
        '    tmpSelCriteria = " AND PLNCode = '" & Trim(tCPlnCode) & "' AND INSCode='" & Trim(tCInsType) & _
        '                     "' AND HLTCode='" & tHLTCode & _
        '                     "' AND PAYCode='" & Trim(tPayor) & _
        '                     "' order by PLNEffDate desc "
        'End If
    ElseIf tmpTableName = "AnnualLimit" Then
        tmpSelField = "Select TOP 1 SuppLimitStatus "
    
        tmpSelCriteria = " AND PLNCode = '" & Trim(tmpPlncode) & "' AND INSCode='" & Trim(tmpInsType) & _
                         "' AND HLTCode='" & tHLTCode & _
                         "' AND AnnEffDate <='" & Format(TmpEDate, "mm/dd/yyyy") & _
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

    TmpREc.CursorLocation = adUseClient
    TmpREc.CursorType = adOpenForwardOnly
    TmpREc.CacheSize = 100
    Set TmpREc = cmd.Execute
    
    If TmpREc.EOF Then
        MsgBox tmpTableName & " (Clinical) Information Is Not Correct", vbOKOnly + vbCritical, DialogBoxTitle
    Else
        Do Until TmpREc.EOF
            With TmpREc
                If tmpTableName = "PLN" Then
                    tCPLNPRMInd = IIf(Len(Trim(TmpREc!PRMInd)), Trim(TmpREc!PRMInd), "N")
                    tCPLNMCOInd = IIf(Len(Trim(TmpREc!MCOInd)), Trim(TmpREc!MCOInd), "N")
                    tCPLNAnnLmtInd = IIf(Len(Trim(TmpREc!ANNLmtInd)), Trim(TmpREc!ANNLmtInd), "N")
                    tCPLNVsTLmtInd = IIf(Len(Trim(TmpREc!VISLmtInd)), Trim(TmpREc!VISLmtInd), "N")
                    tCPLNAgeLimit = IIf(Len(Trim(TmpREc!PlnAgeLimit)), Trim(TmpREc!PlnAgeLimit), "")
                    tCPlnDCStatus = IIf(Len(Trim(TmpREc!DCStatus)), Trim(TmpREc!DCStatus), "")
                    tCPayor = IIf(Len(Trim(TmpREc!PAYCode)), Trim(TmpREc!PAYCode), tPayor)
                    tCPlnPDentStatus = IIf(Len(Trim(TmpREc!PDentalStatus)), Trim(TmpREc!PDentalStatus), "")
                    tCPlnSDentStatus = IIf(Len(Trim(TmpREc!SDentalStatus)), Trim(TmpREc!SDentalStatus), "")
                    tCPlnPSpecStatus = IIf(Len(Trim(TmpREc!PSpecStatus)), Trim(TmpREc!PSpecStatus), "")
                    tCPlnSSpecStatus = IIf(Len(Trim(TmpREc!SSpecStatus)), Trim(TmpREc!SSpecStatus), "")
                    tCPlnPMepStatus = IIf(Len(Trim(TmpREc!PMepStatus)), Trim(TmpREc!PMepStatus), "")
                    tCPlnSMepStatus = IIf(Len(Trim(TmpREc!SMepStatus)), Trim(TmpREc!SMepStatus), "")
                ElseIf tmpTableName = "AnnualLimit" Then
                    tCAnnSuppLmtStatus = IIf(Len(Trim(TmpREc!SuppLimitStatus)), TmpREc!SuppLimitStatus, "A")
                End If
                TmpREc.MoveNext
            End With
        Loop
    End If
    TmpREc.Close
    Set TmpREc = Nothing
End Sub

Private Sub SaveMBMAmendHistory(tmpMBMNo As String, tmpPOlNo As String, tmpIcNo As String, tmpName As String _
            , tmpPolExpDate As String, tmpPlncode As String, tmpEndtType As String _
            , tmpEndtEffDate As String, tmpSubmitDate As String)
Dim MBMHisSql As String
    
    tmpPolExpDate = IIf(IsDate(tmpPolExpDate), Format(tmpPolExpDate), "")
    tmpEndtEffDate = IIf(IsDate(tmpEndtEffDate), Format(tmpEndtEffDate), "")
    tmpSubmitDate = IIf(IsDate(tmpSubmitDate), Format(tmpSubmitDate), "")
    MBMHisSql = "EXEC Upload_MBMAmendHistory '" & tmpMBMNo & "', '" & tmpPOlNo & "', '" & tmpIcNo & "', " & _
                "'" & tmpName & "','" & tmpPolExpDate & "', '" & tmpPlncode & "','" & tmpEndtType & "', '" & _
                "'" & tmpEndtEffDate & "','" & tmpSubmitDate & "','" & Format(Date, "yyyy/mm/dd") & "','" & Logon & "'"
    medixmasterconn.Execute MBMHisSql
End Sub

Private Sub SaveMBMOthers(tmpDataType As String, tmpPrem As Double, tmpBPrem As Double, tmpBMco As Double, tmpMCO As Double, tmpPolDis As Double, tmpRewDis As Double, tmpVClmNonPayFr As String, tmpVClmNonPayTo As String, tmpVEffDate As String, tmpVGuaRenewal As String)
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

    tmpVlblStatus = Mid(cboLBLStatus, 1, 1)
    tmpVlblStatus = IIf(Len(tmpVlblStatus), tmpVlblStatus, "N")
    tmpVAgentCode = Trim(Mid(TextLine, 27, 15))
    tmpVGrpCompName = Trim(Mid(TextLine, 395, 40))
    tmpVBRNCode = Trim(Mid(TextLine, 2939, 15))
    tmpVDateJoin = Mid(TextLine, 3012, 2) & "/" & Mid(TextLine, 3014, 2) & "/" & Mid(TextLine, 3016, 4)
    tmpVDateJoin = IIf(IsDate(tmpVDateJoin), Format(tmpVDateJoin, "YYYY/MM/DD"), Format(tmpVEffDate, "YYYY/MM/DD"))
    tmpVPrevTOPln = Trim(Mid(TextLine, 3371, 4))
    tmpVPaymentMode = Trim(Mid(TextLine, 3375, 1))
    tmpVPolSubNo1 = Trim(Mid(TextLine, 3376, 10))
    tmpVPolSubNo2 = Trim(Mid(TextLine, 3386, 10))
    tmpVPrevPolNo = Trim(Mid(TextLine, 2954, 25))
    tmpVPrevMBMNo = Trim(Mid(TextLine, 2979, 18))
    tmpVEmpNo = Trim(Mid(TextLine, 2997, 15))
    tmpVDepartment = Trim(Mid(TextLine, 3221, 30))
    tmpVUpgradePln = ""
    tmpVMBMNewPolEffDate = ""
    If Trim(Mid(TextLine, 3592, 4)) <> "" And Trim(Mid(TextLine, 3596, 8)) <> "" Then
       tmpVUpgradePln = Trim(Mid(TextLine, 3592, 4))
       tmpVMBMNewPolEffDate = Mid(TextLine, 3596, 2) & "/" & Mid(TextLine, 3598, 2) & "/" & Mid(TextLine, 3600, 4)
    End If
    tmpVMBMNewPolEffDate = IIf(IsDate(tmpVMBMNewPolEffDate), Format(tmpVMBMNewPolEffDate, "YYYY/MM/DD"), "")
    tmpVEndtRefNo = Trim(Mid(TextLine, 3604, 20))
    tmpVPolCondition = Trim(Mid(TextLine, 3800, 5))
    tmpVUnExcess = IIf(IsNumeric(Trim(Mid(TextLine, 3830, 12))), Trim(Mid(TextLine, 3830, 12)), 0)
    tmpVAgentCode = Trim(Replace(tmpVAgentCode, "'", "''"))

    MBMOWSql = "EXEC Upload_MBMOthersNew '" & Trim(tMBMNo) & "','" & tmpVAgentCode & "','" & tmpVlblStatus & "', " & _
                "'" & tMBMNo & "'," & tmpPrem & "," & tmpBPrem & "," & tmpBMco & "," & tmpMCO & ", " & _
                "'" & tmpDataType & "'," & tmpPolDis & "," & tmpRewDis & ",'" & tmpVGrpCompName & "', '" & tmpVPrevTOPln & "', " & _
                "'" & tmpVDateJoin & "','" & tmpVPrevPolNo & "','" & tmpVPrevMBMNo & "','" & tmpVEmpNo & "', " & _
                "'" & tmpVDepartment & "','" & tmpVPaymentMode & "','" & tmpVPolSubNo1 & "','" & tmpVPolSubNo2 & "','" & tmpVBRNCode & "', " & _
                "'" & tmpVUpgradePln & "','" & tmpVMBMNewPolEffDate & "','" & tmpVEndtRefNo & "','" & tmpVPolCondition & "'," & tmpVUnExcess & ",'" & tmpVGuaRenewal & "'," & _
                "'" & tmpVClmNonPayFr & "','" & tmpVClmNonPayTo & "'"
                   
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
    tmpVPayPremNet = IIf(IsNumeric(Trim(Mid(TextLine, 3684, 12))), Trim(Mid(TextLine, 3684, 12)), 0)
    tmpVPayPremGross = IIf(IsNumeric(Trim(Mid(TextLine, 3696, 12))), Trim(Mid(TextLine, 3696, 12)), 0)
    tmpVPayMcoNet = IIf(IsNumeric(Trim(Mid(TextLine, 3708, 12))), Trim(Mid(TextLine, 3708, 12)), 0)
    tmpVPayMCOGross = IIf(IsNumeric(Trim(Mid(TextLine, 3720, 12))), Trim(Mid(TextLine, 3720, 12)), 0)
    tmpVRnBStatus = Trim(Mid(TextLine, 3762, 2))
    tmpVLoadingPrem = IIf(IsNumeric(Trim(Mid(TextLine, 3764, 12))), Trim(Mid(TextLine, 3764, 12)), 0)
    tmpVFamDiscAmt = IIf(IsNumeric(Trim(Mid(TextLine, 3776, 12))), Trim(Mid(TextLine, 3776, 12)), 0)
    tmpVRenewDiscAmt = IIf(IsNumeric(Trim(Mid(TextLine, 3788, 12))), Trim(Mid(TextLine, 3788, 12)), 0)
    tmpVMBMBankAccNo = Trim(Mid(TextLine, 3805, 25))
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
Dim tmpVNat, tmpVHomeTel, tmpVOffTel, tmpVMobilePhone As String
    
    TmpVSRVCodeList = "ME"
    tmpVNat = Trim(Mid(TextLine, 302, 10))
    tmpVSalutation = Trim(Mid(TextLine, 42, 5))
    tmpVLanguage = Trim(Mid(TextLine, 318, 1))
    tmpVHomeTel = Trim(Mid(TextLine, 237, 12))
    tmpVOffTel = Trim(Mid(TextLine, 249, 12))
    tmpVMobilePhone = Trim(Mid(TextLine, 261, 12))
    tmpVIcBcPp2nd = Trim(Mid(TextLine, 3742, 20))
    tmpVAppName = Trim(Mid(TextLine, 3624, 60))
    tmpVAppName = Trim(Replace(tmpVAppName, "'", "''"))
    tmpVSalutation = Trim(Replace(tmpVSalutation, "'", "''"))
    tmpVNat = Trim(Replace(tmpVNat, "'", "''"))

    MBMTwoWSql = "EXEC Upload_MBMTwo '" & Trim(tMBMNo) & " ','" & tmpVSalutation & "','" & tmpVHomeTel & "','" & tmpVMobilePhone & "'," & _
                 "'" & tmpVOffTel & "','" & tmpVNat & "','" & tmpVLanguage & "','" & TmpVSRVCodeList & "', '" & Logon & "'," & _
                 "'" & Logon & "','" & tmpVClinical & "','" & tmpVAppName & "','" & tmpVIcBcPp2nd & "','A'"
    medixmasterconn.Execute MBMTwoWSql

End Sub

Private Sub SaveMBMCLNOthers(TmpCISMBMType As String, tCPayor As String, tmpVPolicy As String, tmpClnStatus As String, _
tmpClnDAtaStatus As String, tmpClnExpStatus As String, tmpCisPlnCOde As String, tmpCisInsType As String, tmpCisWarranty As String, _
CISBasicPrem As Double, CisPrem As Double, CisBasicMco As Double, CisMCO As Double, CisBalLmt As Double, CisVisitLmt As Double, _
CisDentalLmt As Double, CisSpecLmt As Double, CisMepLmt As Double, tmpCisVEffDate As String, tmpCisVExpDate As String, tmpCisVCardType As String)
Dim StrCLNSql As String
Dim tEffDate, tExpDate As String
    tEffDate = IIf(IsDate(tmpCisVEffDate), Format(tmpCisVEffDate, "YYYY/MM/DD"), "")
    tExpDate = IIf(IsDate(tmpCisVExpDate), Format(tmpCisVExpDate, "YYYY/MM/DD"), "")

    StrCLNSql = "EXEC Upload_MBMClnOthers '" & TmpCISMBMType & "','" & Trim(tCPayor) & "','" & Trim(tMBMNo) & "'," & _
                "'" & tmpVPolicy & "','" & tmpClnStatus & "','" & tmpClnDAtaStatus & "','" & tmpClnExpStatus & "', " & _
                "'" & tmpCisPlnCOde & "','" & tmpCisInsType & "','" & tmpCisWarranty & "'," & CISBasicPrem & "," & CisPrem & ", " & _
                "" & CisBasicMco & "," & CisMCO & "," & CisBalLmt & "," & CisVisitLmt & "," & _
                "" & CisDentalLmt & "," & CisSpecLmt & "," & CisMepLmt & "," & _
                "'" & tEffDate & "','" & tExpDate & "','" & tmpCisVCardType & "'"
    medixmasterconn.Execute StrCLNSql
End Sub

Private Sub SaveMBMCovPerson(tmpVDataType As String, TmpCovID As String, tDOB As String, _
tmpVMemberType As String, tmpVMBMName As String, tmpVMBMStatus As String, tmpVBordxDate As String, tmpVBatchNo As String, _
tmpVIC As String, tmpVSex As String, tmpVPolDis As Double, tmpVRenewDis As Double, tmpVPlanCode As String, _
tAgeCode As String, tAdultChild As String, tmpVAnnLmt As Double, tmpVBasicMco As Double, tmpVMcoAmt As Double, _
tmpVBasicPrem As Double, tmpVPremium As Double, tmpSuppExpDate As String, _
tmpVClmNonPayFr As String, tmpVClmNonPayTo As String)

Dim SuppWSql As String
Dim tmpSExpDate As String
    
    tmpSExpDate = IIf(IsDate(tmpSuppExpDate), Format(tmpSuppExpDate, "yyyy/mm/dd"), "")
    SuppWSql = "EXEC Upload_MBMCovPersonsNew '" & tmpVDataType & "','" & tMBMNo & "','" & TmpCovID & "','" & tDOB & "'," & _
                "'" & tmpVMemberType & "','" & tmpVMBMName & "','" & tmpVMBMStatus & "','" & tmpVBordxDate & "','" & tmpVBatchNo & "', " & _
                "'" & tmpVIC & "','" & tmpVSex & "'," & tmpVPolDis & "," & tmpVRenewDis & ",'" & tmpVPlanCode & "', " & _
                "'" & tAgeCode & "','" & tAdultChild & "'," & tmpVAnnLmt & "," & tmpVAnnLmt & "," & tmpVBasicMco & "," & tmpVMcoAmt & ", " & _
                "" & tmpVBasicPrem & "," & tmpVPremium & ",'" & Format(Date, "yyyy/mm/dd") & "','" & tmpSuppExpDate & "'," & _
                "'" & tmpVClmNonPayFr & "','" & tmpVClmNonPayTo & "'"
    medixmasterconn.Execute SuppWSql
End Sub

Private Sub SaveMBMCovPersonTwo(tmpVCovID As String)
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
    tmpVPayPremNet = IIf(IsNumeric(Trim(Mid(TextLine, 3684, 12))), Trim(Mid(TextLine, 3684, 12)), 0)
    tmpVPayPremGross = IIf(IsNumeric(Trim(Mid(TextLine, 3696, 12))), Trim(Mid(TextLine, 3696, 12)), 0)
    tmpVPayMcoNet = IIf(IsNumeric(Trim(Mid(TextLine, 3708, 12))), Trim(Mid(TextLine, 3708, 12)), 0)
    tmpVPayMCOGross = IIf(IsNumeric(Trim(Mid(TextLine, 3720, 12))), Trim(Mid(TextLine, 3720, 12)), 0)
    tmpVLoadingPrem = IIf(IsNumeric(Trim(Mid(TextLine, 3764, 12))), Trim(Mid(TextLine, 3764, 12)), 0)
    tmpVFamDiscAmt = IIf(IsNumeric(Trim(Mid(TextLine, 3776, 12))), Trim(Mid(TextLine, 3776, 12)), 0)
    tmpVRenewDiscAmt = IIf(IsNumeric(Trim(Mid(TextLine, 3788, 12))), Trim(Mid(TextLine, 3788, 12)), 0)
    tmpVUnExcess = IIf(IsNumeric(Trim(Mid(TextLine, 3830, 12))), Trim(Mid(TextLine, 3830, 12)), 0)
    tmpVSalutation = Trim(Replace(tmpVSalutation, "'", "''"))
    tmpVIcBcPp2nd = Trim(Mid(TextLine, 3742, 20))
    Supp2WSql = "EXEC Upload_MBMCovPersonsTwo '" & Trim(tMBMNo) & "','" & tmpVCovID & "','" & tmpVSalutation & "', " & _
        "'" & TmpVOccupation & " ','" & TmpVWorkNature & "','" & TmpVBlood & "','" & tmpVAllergies & "','" & TmpVSmoking & "', " & _
        "'" & TmpVAlcohol & "','" & tmpVPrevPlnCode & "','" & tmpVPrevInsCom & "','" & tmpVPayRemarks & "','" & tmpVWarrExc & "', " & _
        "'" & tmpVPayNo & "','" & tmpVPaySeqNo & "','" & tmpVClientNo & "','" & tmpVClientID & "'," & tmpVRBAmt & "," & _
        "" & tmpVPayPremNet & ", " & tmpVPayMcoNet & ", " & tmpVPayMCOGross & ", " & _
        "'" & tmpVIcBcPp2nd & "', " & tmpVLoadingPrem & ", " & tmpVFamDiscAmt & ", " & tmpVRenewDiscAmt & ", " & tmpVPayPremGross & "," & _
        "" & tmpVUnExcess
    medixmasterconn.Execute Supp2WSql
End Sub

Private Sub SaveMBMCovPersonCLN(TmpCISMBMType As String, tmpVCovID As String, tmpClnStatus As String, _
tmpCisPlnCOde As String, tmpCisWarranty As String, CISBasicPrem As Double, CisPrem As Double, _
CisBasicMco As Double, CisMCO As Double, CisBalLmt As Double, CisVisitLmt As Double, _
CisDentalLmt As Double, CisSpecLmt As Double, CisMepLmt As Double, tmpSuppCisExpDate As String, _
tmpCisVWarranty As String, tmpCisVCardType As String, tmpCisVPlnCode As String)

Dim CisSuppSql As String
Dim tmpSExpDate As String
    
    tmpSExpDate = IIf(IsDate(CDate(tmpSuppCisExpDate)), Format(tmpSuppCisExpDate, "yyyy/mm/dd"), "")

    CisSuppSql = "EXEC Upload_MBMCovPersonsCis '" & tMBMNo & "','" & tmpVCovID & "', " & _
        "" & CisMCO & "," & CisBasicMco & "," & CisPrem & "," & CISBasicPrem & ", " & _
        "" & CisBalLmt & "," & CisVisitLmt & "," & CisDentalLmt & "," & CisSpecLmt & ", " & _
        "" & CisMepLmt & ",'" & Format(Date, "yyyy/mm/dd") & "','" & Logon & "'," & _
        "'" & Format(Date, "yyyy/mm/dd") & "','" & Logon & "','" & tmpSuppCisExpDate & "'," & _
        "'" & tmpCisVCardType & "','" & tmpCisVWarranty & "','" & tmpClnStatus & "','" & tmpCisVPlnCode & "'"
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
    If tmpVTakeOver = "C" Or tmpVTakeOver = "Y" Then
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

Private Sub SPSearchMBM(tmpSelFields As String, tmpTableName As String, tmpSelCriteria As String, TmpRecFound As Boolean, tmpCode As String, tmpDesc As String)
Dim TmpREc As New ADODB.Recordset
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
End Sub



