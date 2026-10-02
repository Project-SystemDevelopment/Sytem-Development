VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmNewEndorsement 
   Caption         =   "Upload MBM"
   ClientHeight    =   7890
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   11415
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   7890
   ScaleWidth      =   11415
   WindowState     =   2  'Maximized
   Begin TabDlg.SSTab SSTab1 
      Height          =   6975
      Left            =   120
      TabIndex        =   23
      Top             =   240
      Width           =   11175
      _ExtentX        =   19711
      _ExtentY        =   12303
      _Version        =   393216
      Tabs            =   2
      TabsPerRow      =   2
      TabHeight       =   520
      TabCaption(0)   =   "Upload Bordx"
      TabPicture(0)   =   "frmNewEndorsement.frx":0000
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "lstERRList"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Frame1"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "Frame2"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).ControlCount=   3
      TabCaption(1)   =   "Generate ASCII File"
      TabPicture(1)   =   "frmNewEndorsement.frx":001C
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Frame8"
      Tab(1).ControlCount=   1
      Begin VB.Frame Frame8 
         Height          =   3375
         Left            =   -74760
         TabIndex        =   45
         Top             =   360
         Width           =   10815
         Begin VB.Frame Frame14 
            Height          =   2535
            Left            =   120
            TabIndex        =   46
            Top             =   120
            Width           =   10455
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
               TabIndex        =   21
               Top             =   1920
               Width           =   960
            End
            Begin VB.Frame Frame16 
               Height          =   1695
               Left            =   120
               TabIndex        =   55
               Top             =   120
               Width           =   10215
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
                  TabIndex        =   17
                  Top             =   1200
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
                  ItemData        =   "frmNewEndorsement.frx":0038
                  Left            =   4800
                  List            =   "frmNewEndorsement.frx":003A
                  Sorted          =   -1  'True
                  TabIndex        =   18
                  Text            =   "B - Batch Bordx"
                  Top             =   1200
                  Width           =   2025
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
                  TabIndex        =   15
                  Text            =   "S - SIHAT MALAYSIA"
                  Top             =   720
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
                  TabIndex        =   16
                  Top             =   720
                  Width           =   5295
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
                  TabIndex        =   62
                  Top             =   240
                  Width           =   1125
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
                  TabIndex        =   61
                  Top             =   1200
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
                  TabIndex        =   59
                  Top             =   720
                  Width           =   1365
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
                  TabIndex        =   58
                  Top             =   720
                  Width           =   915
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
                  TabIndex        =   12
                  Top             =   240
                  Width           =   2025
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
                  TabIndex        =   13
                  Top             =   240
                  Width           =   1905
               End
               Begin VB.Label txtFilName 
                  BackColor       =   &H00FFFFFF&
                  Height          =   255
                  Left            =   8160
                  TabIndex        =   14
                  Top             =   240
                  Width           =   1905
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
                  TabIndex        =   57
                  Top             =   240
                  Width           =   765
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
                  TabIndex        =   56
                  Top             =   270
                  Width           =   1245
               End
            End
            Begin VB.Frame BatchFrame 
               Height          =   600
               Left            =   120
               TabIndex        =   52
               Top             =   1800
               Width           =   9015
               Begin VB.TextBox txtBatchNo1 
                  Height          =   285
                  Left            =   5640
                  TabIndex        =   20
                  Top             =   180
                  Width           =   2500
               End
               Begin MSMask.MaskEdBox txtReceiptDate 
                  Height          =   285
                  Left            =   1530
                  TabIndex        =   19
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
                  TabIndex        =   54
                  Top             =   180
                  Width           =   675
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
                  TabIndex        =   53
                  Top             =   225
                  Width           =   1515
               End
            End
            Begin VB.Frame BordxFrame 
               Height          =   600
               Left            =   120
               TabIndex        =   47
               Top             =   1800
               Visible         =   0   'False
               Width           =   9015
               Begin MSMask.MaskEdBox txtBordxTo 
                  Height          =   285
                  Left            =   5730
                  TabIndex        =   48
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
                  TabIndex        =   49
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
                  TabIndex        =   51
                  Top             =   225
                  Width           =   1275
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
                  TabIndex        =   50
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
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   570
            Left            =   120
            TabIndex        =   63
            Top             =   2640
            Width           =   10500
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
               TabIndex        =   67
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
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   270
               Left            =   1200
               TabIndex        =   66
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
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   270
               Left            =   4920
               TabIndex        =   65
               Top             =   240
               Width           =   1200
            End
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
               TabIndex        =   64
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
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   120
         TabIndex        =   34
         Top             =   6240
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
            TabIndex        =   44
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
            TabIndex        =   43
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
            TabIndex        =   42
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
            TabIndex        =   41
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
            TabIndex        =   40
            Top             =   240
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
            TabIndex        =   39
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
            TabIndex        =   38
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
            TabIndex        =   35
            Top             =   240
            Width           =   960
         End
      End
      Begin VB.Frame Frame1 
         Height          =   1815
         Left            =   120
         TabIndex        =   24
         Top             =   360
         Width           =   10935
         Begin VB.TextBox txtClnStatus 
            Height          =   285
            Left            =   8880
            TabIndex        =   84
            Text            =   "N"
            Top             =   480
            Visible         =   0   'False
            Width           =   975
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
            Top             =   1395
            Width           =   320
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
            TabIndex        =   4
            Top             =   1380
            Width           =   2625
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
            TabIndex        =   6
            Top             =   240
            Width           =   1695
         End
         Begin VB.ComboBox CboBack2Back 
            Height          =   315
            Left            =   6000
            Sorted          =   -1  'True
            Style           =   2  'Dropdown List
            TabIndex        =   7
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
            Top             =   825
            Width           =   1695
         End
         Begin VB.Frame Frame13 
            Height          =   615
            Left            =   5520
            TabIndex        =   25
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
         End
         Begin VB.Label Label8 
            Caption         =   "Clinical Status"
            Height          =   375
            Left            =   8040
            TabIndex        =   85
            Top             =   480
            Visible         =   0   'False
            Width           =   855
         End
         Begin VB.Label Label15 
            Caption         =   "BTB Grantee"
            Height          =   255
            Left            =   4800
            TabIndex        =   33
            Top             =   570
            Width           =   1215
         End
         Begin VB.Label Label14 
            Caption         =   "Label Status"
            Height          =   255
            Left            =   4800
            TabIndex        =   32
            Top             =   840
            Width           =   975
         End
         Begin VB.Label Label13 
            Caption         =   "Verify I/C"
            Height          =   255
            Left            =   4800
            TabIndex        =   31
            Top             =   255
            Width           =   855
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
            TabIndex        =   30
            Top             =   135
            Width           =   1770
         End
         Begin VB.Label Label5 
            Caption         =   "New/ Endt File"
            Height          =   255
            Left            =   240
            TabIndex        =   29
            Top             =   825
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
            TabIndex        =   28
            Top             =   1440
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
            TabIndex        =   27
            Top             =   480
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
            TabIndex        =   26
            Top             =   1140
            Width           =   960
         End
      End
      Begin MSComctlLib.ListView lstERRList 
         Height          =   3975
         Left            =   120
         TabIndex        =   11
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
   Begin VB.Frame Frame4 
      Height          =   615
      Left            =   120
      TabIndex        =   68
      Top             =   7200
      Width           =   11175
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
         TabIndex        =   83
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
         TabIndex        =   82
         Top             =   240
         Width           =   1080
      End
      Begin VB.CommandButton cmdStop 
         Caption         =   "Stop"
         Height          =   315
         Left            =   7440
         TabIndex        =   81
         Top             =   240
         Width           =   960
      End
      Begin VB.TextBox txtTempMemberNo 
         Height          =   285
         Left            =   960
         TabIndex        =   80
         Text            =   "txtTempMemberNo"
         Top             =   240
         Visible         =   0   'False
         Width           =   375
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
         TabIndex        =   79
         Text            =   "txtTempERRMessage"
         Top             =   240
         Visible         =   0   'False
         Width           =   345
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
         TabIndex        =   78
         Text            =   "txtTempRecordLine"
         Top             =   240
         Visible         =   0   'False
         Width           =   345
      End
      Begin VB.TextBox txtTempPLNCode 
         Height          =   285
         Left            =   2040
         TabIndex        =   77
         Text            =   "txtTempPLNCode"
         Top             =   240
         Visible         =   0   'False
         Width           =   375
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
         TabIndex        =   76
         Text            =   "txtTempINSCode"
         Top             =   240
         Visible         =   0   'False
         Width           =   345
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
         TabIndex        =   75
         Text            =   "txtImportFile"
         Top             =   240
         Visible         =   0   'False
         Width           =   345
      End
      Begin VB.TextBox txtMBMAdultChild 
         Height          =   285
         Left            =   5400
         TabIndex        =   74
         Top             =   240
         Visible         =   0   'False
         Width           =   615
      End
      Begin VB.TextBox txtAccessFees 
         Height          =   285
         Left            =   4680
         TabIndex        =   73
         Top             =   240
         Visible         =   0   'False
         Width           =   735
      End
      Begin VB.TextBox txtMBMMCOFees 
         Height          =   285
         Left            =   3960
         TabIndex        =   72
         Top             =   240
         Visible         =   0   'False
         Width           =   735
      End
      Begin VB.TextBox txtMBMPremium 
         Height          =   285
         Left            =   3240
         TabIndex        =   71
         Top             =   240
         Visible         =   0   'False
         Width           =   735
      End
      Begin VB.TextBox txtDateJoin 
         Height          =   375
         Left            =   0
         TabIndex        =   70
         Text            =   "Text1"
         Top             =   240
         Visible         =   0   'False
         Width           =   255
      End
      Begin VB.TextBox tempAnnLmt 
         Height          =   285
         Left            =   5400
         TabIndex        =   69
         Text            =   "Text1"
         Top             =   240
         Visible         =   0   'False
         Width           =   975
      End
      Begin MSComDlg.CommonDialog CommonDialog1 
         Left            =   2520
         Top             =   120
         _ExtentX        =   847
         _ExtentY        =   847
         _Version        =   393216
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
      TabIndex        =   22
      Top             =   0
      Width           =   11415
   End
End
Attribute VB_Name = "frmNewEndorsement"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim MBMCover As New ADODB.Recordset
Dim MBM As New ADODB.Recordset
Dim MBMOthers As New ADODB.Recordset
Dim CLNOthers As New ADODB.Recordset
Dim MBMTwo As New ADODB.Recordset
Dim MBMOthersTwo As New ADODB.Recordset
Dim MBMCovTwo As New ADODB.Recordset
Dim MBMXrefDB As New ADODB.Recordset
Dim MBMCoveredPersons As New ADODB.Recordset
Dim TextLine
Dim Message As String
Dim RECLine As String
Dim recordCount As Integer
Dim CoveredID As Integer
Dim CoveredIDCount As String
Dim VNewPayor, MNumber As String
Dim INSCode As String
Dim MBMPolicyNo As String
Dim MBMEmployeeNo As String
Dim PayEffDate As String
Dim PayorEXPDate As String
Dim InsuredCode As String
Dim PlanCode As String
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
Public intExit As Integer

Private Sub cboHLTCode_Change()
Dim PAY As New ADODB.Recordset
Dim StrSql As String
    
    StrSql = "Select  PAYCode, PAYCompanyName From PAY where HLTCOde ='" & Mid(Trim(cboHLTCode), 1, 1) & "'"
    PAY.Open StrSql, medixmasterconn, adOpenKeyset, adLockOptimistic
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
Dim StrSql As String
    
    StrSql = "Select  PAYCode, PAYCompanyName From PAY where HLTCOde ='" & Mid(Trim(cboHLTCode), 1, 1) & "'"
    PAY.Open StrSql, medixmasterconn, adOpenKeyset, adLockOptimistic
    cboPAYCode.Clear
    Do Until PAY.EOF
       cboPAYCode.AddItem ChkStr(PAY!PAYCode) & " - " & ChkStr(PAY!PAYCompanyName)
       PAY.MoveNext
    Loop
    PAY.Close
    Set PAY = Nothing
End Sub

Private Sub cboHLTCode_KeyPress(KeyAscii As Integer)
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

Private Sub cboHLTCode1_LostFocus()
    cboHLTCode1.Text = ChkStr(cboHLTCode1.Text)
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
Private Sub CmdGenerate_Click()
Dim StrSql As String
Dim TotalPrincipal As Integer
Dim TotalSupplementary As Integer
Dim RecordSaved
Dim MBMOthers As New ADODB.Recordset
Dim MBMCoveredPersons As New ADODB.Recordset
Dim ASCII As New ADODB.Recordset
Dim textField As String
Dim header As String
Dim Trailer As String

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
                           
'            strsql = "SELECT * From MBM Where " & Exportsql
            StrSql = "SELECT MBMNumber, MBMMemberType, INSCode, MBMPolicyNo, MBMPayorEffDate, MBMIcBcPp, " & _
                    " MBMName From MBM Where " & Exportsql

            MBM.Open StrSql, medixmasterconn, adOpenKeyset, adLockOptimistic

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
End Sub

Private Sub PrincipalASCII()
'Dim strsql As String
'Dim MBM As New ADODB.Recordset
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
    
'    strsql = "SELECT MBMNumber,MBMMemberType,INSCode,MBMPolicyNo,MBMPayorEffDate,MBMName,  From MBM Where " & Exportsql
'    MBM.Open strsql, medixmasterconn, adOpenKeyset, adLockOptimistic
    
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
        
        If Mid(cboASCIIFileVer, 1, 2) = "01" Then
            DataString = DataString & l(MBMGroupCompany, 40)
            DataString = DataString & l("", 6) & Chr(13)
        ElseIf Mid(cboASCIIFileVer, 1, 2) = "02" Then
            DataString = DataString & l("00", 2)
            DataString = DataString & l(MBMPayEffDate, 8)
            DataString = DataString & l(MBMGroupCompany, 40)
            DataString = DataString & l("", 6) & Chr(13)
        ElseIf Mid(cboASCIIFileVer, 1, 2) = "03" Then
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
            
        End If
        
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
        
        'Debug.Print DataString
        
    Loop
End Sub

Private Sub SupplementaryASCII()
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
        'TotalSupplementary = TotalSupplementary + 1
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
        If Mid(cboASCIIFileVer, 1, 2) <> "01" Then
            DataString = DataString & l(Format(MBMCoveredPersons!MBMCCoverID, "00"), 2)
            DataString = DataString & l(MBMPayEffDate, 8)
            DataString = DataString & l("", 40)
        End If
        If Mid(cboASCIIFileVer, 1, 2) = "03" Then
            DataString = DataString & l("", 8)
            DataString = DataString & l("", 4)
            DataString = DataString & l("", 5)
            DataString = DataString & l(XCPayNo, 8)
            DataString = DataString & l(XCpaySeqNo, 2)
            DataString = DataString & l(XCClientNo, 8)
            DataString = DataString & l(XCClientID, 2)
            DataString = DataString & l(MBMPolSubNo1, 10)
            DataString = DataString & l("", 64) & Chr(13)
        Else
            DataString = DataString & l("", 6) & Chr(13)
        End If
        MBMCoveredPersons.MoveNext

        Print #1, DataString
        Print #2, DataString
    Loop
    MBMCoveredPersons.Close
    Set MBMCoveredPersons = Nothing
    MBMCov2.Close
    Set MBMCov2 = Nothing
End Sub


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

Private Sub cboHLTCode1_Change()
Dim PAY As New ADODB.Recordset
Dim StrSql As String
    
    StrSql = "Select  PAYCode, PAYCompanyName From PAY where HLTCOde ='" & Mid(Trim(cboHLTCode1), 1, 1) & "'"
    PAY.Open StrSql, medixmasterconn, adOpenKeyset, adLockOptimistic
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
Dim StrSql As String
    
    StrSql = "Select  PAYCode, PAYCompanyName From PAY where HLTCOde ='" & Mid(Trim(cboHLTCode1), 1, 1) & "'"
    PAY.Open StrSql, medixmasterconn, adOpenKeyset, adLockOptimistic
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
'    NewText = LCase(Left(cboHLTCode1.Text, cboHLTCode1.SelStart) + Chr(KeyAscii) + Mid(cboHLTCode1.Text, cboHLTCode1.SelStart + cboHLTCode1.SelLength + 1))
    
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
  'CommonDialog1.FilterIndex = 2
  CommonDialog1.ShowOpen
  txtImportFilename = CommonDialog1.FileTitle
  Exit Sub
  
ErrHandler:
  'User pressed the Cancel button
  Exit Sub
End Sub

Private Sub cmdPrint_Click()
    Call ExportToExcelError(lstERRList)
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
    Else
        RecordSaved = MsgBox("Do you want to upload this file?", vbYesNo + vbExclamation)
        If RecordSaved = vbYes Then
            Screen.MousePointer = vbHourglass
            cmdUpload.Enabled = False
            cmdExit.Enabled = False

            If Mid(txtFileName, 1, 1) = "N" Then
                Call GenerateMember
            ElseIf Mid(txtFileName, 1, 1) = "E" Then
                Call GenerateENDtMember
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
    intExit = 0
    txtPath = SRVPath & "ASCI"
    txtLocalFileLocation = "C:\"
    Call ComboList
End Sub

'---------------------Start ComboLisiting--------------------------
Private Sub ComboListing()
    Dim HLT As New ADODB.Recordset
    Dim PAY As New ADODB.Recordset
    Dim StrSql As String
    
    HLT.Open "HLT", medixmasterconn, adOpenKeyset, adLockOptimistic
    
'    cboHLTCode1.Clear
'    cboHLTCode.Clear
    Do Until HLT.EOF
       cboHLTCode1.AddItem HLT!HLTCode & " - " & HLT!HLTDescription
       cboHLTCode.AddItem HLT!HLTCode & " - " & HLT!HLTDescription
       HLT.MoveNext
    Loop
    
    txtPAYCode.Clear
    StrSql = "Select  PAYCode, PAYCompanyName From PAY where HLTCOde ='" & Mid(Trim(cboHLTCode1), 1, 1) & "'"
    PAY.Open StrSql, medixmasterconn, adOpenKeyset, adLockOptimistic
    
    Do Until PAY.EOF
       txtPAYCode.AddItem ChkStr(PAY!PAYCode) & " - " & ChkStr(PAY!PAYCompanyName)
       PAY.MoveNext
    Loop
    PAY.Close
    Set PAY = Nothing

    txtFileName.AddItem "N" & " - " & "New File"
    txtFileName.AddItem "E" & " - " & "Endorsement File"
    
    cboASCIIFileVer.AddItem "01" & " - " & "Version 185"
    cboASCIIFileVer.AddItem "02" & " - " & "Version 195"
    cboASCIIFileVer.AddItem "03" & " - " & "Version 300"

    CboVerifyIC.AddItem "Y" & " - " & "Yes"
    CboVerifyIC.AddItem "N" & " - " & "No"
    CboVerifyIC.Text = "Y" & " - " & "Yes"
        
    txtFileVersion.AddItem "01" & " - " & "Version 520"
    txtFileVersion.AddItem "02" & " - " & "Version 790"
    txtFileVersion.AddItem "03" & " - " & "Version 4800"
    txtFileVersion.AddItem "04" & " - " & "Version 4800(v4)"
    
    cboLBLStatus.AddItem "Y" & " - " & "YES"
    cboLBLStatus.AddItem "N" & " - " & "NO"
    
    CboBack2Back.AddItem "Y" & " - " & "Yes"
    CboBack2Back.AddItem "N" & " - " & "No"
    CboBack2Back.Text = "N" & " - " & "No"
    
    HLT.Close
    Set HLT = Nothing
End Sub
Private Sub ComboList()
    Dim HLT As New ADODB.Recordset
    Dim PAY As New ADODB.Recordset
    Dim StrSql As String
    
    StrSql = "Select  PAYCode, PAYCompanyName From PAY where HLTCOde ='" & Mid(Trim(cboHLTCode), 1, 1) & "'"
    PAY.Open StrSql, medixmasterconn, adOpenKeyset, adLockOptimistic
    cboPAYCode.Clear
    Do Until PAY.EOF
       cboPAYCode.AddItem ChkStr(PAY!PAYCode) & " - " & ChkStr(PAY!PAYCompanyName)
       PAY.MoveNext
    Loop
    PAY.Close
    Set PAY = Nothing

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

Private Sub txtFileName_Change()
If Mid(txtFileName, 1, 1) = "N" Then
        txtFileVersion.Enabled = True
        txtFileVersion.Clear
        txtFileVersion.AddItem "01" & " - " & "Version 520"
        txtFileVersion.AddItem "02" & " - " & "Version 790"
        txtFileVersion.AddItem "03" & " - " & "Version 4800"
        txtFileVersion.AddItem "04" & " - " & "Version 4800(v4)"
        CboBack2Back.Enabled = True
    ElseIf Mid(txtFileName, 1, 1) = "E" Then
        txtFileVersion.Clear
        txtFileVersion.AddItem "01" & " - " & "Version 810"
        txtFileVersion.AddItem "02" & " - " & "Version 4830"
        txtFileVersion.AddItem "03" & " - " & "Version 4830(v4)"
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
        txtFileVersion.AddItem "04" & " - " & "Version 4800(v4)"
        CboBack2Back.Enabled = True
    ElseIf Mid(txtFileName, 1, 1) = "E" Then
        txtFileVersion.Clear
        txtFileVersion.AddItem "01" & " - " & "Version 810"
        txtFileVersion.AddItem "02" & " - " & "Version 4830"
        txtFileVersion.AddItem "03" & " - " & "Version 4830(v4)"
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
        txtFileVersion.Enabled = True
        txtFileVersion.Clear
        txtFileVersion.AddItem "01" & " - " & "Version 520"
        txtFileVersion.AddItem "02" & " - " & "Version 790"
        txtFileVersion.AddItem "03" & " - " & "Version 4800"
        txtFileVersion.AddItem "04" & " - " & "Version 4800(v4)"
        CboBack2Back.Enabled = True
    ElseIf Mid(txtFileName, 1, 1) = "E" Then
        txtFileVersion.Clear
        txtFileVersion.AddItem "01" & " - " & "Version 810"
        txtFileVersion.AddItem "02" & " - " & "Version 4830"
        txtFileVersion.AddItem "03" & " - " & "Version 4830(v4)"
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
Dim ErrorCount As String
Dim TotalPrincipalCount As Integer
Dim totalSuppCount As Integer
Dim ListError As ListItem
Dim Msg As VerifyError
Dim Check As Boolean
Dim strsqlPLN As String
Dim PLN As New ADODB.Recordset
Dim Find23Age As String
Dim StrSql As String
Dim NewPLNCode As String
Dim VNewIC, VNewDOB, VNewSex, VNewPostCode, VNewState, VNewCity As String
Dim VNewEffDate As String
Dim VNewExpDate As String
Dim VNewPlanCode, VNewInsured As String
Dim VNewDataType, VRenewal, VNewPolicyNo, vNewName As String
Dim VNewCLnPlan, VNewCLnIns, VNewClnEffDate, VNewClnExpDate As String
Dim MBMICArray() As String
Dim MBMICCnt As Integer
Dim ICCount As Integer
Dim XrefSql As String
Dim XrefDB As New ADODB.Recordset
Dim PrincPayEffdate As String
    
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
                    txtTotalError = txtTotalError + 1
                    Call ErrorListing
                End If
                If Trim(Mid(TextLine, 2, 2)) = "" Or Trim(Mid(TextLine, 2, 2)) <> Mid(txtPAYCode, 1, 2) Then
                    Call CreateErrorList(recordCount, "", "", "Invalid Insurer's Code " & Trim(Mid(TextLine, 2, 2)))
                    txtTotalError = txtTotalError + 1
                    Call ErrorListing
                End If
                If IsDate(txtSubmissionDate) = False Then
                    Call CreateErrorList(recordCount, "", "", "Invalid Submission Date !")
                    txtTotalError = txtTotalError + 1
                    Call ErrorListing
                End If
                VNewPayor = Mid(TextLine, 2, 2)
            End If
                             
            If recordCount >= 2 And Mid(TextLine, 1, 1) <> "T" Then
                VNewPolicyNo = Mid(TextLine, 2, 25)
                vNewName = Mid(TextLine, 47, 60)
                VNewDataType = Mid(TextLine, 1, 1)
                
                VNewCity = Mid(TextLine, 197, 20)
                VNewPostCode = Mid(TextLine, 217, 5)
                VNewState = Mid(TextLine, 222, 15)
                VNewIC = Mid(TextLine, 273, 20)
                VNewDOB = Mid(TextLine, 293, 2) & "/" & Mid(TextLine, 295, 2) & "/" & Mid(TextLine, 297, 4)
                VNewSex = Mid(TextLine, 301, 1)
                VNewEffDate = Mid(TextLine, 374, 2) & "/" & Mid(TextLine, 376, 2) & "/" & Mid(TextLine, 378, 4)
                VNewExpDate = Mid(TextLine, 382, 2) & "/" & Mid(TextLine, 384, 2) & "/" & Mid(TextLine, 386, 4)
                VNewPlanCode = Mid(TextLine, 390, 2)
                VNewInsured = Mid(TextLine, 392, 1)
                VRenewal = Mid(TextLine, 433, 1)
                If Mid(TextLine, 1, 1) = "P" Then
                    PrincPayEffdate = Mid(TextLine, 374, 2) & "/" & Mid(TextLine, 376, 2) & "/" & Mid(TextLine, 378, 4)
                End If
                If Trim(Mid(txtFileVersion, 1, 2)) = "02" Or Trim(Mid(txtFileVersion, 1, 2)) = "03" Then
                    VNewPlanCode = Mid(TextLine, 390, 4)
                    VNewInsured = Mid(TextLine, 394, 1)
                    VRenewal = Mid(TextLine, 435, 1)
                End If
                
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
                    If Mid(TextLine, 1, 1) = "P" Then
                        PrincPayEffdate = Mid(TextLine, 404, 2) & "/" & Mid(TextLine, 406, 2) & "/" & Mid(TextLine, 408, 4)
                    End If
                End If
' compulsory for all MBM - HIS & Clinical
                If Trim(vNewName) = "" Then
                    Call CreateErrorList(recordCount, VNewPolicyNo, vNewName, "Name Is Blank")
                    txtTotalError = txtTotalError + 1
                    Call ErrorListing
                End If
                If Trim(VNewIC) = "" Then
                    Call CreateErrorList(recordCount, VNewPolicyNo, vNewName, "IC,BC or PP Is Blank")
                    txtTotalError = txtTotalError + 1
                    Call ErrorListing
                End If
                If IsDate(Trim(VNewDOB)) = False Then
                    Call CreateErrorList(recordCount, VNewPolicyNo, vNewName, "Invalid Date format ~ Date Of Birth")
                    txtTotalError = txtTotalError + 1
                    Call ErrorListing
                End If
                If Trim(VNewSex) = "" Then
                    Call CreateErrorList(recordCount, VNewPolicyNo, vNewName, "Sex Is Blank")
                    txtTotalError = txtTotalError + 1
                    Call ErrorListing
                End If
                'P-Principal, H, W, S, D
                If Trim(Mid(TextLine, 1, 1)) <> "P" And Trim(Mid(TextLine, 1, 1)) <> "H" And Trim(Mid(TextLine, 1, 1)) <> "W" And Trim(Mid(TextLine, 1, 1)) <> "S" And Trim(Mid(TextLine, 1, 1)) <> "D" Then
                    Call CreateErrorList(recordCount, VNewPolicyNo, vNewName, "Invalid Data Type, should be either 'H','W','S' or 'D' only")
                    txtTotalError = txtTotalError + 1
                    Call ErrorListing
                End If
                If Trim(Mid(TextLine, 107, 90)) = "" Then ' Add1+Add2+Add3 for Princ only
                    If Mid(TextLine, 1, 1) = "P" Then
                        Call CreateErrorList(recordCount, VNewPolicyNo, vNewName, "Address Is Blank")
                        txtTotalError = txtTotalError + 1
                        Call ErrorListing
                    End If
                End If
' compulsory for all MBM - HIS & Clinical

                If Trim(txtClnStatus) = "Y" And Trim(VNewPlanCode) <> "" Then
                    Call CreateErrorList(recordCount, VNewPolicyNo, vNewName, "Invalid Plan Code for 'DC' !! ")
                    txtTotalError = txtTotalError + 1
                    Call ErrorListing
                End If
                
                If Trim(VNewPlanCode) = "" And VNewDataType = "P" Then
                    Call CreateErrorList(recordCount, VNewPolicyNo, vNewName, "HIS Plan Code is null")
                    txtTotalError = txtTotalError + 1
                    Call ErrorListing
                End If
                
                If Mid(CboVerifyIC, 1, 1) = "Y" Then
                    If VNewIC <> "" Then
                        XrefSql = "Select MBMNumber, MBMICBCPP From MBMCrossReference Where MBMICBCPP = '" & Trim(VNewIC) & "' And MBMStatus = 'A'"
                        XrefDB.Open XrefSql, medixmasterconnMaint, adOpenKeyset, adLockOptimistic
    
                        If XrefDB.recordCount > 0 Then
                            Do Until XrefDB.EOF
                                Call CreateErrorList(recordCount, VNewPolicyNo, vNewName, "Same IC " & Trim(VNewIC) & " Existed in Membership No. :" & XrefDB!MBMNumber)
                                txtTotalError = txtTotalError + 1
                                Call ErrorListing
                                XrefDB.MoveNext
                            Loop
                        End If
                        XrefDB.Close
                        Set XrefDB = Nothing
                    End If
                End If
                
' HIS verification
                If Trim(VNewPlanCode) <> "" Then
                    If Trim(VNewPolicyNo) = "" Then  ' Policy No.
                        Call CreateErrorList(recordCount, VNewPolicyNo, vNewName, "Policy No Is Blank")
                        txtTotalError = txtTotalError + 1
                        Call ErrorListing
                    End If
                    If Trim(VNewDataType) <> "P" Then
                        If Trim(Mid(TextLine, 1, 1)) = "S" Or Trim(Mid(TextLine, 1, 1)) = "D" Then
                            Find23Age = Get23Age(CDate(PrincPayEffdate), CDate(VNewDOB))
                            If Find23Age >= 23 Then
                                Call CreateErrorList(recordCount, VNewPolicyNo, vNewName, "Supplementary age is >= 23 ")
                                txtTotalError = txtTotalError + 1
                                Call ErrorListing
                            End If
                        End If
                    End If
                    If Trim(VRenewal) <> "" Then
                        If VRenewal <> "R" And VRenewal <> "N" And VRenewal <> "C" And VRenewal <> "Y" And VRenewal <> "Y" And VRenewal <> "T" And VRenewal <> "P" Then
                            Call CreateErrorList(recordCount, VNewPolicyNo, vNewName, "Invalid Renewal Code !")
                            txtTotalError = txtTotalError + 1
                            Call ErrorListing
                        End If
                    End If
                    
                    If Trim(VNewDataType) = "P" Then
                        Call VerifyMBMByPassingPA("HIS", VNewPolicyNo, vNewName, VNewDataType, VNewPlanCode, VNewInsured, VNewEffDate, VNewExpDate)

' ***------- Check for Member Topup
                        strsqlPLN = "select PLNTopUpStatus from PLN where PLNCode = '" & Trim(VNewPlanCode) & "'"
                        PLN.Open strsqlPLN, medixmasterconn, adOpenKeyset, adLockOptimistic
                        If PLN.recordCount <= 0 Then
                            Call CreateErrorList(recordCount, VNewPolicyNo, vNewName, "Plan Code Not Found ! ")
                            txtTotalError = txtTotalError + 1
                            Call ErrorListing
                        Else
                            If PLN!PLNTopUpStatus = "Y" Then
                                XrefSql = "Select Top 1 MBMNumber, INSCode, MBMTopupNumber From MBMCrossReference Where MBMICBCPP = '" & Trim(VNewIC) & _
                                        "' And MBMStatus = 'A' AND PayCode ='" & VNewPayor & "' AND INSCode='" & VNewInsured & _
                                        "' ORDER BY MBMPayorEffDate DESC, MBMNumber "
                                XrefDB.Open XrefSql, medixmasterconnMaint, adOpenKeyset, adLockOptimistic
                               
                                If XrefDB.recordCount <= 0 Then
                                    Call CreateErrorList(recordCount, VNewPolicyNo, vNewName, "NO Member Record found for Topup ! ")
                                    txtTotalError = txtTotalError + 1
                                    Call ErrorListing
                                Else
                                    If Trim(XrefDB!MBMTopupNumber) <> "" Or IsNull(XrefDB!MBMTopupNumber) = False Then
                                        Call CreateErrorList(recordCount, VNewPolicyNo, vNewName, "NO Record found for Topup ! ")
                                        txtTotalError = txtTotalError + 1
                                        Call ErrorListing
                                    End If
                                    If Trim(XrefDB!INSCode) <> Trim(VNewInsured) Then
                                        Call CreateErrorList(recordCount, VNewPolicyNo, vNewName, "Topup Member Insured Type Is not SAME! ")
                                        txtTotalError = txtTotalError + 1
                                        Call ErrorListing
                                    End If
                                End If
                                XrefDB.Close
                                Set XrefDB = Nothing
                            End If
                        End If
                        PLN.Close
                        Set PLN = Nothing
' ***------ End of Checking for Topup
                    End If
                    
                End If
                
' *** end of HIS Verification
               

' ************* CIS Varification
                If Trim(Mid(txtFileVersion, 1, 2)) = "03" Or Trim(Mid(txtFileVersion, 1, 2)) = "04" Then
                    VNewCLnPlan = Mid(TextLine, 6708, 4)
                    VNewCLnIns = Mid(TextLine, 6712, 1)
                    VNewClnEffDate = Mid(TextLine, 6713, 2) & "/" & Mid(TextLine, 6715, 2) & "/" & Mid(TextLine, 6717, 4)
                    VNewClnExpDate = Mid(TextLine, 6721, 2) & "/" & Mid(TextLine, 6723, 2) & "/" & Mid(TextLine, 6725, 4)
                    If Trim(VNewCLnPlan) = "" And VNewDataType = "P" Then
                        Call CreateErrorList(recordCount, VNewPolicyNo, vNewName, "Clinical Plan Is NULL ! ")
                        txtTotalError = txtTotalError + 1
                        Call ErrorListing
                    End If
                    If Trim(txtClnStatus) = "Y" And VNewDataType = "P" And Trim(VNewCLnPlan) = "" Then
                        Call CreateErrorList(recordCount, VNewPolicyNo, vNewName, "Clinical Plan Code cannot be null !! ")
                        txtTotalError = txtTotalError + 1
                        Call ErrorListing
                    End If
                    If Trim(VNewCLnPlan) <> "" Then
                        Call VerifyMBMByPassingPA("CIS", VNewDataType, " ", vNewName, VNewCLnPlan, VNewCLnIns, VNewClnEffDate, VNewClnExpDate)
                    End If
                End If
' ************* End of CIS verify
' ************* Chk 4 Duplicate Rec in ascii file
                If UBound(MBMICArray) > 0 Then
                    For ICCount = 0 To UBound(MBMICArray)
                        If Trim(MBMICArray(ICCount)) = Trim(VNewIC) Then
                            Call CreateErrorList(recordCount, VNewPolicyNo, vNewName, "Same IC Number Found ! ")
                            txtTotalError = txtTotalError + 1
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
    recordCount = 0
    
End Sub
'---------------------End Verification of New File--------------------------

'---------------------Start Verification of ENDt File-----------------------

Private Sub VerifyENDtMember()
    Dim PayorCode As String
    Dim ErrorCount As String
    Dim TotPrincipal As Integer
    Dim TotSupplementary As Integer
    Dim ListError As ListItem
    Dim Msg As VerifyError
    Dim StrSql As String
    Dim VEndtIC, VEndtDOB, VEndtSuppID As String
    Dim VEndtPostCode, VEndtState, VEndtCity As String
    Dim VEndtEffDate, VEndtExpDate As String
    Dim VEndtPlanCode, VEndtInsured As String
    Dim VEndtDataType, VEndtPolicyNo, vEndtName As String
    Dim VEndtCLnPlan, VEndtCLnIns, VEndtClnEffDate, VEndtClnExpDate As String
    Dim CASX As New ADODB.Recordset
    Dim CASXSql As String
    Dim Find23Age As Integer
    lstERRList.listitems.Clear
    ErrorCount = 0
    Screen.MousePointer = vbHourglass
'    ERR.Open "ERR", medixmasterconn, adOpenKeyset, adLockOptimistic
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
          
                
                
                If Trim(vEndtName) = "" Then
                    Call CreateErrorList(recordCount, VEndtPolicyNo, vEndtName, "Name Is Blank")
                    ErrorCount = ErrorCount + 1
                    Call ErrorListing
                End If
                If Trim(Mid(TextLine, 1, 1)) = "C" And Trim(VEndtDataType) = "P" Then
                    CASXSql = "Select CASStatus FROM CASCrossReference WHERE CASStatus='O' and MBMNumber='" & Trim(Mid(TextLine, 11, 18)) & "'"
                    CASX.Open CASXSql, medixmasterconnMaint, adOpenKeyset, adLockOptimistic
                    If CASX.recordCount > 0 Then
                        Call CreateErrorList(recordCount, VEndtPolicyNo, vEndtName, "There is at least an outstanding case under this member!")
                        ErrorCount = ErrorCount + 1
                        Call ErrorListing
                    End If
                    CASX.Close
                    Set CASX = Nothing
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
                    If Trim(VEndtPlanCode) <> "" Then
                        If Trim(VEndtPolicyNo) = "" Then
                            Call CreateErrorList(recordCount, VEndtPolicyNo, vEndtName, "Policy No Is Blank")
                            ErrorCount = ErrorCount + 1
                            Call ErrorListing
                        End If
                    End If
                    If Trim(VEndtDataType) = "S" Or Trim(VEndtDataType) = "D" Then
                        If Mid(TextLine, 1, 1) = "I" Or Mid(TextLine, 1, 1) = "S" Then
                            Find23Age = Get23Age(CDate(VEndtEffDate), CDate(VEndtDOB))
                            If Find23Age >= 23 Then
                                Call CreateErrorList(recordCount, VEndtPolicyNo, vEndtName, "Supplementary age is >= 23 ")
                                ErrorCount = ErrorCount + 1
                                Call ErrorListing
                            End If
                        End If
                    End If
                    If Mid(CboVerifyIC, 1, 1) = "Y" Then
                        If Trim(VEndtIC) <> "" Then
                            StrSql = "Select MBMNumber, MBMICBCPP From MBM Where MBMICBCPP = '" & Trim(VEndtIC) & "' And MBMStatus = 'A'"
                            MBM.Open StrSql, medixmasterconn, adOpenKeyset, adLockOptimistic
        
                            If MBM.recordCount > 0 Then
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

'    ERR.Close
    txtTotalError = ErrorCount
    txtTotalPrincipal = TotPrincipal
    txtTotalSupplementary = TotSupplementary
    recordCount = 0
    
End Sub
'**************************End Verification of ENDt File*********************************

Private Sub CreateErrorList(RecordLine, Policyno, Name, Msg)
        
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
On ERR GoTo ErrorSave
Dim startTrans As Boolean
Dim StrSql As String
Dim PayorCode As String
Dim recordCount As Integer
Dim CovCnt As Integer

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
            
            medixmasterconn.beginTrans
            medixmasterconnMaint.beginTrans
            startTrans = True
            txtBatchNo = CheckBatch(Mid(txtPAYCode, 1, 2), txtSubmissionDate, Mid(txtFileName, 1, 1))
            Do While Not EOF(1)   ' Loop until end of file.
               Line Input #1, TextLine   ' Read line into variable.
                recordCount = recordCount + 1
            
                If Mid(TextLine, 1, 1) = "P" And Mid(TextLine, 1, 1) <> "T" Then
                    
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
            
            MsgBox "Record Import Sucessfully...", vbOKOnly + vbInformation, DialogBoxTitle
            medixmasterconn.CommitTrans
            medixmasterconnMaint.CommitTrans
            startTrans = False
            Screen.MousePointer = vbDefault
            Close #1
            Exit Sub
        End If
    End If
    Close #1   ' Close file.
ErrorSave:
        Screen.MousePointer = vbDefault
        If startTrans = True Then
            medixmasterconn.RollbackTrans
            medixmasterconnMaint.RollbackTrans
        End If
        MsgBox ERR.Description
End Sub

Private Sub SavePrincipalDataImport(Payor As String, CoveredIDCount As Integer)
Dim MemberDOB As String
Dim FindAge As String
Dim AgeValue As String
Dim AnnualLimitValue As String
Dim GetProString As String
Dim tmpPreMemberNo As String
Dim ListMember As ListItem
Dim TmpCityCode, tmpPostCode, TmpSTACode As String
Dim TmpIC, TmpDOB, TmpSex As String
Dim TmpBRNCode, TmpRaceCode As String
Dim TmpTakeOver, TmpRenewal, TmpAppName As String
Dim TMPMBM As New ADODB.Recordset
Dim tmpPlan As New ADODB.Recordset
Dim TmpSql, TmpSql2 As String
Dim Check As Boolean
Dim HSClnInd As String
Dim tmpPolDis As Double
Dim tmpRenewDis As Double
Dim DiscPlan As New ADODB.Recordset
Dim TmpTopupNo As String
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
Dim tmpINSCode, TmpPlanCode As String
Dim TmpEffDate, TmpExpDate, TmpBordxDate As String
Dim TmpMBMStatus, TmpDataStatus As String
Dim TmpPolicy, TmpAgeCode As String
Dim TmpMBMNumber, TmpMBMName As String
Dim SuppLmtStatus As String

            If Mid(TextLine, 1, 1) = "P" And Mid(TextLine, 1, 1) <> "T" Then
                
                TmpPolicy = Mid(TextLine, 2, 25)
                TmpCovID = "00"
                TmpMBMName = Mid(TextLine, 47, 60)
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
                TmpTopupNo = ""
                TmpTopupInd = ""
                
                If Mid(txtFileVersion, 1, 2) <> "04" Then
                    TmpCityCode = Mid(TextLine, 197, 20)
                    tmpPostCode = Mid(TextLine, 217, 5)
                    TmpSTACode = Mid(TextLine, 222, 15)
                    TmpIC = Mid(TextLine, 273, 20)
                    TmpSex = Mid(TextLine, 301, 1)
                    TmpRaceCode = Mid(TextLine, 319, 1)
                    TmpEffDate = Format(Mid(TextLine, 374, 2) & "/" & Mid(TextLine, 376, 2) & "/" & Mid(TextLine, 378, 4), "yyyy/mm/dd")
                    TmpExpDate = Format(Mid(TextLine, 382, 2) & "/" & Mid(TextLine, 384, 2) & "/" & Mid(TextLine, 386, 4), "yyyy/mm/dd")
                    TmpDOB = Format(Mid(TextLine, 293, 2) & "/" & Mid(TextLine, 295, 2) & "/" & Mid(TextLine, 297, 4), "yyyy/mm/dd")
                    TmpPlanCode = Trim(Mid(TextLine, 390, 4))
                    tmpINSCode = Mid(TextLine, 394, 1)
                    TmpTakeOver = Mid(TextLine, 435, 1)
                    TmpRenewal = Mid(TextLine, 435, 1)
                    TmpBRNCode = ""
                    
                    PayEffDate = Mid(TextLine, 374, 2) & "/" & Mid(TextLine, 376, 2) & "/" & Mid(TextLine, 378, 4)
                    PayorEXPDate = Mid(TextLine, 382, 2) & "/" & Mid(TextLine, 384, 2) & "/" & Mid(TextLine, 386, 4)
                    MemberDOB = Mid(TextLine, 293, 2) & "/" & Mid(TextLine, 295, 2) & "/" & Mid(TextLine, 297, 4)
                    PlanCode = Trim(Mid(TextLine, 390, 4))
                    InsuredCode = Mid(TextLine, 394, 1)
                    
                    If Mid(txtFileVersion, 1, 2) = "01" Then
                        PlanCode = Trim(Mid(TextLine, 390, 2))
                        InsuredCode = Mid(TextLine, 392, 1)
                        
                        TmpPlanCode = Trim(Mid(TextLine, 390, 2))
                        tmpINSCode = Mid(TextLine, 392, 1)
                        TmpTakeOver = Mid(TextLine, 433, 1)
                        TmpRenewal = Mid(TextLine, 433, 1)
                        TmpBRNCode = Mid(TextLine, 497, 15)
                    End If
                    If Mid(txtFileVersion, 1, 2) = "02" Then
                        TmpBRNCode = Mid(TextLine, 499, 15)
                    ElseIf Mid(txtFileVersion, 1, 2) = "03" Then
                        TmpBRNCode = Mid(TextLine, 2939, 15)
                    End If
                Else
                    TmpCityCode = Mid(TextLine, 197, 50)
                    tmpPostCode = Mid(TextLine, 247, 5)
                    TmpSTACode = Mid(TextLine, 252, 15)
                    TmpIC = Mid(TextLine, 303, 20)
                    TmpSex = Mid(TextLine, 331, 1)
                    TmpRaceCode = Mid(TextLine, 349, 1)
                    TmpTakeOver = Mid(TextLine, 465, 1)
                    TmpRenewal = Mid(TextLine, 465, 1)
                    TmpBRNCode = Mid(TextLine, 2969, 15)
                    TmpPlanCode = Trim(Mid(TextLine, 420, 4))
                    tmpINSCode = Mid(TextLine, 424, 1)
                    
                    TmpEffDate = Format(Mid(TextLine, 404, 2) & "/" & Mid(TextLine, 406, 2) & "/" & Mid(TextLine, 408, 4), "yyyy/mm/dd")
                    TmpExpDate = Format(Mid(TextLine, 412, 2) & "/" & Mid(TextLine, 414, 2) & "/" & Mid(TextLine, 416, 4), "yyyy/mm/dd")
                    TmpDOB = Format(Mid(TextLine, 323, 2) & "/" & Mid(TextLine, 325, 2) & "/" & Mid(TextLine, 327, 4), "yyyy/mm/dd")
                   
                    PlanCode = Trim(Mid(TextLine, 420, 4))
                    InsuredCode = Mid(TextLine, 424, 1)
                    PayEffDate = Mid(TextLine, 404, 2) & "/" & Mid(TextLine, 406, 2) & "/" & Mid(TextLine, 408, 4)
                    PayorEXPDate = Mid(TextLine, 412, 2) & "/" & Mid(TextLine, 414, 2) & "/" & Mid(TextLine, 416, 4)
                    MemberDOB = Mid(TextLine, 323, 2) & "/" & Mid(TextLine, 325, 2) & "/" & Mid(TextLine, 327, 4)
                   
                End If
                SuppLmtStatus = "A"
                TmpMemberType = "P"
                
                If TmpTakeOver = "C" Or TmpTakeOver = "Y" Then
                    TmpRenewal = "N"
                End If
        '------------------------------------------------------------------------------------------------------------------------
                If Trim(PlanCode) <> "" Then 'HIS prem, mco, annLmt, Access Calc.
                    strsqlannlmt = "Select SuppLimitStatus, PLNCode from AnnualLimit where PLNCode = '" & Trim(PlanCode) & "'"
                    AnnLmt.Open strsqlannlmt, medixmasterconn, adOpenKeyset, adLockOptimistic
                    
                    If Trim(AnnLmt!SuppLimitStatus) <> "" Then
                        SuppLmtStatus = Trim(AnnLmt!SuppLimitStatus)
                    End If
                    AnnLmt.Close
                    Set AnnLmt = Nothing
                    
                    If Trim(SuppLmtStatus) = "A" Then
                        TmpMemberType = "P"
                    Else
                        TmpMemberType = "Q"
                    End If
                    
                    FindAge = GetAge(CDate(PayEffDate), CDate(MemberDOB))
                    AgeValue = Mid(GetAgeBand(Mid(cboHLTCode1, 1, 1), FindAge), 1, 2)
                    TmpAgeCode = AgeValue
                    TmpAdultChild = "C"
                    If AgeValue <> "01" Then
                        TmpAdultChild = "A"
                    End If
                    
                    txtMBMPremium = Format(GetPremium(InsuredCode, Payor, AgeValue, PlanCode, Mid(cboHLTCode1, 1, 1), CDate(PayEffDate)), "#,##0.00")
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
                        txtMBMMCOFees = Format(GetMCO(InsuredCode, Mid(cboHLTCode1, 1, 1), "C", CDate(PayEffDate), PlanCode), "#,##0.00")
                    Else
                        txtMBMMCOFees = Format(GetMCO(InsuredCode, Mid(cboHLTCode1, 1, 1), "A", CDate(PayEffDate), PlanCode), "#,##0.00")
                    End If
                    BasicMCO = txtMBMMCOFees
                    txtMBMMCOFees = InsertMCOFees(CDate(PayorEXPDate), CDate(PayEffDate), Format(txtMBMMCOFees, "#,##0.00"))
                    txtMBMMCOFees = Round(txtMBMMCOFees)
                    AnnualLimitValue = GetAnnualLimit(InsuredCode, PlanCode, Mid(cboHLTCode1, 1, 1), CDate(PayEffDate))
                    TmpAvLmt = CDbl(AnnualLimitValue)
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
                TmpMBMNumber = tmpPreMemberNo
                CLNPlnSql = "SELECT AnnLmtIND, VisLmtIND, DCStatus, PAYCode, PDentalStatus, SDentalStatus, " & _
                            "PSpecStatus, SSpecStatus, PMEPStatus, SMEPStatus from PLN where PLNCode='" & Trim(CisPlnCode) & _
                        "' AND HLTCode='" & Mid(cboHLTCode1, 1, 1) & "'"
                CLNPln.Open CLNPlnSql, MediConnCISDB, adOpenKeyset, adLockOptimistic
                
                CISCovSuppLMtStat = ""
                CISDCStatus = ""
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
                        CISCovSuppLMtStat = CLnAnnLmt!SuppLimitStatus
                        CLnAnnLmt.Close
                        Set CLnAnnLmt = Nothing
                    End If
                End If
                CLNPln.Close
                Set CLNPln = Nothing
                HSClnInd = "H"
                If Trim(PlanCode) = "" And Trim(CisPlnCode) <> "" Then
                    TmpHSClnInd = "C"
'                    tmpClnStatus = "A"         'Write in MBMCLNOthers
'                    tmpClnDAtaStatus = "A"     'Write in MBMCLNOthers
                End If
                If Trim(PlanCode) <> "" And Trim(CisPlnCode) = "" Then
                    TmpHSClnInd = "H"
                End If
                If Trim(CisPlnCode) <> "" And Trim(PlanCode) <> "" Then
                    If Trim(CISDCStatus) = "Y" Then
                        TmpHSClnInd = "M"
                    Else
                        TmpHSClnInd = "B"
                    End If
'                    tmpClnStatus = "A"         'Write in MBMCLNOthers
'                    tmpClnDAtaStatus = "A"     'Write in MBMCLNOthers
                End If
                TmpMBMNumber = tmpPreMemberNo
                txtTempMemberNo = tmpPreMemberNo
                 
' ---------- write Previous MBM no. to new MBMTopupMember field
                TmpSql = "select * from PLN where PLNCode = '" & Trim(PlanCode) & "'"
                tmpPlan.Open TmpSql, medixmasterconn, adOpenKeyset, adLockOptimistic
                TmpTopupNo = ""
                TmpTopupNo = ""
                If tmpPlan.recordCount Then
                    If tmpPlan!PLNTopUpStatus = "Y" Then
                        XrefSql = "Select TOP 1 MBMNumber, MBMTopupNumber, MBMTopupIND From MBMCrossReference Where MBMICBCPP = '" & Trim(TmpIC) & _
                                    "' And MBMStatus = 'A' AND PayCode ='" & Mid(txtPAYCode, 1, 2) & "' AND INSCode='" & Trim(InsuredCode) & _
                                    "' ORDER BY MBMPayorEffDate DESC, MBMNumber "
                        XrefDB.Open XrefSql, medixmasterconnMaint, adOpenKeyset, adLockOptimistic
 
                        
                        If XrefDB.recordCount Then
                            TmpSql2 = "Select MBMTopupNumber, MBMTopupIND From MBM Where MBMNumber = '" & Trim(XrefDB!MBMNumber) & "'"
                            TMPMBM.Open TmpSql2, medixmasterconn, adOpenKeyset, adLockOptimistic
                            TmpTopupNo = XrefDB!MBMNumber
                            TmpTopupNo = "Y"
                        End If
                        XrefDB.Close
                        Set XrefDB = Nothing
                    End If
                End If
                tmpPlan.Close
                Set tmpPlan = Nothing
    TmpAdd1 = Trim(Replace(TmpAdd1, "'", "''"))
    TmpAdd2 = Trim(Replace(TmpAdd2, "'", "''"))
    TmpAdd3 = Trim(Replace(TmpAdd3, "'", "''"))
    TmpCityCode = Trim(Replace(TmpCityCode, "'", "''"))
    TmpSTACode = Trim(Replace(TmpSTACode, "'", "''"))
    TmpMBMName = Trim(Replace(TmpMBMName, "'", "''"))
    TmpPolicy = Trim(Replace(TmpPolicy, "'", "''"))

    MBMWSql = "Insert Into MBM (PAYCode, HLTCode, MBMNumber, MBMPolicyNo, MBMCOVID, MBMIcBcPp, " & _
              "MBMName, MBMDOB, MBMSex, RACCode, MBMStatus, MBMDataStatus, MBMMemberType, " & _
              "MBMAddress1, MBMAddress2, MBMAddress3, CTYCode, MBMPostCode, STACode, " & _
              "MBMExpiredStatus, MBMValueDate, MBMPayorEffDate, MBMPayorEXPDate, " & _
              "MBMTakeOver, MBMRenewFlag, BRNCode, MBMBordxDate, MBMBatchNo, " & _
              "PLNCode, INSCode, AgeCode, MBMAdultChild, MBMAvailableLimit, " & _
              "HSClnInd, MBMTopupNumber, MBMTopupInd) " & _
              "VALUES('" & tmpPayor & "', '" & tmpHLTCode & "', '" & TmpMBMNumber & "', '" & TmpPolicy & "', '" & TmpCovID & "', '" & TmpIC & "', " & _
              "'" & TmpMBMName & "', '" & TmpDOB & "','" & TmpSex & "','" & TmpRaceCode & "', '" & TmpMBMStatus & "','" & TmpDataStatus & "', '" & TmpMemberType & "', " & _
              "'" & TmpAdd1 & "', '" & TmpAdd2 & "', '" & TmpAdd3 & "', '" & TmpCityCode & "', '" & tmpPostCode & "', '" & TmpSTACode & "', " & _
              "'" & TmpExpStatus & "', '" & TmpValueDate & "', '" & TmpEffDate & "', '" & TmpExpDate & "', " & _
              "'" & TmpTakeOver & "', '" & TmpRenewal & "', '" & TmpBRNCode & "', '" & TmpBordxDate & "', '" & Trim(tmpBatchNo) & "', " & _
              "'" & TmpPlanCode & "', '" & tmpINSCode & "', '" & TmpAgeCode & "', '" & TmpAdultChild & "', " & TmpAvLmt & ", " & _
              "'" & TmpHSClnInd & "', '" & Trim(TmpTopupNo) & "', '" & Trim(TmpTopupInd) & "')"
                
    medixmasterconn.Execute MBMWSql
' ---------- topup Member info
' ----- Policy year calculation
'                TmpSql2 = "Select MBMNumber From MBM Where MBMICBCPP = '" & Trim(TmpIC) & _
'                            "' And MBMStatus = 'A' AND PayCode ='" & Mid(txtPAYCode, 1, 2) & "'"
'                TMPMBM.Open TmpSql2, medixmasterconn, adOpenKeyset, adLockOptimistic

'                If TMPMBM.recordCount Then
'                    !MBMPolicyYear = TMPMBM.recordCount + 1
'                Else
'                    !MBMPolicyYear = 1
'                End If
'                TMPMBM.Close
'                Set TMPMBM = Nothing
' ----- Policy year calculation
            End If
        Call SaveMBMXRef(TmpMBMNumber, "00", Trim(TmpPolicy), TmpIC, TmpMBMName, TmpEffDate, TmpExpDate, TmpAgeCode, tmpINSCode, TmpPlanCode, TmpBordxDate, TmpTopupNo, TmpTopupNo, TmpDataStatus, TmpMBMStatus, TmpDOB, tmpHLTCode, tmpPayor, tmpBatchNo)
        If intExit = 1 Then
            Check = False
  '          Exit Do
        End If
        
  '      Loop Until Check = False

    If Mid(txtFileVersion, 1, 2) = "04" Then
        PayEffDate = Mid(TextLine, 404, 2) & "/" & Mid(TextLine, 406, 2) & "/" & Mid(TextLine, 408, 4)
    Else
        PayEffDate = Mid(TextLine, 374, 2) & "/" & Mid(TextLine, 376, 2) & "/" & Mid(TextLine, 378, 4)
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
Dim TmpIC, TmpSex As String
Dim TmpDOB, TmpAgeCode As String
Dim tmpPolDis As Double
Dim tmpRenewDis As Double
Dim TmpRelCode, TmpCovID As String
Dim TmpMemberType As String
Dim TmpAdultChild, TmpAvLmt As String
Dim TmpPremium, TmpBasicPrem As String
Dim TmpBordxDate, tmpBatchNo As String
Dim TmpStatus As String
Dim TmpMBMNumber, TmpMBMName As String
Dim SuppWSql As String
Dim StrDateSql, StrWDateSql As String

    CoveredIDCount = Format(CoveredIDCount + 1, "00")
    
    TmpRelCode = Trim(Mid(TextLine, 1, 1))
    TmpMBMNumber = Trim(txtTempMemberNo)
    TmpCovID = Format(CoveredIDCount, "00")
    TmpMemberType = "C"
    TmpMBMName = Trim(Mid(TextLine, 47, 60))
    TmpStatus = "A"
    TmpBordxDate = Format(txtSubmissionDate, "yyyy/mm/dd")
    tmpBatchNo = txtBatchNo
    TmpIC = Trim(Mid(TextLine, 273, 20))
    TmpSex = Mid(TextLine, 301, 1)
    TmpDOB = Trim(Mid(TextLine, 293, 2) & "/" & Mid(TextLine, 295, 2) & "/" & Mid(TextLine, 297, 4))
    MemberDOB = Trim(Mid(TextLine, 293, 2) & "/" & Mid(TextLine, 295, 2) & "/" & Mid(TextLine, 297, 4))
    If Mid(txtFileVersion, 1, 2) = "04" Then
        TmpIC = Trim(Mid(TextLine, 303, 20))
        TmpSex = Mid(TextLine, 331, 1)
        TmpDOB = Trim(Mid(TextLine, 323, 2) & "/" & Mid(TextLine, 325, 2) & "/" & Mid(TextLine, 327, 4))
        MemberDOB = Trim(Mid(TextLine, 323, 2) & "/" & Mid(TextLine, 325, 2) & "/" & Mid(TextLine, 327, 4))
    End If
    
    tmpPolDis = IIf(IsNumeric(Trim(Mid(TextLine, 3732, 5))), Trim(Mid(TextLine, 3732, 5)), 0)
    tmpRenewDis = IIf(IsNumeric(Trim(Mid(TextLine, 3737, 5))), Trim(Mid(TextLine, 3737, 5)), 0)
    If Mid(txtFileVersion, 1, 2) = "01" Then
        SuppPlanCode = Trim(Mid(TextLine, 390, 2))
    ElseIf Mid(txtFileVersion, 1, 2) = "02" Then
        SuppPlanCode = Trim(Mid(TextLine, 390, 4))
    ElseIf Mid(txtFileVersion, 1, 2) = "03" Then
        SuppPlanCode = Trim(Mid(TextLine, 390, 4))
    ElseIf Mid(txtFileVersion, 1, 2) = "04" Then
        SuppPlanCode = Trim(Mid(TextLine, 420, 4))
    End If
    FindAge = GetAge(CDate(PayEffDate), CDate(MemberDOB))
    AgeValue = Mid(GetAgeBand(Mid(cboHLTCode1, 1, 1), FindAge), 1, 2)
    TmpAgeCode = AgeValue
    TmpAdultChild = "A"
    If AgeValue = "01" Then
        TmpAdultChild = "C"
    End If
    If Trim(SuppPlanCode) = "" Then
        SuppPlanCode = PlanCode
    End If
    
    StrDateSql = ""
    If Trim(TmpDOB) <> "" Then
        TmpDOB = Format(TmpDOB, "yyyy/mm/dd")
        StrDateSql = StrDateSql & ", MBMCDOB "
        StrWDateSql = StrWDateSql & ",'" & TmpDOB & "'"
    End If
    tmpClmNonPayFr = Mid(TextLine, 3576, 2) & "/" & Mid(TextLine, 3578, 2) & "/" & Mid(TextLine, 3580, 4)
    If IsDate(tmpClmNonPayFr) Then
        tmpClmNonPayFr = Format(tmpClmNonPayFr, "yyyy/mm/dd")
        StrDateSql = StrDateSql & ", MBMCCLMNonPayFr "
        StrWDateSql = StrWDateSql & ",'" & tmpClmNonPayFr & "'"
    End If
    tmpClmNonPayTo = Mid(TextLine, 3584, 2) & "/" & Mid(TextLine, 3586, 2) & "/" & Mid(TextLine, 3588, 4)
    If IsDate(tmpClmNonPayTo) Then
        tmpClmNonPayTo = Format(tmpClmNonPayTo, "yyyy/mm/dd")
        StrDateSql = StrDateSql & ", MBMCCLMNonPayTo "
        StrWDateSql = StrWDateSql & ",'" & tmpClmNonPayTo & "'"
    End If
   
    strSptLmt = "Select SuppLimitStatus, PLNCode from AnnualLimit where PLNCode = '" & Trim(SuppPlanCode) & "'"
    SptPlnLmt.Open strSptLmt, medixmasterconn, adOpenKeyset, adLockOptimistic
    SptLmtInd = "A"
    
    If Trim(SptPlnLmt!SuppLimitStatus) <> "" Then
        SptLmtInd = Trim(SptPlnLmt!SuppLimitStatus)
    End If
    SptPlnLmt.Close
    Set SptPlnLmt = Nothing
    
    strSptPrem = "Select SuppPrmStatus, PLNCode from PRM where PLNCode = '" & Trim(SuppPlanCode) & "'"
    SptPlnPrem.Open strSptPrem, medixmasterconn, adOpenKeyset, adLockOptimistic
    SptPremInd = "A"
    If Trim(SptPlnPrem!SuppPrmStatus) <> "" Then
        SptPremInd = Trim(SptPlnPrem!SuppPrmStatus)
    End If
    SptPlnPrem.Close
    Set SptPlnPrem = Nothing
    If Trim(SptLmtInd) = "B" And Format(CoveredIDCount, "00") = "01" Or Trim(SptLmtInd) = "C" Then
        tempSAnnLmt = Format(GetSuppLimit(InsuredCode, SuppPlanCode, Mid(cboHLTCode1, 1, 1), CDate(PayEffDate)), "#,##0.00")
        TmpAvLmt = tempSAnnLmt
    End If
    If Trim(SptPremInd) = "B" And Format(CoveredIDCount, "00") = "01" Or Trim(SptPremInd) = "C" Then
        tempSPremium = Format(GetSPremium(InsuredCode, Payor, AgeValue, SuppPlanCode, Mid(cboHLTCode1, 1, 1), CDate(PayEffDate)), "#,##0.00")
        tempSPremium = InsertPremium(CDate(PayorEXPDate), CDate(PayEffDate), Format(tempSPremium, "#,##0.00"))
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
' ******* Clinical Supp Limit
        
    If Trim(CISCovSuppLMtStat) = "B" And Format(CoveredIDCount, "00") = "1" Or Trim(CISCovSuppLMtStat) = "C" Then
        Call SaveCISCovPerson(Payor, CoveredIDCount, "N")
    End If
    
' ****** End of Clinical Supp Limit
    TmpPremium = CDbl(IIf(IsNumeric(TmpPremium) = True, TmpPremium, 0))
    TmpBasicPrem = CDbl(IIf(IsNumeric(TmpBasicPrem) = True, TmpBasicPrem, 0))
    tmpPolDis = CDbl(IIf(IsNumeric(tmpPolDis) = True, tmpPolDis, 0))
    tmpRenewDis = CDbl(IIf(IsNumeric(tmpRenewDis) = True, tmpRenewDis, 0))
    TmpAvLmt = CDbl(IIf(IsNumeric(TmpAvLmt) = True, TmpAvLmt, 0))
    TmpMBMName = Replace(TmpMBMName, "'", "''")
    
    SuppWSql = "Insert into MBMCoveredPersons(MBMCRELCode, MBMCNumber, MBMCCoverID, " & _
                "MBMCMemberType, MBMCName, MBMCStatus, MBMCBordxDate, MBMCBatchNo, " & _
                "MBMCICBCPPNo, MBMCSex, MBMCPolicyDisc, MBMCRenewalDisc, MBMCPLNCode, " & _
                "MBMCAgeband, MBMCAdultChild, MBMCAnnualLimit, MBMCAvailableLimit, " & _
                "MBMCBasicPremium, MBMCPremium "
    If Trim(StrDateSql) <> "" Then
        SuppWSql = SuppWSql & StrDateSql
    End If
    SuppWSql = SuppWSql & ")"
    SuppWSql = SuppWSql & "Values('" & TmpRelCode & "','" & TmpMBMNumber & "','" & TmpCovID & "', " & _
                "'" & TmpMemberType & "','" & TmpMBMName & "','" & TmpStatus & "','" & TmpBordxDate & "','" & tmpBatchNo & "', " & _
                "'" & TmpIC & "','" & TmpSex & "'," & tmpPolDis & "," & tmpRenewDis & ",'" & SuppPlanCode & "', " & _
                "'" & AgeValue & "','" & TmpAdultChild & "'," & TmpAvLmt & "," & TmpAvLmt & ", " & _
                "" & TmpBasicPrem & "," & TmpPremium
    If Trim(StrWDateSql) <> "" Then
        SuppWSql = SuppWSql & StrWDateSql
    End If
    SuppWSql = SuppWSql & ")"
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
Dim CisAnnLmtType As String
Dim CisSuppSql As String
Dim TmpCovID As String
Dim MemberDOB As String
Dim SysDate As String

    CoveredIDCount = Format(CoveredIDCount, "00")
    If Trim(BordxFmtType) = "N" Then
        If Mid(txtFileVersion, 1, 2) <> "04" Then
            MemberDOB = Mid(TextLine, 293, 2) & "/" & Mid(TextLine, 295, 2) & "/" & Mid(TextLine, 297, 4)
        Else
            MemberDOB = Mid(TextLine, 323, 2) & "/" & Mid(TextLine, 325, 2) & "/" & Mid(TextLine, 327, 4)
        End If
    ElseIf Trim(BordxFmtType) = "E" Then
        If Mid(txtFileVersion, 1, 2) <> "03" Then
            MemberDOB = Mid(TextLine, 320, 2) & "/" & Mid(TextLine, 322, 2) & "/" & Mid(TextLine, 324, 4)
        Else
            MemberDOB = Mid(TextLine, 350, 2) & "/" & Mid(TextLine, 352, 2) & "/" & Mid(TextLine, 354, 4)
        End If
    End If
    TmpMBMNumber = Trim(txtTempMemberNo)
    TmpCovID = Format(CoveredIDCount, "00")
    FindAge = GetAge(CDate(PayEffDate), CDate(MemberDOB))
    AgeValue = Mid(GetAgeBand(Mid(cboHLTCode1, 1, 1), FindAge), 1, 2)
    

' ******* Clinical Supp Limit
        
    If Trim(CISCovSuppLMtStat) = "B" And Format(CoveredIDCount, "00") = "1" Or Trim(CISCovSuppLMtStat) = "C" Then
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
        If Trim(CISPDentStatus) = "S" Then 'stand alone dental limit
            CisAnnLmtType = "SDentalAnnLimit"
            CisDentalLmt = GetCISLimit(CisInsType, CisPlnCode, Mid(cboHLTCode1, 1, 1), CDate(CisEffDate), CisAnnLmtType)
        End If
        If Trim(CISPSpecStatus) = "S" Then 'stand alone Spec limit
            CisAnnLmtType = "SSpecAnnLimit"
            CisSpecLmt = GetCISLimit(CisInsType, CisPlnCode, Mid(cboHLTCode1, 1, 1), CDate(CisEffDate), CisAnnLmtType)
        End If
        If Trim(CISPMepStatus) = "S" Then 'stand alone MEP limit
            CisAnnLmtType = "SMepAnnLimit"
            CisMepLmt = GetCISLimit(CisInsType, CisPlnCode, Mid(cboHLTCode1, 1, 1), CDate(CisEffDate), CisAnnLmtType)
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
    CisSuppSql = "Insert into MBMCoveredPersonsCLN(MBMCNumber, MBMCCoverID, " & _
               "MBMCCLNMCO, MBMCCLNBasicMCO, MBMCCLNPrem, MBMCCLNBasicPrem, " & _
                "MBMCCLNBalLmt, MBMCCLNVisitLmt, MBMCCLNDentalLmt, MBMCCLNSpecLmt, " & _
                "MBMCCLNMEPLmt, MBMCCLNRegDate, MBMCCLNRegUser, " & _
                "MBMCCLNLastUpdateDate, MBMCCLNLastUpdateUser)"
    
    CisSuppSql = CisSuppSql & "Values('" & TmpMBMNumber & "','" & TmpCovID & "', " & _
                "" & CISSuppMCO & "," & CISSuppBMCO & "," & CISSuppPrem & "," & CISSuppBPrem & ", " & _
                "" & CISSuppAnnLmt & "," & CISSuppVisLmt & "," & CisDentalLmt & "," & CisSpecLmt & ", " & _
                "" & CisMepLmt & ",'" & SysDate & "','" & Logon & "'," & _
                "'" & SysDate & "','" & Logon & "')"
    medixmasterconn.Execute CisSuppSql
End Sub

Private Sub SaveOthersDataImport()
Dim tmpPreMemberNo As String
Dim TempMBM As New ADODB.Recordset
Dim AA, QQ As String
Dim TmpIC, TempSql As String
Dim tmpClmNonPayFr As String
Dim tmpClmNonPayTo As String
Dim Tmp1stJoinedDate As String
Dim tmpAgent, tmplblStatus, TmpAllergic As String
Dim tmp1stMBMNo, TmpPremium, TmpBasicPrem As String
Dim tmpBasicMco, tmpBasicEa, tmpMcoFee As String
Dim TmpRelCode, tmpPolDis, tmpRenewDis As String
Dim tmpGrpComp, tmpPrevPolNo, tmpPrevMBMNo, tmpEmployeeNO As String
Dim tmpBranch, tmpDepart, tmpPolSubNo1, tmpPolSubNo2, tmpInstall As String
Dim tmpMBMNewPln, TmpMBMNewPolEffDate As String
Dim tmpEndtRefNo, TmpPolCondition As String
Dim MBMOWSql As String
Dim StrDateSql, StrWDateSql As String
Dim tmpUnExcess As String

     TmpRelCode = "P"
     tmplblStatus = ""
     tmpAgent = Trim(Mid(TextLine, 27, 15))
     TmpAllergic = Trim(Mid(TextLine, 354, 20))
     If Mid(txtFileVersion, 1, 2) = "04" Then
         TmpAllergic = Trim(Mid(TextLine, 384, 20))
     End If
     If Mid(cboLBLStatus, 1, 1) = "Y" Then
         tmplblStatus = Mid(cboLBLStatus, 1, 1)
     End If
     tmp1stMBMNo = Trim(txtTempMemberNo)
     tmpPolDis = Trim(Mid(TextLine, 3732, 5))
     tmpRenewDis = Trim(Mid(TextLine, 3737, 5))
     TmpPolCondition = Trim(Mid(TextLine, 3800, 5))
     If Mid(txtFileVersion, 1, 2) = "01" Then
         If Len(Mid(TextLine, 393, 40)) Then
             tmpGrpComp = Trim(Mid(TextLine, 393, 40))
         End If
         
     ElseIf Mid(txtFileVersion, 1, 2) = "02" Then
         tmpGrpComp = Trim(Mid(TextLine, 395, 40))
         tmpPrevPolNo = Trim(Mid(TextLine, 514, 25))
         tmpEmployeeNO = Trim(Mid(TextLine, 557, 15))
         Tmp1stJoinedDate = Mid(TextLine, 572, 2) & "/" & Mid(TextLine, 574, 2) & "/" & Mid(TextLine, 576, 4)
     ElseIf Mid(txtFileVersion, 1, 2) = "03" Then
         tmpGrpComp = Trim(Mid(TextLine, 395, 40))
         tmpBranch = Trim(Mid(TextLine, 2939, 15))
         tmpPrevPolNo = Trim(Mid(TextLine, 2954, 25))
         tmpPrevMBMNo = Trim(Mid(TextLine, 2979, 18))
         tmpEmployeeNO = Trim(Mid(TextLine, 2997, 15))
         Tmp1stJoinedDate = Mid(TextLine, 3012, 2) & "/" & Mid(TextLine, 3014, 2) & "/" & Mid(TextLine, 3016, 4)
     ElseIf Mid(txtFileVersion, 1, 2) = "04" Then
         tmpGrpComp = Trim(Mid(TextLine, 425, 40))
         tmpBranch = Trim(Mid(TextLine, 2969, 15))
         tmpPrevPolNo = Trim(Mid(TextLine, 2984, 25))
         tmpPrevMBMNo = Trim(Mid(TextLine, 3009, 18))
         tmpEmployeeNO = Trim(Mid(TextLine, 3027, 15))
         Tmp1stJoinedDate = Mid(TextLine, 3042, 2) & "/" & Mid(TextLine, 3044, 2) & "/" & Mid(TextLine, 3046, 4)
    End If
    If Mid(txtFileVersion, 1, 2) = "03" Or Mid(txtFileVersion, 1, 2) = "04" Then
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
     End If
     tmpPolDis = Trim(Mid(TextLine, 3732, 5))
     tmpRenewDis = Trim(Mid(TextLine, 3737, 5))
     tmpUnExcess = Trim(Mid(TextLine, 3830, 12))
'------------------
    StrDateSql = ""
    StrWDateSql = ""
    If IsDate(Tmp1stJoinedDate) Then
        Tmp1stJoinedDate = Format(Tmp1stJoinedDate, "YYYY/MM/DD")
        StrDateSql = StrDateSql & ", MBMFirstJoinedDate "
        StrWDateSql = StrWDateSql & ",'" & Tmp1stJoinedDate & "'"
    End If
    If IsDate(tmpClmNonPayFr) Then
        tmpClmNonPayFr = Format(tmpClmNonPayFr, "YYYY/MM/DD")
        StrDateSql = StrDateSql & ", MBMCLMNonPayFr "
        StrWDateSql = StrWDateSql & ",'" & tmpClmNonPayFr & "'"
    End If
    If IsDate(tmpClmNonPayTo) Then
        tmpClmNonPayTo = Format(tmpClmNonPayTo, "YYYY/MM/DD")
        StrDateSql = StrDateSql & ", MBMCLMNonPayTo "
        StrWDateSql = StrWDateSql & ",'" & tmpClmNonPayTo & "'"
    End If
    If IsDate(TmpMBMNewPolEffDate) Then
        TmpMBMNewPolEffDate = Format(TmpMBMNewPolEffDate, "YYYY/MM/DD")
        StrDateSql = StrDateSql & ", MBMNewPolEffDate "
        StrWDateSql = StrWDateSql & ",'" & TmpMBMNewPolEffDate & "'"
    End If
    
    TmpPremium = CDbl(IIf(IsNumeric(txtMBMPremium) = True, txtMBMPremium, 0))
    TmpBasicPrem = CDbl(IIf(IsNumeric(BasicPrem) = True, BasicPrem, 0))
    tmpBasicMco = CDbl(IIf(IsNumeric(BasicMCO) = True, BasicMCO, 0))
    tmpMcoFee = CDbl(IIf(IsNumeric(txtMBMMCOFees) = True, txtMBMMCOFees, 0))
    tmpPolDis = CDbl(IIf(IsNumeric(tmpPolDis) = True, tmpPolDis, 0))
    tmpRenewDis = CDbl(IIf(IsNumeric(tmpRenewDis) = True, tmpRenewDis, 0))
    tmpUnExcess = CDbl(IIf(IsNumeric(tmpUnExcess) = True, tmpUnExcess, 0))
    tmpGrpComp = Replace(tmpGrpComp, "'", "''")
    TmpAllergic = Replace(TmpAllergic, "'", "''")
    tmpDepart = Replace(tmpDepart, "'", "''")
    MBMOWSql = "Insert Into MBMOthers(MBMNumber, MBMAgentCode, LabelStatus, MBMAllergic, " & _
                 "MBMFirstMEMNo, MBMPremium, MBMBasicPrem, MBMBasicMCO, MBMMCOFee, " & _
                 "RELCODE, MBMPolicyDisc, MBMRenewalDisc, MBMGroupCompany,  " & _
                 "MBMPrevPolNo, MBMPrevMemNo, MBMEmployeeNo, " & _
                 "MBMDepartment, MBMInstallment, MBMPolSubNo1, MBMPOLSubNo2, MBMBranch, " & _
                 " MBMNewPlanCode, MBMENDtRefNo, MBMPolCondition, MBMUndExcess "
  If Trim(StrDateSql) <> "" Then
        MBMOWSql = MBMOWSql & StrDateSql
  End If
  MBMOWSql = MBMOWSql & ")"
  MBMOWSql = MBMOWSql & "Values('" & Trim(txtTempMemberNo) & "','" & tmpAgent & "','" & tmplblStatus & "','" & TmpAllergic & "', " & _
                 "'" & tmp1stMBMNo & "'," & TmpPremium & "," & TmpBasicPrem & "," & tmpBasicMco & "," & tmpMcoFee & ", " & _
                 "'" & TmpRelCode & "'," & tmpPolDis & "," & tmpRenewDis & ",'" & tmpGrpComp & "', " & _
                 "'" & tmpPrevPolNo & "','" & tmpPrevMBMNo & "','" & tmpEmployeeNO & "', " & _
                 "'" & tmpDepart & "','" & tmpInstall & "','" & tmpPolSubNo1 & "','" & tmpPolSubNo2 & "','" & tmpBranch & "', " & _
                 "'" & tmpMBMNewPln & "','" & tmpEndtRefNo & "','" & TmpPolCondition & "'," & tmpUnExcess
  If Trim(StrWDateSql) <> "" Then
        MBMOWSql = MBMOWSql & StrWDateSql
  End If
  MBMOWSql = MBMOWSql & ")"
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
            If Mid(txtFileVersion, 1, 2) = "04" Then
                EndtInsType = Mid(TextLine, 451, 1)
            Else
                EndtInsType = Mid(TextLine, 421, 1)
            End If
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

    Call EndtCancelPrincipalData
    Call EndtMBMRefund
    Call EndtCancelCoveredPersonsData
End Sub

Private Sub EndtCancelCoveredPersonsData()
Dim StrSql As String
Dim MBMCoveredPersons As New ADODB.Recordset
Dim LastCoveredID  As Integer

    StrSql = "Select MBMCStatus from MBMCoveredPersons Where MBMCNumber ='" & Mid(TextLine, 11, 18) & "'"
    MBMCoveredPersons.Open StrSql, medixmasterconn, adOpenKeyset, adLockOptimistic
    
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

Private Sub EndtCancelPrincipalData()
Dim StrSql As String
Dim MBM As New ADODB.Recordset
Dim MemberAmendHis As New ADODB.Recordset
Dim MBMTwoSql As String
Dim MBMTwo As New ADODB.Recordset
Dim XMBM As New ADODB.Recordset
Dim XMBMSql As String
Dim tmpMBMNo As String

    tmpMBMNo = Trim(Mid(TextLine, 11, 18))
    StrSql = "Select MBMNumber, MBMStatus, MBMLastEndorseStatus, MBMENDtBordxDate,MBMENDtEffDate, " & _
             "MBMPolicyNo, MBMICBcPp, MBMName, MBMPayorEXPDate, PLNCode " & _
             " from MBM Where MBMNumber ='" & tmpMBMNo & "'"
    MBM.Open StrSql, medixmasterconn, adOpenKeyset, adLockOptimistic

    If MBM.recordCount Then
        MemberAmendHis.Open "MBMAmendHistory", medixmasterconn, adOpenKeyset, adLockOptimistic
    
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
            !MBMAmendDate = Date
            !MBMAmendUser = Logon
            MemberAmendHis.Update
        End With
        MemberAmendHis.Close
        Set MemberAmendHis = Nothing
            
        With MBM
            If Len(MBM!MBMStatus) Then
                !MBMStatus = "C"
                !MBMLastEndorseStatus = "C"
                !MBMENDtBordxDate = CDate(txtSubmissionDate)
                !MBMENDtEffDate = Mid(TextLine, 2, 2) & "/" & Mid(TextLine, 4, 2) & "/" & Mid(TextLine, 6, 4)
                MBM.Update
            End If
        End With
        txtTempMemberNo = Trim(Mid(TextLine, 11, 18))
        
        MBMTwoSql = "Select MBMLastAdjustmentDate, MBMLastUpdateUser from MBMTwo Where MBMNumber ='" & Trim(Mid(TextLine, 11, 18)) & "'"
        MBMTwo.Open MBMTwoSql, medixmasterconn, adOpenKeyset, adLockOptimistic
        If MBMTwo.recordCount Then
            MBMTwo!MBMLastAdjustmentDate = Date
            MBMTwo!MBMLastUpdateUser = Logon
            MBMTwo.Update
        End If
        MBMTwo.Close
        Set MBMTwo = Nothing
        
        XMBMSql = "Select MBMStatus from MBMCrossReference Where MBMNumber ='" & Trim(Mid(TextLine, 11, 18)) & "'"
        XMBM.Open XMBMSql, medixmasterconnMaint, adOpenKeyset, adLockOptimistic

        If XMBM.recordCount > 0 Then
            XMBM!MBMStatus = "C"
            XMBM.Update
        End If
        XMBM.Close
        Set XMBM = Nothing
    End If
    MBM.Close
    Set MBM = Nothing
End Sub

Private Sub EndtMBMRefund()
Dim MBMRefund As New ADODB.Recordset
Dim StrSql As String
Dim MCORefundValue As String
Dim MBMMCOFees As Double
Dim PremiumRefundValue As String
Dim MBMAccessFees As Double
Dim MBMAccessFeesEff As String
Dim MBMAccessFeesRef As String
Dim MCOEffective As String
Dim PremiumEffective As String
Dim MBMPremium As Double
Dim PlanCode As String
Dim InsuredCode As String
Dim HLTCode As String
Dim Payor As String
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
    
    StrSql = "Select * from MBMRefund Where MBMNumber = '" & Trim(Mid(TextLine, 11, 18)) & "'"
    MBMRefund.Open StrSql, medixmasterconn, adOpenKeyset, adLockOptimistic
    
    strsqlCAS = "Select MBMPremium, MBMMCOFee, MBMAccessFee from MBMOthers where MBMNumber = '" & Trim(Mid(TextLine, 11, 18)) & "'"
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
    MBMRefund.Close
    Set MBMRefund = Nothing
    MBMOthers.Close
    Set MBMOthers = Nothing
    
End Sub

Private Sub EndtDeleteSupplementary()
'On Error Resume Next
Dim strdelcov As String
Dim MBMCov As New ADODB.Recordset
Dim VEndtSuppID As String
'    strsql = "Select * from MBMCov Where MBMNumber = '" & Mid(TextLine, 11, 18) & "'"
'    MBMCov.Open strsql, medixmasterconn, adOpenKeyset, adLockOptimistic
    VEndtSuppID = Trim(Mid(TextLine, 2963, 2))
    If Mid(txtFileVersion, 1, 2) = "03" Then
       VEndtSuppID = Trim(Mid(TextLine, 2993, 2))
    End If
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

    StrSql = "Select * from MBMCoveredPersons Where MBMCNumber = '" & Trim(Mid(TextLine, 11, 18)) & "'"
    MBMCov.Open StrSql, medixmasterconn, adOpenKeyset, adLockOptimistic

    LastCoveredID = GetLastCoveredID(Trim(Mid(TextLine, 11, 18)))
        MBMCov.AddNew
        With MBMCov
            !MBMCNumber = Trim(Mid(TextLine, 11, 18))
            !MBMCRELCode = Mid(TextLine, 10, 1)
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
                !MBMCICBCPPNo = Trim(Mid(TextLine, 330, 20))
                !MBMCDOB = Mid(TextLine, 350, 2) & "/" & Mid(TextLine, 352, 2) & "/" & Mid(TextLine, 354, 4)
                !MBMCSex = Mid(TextLine, 358, 1)
                !MBMCAllergic = Mid(TextLine, 411, 20)
            Else
                !MBMCICBCPPNo = Trim(Mid(TextLine, 300, 20))
                !MBMCDOB = Mid(TextLine, 320, 2) & "/" & Mid(TextLine, 322, 2) & "/" & Mid(TextLine, 324, 4)
                !MBMCSex = Mid(TextLine, 328, 1)
                !MBMCAllergic = Mid(TextLine, 381, 20)
            End If
            !MBMCPolicyDisc = IIf(IsNumeric(Trim(Mid(TextLine, 3761, 5))), Trim(Mid(TextLine, 3761, 5)), 0)
            !MBMCRenewalDisc = IIf(IsNumeric(Trim(Mid(TextLine, 3766, 5))), Trim(Mid(TextLine, 3766, 5)), 0)

'            If Mid(txtFileVersion, 1, 2) = "01" Then
'                !MBMCExclusion = Mid(TextLine, 463, 60)
'            ElseIf Mid(txtFileVersion, 1, 2) = "02" Then
'                !MBMCExclusion = Mid(TextLine, 463, 2500)
'            ElseIf Mid(txtFileVersion, 1, 2) = "03" Then
'                !MBMCExclusion = Mid(TextLine, 493, 2500)
'            End If
            
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

    strsql2 = "select * from MBM where MBMNumber = '" & Mid(TextLine, 11, 18) & "'"
    MBM.Open "MBM", medixmasterconn, adOpenKeyset, adLockOptimistic
    
    If Mid(txtFileVersion, 1, 2) = "01" Then
        StrSql = "Select * from MBMCoveredPersons Where MBMCNumber = '" & Mid(TextLine, 11, 18) & "' AND MBMCICBCPPNo='" & Mid(TextLine, 74, 60) & "'"
        MBMCov.Open StrSql, medixmasterconn, adOpenKeyset, adLockOptimistic
        If MBMCov.recordCount = 0 Then
            MBMCov.Close
            Set MBMCov = Nothing
            StrSql = "Select * from MBMCoveredPersons Where MBMCNumber = '" & Mid(TextLine, 11, 18) & "' AND MBMCICBCPPNo='" & Mid(TextLine, 300, 20) & "'"
            MBMCov.Open StrSql, medixmasterconn, adOpenKeyset, adLockOptimistic
        End If
    Else
        If Mid(txtFileVersion, 1, 2) = "02" Then 'Version 4830
            SuppID = Mid(TextLine, 2963, 2)
        ElseIf Mid(txtFileVersion, 1, 2) = "03" Then 'Version 4830 (MAA Gen)
            SuppID = Mid(TextLine, 2993, 2)
        End If
        StrSql = "Select * from MBMCoveredPersons Where MBMCNumber = '" & Mid(TextLine, 11, 18) & "' and MBMCCoverID = '" & SuppID & "'"
        MBMCov.Open StrSql, medixmasterconn, adOpenKeyset, adLockOptimistic
    End If
    PayEffDate = MBM!MBMPayorEffDate
    LastCoveredID = GetLastCoveredID(Mid(TextLine, 11, 18))
    If MBMCov.recordCount Then
        With MBMCov
            !MBMCName = Mid(TextLine, 74, 60)
            
            If Mid(txtFileVersion, 1, 2) = "01" Or Mid(txtFileVersion, 1, 2) = "02" Then
                !MBMCSex = Mid(TextLine, 328, 1)
                !MBMCDOB = Mid(TextLine, 320, 2) & "/" & Mid(TextLine, 322, 2) & "/" & Mid(TextLine, 324, 4)
                !MBMCICBCPPNo = Mid(TextLine, 300, 20)
'                !MBMCOccupation = Mid(TextLine, 347, 30)
                !MBMCWorkNature = Mid(TextLine, 377, 1)
                !MBMCBloodType = Mid(TextLine, 378, 3)
                !MBMCAllergic = Mid(TextLine, 381, 20)
                If Mid(txtFileVersion, 1, 2) = "01" Then
'                    !MBMCExclusion = Mid(TextLine, 463, 60)
                    !MBMCSmoking = Mid(TextLine, 523, 1)
                    !MBMCAlcohol = Mid(TextLine, 524, 1)
                ElseIf Mid(txtFileVersion, 1, 2) = "02" Then
'                    !MBMCExclusion = Mid(TextLine, 463, 2500)
                    !MBMCSmoking = Mid(TextLine, 2965, 1)
                    !MBMCAlcohol = Mid(TextLine, 2966, 1)
                End If
                    
            ElseIf Mid(txtFileVersion, 1, 2) = "03" Then
                !MBMCSex = Mid(TextLine, 358, 1)
                !MBMCDOB = Mid(TextLine, 350, 2) & "/" & Mid(TextLine, 352, 2) & "/" & Mid(TextLine, 354, 4)
                !MBMCICBCPPNo = Mid(TextLine, 330, 20)
'                !MBMCOccupation = Mid(TextLine, 377, 30)
                !MBMCWorkNature = Mid(TextLine, 407, 1)
                !MBMCBloodType = Mid(TextLine, 408, 3)
                !MBMCAllergic = Mid(TextLine, 411, 20)
'                !MBMCExclusion = Mid(TextLine, 493, 2500)
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
                If Mid(txtFileVersion, 1, 2) = "01" Or Mid(txtFileVersion, 1, 2) = "02" Then
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
                    
                ElseIf Mid(txtFileVersion, 1, 2) = "03" Then
                    !CTYCode = Mid(TextLine, 224, 50)
                    !MBMPostCode = Mid(TextLine, 274, 5)
                    !STACode = Mid(TextLine, 279, 15)
                    !MBMICBcPp = Mid(TextLine, 330, 20)
                    !MBMDOB = Mid(TextLine, 350, 2) & "/" & Mid(TextLine, 352, 2) & "/" & Mid(TextLine, 354, 4)
                    !MBMSex = Mid(TextLine, 358, 1)
                    !MBMPayorEffDate = Mid(TextLine, 431, 2) & "/" & Mid(TextLine, 433, 2) & "/" & Mid(TextLine, 435, 4)
                    !MBMPayorEXPDate = Mid(TextLine, 439, 2) & "/" & Mid(TextLine, 441, 2) & "/" & Mid(TextLine, 443, 4)
                    !PLNCode = Mid(TextLine, 447, 4)
                    !INSCode = Mid(TextLine, 451, 1)
                    
                End If
                InsuredCode = Trim(!INSCode)
                PlanCode = Trim(!PLNCode)
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
                
                    'TmpSql = "select * from DiscPlan where PLNCode = '" & Trim(PlanCode) & "'"
                    'DiscPlan.Open TmpSql, medixmasterconnMaint, adOpenKeyset, adLockOptimistic
                    'If DiscPlan.recordCount Then
                        If TempPolDis > 0 Then
                            txtMBMPremium = txtMBMPremium * (100 - TempPolDis) / 100
                        End If
                        If TempRenewDis > 0 Then
                            txtMBMPremium = txtMBMPremium * (100 - TempRenewDis) / 100
                        End If
                    'End If
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

Private Sub EndtSaveOthersDataImport()
'On Error Resume Next
'On Error GoTo error1
    Dim tmpPreMemberNo As String
    Dim MBMOthers As New ADODB.Recordset
    Dim tmpClmNonPayFr As String
    Dim tmpClmNonPayTo As String
    
        MBMOthers.Open "MBMOthers", medixmasterconn, adOpenKeyset, adLockOptimistic
       'medixmasterconn.BeginTrans
        
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

            If Mid(txtFileVersion, 1, 2) = "01" Then
                !MBMPrevPolNo = Mid(TextLine, 541, 25)
                !MBMPrevMemNo = Mid(TextLine, 566, 18)
                !MBMEmployeeNo = Mid(TextLine, 584, 15)
            ElseIf Mid(txtFileVersion, 1, 2) = "02" Or Mid(txtFileVersion, 1, 2) = "03" Then
                If Mid(txtFileVersion, 1, 2) = "02" Then
                    !MBMBranch = Mid(TextLine, 2968, 15)
                    !MBMPrevPolNo = Mid(TextLine, 2983, 25)
                    !MBMPrevMemNo = Mid(TextLine, 3008, 18)
                    !MBMEmployeeNo = Mid(TextLine, 3026, 15)
                ElseIf Mid(txtFileVersion, 1, 2) = "03" Then
                    !MBMBranch = Mid(TextLine, 2998, 15)
                    !MBMPrevPolNo = Mid(TextLine, 3013, 25)
                    !MBMPrevMemNo = Mid(TextLine, 3038, 18)
                    !MBMEmployeeNo = Mid(TextLine, 3056, 15)
                     
                End If
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
Dim ListMember As ListItem
Dim MBMCover As New ADODB.Recordset
        
     MBMCover.Open "MBMCoveredPersons", medixmasterconn, adOpenKeyset, adLockOptimistic
    ' medixmasterconn.BeginTrans
     CoveredIDCount = Format(CoveredIDCount + 1, "00")
     
     MBMCover.AddNew
     With MBMCover
'            If Mid(TextLine, 1, 1) <> "P" Then
             !MBMCRELCode = Mid(TextLine, 10, 1)
             !MBMCCoverID = CoveredIDCount
             !MBMCMemberType = "C"
             !MBMCNumber = Trim(txtTempMemberNo)
'            End If
         !MBMCName = Mid(TextLine, 74, 60)
         
         !MBMCStatus = "A"
         !MBMCName = Trim(Mid(TextLine, 74, 60))
         If Mid(txtFileVersion, 1, 2) = "03" Then
             !MBMCICBCPPNo = Trim(Mid(TextLine, 330, 20))
             !MBMCDOB = Mid(TextLine, 350, 2) & "/" & Mid(TextLine, 352, 2) & "/" & Mid(TextLine, 354, 4)
             !MBMCSex = Mid(TextLine, 358, 1)
             !MBMCAllergic = Mid(TextLine, 411, 20)
         Else
             !MBMCICBCPPNo = Trim(Mid(TextLine, 300, 20))
             !MBMCDOB = Mid(TextLine, 320, 2) & "/" & Mid(TextLine, 322, 2) & "/" & Mid(TextLine, 324, 4)
             !MBMCSex = Mid(TextLine, 328, 1)
             !MBMCAllergic = Mid(TextLine, 381, 20)
         End If
        !MBMCPolicyDisc = IIf(IsNumeric(Trim(Mid(TextLine, 3761, 5))), Trim(Mid(TextLine, 3761, 5)), 0)
        !MBMCRenewalDisc = IIf(IsNumeric(Trim(Mid(TextLine, 3766, 5))), Trim(Mid(TextLine, 3766, 5)), 0)
        
        If Mid(txtFileVersion, 1, 1) = "01" Or Mid(txtFileVersion, 1, 1) = "02" Then
            !MBMCDOB = Mid(TextLine, 320, 2) & "/" & Mid(TextLine, 322, 2) & "/" & Mid(TextLine, 324, 4)
            !MBMCSex = Mid(TextLine, 328, 1)
'            If Mid(txtFileVersion, 1, 1) = "01" Then
'                !MBMCExclusion = Mid(TextLine, 463, 60)
'            ElseIf Mid(txtFileVersion, 1, 1) = "02" Then
'                !MBMCExclusion = Mid(TextLine, 463, 2500)
'            End If
            
        ElseIf Mid(txtFileVersion, 1, 1) = "03" Then
            !MBMCDOB = Mid(TextLine, 350, 2) & "/" & Mid(TextLine, 352, 2) & "/" & Mid(TextLine, 354, 4)
            !MBMCSex = Mid(TextLine, 358, 1)
'            !MBMCExclusion = Mid(TextLine, 493, 2500)
        End If
        
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

    CISPayor = IIf(Trim(CISDCStatus) = "Y", "DC", Mid(txtPAYCode, 1, 2))
    If Trim(BordxFmtType) = "N" Then
        If Mid(txtFileVersion, 1, 2) <> "04" Then
            MemberDOB = Mid(TextLine, 293, 2) & "/" & Mid(TextLine, 295, 2) & "/" & Mid(TextLine, 297, 4)
        Else
            MemberDOB = Mid(TextLine, 323, 2) & "/" & Mid(TextLine, 325, 2) & "/" & Mid(TextLine, 327, 4)
        End If
        PolNO = Trim(Mid(TextLine, 2, 25))
        CisWarranty = Trim(Mid(TextLine, 6729, 2500))
    ElseIf Trim(BordxFmtType) = "E" Then
        If Mid(txtFileVersion, 1, 2) <> "03" Then
            MemberDOB = Mid(TextLine, 320, 2) & "/" & Mid(TextLine, 322, 2) & "/" & Mid(TextLine, 324, 4)
        Else
            MemberDOB = Mid(TextLine, 350, 2) & "/" & Mid(TextLine, 352, 2) & "/" & Mid(TextLine, 354, 4)
        End If
        PolNO = Trim(Mid(TextLine, 29, 25))
        CisWarranty = Trim(Mid(TextLine, 6758, 2500))
    End If
    tmpClnStatus = "A"
    tmpClnDAtaStatus = "A"
    tmpClnExpStatus = "N"
    
    FindAge = GetAge(CDate(CisEffDate), CDate(MemberDOB))
    AgeValue = Mid(GetAgeBand(Mid(cboHLTCode1, 1, 1), FindAge), 1, 2)
                   
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
    StrCLNSql = "Insert Into MBMCLNOthers(MBMCLNMemberType, PAYCode, MBMNumber," & _
                "MBMPolNumber, MBMCLNStatus, MBMCLNDataStatus, MBMCLNExpiryStatus, " & _
                "PLNType, InsType, MBMCLNWarranty, MBMCLNBasicPrem, MBMCLNPremium,  " & _
                "MBMCLNBasicMCO, MBMCLNMCOFee, MBMCLNBalLmt, MBMCLNVisitLmt, " & _
                "MBMDentalLmt, MBMSpecLmt, MBMMEPLmt, " & _
                "MBMPolEffDate, MBMPolExpDate)" & _
                "Values('" & TmpMBMType & "','" & Trim(CISPayor) & "','" & Trim(txtTempMemberNo) & "'," & _
                "'" & PolNO & "','" & tmpClnStatus & "','" & tmpClnDAtaStatus & "','" & tmpClnExpStatus & "', " & _
                "'" & Trim(CisPlnCode) & "','" & CisInsType & "','" & CisWarranty & "'," & CISBasicPrem & "," & CisPrem & ", " & _
                "" & CisBasicMco & "," & CisMCO & "," & CisBalLmt & "," & CisVisitLmt & "," & _
                "" & CisDentalLmt & "," & CisSpecLmt & "," & CisMepLmt & "," & _
                "'" & tmpCisEffDate & "','" & tmpCisExpDate & "')"
    medixmasterconn.Execute StrCLNSql
      

End Sub

Private Sub SaveMBMTwo(BordxFmtType As String)
Dim TmpTelNoH, TmpTelNoO, TmpTelNoM, TmpSRVCodeList As String
Dim TmpNatCode, TmpLgnCode, TmpClinical As String
Dim TmpSalutation, TmpAppName, TmpIcBcPp2nd As String
Dim MBMTwoWSql As String

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
            If Mid(txtFileVersion, 1, 2) = "04" Then
                TmpTelNoH = Mid(TextLine, 267, 12)
                TmpTelNoM = Mid(TextLine, 279, 12)
                TmpTelNoO = Mid(TextLine, 291, 12)
                TmpNatCode = Mid(TextLine, 332, 10)
                TmpLgnCode = Mid(TextLine, 348, 1)
            End If
        ElseIf Trim(BordxFmtType) = "E" Then
            TmpSalutation = Trim(Mid(TextLine, 69, 5))
            TmpTelNoH = Trim(Mid(TextLine, 264, 12))
            TmpTelNoM = Trim(Mid(TextLine, 276, 12))
            TmpTelNoO = Trim(Mid(TextLine, 288, 12))
            TmpNatCode = Trim(Mid(TextLine, 329, 10))
            TmpLgnCode = Trim(Mid(TextLine, 345, 1))
            TmpAppName = Trim(Mid(TextLine, 3653, 60))
            TmpIcBcPp2nd = Trim(Mid(TextLine, 3771, 20))
            If Mid(txtFileVersion, 1, 2) = "03" Then
                TmpTelNoH = Trim(Mid(TextLine, 294, 12))
                TmpTelNoM = Trim(Mid(TextLine, 306, 12))
                TmpTelNoO = Trim(Mid(TextLine, 318, 12))
                TmpNatCode = Trim(Mid(TextLine, 359, 10))
                TmpLgnCode = Trim(Mid(TextLine, 375, 1))
            End If
        End If
        If Trim(PlanCode) = "" Then
            TmpClinical = "Y"
        ElseIf Trim(CisPlnCode) <> "" Then
            TmpClinical = "Y"
        End If
        TmpAppName = Replace(TmpAppName, "'", "''")
        TmpSalutation = Replace(TmpSalutation, "'", "''")
        MBMTwoWSql = "Insert Into MBMTwo(MBMNumber, MBMSalutation, MBMTelnoH, MBMTelNoM, " & _
                     "MBMTelNoO, NATCode, LGNCode, SRVCodeList, MBMRegUser, " & _
                     "MBMLastUpdateUser, MBMClinical, MBMAppName, MBMIcBcPp2nd)" & _
                     "Values('" & Trim(txtTempMemberNo) & "','" & TmpSalutation & "','" & TmpTelNoH & "','" & TmpTelNoM & "'," & _
                     "'" & TmpTelNoO & "','" & TmpNatCode & "','" & TmpLgnCode & "','" & TmpSRVCodeList & "', '" & Logon & "'," & _
                     "'" & Logon & "','" & TmpClinical & "','" & TmpAppName & "','" & TmpIcBcPp2nd & "')"
        medixmasterconn.Execute MBMTwoWSql
End Sub
Private Sub SaveMBMOthersTwo(BordxFmtType As String)
Dim TmpWeight, TmpHeight, TmpWorkNature As String
Dim TmpBlood, TmpOccupation, TmpSmoking, TmpAlcohol As String
Dim TmpMaritalStatus, TmpBTBG As String
Dim TmpGrpCompAdd1, TmpGrpCompAdd2, TmpGrpCompAdd3, TmpGrpCompAdd4 As String
Dim POLIssuedDate As String
Dim TmpRiskNo, TmpTransNo, TmpClientNo, TmpClientID As String
Dim TmpPayNo, TmpPaySeqNo As String
Dim TmpRBAmt As String
Dim TmpPrevInsCom, TmpPrevPlnCode, TmpPayRemarks As String
Dim TmpCoPayRB, TmpstaffDec, TmpMBMBankAccNo As String
Dim TmpGenWP, TmpPreExWP, TmpSpeIllWP As String
Dim TmpGenEndWP, TmpPreExEndWP, TmpSpeIllEndWP As String
Dim TmpPayPremNet As Double
Dim TmpPayMcoNet As Double
Dim TmpPayMCOGross As Double
Dim TmpPayPremGross As Double
Dim TmpLoadingPrem As Double
Dim TmpFamDiscAmt As Double
Dim TmpRenewDiscAmt As Double
Dim TmpExclusion As String
Dim MBMOTwoWSql As String
Dim StrDateSql As String
Dim StrWDateSql As String

    TmpGenWP = ""
    TmpPreExWP = ""
    TmpSpeIllWP = ""
    TmpMBMBankAccNo = ""
    TmpstaffDec = ""
    If Trim(BordxFmtType) = "N" Then
        TmpGrpCompAdd1 = Trim(Mid(TextLine, 3251, 30))
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
        TmpRBAmt = Trim(Mid(TextLine, 3448, 8))
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
        
        TmpWeight = Trim(Mid(TextLine, 312, 3))
        TmpHeight = Trim(Mid(TextLine, 315, 3))
        TmpWorkNature = Mid(TextLine, 350, 1)
        TmpBlood = Trim(Mid(TextLine, 351, 3))
        TmpOccupation = Trim(Mid(TextLine, 320, 30))
        If Mid(txtFileVersion, 1, 2) = "01" Then
            TmpSmoking = Mid(TextLine, 494, 1)
            TmpAlcohol = Mid(TextLine, 495, 1)
            TmpMaritalStatus = Mid(TextLine, 496, 1)
            TmpExclusion = Trim(Mid(TextLine, 434, 60))
        ElseIf Mid(txtFileVersion, 1, 2) = "02" Then
            TmpSmoking = Mid(TextLine, 496, 1)
            TmpAlcohol = Mid(TextLine, 497, 1)
            TmpMaritalStatus = Mid(TextLine, 498, 1)
            TmpExclusion = Trim(Mid(TextLine, 436, 60))
        ElseIf Mid(txtFileVersion, 1, 2) = "03" Then
            TmpSmoking = Mid(TextLine, 2936, 1)
            TmpAlcohol = Mid(TextLine, 2937, 1)
            TmpMaritalStatus = Mid(TextLine, 2938, 1)
            TmpExclusion = Trim(Mid(TextLine, 436, 2500))
            TmpMBMBankAccNo = Trim(Mid(TextLine, 3805, 25))
        End If
       If Mid(txtFileVersion, 1, 2) = "04" Then
            TmpWeight = Trim(Mid(TextLine, 342, 3))
            TmpHeight = Trim(Mid(TextLine, 345, 3))
            TmpWorkNature = Mid(TextLine, 380, 1)
            TmpBlood = Mid(TextLine, 381, 3)
            TmpOccupation = Mid(TextLine, 350, 30)
            TmpSmoking = Mid(TextLine, 2966, 1)
            TmpAlcohol = Mid(TextLine, 2967, 1)
            TmpMaritalStatus = Mid(TextLine, 2968, 1)
            TmpExclusion = Trim(Mid(TextLine, 466, 2500))
            TmpMBMBankAccNo = Trim(Mid(TextLine, 3805, 25))
            TmpstaffDec = Trim(Mid(TextLine, 3050, 1))
            If IsNumeric(Mid(TextLine, 3051, 5)) Then
                TmpGenWP = Trim(Mid(TextLine, 3051, 5))
            End If
            
            If IsNumeric(Mid(TextLine, 3056, 5)) Then
                TmpPreExWP = Trim(Mid(TextLine, 3056, 5))
            End If
            
            If IsNumeric(Mid(TextLine, 3061, 5)) Then
                TmpSpeIllWP = Trim(Mid(TextLine, 3061, 5))
            End If
        End If
    ElseIf Trim(BordxFmtType) = "E" Then
        
        TmpWeight = Trim(Mid(TextLine, 339, 3))
        TmpHeight = Trim(Mid(TextLine, 342, 3))
        TmpWorkNature = Mid(TextLine, 377, 1)
        TmpBlood = Trim(Mid(TextLine, 378, 3))
        TmpOccupation = Trim(Mid(TextLine, 347, 30))
        TmpSmoking = Mid(TextLine, 523, 1)
        TmpAlcohol = Mid(TextLine, 524, 1)
        TmpMaritalStatus = Mid(TextLine, 525, 1)
        TmpExclusion = Trim(Mid(TextLine, 463, 60))
        If Mid(txtFileVersion, 1, 2) = "02" Then
            TmpSmoking = Mid(TextLine, 2965, 1)
            TmpAlcohol = Mid(TextLine, 2966, 1)
            TmpMaritalStatus = Mid(TextLine, 2967, 1)
            TmpExclusion = Trim(Mid(TextLine, 463, 2500))
        End If
        If Mid(txtFileVersion, 1, 2) = "03" Then
            TmpWeight = Trim(Mid(TextLine, 369, 3))
            TmpHeight = Trim(Mid(TextLine, 372, 3))
            TmpWorkNature = Mid(TextLine, 407, 1)
            TmpBlood = Trim(Mid(TextLine, 408, 3))
            TmpOccupation = Mid(TextLine, 377, 30)
            TmpSmoking = Mid(TextLine, 2995, 1)
            TmpAlcohol = Mid(TextLine, 2996, 1)
            TmpMaritalStatus = Mid(TextLine, 2997, 1)
            TmpstaffDec = Trim(Mid(TextLine, 3079, 1))
            TmpGenWP = Trim(Mid(TextLine, 3080, 5))
            TmpPreExWP = Trim(Mid(TextLine, 3085, 5))
            TmpSpeIllWP = Trim(Mid(TextLine, 3090, 5))
            TmpExclusion = Trim(Mid(TextLine, 493, 2500))
        End If
        TmpGrpCompAdd1 = Trim(Mid(TextLine, 3280, 30))
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
        TmpRBAmt = Mid(TextLine, 3477, 8)
        TmpPrevInsCom = Trim(Mid(TextLine, 3585, 60))
        TmpPrevPlnCode = Trim(Mid(TextLine, 3462, 15))
        TmpPayRemarks = Trim(Mid(TextLine, 3545, 60))
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
        
    StrDateSql = ""
    StrWDateSql = ""
    If IsDate(POLIssuedDate) Then
        POLIssuedDate = Format(POLIssuedDate, "YYYY/MM/DD")
        StrDateSql = StrDateSql & ", MBMPOLIssuedDate "
        StrWDateSql = StrWDateSql & ",'" & POLIssuedDate & "'"
    End If
    If IsNumeric(TmpGenWP) Then
        TmpGenEndWP = Format(DateAdd("d", TmpGenWP, PayEffDate), "yyyy/mm/dd")
        StrDateSql = StrDateSql & ", MBMGenEndWP "
        StrWDateSql = StrWDateSql & ",'" & TmpGenEndWP & "'"
    End If
    If IsNumeric(TmpPreExWP) Then
        TmpPreExEndWP = Format(DateAdd("d", TmpPreExWP, PayEffDate), "yyyy/mm/dd")
        StrDateSql = StrDateSql & ", MBMPreExEndWP "
        StrWDateSql = StrWDateSql & ",'" & TmpPreExEndWP & "'"
    End If
    If IsNumeric(TmpSpeIllWP) Then
        TmpSpeIllEndWP = Format(DateAdd("d", TmpSpeIllWP, PayEffDate), "yyyy/mm/dd")
        StrDateSql = StrDateSql & ", MBMSpeIllEndWP "
        StrWDateSql = StrWDateSql & ",'" & TmpPreExEndWP & "'"
    End If
    
    TmpRBAmt = CDbl(IIf(IsNumeric(TmpRBAmt) = True, TmpRBAmt, 0))
    TmpGenWP = CDbl(IIf(IsNumeric(TmpGenWP) = True, TmpGenWP, 0))
    TmpPreExWP = CDbl(IIf(IsNumeric(TmpPreExWP) = True, TmpPreExWP, 0))
    TmpSpeIllWP = CDbl(IIf(IsNumeric(TmpSpeIllWP) = True, TmpSpeIllWP, 0))

    TmpOccupation = Replace(TmpOccupation, "'", "''")
    TmpGrpCompAdd1 = Replace(TmpGrpCompAdd1, "'", "''")
    TmpGrpCompAdd2 = Replace(TmpGrpCompAdd2, "'", "''")
    TmpGrpCompAdd3 = Replace(TmpGrpCompAdd3, "'", "''")
    TmpGrpCompAdd4 = Replace(TmpGrpCompAdd4, "'", "''")
    TmpPrevInsCom = Replace(TmpPrevInsCom, "'", "''")
    TmpExclusion = Replace(TmpExclusion, "'", "''")
    TmpMBMBankAccNo = Replace(TmpMBMBankAccNo, "'", "''")
    MBMOTwoWSql = "Insert Into MBMOthersTwo(MBMNumber,MBMWeight,MBMHeight,MBMWorkNature, " & _
                  "MBMBlood, MBMOccupation, MBMSmoking, MBMAlcohol, MBMMaritalStatus, " & _
                  "MBMGRPCompanyAdd1, MBMGRPCompanyAdd2, MBMGRPCompanyAdd3, MBMGRPCompanyAdd4, " & _
                  "MBMRiskNo, MBMTransNo, MBMPayNo, MBMPaySeqNo, MBMBankACNo , " & _
                  "MBMClientNo, MBMClientID, MBMRBAmt, " & _
                  "MBMPrevInsCom, MBMPrevPlnCode, MBMPayRemarks, MBMCoPayRB, MBMStaffDec, MBMExclusion," & _
                  "MBMBTBG, MBMPayorPremNet,  MBMPayorMcoNet, MBMPayorMCoGross, MBMLoadingPrem, " & _
                  "MBMFamDiscAmt, MBMRenewDiscAmt, " & _
                  "MBMGenWP, MBMPreExWP, MBMSpeIllWP,MBMPayorPremGross "
                  
    If Trim(StrDateSql) <> "" Then
          MBMOTwoWSql = MBMOTwoWSql & StrDateSql
    End If
    MBMOTwoWSql = MBMOTwoWSql & ")"
    
    MBMOTwoWSql = MBMOTwoWSql & "Values ('" & Trim(txtTempMemberNo) & "', '" & TmpWeight & "', '" & TmpHeight & "','" & TmpWorkNature & "', " & _
                  "'" & TmpBlood & "','" & TmpOccupation & "','" & TmpSmoking & "','" & TmpAlcohol & "','" & TmpMaritalStatus & "', " & _
                  "'" & TmpGrpCompAdd1 & "','" & TmpGrpCompAdd2 & "','" & TmpGrpCompAdd3 & "','" & TmpGrpCompAdd4 & "', " & _
                  "'" & TmpRiskNo & "','" & TmpTransNo & "','" & TmpPayNo & "','" & TmpPaySeqNo & "','" & TmpMBMBankAccNo & "'," & _
                  "'" & TmpClientNo & "','" & TmpClientID & "'," & TmpRBAmt & "," & _
                  "'" & TmpPrevInsCom & "','" & TmpPrevPlnCode & "','" & TmpPayRemarks & "','" & TmpCoPayRB & "','" & TmpstaffDec & "','" & TmpExclusion & "', " & _
                  "'" & TmpBTBG & "'," & TmpPayPremNet & "," & TmpPayMcoNet & "," & TmpPayMCOGross & "," & TmpLoadingPrem & ", " & _
                  "" & TmpFamDiscAmt & "," & TmpRenewDiscAmt & ", " & _
                  "" & TmpGenWP & "," & TmpPreExWP & "," & TmpSpeIllWP & " ," & TmpPayPremGross & ""
                  
    If Trim(StrWDateSql) <> "" Then
          MBMOTwoWSql = MBMOTwoWSql & StrWDateSql
    End If
    MBMOTwoWSql = MBMOTwoWSql & ")"
    
    medixmasterconn.Execute MBMOTwoWSql
    
End Sub

Private Sub SaveCovPersonsTwo(CovCnt As Integer, BordxFmtType As String)
Dim TmpBlood, TmpOccupation, TmpSmoking, TmpAlcohol, TmpWorkNature  As String
Dim TmpSalutation, TmpMaritalStatus, TmpAllergic As String
Dim TmpRiskNo, TmpTransNo, TmpClientNo, TmpClientID As String
Dim TmpPayNo, TmpPaySeqNo As String
Dim TmpRBAmt As String
Dim TmpPrevInsCom, TmpPrevPlnCode, TmpPayRemarks As String
Dim TmpCoPayRB, TmpstaffDec As String
Dim TmpGenWP, TmpPreExWP, TmpSpeIllWP As String
Dim TmpGenEndWP, TmpPreExEndWP, TmpSpeIllEndWP As String
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
Dim StrDateSql, StrWDateSql As String

    TmpGenWP = ""
    TmpPreExWP = ""
    TmpSpeIllWP = ""
    If Trim(BordxFmtType) = "N" Then
        TmpOccupation = Mid(TextLine, 320, 30)
        TmpWorkNature = Mid(TextLine, 350, 1)
        TmpBlood = Mid(TextLine, 351, 3)
        TmpAllergic = Mid(TextLine, 354, 20)
        TmpSmoking = Mid(TextLine, 494, 1)
        TmpAlcohol = Mid(TextLine, 495, 1)
        TmpSalutation = Trim(Mid(TextLine, 42, 5))
        TmpSuppExclusion = Trim(Mid(TextLine, 434, 60))
        If Mid(txtFileVersion, 1, 2) = "02" Then
            TmpSmoking = Mid(TextLine, 496, 1)
            TmpAlcohol = Mid(TextLine, 497, 1)
            TmpSuppExclusion = Trim(Mid(TextLine, 390, 4))
        ElseIf Mid(txtFileVersion, 1, 2) = "03" Then
            TmpSmoking = Mid(TextLine, 2936, 1)
            TmpAlcohol = Mid(TextLine, 2937, 1)
            TmpSuppExclusion = Trim(Mid(TextLine, 436, 2500))
        ElseIf Mid(txtFileVersion, 1, 2) = "04" Then
            TmpOccupation = Trim(Mid(TextLine, 350, 30))
            TmpWorkNature = Mid(TextLine, 380, 1)
            TmpBlood = Mid(TextLine, 381, 3)
            TmpAllergic = Trim(Mid(TextLine, 384, 20))
            TmpSmoking = Mid(TextLine, 2966, 1)
            TmpAlcohol = Mid(TextLine, 2967, 1)
            TmpSuppExclusion = Trim(Mid(TextLine, 466, 2500))
        End If
        
        TmpPrevPlnCode = Trim(Mid(TextLine, 3433, 15))
        TmpPrevInsCom = Trim(Mid(TextLine, 3456, 60))
        TmpPayRemarks = Trim(Mid(TextLine, 3516, 60))

        TmpPayNo = Trim(Mid(TextLine, 3413, 8))
        TmpPaySeqNo = Trim(Mid(TextLine, 3421, 2))
        TmpClientNo = Trim(Mid(TextLine, 3423, 8))
        TmpClientID = Trim(Mid(TextLine, 3431, 2))
        TmpRBAmt = Trim(Mid(TextLine, 3448, 8))
        
        TmpstaffDec = Trim(Mid(TextLine, 3050, 1))
        TmpGenWP = Mid(TextLine, 3051, 5)
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
        TmpSmoking = Mid(TextLine, 523, 1)
        TmpAlcohol = Mid(TextLine, 524, 1)
        TmpAllergic = Trim(Mid(TextLine, 381, 20))
        TmpSuppExclusion = Trim(Mid(TextLine, 463, 60))
        
        If Mid(txtFileVersion, 1, 2) = "02" Then
            TmpSmoking = Mid(TextLine, 2965, 1)
            TmpAlcohol = Mid(TextLine, 2966, 1)
            TmpSuppExclusion = Trim(Mid(TextLine, 463, 2500))
        End If
        If Mid(txtFileVersion, 1, 2) = "03" Then
            TmpWorkNature = Mid(TextLine, 407, 1)
            TmpBlood = Trim(Mid(TextLine, 408, 3))
            TmpOccupation = Trim(Mid(TextLine, 377, 30))
            TmpSmoking = Mid(TextLine, 2995, 1)
            TmpAlcohol = Mid(TextLine, 2996, 1)
            TmpAllergic = Trim(Mid(TextLine, 411, 20))
            TmpGenWP = Mid(TextLine, 3080, 5)
            TmpPreExWP = Mid(TextLine, 3085, 5)
            TmpSpeIllWP = Mid(TextLine, 3090, 5)
            TmpSuppExclusion = Trim(Mid(TextLine, 493, 2500))
        End If
        TmpPrevPlnCode = Trim(Mid(TextLine, 3462, 15))
        TmpPrevInsCom = Trim(Mid(TextLine, 3485, 60))
        TmpPayRemarks = Trim(Mid(TextLine, 3545, 60))

        TmpPayNo = Trim(Mid(TextLine, 3442, 8))
        TmpPaySeqNo = Trim(Mid(TextLine, 3450, 2))
        TmpClientNo = Trim(Mid(TextLine, 3452, 8))
        TmpClientID = Trim(Mid(TextLine, 3460, 2))
        TmpRBAmt = Mid(TextLine, 3477, 8)
        TmpstaffDec = Trim(Mid(TextLine, 3079, 1))
        TmpPayorPremNet = IIf(IsNumeric(Trim(Mid(TextLine, 3713, 12))), Trim(Mid(TextLine, 3713, 12)), 0)
        TmpPayorPremGross = IIf(IsNumeric(Trim(Mid(TextLine, 3725, 12))), Trim(Mid(TextLine, 3725, 12)), 0)
        TmpPayorMcoNet = IIf(IsNumeric(Trim(Mid(TextLine, 3737, 12))), Trim(Mid(TextLine, 3737, 12)), 0)
        TmpPayorMCoGross = IIf(IsNumeric(Trim(Mid(TextLine, 3749, 12))), Trim(Mid(TextLine, 3749, 12)), 0)
        LoadingPrem = IIf(IsNumeric(Trim(Mid(TextLine, 3793, 12))), Trim(Mid(TextLine, 3793, 12)), 0)
        FamDiscAmt = IIf(IsNumeric(Trim(Mid(TextLine, 3805, 12))), Trim(Mid(TextLine, 3805, 12)), 0)
        RenewDiscAmt = IIf(IsNumeric(Trim(Mid(TextLine, 3817, 12))), Trim(Mid(TextLine, 3817, 12)), 0)
    End If
    TmpSuppExclusion = Replace(TmpSuppExclusion, "'", "''")
    TmpOccupation = Replace(TmpOccupation, "'", "''")
    TmpPrevInsCom = Replace(TmpPrevInsCom, "'", "''")
    TmpSalutation = Replace(TmpSalutation, "'", "''")
    StrDateSql = ""
    StrWDateSql = ""
    If IsNumeric(TmpGenWP) Then
        TmpGenEndWP = Format(DateAdd("d", TmpGenWP, PayEffDate), "yyyy/mm/dd")
        StrDateSql = StrDateSql & ", MBMCGenEndWP "
        StrWDateSql = StrWDateSql & ", '" & TmpGenEndWP & "'"
    End If
    If IsNumeric(TmpPreExWP) Then
        TmpPreExEndWP = Format(DateAdd("d", TmpPreExWP, PayEffDate), "yyyy/mm/dd")
        StrDateSql = StrDateSql & ", MBMCPreExEndWP "
        StrWDateSql = StrWDateSql & ", '" & TmpPreExEndWP & "'"
    End If
    If IsNumeric(TmpSpeIllWP) Then
        TmpSpeIllEndWP = Format(DateAdd("d", TmpSpeIllWP, PayEffDate), "yyyy/mm/dd")
        StrDateSql = StrDateSql & ", MBMSpeIllEndWP "
        StrWDateSql = StrWDateSql & ", '" & TmpSpeIllEndWP & "'"
    End If
    
    TmpRBAmt = CDbl(IIf(IsNumeric(TmpRBAmt) = True, TmpRBAmt, 0))

    TmpCovID = Format(CovCnt, "00")
    Supp2WSql = "Insert Into MBMCoveredPersonsTwo(MBMCNumber, MBMCCoverID, MBMCSalutation, " & _
                "MBMCOccupation, MBMCWorkNature, MBMCBloodType, MBMCAllergic, MBMCSmoking,  " & _
                "MBMCAlcohol, MBMCPrevPlnCode, MBMCPrevInsCom, MBMCPayRemarks, MBMCExclusion, " & _
                "MBMCPayNo, MBMCPaySeqNo, MBMCClientNo, MBMCClientID, MBMCRBAmt, MBMCstaffDec, " & _
                "MBMCPayorPremNet,  MBMCPayorMcoNet, MBMCPayorMCoGross,  " & _
                "MBMCIcBcPp2nd, MBMCLoadingPrem, MBMCFamDiscAmt, MBMCRenewDiscAmt,MBMCPayorPremGross "
    If Trim(StrDateSql) <> "" Then
       Supp2WSql = Supp2WSql & StrDateSql
    End If
        Supp2WSql = Supp2WSql & ")"
        Supp2WSql = Supp2WSql & "Values('" & Trim(txtTempMemberNo) & "','" & TmpCovID & "','" & TmpSalutation & "', " & _
                "'" & TmpOccupation & " ','" & TmpWorkNature & "','" & TmpBlood & "','" & TmpAllergic & "','" & TmpSmoking & "', " & _
                "'" & TmpAlcohol & "','" & TmpPrevPlnCode & "','" & TmpPrevInsCom & "','" & TmpPayRemarks & "','" & TmpSuppExclusion & "', " & _
                "'" & TmpPayNo & "','" & TmpPaySeqNo & "','" & TmpClientNo & "','" & TmpClientID & "'," & TmpRBAmt & ",'" & TmpstaffDec & "', " & _
                "" & TmpPayorPremNet & ", " & TmpPayorMcoNet & ", " & TmpPayorMCoGross & ", " & _
                "'" & TmpIcBcPp2nd & "', " & LoadingPrem & ", " & FamDiscAmt & ", " & RenewDiscAmt & ", " & TmpPayorPremGross & " "
    If Trim(StrWDateSql) <> "" Then
        Supp2WSql = Supp2WSql & StrWDateSql
    End If
        Supp2WSql = Supp2WSql & ")"
    medixmasterconn.Execute Supp2WSql
End Sub
Private Function Get23Age(EffectiveDate As Date, DateOfBirth As Date)
    Dim CheckDate As Long
    Dim CheckYear As Integer
    Dim CheckDOB As Integer
    
    CheckDate = Year(EffectiveDate) - Year(DateOfBirth)
    
    If CheckDate > 23 Then
        Get23Age = CheckDate
    Else
        Get23Age = CheckDate
        If Month(DateOfBirth) > Month(EffectiveDate) Then
            Get23Age = CheckDate - 1
        Else
            If Day(DateOfBirth) > Day(EffectiveDate) Then
                Get23Age = CheckDate - 1
            End If
        End If
    End If
End Function

Private Sub SaveMBMXRef(XNumber, XCovID, XPolNo, XIC, XName, XEffdate, XExpDate, XAgeCode, XInsCode, XPlan, XBordxDate, XTopupNo, XTopupIND, XDataStatus, XStatus, XDOb, XHltCode, XPayor, XBatchNo)
Dim XrefWSql As String
    XPolNo = Replace(XPolNo, "'", "''")
    XName = Replace(XName, "'", "''")
    
    XrefWSql = "Insert Into MBMCrossReference(MBMNumber, MBMCOVID, MBMPolicyNo, MBMIcBcPp, " & _
              "MBMName, MBMPayorEffDate, MBMPayorEXPDate, AgeCode, INSCode, PLNCode, " & _
              "MBMBordxDate, MBMTopupNumber, MBMTopupInd, " & _
              "MBMDataStatus, MBMStatus, MBMDOB, HLTCode, PAYCode, MBMBatchNo)  " & _
              "Values('" & XNumber & "', '" & Format(XCovID, "00") & "', '" & XPolNo & "', '" & XIC & "'," & _
              "'" & XName & "','" & XEffdate & "','" & XExpDate & "', '" & XAgeCode & "','" & XInsCode & "','" & XPlan & "'," & _
              "'" & XBordxDate & "','" & XTopupNo & "', '" & XTopupIND & "'," & _
              "'" & XDataStatus & "','" & XStatus & "', '" & XDOb & "','" & XHltCode & "', '" & XPayor & "', '" & Trim(XBatchNo) & "')"
    medixmasterconnMaint.Execute XrefWSql
End Sub

Private Sub txtPAYCode_LostFocus()
Dim PayRec As New ADODB.Recordset
Dim StrSql As String
    txtClnStatus = "N"
    If Trim(txtPAYCode) <> "" Then
        StrSql = "Select CLNStatus From PAY where PAYCode ='" & Mid(Trim(txtPAYCode), 1, 2) & "'"
        PayRec.Open StrSql, medixmasterconn, adOpenKeyset, adLockOptimistic
        If PayRec.recordCount Then
           txtClnStatus = IIf(Len(Trim(PayRec!CLNStatus)), Trim(PayRec!CLNStatus), "N")
        End If
        PayRec.Close
        Set PayRec = Nothing
    End If
End Sub

Private Sub VerifyMBMByPassingPA(tmpProdType, tmpPolNO, tmpName, tmpDataType, TmpPlanCode, TmpInsType, TmpEffDate, TmpExpDate As String)
Dim CoverDays As String

    If Trim(TmpInsType) <> "I" And Trim(TmpInsType) <> "H" And Trim(TmpInsType) <> "G" And Trim(TmpInsType) <> "F" Then
        Call CreateErrorList(recordCount, tmpPolNO, tmpName, "Invalid " & tmpProdType & " Insured Type")
        txtTotalError = txtTotalError + 1
        Call ErrorListing
    End If
    
    If IsDate(Trim(TmpEffDate)) = False Then
        Call CreateErrorList(recordCount, tmpPolNO, tmpName, "Invalid " & tmpProdType & " Date format ~ Effective Date")
        txtTotalError = txtTotalError + 1
        Call ErrorListing
    End If
                        
    If IsDate(Trim(TmpExpDate)) = False Then
        Call CreateErrorList(recordCount, tmpPolNO, tmpName, "Invalid " & tmpProdType & " Date format ~ Expiry Date ")
        txtTotalError = txtTotalError + 1
        Call ErrorListing
    End If
    
    If IsDate(Trim(TmpEffDate)) And IsDate(Trim(TmpExpDate)) Then
        CoverDays = CDate(TmpExpDate) - CDate(TmpEffDate)
        If CoverDays < 30 Or CoverDays > 365 Then
            Call CreateErrorList(recordCount, tmpPolNO, tmpName, tmpProdType & " Policy Coverage must be > 30 and < 365 days ")
            txtTotalError = txtTotalError + 1
            Call ErrorListing
        End If
    End If


End Sub
