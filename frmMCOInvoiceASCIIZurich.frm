VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form frmMCOInvoiceASCIIZurich 
   Caption         =   "Form1"
   ClientHeight    =   7185
   ClientLeft      =   60
   ClientTop       =   510
   ClientWidth     =   11685
   LinkTopic       =   "Form1"
   ScaleHeight     =   7185
   ScaleWidth      =   11685
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame8 
      Height          =   1215
      Left            =   120
      TabIndex        =   0
      Top             =   240
      Width           =   11535
      Begin VB.ComboBox CboExportStatus 
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
         Left            =   7680
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   3
         Top             =   1200
         Visible         =   0   'False
         Width           =   1335
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
         Left            =   4680
         Sorted          =   -1  'True
         TabIndex        =   1
         Top             =   240
         Width           =   4575
      End
      Begin MSComDlg.CommonDialog CommonDialog2 
         Left            =   10680
         Top             =   600
         _ExtentX        =   847
         _ExtentY        =   847
         _Version        =   393216
      End
      Begin MSMask.MaskEdBox txtReceiptDate 
         Height          =   285
         Left            =   4680
         TabIndex        =   2
         Top             =   600
         Width           =   1305
         _ExtentX        =   2302
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
      Begin MSMask.MaskEdBox txttodate 
         Height          =   285
         Left            =   6480
         TabIndex        =   4
         Top             =   600
         Width           =   1305
         _ExtentX        =   2302
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
      Begin VB.Label txtFilName 
         BackColor       =   &H00FFFFC0&
         Height          =   255
         Left            =   1680
         TabIndex        =   14
         Top             =   840
         Width           =   1905
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
         Left            =   360
         TabIndex        =   13
         Top             =   900
         Width           =   1245
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
         Left            =   3840
         TabIndex        =   12
         Top             =   285
         Width           =   915
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
         Left            =   360
         TabIndex        =   11
         Top             =   615
         Width           =   765
      End
      Begin VB.Label txtPath 
         BackColor       =   &H00FFFFC0&
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
         Left            =   1680
         TabIndex        =   10
         Top             =   240
         Width           =   1905
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
         Left            =   360
         TabIndex        =   9
         Top             =   315
         Width           =   1125
      End
      Begin VB.Label Label26 
         Caption         =   "Bordx Date"
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
         Left            =   3840
         TabIndex        =   8
         Top             =   645
         Width           =   915
      End
      Begin VB.Label txtLocalFileLocation 
         BackColor       =   &H00FFFFC0&
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
         Left            =   1680
         TabIndex        =   7
         Top             =   540
         Width           =   1905
      End
      Begin VB.Label Label1 
         Caption         =   "Export Status"
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
         Left            =   6480
         TabIndex        =   6
         Top             =   1200
         Visible         =   0   'False
         Width           =   1035
      End
      Begin VB.Label Label26 
         Caption         =   "To"
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
         Left            =   6120
         TabIndex        =   5
         Top             =   645
         Width           =   315
      End
   End
   Begin MSComctlLib.ListView lstMBM 
      Height          =   3855
      Left            =   120
      TabIndex        =   42
      Top             =   3240
      Width           =   11535
      _ExtentX        =   20346
      _ExtentY        =   6800
      View            =   3
      Sorted          =   -1  'True
      MultiSelect     =   -1  'True
      LabelWrap       =   0   'False
      HideSelection   =   0   'False
      AllowReorder    =   -1  'True
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
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
      NumItems        =   19
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "MBMFileName"
         Object.Width           =   1764
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   1
         Text            =   "PolicyNumber"
         Object.Width           =   2117
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   2
         Text            =   "ICNumber"
         Object.Width           =   1411
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Text            =   "Name"
         Object.Width           =   1764
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   4
         Text            =   "PlanCode"
         Object.Width           =   1764
      EndProperty
      BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   5
         Key             =   "MBMNo"
         Text            =   "Effective Date"
         Object.Width           =   2469
      EndProperty
      BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   6
         Text            =   "Expiry Date"
         Object.Width           =   1764
      EndProperty
      BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   7
         Text            =   "Days"
         Object.Width           =   2822
      EndProperty
      BeginProperty ColumnHeader(9) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   8
         Text            =   "MCO Fee"
         Object.Width           =   1764
      EndProperty
      BeginProperty ColumnHeader(10) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   9
         Text            =   "Refund Fee"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(11) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   10
         Text            =   "Endorsement No"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(12) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   11
         Text            =   "Invoice Number"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(13) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   12
         Text            =   "Invoice Date"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(14) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   13
         Text            =   "FileType"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(15) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   14
         Text            =   "MBMNumber"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(16) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   15
         Text            =   "MBMCoverID"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(17) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   16
         Text            =   "GST"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(18) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   17
         Text            =   "GSTPerc"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(19) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   18
         Text            =   "GSTCancel"
         Object.Width           =   2540
      EndProperty
   End
   Begin VB.Frame Frame1 
      Height          =   1695
      Left            =   120
      TabIndex        =   15
      Top             =   1440
      Width           =   11535
      Begin VB.CommandButton CmdExit 
         Cancel          =   -1  'True
         Caption         =   "E&xit"
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
         Left            =   10560
         TabIndex        =   20
         Top             =   855
         Width           =   840
      End
      Begin VB.CommandButton cmdGenerate 
         Caption         =   "&OK"
         Default         =   -1  'True
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
         Left            =   9720
         TabIndex        =   19
         Top             =   855
         Width           =   840
      End
      Begin VB.TextBox TxtInvoiceNo 
         Height          =   285
         Left            =   6120
         TabIndex        =   18
         Top             =   600
         Visible         =   0   'False
         Width           =   855
      End
      Begin VB.TextBox txtTotalFee 
         Height          =   285
         Left            =   7200
         TabIndex        =   17
         Top             =   600
         Visible         =   0   'False
         Width           =   735
      End
      Begin VB.CommandButton CmdPrint 
         Caption         =   "&Print"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   8880
         TabIndex        =   16
         Top             =   855
         Width           =   855
      End
      Begin VB.CommandButton CmdClear 
         Caption         =   "&Clear"
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
         Left            =   10320
         TabIndex        =   21
         Top             =   840
         Visible         =   0   'False
         Width           =   840
      End
      Begin VB.Label Label17 
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
         Index           =   0
         Left            =   120
         TabIndex        =   41
         Top             =   270
         Width           =   1080
      End
      Begin VB.Label Label16 
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
         Index           =   0
         Left            =   9840
         TabIndex        =   40
         Top             =   1200
         Visible         =   0   'False
         Width           =   1560
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
         Left            =   1320
         TabIndex        =   39
         Top             =   240
         Width           =   960
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
         Left            =   11520
         TabIndex        =   38
         Top             =   1200
         Width           =   960
      End
      Begin VB.Label Label16 
         Caption         =   "Total Cancellation Fee"
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
         Index           =   1
         Left            =   2400
         TabIndex        =   37
         Top             =   960
         Width           =   1920
      End
      Begin VB.Label Label17 
         Caption         =   "Total MCO Fee"
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
         Index           =   1
         Left            =   2400
         TabIndex        =   36
         Top             =   270
         Width           =   1200
      End
      Begin VB.Label lblMCOFees 
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
         Left            =   4080
         TabIndex        =   35
         Top             =   240
         Width           =   960
      End
      Begin VB.Label lblMCOCFee 
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
         Left            =   4080
         TabIndex        =   34
         Top             =   960
         Width           =   960
      End
      Begin VB.Label Label17 
         Caption         =   "Total Reins. Fee"
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
         Index           =   2
         Left            =   2400
         TabIndex        =   33
         Top             =   600
         Width           =   1200
      End
      Begin VB.Label lblMCORFee 
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
         Left            =   4080
         TabIndex        =   32
         Top             =   600
         Width           =   960
      End
      Begin VB.Label Label6 
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
         Left            =   11520
         TabIndex        =   31
         Top             =   1440
         Width           =   960
      End
      Begin VB.Label Label17 
         Caption         =   "Total Invoice Amount"
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
         Index           =   3
         Left            =   9840
         TabIndex        =   30
         Top             =   1470
         Width           =   1200
      End
      Begin VB.Label lblRPrin 
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
         TabIndex        =   29
         Top             =   600
         Width           =   960
      End
      Begin VB.Label Label17 
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
         Index           =   4
         Left            =   120
         TabIndex        =   28
         Top             =   630
         Width           =   1080
      End
      Begin VB.Label Label17 
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
         Index           =   5
         Left            =   120
         TabIndex        =   27
         Top             =   990
         Width           =   1080
      End
      Begin VB.Label lblCPrin 
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
         TabIndex        =   26
         Top             =   960
         Width           =   960
      End
      Begin VB.Label txtSuppamount 
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
         Left            =   4080
         TabIndex        =   25
         Top             =   1320
         Width           =   960
      End
      Begin VB.Label Label17 
         Caption         =   "Total Supp MCO Fee"
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
         Index           =   6
         Left            =   2400
         TabIndex        =   24
         Top             =   1350
         Width           =   1680
      End
      Begin VB.Label TxtSuppNumber 
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
         TabIndex        =   23
         Top             =   1320
         Width           =   960
      End
      Begin VB.Label Label17 
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
         Index           =   7
         Left            =   120
         TabIndex        =   22
         Top             =   1350
         Width           =   1080
      End
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      BackColor       =   &H00404040&
      Caption         =   "MCO Invoice ASCII"
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
      Height          =   225
      Left            =   0
      TabIndex        =   43
      Top             =   0
      Width           =   11775
   End
End
Attribute VB_Name = "frmMCOInvoiceASCIIZurich"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Exportsql As String
Private textField As String
Private strCount
Private strCount2

Private Sub cmdPrint_Click()
    If lstMBM.ListItems.Count > 0 Then
        Call ExportListViewtoExcel(lstMBM)
    Else
        MsgBox "Please select record!", vbCritical + vbOKOnly, DialogBoxTitle
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

Private Sub CmdClear_Click()
    Call ClearForm(Me)
End Sub

Private Sub CmdGenerate_Click()
Dim TotalPrincipal As Integer
Dim TotalSupplementary As Integer
Dim RecordSaved
Dim startTrans As Boolean
    cmdGenerate.Enabled = False
    startTrans = False
  '  medixmasterconn.Open "File Name=" & BOSPath & "Conn\mediconnect.UDL"
    
    If cboPAYCode = "" Then
        MsgBox "Please Enter Payor Code", vbOKOnly + vbCritical, DialogBoxTitle
        cboPAYCode.SetFocus
    ElseIf txtReceiptDate = "__/__/____" Then
        MsgBox "Please Enter Bordx Date ", vbOKOnly + vbCritical, DialogBoxTitle
        txtReceiptDate.SetFocus
    ElseIf txttodate = "__/__/____" Then
        MsgBox "Please Enter Bordx Date ", vbOKOnly + vbCritical, DialogBoxTitle
        txttodate.SetFocus
    Else
      RecordSaved = MsgBox("Are you sure?", vbYesNo + vbExclamation, DialogBoxTitle)
        If RecordSaved = vbYes Then
    
            Screen.MousePointer = vbHourglass
                startTrans = True
                medixmasterconn.beginTrans
                Call PrincipalASCII
                If startTrans = True Then
                    medixmasterconn.commitTrans
                End If
            Screen.MousePointer = vbDefault
        End If
    cmdGenerate.Enabled = True
    End If
    cmdGenerate.Enabled = True
    Exit Sub
End Sub

Private Sub PrincipalASCII()
Dim strsql As String
Dim MBM As New ADODB.Recordset
Dim BordxDate As String
Dim PAYCode As String
Dim TotalPrincipal As Integer
Dim Header As String
Dim ii As Integer
Dim DataString As String
Dim Trailer As String
Dim rst1 As New ADODB.Recordset
Dim rst2 As New ADODB.Recordset
Dim strCriteria As String
'Dim strCriteria1 As String
Dim Prm1 As New ADODB.Parameter
'Dim CovPrm1 As New ADODB.Parameter
Dim cmd As New ADODB.Command
Dim cmd1 As New ADODB.Command
Dim cmd2 As New ADODB.Command
Dim cmd3 As New ADODB.Command
Dim cmd4 As New ADODB.Command
Dim cmd5 As New ADODB.Command
Dim cmd6 As New ADODB.Command

Dim Suppcmd As New ADODB.Command
Dim SuppPrm As New ADODB.Parameter
Dim Supprst As New ADODB.Recordset
Dim Prm2 As New ADODB.Parameter
Dim Prm3 As New ADODB.Parameter
Dim prm4 As New ADODB.Parameter
Dim Prm5 As New ADODB.Parameter
Dim prm6 As New ADODB.Parameter

Dim TotalP As Integer
Dim TotalSupp As Integer
Dim TotalMCO As Double
Dim TotalMCOSupp As Double
Dim TotalR As Integer
Dim TotalRMCO As Double
Dim TotalC As Integer
Dim TotalCMCO As Double
Dim Record As ListItem
Dim Record1 As ListItem
Dim tmpInvoiceDate As String
Dim tmpDays As String
Dim tmpTrailer As String
Dim AA As Integer
    AA = 0
    tmpInvoiceDate = Format(GetServerTime, "DD/MM/YYYY")
    TotalP = 0
    TotalR = 0
    TotalC = 0
    tmpTrailer = 0
    strCriteria = ""
   
        Screen.MousePointer = vbHourglass
        lstMBM.ListItems.Clear
        strCriteria = ""
        'strCriteria1 = ""
        'strCriteria = strCriteria & " AND dbo.MBM.PayCode =  '" & Mid(cboPAYCode, 1, 2) & "' AND MBMBordxDate >=  '" & Format(txtReceiptDate, "yyyy/mm/dd") & "' AND MBMBordxDate <=  '" & Format(txttodate.Text, "yyyy/mm/dd") & "'"
        strCriteria = strCriteria & " AND  dbo.ZurichPay.PAYGRPCode =  '" & Mid(cboPAYCode, 1, 3) & "' AND MBMBordxDate >=  '" & Format(txtReceiptDate, "yyyy/mm/dd") & "' AND MBMBordxDate <=  '" & Format(txttodate.Text, "yyyy/mm/dd") & "'"
        
       ' strCriteria1 = strCriteria1 & " AND  dbo.ZurichPay.PAYGRPCode =  '" & Mid(cboPAYCode, 1, 3) & "' AND MBMCoveredPersons.MBMCBordxDate >=  '" & Format(txtReceiptDate, "yyyy/mm/dd") & "' AND MBMCoveredPersons.MBMCBordxDate <=  '" & Format(txttodate.Text, "yyyy/mm/dd") & "'"
        
        
        'strCriteria = strCriteria & " AND MBMBordxDate =  '" & Format(txtReceiptDate, "yyyy/mm/dd") & "'"
        Set cmd.ActiveConnection = medixmasterconn
        cmd.CommandText = "SP_MCOASCII_New"
        cmd.CommandType = adCmdStoredProc
        cmd.CommandTimeout = 300
        Set Prm2 = cmd.CreateParameter("SelCriterial", adVarChar, adParamInput, 4000, strCriteria)
        cmd.Parameters.Append Prm2
        'Set CovPrm1 = cmd.CreateParameter("SelCriterial1", adVarChar, adParamInput, 4000, strCriteria1)
        'cmd.Parameters.Append CovPrm1
        rst2.CursorLocation = adUseClient
        rst2.CursorType = adOpenForwardOnly
        rst2.CacheSize = 100
        Dim rCnt As Integer
        rCnt = 0
        Set rst2 = cmd.Execute
        Do Until rst2.EOF
            With rst2
            rCnt = rCnt + 1
                Set Record = lstMBM.ListItems.Add(, , IIf(Len(rst2!MBMFilename), rst2!MBMFilename, ""))
                If Len(rst2!MBMPolicyNo) Then
                   Record.SubItems(1) = rst2!MBMPolicyNo
                End If
                If Len(rst2!MBMIcBcPp) Then
                   Record.SubItems(2) = rst2!MBMIcBcPp
                End If
                If Len(rst2!MBMName) Then
                   Record.SubItems(3) = rst2!MBMName
                End If
                If Len(rst2!PLNPayorPlan) Then
                   Record.SubItems(4) = rst2!PLNPayorPlan
                End If
                If Len(rst2!MBMPayorEffDate) Then
                   Record.SubItems(5) = rst2!MBMPayorEffDate
                End If
                If Len(rst2!MBMPayorEXPDate) Then
                   Record.SubItems(6) = rst2!MBMPayorEXPDate
                End If
                If (Len(rst2!MBMPayorEXPDate) > 0) Then
                Record.SubItems(7) = (rst2!MBMPayorEXPDate - rst2!MBMPayorEffDate) + 1
                End If
                If Len(rst2!MBMMCOFee) Then
                    Record.SubItems(8) = Format(rst2!MBMMCOFee, "#,##0.00")
                End If
                If Len(rst2!MBMENDtRefNo) Then
                    Record.SubItems(10) = rst2!MBMENDtRefNo
                End If
                Record.SubItems(13) = "N"
                Record.SubItems(14) = rst2!MBMNumber
                Record.SubItems(15) = rst2!MBMCCoverID
                Record.SubItems(16) = IIf(Len(rst2!MBMGSTAmt), rst2!MBMGSTAmt, 0)
                Record.SubItems(17) = IIf(Len(rst2!MBMGSTPerc), rst2!MBMGSTPerc, 0)
                Record.SubItems(8) = "0"
                
                TotalP = TotalP + 1
                If Len(rst2!MBMMCOFee) Then
                    TotalMCO = TotalMCO + rst2!MBMMCOFee
                End If
            End With
            rst2.MoveNext
            Loop
            rst2.Close
            Set rst2 = Nothing
            txtTotalPrincipals = TotalP
            lblMCOFees = TotalMCO
            
        'Cancellation
        strCriteria = ""
        'strCriteria = strCriteria & " AND dbo.MBM.PayCode =  '" & Mid(cboPAYCode, 1, 2) & "' AND MBMRefund.MBMBordxDate >=  '" & Format(txtReceiptDate, "yyyy/mm/dd") & "' AND MBMRefund.MBMBordxDate <=  '" & Format(txttodate.Text, "yyyy/mm/dd") & "'"
        strCriteria = strCriteria & " AND  dbo.ZurichPay.PAYGRPCode =  '" & Mid(cboPAYCode, 1, 3) & "' AND MBMRefund.MBMBordxDate >=  '" & Format(txtReceiptDate, "yyyy/mm/dd") & "' AND MBMRefund.MBMBordxDate <=  '" & Format(txttodate.Text, "yyyy/mm/dd") & "'"
        Set cmd2.ActiveConnection = medixmasterconn
        'cmd2.CommandText = "SP_MCOASCII_Refund"
        cmd2.CommandText = "SP_MCOASCII_cancellation"
        cmd2.CommandType = adCmdStoredProc
        cmd2.CommandTimeout = 300
        Set prm6 = cmd2.CreateParameter("SelCriterial", adVarChar, adParamInput, 2000, strCriteria)
        cmd2.Parameters.Append prm6
        rst2.CursorLocation = adUseClient
        rst2.CursorType = adOpenForwardOnly
        rst2.CacheSize = 100
        Set rst2 = cmd2.Execute
        rCnt = 0
        Do Until rst2.EOF
            With rst2
            rCnt = rCnt + 1
                Set Record = lstMBM.ListItems.Add(, , !MBMFilename)
                If Len(rst2!MBMPolicyNo) Then
                   Record.SubItems(1) = rst2!MBMPolicyNo
                End If
                If Len(rst2!MBMIcBcPp) Then
                   Record.SubItems(2) = rst2!MBMIcBcPp
                End If
                If Len(rst2!MBMName) Then
                   Record.SubItems(3) = rst2!MBMName
                End If
                If Len(rst2!PLNPayorPlan) Then
                   Record.SubItems(4) = rst2!PLNPayorPlan
                End If
                If Len(rst2!MBMPayorEffDate) Then
                   Record.SubItems(5) = rst2!MBMPayorEffDate
                End If
                If Len(rst2!MBMPayorEXPDate) Then
                   Record.SubItems(6) = rst2!MBMPayorEXPDate
                End If
                If (Len(rst2!MBMPayorEXPDate) > 0) Then
                Record.SubItems(7) = (rst2!MBMPayorEXPDate - rst2!MBMPayorEffDate) + 1
                End If
                If Len(rst2!MBMMCOFee) Then
                    Record.SubItems(9) = Format(rst2!MBMMCOFee, "#,##0.00")
                End If
                If Len(rst2!MBMENDtRefNo) Then
                    Record.SubItems(10) = rst2!MBMENDtRefNo
                End If
                Record.SubItems(13) = "C"
                Record.SubItems(14) = rst2!MBMNumber
                Record.SubItems(16) = 0
                Record.SubItems(17) = IIf(Len(rst2!MBMGSTPerc), rst2!MBMGSTPerc, 0)
                Record.SubItems(18) = IIf(Len(rst2!MBMGSTAmt), rst2!MBMGSTAmt, 0)
                TotalP = TotalP + 1
                If Len(rst2!MBMMCOFee) Then
                    TotalMCO = TotalMCO + rst2!MBMMCOFee
                End If
            End With
            rst2.MoveNext
            Loop
            rst2.Close
            Set rst2 = Nothing
            lblCPrin = TotalC
            lblMCOCFee = TotalCMCO
            
            Call IncInvoiceNo
            
            For AA = 1 To lstMBM.ListItems.Count
                    If Trim(lstMBM.ListItems(AA).SubItems(13)) = "N" Then
                        Call UpdateINVPrintingFile(Trim(lstMBM.ListItems(AA).SubItems(14)), Trim(lstMBM.ListItems(AA).SubItems(15)))
                    ElseIf Trim(lstMBM.ListItems(AA).SubItems(13)) = "E" Then
                        Call UpdateINVReinstate(Trim(lstMBM.ListItems(AA).SubItems(14)))
                    ElseIf Trim(lstMBM.ListItems(AA).SubItems(13)) = "C" Then
                        Call UpdateINVRefund(Trim(lstMBM.ListItems(AA).SubItems(14)))
                        'Call UpdateINVPrintingFile(Trim(lstMBM.ListItems(AA).SubItems(14)))
                    End If
            Next AA
            
            
            
    If lstMBM.ListItems.Count = 0 Then
      '  MsgBox "No record found!", vbOKOnly + vbCritical, DialogBoxTitle
    Else
    
       txtTotalFee = CDbl(TotalMCO) + CDbl(TotalRMCO) + CDbl(TotalCMCO)
       
       Header = "F" & "MCO02" & TxtInvoiceNo & "         " & Mid(tmpInvoiceDate, 1, 2) & Mid(tmpInvoiceDate, 4, 2) & Mid(tmpInvoiceDate, 7, 4) & txtTotalFee
       txtFilName = "F" & "MCO02" & TxtInvoiceNo
       Open txtPath & "\" & txtFilName For Output As #1
       Open txtLocalFileLocation & txtFilName For Output As #2
            
       Print #1, Trim(Header)
       Print #2, Trim(Header)
                    
       AA = 0
       For AA = 1 To lstMBM.ListItems.Count
         DataString = l("I", 1)
         DataString = DataString & l(lstMBM.ListItems(AA).Text, 20)
         DataString = DataString & l(lstMBM.ListItems(AA).SubItems(1), 15)
         DataString = DataString & l(lstMBM.ListItems(AA).SubItems(2), 20)
         DataString = DataString & l(lstMBM.ListItems(AA).SubItems(3), 60)
         DataString = DataString & l(lstMBM.ListItems(AA).SubItems(4), 8)
         DataString = DataString & l(Format(lstMBM.ListItems(AA).SubItems(5), "ddmmyyyy"), 8)
         DataString = DataString & l(Format(lstMBM.ListItems(AA).SubItems(6), "ddmmyyyy"), 8)
         DataString = DataString & l(lstMBM.ListItems(AA).SubItems(7), 3)
         DataString = DataString & l(Format(lstMBM.ListItems(AA).SubItems(8), "#,##0.00"), 12)
         DataString = DataString & l(Format(lstMBM.ListItems(AA).SubItems(9), "#,##0.00"), 12)
         DataString = DataString & l(lstMBM.ListItems(AA).SubItems(10), 25)
         DataString = DataString & l("", 60)
         DataString = DataString & l(Format(lstMBM.ListItems(AA).SubItems(16), "#,##0.00"), 12) & Chr(13)
         
         'If IsNumeric(lstMBM.ListItems(AA).SubItems(16)) Then
         '   If lstMBM.ListItems(AA).SubItems(16) > 0 Then
         '       DataString = DataString & l(Format(lstMBM.ListItems(AA).SubItems(16), "#,##0.00"), 12)
         '       DataString = DataString & l(Format("0", "#,##0.00"), 12)
         '       DataString = DataString & l(Format(lstMBM.ListItems(AA).SubItems(17), "#,##0.00"), 6)
         '   Else
         '       DataString = DataString & l(Format("0", "#,##0.00"), 12)
         '       DataString = DataString & l(Format(lstMBM.ListItems(AA).SubItems(18), "#,##0.00"), 12)
         '       DataString = DataString & l(Format(lstMBM.ListItems(AA).SubItems(17), "#,##0.00"), 6)
         '   End If
         'Else
         '   DataString = DataString & l(Format("0.00", "#,##0.00"), 12)
         '   DataString = DataString & l(Format("0.00", "#,##0.00"), 12)
         '   DataString = DataString & l(Format("0.00", "#,##0.00"), 6)
        '
        ' End If
         
         
         
          
         Print #1, DataString
         Print #2, DataString
            
        Next AA
            
            tmpTrailer = CDbl(TotalP) + CDbl(TotalR) + CDbl(TotalC) + CDbl(TotalSupp)
            Trailer = "T" & l(tmpTrailer, 10)
                'Debug.Print Trailer
                Print #1, Trailer
                Print #2, Trailer
                cmdGenerate.Enabled = False
                Close #1
                Close #2
                
              '  MsgBox "File Generated Sucessfully...", vbOKOnly + vbInformation, DialogBoxTitle
                cmdGenerate.Enabled = True
    End If
    
    TotalP = 0
    TotalR = 0
    TotalC = 0
    TotalSupp = 0
    
     'Reinstatement
      'If Record.ListSubItems.Count > 0 Then
      '  Record.ListSubItems.Clear
      'End If
         lstMBM.ListItems.Clear
         
        strCriteria = ""
        'strCriteria = strCriteria & " AND dbo.MBM.PayCode =  '" & Mid(cboPAYCode, 1, 2) & "' AND MBMBordxDate >=  '" & Format(txtReceiptDate, "yyyy/mm/dd") & "' AND MBMBordxDate <=  '" & Format(txttodate.Text, "yyyy/mm/dd") & "'"
        strCriteria = strCriteria & " AND  dbo.ZurichPay.PAYGRPCode =  '" & Mid(cboPAYCode, 1, 3) & "' AND MBMBordxDate >=  '" & Format(txtReceiptDate, "yyyy/mm/dd") & "' AND MBMBordxDate <=  '" & Format(txttodate.Text, "yyyy/mm/dd") & "'"
        Set cmd3.ActiveConnection = medixmasterconn
        cmd3.CommandText = "SP_MCOASCII_ReNew"
        cmd3.CommandType = adCmdStoredProc
        cmd3.CommandTimeout = 300
        Set Prm5 = cmd3.CreateParameter("SelCriterial", adVarChar, adParamInput, 2000, strCriteria)
        cmd3.Parameters.Append Prm5
        rst2.CursorLocation = adUseClient
        rst2.CursorType = adOpenForwardOnly
        rst2.CacheSize = 100
        Set rst2 = cmd3.Execute
       
         Do Until rst2.EOF
           With rst2
                Set Record1 = lstMBM.ListItems.Add(, , rst2!MBMFilename)
                If Len(rst2!MBMPolicyNo) Then
                   Record1.SubItems(1) = rst2!MBMPolicyNo
                End If
                If Len(rst2!MBMIcBcPp) Then
                   Record1.SubItems(2) = rst2!MBMIcBcPp
                End If
                If Len(rst2!MBMName) Then
                   Record1.SubItems(3) = rst2!MBMName
                End If
                If Len(rst2!PLNPayorPlan) Then
                   Record1.SubItems(4) = rst2!PLNPayorPlan
                End If
                'If Len(rst2!MBMReinstateDate) Then
                '   Record1.SubItems(5) = rst2!MBMReinstateDate
                'End If
                'MBMPayorEffDate
                If Len(rst2!MBMPayorEffDate) Then
                   Record1.SubItems(5) = rst2!MBMPayorEffDate
                End If
                If Len(rst2!MBMPayorEXPDate) Then
                   Record1.SubItems(6) = rst2!MBMPayorEXPDate
                End If
                Record1.SubItems(7) = (rst2!MBMPayorEXPDate - rst2!MBMPayorEffDate) + 1
                If Len(rst2!MBMMCOFee) Then
                   Record1.SubItems(8) = Format(rst2!MBMMCOFee, "#,##0.00")
                End If
                If Len(rst2!MBMENDtRefNo) Then
                    Record1.SubItems(10) = rst2!MBMENDtRefNo
                End If
                Record1.SubItems(13) = "N"
                Record1.SubItems(14) = rst2!MBMNumber
                Record1.SubItems(15) = rst2!MBMCCoverID
                
                Record1.SubItems(16) = IIf(Len(rst2!MBMGSTAmt), rst2!MBMGSTAmt, 0)
                Record1.SubItems(17) = IIf(Len(rst2!MBMGSTPerc), rst2!MBMGSTPerc, 0)
                Record1.SubItems(18) = 0
                TotalR = TotalR + 1
                If Len(rst2!MBMMCOFee) Then
                    TotalRMCO = TotalRMCO + rst2!MBMMCOFee
                End If
            End With
            rst2.MoveNext
            Loop
            rst2.Close
            Set rst2 = Nothing
           
        strCriteria = ""
        'strCriteria = strCriteria & " AND dbo.MBM.PayCode =  '" & Mid(cboPAYCode, 1, 2) & "' AND MBMBordxDate >=  '" & Format(txtReceiptDate, "yyyy/mm/dd") & "' AND MBMBordxDate <=  '" & Format(txttodate.Text, "yyyy/mm/dd") & "'"
         strCriteria = strCriteria & " AND  dbo.ZurichPay.PAYGRPCode =  '" & Mid(cboPAYCode, 1, 3) & "' AND MBMBordxDate >=  '" & Format(txtReceiptDate, "yyyy/mm/dd") & "' AND MBMBordxDate <=  '" & Format(txttodate.Text, "yyyy/mm/dd") & "'"
        Set cmd4.ActiveConnection = medixmasterconn
        cmd4.CommandText = "SP_MCOASCII_Endorsmnt"
        cmd4.CommandType = adCmdStoredProc
        cmd4.CommandTimeout = 300
        Set Prm3 = cmd4.CreateParameter("SelCriterial", adVarChar, adParamInput, 2000, strCriteria)
        cmd4.Parameters.Append Prm3
        rst2.CursorLocation = adUseClient
        rst2.CursorType = adOpenForwardOnly
        rst2.CacheSize = 100
        Set rst2 = cmd4.Execute
          Do Until rst2.EOF
            With rst2
                Set Record1 = lstMBM.ListItems.Add(, , rst2!MBMFilename)
                If Len(rst2!MBMPolicyNo) Then
                   Record1.SubItems(1) = rst2!MBMPolicyNo
                End If
                If Len(rst2!MBMIcBcPp) Then
                   Record1.SubItems(2) = rst2!MBMIcBcPp
                End If
               If Len(rst2!MBMName) Then
                   Record1.SubItems(3) = rst2!MBMName
                End If
                If Len(rst2!PLNPayorPlan) Then
                   Record1.SubItems(4) = rst2!PLNPayorPlan
                End If
                If Len(rst2!MBMPayorEffDate) Then
                   Record1.SubItems(5) = rst2!MBMPayorEffDate
                End If
                If Len(rst2!MBMPayorEXPDate) Then
                   Record1.SubItems(6) = rst2!MBMPayorEXPDate
                End If
               Record1.SubItems(7) = (rst2!MBMPayorEXPDate - rst2!MBMPayorEffDate) + 1
                If Len(rst2!MBMMCOFee) Then
                   Record1.SubItems(8) = Format(rst2!MBMMCOFee, "#,##0.00")
                End If
                If Len(rst2!MBMENDtRefNo) Then
                    Record1.SubItems(10) = rst2!MBMENDtRefNo
                End If
                Record1.SubItems(13) = "N"
                Record1.SubItems(14) = rst2!MBMNumber
                Record1.SubItems(16) = 0
                Record1.SubItems(17) = 0
                Record1.SubItems(18) = 0
                
                TotalR = TotalR + 1
                If Len(rst2!MBMMCOFee) Then
                    TotalRMCO = TotalRMCO + rst2!MBMMCOFee
                End If
            End With
           rst2.MoveNext
            Loop
        
            rst2.Close
            Set rst2 = Nothing
            
            
        strCriteria = ""
        'strCriteria = strCriteria & " AND PayCode =  '" & Mid(cboPAYCode, 1, 2) & "' AND MBMEndtBordxDate >=  '" & Format(txtReceiptDate, "yyyy/mm/dd") & "' AND MBMEndtBordxDate <=  '" & Format(txttodate.Text, "yyyy/mm/dd") & "'"
        strCriteria = strCriteria & " AND  PAYGRPCode =  '" & Mid(cboPAYCode, 1, 3) & "' AND MBMEndtBordxDate >=  '" & Format(txtReceiptDate, "yyyy/mm/dd") & "' AND MBMEndtBordxDate <=  '" & Format(txttodate.Text, "yyyy/mm/dd") & "'"
        Set cmd1.ActiveConnection = medixmasterconn
        cmd1.CommandText = "SP_MCOASCII_Reinstatement"
        'cmd1.CommandText = "SP_MCOASCII_NewTwo"
        cmd1.CommandType = adCmdStoredProc
        cmd1.CommandTimeout = 300
        Set prm4 = cmd1.CreateParameter("SelCriterial", adVarChar, adParamInput, 2000, strCriteria)
        cmd1.Parameters.Append prm4
        rst2.CursorLocation = adUseClient
        rst2.CursorType = adOpenForwardOnly
        rst2.CacheSize = 100
        Set rst2 = cmd1.Execute
        Set Record1 = Nothing
        Do Until rst2.EOF
            With rst2
                Set Record1 = lstMBM.ListItems.Add(, , rst2!MBMFilename)
                If Len(rst2!MBMPolicyNo) Then
                   Record1.SubItems(1) = rst2!MBMPolicyNo
                End If
                If Len(rst2!MBMIcBcPp) Then
                   Record1.SubItems(2) = rst2!MBMIcBcPp
                End If
                If Len(rst2!MBMName) Then
                   Record1.SubItems(3) = rst2!MBMName
                End If
                If Len(rst2!PLNPayorPlan) Then
                   Record1.SubItems(4) = rst2!PLNPayorPlan
                End If
                If Len(rst2!MBMReinstateDate) Then
                   Record1.SubItems(5) = rst2!MBMReinstateDate
                End If
                If Len(rst2!MBMPayorEXPDate) Then
                   Record1.SubItems(6) = rst2!MBMPayorEXPDate
                End If
                Record1.SubItems(7) = (rst2!MBMPayorEXPDate - rst2!MBMReinstateDate) + 1
                If Len(rst2!MBMMCOFee) Then
                   Record1.SubItems(8) = Format(rst2!MBMMCOFee, "#,##0.00")
                End If
                If Len(rst2!MBMENDtRefNo) Then
                    Record1.SubItems(10) = rst2!MBMENDtRefNo
                End If
                Record1.SubItems(13) = "E"
                Record1.SubItems(14) = rst2!MBMNumber
                Record1.SubItems(16) = IIf(Len(rst2!MBMGSTAmt), rst2!MBMGSTAmt, 0)
                Record1.SubItems(17) = IIf(Len(rst2!MBMGSTPerc), rst2!MBMGSTPerc, 0)
                Record1.SubItems(18) = 0
                TotalR = TotalR + 1
                If Len(rst2!MBMMCOFee) Then
                    TotalRMCO = TotalRMCO + rst2!MBMMCOFee
                End If
            End With
            rst2.MoveNext
            Loop
            rst2.Close
            Set rst2 = Nothing
            
            
            
            lblRPrin = TotalR
            lblMCORFee = TotalRMCO
            txtTotalFee = CDbl(TotalMCO) + CDbl(TotalRMCO) + CDbl(TotalCMCO)
            
            'Call IncInvoiceNo
             For AA = 1 To lstMBM.ListItems.Count
                    If Trim(lstMBM.ListItems(AA).SubItems(13)) = "R" Then
                        Call UpdateINVPrintingFile(Trim(lstMBM.ListItems(AA).SubItems(14)), Trim(lstMBM.ListItems(AA).SubItems(15)))
                    ElseIf Trim(lstMBM.ListItems(AA).SubItems(13)) = "E" Then
                        Call UpdateINVReinstate(Trim(lstMBM.ListItems(AA).SubItems(14)))
                    ElseIf Trim(lstMBM.ListItems(AA).SubItems(13)) = "C" Then
                        Call UpdateINVRefund(Trim(lstMBM.ListItems(AA).SubItems(14)))
                    ElseIf Trim(lstMBM.ListItems(AA).SubItems(13)) = "N" Then
                        Call UpdateINVPrintingFile(Trim(lstMBM.ListItems(AA).SubItems(14)), Trim(lstMBM.ListItems(AA).SubItems(15)))
                    End If
            Next AA
          
            
    If lstMBM.ListItems.Count <> 0 Then
       Header = "F" & "MCO02" & TxtInvoiceNo & "         " & Mid(tmpInvoiceDate, 1, 2) & Mid(tmpInvoiceDate, 4, 2) & Mid(tmpInvoiceDate, 7, 4) & txtTotalFee
       txtFilName = "F" & "MCO02" & TxtInvoiceNo & "Renew"
       Open txtPath & "\" & txtFilName For Output As #1
       Open txtLocalFileLocation & txtFilName For Output As #2
       
       AA = 0
       For AA = 1 To lstMBM.ListItems.Count
         DataString = l("I", 1)
         DataString = DataString & l(lstMBM.ListItems(AA).Text, 20)
         DataString = DataString & l(lstMBM.ListItems(AA).SubItems(1), 15)
         DataString = DataString & l(lstMBM.ListItems(AA).SubItems(2), 20)
         DataString = DataString & l(lstMBM.ListItems(AA).SubItems(3), 60)
         DataString = DataString & l(lstMBM.ListItems(AA).SubItems(4), 8)
         DataString = DataString & l(Format(lstMBM.ListItems(AA).SubItems(5), "ddmmyyyy"), 8)
         DataString = DataString & l(Format(lstMBM.ListItems(AA).SubItems(6), "ddmmyyyy"), 8)
         DataString = DataString & l(lstMBM.ListItems(AA).SubItems(7), 3)
         DataString = DataString & l(Format(lstMBM.ListItems(AA).SubItems(8), "#,##0.00"), 12)
         DataString = DataString & l(Format(lstMBM.ListItems(AA).SubItems(9), "#,##0.00"), 12)
         DataString = DataString & l(lstMBM.ListItems(AA).SubItems(10), 25)
         DataString = DataString & l("", 60)
         DataString = DataString & l(Format(lstMBM.ListItems(AA).SubItems(16), "#,##0.00"), 12) & Chr(13)
         'If IsNumeric(lstMBM.ListItems(AA).SubItems(16)) Then
         '   If lstMBM.ListItems(AA).SubItems(16) > 0 Then
                'DataString = DataString & l(Format(lstMBM.ListItems(AA).SubItems(16), "#,##0.00"), 25)
                'DataString = DataString & l(Format("0", "#,##0.00"), 12)
                'DataString = DataString & l(Format(lstMBM.ListItems(AA).SubItems(17), "#,##0.00"), 6)
         '   Else
         '       DataString = DataString & l(Format("0", "#,##0.00"), 12)
         '       DataString = DataString & l(Format(lstMBM.ListItems(AA).SubItems(18), "#,##0.00"), 12)
         '       DataString = DataString & l(Format(lstMBM.ListItems(AA).SubItems(17), "#,##0.00"), 6)
         '   End If
         'Else
         '   DataString = DataString & l(Format("0.00", "#,##0.00"), 12)
         '   DataString = DataString & l(Format("0.00", "#,##0.00"), 12)
         '   DataString = DataString & l(Format("0.00", "#,##0.00"), 6)
         
         'End If
         
          
         Print #1, DataString
         Print #2, DataString
            
        Next AA
            
            
       Print #1, Trim(Header)
       Print #2, Trim(Header)
                    
      
            
            tmpTrailer = CDbl(TotalP) + CDbl(TotalR) + CDbl(TotalC) + CDbl(TotalSupp)
            Trailer = "T" & l(tmpTrailer, 10)
                'Debug.Print Trailer
                Print #1, Trailer
                Print #2, Trailer
                cmdGenerate.Enabled = False
                Close #1
                Close #2
                MsgBox "File Generated Sucessfully...", vbOKOnly + vbInformation, DialogBoxTitle
                cmdGenerate.Enabled = True
    Else
        MsgBox "No record found!", vbOKOnly + vbCritical, DialogBoxTitle
    End If
    
    Close #1
    Close #2
    'End If
End Sub

Private Sub UpdateINVPrintingFile(tmpMBMNo As String, Optional tmpMBMCovID As String)
Dim strsql As String
Dim InvoiceDate As String
    
    InvoiceDate = Format(GetServerTime, "YYYY/MM/DD")
    If tmpMBMCovID = "00" Then   'Update file for MBMOthers Table
        strsql = " update MBMOthers set " & _
                 " MBMInvoiceDate = '" & InvoiceDate & "', " & _
                 " MBMInvoiceNo = '" & Trim(TxtInvoiceNo) & "', " & _
                 " MBMInvoiceBy = '" & Logon & "' " & _
                 " Where MBMNumber = '" & Trim(tmpMBMNo) & "'"
                 medixmasterconn.Execute strsql
    Else    'update file for MBMCoveredPersons Table
        strsql = " update MBMCoveredPersons set " & _
                 " MBMInvoiceDate = '" & InvoiceDate & "', " & _
                 " MBMInvoiceNo = '" & Trim(TxtInvoiceNo) & "', " & _
                 " MBMInvoiceBy = '" & Logon & "' " & _
                 " Where MBMCNumber = '" & Trim(tmpMBMNo) & "'" & _
                 " and MBMCCoverID = '" & tmpMBMCovID & "'"
                 medixmasterconn.Execute strsql
    End If
End Sub
Private Sub UpdateINVRefund(tmpMBMNo As String)
Dim strsql As String
Dim InvoiceDate As String
    
    InvoiceDate = Format(GetServerTime, "YYYY/MM/DD")
    'If TmpCovID = "00" Then   'Update file for MBMOthers Table
        strsql = " update MBMRefund set " & _
                 " MBMRefInvoiceDate = '" & InvoiceDate & "', " & _
                 " MBMRefInvoiceNo = '" & Trim(TxtInvoiceNo) & "', " & _
                 " MBMRefInvoiceBy = '" & Logon & "' " & _
                 " Where MBMNumber = '" & Trim(tmpMBMNo) & "'"
                 medixmasterconn.Execute strsql
    'Else    'update file for MBMCoveredPersons Table
    '    strsql = " update MBMCoveredPersons set " & _
    '             " MBMInvoiceDate = '" & InvoiceDate & "', " & _
    '             " MBMInvoiceNo = '" & Trim(TxtInvoiceNo) & "', " & _
    '             " MBMInvoiceBy = '" & Logon & "' " & _
    '             " Where MBMCNumber = '" & Trim(tmpMBMNo) & "'" & _
    '             " and MBMCCoverID = '" & TmpCovID & "'"
    '             medixmasterconn.Execute strsql
    'End If
End Sub

Private Sub UpdateINVReinstate(tmpMBMNo As String)
Dim strsql As String
Dim InvoiceDate As String
    
    InvoiceDate = Format(GetServerTime, "YYYY/MM/DD")
        strsql = " update MBMReinstatement set " & _
                 " MBMInvoiceDate = '" & InvoiceDate & "', " & _
                 " MBMInvoiceNo = '" & Trim(TxtInvoiceNo) & "', " & _
                 " MBMInvoiceBy = '" & Logon & "' " & _
                 " Where MBMNumber = '" & Trim(tmpMBMNo) & "'"
                 medixmasterconn.Execute strsql
End Sub
Private Sub IncInvoiceNo()
Dim MBMMCOINVQ As New ADODB.Recordset
Dim strsql As String
Dim InvoiceNo As String

    strsql = "Select * from MBMMCOINVQ "
    MBMMCOINVQ.Open strsql, medixmasterconn, adOpenKeyset, adLockOptimistic
    
    If MBMMCOINVQ.recordCount = 0 Then
        'With MBMMCOINVQ
        '    .AddNew
        '        !PAYCode = tmpPayor
        '        !Prefix = "S"
        '        !MCOInvoice_No = "00001"
        '        InvoiceNo = "00001"
        '    .Update
        'End With
    Else
        With MBMMCOINVQ
            !prefix = "S"
            !MCOInvoice_No = Format(!MCOInvoice_No + 1, "00000")
            InvoiceNo = Format(!MCOInvoice_No, "00000")
        .Update
        End With
    End If
    TxtInvoiceNo = "S" & InvoiceNo
    MBMMCOINVQ.Close
    Set MBMMCOINVQ = Nothing
End Sub
Private Sub CmdExit_Click()
'Dim RecordExit
'    RecordExit = MsgBox("Quit this form?", vbYesNo + vbExclamation)
'    If RecordExit = vbYes Then
'        bExit = True
        Unload Me
'    End If
End Sub

Private Sub Form_Load()
        
    txtPath = DocPath & "ASCI"
    txtLocalFileLocation = "C:\"
    Call ComboListing
    
End Sub

Private Sub txtReceiptDate_LostFocus()
    If IsDate(txtReceiptDate) Then
    Else
        If txtReceiptDate <> "__/__/____" Then
            MsgBox "Invalid Date Format! ", vbCritical
            txtReceiptDate.SetFocus
        End If
    End If
End Sub

Private Sub cboPAYCode_Change()
If cboPAYCode.Text = "" Then
    If InStr(cboPAYCode.Text, "null") Then
       cboPAYCode.Enabled = False
    End If
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

Private Sub ComboListing()
Dim PAY As New ADODB.Recordset
Dim cmd As New ADODB.Command
    
  '  medixmasterconn.Open "File Name=" & BOSPath & "Conn\mediconnect.UDL"
    Set cmd.ActiveConnection = medixmasterconn
    cmd.CommandText = "SearchPayorMAA"
    cmd.CommandType = adCmdStoredProc
    cmd.CommandTimeout = 15
    PAY.CursorLocation = adUseClient
    PAY.CursorType = adOpenForwardOnly
    PAY.CacheSize = 100
    Set PAY = cmd.Execute
    Do Until PAY.EOF
        With PAY
            cboPAYCode.AddItem !PAYGRPCode & " - " & Trim(!PAYCompanyName)
            PAY.MoveNext
        End With
    Loop
    PAY.Close
    Set PAY = Nothing
    
  '  medixmasterconn.Close
  '  Set medixmasterconn = Nothing
    CboExportStatus.AddItem "N - New"
    CboExportStatus.AddItem "R - Reprint"
    CboExportStatus.Text = "N - New"
End Sub
