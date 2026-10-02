VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "mscomctl.ocx"
Begin VB.Form frmStorageDoc 
   Caption         =   "Form1"
   ClientHeight    =   7935
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   11880
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   7935
   ScaleWidth      =   11880
   WindowState     =   2  'Maximized
   Begin VB.Frame Frame2 
      Caption         =   "Search"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1455
      Left            =   120
      TabIndex        =   40
      Top             =   240
      Width           =   10335
      Begin VB.ComboBox cboDOCType 
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
         ItemData        =   "frmStorageDoc.frx":0000
         Left            =   1680
         List            =   "frmStorageDoc.frx":0002
         Style           =   2  'Dropdown List
         TabIndex        =   0
         Top             =   200
         Width           =   4095
      End
      Begin VB.TextBox TxtstrDOCNo2 
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
         Left            =   4200
         MaxLength       =   20
         TabIndex        =   2
         Top             =   480
         Width           =   1575
      End
      Begin VB.TextBox TxtstrDOCNo 
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
         Left            =   1680
         MaxLength       =   20
         TabIndex        =   1
         Top             =   480
         Width           =   1455
      End
      Begin VB.ComboBox CboFileStatus1 
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
         ItemData        =   "frmStorageDoc.frx":0004
         Left            =   4200
         List            =   "frmStorageDoc.frx":0006
         TabIndex        =   4
         Text            =   "A - All"
         Top             =   720
         Width           =   1575
      End
      Begin VB.ComboBox CboCarton1 
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
         Left            =   1680
         TabIndex        =   3
         Top             =   720
         Width           =   1455
      End
      Begin VB.ComboBox CboUser1 
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
         Left            =   1680
         Sorted          =   -1  'True
         TabIndex        =   5
         Top             =   1005
         Width           =   4095
      End
      Begin MSMask.MaskEdBox MaskstrDORqst 
         Height          =   285
         Left            =   8520
         TabIndex        =   6
         Top             =   180
         Width           =   1695
         _ExtentX        =   2990
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
      Begin MSMask.MaskEdBox MaskstrDODelivery 
         Height          =   285
         Left            =   8520
         TabIndex        =   7
         Top             =   420
         Width           =   1695
         _ExtentX        =   2990
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
      Begin MSMask.MaskEdBox MaskstrDOReturn 
         Height          =   285
         Left            =   8520
         TabIndex        =   8
         Top             =   675
         Width           =   1695
         _ExtentX        =   2990
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
      Begin VB.Label Label31 
         Caption         =   "Document Type"
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
         Left            =   240
         TabIndex        =   49
         Top             =   240
         Width           =   1455
      End
      Begin VB.Label Label7 
         Caption         =   "Date Of Return"
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
         Left            =   6960
         TabIndex        =   48
         Top             =   720
         Width           =   1455
      End
      Begin VB.Label Label5 
         Caption         =   "Date Of Delivery"
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
         Left            =   6960
         TabIndex        =   47
         Top             =   435
         Width           =   1575
      End
      Begin VB.Label Label3 
         Caption         =   "Date Of Request"
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
         Left            =   6960
         TabIndex        =   46
         Top             =   180
         Width           =   1575
      End
      Begin VB.Label Label2 
         Caption         =   "Carton No"
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
         Left            =   240
         TabIndex        =   45
         Top             =   765
         Width           =   975
      End
      Begin VB.Label Label1 
         Caption         =   "Document No"
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
         Left            =   240
         TabIndex        =   44
         Top             =   495
         Width           =   1335
      End
      Begin VB.Label Label11 
         Caption         =   "File Status"
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
         Left            =   3240
         TabIndex        =   43
         Top             =   780
         Width           =   1215
      End
      Begin VB.Label Label4 
         Caption         =   "Request By"
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
         Left            =   240
         TabIndex        =   42
         Top             =   1080
         Width           =   975
      End
      Begin VB.Label Label15 
         Caption         =   "To"
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
         Left            =   3480
         TabIndex        =   41
         Top             =   525
         Width           =   375
      End
   End
   Begin VB.Frame FrameuPDATE 
      Enabled         =   0   'False
      Height          =   1455
      Left            =   120
      TabIndex        =   31
      Top             =   5760
      Width           =   10335
      Begin VB.ComboBox cboDOCTypeUpd 
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
         Left            =   1200
         TabIndex        =   9
         Top             =   200
         Width           =   3375
      End
      Begin VB.TextBox txtDocNo 
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
         MaxLength       =   20
         MultiLine       =   -1  'True
         TabIndex        =   10
         Top             =   480
         Width           =   3375
      End
      Begin VB.ComboBox CboCarton2 
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
         Left            =   1200
         TabIndex        =   11
         Top             =   720
         Width           =   1575
      End
      Begin VB.ComboBox CboFileStatus2 
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
         ItemData        =   "frmStorageDoc.frx":0008
         Left            =   3720
         List            =   "frmStorageDoc.frx":000A
         TabIndex        =   12
         Text            =   "I - In"
         Top             =   720
         Width           =   855
      End
      Begin VB.ComboBox CboUser2 
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
         Left            =   6720
         Sorted          =   -1  'True
         TabIndex        =   13
         Top             =   165
         Width           =   3495
      End
      Begin VB.ComboBox cboDept 
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
         Left            =   1200
         TabIndex        =   52
         Top             =   1010
         Width           =   3375
      End
      Begin MSMask.MaskEdBox txtDORqst 
         Height          =   285
         Left            =   6720
         TabIndex        =   14
         Top             =   450
         Width           =   1215
         _ExtentX        =   2143
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
      Begin MSMask.MaskEdBox txtDOD 
         Height          =   285
         Left            =   9000
         TabIndex        =   15
         Top             =   450
         Width           =   1215
         _ExtentX        =   2143
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
      Begin MSMask.MaskEdBox txtDOReturn 
         Height          =   285
         Left            =   6720
         TabIndex        =   16
         Top             =   705
         Width           =   1215
         _ExtentX        =   2143
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
      Begin VB.TextBox TxtRemarks 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   6720
         MaxLength       =   200
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   21
         Top             =   960
         Width           =   3495
      End
      Begin VB.Label Label10 
         Caption         =   "Dept"
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
         Left            =   240
         TabIndex        =   53
         Top             =   1000
         Width           =   975
      End
      Begin VB.Label Label9 
         Caption         =   "Doc No"
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
         Left            =   240
         TabIndex        =   51
         Top             =   480
         Width           =   735
      End
      Begin VB.Label Label6 
         Caption         =   "Doc Type"
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
         Left            =   240
         TabIndex        =   39
         Top             =   230
         Width           =   975
      End
      Begin VB.Label Label28 
         Caption         =   "Request By"
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
         TabIndex        =   38
         Top             =   240
         Width           =   975
      End
      Begin VB.Label Label27 
         Caption         =   "File Status"
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
         Left            =   2880
         TabIndex        =   37
         Top             =   765
         Width           =   855
      End
      Begin VB.Label Label26 
         Caption         =   "Carton No"
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
         Left            =   240
         TabIndex        =   36
         Top             =   750
         Width           =   975
      End
      Begin VB.Label Label25 
         Caption         =   "Date Of Req"
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
         TabIndex        =   35
         Top             =   495
         Width           =   1335
      End
      Begin VB.Label Label20 
         Caption         =   "Date Of Del"
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
         Left            =   8040
         TabIndex        =   34
         Top             =   480
         Width           =   1335
      End
      Begin VB.Label Label17 
         Caption         =   "Date Of Ret"
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
         TabIndex        =   33
         Top             =   735
         Width           =   1335
      End
      Begin VB.Label Label8 
         Caption         =   "Remarks"
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
         TabIndex        =   32
         Top             =   960
         Width           =   735
      End
   End
   Begin VB.Frame Frame5 
      Height          =   735
      Left            =   120
      TabIndex        =   20
      Top             =   7200
      Width           =   10395
      Begin VB.TextBox lblCurrent 
         Height          =   285
         Left            =   240
         TabIndex        =   54
         Top             =   240
         Width           =   615
      End
      Begin VB.CommandButton CmdStop 
         Caption         =   "St&op"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   350
         Left            =   6600
         TabIndex        =   26
         Top             =   240
         Width           =   735
      End
      Begin VB.CommandButton CmdClear 
         Caption         =   "&Clear"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   350
         Left            =   7320
         TabIndex        =   25
         Top             =   240
         Width           =   735
      End
      Begin VB.CommandButton CmdExit 
         Cancel          =   -1  'True
         Caption         =   "E&xit"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   350
         Left            =   9600
         TabIndex        =   24
         Top             =   240
         Width           =   735
      End
      Begin VB.CommandButton CmdSave 
         Caption         =   "Sa&ve"
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
         Height          =   350
         Left            =   8880
         TabIndex        =   19
         Top             =   240
         Width           =   735
      End
      Begin VB.CommandButton CmdSearch 
         Caption         =   "&Search"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   350
         Left            =   5880
         TabIndex        =   23
         Top             =   240
         Width           =   735
      End
      Begin VB.CommandButton CmdEdit 
         Caption         =   "&Edit"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   350
         Left            =   5160
         TabIndex        =   18
         Top             =   240
         Width           =   705
      End
      Begin VB.CommandButton CmdAdd 
         Caption         =   "&Add"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   350
         Left            =   4485
         TabIndex        =   17
         Top             =   240
         Width           =   675
      End
      Begin VB.CommandButton CmdPrintFile 
         Caption         =   "&Print File"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   350
         Left            =   8085
         TabIndex        =   22
         Top             =   240
         Width           =   800
      End
      Begin VB.Label Label30 
         Caption         =   "No. of Rec. checked"
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
         Left            =   2400
         TabIndex        =   29
         Top             =   285
         Width           =   1455
      End
      Begin VB.Label LblRecChk 
         Alignment       =   2  'Center
         BackColor       =   &H80000018&
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
         Left            =   3960
         TabIndex        =   28
         Top             =   285
         Width           =   405
      End
      Begin VB.Label Label14 
         Caption         =   "Records"
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
         Left            =   1200
         TabIndex        =   27
         Top             =   285
         Width           =   735
      End
   End
   Begin MSComctlLib.ListView Lststr 
      Height          =   3975
      Left            =   120
      TabIndex        =   50
      Top             =   1800
      Width           =   10335
      _ExtentX        =   18230
      _ExtentY        =   7011
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      Checkboxes      =   -1  'True
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
      NumItems        =   11
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Doc. No"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "CartonNo"
         Object.Width           =   2117
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   2
         Text            =   "Carton Status"
         Object.Width           =   2117
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Text            =   "Doc. Status"
         Object.Width           =   1764
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   4
         Text            =   "Date Register"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   5
         Text            =   "RequestBy"
         Object.Width           =   2293
      EndProperty
      BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   6
         Text            =   "Date Request"
         Object.Width           =   2118
      EndProperty
      BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   7
         Text            =   "Date Delivery"
         Object.Width           =   2118
      EndProperty
      BeginProperty ColumnHeader(9) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   8
         Text            =   "Date Return"
         Object.Width           =   2117
      EndProperty
      BeginProperty ColumnHeader(10) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   9
         Text            =   "remarks"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(11) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   10
         Text            =   "Dept"
         Object.Width           =   2540
      EndProperty
   End
   Begin VB.Label Label12 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00404040&
      Caption         =   "Storage / Cartons"
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
      TabIndex        =   30
      Top             =   0
      Width           =   11805
   End
End
Attribute VB_Name = "frmStorageDoc"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public intExit As Integer
Dim bExit As Boolean
Dim EditFlag As Boolean
Dim listloop1 As Integer
Dim Check As Boolean
Dim CheckCnt As Integer
Dim CaseNoArray() As String
Dim strAppend As String


Private Sub CboCarton1_KeyPress(KeyAscii As Integer)
Dim NewText As String
Dim ValidCount As Integer
Dim ValidValue As String
Dim i As Integer
    If KeyAscii >= 32 And KeyAscii <> 127 Then
    NewText = LCase(Left(CboCarton1.Text, CboCarton1.SelStart) + Chr(KeyAscii) + Mid(CboCarton1.Text, CboCarton1.SelStart + CboCarton1.SelLength + 1))
 
    ValidCount = 0
    ValidValue = ""
 
    For i = 0 To CboCarton1.ListCount - 1
 
        If NewText = LCase(Left(CboCarton1.List(i), Len(NewText))) Then
                ValidCount = ValidCount + 1
                ValidValue = CboCarton1.List(i)
        End If
        Next
        If ValidCount <= 1 Then KeyAscii = 0
            If ValidCount = 1 Then
                CboCarton1.Text = ValidValue
                CboCarton1.SelStart = 0
                CboCarton1.SelLength = Len(ValidValue)
            End If
        End If
End Sub

Private Sub CboCarton2_KeyPress(KeyAscii As Integer)
Dim NewText As String
Dim ValidCount As Integer
Dim ValidValue As String
Dim i As Integer
    If KeyAscii >= 32 And KeyAscii <> 127 Then
    NewText = LCase(Left(CboCarton2.Text, CboCarton2.SelStart) + Chr(KeyAscii) + Mid(CboCarton2.Text, CboCarton2.SelStart + CboCarton2.SelLength + 1))
 
    ValidCount = 0
    ValidValue = ""
 
    For i = 0 To CboDept.ListCount - 1
 
        If NewText = LCase(Left(CboCarton2.List(i), Len(NewText))) Then
                ValidCount = ValidCount + 1
                ValidValue = CboCarton2.List(i)
        End If
        Next
        If ValidCount <= 1 Then KeyAscii = 0
            If ValidCount = 1 Then
                CboCarton2.Text = ValidValue
                CboCarton2.SelStart = 0
                CboCarton2.SelLength = Len(ValidValue)
            End If
        End If
End Sub

Private Sub cboDept_KeyPress(KeyAscii As Integer)
Dim NewText As String
Dim ValidCount As Integer
Dim ValidValue As String
Dim i As Integer
    If KeyAscii >= 32 And KeyAscii <> 127 Then
    NewText = LCase(Left(CboDept.Text, CboDept.SelStart) + Chr(KeyAscii) + Mid(CboDept.Text, CboDept.SelStart + CboDept.SelLength + 1))
 
    ValidCount = 0
    ValidValue = ""
 
    For i = 0 To CboDept.ListCount - 1
 
        If NewText = LCase(Left(CboDept.List(i), Len(NewText))) Then
                ValidCount = ValidCount + 1
                ValidValue = CboDept.List(i)
        End If
        Next
        If ValidCount <= 1 Then KeyAscii = 0
            If ValidCount = 1 Then
                CboDept.Text = ValidValue
                CboDept.SelStart = 0
                CboDept.SelLength = Len(ValidValue)
            End If
        End If
End Sub

Private Sub cboDOCType_Change()
    CmdSave.Enabled = False
End Sub

Private Sub cboDOCType_Click()
    CmdSave.Enabled = False
End Sub

Private Sub cboDOCTypeUpd_KeyPress(KeyAscii As Integer)
Dim NewText As String
Dim ValidCount As Integer
Dim ValidValue As String
Dim i As Integer
    If KeyAscii >= 32 And KeyAscii <> 127 Then
    NewText = LCase(Left(cboDOCTypeUpd.Text, cboDOCTypeUpd.SelStart) + Chr(KeyAscii) + Mid(cboDOCTypeUpd.Text, cboDOCTypeUpd.SelStart + cboDOCTypeUpd.SelLength + 1))
 
    ValidCount = 0
    ValidValue = ""
 
    For i = 0 To cboDOCTypeUpd.ListCount - 1
 
        If NewText = LCase(Left(cboDOCTypeUpd.List(i), Len(NewText))) Then
                ValidCount = ValidCount + 1
                ValidValue = cboDOCTypeUpd.List(i)
        End If
        Next
        If ValidCount <= 1 Then KeyAscii = 0
            If ValidCount = 1 Then
                cboDOCTypeUpd.Text = ValidValue
                cboDOCTypeUpd.SelStart = 0
                cboDOCTypeUpd.SelLength = Len(ValidValue)
            End If
        End If

End Sub

Private Sub CboFileStatus1_KeyPress(KeyAscii As Integer)
Dim NewText As String
Dim ValidCount As Integer
Dim ValidValue As String
Dim i As Integer
    If KeyAscii >= 32 And KeyAscii <> 127 Then
    NewText = LCase(Left(CboFileStatus1.Text, CboFileStatus1.SelStart) + Chr(KeyAscii) + Mid(CboFileStatus1.Text, CboFileStatus1.SelStart + CboFileStatus1.SelLength + 1))
 
    ValidCount = 0
    ValidValue = ""
 
    For i = 0 To CboFileStatus1.ListCount - 1
 
        If NewText = LCase(Left(CboFileStatus1.List(i), Len(NewText))) Then
                ValidCount = ValidCount + 1
                ValidValue = CboFileStatus1.List(i)
        End If
        Next
        If ValidCount <= 1 Then KeyAscii = 0
            If ValidCount = 1 Then
                CboFileStatus1.Text = ValidValue
                CboFileStatus1.SelStart = 0
                CboFileStatus1.SelLength = Len(ValidValue)
            End If
        End If
End Sub

Private Sub CboFileStatus2_KeyPress(KeyAscii As Integer)
Dim NewText As String
Dim ValidCount As Integer
Dim ValidValue As String
Dim i As Integer
    If KeyAscii >= 32 And KeyAscii <> 127 Then
    NewText = LCase(Left(CboFileStatus2.Text, CboFileStatus2.SelStart) + Chr(KeyAscii) + Mid(CboFileStatus2.Text, CboFileStatus2.SelStart + CboFileStatus2.SelLength + 1))
 
    ValidCount = 0
    ValidValue = ""
 
    For i = 0 To CboFileStatus2.ListCount - 1
 
        If NewText = LCase(Left(CboFileStatus2.List(i), Len(NewText))) Then
                ValidCount = ValidCount + 1
                ValidValue = CboFileStatus2.List(i)
        End If
        Next
        If ValidCount <= 1 Then KeyAscii = 0
            If ValidCount = 1 Then
                CboFileStatus2.Text = ValidValue
                CboFileStatus2.SelStart = 0
                CboFileStatus2.SelLength = Len(ValidValue)
            End If
        End If

End Sub

Private Sub CboUser1_KeyPress(KeyAscii As Integer)
Dim NewText As String
Dim ValidCount As Integer
Dim ValidValue As String
Dim i As Integer
    If KeyAscii >= 32 And KeyAscii <> 127 Then
    NewText = LCase(Left(CboUser1.Text, CboUser1.SelStart) + Chr(KeyAscii) + Mid(CboUser1.Text, CboUser1.SelStart + CboUser1.SelLength + 1))
 
    ValidCount = 0
    ValidValue = ""
 
    For i = 0 To CboUser1.ListCount - 1
 
        If NewText = LCase(Left(CboUser1.List(i), Len(NewText))) Then
                ValidCount = ValidCount + 1
                ValidValue = CboUser1.List(i)
        End If
        Next
        If ValidCount <= 1 Then KeyAscii = 0
            If ValidCount = 1 Then
                CboUser1.Text = ValidValue
                CboUser1.SelStart = 0
                CboUser1.SelLength = Len(ValidValue)
            End If
        End If

End Sub

Private Sub CboUser2_KeyPress(KeyAscii As Integer)
Dim NewText As String
Dim ValidCount As Integer
Dim ValidValue As String
Dim i As Integer
    If KeyAscii >= 32 And KeyAscii <> 127 Then
    NewText = LCase(Left(CboUser2.Text, CboUser2.SelStart) + Chr(KeyAscii) + Mid(CboUser2.Text, CboUser2.SelStart + CboUser2.SelLength + 1))
 
    ValidCount = 0
    ValidValue = ""
 
    For i = 0 To CboUser2.ListCount - 1
 
        If NewText = LCase(Left(CboUser2.List(i), Len(NewText))) Then
                ValidCount = ValidCount + 1
                ValidValue = CboUser2.List(i)
        End If
        Next
        If ValidCount <= 1 Then KeyAscii = 0
            If ValidCount = 1 Then
                CboUser2.Text = ValidValue
                CboUser2.SelStart = 0
                CboUser2.SelLength = Len(ValidValue)
            End If
        End If
End Sub

Private Sub cmdAdd_Click()
    EditFlag = False
    CmdEdit.Enabled = False
        If CheckCnt > 0 Then
            cboDOCTypeUpd = ""
            CboCarton2 = ""
            CboCarton2.Text = ""
            CboDept = ""
            CboDept.Text = ""
            CboFileStatus2.Text = ""
            CboUser2.Text = ""
            txtDORqst = "__/__/____"
            txtDOD = "__/__/____"
            txtDOReturn = "__/__/____"
            TxtRemarks = ""
            FrameuPDATE.Enabled = True
            'cboDOCTypeUpd.SetFocus
            CmdSave.Enabled = True
            CmdSave.Caption = "&Save"
            cboDOCTypeUpd = cboDocType
            Call EntriesEnable(True)
        Else
            MsgBox "Please checked at least 1 Claim ! ", vbOKOnly + vbCritical
        End If
    

End Sub

Private Sub CmdClear_Click()
    Call ClearForm(Me)
    Check = False
    EditFlag = False
    Call EntriesEnable(True)
    Lststr.ListItems.Clear
    FrameuPDATE.Enabled = False
    ReDim CaseNoArray(0)
End Sub

Private Sub cmdEdit_Click()
    EditFlag = True
    CmdSave.Enabled = True
    CmdSave.Caption = "&Update"
    FrameuPDATE.Enabled = True
    Call EntriesEnable(True)
    cboDOCTypeUpd.SetFocus
End Sub

Private Sub cmdExit_Click()
    Unload Me
End Sub

Private Sub CmdPrintFile_Click()
    Screen.MousePointer = vbHourglass
        If Lststr.ListItems.Count <> 0 Then
            Call ExportListViewtoExcel(Lststr)
        Else
            MsgBox "Report cannot be printed without any record.", vbInformation
        End If
    Screen.MousePointer = vbDefault
End Sub

Private Sub cmdSave_Click()
Dim AA, CASVer As Integer
Dim RecordSaved
Dim DOCNo As String

    Screen.MousePointer = vbHourglass
    CmdSave.Enabled = False
   ' medixmasterconnWS.Open "File Name=" & BOSPath & "Conn\mediconnectWS.UDL"

        If Trim(cboDOCTypeUpd) = "" Then
            MsgBox "Please Enter Document Type", vbOKOnly + vbCritical, DialogBoxTitle
            'cboDOCTypeUpd.SetFocus
        ElseIf Trim(CboCarton2) = "" Then
            MsgBox "Please Enter Carton Number", vbOKOnly + vbCritical, DialogBoxTitle
            CboCarton2.SetFocus
        ElseIf Trim(Mid(CboFileStatus2, 1, 1)) = "" Then
            MsgBox "Please Enter File Status", vbOKOnly + vbCritical, DialogBoxTitle
            CboFileStatus1.SetFocus
        ElseIf CboUser2 <> "" And txtDORqst = "__/__/____" Then
            MsgBox "Please Enter Date Of Request", vbOKOnly + vbCritical, DialogBoxTitle
            txtDORqst.SetFocus
        ElseIf CboUser2 <> "" And Format(txtDOD, "YYYYMMDD") < Format(txtDORqst, "YYYYMMDD") Then
            MsgBox "Invalid Date ! Date of Delivery must greater than Date of Request ", vbOKOnly + vbCritical, DialogBoxTitle
            txtDOD = "__/__/____"
            txtDOD.SetFocus
        ElseIf CboUser2 <> "" And Format(txtDOReturn, "YYYYMMDD") < Format(txtDOD, "YYYYMMDD") Then
            MsgBox "Invalid Date ! Date of Return must greater than Date of Delivery ", vbOKOnly + vbCritical, DialogBoxTitle
            txtDOReturn = "__/__/____"
            txtDOReturn.SetFocus
        Else
            RecordSaved = MsgBox("Do you want to save the record?", vbYesNo + vbExclamation)
            If RecordSaved = vbYes Then
                If EditFlag = True Then
                   Erase CaseNoArray
                   ReDim Preserve CaseNoArray(1)
                   CaseNoArray(1) = Lststr.SelectedItem.Text
                   
                End If
                For AA = 0 To UBound(CaseNoArray)
                    If Trim(CaseNoArray(AA)) <> "" Then
                        delimiter1 = InStr(1, CaseNoArray(AA))
                        DOCNo = Trim(Mid(CaseNoArray(AA), 1, IIf(delimiter1 = 0, 10, delimiter1 - 1)))
                    
                        Call SaveStore(DOCNo)
                    End If
                Next AA
                Erase CaseNoArray
                Arraycnt = 0
                ReDim CaseNoArray(0)
                If EditFlag = False Then
                    MsgBox "Record Saved...", vbOKOnly + vbInformation, DialogBoxTitle
                Else
                    MsgBox "Record Updated...", vbOKOnly + vbInformation, DialogBoxTitle
                End If
                'Call ClearStoreInfo
                'If ChkScan.value = 1 Then
                '    TxtstrCasNo = Lststr.SelectedItem.Text
                'End If
                Lststr.ListItems.Clear
                Call ClearStoreInfo
                Call StorageListing(strAppend)
            End If
        End If
  '  medixmasterconnWS.Close
 '   Set medixmasterconnWS = Nothing
    Call EntriesEnable(False)
    Screen.MousePointer = vbDefault

End Sub

Private Sub cmdSearch_Click()
    CmdAdd.Enabled = True
    CmdSave.Enabled = False
    CmdSearch.Enabled = False
    Screen.MousePointer = vbHourglass
  '  medixmasterconnWS.Open "File Name=" & BOSPath & "Conn\mediconnectWS.UDL"
        
        Erase CaseNoArray
        ReDim CaseNoArray(0)
        Call StorageSearchRecord
        Screen.MousePointer = Default
        
 '   medixmasterconnWS.Close
 '   Set medixmasterconnWS = Nothing
    CmdSearch.Enabled = True
    Screen.MousePointer = Default
End Sub

Private Sub CmdStop_Click()
    intExit = 1
End Sub

Private Sub Form_Load()
    intExit = 0
    Check = False
    CheckCnt = 0
    EditFlag = False
    
  '  medixmasterconnWS.Open "File Name=" & BOSPath & "Conn\mediconnectWS.UDL"
  '  medixmasterconnMaint.Open "File Name=" & BOSPath & "Conn\MediConnHISMaint.UDL"
    Call ComboListing
  '  medixmasterconnWS.Close
  '  Set medixmasterconnWS = Nothing
  '  medixmasterconnMaint.Close
   ' Set medixmasterconnMaint = Nothing
'
End Sub

Private Sub ComboListing()
Dim USRRec As New ADODB.Recordset
Dim CRTRec As New ADODB.Recordset
Dim DeptRec As New ADODB.Recordset
Dim CMD As New ADODB.Command
    
    CboFileStatus1.Clear
    CboFileStatus1.AddItem "A - All"
    CboFileStatus1.AddItem "I - In"
    CboFileStatus1.AddItem "O - Out"
    CboFileStatus1.ListIndex = 0
    
    CboFileStatus2.Clear
    CboFileStatus2.AddItem "I - In"
    CboFileStatus2.AddItem "O - Out"
        
    cboDocType.Clear
    cboDocType.AddItem "B - Bordx"
    cboDocType.AddItem "P - Payment Voucher"
    cboDocType.AddItem "R - Receipt"
    cboDocType.Text = "B - Bordx"
    
    cboDOCTypeUpd.Clear
    cboDOCTypeUpd.AddItem "B - Bordx"
    cboDOCTypeUpd.AddItem "P - Payment Voucher"
    cboDOCTypeUpd.AddItem "R - Receipt"

    Set CMD.ActiveConnection = medixmasterconnWS
    CMD.CommandText = "SearchCarton"
    CMD.CommandType = adCmdStoredProc
    CMD.CommandTimeout = 15
    CRTRec.CursorLocation = adUseClient
    CRTRec.CursorType = adOpenForwardOnly
    CRTRec.CacheSize = 100
    Set CRTRec = CMD.Execute
    
    CboCarton1.Clear
    CboCarton2.Clear
    Do Until CRTRec.EOF
        With CRTRec
            CboCarton1.AddItem !CRTCartonNo
            CboCarton2.AddItem !CRTCartonNo
            CRTRec.MoveNext
        End With
    Loop
    
    CRTRec.Close
    Set CRTRec = Nothing
    
    Set CMD.ActiveConnection = medixmasterconnMaint
    CMD.CommandText = "SearchUser"
    CMD.CommandType = adCmdStoredProc
    CMD.CommandTimeout = 15
    USRRec.CursorLocation = adUseClient
    USRRec.CursorType = adOpenForwardOnly
    USRRec.CacheSize = 100
    Set USRRec = CMD.Execute
    CboUser1.Clear
    CboUser2.Clear
    Do Until USRRec.EOF
        With USRRec
            CboUser1.AddItem IIf(Len(!USRName), !USRName, "")
            CboUser2.AddItem IIf(Len(!USRName), !USRName, "")
            USRRec.MoveNext
        End With
    Loop
    USRRec.Close
    Set USRRec = Nothing
    
    Set CMD.ActiveConnection = medixmasterconnMaint
    CMD.CommandText = "SearchDPT"
    CMD.CommandType = adCmdStoredProc
    CMD.CommandTimeout = 15
    DeptRec.CursorLocation = adUseClient
    DeptRec.CursorType = adOpenForwardOnly
    DeptRec.CacheSize = 100
    Set DeptRec = CMD.Execute
    CboDept.Clear
    Do Until DeptRec.EOF
    With DeptRec
        CboDept.AddItem !DPTDescription
        DeptRec.MoveNext
    End With
    Loop
    DeptRec.Close
    Set DeptRec = Nothing

End Sub

Private Function StorageSearchRecord()
Dim strsql As String
Dim sql As String
Dim storage As New ADODB.Recordset
Dim TempSearchDocNoFr As String
Dim TempSearchDocNoTo As String

    Call EntriesEnable(False)
    
    strAppend = ""
    If Mid(cboDocType, 1, 1) = "B" Then
        If Len(Trim(TxtstrDOCNo)) Then
            strAppend = strAppend + " AND BordxRefNo >= '" & RemStr(Trim(TxtstrDOCNo)) & "'"
        End If
        If Len(Trim(TxtstrDOCNo2)) Then
            strAppend = strAppend + " AND BordxRefNo <= '" & RemStr(Trim(TxtstrDOCNo2)) & "'"
        End If
    ElseIf Mid(cboDocType, 1, 1) = "P" Then
        If Len(Trim(TxtstrDOCNo)) Then
            strAppend = strAppend + " AND SearchPVNo >= '" & RemStr(Trim(TxtstrDOCNo)) & "'"
        End If
        If Len(Trim(TxtstrDOCNo2)) Then
            strAppend = strAppend + " And SearchPVNo <= '" & RemStr(Trim(TxtstrDOCNo2)) & "'"
        End If
    ElseIf Mid(cboDocType, 1, 1) = "R" Then
        If Len(Trim(TxtstrDOCNo)) Then
            strAppend = strAppend + " AND RecNo >= '" & RemStr(Trim(TxtstrDOCNo)) & "'"
        End If
        If Len(Trim(TxtstrDOCNo2)) Then
            strAppend = strAppend + " AND RecNo <= '" & RemStr(Trim(TxtstrDOCNo2)) & "'"
        End If
    End If
    If Len(Trim(CboCarton1)) Then
        strAppend = strAppend + " AND STRCartonNo = '" & RemStr(Trim(CboCarton1)) & "'"
    End If
    
    If Trim(CboFileStatus1) <> "" And Mid(CboFileStatus1, 1, 1) <> "A" Then
        strAppend = strAppend + " AND STRFileStatus = '" & RemStr(Trim(Mid(CboFileStatus1, 1, 1))) & "'"
    End If
    
    If Len(Trim(CboUser1)) Then
        strAppend = strAppend + " AND STRRequestBy = '" & RemStr(Trim(CboUser1)) & "'"
    End If
    
    If IsDate(MaskstrDORqst) = True Then
        strAppend = strAppend + " And DORequest = '" & Format(MaskstrDORqst, "mm/dd/yyyy") & "'"
    End If
    
    If IsDate(MaskstrDODelivery) = True Then
        strAppend = strAppend + " And DODelivery = '" & Format(MaskstrDODelivery, "mm/dd/yyyy") & "'"
    End If
    
    If IsDate(MaskstrDOReturn) = True Then
        strAppend = strAppend + " And DOReturn = '" & Format(MaskstrDOReturn, "mm/dd/yyyy") & "'"
    End If

    Call StorageListing(strAppend)
    
End Function

Private Sub StorageListing(Optional strAppend As String)
Dim rst2 As New ADODB.Recordset
Dim rstCount As New ADODB.Recordset
Dim Record As ListItem
Dim sql As String
Dim sqlCount As String
Dim total As String
Dim Current As String
Dim strAppendCase As Integer
Dim CMD As New ADODB.Command
Dim Prm1 As ADODB.Parameter
Dim Prm2 As ADODB.Parameter
Dim TempPVNo As String
Dim TempRecNo As String
    total = 0
    Current = 0
    TempPVNo = ""
    TempRecNo = ""
    Lststr.ListItems.Clear
    If Mid(cboDocType, 1, 1) = "B" Then
        strAppendCase = "1"
    ElseIf Mid(cboDocType, 1, 1) = "P" Then
        strAppendCase = "2"
    ElseIf Mid(cboDocType, 1, 1) = "R" Then
        strAppendCase = "3"
    End If
    Set CMD.ActiveConnection = medixmasterconnWS
    CMD.CommandText = "Storage_Documents"
    CMD.CommandType = adCmdStoredProc
    CMD.CommandTimeout = 300
    Set Prm1 = CMD.CreateParameter("StrAppend", adVarChar, adParamInput, 2000, strAppend)
    CMD.Parameters.Append Prm1
    Set Prm2 = CMD.CreateParameter("strAppendCase", adVarChar, adParamInput, 255, strAppendCase)
    CMD.Parameters.Append Prm2
    rst2.CursorLocation = adUseClient
    rst2.CursorType = adOpenForwardOnly
    rst2.CacheSize = 100
    Set rst2 = CMD.Execute
                
    If rst2.EOF = True Then
        MsgBox "No record found!", vbOKOnly + vbCritical, DialogBoxTitle
        CmdAdd.Enabled = True
    Else
    
    Do
        Do Until rst2.EOF
            With rst2
                If strAppendCase = "1" Then
                    Set Record = Lststr.ListItems.Add(, , IIf(Len(!BordxRefNo), !BordxRefNo, ""))
                ElseIf strAppendCase = "2" Then
                    Set Record = Lststr.ListItems.Add(, , IIf(Len(!SearchPVNo), !SearchPVNo, ""))
                ElseIf strAppendCase = "3" Then
                    Set Record = Lststr.ListItems.Add(, , IIf(Len(!RecNo), !RecNo, ""))
                    
                End If
                
                If Len(!STRCartonNo) Then
                    Record.SubItems(1) = !STRCartonNo
                End If
                If Len(!STRDocType) Then
                    Record.SubItems(2) = !STRDocType
                End If
                If Len(!STRFileStatus) Then
                    Record.SubItems(3) = !STRFileStatus
                End If
                If Len(!LastUpdateDate) Then
                    Record.SubItems(4) = Format(!LastUpdateDate, "dd/mm/yyyy")
                End If
                If Len(!STRRequestBy) Then
                    Record.SubItems(5) = !STRRequestBy
                End If
                If Len(!DORequest) Then
                    Record.SubItems(6) = Format(!DORequest, "dd/mm/yyyy")
                End If
                If Len(!DODelivery) Then
                    Record.SubItems(7) = Format(!DODelivery, "dd/mm/yyyy")
                End If
                If Len(!DOReturn) Then
                    Record.SubItems(8) = Format(!DOReturn, "dd/mm/yyyy")
                End If
                If Len(!STRRemarks) Then
                    Record.SubItems(9) = !STRRemarks
                End If
                If Len(!STRDept) Then
                    Record.SubItems(10) = !STRDept
                End If
                Current = Current + 1
                lblCurrent = Current
            End With
            rst2.MoveNext
            DoEvents
                If intExit = 1 Then
                   Check = False
                   Exit Do
                End If
        Loop
    Loop Until Check = False
End If
intExit = 0
rst2.Close
Set rst2 = Nothing
CmdExit.Enabled = True
End Sub


Private Sub Lststr_Click()
    
    Call ClearStoreInfo
    If Lststr.ListItems.Count <> 0 Then
        If Len(Trim(Lststr.SelectedItem.SubItems(1))) Then
            txtDocNo = Lststr.SelectedItem.Text
            If Len(Trim(Lststr.SelectedItem.SubItems(1))) Then
                CboCarton2 = Lststr.SelectedItem.SubItems(1)
            End If
            If Len(Trim(Lststr.SelectedItem.SubItems(2))) Then
                cboDOCTypeUpd = Lststr.SelectedItem.SubItems(2)
            End If
            If Len(Trim(Lststr.SelectedItem.SubItems(3))) Then
                CboFileStatus2 = Lststr.SelectedItem.SubItems(3)
            End If
            'If Len(Trim(Lststr.SelectedItem.SubItems(4))) Then
            '    CboFileStatus2 = Lststr.SelectedItem.SubItems(4)
            'End If
            If Len(Trim(Lststr.SelectedItem.SubItems(5))) Then
                CboUser2 = Lststr.SelectedItem.SubItems(5)
            End If
            If Len(Trim(Lststr.SelectedItem.SubItems(6))) Then
                txtDORqst = Lststr.SelectedItem.SubItems(6)
            End If
            If Len(Trim(Lststr.SelectedItem.SubItems(7))) Then
                txtDOD = Lststr.SelectedItem.SubItems(7)
            End If
            If Len(Trim(Lststr.SelectedItem.SubItems(8))) Then
                txtDOReturn = Lststr.SelectedItem.SubItems(8)
            End If
            If Len(Trim(Lststr.SelectedItem.SubItems(9))) Then
                TxtRemarks = Lststr.SelectedItem.SubItems(9)
            End If
            If Len(Trim(Lststr.SelectedItem.SubItems(10))) Then
                CboDept = Lststr.SelectedItem.SubItems(10)
            End If
            CmdEdit.Enabled = True
        Else
            CmdEdit.Enabled = False
        End If
    End If
End Sub

Private Sub Lststr_ItemCheck(ByVal Item As MSComctlLib.ListItem)
    CmdSave.Enabled = True
    FrameuPDATE.Enabled = True
    'medixmasterconnWS.Open "File Name=" & BOSPath & "Conn\mediconnectWS.UDL"
        Call ProcessItemChecked(Item)
    'medixmasterconnWS.Close
    'Set medixmasterconnWS = Nothing
End Sub

Private Function ProcessItemChecked(Lststr As ListItem, Optional Partial As Boolean)
Dim BB, newbound  As Integer
Dim tempArray() As String
Dim strsql As String
Dim StoRec As New ADODB.Recordset
Dim CaseNo As String
Dim tmpSelCasNo As String

Static Arraycnt As Integer
    
    CaseNo = Lststr.Text
    strsql = "Select STRDocumentNo From StorageDocument where STRDocumentNo = '" & Lststr.Text & "'"
    StoRec.Open strsql, medixmasterconnWS, adOpenKeyset, adLockPessimistic

    If Lststr.Checked = True Then
        If StoRec.recordCount = 0 Then
            ReDim Preserve CaseNoArray(Arraycnt)
            CaseNoArray(Arraycnt) = Lststr.Text
            Arraycnt = Arraycnt + 1
            CheckCnt = CheckCnt + 1
            LblRecChk = CheckCnt
        Else
            MsgBox "Claim has been stored!", vbCritical
            Lststr.Checked = False
        End If
    Else
        Arraycnt = Arraycnt - 1
        CheckCnt = CheckCnt - 1
        LblRecChk = IIf(CheckCnt < 0, 0, CheckCnt)

        If UBound(CaseNoArray) <> 0 Then
            For BB = 0 To UBound(CaseNoArray)
                tmpSelCasNo = Lststr.Text
                If tmpSelCasNo <> CaseNoArray(BB) Then
                    ReDim Preserve tempArray(newbound)
                    tempArray(newbound) = CaseNoArray(BB)
                    newbound = newbound + 1
                End If
            Next BB
            
            Erase CaseNoArray
            ReDim CaseNoArray(0)
            
            If UBound(tempArray) = 0 Then
                Arraycnt = 1
            Else
                Arraycnt = UBound(tempArray) + 1
            End If
            ReDim CaseNoArray(UBound(tempArray))
            For BB = 0 To UBound(tempArray)
                CaseNoArray(BB) = tempArray(BB)
            Next BB
            Erase tempArray
            ReDim tempArray(0)
        End If
    End If
    StoRec.Close
    Set StoRec = Nothing
End Function

Private Sub SaveStore(DOCNo As String)
Dim RecordSaved
Dim SaveSTORAGE As New ADODB.Recordset
Dim strsql As String
Dim sql2 As String
Dim Store As New ADODB.Recordset
Dim CLM As New ADODB.Recordset
Dim CARTON As New ADODB.Recordset
    
'strsql = ""
    strsql = "Select STRDocumentNo,STRCartonNo,STRFileStatus, " & _
                "STRRequestBy, STRDocType,DORequest,DODelivery,DOReturn, " & _
                "DORequest,STRRemarks,LastUpdateBy,LastUpdateDate,STRDept " & _
                "From StorageDocument Where STRDocumentNo = '" & Trim(DOCNo) & "'"
    
    SaveSTORAGE.Open strsql, medixmasterconnWS, adOpenKeyset, adLockOptimistic
    
    If SaveSTORAGE.recordCount > 0 And EditFlag = False Then
        MsgBox "Record Duplicated! Please Change The Corresponding Field", vbCritical, "Error"
        SaveSTORAGE.Close
        CboCarton2.SetFocus
        CboCarton2.SelStart = 0
        CboCarton2.SelLength = Len(CboCarton2)
        Set SaveSTORAGE = Nothing
        Exit Sub
    Else
        If EditFlag = False Then
            SaveSTORAGE.AddNew
        End If
        With SaveSTORAGE
            !STRDocumentNo = DOCNo
            !STRCartonNo = ChkStr(CboCarton2)
            !STRFileStatus = Trim(Mid(CboFileStatus2, 1, 1))
            !STRRequestBy = ChkStr(CboUser2)
            !STRDocType = Mid(cboDocType, 1, 1)
            !STRDept = CboDept
            If CboUser2 <> "" Then
                !DORequest = txtDORqst
                If txtDOD <> "__/__/____" Then
                    !DODelivery = txtDOD
                Else
                    !DODelivery = Null
                End If
                If txtDOReturn <> "__/__/____" Then
                    !DOReturn = txtDOReturn
                Else
                    !DOReturn = Null
                End If
            Else
                !DODelivery = Null
                !DOReturn = Null
                !DORequest = Null
            End If
            !STRRemarks = ChkStr(TxtRemarks)
            !LastUpdateBy = Logon
            !LastUpdateDate = Date
        End With
        SaveSTORAGE.Update
        
        SaveSTORAGE.Close
        Set SaveSTORAGE = Nothing
    End If
End Sub

Private Sub ClearStoreInfo()
    cboDOCTypeUpd = ""
    txtDocNo = ""
    CboCarton1.Text = ""
    CboFileStatus1.Text = ""
    CboUser1.Text = ""
    MaskstrDORqst = "__/__/____"
    MaskstrDODelivery = "__/__/____"
    MaskstrDOReturn = "__/__/____"
    CboCarton2.Text = ""
    CboDept = ""
    CboDept.Text = ""
    CboFileStatus2.Text = ""
    CboUser2.Text = ""
    txtDORqst = "__/__/____"
    txtDOD = "__/__/____"
    txtDOReturn = "__/__/____"
    TxtRemarks = ""
End Sub

Private Sub EntriesEnable(status As Boolean)
    '----- STORAGE
    cboDOCTypeUpd.Enabled = status
    CboCarton2.Enabled = status
    CboFileStatus2.Enabled = status
    CboUser2.Enabled = status
    txtDORqst.Enabled = status
    txtDOD.Enabled = status
    txtDOReturn.Enabled = status
    TxtRemarks.Enabled = status
    CboDept.Enabled = status
End Sub

Private Sub Lststr_KeyDown(KeyCode As Integer, Shift As Integer)
    Call ClearStoreInfo
    If Lststr.ListItems.Count <> 0 Then
        If Len(Trim(Lststr.SelectedItem.SubItems(1))) Then
            txtDocNo = Lststr.SelectedItem.Text
            If Len(Trim(Lststr.SelectedItem.SubItems(1))) Then
                CboCarton2 = Lststr.SelectedItem.SubItems(1)
            End If
            If Len(Trim(Lststr.SelectedItem.SubItems(2))) Then
                cboDOCTypeUpd = Lststr.SelectedItem.SubItems(2)
            End If
            If Len(Trim(Lststr.SelectedItem.SubItems(3))) Then
                CboFileStatus2 = Lststr.SelectedItem.SubItems(3)
            End If
            'If Len(Trim(Lststr.SelectedItem.SubItems(4))) Then
            '    CboFileStatus2 = Lststr.SelectedItem.SubItems(4)
            'End If
            If Len(Trim(Lststr.SelectedItem.SubItems(5))) Then
                CboUser2 = Lststr.SelectedItem.SubItems(5)
            End If
            If Len(Trim(Lststr.SelectedItem.SubItems(6))) Then
                txtDORqst = Lststr.SelectedItem.SubItems(6)
            End If
            If Len(Trim(Lststr.SelectedItem.SubItems(7))) Then
                txtDOD = Lststr.SelectedItem.SubItems(7)
            End If
            If Len(Trim(Lststr.SelectedItem.SubItems(8))) Then
                txtDOReturn = Lststr.SelectedItem.SubItems(8)
            End If
            If Len(Trim(Lststr.SelectedItem.SubItems(9))) Then
                TxtRemarks = Lststr.SelectedItem.SubItems(9)
            End If
            If Len(Trim(Lststr.SelectedItem.SubItems(10))) Then
                CboDept = Lststr.SelectedItem.SubItems(10)
            End If
        End If
    End If

End Sub

Private Sub Lststr_KeyUp(KeyCode As Integer, Shift As Integer)
    Call ClearStoreInfo
    If Lststr.ListItems.Count <> 0 Then
        If Len(Trim(Lststr.SelectedItem.SubItems(1))) Then
            txtDocNo = Lststr.SelectedItem.Text
            If Len(Trim(Lststr.SelectedItem.SubItems(1))) Then
                CboCarton2 = Lststr.SelectedItem.SubItems(1)
            End If
            If Len(Trim(Lststr.SelectedItem.SubItems(2))) Then
                cboDOCTypeUpd = Lststr.SelectedItem.SubItems(2)
            End If
            If Len(Trim(Lststr.SelectedItem.SubItems(3))) Then
                CboFileStatus2 = Lststr.SelectedItem.SubItems(3)
            End If
            'If Len(Trim(Lststr.SelectedItem.SubItems(4))) Then
            '    CboFileStatus2 = Lststr.SelectedItem.SubItems(4)
            'End If
            If Len(Trim(Lststr.SelectedItem.SubItems(5))) Then
                CboUser2 = Lststr.SelectedItem.SubItems(5)
            End If
            If Len(Trim(Lststr.SelectedItem.SubItems(6))) Then
                txtDORqst = Lststr.SelectedItem.SubItems(6)
            End If
            If Len(Trim(Lststr.SelectedItem.SubItems(7))) Then
                txtDOD = Lststr.SelectedItem.SubItems(7)
            End If
            If Len(Trim(Lststr.SelectedItem.SubItems(8))) Then
                txtDOReturn = Lststr.SelectedItem.SubItems(8)
            End If
            If Len(Trim(Lststr.SelectedItem.SubItems(9))) Then
                TxtRemarks = Lststr.SelectedItem.SubItems(9)
            End If
            If Len(Trim(Lststr.SelectedItem.SubItems(10))) Then
                CboDept = Lststr.SelectedItem.SubItems(10)
            End If
            
        End If
    End If
End Sub

Private Sub MaskstrDODelivery_LostFocus()
    If IsDate(MaskstrDODelivery) Then
    Else
        If MaskstrDODelivery <> "__/__/____" Then
            MsgBox "Invalid Date Format! ", vbCritical
            MaskstrDODelivery = "__/__/____"
            MaskstrDODelivery.SetFocus
        End If
    End If
End Sub

Private Sub MaskstrDOReturn_LostFocus()
    If IsDate(MaskstrDOReturn) Then

    Else
        If MaskstrDOReturn <> "__/__/____" Then
            MsgBox "Invalid Date Format! ", vbCritical
            MaskstrDOReturn = "__/__/____"
            MaskstrDOReturn.SetFocus
        End If
    End If

End Sub

Private Sub MaskstrDORqst_LostFocus()
    If IsDate(MaskstrDORqst) Then
        Else
            If MaskstrDORqst <> "__/__/____" Then
                MsgBox "Invalid Date Format! ", vbCritical
                MaskstrDORqst = "__/__/____"
                MaskstrDORqst.SetFocus
            End If
    End If
End Sub

Private Sub txtDOD_LostFocus()
    If IsDate(txtDOD) Then
        Else
            If txtDOD <> "__/__/____" Then
                MsgBox "Invalid Date Format! ", vbCritical
                txtDOD = "__/__/____"
                txtDOD.SetFocus
            End If
    End If

End Sub


Private Sub txtDOReturn_LostFocus()
    If IsDate(txtDOReturn) Then
        Else
            If txtDOReturn <> "__/__/____" Then
                MsgBox "Invalid Date Format! ", vbCritical
                txtDOReturn = "__/__/____"
                txtDOReturn.SetFocus
            End If
    End If
End Sub

Private Sub txtDORqst_LostFocus()
    If IsDate(txtDORqst) Then
        Else
            If txtDORqst <> "__/__/____" Then
                MsgBox "Invalid Date Format! ", vbCritical
                txtDORqst = "__/__/____"
                txtDORqst.SetFocus
            End If
    End If
End Sub
