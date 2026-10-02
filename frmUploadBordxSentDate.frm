VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "mscomctl.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form frmUploadBordxSentDate 
   Caption         =   "Upload Bordx Sent Date"
   ClientHeight    =   6045
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   8520
   LinkTopic       =   "Form1"
   ScaleHeight     =   6045
   ScaleWidth      =   8520
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox txtFileName 
      Enabled         =   0   'False
      Height          =   285
      Left            =   150
      TabIndex        =   2
      Top             =   420
      Width           =   7695
   End
   Begin VB.CommandButton cmdFileSearch 
      Caption         =   ">"
      Height          =   255
      Left            =   7920
      TabIndex        =   1
      Top             =   420
      Width           =   375
   End
   Begin MSComctlLib.ListView lstErrorList 
      Height          =   4455
      Left            =   120
      TabIndex        =   0
      Top             =   1440
      Width           =   8295
      _ExtentX        =   14631
      _ExtentY        =   7858
      View            =   3
      LabelEdit       =   1
      MultiSelect     =   -1  'True
      LabelWrap       =   0   'False
      HideSelection   =   0   'False
      AllowReorder    =   -1  'True
      Checkboxes      =   -1  'True
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   6
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Line No"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Bordx No"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "Approval Code"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Text            =   "Approval Date"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   4
         Text            =   "Bordx Amount"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   5
         Text            =   "ReturnDate"
         Object.Width           =   2540
      EndProperty
   End
   Begin VB.Frame Frame1 
      Height          =   1095
      Left            =   80
      TabIndex        =   3
      Top             =   240
      Width           =   8415
      Begin VB.CommandButton cmdUpload 
         Caption         =   "App Date"
         Enabled         =   0   'False
         Height          =   375
         Left            =   4200
         TabIndex        =   8
         Top             =   600
         Width           =   975
      End
      Begin VB.CommandButton cmdVerify 
         Caption         =   "Verify App"
         Height          =   375
         Left            =   3240
         TabIndex        =   7
         Top             =   600
         Width           =   975
      End
      Begin VB.CommandButton cmdExit 
         Caption         =   "Exit"
         Height          =   375
         Left            =   7320
         TabIndex        =   6
         Top             =   600
         Width           =   975
      End
      Begin VB.CommandButton cmdReturnDate 
         Caption         =   "Return Date"
         Enabled         =   0   'False
         Height          =   375
         Left            =   6240
         TabIndex        =   5
         Top             =   600
         Width           =   1095
      End
      Begin VB.CommandButton cmdVerifyReturn 
         Caption         =   "Verify Return"
         Height          =   375
         Left            =   5160
         TabIndex        =   4
         Top             =   600
         Width           =   1095
      End
      Begin MSComDlg.CommonDialog CommonDialog1 
         Left            =   1920
         Top             =   960
         _ExtentX        =   847
         _ExtentY        =   847
         _Version        =   393216
      End
   End
   Begin VB.Label Label28 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00404040&
      Caption         =   "Upload Bordx Sent Date"
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
      TabIndex        =   9
      Top             =   0
      Width           =   8445
   End
End
Attribute VB_Name = "frmUploadBordxSentDate"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
