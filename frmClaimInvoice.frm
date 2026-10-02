VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmClaimInvoice 
   Caption         =   "Claim Invoices"
   ClientHeight    =   5925
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   8100
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   5925
   ScaleWidth      =   8100
   StartUpPosition =   3  'Windows Default
   Begin TabDlg.SSTab SSTab1 
      Height          =   4935
      Left            =   120
      TabIndex        =   5
      Top             =   720
      Width           =   7815
      _ExtentX        =   13785
      _ExtentY        =   8705
      _Version        =   393216
      Tabs            =   2
      Tab             =   1
      TabsPerRow      =   2
      TabHeight       =   520
      TabCaption(0)   =   "Hospitalisation"
      TabPicture(0)   =   "frmClaimInvoice.frx":0000
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "Frame1"
      Tab(0).ControlCount=   1
      TabCaption(1)   =   "Invoices"
      TabPicture(1)   =   "frmClaimInvoice.frx":001C
      Tab(1).ControlEnabled=   -1  'True
      Tab(1).Control(0)=   "Frame17"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).ControlCount=   1
      Begin VB.Frame Frame1 
         Caption         =   "Hospitalisation"
         Height          =   2775
         Left            =   -74880
         TabIndex        =   18
         Top             =   720
         Width           =   7335
         Begin VB.TextBox Text1 
            Height          =   285
            Left            =   1680
            TabIndex        =   21
            Text            =   "Hospital"
            Top             =   480
            Width           =   5295
         End
         Begin VB.TextBox Text2 
            Height          =   285
            Left            =   1680
            TabIndex        =   20
            Text            =   "C - Cash"
            Top             =   840
            Width           =   1335
         End
         Begin VB.TextBox Text3 
            Height          =   285
            Left            =   5160
            TabIndex        =   19
            Text            =   "CH - Hospitalisation"
            Top             =   840
            Width           =   1695
         End
         Begin MSMask.MaskEdBox txtDateAdmitted 
            Height          =   285
            Left            =   1680
            TabIndex        =   22
            Top             =   1200
            Width           =   1560
            _ExtentX        =   2752
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
         Begin MSMask.MaskEdBox txtDateDischarge 
            Height          =   285
            Left            =   5160
            TabIndex        =   23
            Top             =   1200
            Width           =   1560
            _ExtentX        =   2752
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
         Begin MSMask.MaskEdBox txtDateOfLoss 
            Height          =   285
            Left            =   1680
            TabIndex        =   24
            Top             =   1575
            Width           =   1560
            _ExtentX        =   2752
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
         Begin VB.Label Label2 
            Caption         =   "Payment Mode"
            Height          =   255
            Left            =   240
            TabIndex        =   30
            Top             =   855
            Width           =   1335
         End
         Begin VB.Label Label3 
            Caption         =   "Payment Type"
            Height          =   255
            Left            =   3840
            TabIndex        =   29
            Top             =   855
            Width           =   1095
         End
         Begin VB.Label Label5 
            Caption         =   "Date Of Loss"
            Height          =   255
            Left            =   240
            TabIndex        =   28
            Top             =   1575
            Width           =   1095
         End
         Begin VB.Label Label8 
            Caption         =   "Date Discharge"
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
            Left            =   3810
            TabIndex        =   27
            Top             =   1245
            Width           =   1275
         End
         Begin VB.Label Label9 
            Caption         =   "Date Admitted"
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
            TabIndex        =   26
            Top             =   1245
            Width           =   1275
         End
         Begin VB.Label Label27 
            Caption         =   "Hospitals"
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
            TabIndex        =   25
            Top             =   525
            Width           =   1140
         End
      End
      Begin VB.Frame Frame17 
         Caption         =   "Invoices"
         Height          =   4335
         Left            =   240
         TabIndex        =   6
         Top             =   480
         Width           =   7215
         Begin VB.ComboBox CboItems 
            Height          =   315
            Left            =   1560
            TabIndex        =   39
            Text            =   "Daily Room And Borad"
            Top             =   1680
            Width           =   3375
         End
         Begin VB.TextBox txtByMBM 
            Height          =   285
            Left            =   1560
            TabIndex        =   33
            Top             =   2040
            Width           =   1695
         End
         Begin VB.TextBox txtPaidMBM 
            Height          =   285
            Left            =   4920
            TabIndex        =   32
            Top             =   2040
            Width           =   1455
         End
         Begin VB.TextBox txtRemark 
            Height          =   285
            Left            =   1560
            TabIndex        =   31
            Top             =   2400
            Width           =   5055
         End
         Begin VB.ComboBox cboHOSCode 
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
            Left            =   1560
            Sorted          =   -1  'True
            TabIndex        =   12
            Top             =   225
            Width           =   5025
         End
         Begin VB.TextBox txtInvoiceNo 
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
            Left            =   1560
            TabIndex        =   11
            Top             =   600
            Width           =   1740
         End
         Begin VB.TextBox txtAmount 
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
            Left            =   1560
            TabIndex        =   10
            Top             =   960
            Width           =   1740
         End
         Begin VB.CommandButton cmdAddHos 
            Caption         =   "Add"
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
            Left            =   4680
            TabIndex        =   9
            Top             =   3960
            Width           =   960
         End
         Begin VB.CommandButton cmdRemoveHos 
            Caption         =   "Remove"
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
            Left            =   5640
            TabIndex        =   8
            Top             =   3960
            Width           =   960
         End
         Begin VB.TextBox txtPaidHos 
            Height          =   285
            Left            =   1560
            TabIndex        =   7
            Top             =   1320
            Width           =   1695
         End
         Begin MSMask.MaskEdBox txtInvoicedDate 
            Height          =   285
            Left            =   4920
            TabIndex        =   13
            Top             =   600
            Width           =   1740
            _ExtentX        =   3069
            _ExtentY        =   503
            _Version        =   393216
            MaxLength       =   10
            Mask            =   "##/##/####"
            PromptChar      =   "_"
         End
         Begin MSComctlLib.ListView lstInvoice 
            Height          =   1140
            Left            =   240
            TabIndex        =   34
            Top             =   2760
            Width           =   6585
            _ExtentX        =   11615
            _ExtentY        =   2011
            View            =   3
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
            NumItems        =   8
            BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Text            =   "Hospital Code"
               Object.Width           =   2540
            EndProperty
            BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   1
               Text            =   "Invoice No"
               Object.Width           =   1764
            EndProperty
            BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   2
               Text            =   "Invoiced Date"
               Object.Width           =   1940
            EndProperty
            BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   3
               Text            =   "Amount"
               Object.Width           =   1587
            EndProperty
            BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   4
               Text            =   "Paid Hospital"
               Object.Width           =   2540
            EndProperty
            BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   5
               Text            =   "Paid Member"
               Object.Width           =   2540
            EndProperty
            BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   6
               Text            =   "Paid By Member"
               Object.Width           =   2540
            EndProperty
            BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   7
               Text            =   "Remark"
               Object.Width           =   2540
            EndProperty
         End
         Begin VB.Label Label1 
            Caption         =   "Items"
            Height          =   255
            Left            =   240
            TabIndex        =   40
            Top             =   1680
            Width           =   975
         End
         Begin VB.Label Label55 
            Caption         =   "Paid By Member"
            Height          =   255
            Left            =   240
            TabIndex        =   38
            Top             =   2040
            Width           =   1215
         End
         Begin VB.Label Label54 
            Caption         =   "Approve Amount"
            Height          =   255
            Left            =   240
            TabIndex        =   37
            Top             =   1320
            Width           =   1335
         End
         Begin VB.Label Label53 
            Caption         =   "Paid to Member"
            Height          =   255
            Left            =   3480
            TabIndex        =   36
            Top             =   2040
            Width           =   1335
         End
         Begin VB.Label Label25 
            Caption         =   "Remark"
            Height          =   255
            Left            =   240
            TabIndex        =   35
            Top             =   2400
            Width           =   735
         End
         Begin VB.Label Label7 
            Caption         =   "Date of Invoice"
            Height          =   255
            Left            =   3480
            TabIndex        =   17
            Top             =   600
            Width           =   1215
         End
         Begin VB.Label Label6 
            Caption         =   "Hospitals"
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
            TabIndex        =   16
            Top             =   270
            Width           =   1140
         End
         Begin VB.Label Label44 
            Caption         =   "Invoice Number"
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
            TabIndex        =   15
            Top             =   645
            Width           =   1275
         End
         Begin VB.Label Label4 
            Caption         =   "Amount"
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
            TabIndex        =   14
            Top             =   960
            Width           =   1275
         End
      End
   End
   Begin VB.Frame Frame2 
      Height          =   615
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   9855
      Begin VB.TextBox txtMBMNumber 
         Height          =   285
         Left            =   5160
         TabIndex        =   2
         Top             =   240
         Width           =   2295
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
         TabIndex        =   1
         Top             =   240
         Width           =   1740
      End
      Begin VB.Label Label26 
         Caption         =   "Member Number"
         Height          =   255
         Left            =   3600
         TabIndex        =   4
         Top             =   240
         Width           =   1575
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
         Left            =   240
         TabIndex        =   3
         Top             =   240
         Width           =   1275
      End
   End
End
Attribute VB_Name = "frmClaimInvoice"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
