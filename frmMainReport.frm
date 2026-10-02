VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "CRYSTL32.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmReport 
   Caption         =   "Report"
   ClientHeight    =   7935
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   11805
   LinkTopic       =   "Form2"
   MDIChild        =   -1  'True
   ScaleHeight     =   7935
   ScaleWidth      =   11805
   WindowState     =   2  'Maximized
   Begin VB.Frame Frame1 
      Caption         =   "Enter Criteria For Searching"
      Height          =   2295
      Left            =   120
      TabIndex        =   32
      Top             =   240
      Width           =   11535
      Begin Crystal.CrystalReport cRpt 
         Left            =   1080
         Top             =   1320
         _ExtentX        =   741
         _ExtentY        =   741
         _Version        =   348160
         WindowControlBox=   -1  'True
         WindowMaxButton =   -1  'True
         WindowMinButton =   -1  'True
         PrintFileLinesPerPage=   60
         WindowShowNavigationCtls=   -1  'True
         WindowShowCancelBtn=   -1  'True
         WindowShowPrintBtn=   -1  'True
         WindowShowExportBtn=   -1  'True
         WindowShowZoomCtl=   -1  'True
         WindowShowProgressCtls=   -1  'True
      End
      Begin VB.ComboBox cboHLTCode 
         Height          =   315
         Left            =   4800
         TabIndex        =   4
         Top             =   1320
         Width           =   2415
      End
      Begin VB.ComboBox cboInsurer 
         Height          =   315
         Left            =   2160
         TabIndex        =   2
         Top             =   840
         Width           =   5055
      End
      Begin VB.Frame Frame2 
         Caption         =   "Sort By"
         Height          =   1455
         Left            =   7560
         TabIndex        =   33
         Top             =   240
         Width           =   3495
         Begin VB.OptionButton optMembership 
            Caption         =   "Membership"
            Height          =   375
            Left            =   1080
            TabIndex        =   6
            Top             =   360
            Value           =   -1  'True
            Width           =   1335
         End
         Begin VB.OptionButton optPolicy 
            Caption         =   "Policy Number"
            Height          =   375
            Left            =   1080
            TabIndex        =   7
            Top             =   840
            Width           =   1455
         End
      End
      Begin VB.TextBox txtBatchNo 
         Height          =   315
         Left            =   2160
         TabIndex        =   3
         Top             =   1320
         Width           =   1575
      End
      Begin VB.CommandButton cmdSearch 
         Caption         =   "Search"
         Height          =   375
         Left            =   5520
         TabIndex        =   8
         Top             =   1800
         Width           =   1215
      End
      Begin VB.CommandButton cmdClose 
         Caption         =   "Close"
         Height          =   375
         Left            =   8160
         TabIndex        =   10
         Top             =   1800
         Width           =   1215
      End
      Begin VB.CommandButton cmdClear 
         Caption         =   "Clear"
         Height          =   375
         Left            =   6840
         TabIndex        =   9
         Top             =   1800
         Width           =   1215
      End
      Begin VB.ComboBox cboReport 
         Height          =   315
         ItemData        =   "frmMainReport.frx":0000
         Left            =   2160
         List            =   "frmMainReport.frx":0010
         TabIndex        =   5
         Top             =   1800
         Width           =   1695
      End
      Begin MSMask.MaskEdBox txtBordxDateTo 
         Height          =   315
         Left            =   6000
         TabIndex        =   1
         Top             =   360
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   556
         _Version        =   393216
         MaxLength       =   10
         Mask            =   "##/##/####"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox txtBordxDateFr 
         Height          =   315
         Left            =   2160
         TabIndex        =   0
         Top             =   360
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   556
         _Version        =   393216
         MaxLength       =   10
         Mask            =   "##/##/####"
         PromptChar      =   "_"
      End
      Begin VB.Label lblHTLCode 
         Caption         =   "Health Type"
         Height          =   255
         Left            =   3840
         TabIndex        =   39
         Top             =   1320
         Width           =   975
      End
      Begin VB.Label Label1 
         Caption         =   "Bordereaux Date From"
         Height          =   255
         Left            =   360
         TabIndex        =   38
         Top             =   360
         Width           =   1695
      End
      Begin VB.Label Label2 
         Caption         =   "Bordereaux Date To"
         Height          =   375
         Left            =   4320
         TabIndex        =   37
         Top             =   360
         Width           =   1575
      End
      Begin VB.Label Label3 
         Caption         =   "Insurer's Code"
         Height          =   255
         Left            =   360
         TabIndex        =   36
         Top             =   840
         Width           =   1095
      End
      Begin VB.Label Label4 
         Caption         =   "Batch No"
         Height          =   255
         Left            =   360
         TabIndex        =   35
         Top             =   1320
         Width           =   1095
      End
      Begin VB.Label Label5 
         Caption         =   "Report"
         Height          =   255
         Left            =   360
         TabIndex        =   34
         Top             =   1800
         Width           =   1215
      End
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   5175
      Left            =   120
      TabIndex        =   11
      Top             =   2760
      Width           =   11535
      _ExtentX        =   20346
      _ExtentY        =   9128
      _Version        =   393216
      Style           =   1
      Tabs            =   4
      TabsPerRow      =   4
      TabHeight       =   520
      TabCaption(0)   =   "Leaflet"
      TabPicture(0)   =   "frmMainReport.frx":0042
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Frame4"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Frame3"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "cmdPrint"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).ControlCount=   3
      TabCaption(1)   =   "MCO Listing"
      TabPicture(1)   =   "frmMainReport.frx":005E
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Frame9"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).Control(1)=   "cmdMCOPrintPreview"
      Tab(1).Control(1).Enabled=   0   'False
      Tab(1).Control(2)=   "Frame5"
      Tab(1).Control(2).Enabled=   0   'False
      Tab(1).Control(3)=   "Frame6"
      Tab(1).Control(3).Enabled=   0   'False
      Tab(1).Control(4)=   "cmdMCOPrint"
      Tab(1).Control(4).Enabled=   0   'False
      Tab(1).ControlCount=   5
      TabCaption(2)   =   "Invoice"
      TabPicture(2)   =   "frmMainReport.frx":007A
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "TxtInvoiceYear"
      Tab(2).Control(0).Enabled=   0   'False
      Tab(2).Control(1)=   "cmdINVPrintPreview"
      Tab(2).Control(1).Enabled=   0   'False
      Tab(2).Control(2)=   "txtInvoiceNo"
      Tab(2).Control(2).Enabled=   0   'False
      Tab(2).Control(3)=   "Frame7"
      Tab(2).Control(3).Enabled=   0   'False
      Tab(2).Control(4)=   "Frame8"
      Tab(2).Control(4).Enabled=   0   'False
      Tab(2).Control(5)=   "cmdInvPrint"
      Tab(2).Control(5).Enabled=   0   'False
      Tab(2).ControlCount=   6
      TabCaption(3)   =   "Card Printing"
      TabPicture(3)   =   "frmMainReport.frx":0096
      Tab(3).ControlEnabled=   0   'False
      Tab(3).Control(0)=   "c"
      Tab(3).Control(0).Enabled=   0   'False
      Tab(3).Control(1)=   "cmdCRDPrint"
      Tab(3).Control(1).Enabled=   0   'False
      Tab(3).Control(2)=   "Frame11"
      Tab(3).Control(2).Enabled=   0   'False
      Tab(3).Control(3)=   "Frame12"
      Tab(3).Control(3).Enabled=   0   'False
      Tab(3).ControlCount=   4
      Begin VB.Frame Frame9 
         Caption         =   "Premium"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   -74760
         TabIndex        =   52
         Top             =   960
         Width           =   11055
         Begin VB.OptionButton optMCOWihtoutPremium 
            Caption         =   "Without Premium"
            Height          =   255
            Left            =   4440
            TabIndex        =   54
            Top             =   160
            Width           =   1815
         End
         Begin VB.OptionButton optMCOPremium 
            Caption         =   "With Premium"
            Height          =   255
            Left            =   840
            TabIndex        =   53
            Top             =   160
            Value           =   -1  'True
            Width           =   1575
         End
      End
      Begin VB.CommandButton c 
         Caption         =   "Print Preview"
         Enabled         =   0   'False
         Height          =   495
         Left            =   -73440
         TabIndex        =   51
         Top             =   4500
         Width           =   1215
      End
      Begin VB.CommandButton cmdCRDPrint 
         Caption         =   "Print"
         Enabled         =   0   'False
         Height          =   495
         Left            =   -74760
         TabIndex        =   49
         Top             =   4500
         Width           =   1215
      End
      Begin VB.Frame Frame11 
         Caption         =   "Listing"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   3015
         Left            =   -74760
         TabIndex        =   47
         Top             =   1440
         Width           =   11055
         Begin MSComctlLib.ListView lvwCRD 
            Height          =   2655
            Left            =   120
            TabIndex        =   48
            Top             =   240
            Width           =   10815
            _ExtentX        =   19076
            _ExtentY        =   4683
            View            =   3
            LabelWrap       =   -1  'True
            HideSelection   =   -1  'True
            _Version        =   393217
            ForeColor       =   -2147483640
            BackColor       =   -2147483643
            BorderStyle     =   1
            Appearance      =   1
            NumItems        =   9
            BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Text            =   "Payor Code"
               Object.Width           =   2540
            EndProperty
            BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   1
               Text            =   "Membership No"
               Object.Width           =   2540
            EndProperty
            BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   2
               Text            =   "Bordereaux Date"
               Object.Width           =   2540
            EndProperty
            BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   3
               Text            =   "Policy No"
               Object.Width           =   2540
            EndProperty
            BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   4
               Text            =   "Health Type"
               Object.Width           =   2540
            EndProperty
            BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   5
               Text            =   "Batch No"
               Object.Width           =   2540
            EndProperty
            BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   6
               Text            =   "Name"
               Object.Width           =   2540
            EndProperty
            BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   7
               Text            =   "Printed"
               Object.Width           =   2540
            EndProperty
            BeginProperty ColumnHeader(9) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   8
               Object.Width           =   2540
            EndProperty
         End
      End
      Begin VB.Frame Frame12 
         Caption         =   "Options"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   855
         Left            =   -74760
         TabIndex        =   44
         Top             =   480
         Width           =   11055
         Begin VB.OptionButton optCRDExport 
            Caption         =   "Export"
            Height          =   255
            Left            =   8400
            TabIndex        =   50
            Top             =   360
            Width           =   855
         End
         Begin VB.OptionButton optCRDReprint 
            Caption         =   "Reprint"
            Height          =   255
            Left            =   4440
            TabIndex        =   46
            Top             =   360
            Width           =   855
         End
         Begin VB.OptionButton optCRDPrint 
            Caption         =   "Print"
            Height          =   255
            Left            =   840
            TabIndex        =   45
            Top             =   360
            Value           =   -1  'True
            Width           =   855
         End
      End
      Begin VB.TextBox TxtInvoiceYear 
         Height          =   375
         Left            =   -70920
         TabIndex        =   43
         Top             =   4620
         Visible         =   0   'False
         Width           =   735
      End
      Begin VB.CommandButton cmdINVPrintPreview 
         Caption         =   "Print Preview"
         Enabled         =   0   'False
         Height          =   495
         Left            =   -73440
         TabIndex        =   42
         Top             =   4500
         Width           =   1215
      End
      Begin VB.CommandButton cmdMCOPrintPreview 
         Caption         =   "Print Preview"
         Enabled         =   0   'False
         Height          =   495
         Left            =   -73440
         TabIndex        =   41
         Top             =   4500
         Width           =   1215
      End
      Begin VB.TextBox txtInvoiceNo 
         Height          =   375
         Left            =   -71760
         TabIndex        =   40
         Top             =   4620
         Visible         =   0   'False
         Width           =   735
      End
      Begin VB.CommandButton cmdPrint 
         Caption         =   "Print"
         Enabled         =   0   'False
         Height          =   495
         Left            =   240
         TabIndex        =   20
         Top             =   4500
         Width           =   1215
      End
      Begin VB.Frame Frame3 
         Caption         =   "Listing"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   3015
         Left            =   240
         TabIndex        =   30
         Top             =   1440
         Width           =   11055
         Begin MSComctlLib.ListView lvwLeaflet 
            Height          =   2655
            Left            =   120
            TabIndex        =   31
            Top             =   240
            Width           =   10815
            _ExtentX        =   19076
            _ExtentY        =   4683
            View            =   3
            LabelWrap       =   -1  'True
            HideSelection   =   -1  'True
            _Version        =   393217
            ForeColor       =   -2147483640
            BackColor       =   -2147483643
            BorderStyle     =   1
            Appearance      =   1
            NumItems        =   8
            BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Text            =   "Payor Code"
               Object.Width           =   2540
            EndProperty
            BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   1
               Text            =   "Membership No"
               Object.Width           =   2540
            EndProperty
            BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   2
               Text            =   "Bordereaux Date"
               Object.Width           =   2540
            EndProperty
            BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   3
               Text            =   "Policy No"
               Object.Width           =   2540
            EndProperty
            BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   4
               Text            =   "Health Type"
               Object.Width           =   2540
            EndProperty
            BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   5
               Text            =   "Batch No"
               Object.Width           =   2540
            EndProperty
            BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   6
               Text            =   "Name"
               Object.Width           =   2540
            EndProperty
            BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   7
               Text            =   "Printed"
               Object.Width           =   2540
            EndProperty
         End
      End
      Begin VB.Frame Frame4 
         Caption         =   "Options"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   855
         Left            =   240
         TabIndex        =   29
         Top             =   480
         Width           =   11055
         Begin VB.OptionButton optPrint 
            Caption         =   "Print"
            Height          =   255
            Left            =   840
            TabIndex        =   12
            Top             =   360
            Value           =   -1  'True
            Width           =   855
         End
         Begin VB.OptionButton optReprint 
            Caption         =   "Reprint"
            Height          =   255
            Left            =   4440
            TabIndex        =   13
            Top             =   360
            Width           =   975
         End
         Begin VB.OptionButton optPrintAll 
            Caption         =   "Print All"
            Height          =   255
            Left            =   8400
            TabIndex        =   14
            Top             =   360
            Width           =   975
         End
      End
      Begin VB.Frame Frame5 
         Caption         =   "Options"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   -74760
         TabIndex        =   28
         Top             =   480
         Width           =   11055
         Begin VB.OptionButton optMCOPrintAll 
            Caption         =   "Print All"
            Height          =   255
            Left            =   8400
            TabIndex        =   17
            Top             =   200
            Width           =   975
         End
         Begin VB.OptionButton optMCOReprint 
            Caption         =   "Reprint"
            Height          =   255
            Left            =   4440
            TabIndex        =   16
            Top             =   200
            Width           =   975
         End
         Begin VB.OptionButton optMCOPrint 
            Caption         =   "Print"
            Height          =   255
            Left            =   840
            TabIndex        =   15
            Top             =   200
            Value           =   -1  'True
            Width           =   855
         End
      End
      Begin VB.Frame Frame6 
         Caption         =   "Listing"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   3015
         Left            =   -74760
         TabIndex        =   26
         Top             =   1440
         Width           =   11055
         Begin MSComctlLib.ListView lvwMCO 
            Height          =   2655
            Left            =   120
            TabIndex        =   27
            Top             =   240
            Width           =   10815
            _ExtentX        =   19076
            _ExtentY        =   4683
            View            =   3
            LabelWrap       =   -1  'True
            HideSelection   =   -1  'True
            _Version        =   393217
            ForeColor       =   -2147483640
            BackColor       =   -2147483643
            BorderStyle     =   1
            Appearance      =   1
            NumItems        =   8
            BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Text            =   "Payor Code"
               Object.Width           =   2540
            EndProperty
            BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   1
               Text            =   "Membership No"
               Object.Width           =   2540
            EndProperty
            BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   2
               Text            =   "Bordereaux Date"
               Object.Width           =   2540
            EndProperty
            BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   3
               Text            =   "Policy No"
               Object.Width           =   2540
            EndProperty
            BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   4
               Text            =   "Health Type"
               Object.Width           =   2540
            EndProperty
            BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   5
               Text            =   "Batch No"
               Object.Width           =   2540
            EndProperty
            BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   6
               Text            =   "Name"
               Object.Width           =   2540
            EndProperty
            BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   7
               Text            =   "Printed"
               Object.Width           =   2540
            EndProperty
         End
      End
      Begin VB.CommandButton cmdMCOPrint 
         Caption         =   "Print"
         Enabled         =   0   'False
         Height          =   495
         Left            =   -74760
         TabIndex        =   25
         Top             =   4500
         Width           =   1215
      End
      Begin VB.Frame Frame7 
         Caption         =   "Listing"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   3015
         Left            =   -74760
         TabIndex        =   23
         Top             =   1440
         Width           =   11055
         Begin MSComctlLib.ListView lvwInvoice 
            Height          =   2655
            Left            =   120
            TabIndex        =   24
            Top             =   240
            Width           =   10815
            _ExtentX        =   19076
            _ExtentY        =   4683
            View            =   3
            LabelWrap       =   -1  'True
            HideSelection   =   -1  'True
            _Version        =   393217
            ForeColor       =   -2147483640
            BackColor       =   -2147483643
            BorderStyle     =   1
            Appearance      =   1
            NumItems        =   5
            BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Text            =   "Payor Code"
               Object.Width           =   2540
            EndProperty
            BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   1
               Text            =   "Bordereaux Date"
               Object.Width           =   2540
            EndProperty
            BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   2
               Text            =   "Health Type"
               Object.Width           =   2540
            EndProperty
            BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   3
               Text            =   "Batch No"
               Object.Width           =   2540
            EndProperty
            BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   4
               Text            =   "Printed"
               Object.Width           =   2540
            EndProperty
         End
      End
      Begin VB.Frame Frame8 
         Caption         =   "Options"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   855
         Left            =   -74760
         TabIndex        =   22
         Top             =   480
         Width           =   11055
         Begin VB.OptionButton optInvPrint 
            Caption         =   "Print"
            Height          =   255
            Left            =   840
            TabIndex        =   18
            Top             =   360
            Value           =   -1  'True
            Width           =   855
         End
         Begin VB.OptionButton optInvReprint 
            Caption         =   "Reprint"
            Height          =   255
            Left            =   8400
            TabIndex        =   19
            Top             =   360
            Width           =   975
         End
      End
      Begin VB.CommandButton cmdInvPrint 
         Caption         =   "Print"
         Enabled         =   0   'False
         Height          =   495
         Left            =   -74760
         TabIndex        =   21
         Top             =   4500
         Width           =   1215
      End
   End
End
Attribute VB_Name = "frmReport"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Const LEAF_PAYCODE = 0
Const LEAF_MEMBERNO = 1
Const LEAF_BORDXDATE = 2
Const LEAF_POLICYNO = 3
Const LEAF_HLTCODE = 4
Const LEAF_BATCHNO = 5
Const LEAF_NAME = 6
Const LEAF_PRINTED = 7

Const MCO_PAYCODE = 0
Const MCO_MEMBERNO = 1
Const MCO_BORDXDATE = 2
Const MCO_POLICYNO = 3
Const MCO_HLTCODE = 4
Const MCO_BATCHNO = 5
Const MCO_NAME = 6
Const MCO_PRINTED = 7

Const INV_PAYCODE = 0
Const INV_BORDXDATE = 1
Const INV_HLTCODE = 2
Const INV_BATCHNO = 3
Const INV_PRINTED = 4


Const MBM_PAYCODE = 0
Const MBM_MEMBERNO = 1
Const MBM_BORDXDATE = 2
Const MBM_POLICYNO = 3
Const MBM_HLTCODE = 4
Const MBM_BATCHNO = 5
Const MBM_NAME = 6
Const MBM_PRINTED = 7

Const CRD_PAYCODE = 0
Const CRD_MEMBERNO = 1
Const CRD_BORDXDATE = 2
Const CRD_POLICYNO = 3
Const CRD_HLTCODE = 4
Const CRD_BATCHNO = 5
Const CRD_NAME = 6
Const CRD_PRINTED = 7

Dim DateFrom As String
Dim DateTo As String
Dim Exportsql As String
Dim MBMNumber As String
Sub Create_File()
    Dim fso, txtfile
    Set fso = CreateObject("Scripting.FileSystemObject")
    Set txtfile = fso.CreateTextFile("c:\My Documents\rpt\Testing.xls", True)
    txtfile.Write ("This is a test. ") ' Write a line.
    ' Write a line with a newline character.
    txtfile.WriteLine ("Testing 1, 2, 3.")
    ' Write three newline characters to the file.
    txtfile.WriteBlankLines (3)
    txtfile.Close
End Sub

Private Sub ExpMBMPrintingToExcel(PrintMode As String)
On Error Resume Next
Dim MyXL As Object
Dim xlApp As excel.Application
Dim xlBook As excel.Workbook
Dim xlSheet As excel.Worksheet
Dim rstExp As New ADODB.Recordset
Dim sql As String
Dim i As Integer
Dim tmpPaCode, tmpBordDate, tmpFileName, tmpBordxDate, tmpMBMNo, tmpPCode
Dim temp1, temp2, temp3, temp4, temp5, temp6, temp7 As Variant
Dim Row As Integer

Dim ii As Integer
Dim MemberNo As String

Dim tmpBDate, tmpBatchNo, tmpPayCode, tmpHLTCode, TmpHLT As String
Dim tmpPrintedBDate, tmpPrintedBatchNo, tmpPrintedPaycode, tmpPrintedHLTcode As String
Dim BDateArray() As String
Dim BatchNoArray() As String
Dim PaycodeArray() As String
Dim HLTcodeArray() As String
Dim cnt As Integer
Dim date1 As String
Dim date2 As String
Dim Rcounter As Integer

    DateFrom = CStr(Month(txtBordxDateFr)) & "/" & CStr(Day(txtBordxDateFr)) & "/" & CStr(Year(txtBordxDateFr))
    DateTo = CStr(Month(txtBordxDateTo)) & "/" & CStr(Day(txtBordxDateTo)) & "/" & CStr(Year(txtBordxDateTo))
    sql = "SELECT * FROM VIEW_CardPrinting"
    sql = sql & " WHERE MBMBordxDate >= '" & DateFrom & "'"
    sql = sql & " AND MBMBordxDate <= '" & DateTo & "'"
    sql = sql & " AND PAYCode = '" & Mid(cboInsurer, 1, 2) & "'"
    sql = sql & " AND HLTCode = '" & Mid(cboHLTCode, 1, 1) & "'"
    
    If txtBatchNo <> "" Then
        sql = sql & " AND MBMBatchNo = '" & txtBatchNo & "'"
    End If
    sql = sql & " ORDER BY PAYCode, MBMBordxDate, HLTCode, MBMBatchNo, MBMNumber, MBMCCoverId"
    rstExp.Open sql, medixmasterconn, adOpenKeyset, adLockOptimistic, adCmdText
    
    Set xlApp = CreateObject("Excel.application")
    Set xlBook = xlApp.Workbooks.Open("C:\My Documents\RPT\testing.xls")
    Set xlSheet = xlBook.Sheets(1)
    
    
    With rstExp
        'tmpPaCode = Mid(cboInsurer, 2, 2)
        tmpPaCode = (Trim(Mid(cboInsurer, 5)))
        'tmpPayCode = Mid(cboInsurer, 1, 2)
        tmpBordDate = Format(txtBordxDateFr, "dd/mm/yyyy")
        tmpBordxDate = Format(tmpBordDate, "yymmdd")
        tmpFileName = Mid(cboInsurer, 1, 2) & tmpBordxDate
        
        
    End With
                
        xlSheet.Cells(1, 1).Value = tmpPaCode
        xlSheet.Cells(2, 1).Value = tmpBordDate
        xlSheet.Cells(3, 1).Value = tmpFileName
          
        Row = 5
        Rcounter = 0
        
'If Dir("c:\My Documents\rpt\Testing.xls") = "testing.xls" Then
 If optCRDExport Then
    cnt = 0
      Rcounter = Rcounter + 1
        For ii = 1 To lvwCRD.listitems.Count
            Set lvwCRD.SelectedItem = lvwCRD.listitems(ii)
            If lvwCRD.SelectedItem.SubItems(CRD_PRINTED) = PrintMode Then
                MemberNo = lvwCRD.SelectedItem.SubItems(CRD_MEMBERNO)
                tmpBDate = lvwCRD.SelectedItem.SubItems(CRD_BORDXDATE)
                tmpHLTCode = lvwCRD.SelectedItem.SubItems(CRD_HLTCODE)
                tmpBatchNo = lvwCRD.SelectedItem.SubItems(CRD_BATCHNO)
                tmpPayCode = lvwCRD.SelectedItem
                
                
                rstExp.MoveFirst
                Do Until rstExp.EOF
                If MemberNo = rstExp!MBMCNumber Then
                    With rstExp
                        temp7 = Rcounter
                        temp1 = ![MBMCname]
                        temp2 = ![MBMCIcBcPpNo]
                        temp3 = ![MBMPolicyNo]
                        tmpMBMNo = MemberNo & "-" & !MBMCCoverID
                        temp4 = tmpMBMNo
                        temp5 = ![PLNCode]
                        temp6 = ![MBMPayorEXPDate]
                    End With
                    xlSheet.Cells(Row, 1).Value = temp7
                    xlSheet.Cells(Row, 2).Value = temp1
                    xlSheet.Cells(Row, 3).Value = temp2
                    xlSheet.Cells(Row, 4).Value = temp3
                    xlSheet.Cells(Row, 5).Value = temp4
                    xlSheet.Cells(Row, 6).Value = temp5
                    xlSheet.Cells(Row, 7).Value = temp6
                    Rcounter = Rcounter + 1
                    Row = Row + 1
                End If
                rstExp.MoveNext
                Loop
                
                
                If tmpBDate = tmpPrintedBDate And tmpBatchNo = tmpPrintedBatchNo And tmpPayCode = tmpPrintedPaycode Then
                Else
                   tmpPrintedBDate = tmpBDate
                   tmpPrintedBatchNo = tmpBatchNo
                   tmpPrintedPaycode = tmpPayCode
                   tmpPrintedHLTcode = tmpHLTCode
                   
                   cnt = cnt + 1
                   ReDim Preserve BDateArray(cnt)
                   ReDim Preserve BatchNoArray(cnt)
                   ReDim Preserve PaycodeArray(cnt)
                   ReDim Preserve HLTcodeArray(cnt)
                   BDateArray(cnt) = tmpBDate
                   BatchNoArray(cnt) = tmpBatchNo
                   PaycodeArray(cnt) = tmpPayCode
                   HLTcodeArray(cnt) = tmpHLTCode
                End If
            End If
        Next ii
End If
    
    If cmdCRDPrint Then
        If optCRDPrint And lvwCRD.SelectedItem.SubItems(CRD_PRINTED) = "YES" Then
            MsgBox "This report already printed, please select 'Reprint' ", vbOKOnly + vbCritical, DialogBoxTitle
        Else
            If optCRDReprint And lvwCRD.SelectedItem.SubItems(CRD_PRINTED) = "NO" Then
                MsgBox "Please select 'Print' option for this report", vbOKOnly + vbCritical, DialogBoxTitle
            Else
                'Update MBMPrinting
                For ii = 1 To UBound(BDateArray)
                    UpdateMBMPrintingFile BDateArray(ii), BatchNoArray(ii), PaycodeArray(ii), HLTcodeArray(ii), "MBM"
                Next ii
                
                'Export the record to the excel application
                tmpHLTCode = lvwCRD.SelectedItem.SubItems(CRD_HLTCODE)
                tmpBatchNo = lvwCRD.SelectedItem.SubItems(CRD_BATCHNO)
                tmpPayCode = lvwCRD.SelectedItem
                DateFrom = Trim(txtBordxDateFr)
                DateTo = Trim(txtBordxDateTo)
            
                ss0 = "{VIEW_CardPrinting.HLTCode} = '" & TmpHLT & "'"
                ss = "{VIEW_CardPrinting.MBMBordxDate} = '" & DateFrom & "'"
                ss1 = "{VIEW_CardPrinting.MBMBordxDate} = '" & DateTo & "'"
                cRpt.ReportFileName = "C:\My Documents\RPT\MBMCardPrinting.rpt"
                cRpt.Destination = crptToPrinter
                cRpt.SelectionFormula = " {VIEW_CardPrinting.MBMBordxDate} >= date(" & Year(txtBordxDateFr) & "," & Month(txtBordxDateFr) & "," & Day(txtBordxDateFr) & ") " & " And " & _
                                        " {VIEW_CardPrinting.MBMBordxDate} <= date(" & Year(txtBordxDateTo) & "," & Month(txtBordxDateTo) & "," & Day(txtBordxDateTo) & ") " & " And " & _
                                        " {VIEW_CardPrinting.PayCode} = '" & tmpPayCode & "'" & " And " & _
                                        " {VIEW_CardPrinting.MBMBatchNo} = " & tmpBatchNo & ""
                cRpt.Action = 1
                
            End If
        End If
    Else
        If cmdCRDPrintPreview Then
           ' If optCRDPrint And lvwCRD.SelectedItem.SubItems(CRD_PRINTED) = "YES" Then
           '     MsgBox "This report already printed, please select 'Reprint' ", vbOKOnly + vbCritical, DialogBoxTitle
           ' Else
           '     If optCRDReprint And lvwCRD.SelectedItem.SubItems(CRD_PRINTED) = "NO" Then
           '         MsgBox "Please select 'Print' option for this report", vbOKOnly + vbCritical, DialogBoxTitle
           '     Else
                    'Update MBMPrinting
           '         For ii = 1 To UBound(BDateArray)
           '             UpdateMBMPrintingFile BDateArray(ii), BatchNoArray(ii), PaycodeArray(ii), HLTcodeArray(ii), "MBM"
           '         Next ii
                
                    'Export the record to the excel application
                    tmpHLTCode = lvwCRD.SelectedItem.SubItems(CRD_HLTCODE)
                    tmpBatchNo = lvwCRD.SelectedItem.SubItems(CRD_BATCHNO)
                    tmpPayCode = lvwCRD.SelectedItem
                    DateFrom = Trim(txtBordxDateFr)
                    DateTo = Trim(txtBordxDateTo)
                
                    ss0 = "{VIEW_CardPrinting.HLTCode} = '" & TmpHLT & "'"
                    ss = "{VIEW_CardPrinting.MBMBordxDate} = '" & DateFrom & "'"
                    ss1 = "{VIEW_CardPrinting.MBMBordxDate} = '" & DateTo & "'"
                    cRpt.ReportFileName = "C:\My Documents\RPT\MBMCardPrinting.rpt"
                    cRpt.Destination = crptToWindow
                    cRpt.SelectionFormula = " {VIEW_CardPrinting.MBMBordxDate} >= date(" & Year(txtBordxDateFr) & "," & Month(txtBordxDateFr) & "," & Day(txtBordxDateFr) & ") " & " And " & _
                                            " {VIEW_CardPrinting.MBMBordxDate} <= date(" & Year(txtBordxDateTo) & "," & Month(txtBordxDateTo) & "," & Day(txtBordxDateTo) & ") " & " And " & _
                                            " {VIEW_CardPrinting.PayCode} = '" & tmpPayCode & "'" & " And " & _
                                            " {VIEW_CardPrinting.MBMBatchNo} = " & tmpBatchNo & ""
                    cRpt.Action = 1
                    cRpt.ParameterFields(0) = "PayCode;" & cboInsurer & ";true"
                    cRpt.ParameterFields(1) = "BordxDate;" & DateFrom & ";true"
                    cRpt.ParameterFields(2) = "FileName;" & tmpFileName & ";true"
                    
                    'rstExp.Close
                    'xlBook.Save
                    'xlBook.Close
            
                    'Set xlBook = Nothing
                    'Set xlApp = Nothing
           ' End If
        'End If
        End If
    End If
rstExp.Close
xlBook.Save
xlBook.Close
Set xlBook = Nothing
Set xlApp = Nothing

End Sub

Private Sub IncInvoiceNo()
On Error Resume Next
    Dim MBMMCOINVQ As New ADODB.Recordset
    Dim strsql As String
    Dim InvoiceNo As String
    Dim Code As String
    
    
    strsql = "Select * from MBMMCOINVQ Where MCOYear = '" & Year(Date) & "'"
    MBMMCOINVQ.Open strsql, medixmasterconn, adOpenKeyset, adLockOptimistic
    
    If MBMMCOINVQ.recordCount = 0 Then
        With MBMMCOINVQ
            .AddNew
                !MCOInvoice_No = "000001"
                InvoiceNo = "000001"
                !MCOYear = Year(Date)
            .Update
        End With
    Else
    
        With MBMMCOINVQ
                !MCOInvoice_No = Format(!MCOInvoice_No + 1, "000000")
                InvoiceNo = Format(!MCOInvoice_No + 1, "000000")
        .Update
        End With
    End If

    txtInvoiceNo = InvoiceNo
    TxtInvoiceYear = Year(Date)
    
    MBMMCOINVQ.Close
    Set MBMMCOINVQ = Nothing
            
End Sub

Private Sub HLTComboListing()
    Dim HLT As New ADODB.Recordset
  
    HLT.Open "HLT", medixmasterconn, adOpenKeyset, adLockOptimistic
        
    Do Until HLT.EOF
        cboHLTCode.AddItem HLT!HLTCode & " - " & Trim(HLT!HLTDescription)
        HLT.MoveNext
    Loop
    HLT.Close
    
End Sub

Sub ComboListing()
    Dim PAY As New ADODB.Recordset
    Dim sqlstr As String
   
    sqlstr = "select PAYCode , PAYCompanyName , PayCompanyFullName from PAY "
    
    PAY.Open sqlstr, medixmasterconn, adOpenKeyset, adLockOptimistic
    
    Do While Not PAY.EOF
        cboInsurer.AddItem (PAY!PayCode) & " - " & (PAY!PAYCompanyName)
        PAY.MoveNext
    Loop
    
    PAY.Close
    Set PAY = Nothing
End Sub

Private Sub RecordSearching(objListView As Object)
On Error Resume Next
Dim rsSearch As New ADODB.Recordset
Dim rsPrint As New ADODB.Recordset
Dim ColumnItem As ListItem
Dim sql As String
Dim itemToPrint As Integer
Dim isPrinted As Boolean
Dim CoverQty As Integer

If cboReport.ListIndex = 2 Then
    objListView.listitems.Clear
    DateFrom = CStr(Month(txtBordxDateFr)) & "/" & CStr(Day(txtBordxDateFr)) & "/" & CStr(Year(txtBordxDateFr))
    DateTo = CStr(Month(txtBordxDateTo)) & "/" & CStr(Day(txtBordxDateTo)) & "/" & CStr(Year(txtBordxDateTo))
    sql = "SELECT * FROM MBM "
    sql = sql & " WHERE MBMBordxDate >= '" & DateFrom & "'"
    sql = sql & " AND MBMBordxDate <= '" & DateTo & "'"
    sql = sql & " AND PAYCode = '" & Mid(cboInsurer, 1, 2) & "'"
    sql = sql & " AND HLTCode = '" & Mid(cboHLTCode, 1, 1) & "'"
    If txtBatchNo <> "" Then
        sql = sql & " AND MBMBatchNo = '" & txtBatchNo & "'"
    End If
    sql = sql & " ORDER BY PAYCode, MBMBordxDate, HLTCode, MBMBatchNo"
    rsSearch.Open sql, medixmasterconn, adOpenKeyset, adLockOptimistic, adCmdText
    
    If rsSearch.recordCount = 0 Then
       MsgBox "Record not found !", vbInformation + vbOKOnly
       Exit Sub
    End If
    
    
    
    Do While Not rsSearch.EOF
        Set ColumnItem = objListView.listitems.Add(, , rsSearch!PayCode)
        ColumnItem.SubItems(1) = rsSearch!MBMBordxDate
        ColumnItem.SubItems(2) = rsSearch!HLTCode
        
        If Len(rsSearch!MBMBatchNo) = 0 Then
            rsSearch!MBMBatchNo = 1
            ColumnItem.SubItems(3) = rsSearch!MBMBatchNo
        Else
            ColumnItem.SubItems(3) = rsSearch!MBMBatchNo
        End If
        
        
        sql = "SELECT * FROM MBMPrinting"
        DateFrom = CStr(Month(rsSearch!MBMBordxDate)) & "/" & CStr(Day(rsSearch!MBMBordxDate)) & "/" & CStr(Year(rsSearch!MBMBordxDate))
        sql = sql & " WHERE MBMBordxDate = '" & DateFrom & "'"
        sql = sql & " AND PAYCode = '" & rsSearch!PayCode & "'"
        sql = sql & " AND HLTCode = '" & rsSearch!HLTCode & "'"
        If txtBatchNo = "" Then
            sql = sql & " AND MBMBatchno = '" & rsSearch!MBMBatchNo & "'"
        Else
            sql = sql & " AND MBMBatchno = '" & txtBatchNo & "'"
        End If
        
        rsPrint.Open sql, medixmasterconn, adOpenKeyset, adLockOptimistic
        Select Case SSTab1.Tab
        Case 0 'Leaflet
            itemToPrint = LEAF_PRINTED
            cmdPrint.Enabled = True
        Case 1 'MCO Listing
            itemToPrint = MCO_PRINTED
            cmdMCOPrint.Enabled = True
            cmdMCOPrintPreview.Enabled = True
        Case 2 'Invoice
            itemToPrint = INV_PRINTED
            cmdInvPrint.Enabled = True
            cmdINVPrintPreview.Enabled = True
        Case 3 'Card Printing
            itemToPrint = CRD_PRINTED
            cmdCRDPrint.Enabled = True
            cmdCRDPrintPreview.Enabled = True
        End Select
        
        If rsPrint.recordCount > 0 Then
            Select Case SSTab1.Tab
            Case 0 'Leaflet
                isPrinted = rsPrint!MBMPrinted
            Case 1 'MCO Listing
                isPrinted = rsPrint!MBMMCOListPrinted
            Case 2 'Invoice
                isPrinted = rsPrint!MBMInvoiced
            Case 3 'Card Printing
                isPrinted = rsPrint!MBMExported
            End Select
            If isPrinted = True Then
                ColumnItem.SubItems(itemToPrint) = "YES"
            Else
                ColumnItem.SubItems(itemToPrint) = "NO"
            End If
        Else
            ColumnItem.SubItems(itemToPrint) = "NO"
        End If
        rsPrint.Close
        rsSearch.MoveNext
    Loop
    
    rsSearch.Close
   
    Set rsPrint = Nothing
    Set rsSearch = Nothing
Else
    objListView.listitems.Clear
    DateFrom = CStr(Month(txtBordxDateFr)) & "/" & CStr(Day(txtBordxDateFr)) & "/" & CStr(Year(txtBordxDateFr))
    DateTo = CStr(Month(txtBordxDateTo)) & "/" & CStr(Day(txtBordxDateTo)) & "/" & CStr(Year(txtBordxDateTo))
    sql = "SELECT * FROM MBM"
    sql = sql & " WHERE MBMBordxDate >= '" & DateFrom & "'"
    sql = sql & " AND MBMBordxDate <= '" & DateTo & "'"
    sql = sql & " AND PAYCode = '" & Mid(cboInsurer, 1, 2) & "'"
    sql = sql & " AND HLTCode = '" & Mid(cboHLTCode, 1, 1) & "'"
    
    If txtBatchNo <> "" Then
        sql = sql & " AND MBMBatchNo = '" & txtBatchNo & "'"
    End If
    sql = sql & " ORDER BY PAYCode, MBMBordxDate, HLTCode, MBMBatchNo"
    rsSearch.Open sql, medixmasterconn, adOpenKeyset, adLockOptimistic, adCmdText
    
    If rsSearch.recordCount = 0 Then
       MsgBox "Record not found !", vbInformation + vbOKOnly
       Exit Sub
    End If
    
    Do While Not rsSearch.EOF
        Set ColumnItem = objListView.listitems.Add(, , rsSearch!PayCode)
        ColumnItem.SubItems(1) = rsSearch!MBMNumber
        ColumnItem.SubItems(2) = rsSearch!MBMBordxDate
        ColumnItem.SubItems(3) = rsSearch!MBMPolicyNo
        ColumnItem.SubItems(4) = rsSearch!HLTCode
        ColumnItem.SubItems(5) = rsSearch!MBMBatchNo
        ColumnItem.SubItems(6) = rsSearch!MBMName
        
        sql = "SELECT * FROM MBMPrinting"
        DateFrom = CStr(Month(rsSearch!MBMBordxDate)) & "/" & CStr(Day(rsSearch!MBMBordxDate)) & "/" & CStr(Year(rsSearch!MBMBordxDate))
        sql = sql & " WHERE MBMBordxDate = '" & DateFrom & "'"
        sql = sql & " AND PAYCode = '" & rsSearch!PayCode & "'"
        sql = sql & " AND HLTCode = '" & rsSearch!HLTCode & "'"
        
        If txtBatchNo = "" Then
            sql = sql & " AND MBMBatchno = '" & rsSearch!MBMBatchNo & "'"
        Else
            sql = sql & " AND MBMBatchno = '" & txtBatchNo & "'"
        End If
        
        rsPrint.Open sql, medixmasterconn, adOpenKeyset, adLockOptimistic
        Select Case SSTab1.Tab
        Case 0 'Leaflet
            itemToPrint = LEAF_PRINTED
            cmdPrint.Enabled = True
        Case 1 'MCO Listing
            itemToPrint = MCO_PRINTED
            cmdMCOPrint.Enabled = True
            cmdMCOPrintPreview.Enabled = True
        Case 2 'Invoice
            itemToPrint = INV_PRINTED
            cmdInvPrint.Enabled = True
            cmdINVPrintPreview.Enabled = True
        Case 3 'Card Printing
            itemToPrint = CRD_PRINTED
            cmdCRDPrint.Enabled = True
            cmdCRDPrintPreview.Enabled = True
        End Select
        
        If rsPrint.recordCount > 0 Then
            Select Case SSTab1.Tab
            Case 0 'Leaflet
                isPrinted = rsPrint!MBMPrinted
            Case 1 'MCO Listing
                isPrinted = rsPrint!MBMMCOListPrinted
            Case 2 'Invoice
                isPrinted = rsPrint!MBMInvoiced
            Case 3 'Card Printing
                isPrinted = rsPrint!MBMExported
            End Select
            If isPrinted = True Then
                ColumnItem.SubItems(itemToPrint) = "YES"
            Else
                ColumnItem.SubItems(itemToPrint) = "NO"
            End If
        Else
            ColumnItem.SubItems(itemToPrint) = "NO"
        End If
        rsPrint.Close
        rsSearch.MoveNext
    Loop
    
    rsSearch.Close
   
    Set rsPrint = Nothing
    Set rsSearch = Nothing
End If
End Sub

Private Sub UpdateMBMPrintingFile(BordxDate As String, BatchNo As String, PayCode As String, HLTCode As String, WhichField As String)
'On Error Resume Next
Dim BordxDate As String
Dim BatchNo As String
Dim PayCode As String
Dim HLTCode As String
Dim WhichField As String
Dim rsUpdate As New ADODB.Recordset
Dim sql As String
    
    'Update MBMPrinting
    sql = "SELECT * FROM MBMPrinting"
    sql = sql & " WHERE MBMBordxDate = '" & Format(CDate(BordxDate), "MM/DD/YYYY") & "'"
    sql = sql & " AND MBMBatchNo = '" & BatchNo & "'"
    sql = sql & " AND PAYCode = '" & PayCode & "'"
    sql = sql & " AND HLTCode = '" & HLTCode & "'"
    
    rsUpdate.Open sql, medixmasterconn, adOpenKeyset, adLockOptimistic, adCmdText
    If rsUpdate.recordCount = 0 Then 'AddNew
        rsUpdate.AddNew
        rsUpdate!MBMBordxDate = BordxDate
        rsUpdate!MBMBatchNo = BatchNo
        rsUpdate!PayCode = PayCode
        rsUpdate!HLTCode = HLTCode
        
        Select Case WhichField
        Case "LEAFLET"
            rsUpdate!MBMPrinted = "1"
            rsUpdate!MBMPrintedDate = Date
        Case "MCO"
            rsUpdate!MBMMCOListPrinted = "1"
            rsUpdate!MBMMCOListDate = Date
        Case "MBM"
            rsUpdate!MBMExpListed = "1"
            rsUpdate!MBMExpListDate = Date
        Case "INV"
            rsUpdate!MBMInvoiced = "1"
            rsUpdate!MBMInvoiceDate = Date
            rsUpdate!MBMInvoiceNo = TxtInvoiceYear & txtInvoiceNo
        Case "CRD"
            rsUpdate!MBMExpListed = "1"
            rdupdate!MBMExpListDate = Date
        End Select
        rsUpdate.Update
    Else 'Update
        Select Case WhichField
        Case "LEAFLET"
            rsUpdate!MBMPrinted = "1"
            rsUpdate!MBMPrintedDate = Date
        Case "MCO"
            rsUpdate!MBMMCOListPrinted = "1"
            rsUpdate!MBMMCOListDate = Date
        Case "MBM"
            rsUpdate!MBMExpListed = "1"
            rsUpdate!MBMExpListDate = Date
        Case "INV"
            rsUpdate!MBMInvoiced = "1"
            rsUpdate!MBMInvoiceDate = Date
            rsUpdate!MBMInvoiceNo = TxtInvoiceYear & txtInvoiceNo
        Case "CRD"
            rsUpdate!MBMExpListed = "1"
            rsUpdate!MBMExpListDate = Date
        End Select
        rsUpdate.Update
    End If
    rsUpdate.Close

End Sub

Private Sub LeafLetPrint(PrintMode As String)
Dim ii As Integer
Dim MemberNo As String

Dim tmpBDate, tmpBatchNo, tmpPayCode, tmpHLTCode As String
Dim tmpPrintedBDate, tmpPrintedBatchNo, tmpPrintedPaycode, tempPrintedHLTcode As String
Dim BDateArray() As String
Dim BatchNoArray() As String
Dim PaycodeArray() As String
Dim HLTcodeArray() As String
Dim cnt As Integer

    cnt = 0
    If PrintMode = "ALL" Then
        For ii = 1 To lvwLeaflet.listitems.Count
         'If lvwLeaflet.ListItems.Count > 0 Then
            Set lvwLeaflet.SelectedItem = lvwLeaflet.listitems(ii)
            MemberNo = lvwLeaflet.SelectedItem.SubItems(LEAF_MEMBERNO)
            tmpBDate = lvwLeaflet.SelectedItem.SubItems(LEAF_BORDXDATE)
            tmpHLTCode = lvwLeaflet.SelectedItem.SubItems(LEAF_HLTCODE)
            tmpBatchNo = lvwLeaflet.SelectedItem.SubItems(LEAF_BATCHNO)
            tmpPayCode = lvwLeaflet.SelectedItem
            
            cRpt.ReportFileName = "C:\My Documents\RPT\Leaflet.rpt"
            cRpt.Destination = crptToPrinter
            cRpt.SelectionFormula = "{VIEW_LEAFLET.MBMNumber} = '" & MemberNo & "'"
            cRpt.Action = 1
            
            If tmpBDate = tmpPrintedBDate And tmpBatchNo = tmpPrintedBatchNo And tmpPayCode = tmpPrintedPaycode And _
               tmpHLTCode = tempPrintedHLTcode Then
            Else
               tmpPrintedBDate = tmpBDate
               tmpPrintedBatchNo = tmpBatchNo
               tmpPrintedPaycode = tmpPayCode
               tempPrintedHLTcode = tmpHLTCode
               cnt = cnt + 1
               ReDim Preserve BDateArray(cnt)
               ReDim Preserve BatchNoArray(cnt)
               ReDim Preserve PaycodeArray(cnt)
               ReDim Preserve HLTcodeArray(cnt)
               BDateArray(cnt) = tmpBDate
               BatchNoArray(cnt) = tmpBatchNo
               PaycodeArray(cnt) = tmpPayCode
               HLTcodeArray(cnt) = tmpHLTCode
            End If
        'End If
        Next ii
    Else
        'For ii = 1 To lvwLeaflet.ListItems.Count
         'If lvwLeaflet.ListItems.Count > 0 Then
            For ii = 1 To lvwLeaflet.listitems.Count
                Set lvwLeaflet.SelectedItem = lvwLeaflet.listitems(ii)   'lvwLeaflet.ListItems(ii)
                If lvwLeaflet.SelectedItem.SubItems(LEAF_PRINTED) = PrintMode Then
                    MemberNo = lvwLeaflet.SelectedItem.SubItems(LEAF_MEMBERNO)
                    tmpBDate = lvwLeaflet.SelectedItem.SubItems(LEAF_BORDXDATE)
                    tmpHLTCode = lvwLeaflet.SelectedItem.SubItems(LEAF_HLTCODE)
                    tmpBatchNo = lvwLeaflet.SelectedItem.SubItems(LEAF_BATCHNO)
                    tmpPayCode = lvwLeaflet.SelectedItem
                    
                    cRpt.ReportFileName = "C:\My Documents\RPT\Leaflet.rpt"
                    cRpt.Destination = crptToPrinter
                    cRpt.SelectionFormula = "{VIEW_LEAFLET.MBMNumber} = '" & MemberNo & "'"
                    cRpt.Action = 1
                    
                
                    If tmpBDate = tmpPrintedBDate And tmpBatchNo = tmpPrintedBatchNo And tmpPayCode = tmpPrintedPaycode And _
                        tmpHLTCode = tempPrintedHLTcode Then
                    Else
                        tmpPrintedBDate = tmpBDate
                        tmpPrintedBatchNo = tmpBatchNo
                        tmpPrintedPaycode = tmpPayCode
                        tempPrintedHLTcode = tmpHLTCode
                   
                        cnt = cnt + 1
                        ReDim Preserve BDateArray(cnt)
                        ReDim Preserve BatchNoArray(cnt)
                        ReDim Preserve PaycodeArray(cnt)
                        ReDim Preserve HLTcodeArray(cnt)
                        BDateArray(cnt) = tmpBDate
                        BatchNoArray(cnt) = tmpBatchNo
                        PaycodeArray(cnt) = tmpPayCode
                        HLTcodeArray(cnt) = tmpHLTCode
                   End If
                End If
            Next ii
        'End If
   End If

    'Update MBMPrinting
    If optPrint And lvwLeaflet.SelectedItem.SubItems(LEAF_PRINTED) = "YES" Then
       MsgBox "This report already printed, please select 'Reprint' ", vbOKOnly + vbCritical, DialogBoxTitle
    Else
        If optReprint And lvwLeaflet.SelectedItem.SubItems(LEAF_PRINTED) = "NO" Then
            MsgBox "Please select 'Print' option for this report", vbOKOnly + vbCritical, DialogBoxTitle
        Else
            For ii = 1 To UBound(BDateArray)
                UpdateMBMPrintingFile BDateArray(ii), BatchNoArray(ii), PaycodeArray(ii), HLTcodeArray(ii), "LEAFLET"
            Next ii
        End If
    End If
End Sub

Private Sub MCOPrint(PrintMode As String)
' On Error Resume Next
Dim ii As Integer
Dim MemberNo As String

Dim tmpBDate, tmpBatchNo, tmpPayCode, tmpHLTCode As String
Dim tmpPrintedBDate, tmpPrintedBatchNo, tmpPrintedPaycode, tmpPrintedHLTcode As String
Dim BDateArray() As String
Dim BatchNoArray() As String
Dim PaycodeArray() As String
Dim HLTcodeArray() As String
Dim cnt As Integer
Dim date1 As String
Dim date2 As String
Dim TmpHLT As String


    cnt = 0
    If PrintMode = "ALL" Then
        For ii = 1 To lvwMCO.listitems.Count
            Set lvwMCO.SelectedItem = lvwMCO.listitems(ii)
            MemberNo = lvwMCO.SelectedItem.SubItems(MCO_MEMBERNO)
            tmpBDate = lvwMCO.SelectedItem.SubItems(MCO_BORDXDATE)
            tmpHLTCode = lvwMCO.SelectedItem.SubItems(MCO_HLTCODE)
            tmpBatchNo = lvwMCO.SelectedItem.SubItems(MCO_BATCHNO)
            tmpPayCode = lvwMCO.SelectedItem
            
            If tmpBDate = tmpPrintedBDate And tmpBatchNo = tmpPrintedBatchNo And tmpPayCode = tmpPrintedPaycode And _
                tmpHLTCode = tmpPrintedHLTcode Then
            Else
               tmpPrintedBDate = tmpBDate
               tmpPrintedBatchNo = tmpBatchNo
               tmpPrintedPaycode = tmpPayCode
               tmpPrintedHLTcode = tmpHLTCode
               
               cnt = cnt + 1
               ReDim Preserve BDateArray(cnt)
               ReDim Preserve BatchNoArray(cnt)
               ReDim Preserve PaycodeArray(cnt)
               ReDim Preserve HLTcodeArray(cnt)
               BDateArray(cnt) = tmpBDate
               BatchNoArray(cnt) = tmpBatchNo
               PaycodeArray(cnt) = tmpPayCode
               HLTcodeArray(cnt) = tmpHLTCode
            End If
        Next ii
    Else
        
        For ii = 1 To lvwMCO.listitems.Count
            Set lvwMCO.SelectedItem = lvwMCO.listitems(ii)
            If lvwMCO.SelectedItem.SubItems(MCO_PRINTED) = PrintMode Then
                MemberNo = lvwMCO.SelectedItem.SubItems(MCO_MEMBERNO)
                tmpBDate = lvwMCO.SelectedItem.SubItems(MCO_BORDXDATE)
                tmpHLTCode = lvwMCO.SelectedItem.SubItems(MCO_HLTCODE)
                tmpBatchNo = lvwMCO.SelectedItem.SubItems(MCO_BATCHNO)
                tmpPayCode = lvwMCO.SelectedItem
                
                If tmpBDate = tmpPrintedBDate And tmpBatchNo = tmpPrintedBatchNo And tmpPayCode = tmpPrintedPaycode Then
                Else
                   tmpPrintedBDate = tmpBDate
                   tmpPrintedBatchNo = tmpBatchNo
                   tmpPrintedPaycode = tmpPayCode
                   tmpPrintedHLTcode = tmpHLTCode
                   
                   cnt = cnt + 1
                   ReDim Preserve BDateArray(cnt)
                   ReDim Preserve BatchNoArray(cnt)
                   ReDim Preserve PaycodeArray(cnt)
                   ReDim Preserve HLTcodeArray(cnt)
                   BDateArray(cnt) = tmpBDate
                   BatchNoArray(cnt) = tmpBatchNo
                   PaycodeArray(cnt) = tmpPayCode
                   HLTcodeArray(cnt) = tmpHLTCode
                End If
            End If
        Next ii
    End If
    
If cmdMCOPrintPreview Then
    If optMCOPremium = True Then
        TmpHLT = lvwMCO.SelectedItem.SubItems(MCO_HLTCODE)
        tmpBatchNo = lvwMCO.SelectedItem.SubItems(MCO_BATCHNO)
        tmpPayCode = lvwMCO.SelectedItem
        DateFrom = Trim(txtBordxDateFr)
        DateTo = Trim(txtBordxDateTo)
        
        ss0 = "{VIEW_MCO_LISTING.HLTCode} = '" & TmpHLT & "'"
        ss = "{VIEW_MCO_LISTING.MBMBordxDate} = '" & DateFrom & "'"
        ss1 = "{VIEW_MCO_LISTING.MBMBordxDate} = '" & DateTo & "'"
        cRpt.ReportFileName = "C:\My Documents\RPT\MCOListing.rpt"
        cRpt.Destination = crptToWindow
        cRpt.WindowState = crptMaximized
        cRpt.SelectionFormula = " {VIEW_MCO_LISTING.MBMBordxDate} >= date(" & Year(txtBordxDateFr) & "," & Month(txtBordxDateFr) & "," & Day(txtBordxDateFr) & ") " & " And " & _
                                " {VIEW_MCO_LISTING.MBMBordxDate} <= date(" & Year(txtBordxDateTo) & "," & Month(txtBordxDateTo) & "," & Day(txtBordxDateTo) & ") " & "AND " & _
                                " {VIEW_MCO_LISTING.HLTCode} = '" & TmpHLT & "'" & "And " & _
                                " {VIEW_MCO_LISTING.PayCode} = '" & tmpPayCode & "'" & " And " & _
                                " {VIEW_MCO_LISTING.MBMBatchNo} = " & tmpBatchNo & ""
        
        cRpt.ParameterFields(0) = "FromDate;" & DateFrom & ";true"
        cRpt.ParameterFields(1) = "LtDate;" & DateTo & ";true"
        cRpt.ParameterFields(2) = "PayorName;" & cboInsurer & ";true"
        cRpt.Action = 1
    Else
        If optMCOWithoutPremium = True Then
            TmpHLT = lvwMCO.SelectedItem.SubItems(MCO_HLTCODE)
            tmpBatchNo = lvwMCO.SelectedItem.SubItems(MCO_BATCHNO)
            tmpPayCode = lvwMCO.SelectedItem
            DateFrom = Trim(txtBordxDateFr)
            DateTo = Trim(txtBordxDateTo)
            
            ss0 = "{VIEW_MCO_LISTING.HLTCode} = '" & TmpHLT & "'"
            ss = "{VIEW_MCO_LISTING.MBMBordxDate} = '" & DateFrom & "'"
            ss1 = "{VIEW_MCO_LISTING.MBMBordxDate} = '" & DateTo & "'"
            cRpt.ReportFileName = "C:\My Documents\RPT\MCOListing.rpt"
            cRpt.Destination = crptToWindow
            cRpt.WindowState = crptMaximized
            cRpt.SelectionFormula = " {VIEW_MCO_LISTING.MBMBordxDate} >= date(" & Year(txtBordxDateFr) & "," & Month(txtBordxDateFr) & "," & Day(txtBordxDateFr) & ") " & " And " & _
                                    " {VIEW_MCO_LISTING.MBMBordxDate} <= date(" & Year(txtBordxDateTo) & "," & Month(txtBordxDateTo) & "," & Day(txtBordxDateTo) & ") " & "AND " & _
                                    " {VIEW_MCO_LISTING.HLTCode} = '" & TmpHLT & "'" & "And " & _
                                    " {VIEW_MCO_LISTING.PayCode} = '" & tmpPayCode & "'" & " And " & _
                                    " {VIEW_MCO_LISTING.MBMBatchNo} = " & tmpBatchNo & ""
            cRpt.ReportTitle
            cRpt.ParameterFields(0) = "FromDate;" & DateFrom & ";true"
            cRpt.ParameterFields(1) = "LtDate;" & DateTo & ";true"
            cRpt.ParameterFields(2) = "PayorName;" & cboInsurer & ";true"
            cRpt.Action = 1
        End If
    End If
Else
    If cmdMCOPrint Then
        'Update MBMPrinting
        If optMCOPrint And lvwMCO.SelectedItem.SubItems(MCO_PRINTED) = "YES" Then
            MsgBox "This report already printed, please select 'Reprint' ", vbOKOnly + vbCritical, DialogBoxTitle
        Else
            If optMCOReprint And lvwMCO.SelectedItem.SubItems(MCO_PRINTED) = "NO" Then
                MsgBox "Please select 'Print' option for this report", vbOKOnly + vbCritical, DialogBoxTitle
            Else
                For ii = 1 To UBound(BDateArray)
                    UpdateMBMPrintingFile BDateArray(ii), BatchNoArray(ii), PaycodeArray(ii), HLTcodeArray(ii), "MCO"
                Next ii
                TmpHLT = lvwMCO.SelectedItem.SubItems(MCO_HLTCODE)
                tmpBatchNo = lvwMCO.SelectedItem.SubItems(MCO_BATCHNO)
                tmpPayCode = lvwMCO.SelectedItem
                DateFrom = Trim(txtBordxDateFr)
                DateTo = Trim(txtBordxDateTo)
            
                ss0 = "{VIEW_MCO_LISTING.HLTCode} = '" & TmpHLT & "'"
                ss = "{VIEW_MCO_LISTING.MBMBordxDate} = '" & DateFrom & "'"
                ss1 = "{VIEW_MCO_LISTING.MBMBordxDate} = '" & DateTo & "'"
                cRpt.ReportFileName = "C:\My Documents\RPT\MCOListing.rpt"
                cRpt.Destination = crptToPrinter
                cRpt.SelectionFormula = " {VIEW_MCO_LISTING.MBMBordxDate} >= date(" & Year(txtBordxDateFr) & "," & Month(txtBordxDateFr) & "," & Day(txtBordxDateFr) & ") " & " And " & _
                                        " {VIEW_MCO_LISTING.MBMBordxDate} <= date(" & Year(txtBordxDateTo) & "," & Month(txtBordxDateTo) & "," & Day(txtBordxDateTo) & ") " & "AND " & _
                                        " {VIEW_MCO_LISTING.HLTCode} = '" & TmpHLT & "'" & "And " & _
                                        " {VIEW_MCO_LISTING.PayCode} = '" & tmpPayCode & "'" & " And " & _
                                        " {VIEW_MCO_LISTING.MBMBatchNo} = " & tmpBatchNo & ""
                                        
                cRpt.ParameterFields(0) = "PayorName;" & cboInsurer & ";true"
                cRpt.ParameterFields(1) = "FromDate;" & DateFrom & ";true"
                cRpt.ParameterFields(2) = "LtDate;" & DateTo & ";true"
                cRpt.Action = 1
            End If
        End If
    End If
End If
End Sub
Private Sub cboReport_Click()
    If cboReport <> "" Then
        SSTab1.Tab = cboReport.ListIndex
    End If
   
End Sub


Private Sub cmdClear_Click()
    txtBordxDateFr = "__/__/____"
    txtBordxDateTo = "__/__/____"
    cboInsurer = ""
    txtBatchNo = ""

If cboReport.ListIndex = 2 Then
    cboHLTCode = ""
End If

End Sub

Private Sub cmdClose_Click()
    Unload Me
End Sub



Private Sub cmdCRDPrint_Click()
If ((txtBordxDateFr) = "__/__/____" Or (txtBordxDateTo) = "__/__/____" Or (cboInsurer) = "" Or _
    (txtBatchNo) = "" Or (cboHLTCode) = "") Then
    MsgBox "Please make sure all the search criteria is entered before print", vbOKOnly + vbCritical, DialogBoxTitle
Else
    
    If optCRDPrint.Value = True Then
        ExpMBMPrintingToExcel "NO"
    ElseIf optCRDReprint.Value = True Then
        ExpMBMPrintingToExcel "YES"
    End If
End If
End Sub

Private Sub cmdCRDPrintPreview_Click()
If ((txtBordxDateFr) = "__/__/____" Or (txtBordxDateTo) = "__/__/____" Or (cboInsurer) = "" Or _
    (txtBatchNo) = "" Or (cboHLTCode) = "") Then
    MsgBox "Please make sure all the search criteria is entered before print", vbOKOnly + vbCritical, DialogBoxTitle
Else
    
    If optCRDPrint.Value = True Then
        ExpMBMPrintingToExcel "NO"
    ElseIf optCRDReprint.Value = True Then
        ExpMBMPrintingToExcel "YES"
    End If
End If
End Sub

Private Sub cmdInvPrint_Click()
'On Error Resume Next
 If ((txtBordxDateFr) = "__/__/____" Or (txtBordxDateTo) = "__/__/____" Or (cboInsurer) = "" Or _
    (txtBatchNo) = "" Or (cboHLTCode) = "") Then
    MsgBox "Please make sure all the search criteria is entered before print", vbOKOnly + vbCritical, DialogBoxTitle
Else
 
    If optInvPrint.Value = True Then
        PrintInvoice "NO"
    ElseIf optInvReprint.Value = True Then
        PrintInvoice "YES"
    End If
End If
End Sub

Private Sub cmdINVPrintPreview_Click()
If optInvPrint.Value = True Then
        PrintInvoice "NO"
    ElseIf optInvReprint.Value = True Then
        PrintInvoice "YES"
    End If
End Sub



Private Sub cmdMCOPrint_Click()
If ((txtBordxDateFr) = "__/__/____" Or (txtBordxDateTo) = "__/__/____" Or (cboInsurer) = "" Or _
    (txtBatchNo) = "" Or (cboHLTCode) = "") Then
    MsgBox "Please make sure all the search criteria is entered before print", vbOKOnly + vbCritical, DialogBoxTitle
Else
    If optMCOPrint.Value = True Then
        MCOPrint "NO"
    ElseIf optMCOReprint.Value = True Then
        MCOPrint "YES"
    Else
        MCOPrint "ALL"
    End If
End If
End Sub

Private Sub cmdMCOPrintPreview_Click()
    If optMCOPrint.Value = True Then
        MCOPrint "NO"
    ElseIf optMCOReprint.Value = True Then
        MCOPrint "YES"
    Else
        MCOPrint "ALL"
    End If
End Sub

Private Sub CmdPrint_Click()
If ((txtBordxDateFr) = "__/__/____" Or (txtBordxDateTo) = "__/__/____" Or (cboInsurer) = "" Or _
    (txtBatchNo) = "" Or (cboHLTCode) = "") Then
    MsgBox "Please make sure all the search criteria is entered before print", vbOKOnly + vbCritical, DialogBoxTitle
Else
    If optPrint.Value = True Then
        LeafLetPrint "NO"
    ElseIf optReprint.Value = True Then
        LeafLetPrint "YES"
    ElseIf optPrintAll.Value = True Then
        LeafLetPrint "ALL"
    End If
End If
End Sub

Private Sub cmdSearch_Click()

    If txtBordxDateFr = "__/__/____" Then
        MsgBox "Please Enter Bordereaux Date From !", vbInformation
        Exit Sub
    End If
    If txtBordxDateTo = "__/__/____" Then
        MsgBox "Please Enter Bordereaux Date To !", vbInformation
        Exit Sub
    End If
    If cboInsurer.Text = "" Then
        MsgBox "Please Enter Insurer's Code !", vbInformation
        Exit Sub
    End If
    If txtBatchNo = "" Then
        MsgBox "Please Enter Batch No !", vbInformation
        Exit Sub
    End If
    If cboHLTCode = "" Then
        MsgBox "Please Enter Health Type !", vbInformation
        Exit Sub
    End If
    
    Select Case SSTab1.Tab
    Case 0  'Leaflet
        RecordSearching lvwLeaflet
    Case 1  'MCO Listing
        RecordSearching lvwMCO
    Case 2  'Invoice
        RecordSearching lvwInvoice
    Case 3  'Card Printing
        RecordSearching lvwCRD
    End Select
End Sub

Private Sub Form_Load()
    Call ComboListing
    Call HLTComboListing
    cboReport.Text = cboReport.list(0)
End Sub


Private Sub SSTab1_Click(PreviousTab As Integer)
    cboReport.ListIndex = SSTab1.Tab
End Sub

Private Sub PrintInvoice(PrintMode As String)
On Error Resume Next
Dim ii As Integer
Dim MemberNo As String
Dim rstInv As New ADODB.Recordset
Dim strsql As String

Dim tmpBDate, tmpBatchNo, tmpPayCode, tmpHLTCode As String
Dim tmpPrintedBDate, tmpPrintedBatchNo, tmpPrintedPaycode, tmpPrintedHLTcode As String
Dim BDateArray() As String
Dim BatchNoArray() As String
Dim PaycodeArray() As String
Dim HLTcodeArray() As String
Dim cnt As Integer
Dim date1 As String
Dim date2 As String
Dim TmpHLT As String
Dim RetmpBDateFr As String
Dim RetmpBDateTo As String

     strsql = "SELECT * from MBMPrinting"
     rstInv.Open strsql, medixmasterconn, adOpenKeyset, adLockOptimistic
    
    cnt = 0
    For ii = 1 To lvwInvoice.listitems.Count
        Set lvwInvoice.SelectedItem = lvwInvoice.listitems(ii)
        If lvwInvoice.SelectedItem.SubItems(INV_PRINTED) = PrintMode Then
           tmpBDate = lvwInvoice.SelectedItem.SubItems(INV_BORDXDATE)
           tmpHLTCode = lvwInvoice.SelectedItem.SubItems(INV_HLTCODE)
           tmpBatchNo = lvwInvoice.SelectedItem.SubItems(INV_BATCHNO)
           tmpPayCode = lvwInvoice.SelectedItem
                
            If tmpBDate = tmpPrintedBDate And tmpBatchNo = tmpPrintedBatchNo And tmpPayCode = tmpPrintedPaycode And tmpHLTCode = tmpPrintedHLTcode Then
                Else
                tmpPrintedBDate = tmpBDate
                tmpPrintedBatchNo = tmpBatchNo
                tmpPrintedPaycode = tmpPayCode
                tmpPrintedHLTcode = tmpHLTCode
                       
                cnt = cnt + 1
                ReDim Preserve BDateArray(cnt)
                ReDim Preserve BatchNoArray(cnt)
                ReDim Preserve PaycodeArray(cnt)
                ReDim Preserve HLTcodeArray(cnt)
                BDateArray(cnt) = tmpBDate
                BatchNoArray(cnt) = tmpBatchNo
                PaycodeArray(cnt) = tmpPayCode
                HLTcodeArray(cnt) = tmpHLTCode
            End If
        End If
    Next ii
    
    If cmdINVPrintPreview Then
        TmpHLT = lvwInvoice.SelectedItem.SubItems(INV_HLTCODE)
        tmpBatchNo = Trim(txtBatchNo)
        tmpPayCode = lvwInvoice.SelectedItem
        DateFrom = Trim(txtBordxDateFr)
        DateTo = Trim(txtBordxDateTo)
    
        ss0 = "{VIEW_Invoice_Report.HLTCode} = '" & TmpHLT & "'"
        ss = "{VIEW_Invoice_Report.MBMBordxDate} = '" & DateFrom & "'"
        ss1 = "{VIEW_Invoice_Report.MBMBordxDate} = '" & DateTo & "'"
        ss2 = "{VIEW_Invoice_Report.MBMBatchNo} = " & tmpBatchNo & ""
        ss3 = "{VIEW_Invoice_Report.PayCode} = '" & tmpPayCode & "'"
                
        cRpt.ReportFileName = "C:\My Documents\RPT\InvoicesReport.rpt"
        cRpt.Destination = crptToWindow
        cRpt.SelectionFormula = " {VIEW_Invoice_Report.MBMBordxDate} >= date(" & Year(txtBordxDateFr) & "," & Month(txtBordxDateFr) & "," & Day(txtBordxDateFr) & ") " & " AND " & _
                                " {VIEW_Invoice_Report.MBMBordxDate} <= date(" & Year(txtBordxDateTo) & "," & Month(txtBordxDateTo) & "," & Day(txtBordxDateTo) & ") " & "AND " & _
                                " {VIEW_Invoice_Report.HLTCode} = '" & TmpHLT & "'" & " AND " & _
                                " {VIEW_Invoice_Report.PayCode} = '" & tmpPayCode & "'" & " AND " & _
                                " {VIEW_Invoice_Report.MBMBatchNo} = " & tmpBatchNo & ""
                                
        cRpt.Action = 1
    Else
        If cmdInvPrint Then
            If optInvPrint And lvwInvoice.SelectedItem.SubItems(INV_PRINTED) = "YES" Then
                MsgBox "This report already printed, please select 'Reprint' ", vbOKOnly + vbCritical, DialogBoxTitle
            Else
            
            If optInvPrint And lvwInvoice.SelectedItem.SubItems(INV_PRINTED) = "NO" Then
                Call IncInvoiceNo
                    
                'Update MBMPrinting
                For ii = 1 To UBound(BDateArray)
                    UpdateMBMPrintingFile BDateArray(ii), BatchNoArray(ii), PaycodeArray(ii), HLTcodeArray(ii), "INV"
                Next ii
                
                TmpHLT = lvwInvoice.SelectedItem.SubItems(INV_HLTCODE)
                tmpBatchNo = Trim(txtBatchNo)
                tmpPayCode = lvwInvoice.SelectedItem
                DateFrom = Trim(txtBordxDateFr)
                DateTo = Trim(txtBordxDateTo)
                 
                ss0 = "{VIEW_Invoice_Report.HLTCode} = '" & TmpHLT & "'"
                ss = "{VIEW_Invoice_Report.MBMBordxDate} = '" & DateFrom & "'"
                ss1 = "{VIEW_Invoice_Report.MBMBordxDate} = '" & DateTo & "'"
                ss2 = "{VIEW_Invoice_Report.MBMBatchNo} = " & tmpBatchNo & ""
                ss3 = "{VIEW_Invoice_Report.PayCode} = '" & tmpPayCode & "'"
                cRpt.ReportFileName = "C:\My Documents\RPT\InvoicesReport.rpt"
                cRpt.Destination = crptToPrinter
                cRpt.SelectionFormula = " {VIEW_Invoice_Report.MBMBordxDate} >= date(" & Year(txtBordxDateFr) & "," & Month(txtBordxDateFr) & "," & Day(txtBordxDateFr) & ") " & " AND " & _
                                        " {VIEW_Invoice_Report.MBMBordxDate} <= date(" & Year(txtBordxDateTo) & "," & Month(txtBordxDateTo) & "," & Day(txtBordxDateTo) & ") " & "AND " & _
                                        " {VIEW_Invoice_Report.HLTCode} = '" & TmpHLT & "'" & " AND " & _
                                        " {VIEW_Invoice_Report.MBMBatchNo} = " & tmpBatchNo & "" & " AND " & _
                                        " {VIEW_Invoice_Report.PayCode} = '" & tmpPayCode & "'"

                cRpt.ParameterFields(0) = "MCOInvoiceNo;" & txtInvoiceNo & ";true"
                cRpt.ParameterFields(1) = "MCOInvoiceYr;" & TxtInvoiceYear & ";true"
                cRpt.Action = 1
            Else
                     If optInvReprint And lvwInvoice.SelectedItem.SubItems(INV_PRINTED) = "NO" Then
                        MsgBox "Please select 'Print' option for this report", vbOKOnly + vbCritical, DialogBoxTitle
                     Else
                        
                        TmpHLT = lvwInvoice.SelectedItem.SubItems(INV_HLTCODE)
                        tmpBatchNo = lvwInvoice.SelectedItem.SubItems(INV_BATCHNO)
                        tmpPayCode = lvwInvoice.SelectedItem
                        DateFrom = Trim(txtBordxDateFr)
                        DateTo = Trim(txtBordxDateTo)
                    
                        ss0 = "{VIEW_Invoice_Report.HLTCode} = '" & TmpHLT & "'"
                        ss = "{VIEW_Invoice_Report.MBMBordxDate} = '" & DateFrom & "'"
                        ss1 = "{VIEW_Invoice_Report.MBMBordxDate} = '" & DateTo & "'"
                        cRpt.ReportFileName = "C:\My Documents\RPT\InvoicesReport.rpt"
                        cRpt.Destination = crptToPrinter
                        cRpt.SelectionFormula = " {VIEW_Invoice_Report.MBMBordxDate} >= date(" & Year(txtBordxDateFr) & "," & Month(txtBordxDateFr) & "," & Day(txtBordxDateFr) & ") " & " And " & _
                                                " {VIEW_Invoice_Report.MBMBordxDate} <= date(" & Year(txtBordxDateTo) & "," & Month(txtBordxDateTo) & "," & Day(txtBordxDateTo) & ") " & "AND " & _
                                                " {VIEW_Invoice_Report.HLTCode} = '" & TmpHLT & "'" & " AND " & _
                                                " {VIEW_Invoice_Report.PayCode} = '" & tmpPayCode & "'" & " And " & _
                                                " {VIEW_Invoice_Report.MBMBatchNo} = " & tmpBatchNo & ""
                        
                       
                        
                        rstInv.MoveFirst
                        Do Until rstInv.EOF
                        If rstInv.recordCount > 0 Then
                           If txtBordxDateFr = rstInv!MBMBordxDate Or txtBordxDateTo = rstInv!MBMBordxDate And _
                              TmpHLT = rstInv!HLTCode And tmpBatchNo = rstInv!MBMBatchNo And tmpPayCode = rstInv!PayCode Then
                            
                                txtInvoiceNo = rstInv!MBMInvoiceNo
                        
                            End If
                        rstInv.MoveNext
                        End If
                        Loop
                        rstInv.Close
                        Set rstInv = Nothing
                        
                        cRpt.ParameterFields(0) = "MCOInvoiceNo;" & txtInvoiceNo & ";true"
                        cRpt.Action = 1
                    End If
                End If
            End If
        End If
    End If
End Sub
