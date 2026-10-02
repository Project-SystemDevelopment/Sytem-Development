VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "tabctl32.ocx"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Begin VB.Form frmStorageEntry 
   Caption         =   "STORAGE"
   ClientHeight    =   7980
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   11880
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   7980
   ScaleWidth      =   11880
   Begin TabDlg.SSTab SSTab1 
      Height          =   6975
      Left            =   120
      TabIndex        =   42
      Top             =   240
      Width           =   11700
      _ExtentX        =   20638
      _ExtentY        =   12303
      _Version        =   393216
      Tabs            =   2
      TabsPerRow      =   2
      TabHeight       =   520
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      TabCaption(0)   =   "Carton Details"
      TabPicture(0)   =   "frmStorageEntry.frx":0000
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Lstcrt"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "FrameCarton"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "chkAll"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).ControlCount=   3
      TabCaption(1)   =   "Storage Details"
      TabPicture(1)   =   "frmStorageEntry.frx":001C
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "FrameuPDATE"
      Tab(1).Control(1)=   "Frame2"
      Tab(1).Control(2)=   "Lststr"
      Tab(1).ControlCount=   3
      Begin VB.CheckBox chkAll 
         Caption         =   "Use Scan"
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
         Left            =   420
         TabIndex        =   76
         Top             =   2820
         Visible         =   0   'False
         Width           =   1095
      End
      Begin VB.Frame FrameuPDATE 
         Enabled         =   0   'False
         Height          =   1215
         Left            =   -74880
         TabIndex        =   62
         Top             =   5640
         Width           =   11415
         Begin VB.TextBox TxtVersion 
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
            Left            =   3720
            MaxLength       =   20
            TabIndex        =   22
            Top             =   240
            Width           =   1095
         End
         Begin VB.TextBox TxtCase1 
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
            TabIndex        =   21
            Top             =   240
            Width           =   1575
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
            Height          =   735
            Left            =   8880
            MaxLength       =   3000
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   29
            Top             =   240
            Width           =   2415
         End
         Begin MSMask.MaskEdBox txtDORqst 
            Height          =   285
            Left            =   6600
            TabIndex        =   26
            Top             =   240
            Width           =   1335
            _ExtentX        =   2355
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
            Left            =   6600
            TabIndex        =   27
            Top             =   480
            Width           =   1335
            _ExtentX        =   2355
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
            Left            =   6600
            TabIndex        =   28
            Top             =   720
            Width           =   1335
            _ExtentX        =   2355
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
            ItemData        =   "frmStorageEntry.frx":0038
            Left            =   3720
            List            =   "frmStorageEntry.frx":003A
            TabIndex        =   24
            Text            =   "I - In"
            Top             =   480
            Width           =   1095
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
            TabIndex        =   23
            Top             =   480
            Width           =   1575
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
            Left            =   1200
            Sorted          =   -1  'True
            TabIndex        =   25
            Top             =   765
            Width           =   3615
         End
         Begin VB.Label Label29 
            Caption         =   "Version"
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
            TabIndex        =   71
            Top             =   240
            Width           =   855
         End
         Begin VB.Label Label6 
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
            TabIndex        =   70
            Top             =   255
            Width           =   990
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
            Left            =   240
            TabIndex        =   69
            Top             =   840
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
            TabIndex        =   68
            Top             =   520
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
            TabIndex        =   67
            Top             =   525
            Width           =   975
         End
         Begin VB.Label Label25 
            Caption         =   "Date Of Request"
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
            Left            =   5160
            TabIndex        =   66
            Top             =   255
            Width           =   1335
         End
         Begin VB.Label Label20 
            Caption         =   "Date Of Delivery"
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
            Left            =   5160
            TabIndex        =   65
            Top             =   525
            Width           =   1335
         End
         Begin VB.Label Label17 
            Caption         =   "Date Of Return"
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
            Left            =   5160
            TabIndex        =   64
            Top             =   795
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
            Left            =   8160
            TabIndex        =   63
            Top             =   240
            Width           =   735
         End
      End
      Begin VB.Frame FrameCarton 
         Height          =   2355
         Left            =   360
         TabIndex        =   52
         Top             =   420
         Width           =   11055
         Begin VB.CheckBox chkscan1 
            Caption         =   "Use Scan"
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
            Left            =   5460
            TabIndex        =   75
            Top             =   360
            Value           =   1  'Checked
            Width           =   1095
         End
         Begin VB.TextBox TxtcrtRemarks 
            Height          =   855
            Left            =   1800
            MaxLength       =   3998
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   8
            Top             =   1380
            Width           =   8535
         End
         Begin MSMask.MaskEdBox MaskcrtDORqst 
            Height          =   285
            Left            =   8640
            TabIndex        =   4
            Top             =   300
            Width           =   1695
            _ExtentX        =   2990
            _ExtentY        =   503
            _Version        =   393216
            MaxLength       =   10
            Mask            =   "##/##/####"
            PromptChar      =   "_"
         End
         Begin MSMask.MaskEdBox MaskcrtDODelivery 
            Height          =   285
            Left            =   8640
            TabIndex        =   5
            Top             =   540
            Width           =   1695
            _ExtentX        =   2990
            _ExtentY        =   503
            _Version        =   393216
            MaxLength       =   10
            Mask            =   "##/##/####"
            PromptChar      =   "_"
         End
         Begin MSMask.MaskEdBox MaskcrtDOReturn 
            Height          =   285
            Left            =   8640
            TabIndex        =   6
            Top             =   780
            Width           =   1695
            _ExtentX        =   2990
            _ExtentY        =   503
            _Version        =   393216
            MaxLength       =   10
            Mask            =   "##/##/####"
            PromptChar      =   "_"
         End
         Begin MSMask.MaskEdBox MaskcrtDODestroy 
            Height          =   315
            Left            =   8640
            TabIndex        =   7
            Top             =   1020
            Width           =   1695
            _ExtentX        =   2990
            _ExtentY        =   556
            _Version        =   393216
            MaxLength       =   10
            Mask            =   "##/##/####"
            PromptChar      =   "_"
         End
         Begin VB.TextBox TxtcrtCartonNo 
            Height          =   315
            Left            =   1800
            MaxLength       =   20
            TabIndex        =   0
            Top             =   360
            Width           =   3495
         End
         Begin VB.TextBox TxtcrtCartonNo2 
            Height          =   315
            Left            =   3720
            MaxLength       =   20
            TabIndex        =   1
            Top             =   360
            Width           =   1575
         End
         Begin VB.ComboBox CboCartonStatus 
            Height          =   315
            ItemData        =   "frmStorageEntry.frx":003C
            Left            =   1800
            List            =   "frmStorageEntry.frx":003E
            TabIndex        =   2
            Text            =   "M - Medix"
            Top             =   640
            Width           =   3495
         End
         Begin MSMask.MaskEdBox MaskcrtDOSend 
            Height          =   285
            Left            =   1800
            TabIndex        =   3
            Top             =   915
            Width           =   3495
            _ExtentX        =   6165
            _ExtentY        =   503
            _Version        =   393216
            MaxLength       =   10
            Mask            =   "##/##/####"
            PromptChar      =   "_"
         End
         Begin VB.Label Label10 
            Caption         =   "Date Of Sending"
            Height          =   255
            Left            =   480
            TabIndex        =   61
            Top             =   975
            Width           =   1455
         End
         Begin VB.Label Label9 
            Caption         =   "Date Of Destroy"
            Height          =   255
            Left            =   7080
            TabIndex        =   60
            Top             =   1020
            Width           =   1215
         End
         Begin VB.Label Label24 
            Caption         =   "Remarks"
            Height          =   255
            Left            =   480
            TabIndex        =   59
            Top             =   1380
            Width           =   855
         End
         Begin VB.Label Label23 
            Caption         =   "Date Of Return"
            Height          =   255
            Left            =   7080
            TabIndex        =   58
            Top             =   780
            Width           =   1335
         End
         Begin VB.Label Label22 
            Caption         =   "Carton Status"
            Height          =   255
            Left            =   480
            TabIndex        =   57
            Top             =   675
            Width           =   975
         End
         Begin VB.Label Label21 
            Caption         =   "Date Of Delivery"
            Height          =   255
            Left            =   7080
            TabIndex        =   56
            Top             =   540
            Width           =   1455
         End
         Begin VB.Label Label19 
            Caption         =   "Date Of Request"
            Height          =   255
            Left            =   7080
            TabIndex        =   55
            Top             =   300
            Width           =   1455
         End
         Begin VB.Label Label18 
            Caption         =   "Carton No"
            Height          =   255
            Left            =   480
            TabIndex        =   54
            Top             =   360
            Width           =   735
         End
         Begin VB.Label Label16 
            Caption         =   "To"
            Height          =   255
            Left            =   3450
            TabIndex        =   53
            Top             =   360
            Width           =   375
         End
      End
      Begin VB.Frame Frame2 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1335
         Left            =   -74880
         TabIndex        =   43
         Top             =   360
         Width           =   11415
         Begin VB.CheckBox ChkScan 
            Caption         =   "Use Scan"
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
            TabIndex        =   13
            Top             =   360
            Value           =   1  'Checked
            Width           =   1095
         End
         Begin MSMask.MaskEdBox MaskstrDORqst 
            Height          =   285
            Left            =   8640
            TabIndex        =   17
            Top             =   360
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
            Left            =   8640
            TabIndex        =   18
            Top             =   600
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
            Left            =   8640
            TabIndex        =   19
            Top             =   840
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
         Begin VB.TextBox TxtstrCasNo2 
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
            Left            =   3960
            MaxLength       =   20
            TabIndex        =   12
            Top             =   360
            Width           =   1575
         End
         Begin VB.TextBox TxtstrCasNo 
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
            MaxLength       =   20
            TabIndex        =   10
            Top             =   360
            Width           =   1575
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
            ItemData        =   "frmStorageEntry.frx":0040
            Left            =   3960
            List            =   "frmStorageEntry.frx":0042
            TabIndex        =   15
            Text            =   "A - All"
            Top             =   600
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
            Left            =   1320
            TabIndex        =   14
            Top             =   600
            Width           =   1575
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
            Left            =   1320
            Sorted          =   -1  'True
            TabIndex        =   16
            Top             =   880
            Width           =   4215
         End
         Begin VB.Label Label7 
            Caption         =   "Date Of Return"
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
            Left            =   7080
            TabIndex        =   51
            Top             =   915
            Width           =   1455
         End
         Begin VB.Label Label5 
            Caption         =   "Date Of Delivery"
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
            Left            =   7080
            TabIndex        =   50
            Top             =   645
            Width           =   1575
         End
         Begin VB.Label Label3 
            Caption         =   "Date Of Request"
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
            Left            =   7080
            TabIndex        =   49
            Top             =   375
            Width           =   1575
         End
         Begin VB.Label Label2 
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
            TabIndex        =   48
            Top             =   645
            Width           =   975
         End
         Begin VB.Label Label1 
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
            TabIndex        =   47
            Top             =   375
            Width           =   990
         End
         Begin VB.Label Label11 
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
            Left            =   3000
            TabIndex        =   46
            Top             =   660
            Width           =   1215
         End
         Begin VB.Label Label4 
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
            Left            =   240
            TabIndex        =   45
            Top             =   960
            Width           =   975
         End
         Begin VB.Label Label15 
            Caption         =   "To"
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
            Left            =   3120
            TabIndex        =   44
            Top             =   360
            Width           =   375
         End
      End
      Begin MSComctlLib.ListView Lstcrt 
         Height          =   3795
         Left            =   360
         TabIndex        =   9
         Top             =   3120
         Width           =   11055
         _ExtentX        =   19500
         _ExtentY        =   6694
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
         NumItems        =   9
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "CartonNo"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Status"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   2
            SubItemIndex    =   2
            Text            =   "Date Register"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   3
            Text            =   "Date Send"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   4
            Text            =   "Date Request"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   5
            Text            =   "Date Delivery"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   6
            Text            =   "Date Return"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   7
            Text            =   "Date Destroy"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(9) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   8
            Text            =   "Remarks"
            Object.Width           =   2540
         EndProperty
      End
      Begin MSComctlLib.ListView Lststr 
         Height          =   3855
         Left            =   -74880
         TabIndex        =   20
         Top             =   1800
         Width           =   11415
         _ExtentX        =   20135
         _ExtentY        =   6800
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
         NumItems        =   13
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "CaseNo"
            Object.Width           =   2293
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Version"
            Object.Width           =   1411
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "CartonNo"
            Object.Width           =   2117
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   2
            SubItemIndex    =   3
            Text            =   "Carton Status"
            Object.Width           =   2117
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   4
            Text            =   "File Status"
            Object.Width           =   1764
         EndProperty
         BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   5
            Text            =   "Bill Amt"
            Object.Width           =   2117
         EndProperty
         BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   2
            SubItemIndex    =   6
            Text            =   "Date Register"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   7
            Text            =   "RequestBy"
            Object.Width           =   2293
         EndProperty
         BeginProperty ColumnHeader(9) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   8
            Text            =   "Date Request"
            Object.Width           =   2118
         EndProperty
         BeginProperty ColumnHeader(10) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   9
            Text            =   "Date Delivery"
            Object.Width           =   2118
         EndProperty
         BeginProperty ColumnHeader(11) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   10
            Text            =   "Date Return"
            Object.Width           =   2117
         EndProperty
         BeginProperty ColumnHeader(12) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   11
            Text            =   "Duration"
            Object.Width           =   1411
         EndProperty
         BeginProperty ColumnHeader(13) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   12
            Text            =   "Remarks"
            Object.Width           =   2540
         EndProperty
      End
   End
   Begin VB.Frame Frame5 
      Height          =   735
      Left            =   120
      TabIndex        =   11
      Top             =   7200
      Width           =   11715
      Begin VB.TextBox lblTotal 
         Height          =   285
         Left            =   1200
         TabIndex        =   79
         Top             =   240
         Width           =   495
      End
      Begin VB.TextBox lblCurrent 
         Height          =   285
         Left            =   120
         TabIndex        =   78
         Top             =   240
         Width           =   615
      End
      Begin VB.CommandButton cmdRemove 
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
         Height          =   350
         Left            =   6910
         TabIndex        =   77
         Top             =   240
         Width           =   810
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
         Left            =   8480
         TabIndex        =   33
         Top             =   240
         Width           =   615
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
         Left            =   9075
         TabIndex        =   34
         Top             =   240
         Width           =   615
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
         Left            =   11120
         TabIndex        =   38
         Top             =   240
         Width           =   530
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
         Left            =   6300
         TabIndex        =   37
         Top             =   240
         Width           =   615
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
         Left            =   7720
         TabIndex        =   32
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
         Left            =   5600
         TabIndex        =   31
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
         Left            =   4920
         TabIndex        =   30
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
         Left            =   10335
         TabIndex        =   36
         Top             =   240
         Width           =   800
      End
      Begin VB.CommandButton CmdListing 
         Caption         =   "&Listing"
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
         Left            =   9675
         TabIndex        =   35
         Top             =   240
         Width           =   675
      End
      Begin VB.CommandButton CmdDelete 
         Caption         =   "&Delete"
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
         Left            =   4920
         TabIndex        =   39
         Top             =   240
         Visible         =   0   'False
         Width           =   675
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
         Left            =   2880
         TabIndex        =   73
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
         Left            =   4440
         TabIndex        =   72
         Top             =   285
         Width           =   405
      End
      Begin VB.Label Label13 
         Caption         =   "of"
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
         Left            =   960
         TabIndex        =   41
         Top             =   285
         Width           =   255
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
         Left            =   1920
         TabIndex        =   40
         Top             =   285
         Width           =   735
      End
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
      TabIndex        =   74
      Top             =   0
      Width           =   11805
   End
End
Attribute VB_Name = "frmStorageEntry"
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
Dim CrtLstArr As Collection
Dim StoLstArr1 As Collection

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
 
    For i = 0 To CboCarton2.ListCount - 1
 
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

Private Sub CboCartonStatus_KeyPress(KeyAscii As Integer)
Dim NewText As String
Dim ValidCount As Integer
Dim ValidValue As String
Dim i As Integer
    If KeyAscii >= 32 And KeyAscii <> 127 Then
    NewText = LCase(Left(CboCartonStatus.Text, CboCartonStatus.SelStart) + Chr(KeyAscii) + Mid(CboCartonStatus.Text, CboCartonStatus.SelStart + CboCartonStatus.SelLength + 1))
 
    ValidCount = 0
    ValidValue = ""
 
    For i = 0 To CboCartonStatus.ListCount - 1
 
        If NewText = LCase(Left(CboCartonStatus.List(i), Len(NewText))) Then
                ValidCount = ValidCount + 1
                ValidValue = CboCartonStatus.List(i)
        End If
        Next
        If ValidCount <= 1 Then KeyAscii = 0
            If ValidCount = 1 Then
                CboCartonStatus.Text = ValidValue
                CboCartonStatus.SelStart = 0
                CboCartonStatus.SelLength = Len(ValidValue)
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
 
        If NewText = LCase(Left(CboUser1.List(i), Len(NewText))) Then
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

Private Sub ChkAll_Click()
Dim SS As Integer
Dim CheckCnt As Integer
    If chkAll.Value = 1 Then
            For SS = 1 To Lstcrt.ListItems.Count
                Lstcrt.ListItems(SS).Checked = True
                CheckCnt = CheckCnt + 1
            Next SS
            LblRecChk = CheckCnt
    Else
            For SS = 1 To Lstcrt.ListItems.Count
                Lstcrt.ListItems(SS).Checked = False
            Next SS
            LblRecChk = 0
    End If
End Sub

Private Sub ChkScan1_Click()
   If chkscan1.Value = 0 Then
        Label16.Visible = True
        TxtcrtCartonNo2.Visible = True
        TxtcrtCartonNo.Width = 1575
    Else
        Call EntriesEnable(True)
        Label16.Visible = False
        TxtcrtCartonNo2.Visible = False
        TxtcrtCartonNo.Width = 3495
        Lstcrt.ListItems.Clear
    End If
    'TxtcrtCartonNo.SetFocus
End Sub

Private Sub ChkScan_Click()
    If ChkScan.Value = 0 Then
        Label15.Visible = True
        TxtstrCasNo2.Visible = True
        ChkScan.Left = 5760
        ChkScan.Top = 380
    Else
        Call EntriesEnable(True)
        Label15.Visible = False
        TxtstrCasNo2.Visible = False
        Lststr.ListItems.Clear
        ChkScan.Left = 3000
        ChkScan.Top = 380
        
    End If
    TxtstrCasNo.SetFocus
End Sub

Private Sub CmdAdd_Click()
    'CmdEdit.Enabled = False
    If SSTab1.Tab = 0 Then
        Lstcrt.ListItems.Clear
        Call EntriesEnable(True)
        CmdSave.Enabled = True
        CmdSave.Caption = "&Save"
        CmdAdd.Enabled = True
        CmdDelete.Enabled = False
        Call ClearCartonInfo
        TxtcrtCartonNo.SetFocus
    Else
        If CheckCnt > 0 Then
            TxtCase1 = ""
            txtVersion = ""
            CboCarton2.Text = ""
            CboCartonStatus.Text = ""
            CboFileStatus2.Text = ""
            CboUser2.Text = ""
            txtDORqst = "__/__/____"
            txtDOD = "__/__/____"
            txtDOReturn = "__/__/____"
            TxtRemarks = ""
            FrameuPDATE.Enabled = True
            CboCarton2.SetFocus
            CmdSave.Enabled = True
            CmdSave.Caption = "&Save"
        Else
            MsgBox "Please checked at least 1 Claim ! ", vbOKOnly + vbCritical
        End If
    End If
End Sub

Private Sub CmdClear_Click()
    Call ClearForm(Me)
    Check = False
    EditFlag = False
    Call EntriesEnable(True)
    CmdDelete.Enabled = False
    Lstcrt.ListItems.Clear
    Lststr.ListItems.Clear
    txtTotalRecord = 0
    chkscan1.Enabled = True
    ChkScan.Value = 1
    chkscan1.Value = 1
    Label15.Visible = False
    TxtstrCasNo2.Visible = False
    Lststr.ListItems.Clear
    CboFileStatus1.ListIndex = 0
    ChkScan.Left = 3000
    ChkScan.Top = 380
    FrameuPDATE.Enabled = False
    ReDim CaseNoArray(0)
    chkAll.Visible = False
    LblRecChk = ""
    lblCurrent = ""
End Sub

Private Sub CmdClear_LostFocus()
    If SSTab1.Tab = 0 Then
        TxtcrtCartonNo.SetFocus
    Else
        TxtstrCasNo.SetFocus
    End If
End Sub

Private Sub cmdDelete_Click()
Dim rst2 As New ADODB.Recordset
Dim sql2 As String
Dim RecordSaved
    
Screen.MousePointer = vbHourglass

RecordSaved = MsgBox("Do you want to delete the record?", vbYesNo + vbExclamation)
If RecordSaved = vbYes Then
    sql2 = ""
    If SSTab1.Tab = 0 Then
        sql2 = "Delete From CARTON Where CrtCartonNo = '" & TxtcrtCartonNo & "'"
    Else
        sql2 = "Delete From STORAGE Where StrCaseNo = '" & TxtstrCasNo & "'"
    End If
    rst2.Open sql2, medixmasterconnWS, adOpenKeyset, adLockOptimistic
    MsgBox "Record Deleted... ", vbOKOnly + vbInformation, DialogBoxTitle
    CmdSave.Enabled = True
    Call ClearForm(Me)
    Call DisableCmdButton(Me)
    Call EntriesEnable(True)
    If SSTab1.Tab = 0 Then
        Call CartonListing
        TxtcrtCartonNo.SetFocus
    Else
        Call StorageListing
        TxtstrCasNo.SetFocus
    End If
End If
Screen.MousePointer = vbDefault
End Sub

Private Sub CmdEdit_Click()
    If SSTab1.Tab = 0 Then
        TxtcrtCartonNo.Enabled = False
        TxtcrtCartonNo2.Enabled = False
        chkscan1.Enabled = False
    End If
    
    EditFlag = True
    CmdDelete.Enabled = False
    CmdSave.Enabled = True
    CmdSave.Caption = "&Update"
    If SSTab1.Tab = 0 Then
        If TxtcrtCartonNo <> "" Then
        Call EntriesEnable(True)
           TxtcrtCartonNo.Enabled = False
           CboCartonStatus.SetFocus
        End If
    End If
    FrameuPDATE.Enabled = False
    If SSTab1.Tab = 1 Then
        FrameuPDATE.Enabled = True
        CboCarton2.SetFocus
    End If
End Sub

Private Sub cmdExit_Click()
    Unload Me
End Sub
Private Function StorageSearchRecord()
Dim strsql As String
Dim strAppend As String
Dim sql As String
Dim storage As New ADODB.Recordset

    strAppend = ""
    
    If ChkScan.Value = 1 Then
        If Len(Trim(TxtstrCasNo)) Then
            strAppend = strAppend + " AND CASNumber = '" & RemStr(Trim(TxtstrCasNo)) & "'"
        End If
    Else
        If Len(Trim(TxtstrCasNo)) Then
            strAppend = strAppend + " AND CASNumber >= '" & RemStr(Trim(TxtstrCasNo)) & "'"
        End If
        If Len(Trim(TxtstrCasNo2)) Then
            strAppend = strAppend + " AND CASNumber <= '" & RemStr(Trim(TxtstrCasNo2)) & "'"
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
    
    total = 0
    Current = 0
    If ChkScan.Value = 0 Then
        Lststr.ListItems.Clear
    End If

    sql = " Select * from View_Storage_CLm where 0 = 0 "
        
    If Len(Trim(strAppend)) Then
        sql = sql + strAppend
    End If
    
    
    rst2.Open sql, medixmasterconnWS, adOpenKeyset, adLockOptimistic
    
    If rst2.recordCount = 0 Then
        MsgBox "No record found!", vbOKOnly + vbCritical, DialogBoxTitle
        CmdAdd.Enabled = True
    Else
    
    Do
        Do Until rst2.EOF
        total = rst2.recordCount
        lblTotal = total
            With rst2
                Set Record = Lststr.ListItems.Add(, , rst2!CasNumber)
                
                If Len(!claimversion) Then
                    Record.SubItems(1) = !claimversion
                End If
                If Len(!STRCartonNo) Then
                    Record.SubItems(2) = !STRCartonNo
                End If
                If Len(!CRTCartonStatus) Then
                    Record.SubItems(3) = !CRTCartonStatus
                End If
                If Len(!STRFileStatus) Then
                    Record.SubItems(4) = !STRFileStatus
                End If
                If Len(!CLMTotalActual) Then
                    Record.SubItems(5) = Format(!CLMTotalActual, "#,##,0.00")
                End If
                If Len(!LastUpdateDate) Then
                    Record.SubItems(6) = Format(!LastUpdateDate, "dd/mm/yyyy")
                End If
                If Len(!STRRequestBy) Then
                    Record.SubItems(7) = !STRRequestBy
                End If
                If Len(!DORequest) Then
                    Record.SubItems(8) = Format(!DORequest, "dd/mm/yyyy")
                End If
                If Len(!DODelivery) Then
                    Record.SubItems(9) = Format(!DODelivery, "dd/mm/yyyy")
                End If
                If Len(!DOReturn) Then
                    Record.SubItems(10) = Format(!DOReturn, "dd/mm/yyyy")
                End If
                If Len(!DOReturn) And Len(!DODelivery) Then
                    Record.SubItems(11) = DateValue(!DOReturn) - DateValue(!DODelivery)
                End If
                If Len(!STRRemarks) Then
                    Record.SubItems(12) = !STRRemarks
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

Private Sub CmdListing_Click()
    CmdSearch.Enabled = False
  '  medixmasterconnWS.Open "File Name=" & BOSPath & "Conn\mediconnectWS.UDL"

If SSTab1.Tab = 1 Then

    If Trim(CboCarton1) = "" Then
        MsgBox "Please select Carton No. ", vbOKOnly + vbCritical, DialogBoxTitle
        CboCarton1.SetFocus
    Else
        Dim objExcel As Excel.Application
        Dim objWorkbook As Excel.Workbook
        Dim objWorksheet As Excel.WORKSHEET
        Dim AA, BB As Integer
        Dim yy, Cnt As Integer
        Dim strsql As String
        Dim Sto As New ADODB.Recordset
        Dim ECol1 As String
        Dim ECol2 As String
        Dim ECol3 As String
           
        Screen.MousePointer = vbHourglass

        If objExcel Is Nothing Then
            Set objExcel = CreateObject("Excel.application")
                objExcel.Interactive = False
            If objExcel.Visible = False Then
                objExcel.Visible = True
            End If
            objExcel.WindowState = xlMaximized
            Set objWorkbook = objExcel.Workbooks.Add(LocCardPath & "Storage\StorageTracking.xlt")
            Set objWorksheet = objWorkbook.Sheets(1)
            objExcel.DisplayAlerts = False
            Set objWorksheet = objWorkbook.Worksheets.Item(1)
        End If
        
        If Err.Number > 0 Then
            MsgBox "Microsoft Excel is Not loaded On this machine.", vbOKOnly + vbCritical, "Error Loading Excel"
            GoTo ExitFunction
        End If
        
        On Error GoTo HANDLE_ERROR
    
        strsql = "select STRCaseNo,STRClmVersion, STRFileStatus from Storage WHERE STRCartonNO  = '" & CboCarton1 & "'"
        Sto.Open strsql, medixmasterconnWS, adOpenKeyset, adLockOptimistic
        'Set objExcel = New excel.Application
        
        If Sto.recordCount <> 0 Then
            RecordSelected = Sto.recordCount
            objWorksheet.Cells(2, "F") = Format(Date, "MM/DD/YYYY")
            objWorksheet.Cells(2, "H") = CboCarton1
            BB = 0
            yy = 0
            Cnt = 0
            ECol1 = Chr(65)
            ECol2 = Chr(66)
            ECol3 = Chr(67)
            
            Do Until Sto.EOF
                BB = BB + 1
                If BB > 50 Then
                    yy = yy + 4
                    BB = 1
                    ECol1 = Chr(65 + yy)
                    ECol2 = Chr(66 + yy)
                    ECol3 = Chr(67 + yy)
                    objWorksheet.Cells(4, ECol1) = "No"
                    objWorksheet.Cells(4, ECol2) = "Case No"
                    objWorksheet.Cells(4, ECol3) = "Checking"
                End If
                
                Cnt = Cnt + 1
                objWorksheet.Cells(4 + CDbl(BB), ECol1) = Cnt
                objWorksheet.Cells(4 + CDbl(BB), ECol2) = Sto!STRCaseNo & "-" & Sto!StrCLMVersion
                objWorksheet.Cells(4 + CDbl(BB), ECol3) = Sto!STRFileStatus
                
                Sto.MoveNext
            
        Loop
            
        End If
            objExcel.DisplayAlerts = True
            objExcel.Interactive = True
            Screen.MousePointer = vbDefault
           ' medixmasterconnWS.Close
           ' Set medixmasterconnWS = Nothing
            CmdSearch.Enabled = True
            Exit Sub
ExitFunction:
            objExcel.DisplayAlerts = True
            objExcel.Interactive = True
            objExcel.Quit
            'medixmasterconnWS.Close
           ' Set medixmasterconnWS = Nothing
            CmdSearch.Enabled = True
            Screen.MousePointer = vbDefault
            Exit Sub
HANDLE_ERROR:
            MsgBox "Export To Excel failed. Encountered the following Error" & vbCrLf & vbCrLf & _
            Err.Number & ": " & Err.Description, vbOKOnly + vbCritical, "Error Exporting To Excel"
            Set objWorksheet = Nothing
            Set objWorkbook = Nothing
            GoTo ExitFunction
    End If
End If

End Sub


Private Sub CmdPrintFile_Click()
Select Case SSTab1.Tab
    Case 0
        Screen.MousePointer = vbHourglass
        
        If Lstcrt.ListItems.Count <> 0 Then
            Call ExportListViewtoExcel(Lstcrt)
        Else
            MsgBox "Report cannot be printed without any record.", vbInformation
        End If
        
        Screen.MousePointer = vbDefault
        'Call ExportToExcelWorksheet(Lstcrt)
    Case 1
         'Call ExportToExcelWorksheet(Lststr)
        Screen.MousePointer = vbHourglass
        
        If Lststr.ListItems.Count <> 0 Then
            Call ExportListViewtoExcel(Lststr)
        Else
            MsgBox "Report cannot be printed without any record.", vbInformation
        End If
        
        Screen.MousePointer = vbDefault
End Select
End Sub

Private Sub CmdRemove_Click()
Dim RecordSaved
Dim CC As Integer
    Screen.MousePointer = vbHourglass
    'medixmasterconnWS.Open "File Name=" & BOSPath & "Conn\mediconnectWS.UDL"

    RecordSaved = MsgBox("Do you want to save the record?", vbYesNo + vbExclamation)
    If RecordSaved = vbYes Then
        If Trim(TxtCase1) <> "" And Trim(txtVersion) <> "" Then
            Call DeleteRecords
        Else
            MsgBox "Please Select record to delete!", vbCritical + vbOKOnly, DialogBoxTitle
        End If
    End If
   ' medixmasterconnWS.Close
    'Set medixmasterconnWS = Nothing
    Screen.MousePointer = vbDefault
End Sub

Private Sub DeleteRecords()
Dim strsql As String
Dim strsqlCLM As String
Dim tmpSql As String
Dim rst As New ADODB.Recordset
    strsql = " Delete from STORAGE Where STRCaseNo = '" & Trim(TxtCase1) & "' AND STRClmVersion = '" & Trim(txtVersion) & "'"
    rst.Open strsql, medixmasterconnWS, adOpenKeyset, adLockOptimistic
    
    tmpSql = " And CASNumber = '" & Trim(TxtCase1) & "' AND ClaimVersion = '" & Trim(txtVersion) & "'"
    
    Lststr.ListItems.Clear
    
    Call StorageListing(tmpSql)
    
    MsgBox "Record has been deleted!", vbCritical + vbOKOnly, DialogBoxTitle
End Sub

Private Sub CmdSave_Click()
Dim AA, SS, a, B, CASVer As Integer
Dim RecordSaved
Dim CASNo As String
Dim strAppend As String
Dim rst2 As New ADODB.Recordset
    
    If SSTab1.Tab = 0 Then
            If Trim(Mid(CboCartonStatus, 1, 1)) = "" Then
                    MsgBox "Please Enter Carton Status", vbOKOnly + vbCritical, DialogBoxTitle
                    CboCartonStatus.SetFocus
                    Exit Sub
            ElseIf Format(MaskcrtDORqst, "YYYYMMDD") < Format(MaskcrtDOSend, "YYYYMMDD") Then
                    MsgBox "Invalid Date ! Date of Request must greater than Date of Send ", vbOKOnly + vbCritical, DialogBoxTitle
                    MaskcrtDORqst = "__/__/____"
                    MaskcrtDORqst.SetFocus
                    Exit Sub
            ElseIf IsDate(MaskcrtDOReturn) < IsDate(MaskcrtDODelivery) Then
                    MsgBox "Invalid Date ! Date of Return must greater than Date of Delivery ", vbOKOnly + vbCritical, DialogBoxTitle
                    MaskcrtDOReturn = "__/__/____"
                    MaskcrtDOReturn.SetFocus
                    Exit Sub
            ElseIf IsDate(MaskcrtDODestroy) < IsDate(MaskcrtDOReturn) Then
                    MsgBox "Invalid Date ! Date of Destroy must greater than Date of Return ", vbOKOnly + vbCritical, DialogBoxTitle
                    MaskcrtDODestroy = "__/__/____"
                    MaskcrtDODestroy.SetFocus
                    Exit Sub
            ElseIf Trim(Mid(CboCartonStatus, 1, 1)) = "R" And Trim(MaskcrtDOReturn) = "__/__/____" Then
                    MsgBox "Date Return cannot be Empty when the status is 'R' !", vbOKOnly + vbCritical, DialogBoxTitle
                    MaskcrtDOReturn.SetFocus
                    Exit Sub
            Else
                Screen.MousePointer = vbHourglass
                If EditFlag = True Then
                    RecordSaved = MsgBox("Do you want to save the record?", vbYesNo + vbExclamation)
                    If RecordSaved = vbYes Then
                        'medixmasterconnWS.Open "File Name=" & BOSPath & "Conn\mediconnectWS.UDL"
                        Set CrtLstArr = New Collection
                        For SS = 1 To Lstcrt.ListItems.Count
                            If Lstcrt.ListItems(SS).Checked = True Then
                                Call SaveRecByCarton(Trim(Lstcrt.ListItems(SS).Text))
                                CrtLstArr.Add (Trim(Lstcrt.ListItems(SS).Text))
                            End If
                        Next SS
                        If EditFlag = True Then
                            MsgBox "Record Updated...", vbOKOnly + vbInformation, DialogBoxTitle
                        End If
                        Call ClearForm(Me)
                        Lstcrt.ListItems.Clear
                
                        For a = 1 To CrtLstArr.Count
                            strAppend = " AND CRTCartonNo = '" & CrtLstArr.Item(a) & "'"
                            Call CartonListing(strAppend)
                        Next a
                        
                        sql = " Select crtcartonno from carton where 0 = 0 "
                        rst2.Open sql, medixmasterconnWS, adOpenKeyset, adLockOptimistic
    
                        total = rst2.recordCount
                        lblTotal = total
        
                        rst2.Close
                        Set rst2 = Nothing
                        LblRecChk = 0
                        
                    End If
                Else:
                   ' medixmasterconnWS.Open "File Name=" & BOSPath & "Conn\mediconnectWS.UDL"
                    Call SaveCarton
                End If
                'medixmasterconnWS.Close
            'Set medixmasterconnWS = Nothing
            End If
' ===================Tab 2 Storage Details=============================================
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
        Set StoLstArr1 = New Collection
      '  medixmasterconnWS.Open "File Name=" & BOSPath & "Conn\mediconnectWS.UDL"
            If EditFlag = True Then
               Erase CaseNoArray
               ReDim Preserve CaseNoArray(1)
               CaseNoArray(1) = Lststr.SelectedItem.Text & "-" & Lststr.SelectedItem.SubItems(1) & "~"
        
            End If
            'For AA = 0 To UBound(CaseNoArray)
            For AA = 1 To UBound(CaseNoArray)
                If Trim(CaseNoArray(AA)) <> "" Then
                    delimiter1 = InStr(1, CaseNoArray(AA), "-")
                    CASNo = Trim(Mid(CaseNoArray(AA), 1, IIf(delimiter1 = 0, 10, delimiter1 - 1)))
                    delimiter2 = InStr(1, CaseNoArray(AA), "~")
                    CASVer = Trim(Mid(CaseNoArray(AA), delimiter1 + 1, delimiter2 - delimiter1 - 1))
                    Call SaveStore(CASNo, CInt(CASVer))
                    StoLstArr1.Add CASNo
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
            Call ClearStoreInfo
            If ChkScan.Value = 1 Then
                TxtstrCasNo = Lststr.SelectedItem.Text
            End If
            Lststr.ListItems.Clear
            
            For a = 1 To StoLstArr1.Count
                    strAppend = strAppend + " AND CASNumber = '" & StoLstArr1.Item(a) & "'"
                    Call StorageListing(strAppend)
            Next a
            
            'sql = " Select * from View_Storage_CLm where 0 = 0 "
            'rst2.Open sql, medixmasterconnWS, adOpenKeyset, adLockOptimistic

            'total = rst2.recordCount
            'lblTotal = total

            'rst2.Close
            'Set rst2 = Nothing
            LblRecChk = 0
            
         '   medixmasterconnWS.Close
         '   Set medixmasterconnWS = Nothing
            'Call StorageSearchRecord
        End If
    End If
    
    CmdEdit.Enabled = False
    CmdSave.Enabled = True
    Screen.MousePointer = vbDefault
End Sub

Private Sub SaveCarton()
Dim RecordSaved
Dim CARTON As New ADODB.Recordset
Dim SaveCarton As New ADODB.Recordset
Dim strsql As String
Dim strAppend As String
Dim tmploop As Integer
Dim tmploopNo As Integer
    tmploop = 0
    tmploopNo = 0
    If TxtcrtCartonNo = "" Then
        MsgBox "Please Enter Carton Number", vbOKOnly + vbCritical, DialogBoxTitle
        TxtcrtCartonNo.SetFocus
    ElseIf Trim(Mid(CboCartonStatus, 1, 1)) = "" Then
        MsgBox "Please Enter Carton Status", vbOKOnly + vbCritical, DialogBoxTitle
        CboCartonStatus.SetFocus
    ElseIf Format(MaskcrtDORqst, "YYYYMMDD") < Format(MaskcrtDOSend, "YYYYMMDD") Then
        MsgBox "Invalid Date ! Date of Request must greater than Date of Send ", vbOKOnly + vbCritical, DialogBoxTitle
        MaskcrtDORqst = "__/__/____"
        MaskcrtDORqst.SetFocus
    ElseIf IsDate(MaskcrtDOReturn) < IsDate(MaskcrtDODelivery) Then
        MsgBox "Invalid Date ! Date of Return must greater than Date of Delivery ", vbOKOnly + vbCritical, DialogBoxTitle
        MaskcrtDOReturn = "__/__/____"
        MaskcrtDOReturn.SetFocus
    ElseIf IsDate(MaskcrtDODestroy) < IsDate(MaskcrtDOReturn) Then
        MsgBox "Invalid Date ! Date of Destroy must greater than Date of Return ", vbOKOnly + vbCritical, DialogBoxTitle
        MaskcrtDODestroy = "__/__/____"
        MaskcrtDODestroy.SetFocus
    ElseIf Trim(Mid(CboCartonStatus, 1, 1)) = "R" And Trim(MaskcrtDOReturn) = "__/__/____" Then
        MsgBox "Date Return cannot be null when the status is 'R' !", vbOKOnly + vbCritical, DialogBoxTitle
        MaskcrtDOReturn.SetFocus
    Else
        strsql = "Select * From CARTON Where CRTCartonNo = '" & RemStr(TxtcrtCartonNo) & "'"
        CARTON.Open strsql, medixmasterconnWS, adOpenKeyset, adLockOptimistic
    
        If CARTON.recordCount > 0 And EditFlag = False Then
            MsgBox "Record Duplicated! Please Change The Corresponding Field", vbCritical, "Error"
            CARTON.Close
            TxtcrtCartonNo.SetFocus
            TxtcrtCartonNo.SelStart = 0
            TxtcrtCartonNo.SelLength = Len(TxtcrtCartonNo)
            Set CARTON = Nothing
            Exit Sub
        Else
            RecordSaved = MsgBox("Do you want to save the record?", vbYesNo + vbExclamation)
            If RecordSaved = vbYes Then
            
                If EditFlag = False Then
                    CARTON.AddNew
                End If
                If TxtcrtCartonNo <> TxtcrtCartonNo2 Then
                    Do While tmploop < TxtcrtCartonNo2
                        tmploop = TxtcrtCartonNo2 + tmploop
                    With CARTON
                        !CRTCartonNo = ChkStr(TxtcrtCartonNo)
                        !CRTCartonStatus = Trim(Mid(CboCartonStatus, 1, 1))
                        If MaskcrtDOSend <> "__/__/____" Then
                            !DoSend = MaskcrtDOSend
                        Else
                            !DoSend = Null
                        End If
                        If MaskcrtDORqst <> "__/__/____" Then
                            !DORequest = MaskcrtDORqst
                        Else
                            !DORequest = Null
                        End If
                        If MaskcrtDODelivery <> "__/__/____" Then
                            !DODelivery = MaskcrtDODelivery
                        Else
                            !DODelivery = Null
                        End If
                        If MaskcrtDOReturn <> "__/__/____" Then
                            !DOReturn = MaskcrtDOReturn
                        Else
                            !DOReturn = Null
                        End If
                        If MaskcrtDODestroy <> "__/__/____" Then
                            !dodestroy = MaskcrtDODestroy
                        Else
                            !dodestroy = Null
                        End If
                        !CRTRemarks = ChkStr(TxtcrtRemarks)
                        !LastUpdateBy = Logon
                        !LastUpdateDate = Date
                    End With
                    CARTON.Update
                    Loop
                Else
                    With CARTON
                        !CRTCartonNo = ChkStr(TxtcrtCartonNo)
                        !CRTCartonStatus = Trim(Mid(CboCartonStatus, 1, 1))
                        If MaskcrtDOSend <> "__/__/____" Then
                            !DoSend = MaskcrtDOSend
                        Else
                            !DoSend = Null
                        End If
                        If MaskcrtDORqst <> "__/__/____" Then
                            !DORequest = MaskcrtDORqst
                        Else
                            !DORequest = Null
                        End If
                        If MaskcrtDODelivery <> "__/__/____" Then
                            !DODelivery = MaskcrtDODelivery
                        Else
                            !DODelivery = Null
                        End If
                        If MaskcrtDOReturn <> "__/__/____" Then
                            !DOReturn = MaskcrtDOReturn
                        Else
                            !DOReturn = Null
                        End If
                        If MaskcrtDODestroy <> "__/__/____" Then
                            !dodestroy = MaskcrtDODestroy
                        Else
                            !dodestroy = Null
                        End If
                        !CRTRemarks = ChkStr(TxtcrtRemarks)
                        !LastUpdateBy = Logon
                        !LastUpdateDate = Date
                    End With
                    CARTON.Update
                End If
                If EditFlag = False Then
                    MsgBox "Record Saved...", vbOKOnly + vbInformation, DialogBoxTitle
                Else
                    MsgBox "Record Updated...", vbOKOnly + vbInformation, DialogBoxTitle
                End If
                CARTON.Close
                Set CARTON = Nothing
                strAppend = ""
                If Len(Trim(TxtcrtCartonNo)) Then
                    strAppend = " AND CRTCartonNo = '" & RemStr(Trim(TxtcrtCartonNo)) & "'"
                End If

                Call ClearForm(Me)
                Call CartonListing(strAppend)
            End If
        End If
        EditFlag = False
    End If
End Sub

Private Sub SaveRecByCarton(CartonNo As String)
Dim RecordSaved
Dim CARTON As New ADODB.Recordset
Dim SaveCarton As New ADODB.Recordset
Dim strsql As String
Dim sql As String
Dim strAppend As String
    
        strsql = "Select * From CARTON Where CRTCartonNo = '" & RemStr(CartonNo) & "'"
        CARTON.Open strsql, medixmasterconnWS, adOpenKeyset, adLockOptimistic
    
                With CARTON
                    !CRTCartonNo = ChkStr(CartonNo)
                    !CRTCartonStatus = Trim(Mid(CboCartonStatus, 1, 1))
                    If MaskcrtDOSend <> "__/__/____" Then
                        !DoSend = MaskcrtDOSend
                    Else
                        !DoSend = Null
                    End If
                    If MaskcrtDORqst <> "__/__/____" Then
                        !DORequest = MaskcrtDORqst
                    Else
                        !DORequest = Null
                    End If
                    If MaskcrtDODelivery <> "__/__/____" Then
                        !DODelivery = MaskcrtDODelivery
                    Else
                        !DODelivery = Null
                    End If
                    If MaskcrtDOReturn <> "__/__/____" Then
                        !DOReturn = MaskcrtDOReturn
                    Else
                        !DOReturn = Null
                    End If
                    If MaskcrtDODestroy <> "__/__/____" Then
                        !dodestroy = MaskcrtDODestroy
                    Else
                        !dodestroy = Null
                    End If
                    !CRTRemarks = ChkStr(TxtcrtRemarks)
                    !LastUpdateBy = Logon
                    !LastUpdateDate = Date
                End With
                CARTON.Update
                
        CARTON.Close
        Set CARTON = Nothing
        
End Sub

Private Sub CmdSearch_Click()
    LblRecChk = 0
    chkAll.Visible = True
    CmdAdd.Enabled = True
    CmdDelete.Enabled = False
    CmdSave.Enabled = False
    CmdSearch.Enabled = False
    Screen.MousePointer = vbHourglass
   ' medixmasterconnWS.Open "File Name=" & BOSPath & "Conn\mediconnectWS.UDL"

    Select Case SSTab1.Tab
        Case 0
            Call CartonSearchRecord
        Case 1
            Erase CaseNoArray
            ReDim CaseNoArray(0)
            If ChkScan.Value = 1 Then
                If TxtstrCasNo = "" Then
                    MsgBox "Please Enter Claim No", vbOKOnly + vbCritical, DialogBoxTitle
                    TxtstrCasNo.SetFocus
                Else
                    listrecord = Lststr.ListItems.Count
                    If listrecord <> 0 Then
                        
                        For listloop1 = 1 To listrecord
                            Call CaseNoCheck
                        Next listloop1
                    End If
                    If Check = True Then
                        MsgBox "Data has been uploaded!", vbOKOnly + vbCritical
                    ElseIf Check = False Then
                        Call StorageSearchRecord
                    End If
                    Check = False
                    Screen.MousePointer = Default
                    TxtstrCasNo = ""
                    TxtstrCasNo.SetFocus
                End If
                TxtstrCasNo.SetFocus
            Else
                Call StorageSearchRecord
            End If
    End Select
 '   medixmasterconnWS.Close
  '  Set medixmasterconnWS = Nothing
    CmdSearch.Enabled = True
    Screen.MousePointer = Default
End Sub

Private Function CartonSearchRecord()
Dim strsql As String
Dim strAppend As String
Dim sql As String
Dim storage As New ADODB.Recordset

    strAppend = ""
    If chkscan1.Value = 1 Then
        If Len(Trim(TxtcrtCartonNo)) Then
            strAppend = strAppend & " AND CRTCartonNo = '" & RemStr(Trim(TxtcrtCartonNo)) & "'"
        End If
    ElseIf Len(Trim(TxtcrtCartonNo)) Or Len(Trim(TxtcrtCartonNo2)) Then
        strAppend = strAppend & " AND CRTCartonNo >= '" & RemStr(Trim(TxtcrtCartonNo)) & _
        "' AND CRTCartonNo <= '" & RemStr(Trim(TxtcrtCartonNo2)) & "'"
    End If
    
    If Len(Trim(Mid(CboCartonStatus, 1, 1))) Then
        strAppend = strAppend + " AND CRTCartonStatus = '" & RemStr(Trim(Mid(CboCartonStatus, 1, 1))) & "'"
    End If
    
    If IsDate(MaskcrtDOSend) = True Then
        strAppend = strAppend + " And DOSend = '" & Format(MaskcrtDOSend, "mm/dd/yyyy") & "'"
    End If
    
    If IsDate(MaskcrtDORqst) = True Then
        strAppend = strAppend + " And DORequest = '" & Format(MaskcrtDORqst, "mm/dd/yyyy") & "'"
    End If
    
    If IsDate(MaskcrtDODelivery) = True Then
        strAppend = strAppend + " And DODelivery = '" & Format(MaskcrtDODelivery, "mm/dd/yyyy") & "'"
    End If
    
    If IsDate(MaskcrtDOReturn) = True Then
        strAppend = strAppend + " And DOReturn = '" & Format(MaskcrtDOReturn, "mm/dd/yyyy") & "'"
    End If
    
    If IsDate(MaskcrtDODestroy) = True Then
        strAppend = strAppend + " And DODestroy = '" & Format(MaskcrtDODestroy, "mm/dd/yyyy") & "'"
    End If
    
    Call CartonListing(strAppend)
End Function

Private Sub CartonListing(Optional strAppend As String)
Dim rst2 As New ADODB.Recordset
Dim rstCount As New ADODB.Recordset
Dim Record As ListItem
Dim sql As String
Dim sqlCount As String
Dim total As String
Dim Current As String
    
    total = 0
    Current = 0
        
    If EditFlag = False Then
        Lstcrt.ListItems.Clear
    End If

    sql = " Select crtcartonno, crtcartonstatus, dosend, dorequest, " & _
          " dodelivery, doreturn, dodestroy, crtremarks, lastupdatedate from carton where 0 = 0 "
        
    If Len(Trim(strAppend)) Then
        sql = sql + strAppend
    End If
    
    rst2.Open sql, medixmasterconnWS, adOpenKeyset, adLockOptimistic
    
    If rst2.recordCount = 0 Then
        MsgBox "No record found!", vbOKOnly + vbCritical, DialogBoxTitle
        CmdAdd.Enabled = True
    Else
    total = rst2.recordCount
    lblTotal = total
    Do
        Do Until rst2.EOF
        
           With rst2
                Set Record = Lstcrt.ListItems.Add(, , rst2!CRTCartonNo)
                
                If Len(rst2!CRTCartonStatus) Then
                    Record.SubItems(1) = rst2!CRTCartonStatus
                End If
                If Len(rst2!LastUpdateDate) Then
                    Record.SubItems(2) = Format(rst2!LastUpdateDate, "dd/mm/yyyy")
                End If
                If Len(rst2!DoSend) Then
                    Record.SubItems(3) = Format(rst2!DoSend, "dd/mm/yyyy")
                End If
                If Len(rst2!DORequest) Then
                    Record.SubItems(4) = Format(rst2!DORequest, "dd/mm/yyyy")
                End If
                If Len(rst2!DODelivery) Then
                    Record.SubItems(5) = Format(rst2!DODelivery, "dd/mm/yyyy")
                End If
                If Len(rst2!DOReturn) Then
                    Record.SubItems(6) = Format(rst2!DOReturn, "dd/mm/yyyy")
                End If
                If Len(rst2!dodestroy) Then
                    Record.SubItems(7) = Format(rst2!dodestroy, "dd/mm/yyyy")
                End If
                If Len(rst2!CRTRemarks) Then
                    Record.SubItems(8) = rst2!CRTRemarks
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

Private Sub CmdStop_Click()
intExit = 1
End Sub

Private Sub SaveStore(CASNo As String, CASVer As Integer)
Dim RecordSaved
Dim SaveSTORAGE As New ADODB.Recordset
Dim strsql As String
Dim sql2 As String
Dim Store As New ADODB.Recordset
Dim CLM As New ADODB.Recordset
Dim CARTON As New ADODB.Recordset
Dim strAppend As String
    

    strsql = "Select * From STORAGE Where STRCaseNo = '" & Trim(CASNo) & "' AND STRClmVersion = '" & Trim(CASVer) & "'"
    
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
            !STRCaseNo = CASNo
            !StrCLMVersion = CASVer
            !STRCartonNo = ChkStr(CboCarton2)
            !STRFileStatus = Trim(Mid(CboFileStatus2, 1, 1))
            !STRRequestBy = ChkStr(CboUser2)
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

Private Sub CaseListing(Optional strAppend As String)
Dim rst2 As New ADODB.Recordset
Dim rst3 As New ADODB.Recordset
Dim Record As ListItem
Dim sql As String
Dim sqlStorage As String
    
    sql = " Select CASNumber From CAS WHERE 0 = 0 "
    
    If Len(Trim(strAppend)) Then
        sql = sql + strAppend
    End If

    rst2.Open sql, medixmasterconn, adOpenKeyset, adLockOptimistic
    
    If rst2.recordCount = 0 Then
        MsgBox "No record found!", vbOKOnly + vbCritical
    Else
        
        sqlStorage = " Select STRCaseNo From Storage WHERE STRCaseNo = '" & Trim(TxtstrCasNo) & "'"
        
        rst3.Open sqlStorage, medixmasterconnWS, adOpenKeyset, adLockOptimistic
        
        If rst3.recordCount > 0 Then
            MsgBox "Data has been saved!", vbOKOnly + vbCritical
        Else
            Do Until rst2.EOF
                     
                With rst2
                    If Len(!CasNumber) Then
                        Set Record = Lststr.ListItems.Add(, , !CasNumber)
                    End If
                End With
                rst2.MoveNext
            Loop
        rst3.Close
        Set rst3 = Nothing
        End If
    End If
    intExit = 0
    rst2.Close
    Set rst2 = Nothing
End Sub

Private Function CaseNoCheck()
    If Trim(ChkStr(TxtstrCasNo)) = ChkStr(Lststr.ListItems(listloop1).Text) Then
        Check = True
    End If
End Function

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If Shift = 0 Then
        Select Case KeyCode
            Case vbKeyF6
                KeyCode = 0
                PrintScreenSnapShot
        End Select
    End If
End Sub

Private Sub Form_Load()
    CmdEdit.Enabled = False
    intExit = 0
    Check = False
    CheckCnt = 0
    Label15.Visible = False
    TxtstrCasNo2.Visible = False
    ChkScan.Left = 3000
    ChkScan.Top = 380
    EditFlag = False
    CmdDelete.Enabled = False
    SSTab1.Tab = 1
  '  medixmasterconnWS.Open "File Name=" & BOSPath & "Conn\mediconnectWS.UDL"
  '  medixmasterconnMaint.Open "File Name=" & BOSPath & "Conn\MediConnHISMaint.UDL"
    Call ComboListing
 '   medixmasterconnWS.Close
 '   Set medixmasterconnWS = Nothing
  '  medixmasterconnMaint.Close
  '  Set medixmasterconnMaint = Nothing
End Sub

'Private Sub Lstcrt_Click()

'    If Lstcrt.listitems.Count <> 0 Then
'        TxtcrtCartonNo = Lstcrt.SelectedItem.Text
'            CboCartonStatus = Lstcrt.SelectedItem.SubItems(1)
'        If Len(Trim(Lstcrt.SelectedItem.SubItems(7))) Then
'            MaskcrtDODestroy = Lstcrt.SelectedItem.SubItems(7)
'        End If
'        If Len(Trim(Lstcrt.SelectedItem.SubItems(4))) Then
'            MaskcrtDORqst = Lstcrt.SelectedItem.SubItems(4)
'        End If
'        If Len(Trim(Lstcrt.SelectedItem.SubItems(3))) Then
'            MaskcrtDOSend = Lstcrt.SelectedItem.SubItems(3)
'        End If
'        If Len(Trim(Lstcrt.SelectedItem.SubItems(5))) Then
'            MaskcrtDODelivery = Lstcrt.SelectedItem.SubItems(5)
'        End If
'        If Len(Trim(Lstcrt.SelectedItem.SubItems(6))) Then
'            MaskcrtDOReturn = Lstcrt.SelectedItem.SubItems(6)
'        End If
'        If Len(Trim(Lstcrt.SelectedItem.SubItems(8))) Then
'            TxtcrtRemarks = Lstcrt.SelectedItem.SubItems(8)
'        End If
'
'    End If
'End Sub

Private Sub Lstcrt_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
    Lstcrt.SortKey = ColumnHeader.Index - 1
    Lstcrt.Sorted = True
    If Lstcrt.SortOrder = lvwAscending Then
        Lstcrt.SortOrder = lvwDescending
    Else
        Lstcrt.SortOrder = lvwAscending
    End If
End Sub


Private Sub Lstcrt_ItemCheck(ByVal Item As MSComctlLib.ListItem)
Dim SS, CheckCnt As Integer
    
    CmdEdit.Enabled = True
    TxtcrtCartonNo = ""
    CboCartonStatus = ""
    TxtcrtRemarks = ""
    'MaskcrtDOSend = "##/##/####"
    'MaskcrtDODelivery = "##/##/####"
    'MaskcrtDOReturn = "##/##/####"
    'MaskcrtDODestroy = "##/##/####"
    
    CheckCnt = 0
    For SS = 1 To Lstcrt.ListItems.Count
        If Lstcrt.ListItems(SS).Checked = True Then
            Lstcrt.ListItems(SS).Selected = True
            CheckCnt = CheckCnt + 1
        End If
    Next SS
    LblRecChk = CheckCnt
    CmdSave.Enabled = False
    
    
End Sub

Private Sub Lststr_Click()
    FrameuPDATE.Enabled = False
    CmdEdit.Enabled = False
    CmdSave.Enabled = False
    CmdSave.Caption = "Sa&ve"
    
    TxtCase1 = ""
    txtVersion = ""
    CboCarton2.Text = ""
    CboCartonStatus.Text = ""
    CboFileStatus2.Text = ""
    CboUser2.Text = ""
    txtDORqst = "__/__/____"
    txtDOD = "__/__/____"
    txtDOReturn = "__/__/____"
    TxtRemarks = ""
    
    If Lststr.ListItems.Count <> 0 Then
        If Trim(Lststr.SelectedItem.SubItems(2)) <> "" Then
            CmdEdit.Enabled = True
            Call ClearStoreInfo
            TxtCase1 = Lststr.SelectedItem.Text
            txtVersion = Lststr.SelectedItem.SubItems(1)
            CboCarton2.Text = Lststr.SelectedItem.SubItems(2)
            CboFileStatus2.Text = Lststr.SelectedItem.SubItems(4)
            CboUser2.Text = Lststr.SelectedItem.SubItems(7)
            txtDORqst = IIf(Len(Lststr.SelectedItem.SubItems(8)), Lststr.SelectedItem.SubItems(8), "__/__/____")
            txtDOD = IIf(Len(Lststr.SelectedItem.SubItems(9)), Lststr.SelectedItem.SubItems(9), "__/__/____")
            txtDOReturn = IIf(Len(Lststr.SelectedItem.SubItems(10)), Lststr.SelectedItem.SubItems(10), "__/__/____")
            TxtRemarks = IIf(Len(Lststr.SelectedItem.SubItems(12)), Lststr.SelectedItem.SubItems(12), "")
        End If
    End If
End Sub

Private Sub Lststr_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
    Lststr.SortKey = ColumnHeader.Index - 1
    Lststr.Sorted = True
    If Lststr.SortOrder = lvwAscending Then
        Lststr.SortOrder = lvwDescending
    Else
        Lststr.SortOrder = lvwAscending
    End If
End Sub

Private Sub MaskcrtDODelivery_LostFocus()
    If IsDate(MaskcrtDODelivery) Then
        If Format(MaskcrtDODelivery, "YYYYMMDD") < Format(MaskcrtDORqst, "YYYYMMDD") Then
            MsgBox "Delivery date must be GREATER than  Date Request ! ", vbCritical
            MaskcrtDODelivery.SetFocus
        End If
    Else
        If MaskcrtDODelivery <> "__/__/____" Then
            MsgBox "Invalid Date Format! ", vbCritical
            MaskcrtDODelivery = "__/__/____"
            MaskcrtDODelivery.SetFocus
        End If
    End If
End Sub

Private Sub MaskcrtDODestroy_LostFocus()
    If IsDate(MaskcrtDODestroy) Then
        If Format(MaskcrtDODestroy, "YYYYMMDD") < Format(MaskcrtDOReturn, "YYYYMMDD") Then
            MsgBox "Destroy Date must be GREATER than Return Date! ", vbCritical
            MaskcrtDODestroy.SetFocus
        End If
    Else
        If MaskcrtDODestroy <> "__/__/____" Then
            MsgBox "Invalid Date Format! ", vbCritical
            MaskcrtDODestroy = "__/__/____"
            MaskcrtDODestroy.SetFocus
        End If
    End If
End Sub

Private Sub MaskcrtDOReturn_LostFocus()
    If IsDate(MaskcrtDOReturn) Then
        If Format(MaskcrtDOReturn, "YYYYMMDD") < Format(MaskcrtDODelivery, "YYYYMMDD") Then
            MsgBox "Return Date must be GREATER than Delivery Date! ", vbCritical
            MaskcrtDOReturn.SetFocus
        End If
    Else
        If MaskcrtDOReturn <> "__/__/____" Then
            MsgBox "Invalid Date Format! ", vbCritical
            MaskcrtDOReturn = "__/__/____"
            MaskcrtDOReturn.SetFocus
        End If
    End If
End Sub

Private Sub MaskcrtDORqst_LostFocus()
    If IsDate(MaskcrtDORqst) Then
        Else
            If Format(MaskcrtDORqst, "YYYYMMDD") < Format(MaskcrtDOSend, "YYYYMMDD") Then
                MsgBox "Request Date must be GREATER than Date Send ! ", vbCritical
                MaskcrtDORqst.SetFocus
            End If
        If MaskcrtDORqst <> "__/__/____" Then
           MsgBox "Invalid Date Format! ", vbCritical
           MaskcrtDORqst = "__/__/____"
           MaskcrtDORqst.SetFocus
        End If
    End If

End Sub

Private Sub MaskcrtDOSend_LostFocus()
    If IsDate(MaskcrtDOSend) Then
        Else
            If MaskcrtDOSend <> "__/__/____" Then
                MsgBox "Invalid Date Format! ", vbCritical
                MaskcrtDOSend = "__/__/____"
                MaskcrtDOSend.SetFocus
            End If
    End If
End Sub

Private Sub MaskstrDODelivery_LostFocus()
    If IsDate(MaskstrDODelivery) Then
        If Format(MaskstrDODelivery, "YYYYMMDD") < Format(MaskstrDORqst, "YYYYMMDD") Then
            MsgBox "Delivery Date must Greater than Request Date ! ", vbCritical
            MaskstrDODelivery.SetFocus
        End If
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
        If Format(MaskstrDOReturn, "YYYYMMDD") < Format(MaskstrDODelivery, "YYYYMMDD") Then
            MsgBox "Return Date must Greater than Delviery Date ! ", vbCritical
            MaskstrDOReturn.SetFocus
        End If
    Else
        If MaskstrDOReturn <> "__/__/____" Then
            MsgBox "Invalid Date Format! ", vbCritical
            MaskstrDOReturn = "__/__/____"
            MaskstrDOReturn.SetFocus
        End If
    End If
End Sub

Private Sub MaskstrDORqst_LostFocus()
    If IsDate(MaskstrDORqst) = False Then
        If MaskstrDORqst <> "__/__/____" Then
            MsgBox "Invalid Date Format! ", vbCritical
            MaskstrDORqst = "__/__/____"
            MaskstrDORqst.SetFocus
        End If
    End If
End Sub

Private Sub SSTab1_Click(PreviousTab As Integer)
    Screen.MousePointer = vbHourglass
    Select Case SSTab1.Tab
        Case 0
            Call ClearCartonInfo
            CmdSave.Default = False
            Lstcrt.ListItems.Clear
        Case 1
            Call ClearStoreInfo
            ChkScan.Value = 1
            Lststr.ListItems.Clear
    End Select
    Screen.MousePointer = Default
End Sub

Private Sub TxtcrtCartonNo_LostFocus()
    TxtcrtCartonNo = ChkStr(TxtcrtCartonNo)
End Sub

Private Sub TxtcrtRemarks_LostFocus()
    TxtcrtRemarks = ChkStr(TxtcrtRemarks)
End Sub

Private Sub txtDOD_LostFocus()
    If IsDate(txtDOD) Then
        If Format(txtDOD, "YYYYMMDD") < Format(txtDORqst, "YYYYMMDD") Then
            MsgBox "Delivery Date must Greater than Request Date ! ", vbCritical
            txtDOD.SetFocus
        End If
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
        If Format(txtDOReturn, "YYYYMMDD") < Format(txtDOD, "YYYYMMDD") Then
            MsgBox "Return Date must Greater than Delivery Date ! ", vbCritical
            txtDOReturn.SetFocus
        End If
    Else
        If txtDOReturn <> "__/__/____" Then
            MsgBox "Invalid Date Format! ", vbCritical
            txtDOReturn = "__/__/____"
            txtDOReturn.SetFocus
        End If
    End If
End Sub

Private Sub txtDORqst_LostFocus()
    If IsDate(txtDORqst) = False Then
        If txtDORqst <> "__/__/____" Then
            MsgBox "Invalid Date Format! ", vbCritical
            txtDORqst = "__/__/____"
            txtDORqst.SetFocus
        End If
    End If
End Sub

Private Sub TxtstrCasNo_GotFocus()
    CmdSearch.Default = True
End Sub

Private Sub TxtstrCasNo_LostFocus()
    TxtstrCasNo = ChkStr(TxtstrCasNo)
    CmdSearch.Default = False
End Sub

Private Sub EntriesEnable(status As Boolean)
    '----- STORAGE
    TxtstrCasNo.Enabled = status
    CboCarton1.Enabled = status
    CboFileStatus1.Enabled = status
    CboUser1.Enabled = status
    MaskstrDORqst.Enabled = status
    MaskstrDODelivery.Enabled = status
    MaskstrDOReturn.Enabled = status
    
    '----- CARTON
    TxtcrtCartonNo.Enabled = status
    CboCartonStatus.Enabled = status
    MaskcrtDODestroy.Enabled = status
    MaskcrtDOSend.Enabled = status
    MaskcrtDORqst.Enabled = status
    MaskcrtDODelivery.Enabled = status
    MaskcrtDOReturn.Enabled = status
    TxtcrtRemarks.Enabled = status
    
End Sub

Private Sub ComboListing()
Dim USRRec As New ADODB.Recordset
Dim CRTRec As New ADODB.Recordset
Dim cmd As New ADODB.Command
    
    CboFileStatus1.Clear
    CboFileStatus1.AddItem "A - All"
    CboFileStatus1.AddItem "I - In"
    CboFileStatus1.AddItem "O - Out"
    CboFileStatus1.ListIndex = 0
    
    CboCartonStatus.Clear
    CboCartonStatus.AddItem "M - Medix"
    CboCartonStatus.AddItem "S - Store"
    CboCartonStatus.AddItem "R - Return"
    
    CboFileStatus2.Clear
    CboFileStatus2.AddItem "I - In"
    CboFileStatus2.AddItem "O - Out"
        
    Set cmd.ActiveConnection = medixmasterconnWS
    cmd.CommandText = "SearchCarton"
    cmd.CommandType = adCmdStoredProc
    cmd.CommandTimeout = 15
    CRTRec.CursorLocation = adUseClient
    CRTRec.CursorType = adOpenForwardOnly
    CRTRec.CacheSize = 100
    Set CRTRec = cmd.Execute
    
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
    
    Set cmd.ActiveConnection = medixmasterconnMaint
    cmd.CommandText = "SearchUser"
    cmd.CommandType = adCmdStoredProc
    cmd.CommandTimeout = 15
    USRRec.CursorLocation = adUseClient
    USRRec.CursorType = adOpenForwardOnly
    USRRec.CacheSize = 100
    Set USRRec = cmd.Execute
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
End Sub

Private Sub txtRemarks_LostFocus()
    TxtRemarks = ChkStr(TxtRemarks)
End Sub

Private Sub Lststr_ItemCheck(ByVal Item As MSComctlLib.ListItem)
    CmdSave.Enabled = True
    FrameuPDATE.Enabled = True
   ' medixmasterconnWS.Open "File Name=" & BOSPath & "Conn\mediconnectWS.UDL"
    Call ProcessItemChecked(Item)
  '  medixmasterconnWS.Close
  '  Set medixmasterconnWS = Nothing

End Sub

Private Function ProcessItemChecked(LstREc As ListItem, Optional Partial As Boolean)
Dim BB, newbound  As Integer
Dim tempArray() As String
Dim strsql As String
Dim StoRec As New ADODB.Recordset
Dim CaseNo As String
Dim tmpSelCasNo As String

Static Arraycnt As Integer
    
    
     CaseNo = LstREc.Text
     strsql = "Select * From Storage where STRCASeNo = '" & Trim(CaseNo) & "' AND STRClmVersion = '" & LstREc.SubItems(1) & "'"
     StoRec.Open strsql, medixmasterconnWS, adOpenKeyset, adLockPessimistic

    
    If LstREc.Checked = True Then
        If StoRec.recordCount = 0 Then
            ReDim Preserve CaseNoArray(Arraycnt)
            CaseNoArray(Arraycnt) = LstREc.Text & "-" & LstREc.SubItems(1) & "~"
            Arraycnt = Arraycnt + 1
            CheckCnt = CheckCnt + 1
            LblRecChk = CheckCnt
            If CheckCnt > 70 Then
                MsgBox "Files Limit Exceeded! ", vbCritical
            End If
        Else
            MsgBox "Claim has been stored!", vbCritical
            LstREc.Checked = False
        End If
    Else
        Arraycnt = Arraycnt - 1
        CheckCnt = CheckCnt - 1
        LblRecChk = IIf(CheckCnt < 0, 0, CheckCnt)

        If UBound(CaseNoArray) <> 0 Then
            For BB = 0 To UBound(CaseNoArray)
                tmpSelCasNo = LstREc.Text & "-" & LstREc.SubItems(1) & "~"
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

Private Sub ClearCartonInfo()

    TxtcrtCartonNo.Text = ""
    CboCartonStatus.Text = ""
    MaskcrtDODestroy = "__/__/____"
    MaskcrtDOSend = "__/__/____"
    MaskcrtDORqst = "__/__/____"
    MaskcrtDODelivery = "__/__/____"
    MaskcrtDOReturn = "__/__/____"
    TxtcrtRemarks.Text = ""
    
End Sub

Private Sub ClearStoreInfo()
    CboCarton1.Text = ""
    CboFileStatus1.Text = ""
    CboUser1.Text = ""
    MaskstrDORqst = "__/__/____"
    MaskstrDODelivery = "__/__/____"
    MaskstrDOReturn = "__/__/____"
    TxtCase1 = ""
    txtVersion = ""
    CboCarton2.Text = ""
    CboCartonStatus.Text = ""
    CboFileStatus2.Text = ""
    CboUser2.Text = ""
    txtDORqst = "__/__/____"
    txtDOD = "__/__/____"
    txtDOReturn = "__/__/____"
    TxtRemarks = ""
End Sub


