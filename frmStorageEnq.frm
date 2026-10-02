VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "tabctl32.ocx"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Begin VB.Form frmStorageEnq 
   Caption         =   "Storage Enquiry"
   ClientHeight    =   8205
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   11880
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   8205
   ScaleWidth      =   11880
   WindowState     =   2  'Maximized
   Begin TabDlg.SSTab SSTab1 
      Height          =   7095
      Left            =   120
      TabIndex        =   25
      Top             =   240
      Width           =   11700
      _ExtentX        =   20638
      _ExtentY        =   12515
      _Version        =   393216
      Tabs            =   2
      Tab             =   1
      TabsPerRow      =   2
      TabHeight       =   520
      TabCaption(0)   =   "Carton Details"
      TabPicture(0)   =   "frmStorageEnq.frx":0000
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "Lstcrt"
      Tab(0).Control(1)=   "Frame4"
      Tab(0).ControlCount=   2
      TabCaption(1)   =   "Storage Details"
      TabPicture(1)   =   "frmStorageEnq.frx":001C
      Tab(1).ControlEnabled=   -1  'True
      Tab(1).Control(0)=   "Lststr"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).Control(1)=   "Frame2"
      Tab(1).Control(1).Enabled=   0   'False
      Tab(1).Control(2)=   "FrameuPDATE"
      Tab(1).Control(2).Enabled=   0   'False
      Tab(1).ControlCount=   3
      Begin VB.Frame FrameuPDATE 
         Enabled         =   0   'False
         Height          =   1215
         Left            =   360
         TabIndex        =   66
         Top             =   5760
         Width           =   11055
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
            TabIndex        =   69
            Top             =   240
            Width           =   2055
         End
         Begin MSMask.MaskEdBox txtDORqst 
            Height          =   285
            Left            =   6600
            TabIndex        =   70
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
            TabIndex        =   71
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
            TabIndex        =   72
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
            TabIndex        =   67
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
            TabIndex        =   68
            Top             =   240
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
            ItemData        =   "frmStorageEnq.frx":0038
            Left            =   3720
            List            =   "frmStorageEnq.frx":003A
            TabIndex        =   73
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
            TabIndex        =   74
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
            TabIndex        =   75
            Top             =   765
            Width           =   3615
         End
         Begin VB.Label Label40 
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
            TabIndex        =   84
            Top             =   240
            Width           =   735
         End
         Begin VB.Label Label39 
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
            TabIndex        =   83
            Top             =   795
            Width           =   1335
         End
         Begin VB.Label Label38 
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
            TabIndex        =   82
            Top             =   525
            Width           =   1335
         End
         Begin VB.Label Label37 
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
            TabIndex        =   81
            Top             =   255
            Width           =   1335
         End
         Begin VB.Label Label36 
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
            TabIndex        =   80
            Top             =   525
            Width           =   975
         End
         Begin VB.Label Label35 
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
            TabIndex        =   79
            Top             =   520
            Width           =   855
         End
         Begin VB.Label Label34 
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
            TabIndex        =   78
            Top             =   840
            Width           =   975
         End
         Begin VB.Label Label33 
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
            TabIndex        =   77
            Top             =   255
            Width           =   990
         End
         Begin VB.Label Label32 
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
            TabIndex        =   76
            Top             =   240
            Width           =   855
         End
      End
      Begin VB.Frame Frame4 
         Height          =   2655
         Left            =   -74640
         TabIndex        =   36
         Top             =   600
         Width           =   11055
         Begin VB.TextBox TxtcrtRemarks 
            Height          =   855
            Left            =   1800
            MaxLength       =   3998
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   13
            Top             =   1560
            Width           =   8415
         End
         Begin MSMask.MaskEdBox MaskcrtDORqst 
            Height          =   285
            Left            =   7440
            TabIndex        =   5
            Top             =   360
            Width           =   1215
            _ExtentX        =   2143
            _ExtentY        =   503
            _Version        =   393216
            MaxLength       =   10
            Mask            =   "##/##/####"
            PromptChar      =   "_"
         End
         Begin MSMask.MaskEdBox MaskcrtDODelivery 
            Height          =   285
            Left            =   7440
            TabIndex        =   7
            Top             =   600
            Width           =   1215
            _ExtentX        =   2143
            _ExtentY        =   503
            _Version        =   393216
            MaxLength       =   10
            Mask            =   "##/##/####"
            PromptChar      =   "_"
         End
         Begin MSMask.MaskEdBox MaskcrtDOReturn 
            Height          =   285
            Left            =   7440
            TabIndex        =   9
            Top             =   840
            Width           =   1215
            _ExtentX        =   2143
            _ExtentY        =   503
            _Version        =   393216
            MaxLength       =   10
            Mask            =   "##/##/####"
            PromptChar      =   "_"
         End
         Begin MSMask.MaskEdBox MaskcrtDODestroy 
            Height          =   315
            Left            =   7440
            TabIndex        =   11
            Top             =   1080
            Width           =   1215
            _ExtentX        =   2143
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
            Width           =   1575
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
            ItemData        =   "frmStorageEnq.frx":003C
            Left            =   1800
            List            =   "frmStorageEnq.frx":003E
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
            Width           =   1575
            _ExtentX        =   2778
            _ExtentY        =   503
            _Version        =   393216
            MaxLength       =   10
            Mask            =   "##/##/####"
            PromptChar      =   "_"
         End
         Begin MSMask.MaskEdBox MaskcrtDORqst2 
            Height          =   285
            Left            =   9000
            TabIndex        =   6
            Top             =   360
            Width           =   1215
            _ExtentX        =   2143
            _ExtentY        =   503
            _Version        =   393216
            MaxLength       =   10
            Mask            =   "##/##/####"
            PromptChar      =   "_"
         End
         Begin MSMask.MaskEdBox MaskcrtDODelivery2 
            Height          =   285
            Left            =   9000
            TabIndex        =   8
            Top             =   600
            Width           =   1215
            _ExtentX        =   2143
            _ExtentY        =   503
            _Version        =   393216
            MaxLength       =   10
            Mask            =   "##/##/####"
            PromptChar      =   "_"
         End
         Begin MSMask.MaskEdBox MaskcrtDOReturn2 
            Height          =   285
            Left            =   9000
            TabIndex        =   10
            Top             =   840
            Width           =   1215
            _ExtentX        =   2143
            _ExtentY        =   503
            _Version        =   393216
            MaxLength       =   10
            Mask            =   "##/##/####"
            PromptChar      =   "_"
         End
         Begin MSMask.MaskEdBox MaskcrtDODestroy2 
            Height          =   315
            Left            =   9000
            TabIndex        =   12
            Top             =   1080
            Width           =   1215
            _ExtentX        =   2143
            _ExtentY        =   556
            _Version        =   393216
            MaxLength       =   10
            Mask            =   "##/##/####"
            PromptChar      =   "_"
         End
         Begin MSMask.MaskEdBox MaskcrtDOSend2 
            Height          =   285
            Left            =   3720
            TabIndex        =   4
            Top             =   915
            Width           =   1575
            _ExtentX        =   2778
            _ExtentY        =   503
            _Version        =   393216
            MaxLength       =   10
            Mask            =   "##/##/####"
            PromptChar      =   "_"
         End
         Begin VB.Label Label31 
            Caption         =   "To"
            Height          =   255
            Left            =   3450
            TabIndex        =   64
            Top             =   975
            Width           =   375
         End
         Begin VB.Label Label30 
            Caption         =   "To"
            Height          =   255
            Left            =   8720
            TabIndex        =   63
            Top             =   1140
            Width           =   375
         End
         Begin VB.Label Label29 
            Caption         =   "To"
            Height          =   255
            Left            =   8720
            TabIndex        =   62
            Top             =   900
            Width           =   375
         End
         Begin VB.Label Label28 
            Caption         =   "To"
            Height          =   255
            Left            =   8720
            TabIndex        =   61
            Top             =   645
            Width           =   375
         End
         Begin VB.Label Label27 
            Caption         =   "To"
            Height          =   255
            Left            =   8720
            TabIndex        =   60
            Top             =   390
            Width           =   375
         End
         Begin VB.Label Label10 
            Caption         =   "Date Of Sending"
            Height          =   255
            Left            =   480
            TabIndex        =   45
            Top             =   975
            Width           =   1455
         End
         Begin VB.Label Label9 
            Caption         =   "Date Of Destroy"
            Height          =   255
            Left            =   6000
            TabIndex        =   44
            Top             =   1140
            Width           =   1215
         End
         Begin VB.Label Label24 
            Caption         =   "Remarks"
            Height          =   255
            Left            =   480
            TabIndex        =   43
            Top             =   1560
            Width           =   975
         End
         Begin VB.Label Label23 
            Caption         =   "Date Of Return"
            Height          =   255
            Left            =   6000
            TabIndex        =   42
            Top             =   900
            Width           =   1335
         End
         Begin VB.Label Label22 
            Caption         =   "Carton Status"
            Height          =   255
            Left            =   480
            TabIndex        =   41
            Top             =   680
            Width           =   975
         End
         Begin VB.Label Label21 
            Caption         =   "Date Of Delivery"
            Height          =   255
            Left            =   6000
            TabIndex        =   40
            Top             =   645
            Width           =   1455
         End
         Begin VB.Label Label19 
            Caption         =   "Date Of Request"
            Height          =   255
            Left            =   6000
            TabIndex        =   39
            Top             =   390
            Width           =   1455
         End
         Begin VB.Label Label18 
            Caption         =   "Carton No"
            Height          =   255
            Left            =   480
            TabIndex        =   38
            Top             =   375
            Width           =   735
         End
         Begin VB.Label Label16 
            Caption         =   "To"
            Height          =   255
            Left            =   3450
            TabIndex        =   37
            Top             =   360
            Width           =   375
         End
      End
      Begin VB.Frame Frame2 
         Height          =   1695
         Left            =   360
         TabIndex        =   26
         Top             =   480
         Width           =   11055
         Begin VB.CheckBox ChkScan 
            Caption         =   "Use Scan"
            Height          =   255
            Left            =   3120
            TabIndex        =   65
            Top             =   240
            Value           =   1  'Checked
            Width           =   1050
         End
         Begin VB.TextBox TxtstrCasNo2 
            Height          =   285
            Left            =   4080
            MaxLength       =   20
            TabIndex        =   15
            Top             =   240
            Width           =   1695
         End
         Begin MSMask.MaskEdBox MaskstrDORqst 
            Height          =   285
            Left            =   8040
            TabIndex        =   19
            Top             =   240
            Width           =   1215
            _ExtentX        =   2143
            _ExtentY        =   503
            _Version        =   393216
            MaxLength       =   10
            Mask            =   "##/##/####"
            PromptChar      =   "_"
         End
         Begin MSMask.MaskEdBox MaskstrDODelivery 
            Height          =   285
            Left            =   8040
            TabIndex        =   21
            Top             =   480
            Width           =   1215
            _ExtentX        =   2143
            _ExtentY        =   503
            _Version        =   393216
            MaxLength       =   10
            Mask            =   "##/##/####"
            PromptChar      =   "_"
         End
         Begin MSMask.MaskEdBox MaskstrDOReturn 
            Height          =   285
            Left            =   8040
            TabIndex        =   23
            Top             =   720
            Width           =   1215
            _ExtentX        =   2143
            _ExtentY        =   503
            _Version        =   393216
            MaxLength       =   10
            Mask            =   "##/##/####"
            PromptChar      =   "_"
         End
         Begin VB.TextBox TxtstrCasNo 
            Height          =   285
            Left            =   1320
            MaxLength       =   20
            TabIndex        =   14
            Top             =   240
            Width           =   1695
         End
         Begin VB.ComboBox CboCarton 
            Height          =   315
            Left            =   1320
            TabIndex        =   16
            Top             =   480
            Width           =   1695
         End
         Begin VB.TextBox TxtstrClmNo 
            Height          =   285
            Left            =   4200
            MaxLength       =   3
            TabIndex        =   27
            Top             =   240
            Visible         =   0   'False
            Width           =   975
         End
         Begin MSMask.MaskEdBox MaskstrDORqst2 
            Height          =   285
            Left            =   9720
            TabIndex        =   20
            Top             =   240
            Width           =   1215
            _ExtentX        =   2143
            _ExtentY        =   503
            _Version        =   393216
            MaxLength       =   10
            Mask            =   "##/##/####"
            PromptChar      =   "_"
         End
         Begin MSMask.MaskEdBox MaskstrDODelivery2 
            Height          =   285
            Left            =   9720
            TabIndex        =   22
            Top             =   480
            Width           =   1215
            _ExtentX        =   2143
            _ExtentY        =   503
            _Version        =   393216
            MaxLength       =   10
            Mask            =   "##/##/####"
            PromptChar      =   "_"
         End
         Begin MSMask.MaskEdBox MaskstrDOReturn2 
            Height          =   285
            Left            =   9720
            TabIndex        =   24
            Top             =   720
            Width           =   1215
            _ExtentX        =   2143
            _ExtentY        =   503
            _Version        =   393216
            MaxLength       =   10
            Mask            =   "##/##/####"
            PromptChar      =   "_"
         End
         Begin VB.ComboBox CboCartonto 
            Height          =   315
            Left            =   4080
            TabIndex        =   86
            Top             =   480
            Width           =   1695
         End
         Begin VB.ComboBox cboPayorFr 
            Height          =   315
            Left            =   1320
            TabIndex        =   89
            Top             =   770
            Width           =   1695
         End
         Begin VB.ComboBox cboPayorTo 
            Height          =   315
            Left            =   4080
            TabIndex        =   90
            Top             =   770
            Width           =   1695
         End
         Begin MSMask.MaskEdBox txtRegDateFr 
            Height          =   285
            Left            =   8040
            TabIndex        =   93
            Top             =   960
            Width           =   1215
            _ExtentX        =   2143
            _ExtentY        =   503
            _Version        =   393216
            MaxLength       =   10
            Mask            =   "##/##/####"
            PromptChar      =   "_"
         End
         Begin MSMask.MaskEdBox txtRegDateTo 
            Height          =   285
            Left            =   9720
            TabIndex        =   94
            Top             =   960
            Width           =   1215
            _ExtentX        =   2143
            _ExtentY        =   503
            _Version        =   393216
            MaxLength       =   10
            Mask            =   "##/##/####"
            PromptChar      =   "_"
         End
         Begin MSMask.MaskEdBox txtSentDateFr 
            Height          =   285
            Left            =   8040
            TabIndex        =   97
            Top             =   1200
            Width           =   1215
            _ExtentX        =   2143
            _ExtentY        =   503
            _Version        =   393216
            MaxLength       =   10
            Mask            =   "##/##/####"
            PromptChar      =   "_"
         End
         Begin MSMask.MaskEdBox txtSentDateTo 
            Height          =   285
            Left            =   9720
            TabIndex        =   98
            Top             =   1200
            Width           =   1215
            _ExtentX        =   2143
            _ExtentY        =   503
            _Version        =   393216
            MaxLength       =   10
            Mask            =   "##/##/####"
            PromptChar      =   "_"
         End
         Begin VB.ComboBox CboUser 
            Height          =   315
            Left            =   1320
            Sorted          =   -1  'True
            TabIndex        =   18
            Top             =   1050
            Width           =   4455
         End
         Begin VB.ComboBox CboFileStatus 
            Height          =   315
            ItemData        =   "frmStorageEnq.frx":0040
            Left            =   1320
            List            =   "frmStorageEnq.frx":0042
            TabIndex        =   17
            Text            =   "I - In"
            Top             =   1300
            Width           =   4455
         End
         Begin VB.Label Label7 
            Caption         =   "Sent Date"
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
            Index           =   2
            Left            =   6960
            TabIndex        =   100
            Top             =   1275
            Width           =   1215
         End
         Begin VB.Label Label26 
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
            Index           =   2
            Left            =   9360
            TabIndex        =   99
            Top             =   1275
            Width           =   375
         End
         Begin VB.Label Label7 
            Caption         =   "Register Date"
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
            Index           =   1
            Left            =   6960
            TabIndex        =   96
            Top             =   1035
            Width           =   1215
         End
         Begin VB.Label Label26 
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
            Index           =   1
            Left            =   9360
            TabIndex        =   95
            Top             =   1035
            Width           =   375
         End
         Begin VB.Label Label2 
            Caption         =   "Payor"
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
            Index           =   1
            Left            =   360
            TabIndex        =   92
            Top             =   800
            Width           =   975
         End
         Begin VB.Label Label8 
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
            Index           =   1
            Left            =   3360
            TabIndex        =   91
            Top             =   800
            Width           =   975
         End
         Begin VB.Label Label8 
            Caption         =   "To"
            Height          =   255
            Index           =   0
            Left            =   3360
            TabIndex        =   88
            Top             =   240
            Width           =   975
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
            Index           =   0
            Left            =   360
            TabIndex        =   87
            Top             =   525
            Width           =   975
         End
         Begin VB.Label Label26 
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
            Index           =   0
            Left            =   9360
            TabIndex        =   59
            Top             =   795
            Width           =   375
         End
         Begin VB.Label Label25 
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
            Left            =   9360
            TabIndex        =   58
            Top             =   525
            Width           =   375
         End
         Begin VB.Label Label20 
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
            Left            =   9360
            TabIndex        =   57
            Top             =   255
            Width           =   375
         End
         Begin VB.Label Label7 
            Caption         =   "Return Date"
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
            Index           =   0
            Left            =   6960
            TabIndex        =   35
            Top             =   795
            Width           =   1215
         End
         Begin VB.Label Label6 
            Caption         =   "Version No"
            Height          =   255
            Left            =   3840
            TabIndex        =   34
            Top             =   360
            Visible         =   0   'False
            Width           =   855
         End
         Begin VB.Label Label5 
            Caption         =   "Delivery Date"
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
            Left            =   6960
            TabIndex        =   33
            Top             =   525
            Width           =   1335
         End
         Begin VB.Label Label3 
            Caption         =   "Request Date"
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
            Left            =   6960
            TabIndex        =   32
            Top             =   255
            Width           =   1335
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
            Left            =   360
            TabIndex        =   31
            Top             =   255
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
            Left            =   360
            TabIndex        =   30
            Top             =   1400
            Width           =   855
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
            Left            =   360
            TabIndex        =   29
            Top             =   1100
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
            Left            =   3360
            TabIndex        =   28
            Top             =   550
            Width           =   375
         End
      End
      Begin MSComctlLib.ListView Lstcrt 
         Height          =   3615
         Left            =   -74640
         TabIndex        =   46
         Top             =   3360
         Width           =   11055
         _ExtentX        =   19500
         _ExtentY        =   6376
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
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
         Height          =   3495
         Left            =   360
         TabIndex        =   47
         Top             =   2280
         Width           =   11055
         _ExtentX        =   19500
         _ExtentY        =   6165
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         FullRowSelect   =   -1  'True
         GridLines       =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         NumItems        =   14
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "CaseNo"
            Object.Width           =   2540
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
            SubItemIndex    =   3
            Text            =   "Carton Status"
            Object.Width           =   1764
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   4
            Text            =   "File Status"
            Object.Width           =   1764
         EndProperty
         BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   5
            Text            =   "Bill Amt"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   6
            Text            =   "Date Sent"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   2
            SubItemIndex    =   7
            Text            =   "Date Register"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(9) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   8
            Text            =   "Request By"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(10) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   9
            Text            =   "Date Request"
            Object.Width           =   2118
         EndProperty
         BeginProperty ColumnHeader(11) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   10
            Text            =   "Date Delivery"
            Object.Width           =   2118
         EndProperty
         BeginProperty ColumnHeader(12) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   11
            Text            =   "Date Return"
            Object.Width           =   2117
         EndProperty
         BeginProperty ColumnHeader(13) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   12
            Text            =   "Duration"
            Object.Width           =   1411
         EndProperty
         BeginProperty ColumnHeader(14) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   13
            Text            =   "Remarks"
            Object.Width           =   2540
         EndProperty
      End
   End
   Begin VB.Frame Frame5 
      Height          =   735
      Left            =   120
      TabIndex        =   48
      Top             =   7320
      Width           =   11640
      Begin VB.TextBox lblTotal 
         Height          =   285
         Left            =   1560
         TabIndex        =   102
         Top             =   240
         Width           =   615
      End
      Begin VB.TextBox lblCurrent 
         Height          =   285
         Left            =   240
         TabIndex        =   101
         Top             =   240
         Width           =   735
      End
      Begin VB.CommandButton CmdListing 
         Caption         =   "&Listing"
         Height          =   350
         Left            =   6460
         TabIndex        =   56
         Top             =   240
         Width           =   800
      End
      Begin VB.CommandButton CmdStop 
         Caption         =   "St&op"
         Height          =   350
         Left            =   8060
         TabIndex        =   53
         Top             =   240
         Width           =   800
      End
      Begin VB.CommandButton CmdClear 
         Caption         =   "&Clear"
         Height          =   350
         Left            =   8850
         TabIndex        =   52
         Top             =   240
         Width           =   800
      End
      Begin VB.CommandButton CmdExit 
         Cancel          =   -1  'True
         Caption         =   "E&xit"
         Height          =   350
         Left            =   10440
         TabIndex        =   51
         Top             =   240
         Width           =   800
      End
      Begin VB.CommandButton CmdSearch 
         Caption         =   "&Search"
         Default         =   -1  'True
         Height          =   350
         Left            =   7260
         TabIndex        =   50
         Top             =   240
         Width           =   800
      End
      Begin VB.CommandButton CmdPrintFile 
         Caption         =   "&Print File"
         Height          =   350
         Left            =   9640
         TabIndex        =   49
         Top             =   240
         Width           =   800
      End
      Begin VB.Label Label13 
         Caption         =   "of"
         Height          =   255
         Left            =   1200
         TabIndex        =   55
         Top             =   285
         Visible         =   0   'False
         Width           =   375
      End
      Begin VB.Label Label14 
         Caption         =   "Records"
         Height          =   255
         Left            =   2280
         TabIndex        =   54
         Top             =   285
         Visible         =   0   'False
         Width           =   735
      End
   End
   Begin VB.Label Label17 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00404040&
      Caption         =   "Storage / Cartons Enquiry"
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
      TabIndex        =   85
      Top             =   0
      Width           =   11805
   End
End
Attribute VB_Name = "frmStorageEnq"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public intExit As Integer
Dim bExit As Boolean
Dim EditFlag As Boolean
Dim Current2 As String
Dim listloop1 As Integer
Dim listloop2 As Integer
Dim Check As Boolean
Dim Check2 As Boolean
Dim CaseNoArray() As String

Private Sub CboCarton_KeyPress(KeyAscii As Integer)
Dim NewText As String
Dim ValidCount As Integer
Dim ValidValue As String
Dim i As Integer
    If KeyAscii >= 32 And KeyAscii <> 127 Then
    NewText = LCase(Left(CboCarton.Text, CboCarton.SelStart) + Chr(KeyAscii) + Mid(CboCarton.Text, CboCarton.SelStart + CboCarton.SelLength + 1))
 
    ValidCount = 0
    ValidValue = ""
 
    For i = 0 To CboCarton.ListCount - 1
 
        If NewText = LCase(Left(CboCarton.List(i), Len(NewText))) Then
                ValidCount = ValidCount + 1
                ValidValue = CboCarton.List(i)
        End If
        Next
        If ValidCount <= 1 Then KeyAscii = 0
            If ValidCount = 1 Then
                CboCarton.Text = ValidValue
                CboCarton.SelStart = 0
                CboCarton.SelLength = Len(ValidValue)
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

Private Sub CboFileStatus_KeyPress(KeyAscii As Integer)
Dim NewText As String
Dim ValidCount As Integer
Dim ValidValue As String
Dim i As Integer
    If KeyAscii >= 32 And KeyAscii <> 127 Then
    NewText = LCase(Left(CboFileStatus.Text, CboFileStatus.SelStart) + Chr(KeyAscii) + Mid(CboFileStatus.Text, CboFileStatus.SelStart + CboFileStatus.SelLength + 1))
 
    ValidCount = 0
    ValidValue = ""
 
    For i = 0 To CboFileStatus.ListCount - 1
 
        If NewText = LCase(Left(CboFileStatus.List(i), Len(NewText))) Then
                ValidCount = ValidCount + 1
                ValidValue = CboFileStatus.List(i)
        End If
    Next
    If ValidCount <= 1 Then KeyAscii = 0
        If ValidCount = 1 Then
            CboFileStatus.Text = ValidValue
            CboFileStatus.SelStart = 0
            CboFileStatus.SelLength = Len(ValidValue)
        End If
    End If

End Sub

Private Sub CboUser_KeyPress(KeyAscii As Integer)
Dim NewText As String
Dim ValidCount As Integer
Dim ValidValue As String
Dim i As Integer
    If KeyAscii >= 32 And KeyAscii <> 127 Then
    NewText = LCase(Left(cboUser.Text, cboUser.SelStart) + Chr(KeyAscii) + Mid(cboUser.Text, cboUser.SelStart + cboUser.SelLength + 1))
 
    ValidCount = 0
    ValidValue = ""
 
    For i = 0 To cboUser.ListCount - 1
 
        If NewText = LCase(Left(cboUser.List(i), Len(NewText))) Then
                ValidCount = ValidCount + 1
                ValidValue = cboUser.List(i)
        End If
    Next
    If ValidCount <= 1 Then KeyAscii = 0
        If ValidCount = 1 Then
            cboUser.Text = ValidValue
            cboUser.SelStart = 0
            cboUser.SelLength = Len(ValidValue)
        End If
    End If
End Sub

Private Sub CmdClear_Click()
    
    If SSTab1.Tab = 1 Then
        'TxtstrCasNo.SetFocus
    End If
    
    If ChkScan.Value = "0" Then
        Call ClearForm(Me)
    End If
    Call EntriesEnable(True)
    Current2 = 0
    Check = False
    Check2 = False
    lblTotal = ""
    Lstcrt.ListItems.Clear
    Lststr.ListItems.Clear
    Call ClearStoreInfo
    txtTotalRecord = 0
End Sub


Private Sub cmdExit_Click()
'Dim RecordExit
'    RecordExit = MsgBox("Quit this form?", vbYesNo + vbExclamation)
'    If RecordExit = vbYes Then
'        bExit = True
        Unload Me
'    End If
End Sub
Private Function StorageSearchRecord()
Dim strsql As String
Dim strAppend As String
Dim sql As String
Dim storage As New ADODB.Recordset

    strAppend = ""
    
    If ChkScan.Value = "1" Then
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
    
'    If Len(Trim(TxtstrClmNo)) Then
'        strappend = strappend + " AND strclmversion = '" & RemStr(Trim(TxtstrClmNo)) & "'"
'    End If
    
    If Len(Trim(CboCarton)) Then
        strAppend = strAppend + " AND STRCartonNo >= '" & RemStr(Trim(CboCarton)) & "'"
    End If
    
    If Len(Trim(CboCartonto)) Then
        strAppend = strAppend + " AND STRCartonNo <= '" & RemStr(Trim(CboCartonto)) & "'"
    End If
    
    
    If Len(Trim(Mid(CboFileStatus, 1, 1))) Then
        strAppend = strAppend + " AND STRFileStatus = '" & RemStr(Trim(Mid(CboFileStatus, 1, 1))) & "'"
    End If
    
    If Len(Trim(cboUser)) Then
        strAppend = strAppend + " AND STRRequestBy = '" & RemStr(Trim(cboUser)) & "'"
    End If
    
    If IsDate(MaskstrDORqst) = True Then
        strAppend = strAppend + " And DORequest >= '" & Format(MaskstrDORqst, "mm/dd/yyyy") & "'"
    End If
    
    If IsDate(MaskstrDODelivery) = True Then
        strAppend = strAppend + " And DODelivery >= '" & Format(MaskstrDODelivery, "mm/dd/yyyy") & "'"
    End If
    
    If IsDate(MaskstrDOReturn) = True Then
        strAppend = strAppend + " And DOReturn >= '" & Format(MaskstrDOReturn, "mm/dd/yyyy") & "'"
    End If
    
    If IsDate(MaskstrDORqst2) = True Then
        strAppend = strAppend + " And DORequest <= '" & Format(MaskstrDORqst2, "mm/dd/yyyy") & "'"
    End If
    
    If IsDate(MaskstrDODelivery2) = True Then
        strAppend = strAppend + " And DODelivery <= '" & Format(MaskstrDODelivery2, "mm/dd/yyyy") & "'"
    End If
    
    If IsDate(MaskstrDOReturn2) = True Then
        strAppend = strAppend + " And DOReturn <= '" & Format(MaskstrDOReturn2, "mm/dd/yyyy") & "'"
    End If
    
    If IsDate(txtRegDateFr) = True Then
        strAppend = strAppend + " And LastUpdateDate >= '" & Format(txtRegDateFr, "mm/dd/yyyy") & "'"
    End If
    
    If IsDate(txtRegDateTo) = True Then
        strAppend = strAppend + " And LastUpdateDate <= '" & Format(txtRegDateTo, "mm/dd/yyyy") & "'"
    End If
    
    If IsDate(txtSentDateFr) = True Then
        strAppend = strAppend + " And DOSend >= '" & Format(txtSentDateFr, "yyyy/mm/dd") & "'"
    End If
    
    If IsDate(txtSentDateTo) = True Then
        strAppend = strAppend + " And DOSend <= '" & Format(txtSentDateTo, "yyyy/mm/dd") & "'"
    End If
    
    If Len(Trim(cboPayorFr)) Then
        strAppend = strAppend + " AND Substring(CASNumber,1,2) >= '" & Mid(cboPayorFr, 1, 2) & "'"
    End If
    
    If Len(Trim(cboPayorTo)) Then
        strAppend = strAppend + " AND Substring(CASNumber,1,2) <= '" & Mid(cboPayorTo, 1, 2) & "'"
    End If
    
    
    Call StorageListing(strAppend)
    
End Function

Private Sub StorageListing(Optional strAppend As String)
Dim rst2 As New ADODB.Recordset
Dim rstCount As New ADODB.Recordset
Dim Record As ListItem
Dim sql As String
Dim total As String
Dim Current As String
    
    total = 0
    Current = 0
    If ChkScan.Value = 0 Then
        Lststr.ListItems.Clear
    End If

    sql = " Select * from View_Storage_CLm where 0 = 0 and STRCartonNo IS NOT NULL "
        
    If Len(Trim(strAppend)) Then
        sql = sql + strAppend
    End If
    
    
    rst2.Open sql, medixmasterconnWS, adOpenKeyset, adLockOptimistic
    
    If rst2.recordCount = 0 Then
        MsgBox "No record found!", vbOKOnly + vbCritical, DialogBoxTitle
'        CmdAdd.Enabled = True
    Else
    
    Do
        Do Until rst2.EOF
        'total = rst2.recordCount
        'lblTotal = total
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
                If Len(!DOSend) Then
                    Record.SubItems(6) = Format(!DOSend, "dd/mm/yyyy")
                End If
                If Len(!LastUpdateDate) Then
                    Record.SubItems(7) = !LastUpdateDate
                End If
                If Len(!STRRequestBy) Then
                    Record.SubItems(8) = !STRRequestBy
                End If
                If Len(!DORequest) Then
                    Record.SubItems(9) = Format(!DORequest, "dd/mm/yyyy")
                End If
                If Len(!DODelivery) Then
                    Record.SubItems(10) = Format(!DODelivery, "dd/mm/yyyy")
                End If
                If Len(!DOReturn) Then
                    Record.SubItems(11) = Format(!DOReturn, "dd/mm/yyyy")
                End If
                If Len(!DOReturn) And Len(!DODelivery) Then
                    Record.SubItems(12) = DateValue(!DOReturn) - DateValue(!DODelivery)
                End If
                If Len(!STRRemarks) Then
                    Record.SubItems(13) = !STRRemarks
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
    
    CmdListing.Enabled = False
    'medixmasterconnWS.Open "File Name=" & BOSPath & "Conn\mediconnectWS.UDL"
    If SSTab1.Tab = 1 Then
        If Trim(CboCarton) = "" Then
            MsgBox "Please select Carton No. ", vbOKOnly + vbCritical, DialogBoxTitle
            CboCarton.SetFocus
        Else
    'On Error Resume Next
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
            ' Don't allow user to affect workbook
        
            strsql = "select STRCaseNo, STRClmVersion, STRFileStatus from Storage WHERE STRCartonNO  = '" & CboCarton & "'"
            Sto.Open strsql, medixmasterconnWS, adOpenKeyset, adLockOptimistic
            'Set objExcel = New excel.Application
            
            If Sto.recordCount <> 0 Then
                RecordSelected = Sto.recordCount
                objWorksheet.Cells(2, "F") = Format(Date, "MM/DD/YYYY")
                objWorksheet.Cells(2, "H") = CboCarton
                BB = 0
                yy = 0
                Cnt = 0
                ECol1 = Chr(65)
                ECol2 = Chr(66)
                ECol3 = Chr(67)
                
                'Do
                Do Until Sto.EOF
                'For i = 1 To RecordSelected
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
                    objWorksheet.Cells(4 + CDbl(BB), ECol2) = Sto!StrCaseNo & "-" & Sto!StrCLMVersion
                    objWorksheet.Cells(4 + CDbl(BB), ECol3) = Sto!STRFileStatus
                    
                    Sto.MoveNext
                
                'Next i
            Loop
                
            End If
                'cmdCredit.Enabled = False
                objExcel.DisplayAlerts = True
                'objWorksheet.Columns.AutoFit
                objExcel.Interactive = True
                Screen.MousePointer = vbDefault
                CmdListing.Enabled = True
                Exit Sub
ExitFunction:
                objExcel.DisplayAlerts = True
                objExcel.Interactive = True
                objExcel.Quit
              '  medixmasterconnWS.Close
              '  Set medixmasterconnWS = Nothing
                Screen.MousePointer = vbDefault
                CmdListing.Enabled = True
                Exit Sub
HANDLE_ERROR:
                MsgBox "Export To Excel failed. Encountered the following Error" & vbCrLf & vbCrLf & _
                Err.Number & ": " & Err.Description, vbOKOnly + vbCritical, "Error Exporting To Excel"
                Set objWorksheet = Nothing
                Set objWorkbook = Nothing
                GoTo ExitFunction
        End If
    End If
    'medixmasterconnWS.Close
   ' Set medixmasterconnWS = Nothing

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
            Screen.MousePointer = vbHourglass
        
            If Lststr.ListItems.Count <> 0 Then
                Call ExportListViewtoExcel(Lststr)
            Else
                MsgBox "Report cannot be printed without any record.", vbInformation
            End If
        
            Screen.MousePointer = vbDefault
        
            'Call ExportToExcelWorksheet(Lststr)
    End Select
End Sub


Private Sub cmdSearch_Click()
    Screen.MousePointer = vbHourglass
    
    CmdSearch.Enabled = False
    CmdExit.Enabled = False
    'medixmasterconnWS.Open "File Name=" & BOSPath & "Conn\mediconnectWS.UDL"
    Select Case SSTab1.Tab
        Case 0
            Call CartonSearchRecord
        Case 1
            
            If ChkScan.Value = 1 Then
    
                If TxtstrCasNo = "" Then
                    MsgBox "Please Enter Case No", vbOKOnly + vbCritical, DialogBoxTitle
                    TxtstrCasNo.SetFocus
                    CmdExit.Enabled = True
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
                    TxtstrCasNo = ""
                    TxtstrCasNo.SetFocus
                End If
                TxtstrCasNo.SetFocus
            Else
                Call StorageSearchRecord
            End If
    End Select
    'medixmasterconnWS.Close
   ' Set medixmasterconnWS = Nothing
    CmdSearch.Enabled = True
    Screen.MousePointer = Default
End Sub

Private Function CartonSearchRecord()
Dim strsql As String
Dim strAppend As String
Dim sql As String
Dim storage As New ADODB.Recordset

    strAppend = ""
    If Len(Trim(TxtcrtCartonNo)) Then
        strAppend = strAppend + " AND CRTCartonNo >= '" & RemStr(Trim(TxtcrtCartonNo)) & "'"
    End If
    
    If Len(Trim(TxtcrtCartonNo2)) Then
        strAppend = strAppend + " AND CRTCartonNo <= '" & RemStr(Trim(TxtcrtCartonNo2)) & "'"
    End If
    
    If Len(Trim(Mid(CboCartonStatus, 1, 1))) Then
        strAppend = strAppend + " AND CRTCartonStatus = '" & RemStr(Trim(Mid(CboCartonStatus, 1, 1))) & "'"
    End If
    
    If IsDate(MaskcrtDOSend) = True Then
        strAppend = strAppend + " And DOSend >= '" & Format(MaskcrtDOSend, "mm/dd/yyyy") & "'"
    End If
    
    If IsDate(MaskcrtDORqst) = True Then
        strAppend = strAppend + " And DORequest >= '" & Format(MaskcrtDORqst, "mm/dd/yyyy") & "'"
    End If
    
    If IsDate(MaskcrtDODelivery) = True Then
        strAppend = strAppend + " And DODelivery >= '" & Format(MaskcrtDODelivery, "mm/dd/yyyy") & "'"
    End If
    
    If IsDate(MaskcrtDOReturn) = True Then
        strAppend = strAppend + " And DOReturn >= '" & Format(MaskcrtDOReturn, "mm/dd/yyyy") & "'"
    End If
    
    If IsDate(MaskcrtDODestroy) = True Then
        strAppend = strAppend + " And DODestroy >= '" & Format(MaskcrtDODestroy, "mm/dd/yyyy") & "'"
    End If
    
    If IsDate(MaskcrtDOSend2) = True Then
        strAppend = strAppend + " And DOSend <= '" & Format(MaskcrtDOSend2, "mm/dd/yyyy") & "'"
    End If
    
    If IsDate(MaskcrtDORqst2) = True Then
        strAppend = strAppend + " And DORequest <= '" & Format(MaskcrtDORqst2, "mm/dd/yyyy") & "'"
    End If
    
    If IsDate(MaskcrtDODelivery2) = True Then
        strAppend = strAppend + " And DODelivery <= '" & Format(MaskcrtDODelivery2, "mm/dd/yyyy") & "'"
    End If
    
    If IsDate(MaskcrtDOReturn2) = True Then
        strAppend = strAppend + " And DOReturn <= '" & Format(MaskcrtDOReturn2, "mm/dd/yyyy") & "'"
    End If
    
    If IsDate(MaskcrtDODestroy2) = True Then
        strAppend = strAppend + " And DODestroy <= '" & Format(MaskcrtDODestroy2, "mm/dd/yyyy") & "'"
    End If
    
    Call CartonListing(strAppend)
End Function

Private Sub CartonListing(Optional strAppend As String)
Dim rst2 As New ADODB.Recordset
Dim rstCount As New ADODB.Recordset
Dim Record As ListItem
Dim sql As String
Dim total As String
Dim Current As String
    
    total = 0
    Current = 0
    
    Lstcrt.ListItems.Clear

    sql = " Select crtcartonno, crtcartonstatus, dosend, dorequest, " & _
          " dodelivery, doreturn, dodestroy, crtremarks, lastupdatedate from carton where 0 = 0 "
        
    If Len(Trim(strAppend)) Then
        sql = sql + strAppend
    End If
    
    rst2.Open sql, medixmasterconnWS, adOpenKeyset, adLockOptimistic
    
    If rst2.recordCount = 0 Then
        MsgBox "No record found!", vbOKOnly + vbCritical, DialogBoxTitle
    Else
    'total = rst2.recordCount
    'lblTotal = total
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
                If Len(rst2!DOSend) Then
                    Record.SubItems(3) = Format(rst2!DOSend, "dd/mm/yyyy")
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
    
    If SSTab1.Tab = 1 Then
        TxtstrCasNo.SetFocus
    End If

End Sub


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
    intExit = 0
    Current2 = 0
    Check = False
    Check2 = False
  '  medixmasterconnWS.Open "File Name=" & BOSPath & "Conn\mediconnectWS.UDL"
  '  medixmasterconnMaint.Open "File Name=" & BOSPath & "Conn\MediConnHISMaint.UDL"
  '  medixmasterconn.Open "File Name=" & BOSPath & "Conn\mediconnect.UDL"
    
    Call ComboListing
  '  medixmasterconnWS.Close
  '  Set medixmasterconnWS = Nothing
  '  medixmasterconnMaint.Close
  ''  Set medixmasterconnMaint = Nothing
   ' medixmasterconn.Close
   ' Set medixmasterconn = Nothing
    
    If SSTab1.Tab = 1 Then
        CmdListing.Visible = True
        Lststr.ListItems.Clear
        Label15.Visible = False
        TxtstrCasNo2.Visible = False
        'TxtstrCasNo.SetFocus
    Else
        CmdListing.Visible = False
        
    End If
End Sub

Private Sub Lstcrt_Click()
If Lstcrt.ListItems.Count <> 0 Then
        TxtcrtCartonNo = Lstcrt.SelectedItem.Text
            CboCartonStatus = Lstcrt.SelectedItem.SubItems(1)
        If Len(Trim(Lstcrt.SelectedItem.SubItems(7))) Then
            MaskcrtDODestroy = Lstcrt.SelectedItem.SubItems(7)
        End If
        If Len(Trim(Lstcrt.SelectedItem.SubItems(4))) Then
            MaskcrtDORqst = Lstcrt.SelectedItem.SubItems(4)
        End If
        If Len(Trim(Lstcrt.SelectedItem.SubItems(3))) Then
            MaskcrtDOSend = Lstcrt.SelectedItem.SubItems(3)
        End If
        If Len(Trim(Lstcrt.SelectedItem.SubItems(5))) Then
            MaskcrtDODelivery = Lstcrt.SelectedItem.SubItems(5)
        End If
        If Len(Trim(Lstcrt.SelectedItem.SubItems(6))) Then
            MaskcrtDOReturn = Lstcrt.SelectedItem.SubItems(6)
        End If
        If Len(Trim(Lstcrt.SelectedItem.SubItems(8))) Then
            TxtcrtRemarks = Lstcrt.SelectedItem.SubItems(8)
        End If
        
    End If
End Sub

Private Sub Lstcrt_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
    Lstcrt.SortKey = ColumnHeader.Index - 1
    Lstcrt.Sorted = True
    If Lstcrt.SortOrder = lvwAscending Then
        Lstcrt.SortOrder = lvwDescending
    Else
        Lstcrt.SortOrder = lvwAscending
    End If
End Sub

Private Sub Lststr_Click()
Dim rst3 As New ADODB.Recordset
Dim strsql As String
Dim CaseNo As String
    
    If Lststr.ListItems.Count <> 0 Then
        If Trim(Lststr.SelectedItem.SubItems(2)) <> "" Then
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

Private Sub SSTab1_Click(PreviousTab As Integer)
 Screen.MousePointer = vbHourglass
    Select Case SSTab1.Tab
    Case 0
         Lstcrt.ListItems.Clear
         CmdListing.Visible = False
    Case 1
         CmdListing.Visible = True
         Lststr.ListItems.Clear
         Call ClearStoreInfo
         TxtstrCasNo.SetFocus
         If ChkScan.Value = 1 Then
            ChkScan.Left = "3120"
            Label15.Visible = False
            TxtstrCasNo2.Visible = False
        Else
            Label15.Visible = True
            TxtstrCasNo2.Visible = True
            TxtstrCasNo.SetFocus
        End If
    End Select
    Screen.MousePointer = Default
End Sub

Private Sub TxtcrtCartonNo_LostFocus()
    TxtcrtCartonNo = ChkStr(TxtcrtCartonNo)
End Sub

Private Sub TxtcrtRemarks_LostFocus()
    TxtcrtRemarks = ChkStr(TxtcrtRemarks)
End Sub

Private Sub TxtstrCasNo_LostFocus()
    TxtstrCasNo = ChkStr(TxtstrCasNo)
End Sub

Private Sub TxtstrClmNo_KeyPress(KeyAscii As Integer)
    KeyAscii = KeyNumber(KeyAscii)
End Sub

Private Sub EntriesEnable(status As Boolean)
    TxtstrCasNo.Enabled = status
'    TxtstrClmNo.Enabled = status
    CboCarton.Enabled = status
    CboFileStatus.Enabled = status
    cboUser.Enabled = status
    MaskstrDORqst.Enabled = status
    MaskstrDODelivery.Enabled = status
    MaskstrDOReturn.Enabled = status
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
Dim PAY As New ADODB.Recordset

    Set cmd.ActiveConnection = medixmasterconnWS
    cmd.CommandText = "SearchCarton"
    cmd.CommandType = adCmdStoredProc
    cmd.CommandTimeout = 15
    CRTRec.CursorLocation = adUseClient
    CRTRec.CursorType = adOpenForwardOnly
    CRTRec.CacheSize = 100
    Set CRTRec = cmd.Execute
    
    CboCarton.Clear
    Do Until CRTRec.EOF
        With CRTRec
            CboCarton.AddItem !CRTCartonNo
            CboCartonto.AddItem !CRTCartonNo
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
    cboUser.Clear
    
    Do Until USRRec.EOF
        With USRRec
            cboUser.AddItem IIf(Len(!USRName), !USRName, "")
            USRRec.MoveNext
        End With
    Loop
    USRRec.Close
    Set USRRec = Nothing
      

    Set cmd.ActiveConnection = medixmasterconn
    cmd.CommandText = "SearchPayor"
    cmd.CommandType = adCmdStoredProc
    cmd.CommandTimeout = 15
    PAY.CursorLocation = adUseClient
    PAY.CursorType = adOpenForwardOnly
    PAY.CacheSize = 100
    Set PAY = cmd.Execute
    
    Do Until PAY.EOF
        With PAY
            cboPayorFr.AddItem !PAYCode & " - " & !PAYCompanyName
            cboPayorTo.AddItem !PAYCode & " - " & !PAYCompanyName
            PAY.MoveNext
        End With
    Loop
    PAY.Close
    Set PAY = Nothing
      
    CboFileStatus.Clear
    CboFileStatus.AddItem "I - In"
    CboFileStatus.AddItem "O - Out"
    CboCartonStatus.Clear
    CboCartonStatus.AddItem "M - Medix"
    CboCartonStatus.AddItem "S - Store"
    CboCartonStatus.AddItem "R - Return"
End Sub
    
Private Sub ChkScan_Click()
    If ChkScan.Value = 1 Then
        ChkScan.Left = "3120"
        Lststr.ListItems.Clear
        Label15.Visible = False
        TxtstrCasNo2.Visible = False
        TxtstrCasNo.SetFocus
    Else
        ChkScan.Left = "5900"
        Lststr.ListItems.Clear
        Label15.Visible = True
        TxtstrCasNo2.Visible = True
        TxtstrCasNo.SetFocus
    End If
End Sub

Private Function CaseNoCheck()
    If Trim(ChkStr(TxtstrCasNo)) = ChkStr(Trim(Lststr.ListItems(listloop1).Text)) Then
        Check = True
    End If
End Function

Private Sub ClearStoreInfo()
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
