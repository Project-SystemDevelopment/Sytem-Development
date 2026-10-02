VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "tabctl32.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "mscomctl.ocx"
Begin VB.Form frmInvoice 
   Caption         =   "Invoice Processing"
   ClientHeight    =   7920
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   11880
   KeyPreview      =   -1  'True
   MDIChild        =   -1  'True
   ScaleHeight     =   7920
   ScaleWidth      =   11880
   WindowState     =   2  'Maximized
   Begin VB.TextBox Text2 
      Alignment       =   1  'Right Justify
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
      MaxLength       =   8
      TabIndex        =   141
      Text            =   "0"
      Top             =   8040
      Visible         =   0   'False
      Width           =   915
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  'Right Justify
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
      Left            =   120
      MaxLength       =   8
      TabIndex        =   140
      Text            =   "0"
      Top             =   8040
      Visible         =   0   'False
      Width           =   915
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   5310
      Left            =   120
      TabIndex        =   123
      Top             =   2070
      Width           =   11655
      _ExtentX        =   20558
      _ExtentY        =   9366
      _Version        =   393216
      Tabs            =   5
      TabsPerRow      =   5
      TabHeight       =   441
      TabCaption(0)   =   "Invoice Entry"
      TabPicture(0)   =   "frmInvoice.frx":0000
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "FrmeOthers"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "FrmeNS"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "FrmeSurgical"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "lstInvoice"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "FrmeHos"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "FrmeOP"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).ControlCount=   6
      TabCaption(1)   =   "Reimbursement Not Cover"
      TabPicture(1)   =   "frmInvoice.frx":001C
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "FrmeNotCover"
      Tab(1).ControlCount=   1
      TabCaption(2)   =   "Remarks"
      TabPicture(2)   =   "frmInvoice.frx":0038
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "Label99"
      Tab(2).Control(1)=   "txtRemarks"
      Tab(2).ControlCount=   2
      TabCaption(3)   =   "General Info"
      TabPicture(3)   =   "frmInvoice.frx":0054
      Tab(3).ControlEnabled=   0   'False
      Tab(3).ControlCount=   0
      TabCaption(4)   =   "Documents"
      TabPicture(4)   =   "frmInvoice.frx":0070
      Tab(4).ControlEnabled=   0   'False
      Tab(4).Control(0)=   "cmdViewDoc"
      Tab(4).Control(1)=   "Frame1(1)"
      Tab(4).Control(2)=   "List1"
      Tab(4).ControlCount=   3
      Begin VB.ListBox List1 
         Height          =   4740
         ItemData        =   "frmInvoice.frx":008C
         Left            =   -74880
         List            =   "frmInvoice.frx":008E
         TabIndex        =   262
         ToolTipText     =   "right click to randomize contents"
         Top             =   360
         Width           =   4455
      End
      Begin VB.Frame Frame1 
         Appearance      =   0  'Flat
         BackColor       =   &H8000000A&
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   1
         Left            =   -74640
         TabIndex        =   238
         Top             =   3960
         Visible         =   0   'False
         Width           =   840
         Begin VB.PictureBox Picture1 
            Height          =   975
            Index           =   0
            Left            =   120
            ScaleHeight     =   915
            ScaleWidth      =   960
            TabIndex        =   248
            Top             =   240
            Visible         =   0   'False
            Width           =   1020
         End
         Begin VB.PictureBox Picture1 
            Height          =   975
            Index           =   1
            Left            =   1200
            ScaleHeight     =   915
            ScaleWidth      =   960
            TabIndex        =   247
            Top             =   240
            Visible         =   0   'False
            Width           =   1020
         End
         Begin VB.PictureBox Picture1 
            Height          =   975
            Index           =   2
            Left            =   2280
            ScaleHeight     =   915
            ScaleWidth      =   960
            TabIndex        =   246
            Top             =   240
            Visible         =   0   'False
            Width           =   1020
         End
         Begin VB.PictureBox Picture1 
            Height          =   975
            Index           =   3
            Left            =   3360
            ScaleHeight     =   915
            ScaleWidth      =   960
            TabIndex        =   245
            Top             =   240
            Visible         =   0   'False
            Width           =   1020
         End
         Begin VB.PictureBox Picture1 
            Height          =   975
            Index           =   4
            Left            =   4440
            ScaleHeight     =   915
            ScaleWidth      =   960
            TabIndex        =   244
            Top             =   240
            Visible         =   0   'False
            Width           =   1020
         End
         Begin VB.PictureBox Picture1 
            Height          =   975
            Index           =   5
            Left            =   5520
            ScaleHeight     =   915
            ScaleWidth      =   960
            TabIndex        =   243
            Top             =   240
            Visible         =   0   'False
            Width           =   1020
         End
         Begin VB.PictureBox Picture1 
            Height          =   975
            Index           =   9
            Left            =   9840
            ScaleHeight     =   915
            ScaleWidth      =   960
            TabIndex        =   242
            Top             =   240
            Visible         =   0   'False
            Width           =   1020
         End
         Begin VB.PictureBox Picture1 
            Height          =   975
            Index           =   8
            Left            =   8760
            ScaleHeight     =   915
            ScaleWidth      =   960
            TabIndex        =   241
            Top             =   240
            Visible         =   0   'False
            Width           =   1020
         End
         Begin VB.PictureBox Picture1 
            Height          =   975
            Index           =   7
            Left            =   7680
            ScaleHeight     =   915
            ScaleWidth      =   960
            TabIndex        =   240
            Top             =   240
            Visible         =   0   'False
            Width           =   1020
         End
         Begin VB.PictureBox Picture1 
            Height          =   975
            Index           =   6
            Left            =   6600
            ScaleHeight     =   915
            ScaleWidth      =   960
            TabIndex        =   239
            Top             =   240
            Visible         =   0   'False
            Width           =   1020
         End
         Begin VB.Label Label1 
            Alignment       =   2  'Center
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   12
               Charset         =   204
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000FF&
            Height          =   375
            Index           =   21
            Left            =   4440
            TabIndex        =   259
            Top             =   720
            Visible         =   0   'False
            Width           =   2580
         End
         Begin VB.Label Label1 
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   6.75
               Charset         =   204
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   225
            Index           =   1
            Left            =   140
            TabIndex        =   258
            Top             =   1200
            Visible         =   0   'False
            Width           =   1020
         End
         Begin VB.Label Label1 
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   6.75
               Charset         =   204
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   225
            Index           =   78
            Left            =   6610
            TabIndex        =   257
            Top             =   1200
            Visible         =   0   'False
            Width           =   1020
         End
         Begin VB.Label Label1 
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   6.75
               Charset         =   204
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   225
            Index           =   77
            Left            =   7690
            TabIndex        =   256
            Top             =   1200
            Visible         =   0   'False
            Width           =   1020
         End
         Begin VB.Label Label1 
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   6.75
               Charset         =   204
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   225
            Index           =   67
            Left            =   8770
            TabIndex        =   255
            Top             =   1200
            Visible         =   0   'False
            Width           =   1020
         End
         Begin VB.Label Label1 
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   6.75
               Charset         =   204
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   225
            Index           =   22
            Left            =   9850
            TabIndex        =   254
            Top             =   1200
            Visible         =   0   'False
            Width           =   1020
         End
         Begin VB.Label Label1 
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   6.75
               Charset         =   204
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   225
            Index           =   18
            Left            =   5530
            TabIndex        =   253
            Top             =   1200
            Visible         =   0   'False
            Width           =   1020
         End
         Begin VB.Label Label1 
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   6.75
               Charset         =   204
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   225
            Index           =   16
            Left            =   4450
            TabIndex        =   252
            Top             =   1200
            Visible         =   0   'False
            Width           =   1020
         End
         Begin VB.Label Label1 
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   6.75
               Charset         =   204
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   225
            Index           =   9
            Left            =   3370
            TabIndex        =   251
            Top             =   1200
            Visible         =   0   'False
            Width           =   1020
         End
         Begin VB.Label Label1 
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   6.75
               Charset         =   204
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   225
            Index           =   3
            Left            =   2290
            TabIndex        =   250
            Top             =   1200
            Visible         =   0   'False
            Width           =   1020
         End
         Begin VB.Label Label1 
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   6.75
               Charset         =   204
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   225
            Index           =   2
            Left            =   1210
            TabIndex        =   249
            Top             =   1200
            Visible         =   0   'False
            Width           =   1020
         End
      End
      Begin VB.Frame FrmeOP 
         Height          =   1710
         Left            =   6300
         TabIndex        =   208
         Top             =   285
         Width           =   5145
         Begin VB.TextBox A11 
            Alignment       =   1  'Right Justify
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
            Left            =   4080
            MaxLength       =   8
            TabIndex        =   67
            Top             =   120
            Width           =   795
         End
         Begin VB.TextBox txtKidneyMonths 
            Height          =   285
            Left            =   1920
            TabIndex        =   70
            Top             =   860
            Width           =   735
         End
         Begin VB.TextBox txtPTreat 
            Alignment       =   1  'Right Justify
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
            Left            =   3000
            MaxLength       =   8
            TabIndex        =   66
            Top             =   120
            Width           =   915
         End
         Begin VB.TextBox txtACC 
            Alignment       =   1  'Right Justify
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
            MaxLength       =   8
            TabIndex        =   65
            Top             =   120
            Width           =   675
         End
         Begin VB.TextBox A12 
            Alignment       =   1  'Right Justify
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
            Left            =   4080
            MaxLength       =   8
            TabIndex        =   68
            Top             =   350
            Width           =   795
         End
         Begin VB.TextBox A13 
            Alignment       =   1  'Right Justify
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
            Left            =   4080
            MaxLength       =   8
            TabIndex        =   69
            Top             =   600
            Width           =   795
         End
         Begin VB.TextBox A14 
            Alignment       =   1  'Right Justify
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
            Left            =   4080
            MaxLength       =   8
            TabIndex        =   71
            Top             =   830
            Width           =   795
         End
         Begin VB.TextBox A27 
            Alignment       =   1  'Right Justify
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
            Left            =   4080
            MaxLength       =   8
            TabIndex        =   72
            Top             =   1080
            Width           =   795
         End
         Begin VB.TextBox A26 
            Alignment       =   1  'Right Justify
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
            Left            =   4080
            MaxLength       =   8
            TabIndex        =   73
            Top             =   1290
            Width           =   795
         End
         Begin VB.Label Label22 
            Caption         =   "Cancer Treatment"
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
            Left            =   120
            TabIndex        =   290
            Top             =   1095
            Width           =   2055
         End
         Begin VB.Label Label22 
            Caption         =   "Kidney Dialysis"
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
            Left            =   120
            TabIndex        =   215
            Top             =   850
            Width           =   2055
         End
         Begin VB.Label Label21 
            Caption         =   "Outpatient Physiotherapy Treatment"
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
            Left            =   120
            TabIndex        =   214
            Top             =   630
            Width           =   2805
         End
         Begin VB.Label Label20 
            Caption         =   "Ambulance Fee"
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
            Left            =   120
            TabIndex        =   213
            Top             =   410
            Width           =   2085
         End
         Begin VB.Label Label8 
            Caption         =   "Accident"
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
            Left            =   120
            TabIndex        =   212
            Top             =   180
            Width           =   735
         End
         Begin VB.Label Label68 
            Caption         =   "Month(s)"
            ForeColor       =   &H00FF0000&
            Height          =   255
            Left            =   2760
            TabIndex        =   211
            Top             =   895
            Width           =   735
         End
         Begin VB.Label Label84 
            Caption         =   "Post-Treat"
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
            Left            =   2160
            TabIndex        =   210
            Top             =   195
            Width           =   735
         End
         Begin VB.Label Label103 
            Caption         =   "Home Nursing Care"
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
            Left            =   120
            TabIndex        =   209
            Top             =   1320
            Width           =   2805
         End
      End
      Begin VB.Frame FrmeHos 
         Height          =   1710
         Left            =   120
         TabIndex        =   167
         Top             =   285
         Width           =   6015
         Begin VB.ComboBox cboRoomType 
            Height          =   315
            Left            =   960
            TabIndex        =   292
            Top             =   1300
            Width           =   1635
         End
         Begin VB.CommandButton cmdHSS 
            Appearance      =   0  'Flat
            Caption         =   ">"
            Height          =   285
            Left            =   2280
            TabIndex        =   283
            Top             =   960
            Width           =   330
         End
         Begin VB.TextBox A4 
            Alignment       =   1  'Right Justify
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
            Left            =   4320
            MaxLength       =   8
            TabIndex        =   37
            Top             =   960
            Width           =   1635
         End
         Begin VB.TextBox A23 
            Alignment       =   1  'Right Justify
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
            Left            =   3000
            MaxLength       =   8
            TabIndex        =   36
            Top             =   960
            Width           =   1035
         End
         Begin VB.TextBox A3 
            Alignment       =   1  'Right Justify
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
            Left            =   960
            MaxLength       =   10
            TabIndex        =   35
            Top             =   960
            Width           =   1275
         End
         Begin VB.TextBox AIT 
            Alignment       =   1  'Right Justify
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
            Height          =   240
            Left            =   3000
            MaxLength       =   8
            TabIndex        =   32
            Top             =   645
            Width           =   555
         End
         Begin VB.TextBox ICU1 
            Alignment       =   1  'Right Justify
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   240
            Left            =   960
            MaxLength       =   5
            TabIndex        =   30
            Top             =   645
            Width           =   675
         End
         Begin VB.TextBox ICUAmount 
            Alignment       =   1  'Right Justify
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   240
            Left            =   1920
            MaxLength       =   8
            TabIndex        =   31
            Top             =   645
            Width           =   675
         End
         Begin VB.TextBox A2 
            Alignment       =   1  'Right Justify
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   240
            Left            =   3960
            MaxLength       =   8
            TabIndex        =   33
            Top             =   645
            Width           =   795
         End
         Begin VB.TextBox A1 
            Alignment       =   1  'Right Justify
            BeginProperty DataFormat 
               Type            =   1
               Format          =   "0.00"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1033
               SubFormatType   =   1
            EndProperty
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
            MaxLength       =   8
            TabIndex        =   27
            Top             =   180
            Width           =   795
         End
         Begin VB.TextBox RoomAmount 
            Alignment       =   1  'Right Justify
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
            MaxLength       =   8
            TabIndex        =   25
            Top             =   180
            Width           =   675
         End
         Begin VB.TextBox DB1 
            Alignment       =   1  'Right Justify
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
            MaxLength       =   5
            TabIndex        =   24
            Top             =   180
            Width           =   675
         End
         Begin VB.TextBox ART 
            Alignment       =   1  'Right Justify
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
            Height          =   255
            Left            =   3000
            MaxLength       =   8
            TabIndex        =   26
            Top             =   180
            Width           =   555
         End
         Begin VB.TextBox A1B 
            Alignment       =   1  'Right Justify
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
            Height          =   255
            Left            =   5040
            MaxLength       =   8
            TabIndex        =   28
            Top             =   180
            Width           =   915
         End
         Begin VB.TextBox txtGrossRB 
            Alignment       =   1  'Right Justify
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
            Height          =   240
            Left            =   5040
            MaxLength       =   8
            TabIndex        =   29
            Top             =   400
            Width           =   915
         End
         Begin VB.TextBox A2B 
            Alignment       =   1  'Right Justify
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
            Left            =   5040
            MaxLength       =   8
            TabIndex        =   34
            Top             =   620
            Width           =   915
         End
         Begin VB.Label Label11 
            Caption         =   "Room Type"
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
            Left            =   120
            TabIndex        =   293
            Top             =   1320
            Width           =   885
         End
         Begin VB.Label Label28 
            Caption         =   "Lmt"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00C00000&
            Height          =   255
            Index           =   2
            Left            =   1680
            TabIndex        =   286
            Top             =   675
            Width           =   465
         End
         Begin VB.Label Label28 
            Caption         =   "Lmt"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00C00000&
            Height          =   255
            Index           =   1
            Left            =   1680
            TabIndex        =   285
            Top             =   195
            Width           =   465
         End
         Begin VB.Label Label19 
            Caption         =   "MRI"
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
            Index           =   21
            Left            =   2640
            TabIndex        =   284
            Top             =   1005
            Width           =   1005
         End
         Begin VB.Label Label26 
            Caption         =   "Days"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FF0000&
            Height          =   255
            Left            =   600
            TabIndex        =   179
            Top             =   675
            Width           =   465
         End
         Begin VB.Label Label18 
            Caption         =   "R && B"
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
            Left            =   120
            TabIndex        =   178
            Top             =   195
            Width           =   435
         End
         Begin VB.Label Label10 
            Caption         =   "ICU"
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
            Left            =   120
            TabIndex        =   177
            ToolTipText     =   "Intensive Care Unit"
            Top             =   675
            Width           =   1125
         End
         Begin VB.Label Label11 
            Caption         =   "HSS"
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
            Left            =   120
            TabIndex        =   176
            Top             =   960
            Width           =   2445
         End
         Begin VB.Label Label12 
            Caption         =   "OT"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   180
            Left            =   4080
            TabIndex        =   175
            Top             =   1005
            Width           =   495
         End
         Begin VB.Label Label28 
            Caption         =   "Days"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00C00000&
            Height          =   255
            Index           =   0
            Left            =   600
            TabIndex        =   174
            Top             =   195
            Width           =   465
         End
         Begin VB.Label Label78 
            Caption         =   "MRI"
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
            Left            =   2280
            TabIndex        =   173
            Top             =   960
            Width           =   495
         End
         Begin VB.Label Label85 
            Caption         =   "Tax"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00C00000&
            Height          =   255
            Left            =   2640
            TabIndex        =   172
            Top             =   195
            Width           =   465
         End
         Begin VB.Label Label86 
            Caption         =   "Tax"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FF0000&
            Height          =   255
            Left            =   2640
            TabIndex        =   171
            Top             =   675
            Width           =   825
         End
         Begin VB.Label Label88 
            Caption         =   "Amt"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FF0000&
            Height          =   255
            Left            =   3600
            TabIndex        =   170
            Top             =   675
            Width           =   825
         End
         Begin VB.Label Label89 
            Caption         =   "Amt"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00C00000&
            Height          =   255
            Left            =   3600
            TabIndex        =   169
            Top             =   180
            Width           =   465
         End
         Begin VB.Label Label98 
            Caption         =   "Gross RB"
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
            Left            =   120
            TabIndex        =   168
            Top             =   450
            Width           =   675
         End
      End
      Begin VB.Frame Frame1 
         Height          =   2055
         Index           =   0
         Left            =   -74880
         TabIndex        =   144
         Top             =   360
         Width           =   8655
         Begin VB.TextBox txtMBMNumber 
            Enabled         =   0   'False
            Height          =   285
            Left            =   1320
            MaxLength       =   25
            TabIndex        =   151
            Top             =   240
            Width           =   2625
         End
         Begin VB.TextBox txtMBMPolNo 
            Enabled         =   0   'False
            Height          =   285
            Left            =   4920
            MaxLength       =   25
            TabIndex        =   150
            Top             =   240
            Width           =   2205
         End
         Begin VB.TextBox txtMBMName 
            Enabled         =   0   'False
            Height          =   285
            Left            =   1320
            MaxLength       =   50
            TabIndex        =   149
            Top             =   600
            Width           =   5805
         End
         Begin VB.TextBox txtEffDate 
            Enabled         =   0   'False
            Height          =   285
            Left            =   1320
            TabIndex        =   148
            Top             =   960
            Width           =   2625
         End
         Begin VB.TextBox txtExpDate 
            Enabled         =   0   'False
            Height          =   285
            Left            =   4920
            TabIndex        =   146
            Top             =   960
            Width           =   2205
         End
         Begin VB.TextBox txtDOB 
            Enabled         =   0   'False
            Height          =   285
            Left            =   4920
            TabIndex        =   145
            Top             =   1320
            Width           =   2200
         End
         Begin VB.TextBox txtHospital 
            Enabled         =   0   'False
            Height          =   285
            Left            =   1320
            TabIndex        =   147
            Top             =   1695
            Width           =   5800
         End
         Begin VB.Label txtPolStatus 
            Alignment       =   2  'Center
            BackColor       =   &H00C0FFFF&
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000FF&
            Height          =   285
            Left            =   1320
            TabIndex        =   160
            Top             =   1320
            Width           =   2625
         End
         Begin VB.Label Label16 
            Caption         =   "Mem Name"
            Height          =   255
            Left            =   120
            TabIndex        =   159
            Top             =   600
            Width           =   1140
         End
         Begin VB.Label Label76 
            Caption         =   "Mem Num"
            Height          =   255
            Left            =   120
            TabIndex        =   158
            Top             =   240
            Width           =   1065
         End
         Begin VB.Label Label73 
            Caption         =   "Policy Num"
            Height          =   255
            Left            =   4080
            TabIndex        =   157
            Top             =   240
            Width           =   1020
         End
         Begin VB.Label Label55 
            Caption         =   "Pol Status"
            Height          =   255
            Left            =   120
            TabIndex        =   156
            Top             =   1320
            Width           =   915
         End
         Begin VB.Label Label59 
            Caption         =   "Eff Date"
            Height          =   255
            Left            =   120
            TabIndex        =   155
            Top             =   960
            Width           =   855
         End
         Begin VB.Label Label60 
            Caption         =   "Exp Date"
            Height          =   210
            Left            =   4005
            TabIndex        =   154
            Top             =   990
            Width           =   735
         End
         Begin VB.Label Label61 
            Caption         =   "Patient DOB"
            Height          =   255
            Left            =   4005
            TabIndex        =   153
            Top             =   1320
            Width           =   1095
         End
         Begin VB.Label Label70 
            Caption         =   "Hospital"
            Height          =   255
            Left            =   120
            TabIndex        =   152
            Top             =   1710
            Width           =   855
         End
      End
      Begin VB.Frame FrmeNotCover 
         Caption         =   "Reimbursement Not-Payable"
         Height          =   1395
         Left            =   -74760
         TabIndex        =   131
         Top             =   360
         Width           =   10455
         Begin VB.TextBox txtOthReimNotCovered 
            Height          =   285
            Left            =   6720
            MaxLength       =   80
            TabIndex        =   105
            Top             =   960
            Width           =   2505
         End
         Begin VB.TextBox txtPreSpec1 
            Alignment       =   1  'Right Justify
            Height          =   285
            Left            =   9360
            MaxLength       =   8
            TabIndex        =   102
            Top             =   225
            Width           =   915
         End
         Begin VB.TextBox txtPreHos60 
            Alignment       =   1  'Right Justify
            Height          =   285
            Left            =   3600
            MaxLength       =   8
            TabIndex        =   98
            Top             =   240
            Width           =   915
         End
         Begin VB.TextBox txtPreGP 
            Alignment       =   1  'Right Justify
            Height          =   285
            Left            =   3600
            MaxLength       =   8
            TabIndex        =   99
            Top             =   480
            Width           =   915
         End
         Begin VB.TextBox txtPreMEdic 
            Alignment       =   1  'Right Justify
            Height          =   285
            Left            =   3600
            MaxLength       =   8
            TabIndex        =   100
            Top             =   720
            Width           =   915
         End
         Begin VB.TextBox txtPreMR 
            Alignment       =   1  'Right Justify
            Height          =   285
            Left            =   3600
            MaxLength       =   8
            TabIndex        =   101
            Top             =   980
            Width           =   915
         End
         Begin VB.TextBox txtPostGHosp 
            Alignment       =   1  'Right Justify
            Height          =   285
            Left            =   9360
            MaxLength       =   8
            TabIndex        =   103
            Top             =   480
            Width           =   915
         End
         Begin VB.TextBox txtpostNotSame 
            Alignment       =   1  'Right Justify
            Height          =   285
            Left            =   9360
            MaxLength       =   8
            TabIndex        =   104
            Top             =   720
            Width           =   915
         End
         Begin VB.TextBox txtOthers 
            Alignment       =   1  'Right Justify
            Height          =   285
            Left            =   9360
            MaxLength       =   8
            TabIndex        =   106
            Top             =   960
            Width           =   915
         End
         Begin VB.Label Label7 
            Caption         =   "Others"
            Height          =   255
            Left            =   5160
            TabIndex        =   139
            Top             =   1035
            Width           =   1455
         End
         Begin VB.Label Label45 
            Caption         =   "Not Same Physician"
            Height          =   255
            Left            =   5160
            TabIndex        =   138
            Top             =   795
            Width           =   1575
         End
         Begin VB.Label Label44 
            Caption         =   "Post Hosp > 31 days"
            Height          =   255
            Left            =   5160
            TabIndex        =   137
            Top             =   555
            Width           =   1695
         End
         Begin VB.Label Label43 
            Caption         =   "Pre-hosp Specialist Fee > 1"
            Height          =   255
            Left            =   5160
            TabIndex        =   136
            Top             =   270
            Width           =   2055
         End
         Begin VB.Label Label42 
            Caption         =   "Medical Report"
            Height          =   255
            Left            =   240
            TabIndex        =   135
            ToolTipText     =   "Pre-hosp Medical Report"
            Top             =   1005
            Width           =   1455
         End
         Begin VB.Label Label41 
            Caption         =   "Pre - Medicine"
            Height          =   255
            Left            =   240
            TabIndex        =   134
            Top             =   735
            Width           =   1455
         End
         Begin VB.Label Label40 
            Caption         =   "GP fee"
            Height          =   255
            Left            =   240
            TabIndex        =   133
            Top             =   480
            Width           =   1215
         End
         Begin VB.Label Label39 
            Caption         =   "Pre-hosp > 60"
            Height          =   255
            Left            =   240
            TabIndex        =   132
            Top             =   240
            Width           =   1335
         End
      End
      Begin MSComctlLib.ListView lstInvoice 
         Height          =   810
         Left            =   120
         TabIndex        =   143
         Top             =   4440
         Width           =   11415
         _ExtentX        =   20135
         _ExtentY        =   1429
         View            =   3
         LabelEdit       =   1
         Sorted          =   -1  'True
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         AllowReorder    =   -1  'True
         FullRowSelect   =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         NumItems        =   118
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Index"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Bordx"
            Object.Width           =   1058
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "WS "
            Object.Width           =   1058
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   3
            Text            =   "WS No"
            Object.Width           =   883
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   4
            Text            =   "Inv No"
            Object.Width           =   2469
         EndProperty
         BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   5
            Text            =   "Inv Date"
            Object.Width           =   1764
         EndProperty
         BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   6
            Text            =   "INV Amount"
            Object.Width           =   2469
         EndProperty
         BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   7
            Text            =   "Pay To"
            Object.Width           =   1411
         EndProperty
         BeginProperty ColumnHeader(9) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   8
            Text            =   "Payee"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(10) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   9
            Text            =   "Claim Type"
            Object.Width           =   882
         EndProperty
         BeginProperty ColumnHeader(11) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   10
            Text            =   "Fol. Up Date"
            Object.Width           =   882
         EndProperty
         BeginProperty ColumnHeader(12) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   11
            Text            =   "DailyRB"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(13) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   12
            Text            =   "RB Rate"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(14) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   13
            Text            =   "ICUDay"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(15) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   14
            Text            =   "ICU Rate"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(16) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   15
            Text            =   "HSS"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(17) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   16
            Text            =   "OT"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(18) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   17
            Text            =   "PHService"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(19) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   18
            Text            =   "IHPysVisit"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(20) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   19
            Text            =   "PHTreatment"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(21) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   20
            Text            =   "MR"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(22) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   21
            Text            =   "Phone Charges"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(23) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   22
            Text            =   "LodgerDay"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(24) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   23
            Text            =   "Lodger Rate"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(25) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   24
            Text            =   "Admin"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(26) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   25
            Text            =   "Tax"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(27) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   26
            Text            =   "Others"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(28) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   27
            Text            =   "CashDay"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(29) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   28
            Text            =   "Cash Rate"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(30) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   29
            Text            =   "XRay"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(31) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   30
            Text            =   "Lab"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(32) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   31
            Text            =   "pre Others"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(33) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   32
            Text            =   "SFee Amount"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(34) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   33
            Text            =   "AFee"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(35) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   34
            Text            =   "EOut-Patient"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(36) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   35
            Text            =   "Ambfee"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(37) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   36
            Text            =   "OPT"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(38) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   37
            Text            =   "MonthlyKidney"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(39) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   38
            Text            =   "Pre60"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(40) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   39
            Text            =   "PreGP"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(41) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   40
            Text            =   "PreMedic"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(42) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   41
            Text            =   "PreMR "
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(43) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   42
            Text            =   "PreSpec"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(44) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   43
            Text            =   "Post_Host"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(45) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   44
            Text            =   "NotSame"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(46) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   45
            Text            =   "Others"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(47) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   46
            Text            =   "Index"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(48) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   47
            Text            =   "SurgicalPhy"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(49) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   48
            Text            =   "SurgicalPost"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(50) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   49
            Text            =   "Others1Desc"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(51) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   50
            Text            =   "Others2Desc"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(52) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   51
            Text            =   "Others2Value"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(53) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   52
            Text            =   "CLMCashAmt"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(54) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   53
            Text            =   "CLMGuardianAmt"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(55) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   54
            Text            =   "BoardAmt"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(56) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   55
            Text            =   "ICUAmnt"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(57) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   56
            Text            =   "KidneyMonth"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(58) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   57
            Text            =   "Surgicalfee"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(59) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   58
            Text            =   "PreNSHosp"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(60) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   59
            Text            =   "In Hospital Phy Visit"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(61) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   60
            Text            =   "Organ Transplant"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(62) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   61
            Text            =   "RBTax"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(63) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   62
            Text            =   "ICU Tax"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(64) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   63
            Text            =   "Lodger Tax"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(65) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   64
            Text            =   "StayType"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(66) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   65
            Text            =   "PayType"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(67) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   66
            Text            =   "Prepared By"
            Object.Width           =   1764
         EndProperty
         BeginProperty ColumnHeader(68) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   67
            Text            =   "Prepared Date"
            Object.Width           =   1764
         EndProperty
         BeginProperty ColumnHeader(69) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   68
            Text            =   "MRI"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(70) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   69
            Text            =   "MXRay"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(71) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   70
            Text            =   "MLab"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(72) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   71
            Text            =   "MPreOthers"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(73) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   72
            Text            =   "MPreHosAmt"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(74) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   73
            Text            =   "Pre-Hosp Specialist Fee"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(75) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   74
            Text            =   "Post Treat"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(76) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   75
            Text            =   "AccPostAmt"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(77) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   76
            Text            =   "NewRBTax"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(78) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   77
            Text            =   "NewICUTax"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(79) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   78
            Text            =   "NewLodgerTax"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(80) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   79
            Text            =   "TotalRB"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(81) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   80
            Text            =   "TotalICU"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(82) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   81
            Text            =   "TotalGAllow"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(83) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   82
            Text            =   "SurgWard"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(84) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   83
            Text            =   "AnaFees"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(85) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   84
            Text            =   "AnaCons"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(86) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   85
            Text            =   "AnaWard"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(87) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   86
            Text            =   "InHosFees"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(88) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   87
            Text            =   "InHosWard"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(89) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   88
            Text            =   "MedRptMBM"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(90) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   89
            Text            =   "TelChargesMBM"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(91) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   90
            Text            =   "GuaAllowMBM"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(92) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   91
            Text            =   "AdminRegMBM"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(93) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   92
            Text            =   "OthersMBM"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(94) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   93
            Text            =   "Gross RB"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(95) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   94
            Text            =   "OthersNoCoPay"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(96) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   95
            Text            =   "OthersNoCoPayDesc"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(97) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   96
            Text            =   "CLMProceduresFee"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(98) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   97
            Text            =   "DIR"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(99) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   98
            Text            =   "DRM"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(100) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   99
            Text            =   "HomeNC"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(101) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   100
            Text            =   "UHCCLMNo"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(102) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   101
            Text            =   "ExcAmt"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(103) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   102
            Text            =   "InvPayAmt"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(104) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   103
            Text            =   "ReimOthRemarks"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(105) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   104
            Text            =   "Pharmacy"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(106) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   105
            Text            =   "MedSupp"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(107) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   106
            Text            =   "NursingCare"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(108) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   107
            Text            =   "NursingSvr"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(109) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   108
            Text            =   "Lab"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(110) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   109
            Text            =   "Radio"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(111) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   110
            Text            =   "OT"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(112) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   111
            Text            =   "Equip"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(113) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   112
            Text            =   "Misc"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(114) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   113
            Text            =   "InitialActAmt"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(115) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   114
            Text            =   "HospName"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(116) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   115
            Text            =   "Cancer Treatment"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(117) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   116
            Text            =   "MaternityStatus"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(118) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   117
            Text            =   "RoomType"
            Object.Width           =   2540
         EndProperty
      End
      Begin VB.Frame FrmeSurgical 
         Height          =   1215
         Left            =   120
         TabIndex        =   180
         Top             =   1920
         Width           =   6015
         Begin VB.TextBox txtPhysician 
            Alignment       =   1  'Right Justify
            Height          =   285
            Left            =   2685
            MaxLength       =   8
            TabIndex        =   44
            Top             =   630
            Width           =   495
         End
         Begin VB.TextBox txtPostTreat 
            Alignment       =   1  'Right Justify
            Height          =   285
            Left            =   3645
            MaxLength       =   8
            TabIndex        =   45
            Top             =   630
            Width           =   540
         End
         Begin VB.TextBox txtSurgicalFee 
            Alignment       =   1  'Right Justify
            Enabled         =   0   'False
            Height          =   285
            Left            =   1320
            MaxLength       =   8
            TabIndex        =   43
            Top             =   600
            Width           =   495
         End
         Begin VB.TextBox txtPreOthers 
            Alignment       =   1  'Right Justify
            Height          =   285
            Left            =   3600
            MaxLength       =   8
            TabIndex        =   40
            Top             =   120
            Width           =   540
         End
         Begin VB.TextBox txtXray 
            Alignment       =   1  'Right Justify
            Height          =   285
            Left            =   1350
            TabIndex        =   38
            Top             =   140
            Width           =   495
         End
         Begin VB.TextBox A5 
            Alignment       =   1  'Right Justify
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
            Left            =   5160
            MaxLength       =   8
            TabIndex        =   41
            Top             =   120
            Width           =   795
         End
         Begin VB.TextBox txtPreLab 
            Alignment       =   1  'Right Justify
            Height          =   285
            Left            =   2685
            MaxLength       =   8
            TabIndex        =   39
            Top             =   120
            Width           =   495
         End
         Begin VB.TextBox txtSurgWard 
            Alignment       =   1  'Right Justify
            Height          =   285
            Left            =   4680
            MaxLength       =   8
            TabIndex        =   46
            Top             =   600
            Width           =   480
         End
         Begin VB.CommandButton cmdMMAEntry 
            Appearance      =   0  'Flat
            Caption         =   ">"
            Height          =   330
            Left            =   1880
            TabIndex        =   49
            Top             =   750
            Width           =   330
         End
         Begin VB.TextBox A24 
            Alignment       =   1  'Right Justify
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
            Left            =   5160
            MaxLength       =   8
            TabIndex        =   42
            Top             =   375
            Width           =   795
         End
         Begin VB.TextBox A6 
            Alignment       =   1  'Right Justify
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
            Left            =   5160
            MaxLength       =   8
            TabIndex        =   47
            Top             =   600
            Width           =   795
         End
         Begin VB.TextBox A7 
            Alignment       =   1  'Right Justify
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
            Left            =   5160
            MaxLength       =   8
            TabIndex        =   52
            Top             =   850
            Width           =   795
         End
         Begin VB.TextBox txtAnaCons 
            Alignment       =   1  'Right Justify
            Height          =   285
            Left            =   2685
            MaxLength       =   8
            TabIndex        =   50
            Top             =   880
            Width           =   495
         End
         Begin VB.TextBox txtAnaWard 
            Alignment       =   1  'Right Justify
            Height          =   285
            Left            =   3645
            MaxLength       =   8
            TabIndex        =   51
            Top             =   880
            Width           =   540
         End
         Begin VB.TextBox txtAnaFees 
            Alignment       =   1  'Right Justify
            Enabled         =   0   'False
            Height          =   285
            Left            =   1320
            MaxLength       =   8
            TabIndex        =   48
            Top             =   880
            Width           =   495
         End
         Begin VB.Label Label94 
            Caption         =   "Ward"
            ForeColor       =   &H00FF0000&
            Height          =   255
            Left            =   3240
            TabIndex        =   183
            Top             =   885
            Width           =   495
         End
         Begin VB.Label Label64 
            Caption         =   "Fee"
            ForeColor       =   &H00C00000&
            Height          =   255
            Left            =   890
            TabIndex        =   181
            Top             =   630
            Width           =   375
         End
         Begin VB.Label Label9 
            Caption         =   "Pre-Hosp"
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
            Left            =   120
            TabIndex        =   194
            Top             =   165
            Width           =   735
         End
         Begin VB.Label Label35 
            Caption         =   "X-Ray"
            ForeColor       =   &H00C00000&
            Height          =   255
            Left            =   890
            TabIndex        =   193
            Top             =   165
            Width           =   495
         End
         Begin VB.Label Label32 
            Caption         =   "Lab"
            ForeColor       =   &H00C00000&
            Height          =   255
            Left            =   2235
            TabIndex        =   192
            Top             =   165
            Width           =   375
         End
         Begin VB.Label Label27 
            Caption         =   "Oth"
            ForeColor       =   &H00C00000&
            Height          =   255
            Left            =   3240
            TabIndex        =   191
            Top             =   165
            Width           =   495
         End
         Begin VB.Label Label14 
            Caption         =   "Ana Fees"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   120
            TabIndex        =   190
            Top             =   930
            Width           =   1635
         End
         Begin VB.Label Label15 
            Caption         =   "Surgical"
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
            Left            =   120
            TabIndex        =   189
            Top             =   675
            Width           =   885
         End
         Begin VB.Label Label83 
            Caption         =   "Pre-Hos Spec"
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
            Left            =   120
            TabIndex        =   188
            Top             =   435
            Width           =   1815
         End
         Begin VB.Label Label67 
            Caption         =   "Post"
            ForeColor       =   &H00FF0000&
            Height          =   255
            Left            =   3240
            TabIndex        =   187
            Top             =   630
            Width           =   375
         End
         Begin VB.Label Label91 
            Caption         =   "Ward"
            ForeColor       =   &H00FF0000&
            Height          =   255
            Left            =   4245
            TabIndex        =   186
            Top             =   630
            Width           =   495
         End
         Begin VB.Label Label92 
            Caption         =   "Cons"
            ForeColor       =   &H00FF0000&
            Height          =   255
            Left            =   2235
            TabIndex        =   185
            Top             =   930
            Width           =   1095
         End
         Begin VB.Label Label93 
            Caption         =   "Fee"
            ForeColor       =   &H00C00000&
            Height          =   255
            Left            =   890
            TabIndex        =   184
            Top             =   930
            Width           =   375
         End
         Begin VB.Label Label66 
            Caption         =   "Phy"
            ForeColor       =   &H00FF0000&
            Height          =   255
            Left            =   2235
            TabIndex        =   182
            Top             =   630
            Width           =   375
         End
      End
      Begin VB.Frame FrmeNS 
         Height          =   1280
         Left            =   120
         TabIndex        =   195
         Top             =   3120
         Width           =   6015
         Begin VB.TextBox txtInPhyVisit 
            Alignment       =   1  'Right Justify
            Height          =   285
            Left            =   1320
            TabIndex        =   58
            Top             =   680
            Width           =   495
         End
         Begin VB.TextBox txtIndex 
            Height          =   285
            Left            =   3720
            MaxLength       =   8
            TabIndex        =   196
            Top             =   1320
            Visible         =   0   'False
            Width           =   1095
         End
         Begin VB.TextBox TxtMLab 
            Alignment       =   1  'Right Justify
            Height          =   285
            Left            =   3320
            MaxLength       =   8
            TabIndex        =   54
            Top             =   165
            Width           =   495
         End
         Begin VB.TextBox A5B 
            Alignment       =   1  'Right Justify
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
            Left            =   5160
            MaxLength       =   8
            TabIndex        =   56
            Top             =   165
            Width           =   795
         End
         Begin VB.TextBox TxtMXRay 
            Alignment       =   1  'Right Justify
            Height          =   285
            Left            =   2320
            MaxLength       =   8
            TabIndex        =   53
            Top             =   165
            Width           =   495
         End
         Begin VB.TextBox TxtMPreOthers 
            Alignment       =   1  'Right Justify
            Height          =   285
            Left            =   4320
            MaxLength       =   8
            TabIndex        =   55
            Top             =   165
            Width           =   480
         End
         Begin VB.TextBox txtInHosFees 
            Alignment       =   1  'Right Justify
            Height          =   285
            Left            =   3320
            MaxLength       =   8
            TabIndex        =   60
            Top             =   680
            Width           =   495
         End
         Begin VB.TextBox txtInHosWardFees 
            Alignment       =   1  'Right Justify
            Height          =   285
            Left            =   2320
            MaxLength       =   8
            TabIndex        =   59
            Top             =   680
            Width           =   495
         End
         Begin VB.TextBox txtProcFees 
            Alignment       =   1  'Right Justify
            Enabled         =   0   'False
            Height          =   285
            Left            =   4290
            MaxLength       =   8
            TabIndex        =   61
            Top             =   680
            Width           =   495
         End
         Begin VB.CommandButton cmdMMANonSurgEntry 
            Appearance      =   0  'Flat
            Caption         =   ">"
            Height          =   300
            Left            =   4780
            TabIndex        =   62
            Top             =   680
            Width           =   360
         End
         Begin VB.TextBox A8 
            Alignment       =   1  'Right Justify
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
            Left            =   5160
            MaxLength       =   8
            TabIndex        =   57
            Top             =   420
            Width           =   795
         End
         Begin VB.TextBox A9 
            Alignment       =   1  'Right Justify
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
            Left            =   5160
            MaxLength       =   8
            TabIndex        =   63
            Top             =   680
            Width           =   795
         End
         Begin VB.TextBox A10 
            Alignment       =   1  'Right Justify
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
            Left            =   5160
            MaxLength       =   8
            TabIndex        =   64
            Top             =   930
            Width           =   795
         End
         Begin VB.Label Label96 
            Caption         =   "Ward"
            ForeColor       =   &H00FF0000&
            Height          =   255
            Left            =   1880
            TabIndex        =   198
            Top             =   675
            Width           =   495
         End
         Begin VB.Label Label17 
            Caption         =   "Phy Fees"
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
            Left            =   120
            TabIndex        =   206
            Top             =   680
            Width           =   2085
         End
         Begin VB.Label Label25 
            Caption         =   "Post Hos Treat"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   225
            Left            =   120
            TabIndex        =   207
            Top             =   960
            Width           =   2805
         End
         Begin VB.Label Label13 
            Caption         =   "Pre-Hos Spec"
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
            Left            =   120
            TabIndex        =   205
            Top             =   440
            Width           =   2580
         End
         Begin VB.Label Label69 
            Caption         =   "Days"
            ForeColor       =   &H00FF0000&
            Height          =   255
            Left            =   880
            TabIndex        =   204
            Top             =   675
            Width           =   615
         End
         Begin VB.Label Label79 
            Caption         =   "Pre-Hos"
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
            Left            =   120
            TabIndex        =   203
            Top             =   200
            Width           =   735
         End
         Begin VB.Label Label80 
            Caption         =   "Oth"
            ForeColor       =   &H00C00000&
            Height          =   255
            Left            =   3920
            TabIndex        =   202
            Top             =   195
            Width           =   495
         End
         Begin VB.Label Label81 
            Caption         =   "Lab"
            ForeColor       =   &H00C00000&
            Height          =   255
            Left            =   2880
            TabIndex        =   201
            Top             =   195
            Width           =   375
         End
         Begin VB.Label Label82 
            Caption         =   "X-Ray"
            ForeColor       =   &H00C00000&
            Height          =   255
            Left            =   1880
            TabIndex        =   200
            Top             =   195
            Width           =   495
         End
         Begin VB.Label Label95 
            Caption         =   "Fee"
            ForeColor       =   &H00C00000&
            Height          =   255
            Left            =   2880
            TabIndex        =   199
            Top             =   675
            Width           =   375
         End
         Begin VB.Label Label100 
            Caption         =   "Proc"
            ForeColor       =   &H00C00000&
            Height          =   255
            Left            =   3920
            TabIndex        =   197
            Top             =   675
            Width           =   375
         End
      End
      Begin VB.Frame FrmeOthers 
         Height          =   2415
         Left            =   6300
         TabIndex        =   216
         Top             =   1920
         Width           =   5145
         Begin VB.CheckBox chkMedRptMBM 
            Caption         =   "Check1"
            Height          =   255
            Left            =   4900
            TabIndex        =   77
            Top             =   360
            Width           =   210
         End
         Begin VB.TextBox txtOthersDesc2 
            Height          =   285
            Left            =   1080
            TabIndex        =   219
            Top             =   2640
            Visible         =   0   'False
            Width           =   2415
         End
         Begin VB.TextBox A20B 
            Alignment       =   1  'Right Justify
            Height          =   285
            Left            =   3600
            TabIndex        =   218
            Top             =   2640
            Visible         =   0   'False
            Width           =   915
         End
         Begin VB.TextBox A17 
            Alignment       =   1  'Right Justify
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
            Left            =   3510
            MaxLength       =   8
            TabIndex        =   82
            Top             =   840
            Width           =   555
         End
         Begin VB.TextBox txtLodgerDay 
            Alignment       =   1  'Right Justify
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
            Left            =   1080
            MaxLength       =   3
            TabIndex        =   79
            Top             =   840
            Width           =   435
         End
         Begin VB.TextBox txtLodgerAmt 
            Alignment       =   1  'Right Justify
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
            Left            =   1995
            MaxLength       =   8
            TabIndex        =   80
            Top             =   840
            Width           =   420
         End
         Begin VB.TextBox A22 
            Alignment       =   1  'Right Justify
            Height          =   285
            Left            =   4080
            TabIndex        =   74
            Top             =   120
            Width           =   795
         End
         Begin VB.TextBox txtLodgerTax 
            Alignment       =   1  'Right Justify
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
            Left            =   3510
            MaxLength       =   8
            TabIndex        =   88
            Top             =   1320
            Width           =   555
         End
         Begin VB.TextBox txtICUTax 
            Alignment       =   1  'Right Justify
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
            Left            =   2400
            MaxLength       =   8
            TabIndex        =   87
            Top             =   1320
            Width           =   495
         End
         Begin VB.TextBox txtRBTax 
            Alignment       =   1  'Right Justify
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
            MaxLength       =   8
            TabIndex        =   86
            Top             =   1320
            Width           =   435
         End
         Begin VB.ComboBox CboPayType 
            Enabled         =   0   'False
            Height          =   315
            Left            =   1680
            TabIndex        =   75
            Text            =   "CboPayType"
            Top             =   360
            Visible         =   0   'False
            Width           =   1695
         End
         Begin VB.TextBox ALT 
            Alignment       =   1  'Right Justify
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
            Left            =   2760
            MaxLength       =   5
            TabIndex        =   81
            Top             =   840
            Width           =   435
         End
         Begin VB.CheckBox chkGuaAllowMBM 
            Caption         =   "Check1"
            Height          =   255
            Left            =   4900
            TabIndex        =   84
            Top             =   840
            Width           =   210
         End
         Begin VB.CheckBox chkAdminRegMBM 
            Caption         =   "Check1"
            Height          =   255
            Left            =   4900
            TabIndex        =   217
            Top             =   1080
            Width           =   210
         End
         Begin VB.CheckBox chkOthersMBM 
            Caption         =   "Check1"
            Height          =   255
            Left            =   4900
            TabIndex        =   92
            Top             =   1560
            Width           =   210
         End
         Begin VB.CheckBox chkTelChargesMBM 
            Caption         =   "Check1"
            Height          =   255
            Left            =   4900
            TabIndex        =   117
            Top             =   600
            Width           =   210
         End
         Begin VB.TextBox A15 
            Alignment       =   1  'Right Justify
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
            Left            =   4080
            MaxLength       =   8
            TabIndex        =   76
            Top             =   360
            Width           =   795
         End
         Begin VB.TextBox A16 
            Alignment       =   1  'Right Justify
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
            Left            =   4080
            MaxLength       =   8
            TabIndex        =   78
            Top             =   600
            Width           =   795
         End
         Begin VB.TextBox ALB 
            Alignment       =   1  'Right Justify
            BeginProperty DataFormat 
               Type            =   1
               Format          =   "0.00"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1033
               SubFormatType   =   1
            EndProperty
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
            Left            =   4080
            MaxLength       =   8
            TabIndex        =   83
            Top             =   840
            Width           =   795
         End
         Begin VB.TextBox A18 
            Alignment       =   1  'Right Justify
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
            Left            =   4080
            MaxLength       =   8
            TabIndex        =   85
            Top             =   1080
            Width           =   795
         End
         Begin VB.TextBox A19 
            Alignment       =   1  'Right Justify
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
            Left            =   4080
            MaxLength       =   8
            TabIndex        =   89
            Top             =   1320
            Width           =   795
         End
         Begin VB.TextBox A20 
            Alignment       =   1  'Right Justify
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
            Left            =   4080
            MaxLength       =   10
            TabIndex        =   91
            Top             =   1560
            Width           =   795
         End
         Begin VB.TextBox A25 
            Alignment       =   1  'Right Justify
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
            Left            =   4080
            MaxLength       =   10
            TabIndex        =   94
            Top             =   1780
            Width           =   795
         End
         Begin VB.TextBox A21 
            Alignment       =   1  'Right Justify
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
            Left            =   4080
            MaxLength       =   8
            TabIndex        =   97
            Top             =   2020
            Width           =   795
         End
         Begin VB.TextBox txtOthersDesc1 
            Height          =   285
            Left            =   1560
            TabIndex        =   90
            Top             =   1560
            Width           =   2500
         End
         Begin VB.TextBox txtOhersNoCoPay 
            Height          =   285
            Left            =   1560
            TabIndex        =   93
            Top             =   1780
            Width           =   2500
         End
         Begin VB.TextBox txtDayCash 
            Alignment       =   1  'Right Justify
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
            MaxLength       =   3
            TabIndex        =   95
            Top             =   2020
            Width           =   555
         End
         Begin VB.TextBox txtRMCash 
            Alignment       =   1  'Right Justify
            Height          =   285
            Left            =   2760
            MaxLength       =   3
            TabIndex        =   96
            Top             =   2020
            Width           =   495
         End
         Begin VB.Label Label52 
            Caption         =   "Days"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FF0000&
            Height          =   255
            Left            =   1560
            TabIndex        =   222
            Top             =   885
            Width           =   465
         End
         Begin VB.Label Label90 
            Caption         =   "Amt"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00C00000&
            Height          =   255
            Left            =   3210
            TabIndex        =   221
            Top             =   840
            Width           =   465
         End
         Begin VB.Label Label56 
            Caption         =   "GH Income"
            Height          =   255
            Left            =   120
            TabIndex        =   237
            ToolTipText     =   "Daily cash allowance at Govt hospital"
            Top             =   2040
            Width           =   1335
         End
         Begin VB.Label Label57 
            Caption         =   "Days"
            ForeColor       =   &H00FF0000&
            Height          =   240
            Left            =   2280
            TabIndex        =   236
            Top             =   2055
            Width           =   495
         End
         Begin VB.Label Label38 
            Caption         =   "Others"
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
            Left            =   120
            TabIndex        =   235
            Top             =   1560
            Width           =   795
         End
         Begin VB.Label Label29 
            Caption         =   "Gua. Allow."
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
            Left            =   120
            TabIndex        =   234
            Top             =   850
            Width           =   1635
         End
         Begin VB.Label Label24 
            Caption         =   "Tel. Charges"
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
            Left            =   120
            TabIndex        =   233
            Top             =   600
            Width           =   2085
         End
         Begin VB.Label Label23 
            Caption         =   "Medical Rpt"
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
            Left            =   120
            TabIndex        =   232
            Top             =   360
            Width           =   1770
         End
         Begin VB.Label Label36 
            Caption         =   "Others2"
            Height          =   255
            Left            =   240
            TabIndex        =   231
            Top             =   2640
            Visible         =   0   'False
            Width           =   1095
         End
         Begin VB.Label Label71 
            Caption         =   "Tax"
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
            Left            =   120
            TabIndex        =   229
            Top             =   1320
            Width           =   885
         End
         Begin VB.Label Label30 
            Caption         =   "Admin/Reg."
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
            Left            =   120
            TabIndex        =   228
            Top             =   1080
            Width           =   1635
         End
         Begin VB.Label Label31 
            Caption         =   "R&&B"
            ForeColor       =   &H00FF0000&
            Height          =   255
            Left            =   1200
            TabIndex        =   227
            Top             =   1320
            Width           =   735
         End
         Begin VB.Label Label72 
            Caption         =   "ICU"
            ForeColor       =   &H00FF0000&
            Height          =   255
            Left            =   2040
            TabIndex        =   226
            Top             =   1320
            Width           =   735
         End
         Begin VB.Label Label74 
            Caption         =   "Lodger"
            ForeColor       =   &H00FF0000&
            Height          =   255
            Left            =   2955
            TabIndex        =   225
            Top             =   1320
            Width           =   735
         End
         Begin VB.Label Label5 
            Caption         =   "Claim Type"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   1920
            TabIndex        =   224
            Top             =   390
            Visible         =   0   'False
            Width           =   1095
         End
         Begin VB.Label Label87 
            Caption         =   "Tax"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FF0000&
            Height          =   255
            Left            =   2460
            TabIndex        =   223
            Top             =   885
            Width           =   825
         End
         Begin VB.Label Label97 
            Caption         =   "Others (No Co-Pay)"
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
            Left            =   120
            TabIndex        =   220
            Top             =   1800
            Width           =   1635
         End
         Begin VB.Label Label62 
            Caption         =   "Organ Trans"
            Height          =   255
            Left            =   120
            TabIndex        =   230
            Top             =   120
            Width           =   1335
         End
      End
      Begin VB.CommandButton cmdViewDoc 
         Caption         =   "View"
         Height          =   390
         Left            =   -72480
         TabIndex        =   263
         Top             =   4200
         Width           =   975
      End
      Begin VB.TextBox txtRemarks 
         Height          =   4605
         Left            =   -74880
         MaxLength       =   500
         TabIndex        =   161
         Top             =   600
         Visible         =   0   'False
         Width           =   11415
      End
      Begin VB.Label Label99 
         Caption         =   "Remarks"
         Height          =   255
         Left            =   -74760
         TabIndex        =   162
         Top             =   360
         Visible         =   0   'False
         Width           =   975
      End
   End
   Begin VB.Frame Frame13 
      Height          =   540
      Left            =   120
      TabIndex        =   120
      Top             =   240
      Width           =   11655
      Begin VB.TextBox txtCLMVersion 
         Alignment       =   2  'Center
         Enabled         =   0   'False
         Height          =   285
         Left            =   2640
         TabIndex        =   261
         Top             =   160
         Visible         =   0   'False
         Width           =   375
      End
      Begin VB.TextBox TxtAccess 
         Enabled         =   0   'False
         Height          =   285
         Left            =   10800
         TabIndex        =   5
         Top             =   160
         Width           =   765
      End
      Begin VB.CommandButton cmdCASEnq 
         Caption         =   "Case Enq"
         Height          =   330
         Left            =   5190
         TabIndex        =   116
         Top             =   150
         Width           =   885
      End
      Begin VB.CommandButton cmdMBMEnq 
         Caption         =   "MBM Enq"
         Height          =   330
         Left            =   4320
         TabIndex        =   115
         Top             =   150
         Width           =   885
      End
      Begin VB.TextBox txtAdmissionDate 
         Enabled         =   0   'False
         Height          =   285
         Left            =   6720
         TabIndex        =   3
         Top             =   160
         Width           =   1275
      End
      Begin VB.TextBox txtDisChargeDate 
         Height          =   285
         Left            =   8640
         TabIndex        =   4
         Top             =   160
         Width           =   1215
      End
      Begin VB.CommandButton cmdShow 
         Appearance      =   0  'Flat
         Caption         =   "&Go"
         Height          =   330
         Left            =   3105
         TabIndex        =   1
         Top             =   150
         Width           =   570
      End
      Begin VB.CommandButton cmdSearch 
         Caption         =   "&Search"
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
         Left            =   3675
         TabIndex        =   2
         Top             =   150
         Width           =   645
      End
      Begin VB.TextBox txtCaseNo 
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
         Left            =   1080
         MaxLength       =   15
         TabIndex        =   0
         Top             =   160
         Width           =   1365
      End
      Begin VB.Label Label19 
         Caption         =   "Exc Paid"
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
         Index           =   4
         Left            =   10200
         TabIndex        =   266
         Top             =   200
         Width           =   1005
      End
      Begin VB.Label Label19 
         Caption         =   "DOD"
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
         Index           =   3
         Left            =   8160
         TabIndex        =   265
         Top             =   200
         Width           =   1005
      End
      Begin VB.Label Label19 
         Caption         =   "DOA"
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
         Left            =   6240
         TabIndex        =   264
         Top             =   200
         Width           =   1005
      End
      Begin VB.Label Label19 
         Caption         =   "-"
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
         Left            =   2520
         TabIndex        =   260
         Top             =   195
         Visible         =   0   'False
         Width           =   75
      End
      Begin VB.Label Label19 
         Caption         =   "Case Num"
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
         Left            =   120
         TabIndex        =   121
         Top             =   240
         Width           =   1005
      End
   End
   Begin VB.Frame Frame2 
      Height          =   550
      Left            =   120
      TabIndex        =   124
      Top             =   7320
      Width           =   11655
      Begin VB.CommandButton cmdUpdDischarge 
         Caption         =   "Disharge Date"
         Enabled         =   0   'False
         Height          =   330
         Left            =   4080
         TabIndex        =   288
         Top             =   160
         Visible         =   0   'False
         Width           =   1215
      End
      Begin VB.CommandButton CmdCancel 
         Caption         =   "&Cancel"
         Height          =   330
         Left            =   3240
         TabIndex        =   111
         Top             =   160
         Width           =   855
      End
      Begin VB.CommandButton CmdRemove 
         Caption         =   "&Remove"
         Enabled         =   0   'False
         Height          =   330
         Left            =   2400
         TabIndex        =   110
         Top             =   160
         Width           =   855
      End
      Begin VB.CommandButton CmdAdd 
         Caption         =   "&Add"
         Enabled         =   0   'False
         Height          =   330
         Left            =   1560
         TabIndex        =   109
         Top             =   160
         Width           =   855
      End
      Begin VB.CommandButton CmdEdit 
         Caption         =   "E&dit"
         Enabled         =   0   'False
         Height          =   330
         Left            =   840
         TabIndex        =   108
         Top             =   160
         Width           =   735
      End
      Begin VB.CommandButton CmdTotal 
         Caption         =   "&Total"
         Enabled         =   0   'False
         Height          =   330
         Left            =   120
         TabIndex        =   107
         Top             =   160
         Width           =   735
      End
      Begin VB.CommandButton CmdExit 
         Cancel          =   -1  'True
         Caption         =   "E&xit"
         Height          =   330
         Left            =   10920
         TabIndex        =   114
         Top             =   120
         Width           =   615
      End
      Begin VB.CommandButton CmdClear 
         Caption         =   "&Clear All"
         Enabled         =   0   'False
         Height          =   330
         Left            =   10200
         TabIndex        =   113
         Top             =   120
         Width           =   735
      End
      Begin VB.CommandButton CmdClaerInv 
         Caption         =   "Clear &Inv"
         Height          =   330
         Left            =   9360
         TabIndex        =   112
         Top             =   120
         Width           =   855
      End
      Begin VB.CommandButton CmdUpdate 
         Caption         =   "&Update"
         Height          =   330
         Left            =   840
         TabIndex        =   125
         Top             =   160
         Visible         =   0   'False
         Width           =   735
      End
      Begin VB.Label Label50 
         Caption         =   "WS?"
         Height          =   255
         Index           =   3
         Left            =   8280
         TabIndex        =   166
         Top             =   195
         Width           =   375
      End
      Begin VB.Label Label50 
         Caption         =   "Bordx?"
         Height          =   255
         Index           =   2
         Left            =   7080
         TabIndex        =   165
         Top             =   195
         Width           =   735
      End
      Begin VB.Label lblLastVersion 
         Alignment       =   2  'Center
         BackColor       =   &H00C0FFFF&
         Height          =   255
         Left            =   6480
         TabIndex        =   130
         Top             =   195
         Width           =   375
      End
      Begin VB.Label lblVersion 
         Alignment       =   2  'Center
         BackColor       =   &H00C0FFFF&
         Height          =   255
         Left            =   5760
         TabIndex        =   129
         Top             =   195
         Width           =   375
      End
      Begin VB.Label Label50 
         Caption         =   "Ver"
         Height          =   255
         Index           =   0
         Left            =   5400
         TabIndex        =   128
         Top             =   195
         Width           =   375
      End
      Begin VB.Label lblBordx 
         Alignment       =   2  'Center
         BackColor       =   &H00C0FFFF&
         Height          =   255
         Left            =   7800
         TabIndex        =   127
         Top             =   195
         Width           =   375
      End
      Begin VB.Label lblWs 
         BackColor       =   &H00C0FFFF&
         Height          =   255
         Left            =   8760
         TabIndex        =   126
         Top             =   195
         Width           =   495
      End
      Begin VB.Label Label50 
         Caption         =   "of"
         Height          =   255
         Index           =   1
         Left            =   6240
         TabIndex        =   164
         Top             =   195
         Width           =   375
      End
   End
   Begin VB.Frame Frame9 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1335
      Left            =   120
      TabIndex        =   122
      Top             =   720
      Width           =   11655
      Begin VB.TextBox txtInitialActAmt 
         Alignment       =   1  'Right Justify
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
         Left            =   10440
         MaxLength       =   10
         TabIndex        =   21
         Top             =   390
         Width           =   1095
      End
      Begin VB.TextBox txtPatientName 
         Enabled         =   0   'False
         Height          =   285
         Left            =   1080
         TabIndex        =   6
         Top             =   140
         Width           =   4215
      End
      Begin VB.ComboBox cboINVStayType 
         Enabled         =   0   'False
         Height          =   315
         Left            =   1080
         TabIndex        =   7
         Top             =   390
         Width           =   4215
      End
      Begin VB.ComboBox cboINVPayType 
         Enabled         =   0   'False
         Height          =   315
         ItemData        =   "frmInvoice.frx":0090
         Left            =   1080
         List            =   "frmInvoice.frx":0092
         TabIndex        =   8
         Text            =   "H - Hospital"
         Top             =   675
         Width           =   855
      End
      Begin VB.ComboBox cboHosCode 
         Enabled         =   0   'False
         Height          =   315
         Left            =   2520
         Sorted          =   -1  'True
         TabIndex        =   9
         Top             =   680
         Width           =   2775
      End
      Begin VB.TextBox txtINVPayType 
         Enabled         =   0   'False
         Height          =   285
         Left            =   5280
         TabIndex        =   13
         Top             =   1440
         Visible         =   0   'False
         Width           =   1215
      End
      Begin VB.TextBox txtModal 
         Enabled         =   0   'False
         Height          =   285
         Left            =   5280
         TabIndex        =   14
         Top             =   1590
         Visible         =   0   'False
         Width           =   1215
      End
      Begin MSMask.MaskEdBox txtFolDate 
         Height          =   285
         Left            =   6120
         TabIndex        =   16
         Top             =   165
         Width           =   1095
         _ExtentX        =   1931
         _ExtentY        =   503
         _Version        =   393216
         Enabled         =   0   'False
         MaxLength       =   10
         Mask            =   "##/##/####"
         PromptChar      =   "_"
      End
      Begin VB.TextBox txtFolDate2 
         Height          =   285
         Left            =   6120
         TabIndex        =   119
         Top             =   165
         Visible         =   0   'False
         Width           =   1065
      End
      Begin VB.ComboBox cboDRM 
         Enabled         =   0   'False
         Height          =   315
         Left            =   8160
         TabIndex        =   18
         Top             =   120
         Width           =   1095
      End
      Begin VB.TextBox txtInvExcess 
         Alignment       =   1  'Right Justify
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
         Left            =   8160
         MaxLength       =   10
         TabIndex        =   20
         Top             =   405
         Width           =   1095
      End
      Begin VB.TextBox txtHOSName 
         Enabled         =   0   'False
         Height          =   285
         Left            =   1080
         MaxLength       =   100
         TabIndex        =   10
         Top             =   960
         Width           =   4215
      End
      Begin VB.TextBox txtInvNo 
         Enabled         =   0   'False
         Height          =   285
         Left            =   6120
         TabIndex        =   11
         Top             =   420
         Width           =   1095
      End
      Begin MSMask.MaskEdBox txtInvDate 
         Height          =   285
         Left            =   6120
         TabIndex        =   12
         Top             =   660
         Width           =   1095
         _ExtentX        =   1931
         _ExtentY        =   503
         _Version        =   393216
         Enabled         =   0   'False
         MaxLength       =   10
         Mask            =   "##/##/####"
         PromptChar      =   "_"
      End
      Begin VB.TextBox txtINVDate2 
         Height          =   285
         Left            =   6120
         TabIndex        =   23
         Top             =   660
         Visible         =   0   'False
         Width           =   1095
      End
      Begin VB.ComboBox cboRevBordx 
         Enabled         =   0   'False
         Height          =   315
         ItemData        =   "frmInvoice.frx":0094
         Left            =   8160
         List            =   "frmInvoice.frx":0096
         TabIndex        =   15
         Top             =   660
         Width           =   1095
      End
      Begin VB.ComboBox cboMaternityStatus 
         Enabled         =   0   'False
         Height          =   315
         ItemData        =   "frmInvoice.frx":0098
         Left            =   10440
         List            =   "frmInvoice.frx":009A
         TabIndex        =   22
         Top             =   650
         Visible         =   0   'False
         Width           =   1095
      End
      Begin VB.TextBox txtUHCNo 
         Alignment       =   1  'Right Justify
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
         Left            =   8160
         MaxLength       =   10
         TabIndex        =   19
         Top             =   360
         Visible         =   0   'False
         Width           =   855
      End
      Begin MSMask.MaskEdBox txtDIRDate 
         Height          =   270
         Left            =   6120
         TabIndex        =   17
         Top             =   960
         Width           =   1095
         _ExtentX        =   1931
         _ExtentY        =   476
         _Version        =   393216
         Enabled         =   0   'False
         MaxLength       =   10
         Mask            =   "##/##/####"
         PromptChar      =   "_"
      End
      Begin VB.Label Label19 
         Caption         =   "Maternity"
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
         Index           =   23
         Left            =   9360
         TabIndex        =   291
         Top             =   660
         Visible         =   0   'False
         Width           =   1005
      End
      Begin VB.Label Label65 
         Caption         =   "Hosp Name"
         Height          =   255
         Index           =   33
         Left            =   120
         TabIndex        =   289
         Top             =   960
         Width           =   1095
      End
      Begin VB.Label Label19 
         Caption         =   "Initial Inv Amt"
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
         Index           =   22
         Left            =   9360
         TabIndex        =   287
         Top             =   420
         Width           =   1005
      End
      Begin VB.Label Label19 
         Caption         =   "Inv Pay. Amt"
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
         Index           =   20
         Left            =   9360
         TabIndex        =   282
         Top             =   150
         Width           =   1005
      End
      Begin VB.Label Label19 
         Caption         =   "Excess"
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
         Index           =   19
         Left            =   7320
         TabIndex        =   281
         Top             =   435
         Width           =   1005
      End
      Begin VB.Label Label19 
         Caption         =   "Total Inv"
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
         Index           =   18
         Left            =   7320
         TabIndex        =   280
         Top             =   960
         Width           =   765
      End
      Begin VB.Label Label19 
         Caption         =   "UHC"
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
         Index           =   17
         Left            =   8160
         TabIndex        =   279
         Top             =   360
         Visible         =   0   'False
         Width           =   1005
      End
      Begin VB.Label Label19 
         Caption         =   "IRM"
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
         Index           =   16
         Left            =   7320
         TabIndex        =   278
         Top             =   150
         Width           =   1005
      End
      Begin VB.Label Label19 
         Caption         =   "DIR"
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
         Index           =   15
         Left            =   5400
         TabIndex        =   277
         Top             =   1005
         Width           =   1005
      End
      Begin VB.Label Label19 
         Caption         =   "Fol. Date"
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
         Index           =   14
         Left            =   5400
         TabIndex        =   276
         Top             =   165
         Width           =   1005
      End
      Begin VB.Label Label19 
         Caption         =   "Rev Bordx"
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
         Index           =   13
         Left            =   7320
         TabIndex        =   275
         Top             =   720
         Width           =   1005
      End
      Begin VB.Label Label19 
         Caption         =   "Treat. Mode"
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
         Index           =   12
         Left            =   4320
         TabIndex        =   274
         Top             =   1620
         Visible         =   0   'False
         Width           =   1005
      End
      Begin VB.Label Label19 
         Caption         =   "Cash Type"
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
         Index           =   11
         Left            =   4320
         TabIndex        =   273
         Top             =   1350
         Visible         =   0   'False
         Width           =   1005
      End
      Begin VB.Label Label19 
         Caption         =   "Inv Date"
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
         Index           =   10
         Left            =   5400
         TabIndex        =   272
         Top             =   720
         Width           =   1005
      End
      Begin VB.Label Label19 
         Caption         =   "Inv Num"
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
         Index           =   9
         Left            =   5400
         TabIndex        =   271
         Top             =   435
         Width           =   1005
      End
      Begin VB.Label Label19 
         Caption         =   "Payee"
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
         Index           =   8
         Left            =   2000
         TabIndex        =   270
         Top             =   720
         Width           =   1005
      End
      Begin VB.Label Label19 
         Caption         =   "Pay To"
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
         Index           =   7
         Left            =   120
         TabIndex        =   269
         Top             =   700
         Width           =   1005
      End
      Begin VB.Label Label19 
         Caption         =   "Stay Type"
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
         Index           =   6
         Left            =   120
         TabIndex        =   268
         Top             =   420
         Width           =   1005
      End
      Begin VB.Label Label19 
         Caption         =   "Patient"
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
         Index           =   5
         Left            =   120
         TabIndex        =   267
         Top             =   150
         Width           =   1005
      End
      Begin VB.Label lblPayAmt 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFFC0&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   10440
         TabIndex        =   163
         Top             =   120
         Width           =   1095
      End
      Begin VB.Label txtINVAmount 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFFC0&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   8160
         TabIndex        =   118
         Top             =   960
         Width           =   1095
      End
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00404040&
      Caption         =   "Invoice Processing"
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
      TabIndex        =   142
      Top             =   0
      Width           =   11895
   End
End
Attribute VB_Name = "frmInvoice"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim RBlimit, ICUlimit, KidneyLimit, CashLimit, GuardianLimit, totalFMemOriginal As Double
Dim SurgicalLimit, AnaLimit, OrganTransLimit As Double
Dim intLastPoz1 As Integer
Dim bExit As Boolean
Dim intIndex As String
Dim EditFlag As Boolean
Dim checking As Boolean
Dim HLTCode, INSCode, PLNCode As String
Dim objword As Word.Application
Dim objdoc As Word.Document
Dim ScanCode As String
Dim MBMBTBGStatus As Boolean
Private HOspitals As String
Dim NewSMPlan As String
Dim Bordxstatus As Boolean
Dim PAYCode As String
Dim CLMUHCNo As String
Dim MMASurgFees, MMAAnaeFees, MMASurgFees13th, MMAAnaeFees13th, MMASurgFees13thMin As Double
Dim TotalClick As Boolean
Dim TmpCASHOSName As String
Dim HOSArr As Collection
Dim tmpHospChk As String

Private Type PointAPI
    X  As Long
    Y  As Long
End Type

Private mbActive                As Boolean
Private mlCurThumb              As Long
Private Const SRCCOPY           As Long = &HCC0020
Private Const STRETCH_HALFTONE  As Long = &H4&
Private Const SW_RESTORE        As Long = &H9&

Private Declare Function CreateSolidBrush Lib "gdi32" (ByVal crColor As Long) As Long
Private Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long
Private Declare Function SelectObject Lib "gdi32" (ByVal hDC As Long, ByVal hObject As Long) As Long
Private Declare Function SetBrushOrgEx Lib "gdi32" (ByVal hDC As Long, ByVal nXOrg As Long, ByVal nYOrg As Long, lpPt As PointAPI) As Long
Private Declare Function SetStretchBltMode Lib "gdi32" (ByVal hDC As Long, ByVal nStretchMode As Long) As Long
Private Declare Function StretchBlt Lib "gdi32" (ByVal hDC As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal nSrcWidth As Long, ByVal nSrcHeight As Long, ByVal dwRop As Long) As Long
Private Declare Function UnrealizeObject Lib "gdi32" (ByVal hObject As Long) As Long

Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long

Private Sub A1_LostFocus()
    Dim a, B, C, d As Double
    
    If IsNumeric(ART) = True Then
        a = ART
    Else
        a = 0
    End If
    If IsNumeric(A1) = True Then
        B = A1
    Else
        B = 0
    End If
    A1B = CDbl(a) + CDbl(B)
End Sub


Private Sub A10_Change()
    Call getTotal
End Sub

Private Sub A11_Change()
    Call getTotal
End Sub

Private Sub A11_LostFocus()
    Call getTotal
End Sub

Private Sub A12_Change()
    Call getTotal
End Sub

Private Sub A12_LostFocus()
    Call getTotal
End Sub

Private Sub A13_Change()
    Call getTotal
End Sub

Private Sub A13_LostFocus()
    Call getTotal
End Sub

Private Sub A14_Change()
    Call getTotal
End Sub

Private Sub A14_LostFocus()
    Call getTotal
End Sub

Private Sub A15_Change()
    Call getTotal
End Sub

Private Sub A15_LostFocus()
    Call getTotal
End Sub

Private Sub A16_Change()
    Call getTotal
End Sub

Private Sub A16_LostFocus()
    Call getTotal
End Sub

Private Sub A17_LostFocus()
    Dim a, B, C, d As Double
    
    If IsNumeric(ALT) = True Then
        a = ALT
    Else
        a = 0
    End If
    If IsNumeric(A17) = True Then
        B = A17
    Else
        B = 0
    End If
    ALB = CDbl(a) + CDbl(B)

End Sub

Private Sub A18_Change()
    Call getTotal
End Sub

Private Sub A18_LostFocus()
    Call getTotal
End Sub

Private Sub A19_Change()
    Call getTotal
End Sub

Private Sub A19_LostFocus()
    Call getTotal
End Sub

Private Sub A1B_Change()
    Call getTotal
End Sub

Private Sub A1B_LostFocus()
    Call getTotal
End Sub

Private Sub A2_LostFocus()
    Dim a, B, C, d As Double
    
    If IsNumeric(AIT) = True Then
        a = AIT
    Else
        a = 0
    End If
    If IsNumeric(A2) = True Then
        B = A2
    Else
        B = 0
    End If
    A2B = CDbl(a) + CDbl(B)

End Sub

Private Sub A20_Change()
    Call getTotal
End Sub

Private Sub A20_LostFocus()
    Call getTotal
End Sub

Private Sub A21_Change()
    Call getTotal
End Sub

Private Sub A21_LostFocus()
    Call getTotal
End Sub

Private Sub A22_Change()
    Call getTotal
End Sub

Private Sub A22_LostFocus()
    Call getTotal
End Sub

Private Sub A23_Change()
    Call getTotal
End Sub

Private Sub A23_LostFocus()
    Call getTotal
End Sub

Private Sub A24_Change()
    Call getTotal
End Sub

Private Sub A25_Change()
    Call getTotal
End Sub

Private Sub A25_LostFocus()
    Call getTotal
End Sub

Private Sub A26_Change()
    Call getTotal
End Sub

Private Sub A26_LostFocus()
    Call getTotal
End Sub

Private Sub A27_Change()
    Call getTotal
End Sub

Private Sub A27_KeyPress(KeyAscii As Integer)
    KeyAscii = KeyNumber(KeyAscii)
End Sub

Private Sub A2B_Change()
    Call getTotal
End Sub

Private Sub A2B_LostFocus()
    Call getTotal
End Sub

Private Sub A3_Change()
    Call getTotal
End Sub

Private Sub A3_LostFocus()
    Call getTotal
End Sub

Private Sub A4_Change()
    Call getTotal
End Sub

Private Sub A4_LostFocus()
    Call getTotal
End Sub

Private Sub A5_Change()
    Call getTotal
End Sub

Private Sub A5B_Change()
    Call getTotal
End Sub

Private Sub A6_Change()
    Call getTotal
End Sub

Private Sub A7_Change()
    Call getTotal
End Sub

Private Sub A8_Change()
    Call getTotal
End Sub

Private Sub A9_Change()
    Call getTotal
End Sub

Private Sub AIT_LostFocus()
Dim a, B, C, d As Double
    
    If IsNumeric(AIT) = True Then
        a = AIT
    Else
        a = 0
    End If
    If IsNumeric(A2) = True Then
        B = A2
    Else
        B = 0
    End If
    A2B = CDbl(a) + CDbl(B)

End Sub

Private Sub ALB_Change()
    Call getTotal
End Sub

Private Sub ALB_LostFocus()
    Call getTotal
End Sub

Private Sub ALT_LostFocus()
    Dim a, B, C, d As Double
    
    If IsNumeric(ALT) = True Then
        a = ALT
    Else
        a = 0
    End If
    If IsNumeric(A17) = True Then
        B = A17
    Else
        B = 0
    End If
    ALB = CDbl(a) + CDbl(B)

End Sub

Private Sub ART_LostFocus()
Dim a, B, C, d As Double
    
    If IsNumeric(ART) = True Then
        a = ART
    Else
        a = 0
    End If
    If IsNumeric(A1) = True Then
        B = A1
    Else
        B = 0
    End If
    A1B = CDbl(a) + CDbl(B)
End Sub

Private Sub cboDRM_KeyPress(KeyAscii As Integer)
Dim NewText As String
Dim ValidCount As Integer
Dim ValidValue As String
Dim i As Integer
    
    If KeyAscii >= 32 And KeyAscii <> 127 Then
    NewText = LCase(Left(cboDRM.Text, cboDRM.SelStart) + Chr(KeyAscii) + Mid(cboDRM.Text, cboDRM.SelStart + cboDRM.SelLength + 1))

    ValidCount = 0
    ValidValue = ""

    For i = 0 To cboDRM.ListCount - 1
 
        If NewText = LCase(Left(cboINVPayType.List(i), Len(NewText))) Then
                ValidCount = ValidCount + 1
                ValidValue = cboDRM.List(i)
        End If
        Next
        If ValidCount <= 1 Then KeyAscii = 0
             If ValidCount = 1 Then
               cboDRM.Text = ValidValue
               cboDRM.SelStart = 0
               cboDRM.SelLength = Len(ValidValue)
             End If
        End If
End Sub


Private Sub cboHosCode_Change()
Dim strsqlHOS As String
Dim HOS As New ADODB.Recordset

    'medixmasterconn.Open "File Name=" & BOSPath & "Conn\mediconnect.UDL"
   '
   '     If PAYCode = "AX" Then
   '         strsqlHOS = "select HOSDiscInd,HOSName from HOS where HOSName = '" & Trim(cboHosCode) & "'"
   '         HOS.Open strsqlHOS, medixmasterconn, adOpenKeyset, adLockOptimistic
   '             If HOS.recordCount Then
   '                 If Trim(HOS!HOSDiscInd) = "Y" Then
   '                     MsgBox "Please ensure hospital discount is given!", vbOKOnly + vbInformation
   '                 End If
   '             End If
   '         HOS.Close
   '         Set HOS = Nothing
   '     End If
   ' medixmasterconn.Close
   ' Set medixmasterconn = Nothing

End Sub

Private Sub cboHosCode_LostFocus()
Dim strsqlHOS As String
Dim HOS As New ADODB.Recordset
Dim tmpchkHopsCode As String
Dim tmpAXAClerical As String
        
    
   ' medixmasterconn.Open "File Name=" & BOSPath & "Conn\mediconnect.UDL"

        If (cboHosCode = "DENTAL CLINIC" Or cboHosCode = "DIALYSIS CENTRE" Or cboHosCode = "FOREIGN HOSPITAL" Or cboHosCode = "GENERAL PRACTITIONER" Or cboHosCode = "GOVERNMENT HOSPITAL" Or cboHosCode = "NON-PANEL HOSPITAL" Or cboHosCode = "SPECIALIST CLINIC") And Trim(txtHOSName) = "" Then
            MsgBox "Please Enter Hospital Name or Clinic Name!", vbOKOnly + vbInformation
            txtHOSName.Enabled = True
        ElseIf (cboHosCode <> "DENTAL CLINIC" And cboHosCode <> "DIALYSIS CENTRE" And cboHosCode <> "FOREIGN HOSPITAL" And cboHosCode <> "GENERAL PRACTITIONER" And cboHosCode <> "GOVERNMENT HOSPITAL" And cboHosCode <> "NON-PANEL HOSPITAL" And cboHosCode <> "SPECIALIST CLINIC") And Trim(txtHOSName) <> "" Then
            txtHOSName.Enabled = False
            txtHOSName = ""
        End If
    
    
    
        If PAYCode = "AX" Then
            strsqlHOS = "select HOSDiscInd,HOSName from HOS where HOSName = '" & Trim(cboHosCode) & "'"
            HOS.Open strsqlHOS, medixmasterconn, adOpenKeyset, adLockOptimistic
                If HOS.recordCount Then
                    If Trim(HOS!HOSDiscInd) = "Y" Then
                        MsgBox "Please ensure hospital discount is given!", vbOKOnly + vbInformation
                    End If
                End If
            HOS.Close
            Set HOS = Nothing
        End If
    If (cboHosCode = "DENTAL CLINIC" Or cboHosCode = "DIALYSIS CENTRE" Or cboHosCode = "FOREIGN HOSPITAL" Or cboHosCode = "GENERAL PRACTITIONER" Or txtHOSCode = "GOVERNMENT HOSPITAL" Or cboHosCode = "NON-PANEL HOSPITAL" Or cboHosCode = "SPECIALIST CLINIC") Then
    
    Else
        If Mid(cboINVPayType, 1, 1) = "H" Then
            tmpchkHopsCode = HOSArr.Item(cboHosCode)
            strsqlHOS = ""
    
            strsqlHOS = "select * from HOS where HOSCode = '" & Trim(tmpchkHopsCode) & "'"
            HOS.Open strsqlHOS, medixmasterconn, adOpenKeyset, adLockOptimistic
                
                If HOS!HOSStatus = "P" Then
                    tmpPrivate = True
                Else
                    tmpPrivate = False
                End If
            HOS.Close
            Set HOS = Nothing
            
            If tmpPrivate = True Then
                If tmpHospChk = "Y" Then
                    If cboHosCode <> "" Then
                        'tmpchkHopsCode = HOSArr.Item(cboHosCode)
                        
                        If PAYCode = "AX" And txtMBMPolNo = "03215666" Then
                            tmpAXAClerical = "AX - AXA AFFIN STAFF (NON-EXECUTIVE)"
                            strsqlHOS = ""
                            strsqlHOS = "select * from PanelHospitals where HOSCode = '" & Trim(tmpchkHopsCode) & "' and PayCode = '" & Trim(PAYCode) & "' and HosPanelCode = '" & tmpAXAClerical & "'"
                            HOS.Open strsqlHOS, medixmasterconn, adOpenKeyset, adLockOptimistic
                        Else
                            strsqlHOS = "select * from PanelHospitals where HOSCode = '" & Trim(tmpchkHopsCode) & "' and PayCode = '" & Trim(PAYCode) & "'"
                            HOS.Open strsqlHOS, medixmasterconn, adOpenKeyset, adLockOptimistic
                        End If
                        
                        If HOS.EOF = True Then
                            MsgBox "The selected hospital is not on the panel!", vbOKOnly + vbInformation
                            CmdAdd.Enabled = False
                        Else
                            CmdAdd.Enabled = True
                        End If
                        HOS.Close
                        Set HOS = Nothing
                    End If
                End If
            End If
        End If
    End If
  '  medixmasterconn.Close
   ' Set medixmasterconn = Nothing

End Sub

Private Sub cboINVStayType_Change()
   If checking = True And IsDate(txtInvDate) = True Then
     FrmeOP.Enabled = True
     FrmeHos.Enabled = True
     FrmeSurgical.Enabled = True
     FrmeNS.Enabled = True
     FrmeOthers.Enabled = True
     FrmeNotCover.Enabled = True
     
     If Trim(Mid(ChkStr(cboINVStayType), 1, 1)) = "F" Then 'Post treatment / fol up
        
        If IsDate(txtDisChargeDate) = True Then
            If DateDiff("d", txtDisChargeDate, txtInvDate) < 0 Then 'ealier than discharge date
                MsgBox "Post Treatment:" & vbNewLine & "  Invoice date should not be earlier than discharge date!", vbExclamation
            End If
        ElseIf Trim(Mid(ChkStr(cboINVStayType), 1, 1)) = "F" And Trim(Mid(txtINVPayType, 1, 1)) = "R" Then
            If IsDate(txtDisChargeDate) = True Then
                If DateDiff("d", txtDisChargeDate, txtInvDate) > 31 Then ' later more than 31 days
                    MsgBox "Over 31 days upon discharged!" & vbNewLine & "  All invoices would not be covered anymore!", vbCritical
                    FrmeOP.Enabled = False
                    FrmeHos.Enabled = False
                    FrmeSurgical.Enabled = False
                    FrmeNS.Enabled = False
                    FrmeOthers.Enabled = False
                End If
            End If
        End If
     ElseIf Trim(Mid(ChkStr(cboINVStayType), 1, 1)) = "P" Then 'pre treatment
         If DateDiff("d", txtAdmissionDate, txtInvDate) > 0 Then 'later than AdmissionDate date
            MsgBox "Pre-Hospitalisation :" & vbNewLine & "  Invoice date should not be later than admission date!", vbExclamation
         ElseIf Trim(Mid(ChkStr(cboINVStayType), 1, 1)) = "P" And Trim(Mid(txtINVPayType, 1, 1)) = "R" Then
            If (CInt(DateDiff("d", txtAdmissionDate, txtInvDate)) < -60) Then            ' Earlier more than 60 days
                MsgBox "More than 60 days before admitted!" & vbNewLine & "  All invoices would not be covered anymore!", vbCritical
                FrmeOP.Enabled = False
                FrmeHos.Enabled = False
                FrmeSurgical.Enabled = False
                FrmeNS.Enabled = False
                FrmeOthers.Enabled = False
            End If
         End If
      End If
      
    End If

End Sub

Private Sub cboINVStayType_Click()
   If checking = True And IsDate(txtInvDate) = True Then
     FrmeOP.Enabled = True
     FrmeHos.Enabled = True
     FrmeSurgical.Enabled = True
     FrmeNS.Enabled = True
     FrmeOthers.Enabled = True
     FrmeNotCover.Enabled = True
     
     If Trim(Mid(ChkStr(cboINVStayType), 1, 1)) = "F" Then 'Post treatment / fol up
        
        If IsDate(txtDisChargeDate) = True Then
            If DateDiff("d", txtDisChargeDate, txtInvDate) < 0 Then 'ealier than discharge date
                MsgBox "Post Treatment:" & vbNewLine & "  Invoice date should not be earlier than discharge date!", vbExclamation
            End If
        ElseIf Trim(Mid(ChkStr(cboINVStayType), 1, 1)) = "F" And Trim(Mid(txtINVPayType, 1, 1)) = "R" Then
            If IsDate(txtDisChargeDate) = True Then
                If DateDiff("d", txtDisChargeDate, txtInvDate) > 31 Then ' later more than 31 days
                    MsgBox "Over 31 days upon discharged!" & vbNewLine & "  All invoices would not be covered anymore!", vbCritical
                    FrmeOP.Enabled = False
                    FrmeHos.Enabled = False
                    FrmeSurgical.Enabled = False
                    FrmeNS.Enabled = False
                    FrmeOthers.Enabled = False
                End If
            End If
        End If
     ElseIf Trim(Mid(ChkStr(cboINVStayType), 1, 1)) = "P" Then 'pre treatment
         If DateDiff("d", txtAdmissionDate, txtInvDate) > 0 Then 'later than AdmissionDate date
            MsgBox "Pre-Hospitalisation :" & vbNewLine & "  Invoice date should not be later than admission date!", vbExclamation
         ElseIf Trim(Mid(ChkStr(cboINVStayType), 1, 1)) = "P" And Trim(Mid(txtINVPayType, 1, 1)) = "R" Then
            If (CInt(DateDiff("d", txtAdmissionDate, txtInvDate)) < -60) Then            ' Earlier more than 60 days
                MsgBox "More than 60 days before admitted!" & vbNewLine & "  All invoices would not be covered anymore!", vbCritical
                FrmeOP.Enabled = False
                FrmeHos.Enabled = False
                FrmeSurgical.Enabled = False
                FrmeNS.Enabled = False
                FrmeOthers.Enabled = False
            End If
         End If
      End If
      
    End If

End Sub

Private Sub cboINVStayType_KeyPress(KeyAscii As Integer)
Dim NewText As String
Dim ValidCount As Integer
Dim ValidValue As String
Dim i As Integer
    
    If KeyAscii >= 32 And KeyAscii <> 127 Then
    NewText = LCase(Left(cboINVStayType.Text, cboINVStayType.SelStart) + Chr(KeyAscii) + Mid(cboINVStayType.Text, cboINVStayType.SelStart + cboINVStayType.SelLength + 1))

    ValidCount = 0
    ValidValue = ""

    For i = 0 To cboINVStayType.ListCount - 1
 
        If NewText = LCase(Left(cboINVStayType.List(i), Len(NewText))) Then
                ValidCount = ValidCount + 1
                ValidValue = cboINVStayType.List(i)
        End If
        Next
        If ValidCount <= 1 Then KeyAscii = 0
             If ValidCount = 1 Then
                cboINVStayType.Text = ValidValue
                cboINVStayType.SelStart = 0
                cboINVStayType.SelLength = Len(ValidValue)
             End If
        End If
End Sub

Private Sub cboMaternityStatus_KeyPress(KeyAscii As Integer)
Dim NewText As String
Dim ValidCount As Integer
Dim ValidValue As String
Dim i As Integer
    
    If KeyAscii >= 32 And KeyAscii <> 127 Then
    NewText = LCase(Left(cboMaternityStatus.Text, cboMaternityStatus.SelStart) + Chr(KeyAscii) + Mid(cboMaternityStatus.Text, cboMaternityStatus.SelStart + cboMaternityStatus.SelLength + 1))

    ValidCount = 0
    ValidValue = ""

    For i = 0 To cboMaternityStatus.ListCount - 1
 
        If NewText = LCase(Left(cboMaternityStatus.List(i), Len(NewText))) Then
                ValidCount = ValidCount + 1
                ValidValue = cboMaternityStatus.List(i)
        End If
        Next
        If ValidCount <= 1 Then KeyAscii = 0
             If ValidCount = 1 Then
               cboMaternityStatus.Text = ValidValue
               cboMaternityStatus.SelStart = 0
               cboMaternityStatus.SelLength = Len(ValidValue)
             End If
        End If
End Sub

Private Sub cboRevBordx_Change()
    If EditFlag = False And Mid(cboRevBordx, 1, 1) = "Y" Then
        cmdUpdate_Click
    End If
End Sub

Private Sub cboRevBordx_Click()
    If EditFlag = False And Mid(cboRevBordx, 1, 1) = "Y" Then
        cmdUpdate_Click
    End If
End Sub

Private Sub cboRevBordx_KeyPress(KeyAscii As Integer)
Dim NewText As String
Dim ValidCount As Integer
Dim ValidValue As String
Dim i As Integer
    
    If KeyAscii >= 32 And KeyAscii <> 127 Then
    NewText = LCase(Left(cboRevBordx.Text, cboRevBordx.SelStart) + Chr(KeyAscii) + Mid(cboRevBordx.Text, cboRevBordx.SelStart + cboRevBordx.SelLength + 1))

    ValidCount = 0
    ValidValue = ""

    For i = 0 To cboRevBordx.ListCount - 1
 
        If NewText = LCase(Left(cboRevBordx.List(i), Len(NewText))) Then
                ValidCount = ValidCount + 1
                ValidValue = cboRevBordx.List(i)
        End If
        Next
        If ValidCount <= 1 Then KeyAscii = 0
             If ValidCount = 1 Then
               cboRevBordx.Text = ValidValue
               cboRevBordx.SelStart = 0
               cboRevBordx.SelLength = Len(ValidValue)
             End If
        End If
End Sub

Private Sub cboRoomType_KeyPress(KeyAscii As Integer)
Dim NewText As String
Dim ValidCount As Integer
Dim ValidValue As String
Dim i As Integer
    
    If KeyAscii >= 32 And KeyAscii <> 127 Then
    NewText = LCase(Left(cboRoomType.Text, cboRoomType.SelStart) + Chr(KeyAscii) + Mid(cboRoomType.Text, cboRoomType.SelStart + cboRoomType.SelLength + 1))

    ValidCount = 0
    ValidValue = ""

    For i = 0 To cboRoomType.ListCount - 1
 
        If NewText = LCase(Left(cboRoomType.List(i), Len(NewText))) Then
                ValidCount = ValidCount + 1
                ValidValue = cboRoomType.List(i)
        End If
        Next
        If ValidCount <= 1 Then KeyAscii = 0
             If ValidCount = 1 Then
                cboRoomType.Text = ValidValue
                cboRoomType.SelStart = 0
                cboRoomType.SelLength = Len(ValidValue)
             End If
        End If

End Sub

Private Sub cmdCASEnq_Click()
    If txtCaseNo <> "" Then
        frmCaseInquiry2016.txtCaseNo = txtCaseNo
        frmCaseInquiry2016.cmdShow_Click
    End If

End Sub


Private Sub cmdHSS_Click()
    'frmHSSDetails.Visible = True
    frmHSSDetails.Show vbModal
     
End Sub

Private Sub cmdMBMEnq_Click()
    If txtMBMNumber <> "" Then
        frmNewMBMEnquiry.txtMBMNumber = txtMBMNumber
        frmNewMBMEnquiry.cmdShow_Click
    End If
End Sub

Private Sub cmdMMAEntry_Click()
    frmMMAEntry.Source = 0
    frmMMAEntry.FrameNSProc.Visible = False
    frmMMAEntry.FrameNSOPCS.Visible = False
    frmMMAEntry.FrameNSMMARate.Visible = False
    frmMMAEntry.FrameNSActFee.Visible = False
    frmMMAEntry.FrameNS13Rate.Visible = False
    frmMMAEntry.FrameNS13RateMin.Visible = False
    frmMMAEntry.FrameSProc.Visible = True
    frmMMAEntry.FrameSOPCS.Visible = True
    frmMMAEntry.FrameSMMARate.Visible = True
    frmMMAEntry.FrameSActFee.Visible = True
    frmMMAEntry.FrameS13Rate.Visible = True
    frmMMAEntry.FrameAnaesMMA.Visible = True
    frmMMAEntry.FrameAnaesAct.Visible = True
    frmMMAEntry.Frame13Rate.Visible = True
    frmMMAEntry.Label6.Visible = True
    frmMMAEntry.FrameS13MinRate.Visible = True
    frmMMAEntry.Label5.Caption = "Surgical"
    frmMMAEntry.Label7.Caption = "Surgical Procedures Entry"
    frmMMAEntry.Visible = True
    
    If lstInvoice.ListItems.Count Then
        'If lstInvoice.ListItems(1).SubItems(1) = "True" Then
        '    frmMMAEntry.FrameSProc.Enabled = False
        '    frmMMAEntry.FrameSOPCS.Enabled = False
        '    frmMMAEntry.FrameSMMARate.Enabled = False
        '    frmMMAEntry.FrameSActFee.Enabled = False
        '    frmMMAEntry.FrameS13Rate.Enabled = False
        '    frmMMAEntry.FrameAnaesMMA.Enabled = False
        '    frmMMAEntry.FrameAnaesAct.Enabled = False
        '    frmMMAEntry.Frame13Rate.Enabled = False
        '    frmMMAEntry.FrameS13MinRate.Enabled = False
        'Else
            frmMMAEntry.FrameSProc.Enabled = True
            frmMMAEntry.FrameSOPCS.Enabled = True
            frmMMAEntry.FrameSMMARate.Enabled = True
            frmMMAEntry.FrameSActFee.Enabled = True
            frmMMAEntry.FrameS13Rate.Enabled = True
            frmMMAEntry.FrameAnaesMMA.Enabled = True
            frmMMAEntry.FrameAnaesAct.Enabled = True
            frmMMAEntry.Frame13Rate.Enabled = True
            frmMMAEntry.FrameS13MinRate.Enabled = True
        'End If
    End If
    
End Sub
Private Sub cmdMMANonSurgEntry_Click()
    frmMMAEntry.Source = 1
    frmMMAEntry.FrameNSProc.Visible = True
    frmMMAEntry.FrameNSOPCS.Visible = True
    frmMMAEntry.FrameNSMMARate.Visible = True
    frmMMAEntry.FrameNSActFee.Visible = True
    frmMMAEntry.FrameNS13Rate.Visible = True
    frmMMAEntry.FrameNS13RateMin.Visible = True
    frmMMAEntry.FrameSProc.Visible = False
    frmMMAEntry.FrameSOPCS.Visible = False
    frmMMAEntry.FrameSMMARate.Visible = False
    frmMMAEntry.FrameS13Rate.Visible = False
    frmMMAEntry.FrameSActFee.Visible = False
    frmMMAEntry.FrameAnaesMMA.Visible = False
    frmMMAEntry.FrameAnaesAct.Visible = False
    frmMMAEntry.Frame13Rate.Visible = False
    frmMMAEntry.Label6.Visible = False
    frmMMAEntry.FrameS13MinRate.Visible = False
    frmMMAEntry.Label5.Caption = "Non-Surgical"
    frmMMAEntry.Label7.Caption = "Non-Surgical Procedures Entry"
    frmMMAEntry.Visible = True
    
     If lstInvoice.ListItems.Count Then
       'If lstInvoice.ListItems(1).SubItems(1) = "True" Then
       '     frmMMAEntry.FrameNSProc.Enabled = False
       '     frmMMAEntry.FrameNSOPCS.Enabled = False
       '     frmMMAEntry.FrameNSMMARate.Enabled = False
       '     frmMMAEntry.FrameNSActFee.Enabled = False
       '     frmMMAEntry.FrameNS13Rate.Enabled = False
       '     frmMMAEntry.FrameNS13RateMin.Enabled = False
       ' Else
            frmMMAEntry.FrameNSProc.Enabled = True
            frmMMAEntry.FrameNSOPCS.Enabled = True
            frmMMAEntry.FrameNSMMARate.Enabled = True
            frmMMAEntry.FrameNSActFee.Enabled = True
            frmMMAEntry.FrameNS13Rate.Enabled = True
            frmMMAEntry.FrameNS13RateMin.Enabled = True
        'End If
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

Private Sub A22_KeyPress(KeyAscii As Integer)
    KeyAscii = KeyNumber(KeyAscii)
End Sub



Private Sub Form_KeyPress(KeyAscii As Integer)
    Call EnterToTab(KeyAscii)
End Sub

Private Sub Form_Load()
    bExit = False
    EditFlag = False
    
    checking = True   ' do the checking on invoice date
    EnableEntries (False)
    Call ComboListing
    intIndex = 0

End Sub

Private Sub ComboListing()
    Dim HOS As New ADODB.Recordset
    Dim STAYTYPE As New ADODB.Recordset
    Dim RECMode As New ADODB.Recordset
    Dim cmd  As New ADODB.Command
    Dim tmpHOSCode As String
    Dim tmpHosName As String
    Dim Room As New ADODB.Recordset
   ' medixmasterconn.Open "File Name=" & BOSPath & "Conn\mediconnect.UDL"
    'medixmasterconnMaint.Open "File Name=" & BOSPath & "Conn\MediConnHISMaint.UDL"
    
    Set cmd.ActiveConnection = medixmasterconnMaint
    cmd.CommandText = "SearchRECMode"
    cmd.CommandType = adCmdStoredProc
    cmd.CommandTimeout = 15
    RECMode.CursorLocation = adUseClient
    RECMode.CursorType = adOpenForwardOnly
    RECMode.CacheSize = 100
    Set RECMode = cmd.Execute

    Do Until RECMode.EOF
        With RECMode
            cboDRM.AddItem !RecModeCode & " - " & !RecModeDesc
            RECMode.MoveNext
        End With
    Loop
        
    RECMode.Close
    Set RECMode = Nothing

    Set cmd.ActiveConnection = medixmasterconn
    cmd.CommandText = "SearchHOS"
    cmd.CommandType = adCmdStoredProc
    cmd.CommandTimeout = 15
    HOS.CursorLocation = adUseClient
    HOS.CursorType = adOpenForwardOnly
    HOS.CacheSize = 100
    Set HOS = cmd.Execute
    Set HOSArr = New Collection
    
    
    Do Until HOS.EOF
        With HOS
            tmpHOSCode = HOS!HOSCode
            tmpHosName = HOS!HosName
            cboHosCode.AddItem Trim(!HosName)
            HOSArr.Add tmpHOSCode, tmpHosName
            HOS.MoveNext
        End With
    Loop
    HOS.Close
    Set HOS = Nothing
        
    Set cmd.ActiveConnection = medixmasterconnMaint
    cmd.CommandText = "SearchRoomType"
    cmd.CommandType = adCmdStoredProc
    cmd.CommandTimeout = 15
    Room.CursorLocation = adUseClient
    Room.CursorType = adOpenForwardOnly
    Room.CacheSize = 100
    Set Room = cmd.Execute

    Do Until Room.EOF
        With Room
            cboRoomType.AddItem !RoomType
            Room.MoveNext
        End With
    Loop
        
    Room.Close
    Set Room = Nothing
        


    Set cmd.ActiveConnection = medixmasterconn
    cmd.CommandText = "SearchStayType"
    cmd.CommandType = adCmdStoredProc
    cmd.CommandTimeout = 15
    STAYTYPE.CursorLocation = adUseClient
    STAYTYPE.CursorType = adOpenForwardOnly
    STAYTYPE.CacheSize = 100
    Set STAYTYPE = cmd.Execute
    
    Do Until STAYTYPE.EOF
        With STAYTYPE
            cboINVStayType.AddItem !PayTypeCode & " - " & Trim(!PayTypeName)
            STAYTYPE.MoveNext
        End With
    Loop

    STAYTYPE.Close
    Set STAYTYPE = Nothing
    
  '  medixmasterconn.Close
   ' Set medixmasterconn = Nothing
   ' medixmasterconnMaint.Close
   ' Set medixmasterconnMaint = Nothing
   '
    cboINVPayType.AddItem "H - Hospital"
    cboINVPayType.AddItem "M - Member"
    cboINVPayType.ListIndex = 0
    
    cboRevBordx.AddItem "N - No"
    cboRevBordx.AddItem "Y - Yes"
    
    
    cboMaternityStatus.AddItem "N - No"
    cboMaternityStatus.AddItem "D - Normal Delivery"
    cboMaternityStatus.AddItem "C - Caesarean"
End Sub

Private Sub ListClaimTYpe(Optional strPayTo As String)
    Dim PAYTYPE As New ADODB.Recordset
    Dim strsql As String
    If ChkStr(strPayTo) = "H" Then
        strsql = "select PAYTypeCode , PAYTypeName   from  PAYType where PAYTypeCode like '%C' "
    Else
        strsql = "select PAYTypeCode , PAYTypeName   from  PAYType where PAYTypeCode like '%R' "
    End If
    PAYTYPE.Open strsql, medixmasterconn, adOpenKeyset, adLockOptimistic
    CboPayType.Clear
    Do Until PAYTYPE.EOF
        CboPayType.AddItem PAYTYPE!PayTypeCode & " - " & Trim(PAYTYPE!PayTypeName)
        PAYTYPE.MoveNext
    Loop
    PAYTYPE.Close
    Set PAYTYPE = Nothing
End Sub

' * ********************** Button
Private Sub cmdClear_Click()
  Call ClearEntries
End Sub

Private Sub CmdClaerInv_Click()
    Call clearInvoice
End Sub

Private Sub ClearEntries()
   On Error Resume Next
   Dim ctl As Control
    For Each ctl In Me.Controls
        If TypeOf ctl Is TextBox Then
            ctl.Text = ""
        ElseIf TypeOf ctl Is ComboBox Then
          '  If ctl.Name = "cboPayMode" Then
          '      cboPayMode.ListIndex = 0
            'If ctl.Name = "CboPayType" Then
            '    CboPayType.ListIndex = 1
            'Else
            '    ctl.Text = ""
            'End If
        ElseIf TypeOf ctl Is ListView Then
            ctl.ListItems.Clear
        ElseIf TypeOf ctl Is MaskEdBox Then
            ctl.mask = ""
            ctl.Text = ""
            ctl.mask = "##/##/####"
        End If
    Next ctl
    
    txtFolDate.Visible = True
    txtFolDate2.Visible = False
    txtFolDate2 = ""
    txtInvDate.Visible = True
    txtINVDate2.Visible = False
    txtINVDate2 = ""
    
    CmdEdit.Enabled = False
    CmdAdd.Enabled = True
    CmdUpdate.Visible = False
    CmdEdit.Visible = True
    'lblMore15Y.Visible = False
    txtPolStatus = ""
    lblCASStatus = ""
    lblBordx = ""
    lblVersion = ""
    lblLastVersion = ""
    lblWs = ""
    txtINVAmount = 0
    chkMedRptMBM.Value = 0
    chkTelChargesMBM.Value = 0
    chkGuaAllowMBM.Value = 0
    chkAdminRegMBM.Value = 0
    chkOthersMBM.Value = 0
    txtCLMVersion.Visible = False
    Label19(1).Visible = False
    Frame1(1).Visible = False
    List1.Clear
End Sub


Private Sub clearInvoice()
   On Error Resume Next
    Dim ctl As Control
    For Each ctl In Me.Controls
        If ctl.TabIndex >= 7 And ctl.TabIndex <= 106 Then
            If TypeOf ctl Is TextBox Then
                If ctl.Name <> "txtCaseNo" Then
                    ctl.Text = ""
                End If
            ElseIf TypeOf ctl Is ComboBox Then
                 'If ctl.Name = "CboPayType" Then
                 '   CboPayType.ListIndex = 1
                'Else
                '    ctl.Text = ""
                'End If
            ElseIf TypeOf ctl Is MaskEdBox Then
                ctl.mask = ""
                ctl.Text = ""
                ctl.mask = "##/##/####"
            End If
        End If
    Next ctl
    txtFolDate.mask = ""
    txtFolDate.Text = ""
    txtFolDate.mask = "##/##/####"
    txtFolDate.Visible = True
    txtFolDate2.Visible = False
    txtFolDate2 = ""
    txtInvDate.Visible = True
    txtINVDate2.Visible = False
    txtINVDate2 = ""
    cboINVPayType.Text = "H - Hospital"
    lblBordx = ""
    lblVersion = ""
    lblWs = ""
    txtINVAmount = 0
    chkMedRptMBM.Value = 0
    chkTelChargesMBM.Value = 0
    chkGuaAllowMBM.Value = 0
    chkAdminRegMBM.Value = 0
    chkOthersMBM.Value = 0
    txtUHCNo = ""
    cboDRM = ""
    cboRevBordx = ""
    txtInitialActAmt = ""
End Sub

Private Sub EnableEntries(status As Boolean)
   On Error Resume Next
   Dim ctl As Control
    For Each ctl In Me.Controls
        If ctl.TabIndex >= 6 And ctl.TabIndex <= 86 Then
            If TypeOf ctl Is TextBox Then
                If ctl.Name <> "txtCaseNo" Then
                    ctl.Enabled = status
                End If
            ElseIf TypeOf ctl Is ComboBox Then
                ctl.Enabled = status
            ElseIf TypeOf ctl Is MaskEdBox Then
                ctl.Enabled = status
            ElseIf TypeOf ctl Is CheckBox Then
                ctl.Enabled = status
            End If
        End If
    Next ctl
    
        
        
    A5.Enabled = False
    A6.Enabled = False
    A5B.Enabled = False
    A11.Enabled = False
    A7.Enabled = False
    A9.Enabled = False
    A19.Enabled = False
    A3.Enabled = False
    txtSurgicalFee.Enabled = False
    txtAnaFees.Enabled = False
    txtProcFees.Enabled = False
    cmdMMANonSurgEntry.Enabled = True
End Sub

Private Sub EnableEntriesIB(status As Boolean)
   On Error Resume Next
   Dim ctl As Control
    'For Each ctl In Me.Controls
    '    If ctl.TabIndex >= 6 And ctl.TabIndex <= 86 Then
    '        If TypeOf ctl Is TextBox Then
    '            If ctl.Name <> "txtCaseNo" Then
    '                ctl.Enabled = status
    '            End If
    '        ElseIf TypeOf ctl Is ComboBox Then
    '            ctl.Enabled = status
    '        ElseIf TypeOf ctl Is MaskEdBox Then
    '            ctl.Enabled = status
    '        ElseIf TypeOf ctl Is CheckBox Then
    '            ctl.Enabled = status
    '        End If
    '    End If
    'Next ctl
    cmdHSS.Enabled = False
    TxtMPreOthers.Enabled = True
    txtACC.Enabled = True
    A13.Enabled = True
    A26.Enabled = True
    txtLodgerDay.Enabled = True
    txtLodgerAmt.Enabled = True
    ALT.Enabled = True
    A17.Enabled = True
    txtOthersDesc1.Enabled = True
    A20.Enabled = True
    txtDayCash.Enabled = True
    txtRMCash.Enabled = True
    A21.Enabled = True
    cboINVStayType.Enabled = True
    cboINVPayType.Enabled = True
    cboHosCode.Enabled = True
    txtHOSName.Enabled = True
    txtFolDate.Enabled = True
    txtInvNo.Enabled = True
    txtInvDate.Enabled = True
    txtDIRDate.Enabled = True
    cboDRM.Enabled = True
    txtInvExcess.Enabled = True
    cboRevBordx.Enabled = True
    txtInitialActAmt.Enabled = True
    txtInPhyVisit.Enabled = True
    txtInHosWardFees.Enabled = True
    A12.Enabled = True
    
    DB1.Enabled = True
    RoomAmount.Enabled = True
    ART.Enabled = True
    A1.Enabled = True
    ICU1.Enabled = True
    ICUAmount.Enabled = True
    AIT.Enabled = True
    A2.Enabled = True
    
    cmdMMANonSurgEntry.Enabled = False
    A5.Enabled = False
    A6.Enabled = False
    A5B.Enabled = False
    A11.Enabled = False
    A7.Enabled = False
    A9.Enabled = False
    A19.Enabled = False
    A3.Enabled = False
    txtSurgicalFee.Enabled = False
    txtAnaFees.Enabled = False
    txtProcFees.Enabled = False
    txtLodgerDay.Enabled = False
    txtLodgerAmt.Enabled = False
    ALT.Enabled = False
    A17.Enabled = False
    txtDayCash.Enabled = False
    txtRMCash.Enabled = False
    A21.Enabled = False
End Sub

Private Sub EnableEntriesChkBox(status As Boolean)
   Dim ctl As Control
    For Each ctl In Me.Controls
        If ctl.TabIndex >= 6 And ctl.TabIndex <= 87 Then
            If TypeOf ctl Is CheckBox Then
                ctl.Enabled = status
            End If
        End If
    Next ctl
End Sub
Private Sub cboINVPayType_Click()
    Dim i  As Integer
    
    If ChkStr(Mid(cboINVPayType, 1, 1)) = "M" Then
        cboHosCode = "Member"
        txtINVPayType = "Reimbursement"
        txtInvExcess = 0
        txtInvExcess.Enabled = False
    Else
        cboHosCode = ""
        txtINVPayType = "Cashless"
        txtInvExcess.Enabled = True
    End If
End Sub

Private Sub cboINVPayType_Change()
    If Mid(ChkStr(cboINVPayType), 1, 1) = "M" Then
        cboHosCode = "Member"
        txtINVPayType = "Reimbursement"
        txtInvExcess = 0
        txtInvExcess.Enabled = False
    Else
        cboHosCode = ""
        txtINVPayType = "Cashless"
        txtInvExcess.Enabled = True
    End If
End Sub

Private Sub CboPayType_click()
     If checking = True And IsDate(txtInvDate) = True Then
     FrmeOP.Enabled = True
     FrmeHos.Enabled = True
     FrmeSurgical.Enabled = True
     FrmeNS.Enabled = True
     FrmeOthers.Enabled = True
     FrmeNotCover.Enabled = True
     
     If Trim(Mid(ChkStr(CboPayType), 1, 1)) = "F" Then 'Post treatment / fol up
       If IsDate(txtDisChargeDate) = True Then
        If DateDiff("d", txtDisChargeDate, txtInvDate) < 0 Then 'ealier than discharge date
            MsgBox "Post Treatment:" & vbNewLine & "  Invoice date should not be earlier than discharge date!", vbExclamation
        ElseIf Trim(Mid(ChkStr(CboPayType), 1, 2)) = "FR" Then
            If DateDiff("d", txtDisChargeDate, txtInvDate) > 31 Then ' later more than 31 days
                MsgBox "Over 31 days upon discharged!" & vbNewLine & "  All invoices would not be covered anymore!", vbCritical
                FrmeOP.Enabled = False
                FrmeHos.Enabled = False
                FrmeSurgical.Enabled = False
                FrmeNS.Enabled = False
                FrmeOthers.Enabled = False
            End If
        End If
       End If
     ElseIf Trim(Mid(ChkStr(CboPayType), 1, 1)) = "P" Then 'pre treatment
         If DateDiff("d", txtAdmissionDate, txtInvDate) > 0 Then 'later than AdmissionDate date
            MsgBox "Pre-Hospitalisation :" & vbNewLine & "  Invoice date should not be later than admission date!", vbExclamation
         ElseIf Trim(Mid(ChkStr(CboPayType), 1, 2)) = "PR" Then
            If (CInt(DateDiff("d", txtAdmissionDate, txtInvDate)) < -60) Then            ' Earlier more than 60 days
                MsgBox "More than 60 days before admitted!" & vbNewLine & "  All invoices would not be covered anymore!", vbCritical
                FrmeOP.Enabled = False
                FrmeHos.Enabled = False
                FrmeSurgical.Enabled = False
                FrmeNS.Enabled = False
                FrmeOthers.Enabled = False
            End If
         End If
      End If
      
    End If
End Sub

Private Sub CboPayType_Change()
   
   If checking = True And IsDate(txtInvDate) = True Then
     FrmeOP.Enabled = True
     FrmeHos.Enabled = True
     FrmeSurgical.Enabled = True
     FrmeNS.Enabled = True
     FrmeOthers.Enabled = True
     FrmeNotCover.Enabled = True
     
     If Trim(Mid(ChkStr(CboPayType), 1, 1)) = "F" Then 'Post treatment / fol up
        
        If IsDate(txtDisChargeDate) = True Then
            If DateDiff("d", txtDisChargeDate, txtInvDate) < 0 Then 'ealier than discharge date
                MsgBox "Post Treatment:" & vbNewLine & "  Invoice date should not be earlier than discharge date!", vbExclamation
            End If
        ElseIf Trim(Mid(ChkStr(CboPayType), 1, 2)) = "FR" Then
            If IsDate(txtDisChargeDate) = True Then
                If DateDiff("d", txtDisChargeDate, txtInvDate) > 31 Then ' later more than 31 days
                    MsgBox "Over 31 days upon discharged!" & vbNewLine & "  All invoices would not be covered anymore!", vbCritical
                    FrmeOP.Enabled = False
                    FrmeHos.Enabled = False
                    FrmeSurgical.Enabled = False
                    FrmeNS.Enabled = False
                    FrmeOthers.Enabled = False
                End If
            End If
        End If
     ElseIf Trim(Mid(ChkStr(CboPayType), 1, 1)) = "P" Then 'pre treatment
         If DateDiff("d", txtAdmissionDate, txtInvDate) > 0 Then 'later than AdmissionDate date
            MsgBox "Pre-Hospitalisation :" & vbNewLine & "  Invoice date should not be later than admission date!", vbExclamation
         ElseIf Trim(Mid(ChkStr(CboPayType), 1, 2)) = "PR" Then
            If (CInt(DateDiff("d", txtAdmissionDate, txtInvDate)) < -60) Then            ' Earlier more than 60 days
                MsgBox "More than 60 days before admitted!" & vbNewLine & "  All invoices would not be covered anymore!", vbCritical
                FrmeOP.Enabled = False
                FrmeHos.Enabled = False
                FrmeSurgical.Enabled = False
                FrmeNS.Enabled = False
                FrmeOthers.Enabled = False
            End If
         End If
      End If
      
    End If
End Sub






Private Sub List1_Click()
On Error Resume Next
Dim lRet    As Long
Dim sFile   As String
    
    PictureFileName = DocPath & "ScannedDoc\" & txtCaseNo & ""
    '\Invoices"
    crt PictureFileName & "\" & List1.List(List1.ListIndex)

    'If Index = 0 Then
        sFile = PictureFileName & "\" & List1.List(List1.ListIndex)
        lRet = ShellExecute(Me.hwnd, "Open", sFile, &H0&, &H0&, SW_RESTORE)
       ' Picture2.Picture = LoadPicture()
       ' Picture2.Picture = LoadPicture(sFile)
End Sub

Sub crt(pth As String)
Dim f
f = FreeFile
Open App.path + "\temp.html" For Output As #f
    Print #f, "<html><body><center><img src=""" & pth & """></center></body></html>"
Close #f
DoEvents
WebBrowser1.Navigate App.path & "\temp.html"
End Sub

Private Sub lstInvoice_KeyDown(KeyCode As Integer, Shift As Integer)
    Call clearInvoice
    Load frmHSSDetails
    frmHSSDetails.Visible = False
    Call ClearHSS
      
      If lstInvoice.ListItems.Count > 0 Then
          CmdRemove.Enabled = True
          CmdEdit.Enabled = True
          CmdEdit.Visible = True
          CmdUpdate.Visible = False
          
          With lstInvoice.SelectedItem
            txtIndex = .Text
            intIndex = .Text
            If Trim(.SubItems(1)) = "" Then
                lblBordx = "No"
            Else
                lblBordx = .SubItems(1)
            End If
            
            If Trim(.SubItems(2)) = "" Then
                lblWs = "No"
            Else
                lblWs = .SubItems(2)
            End If
            
            If Trim(.SubItems(3)) = "" Then
                lblVersion = 1
            Else
                lblVersion = .SubItems(3)
            End If
            
            If Trim(.SubItems(4)) <> "" Then
                txtInvNo = .SubItems(4)
            End If
            
            If IsDate(.SubItems(5)) = True Then
                txtInvDate = Format(.SubItems(5), "dd/mm/yyyy")
            End If
            
            txtINVAmount = .SubItems(6)
            cboINVPayType = .SubItems(7)
            cboHosCode = .SubItems(8)
            CboPayType = .SubItems(9)
            txtFolDate.mask = ""
            txtFolDate = .SubItems(10)
            txtFolDate.mask = "##/##/####"
            TxtAccess = .SubItems(46)
            ' Hospital
            DB1 = .SubItems(11)
            RoomAmount = .SubItems(12)
            A1 = .SubItems(54)
            ICU1 = .SubItems(13)
            ICUAmount = .SubItems(14)
            A2 = .SubItems(55)
            A2B = .SubItems(80)
            A3 = .SubItems(15)
            
            A4 = .SubItems(16)
            ' Surgical
            txtXray = .SubItems(29)
            txtPreLab = .SubItems(30)
            txtPreOthers = .SubItems(31)
            A5 = .SubItems(17) 'Pre hosp
            A24 = .SubItems(73)
            
            TxtMXRay = .SubItems(69)
            TxtMLab = .SubItems(70)
            TxtMPreOthers = .SubItems(71)
            A5B = .SubItems(72) 'Pre hosp 2
            
            A6 = .SubItems(32) ' Surgical
            txtSurgicalFee = .SubItems(57)
            txtPhysician = .SubItems(47)
            txtPostTreat = .SubItems(48)
            
            A7 = .SubItems(33) ' ANa
            
            'Non -Surgical
            A8 = .SubItems(58)  ' Pre Hosp
            A9 = .SubItems(18)
            A10 = .SubItems(19)
            txtInPhyVisit = .SubItems(59)
            
            ' out patient
            A11 = .SubItems(75)
            txtACC = .SubItems(34)
            txtPTreat = .SubItems(74)
            A12 = .SubItems(35)
            A13 = .SubItems(36)
            txtKidneyMonths = .SubItems(56)
            A14 = .SubItems(37)
                            
            ' OThers
            A22 = .SubItems(60)
            
            A15 = .SubItems(20)  'MR
            A16 = .SubItems(21)   ' phone
            txtLodgerDay = .SubItems(22)  'lodger
            txtLodgerAmt = .SubItems(23)
            A17 = .SubItems(53)
            ALB = .SubItems(81)
            A18 = .SubItems(24) ' admin
            A19 = .SubItems(25) 'tax
            A20 = .SubItems(26) ' other1
            
            txtOthersDesc1 = .SubItems(49)
            A20B = .SubItems(51) ' other2
            txtOthersDesc2 = .SubItems(50)
            txtDayCash = .SubItems(27)
            txtRMCash = .SubItems(28)
            A21 = .SubItems(52)                'Cash allowance
            
            ' Reimb not payable
            txtPreHos60 = .SubItems(38)
            txtPreMEdic = .SubItems(40)
            txtPreGP = .SubItems(39)
            txtPreSpec1 = .SubItems(42)
            txtPostGHosp = .SubItems(43)
            txtPreMR = .SubItems(41)
            txtpostNotSame = .SubItems(44)
            txtOthers = .SubItems(45)
            txtRBTax = .SubItems(61)
            txtICUTax = .SubItems(62)
            txtLodgerTax = .SubItems(63)
            cboINVStayType = .SubItems(64)
            txtINVPayType = .SubItems(65)
            A23 = .SubItems(68)

            
            ART = .SubItems(76)
            AIT = .SubItems(77)
            ALT = .SubItems(78)
            A1B = .SubItems(79)
            A2B = .SubItems(80)
            txtSurgWard = .SubItems(82)
            txtAnaFees = .SubItems(83)
            txtAnaCons = .SubItems(84)
            txtAnaWard = .SubItems(85)
            txtInHosFees = .SubItems(86)
            txtInHosWardFees = .SubItems(87)
            If Trim(.SubItems(88)) = "Y" Then
                chkMedRptMBM.Value = 1
            Else
                chkMedRptMBM.Value = 0
            End If
            If Trim(.SubItems(89)) = "Y" Then
                chkTelChargesMBM.Value = 1
            Else
                chkTelChargesMBM.Value = 0
            End If
            If Trim(.SubItems(90)) = "Y" Then
                chkGuaAllowMBM.Value = 1
            Else
                chkGuaAllowMBM.Value = 0
            End If
            If Trim(.SubItems(91)) = "Y" Then
                chkAdminRegMBM.Value = 1
            Else
                chkAdminRegMBM.Value = 0
            End If
            If Trim(.SubItems(92)) = "Y" Then
                chkOthersMBM.Value = 1
            Else
                chkOthersMBM.Value = 0
            End If
            
            txtGrossRB = .SubItems(93)
            A25 = .SubItems(94)
            txtOhersNoCoPay = .SubItems(95)
            txtProcFees = .SubItems(96)
            txtDIRDate.mask = ""
            txtDIRDate = .SubItems(97)
            txtDIRDate.mask = "##/##/####"
            'txtDIRDate = .SubItems(97)
            cboDRM = .SubItems(98)
            A26 = .SubItems(99)
            txtUHCNo = .SubItems(100)
            txtInvExcess = .SubItems(101)
            lblPayAmt = .SubItems(102)
            txtOthReimNotCovered = .SubItems(103)
            frmHSSDetails.txtInvPharmacy = .SubItems(104)
            frmHSSDetails.txtMedSupp = .SubItems(105)
            frmHSSDetails.txtNursingCare = .SubItems(106)
            frmHSSDetails.txtNursingServ = .SubItems(107)
            frmHSSDetails.txtLab = .SubItems(108)
            frmHSSDetails.txtRadiology = .SubItems(109)
            frmHSSDetails.txtOTSupp = .SubItems(110)
            frmHSSDetails.txtEquipChrg = .SubItems(111)
            frmHSSDetails.txtMisc = .SubItems(112)
            txtInitialActAmt = IIf(Len(.SubItems(113)), .SubItems(113), 0)
            txtHOSName = IIf(Len(.SubItems(114)), .SubItems(114), "")
            A27 = IIf(Len(.SubItems(115)), .SubItems(115), 0)
            cboMaternityStatus = IIf(Len(.SubItems(116)), .SubItems(116), "")
            Call getTotal
        End With
        EnableEntries (False)
        CmdEdit.Enabled = True
    End If
End Sub

Private Sub lstInvoice_KeyUp(KeyCode As Integer, Shift As Integer)
    Call clearInvoice
    Load frmHSSDetails
    frmHSSDetails.Visible = False
    Call ClearHSS
      
      If lstInvoice.ListItems.Count > 0 Then
          CmdRemove.Enabled = True
          CmdEdit.Enabled = True
          CmdEdit.Visible = True
          CmdUpdate.Visible = False
          
          With lstInvoice.SelectedItem
            txtIndex = .Text
            intIndex = .Text
            If Trim(.SubItems(1)) = "" Then
                lblBordx = "No"
            Else
                lblBordx = .SubItems(1)
            End If
            
            If Trim(.SubItems(2)) = "" Then
                lblWs = "No"
            Else
                lblWs = .SubItems(2)
            End If
            
            If Trim(.SubItems(3)) = "" Then
                lblVersion = 1
            Else
                lblVersion = .SubItems(3)
            End If
            
            If Trim(.SubItems(4)) <> "" Then
                txtInvNo = .SubItems(4)
            End If
            
            If IsDate(.SubItems(5)) = True Then
                txtInvDate = Format(.SubItems(5), "dd/mm/yyyy")
            End If
            
            txtINVAmount = .SubItems(6)
            cboINVPayType = .SubItems(7)
            cboHosCode = .SubItems(8)
            CboPayType = .SubItems(9)
            txtFolDate.mask = ""
            txtFolDate = .SubItems(10)
            txtFolDate.mask = "##/##/####"
            TxtAccess = .SubItems(46)
            ' Hospital
            DB1 = .SubItems(11)
            RoomAmount = .SubItems(12)
            A1 = .SubItems(54)
            ICU1 = .SubItems(13)
            ICUAmount = .SubItems(14)
            A2 = .SubItems(55)
            A2B = .SubItems(80)
            A3 = .SubItems(15)
            
            A4 = .SubItems(16)
            ' Surgical
            txtXray = .SubItems(29)
            txtPreLab = .SubItems(30)
            txtPreOthers = .SubItems(31)
            A5 = .SubItems(17) 'Pre hosp
            A24 = .SubItems(73)
            
            TxtMXRay = .SubItems(69)
            TxtMLab = .SubItems(70)
            TxtMPreOthers = .SubItems(71)
            A5B = .SubItems(72) 'Pre hosp 2
            
            A6 = .SubItems(32) ' Surgical
            txtSurgicalFee = .SubItems(57)
            txtPhysician = .SubItems(47)
            txtPostTreat = .SubItems(48)
            
            A7 = .SubItems(33) ' ANa
            
            'Non -Surgical
            A8 = .SubItems(58)  ' Pre Hosp
            A9 = .SubItems(18)
            A10 = .SubItems(19)
            txtInPhyVisit = .SubItems(59)
            
            ' out patient
            A11 = .SubItems(75)
            txtACC = .SubItems(34)
            txtPTreat = .SubItems(74)
            A12 = .SubItems(35)
            A13 = .SubItems(36)
            txtKidneyMonths = .SubItems(56)
            A14 = .SubItems(37)
                            
            ' OThers
            A22 = .SubItems(60)
            
            A15 = .SubItems(20)  'MR
            A16 = .SubItems(21)   ' phone
            txtLodgerDay = .SubItems(22)  'lodger
            txtLodgerAmt = .SubItems(23)
            A17 = .SubItems(53)
            ALB = .SubItems(81)
            A18 = .SubItems(24) ' admin
            A19 = .SubItems(25) 'tax
            A20 = .SubItems(26) ' other1
            
            txtOthersDesc1 = .SubItems(49)
            A20B = .SubItems(51) ' other2
            txtOthersDesc2 = .SubItems(50)
            txtDayCash = .SubItems(27)
            txtRMCash = .SubItems(28)
            A21 = .SubItems(52)                'Cash allowance
            
            ' Reimb not payable
            txtPreHos60 = .SubItems(38)
            txtPreMEdic = .SubItems(40)
            txtPreGP = .SubItems(39)
            txtPreSpec1 = .SubItems(42)
            txtPostGHosp = .SubItems(43)
            txtPreMR = .SubItems(41)
            txtpostNotSame = .SubItems(44)
            txtOthers = .SubItems(45)
            txtRBTax = .SubItems(61)
            txtICUTax = .SubItems(62)
            txtLodgerTax = .SubItems(63)
            cboINVStayType = .SubItems(64)
            txtINVPayType = .SubItems(65)
            A23 = .SubItems(68)

            
            ART = .SubItems(76)
            AIT = .SubItems(77)
            ALT = .SubItems(78)
            A1B = .SubItems(79)
            A2B = .SubItems(80)
            txtSurgWard = .SubItems(82)
            txtAnaFees = .SubItems(83)
            txtAnaCons = .SubItems(84)
            txtAnaWard = .SubItems(85)
            txtInHosFees = .SubItems(86)
            txtInHosWardFees = .SubItems(87)
            If Trim(.SubItems(88)) = "Y" Then
                chkMedRptMBM.Value = 1
            Else
                chkMedRptMBM.Value = 0
            End If
            If Trim(.SubItems(89)) = "Y" Then
                chkTelChargesMBM.Value = 1
            Else
                chkTelChargesMBM.Value = 0
            End If
            If Trim(.SubItems(90)) = "Y" Then
                chkGuaAllowMBM.Value = 1
            Else
                chkGuaAllowMBM.Value = 0
            End If
            If Trim(.SubItems(91)) = "Y" Then
                chkAdminRegMBM.Value = 1
            Else
                chkAdminRegMBM.Value = 0
            End If
            If Trim(.SubItems(92)) = "Y" Then
                chkOthersMBM.Value = 1
            Else
                chkOthersMBM.Value = 0
            End If
            txtGrossRB = .SubItems(93)
            A25 = .SubItems(94)
            txtOhersNoCoPay = .SubItems(95)
            txtProcFees = .SubItems(96)
            txtDIRDate.mask = ""
            txtDIRDate = .SubItems(97)
            txtDIRDate.mask = "##/##/####"
            'txtDIRDate = .SubItems(97)
            cboDRM = .SubItems(98)
            A26 = .SubItems(99)
            txtUHCNo = .SubItems(100)
            txtInvExcess = .SubItems(101)
            lblPayAmt = .SubItems(102)
            txtOthReimNotCovered = .SubItems(103)
            frmHSSDetails.txtInvPharmacy = .SubItems(104)
            frmHSSDetails.txtMedSupp = .SubItems(105)
            frmHSSDetails.txtNursingCare = .SubItems(106)
            frmHSSDetails.txtNursingServ = .SubItems(107)
            frmHSSDetails.txtLab = .SubItems(108)
            frmHSSDetails.txtRadiology = .SubItems(109)
            frmHSSDetails.txtOTSupp = .SubItems(110)
            frmHSSDetails.txtEquipChrg = .SubItems(111)
            frmHSSDetails.txtMisc = .SubItems(112)
            txtInitialActAmt = IIf(Len(.SubItems(113)), .SubItems(113), 0)
            txtHOSName = IIf(Len(.SubItems(114)), .SubItems(114), "")
            A27 = IIf(Len(.SubItems(115)), .SubItems(115), 0)
            cboMaternityStatus = IIf(Len(.SubItems(116)), .SubItems(116), "")
            Call getTotal
        End With
        EnableEntries (False)
        CmdEdit.Enabled = True
    End If
End Sub


Private Sub lstLGListing_ItemCheck(ByVal Item As MSComctlLib.ListItem)
    If Item.Checked = True Then
        ItemCheck = True
        If Item.SubItems(14) = "Y" Then
            MsgBox "LG has been proccessed!", vbCritical
            Item.Checked = False
            
        Else
            Item.Text = "Checked"
        End If
    Else
        If Item.Text = "Checked" Then
            Item.Text = ""
        End If
    End If
End Sub

Private Sub Picture1_Click(Index As Integer)
On Error Resume Next
Dim lRet    As Long
Dim sFile   As String

    If Index = 0 Then
        sFile = DocPath & "ScannedDoc\" & txtCaseNo & "\Invoices\IN" & txtCaseNo & "-" & txtCLMVersion & "(1).jpg"
        lRet = ShellExecute(Me.hwnd, "Open", sFile, &H0&, &H0&, SW_RESTORE)
    ElseIf Index = 1 Then
        sFile = DocPath & "ScannedDoc\" & txtCaseNo & "\Invoices\IN" & txtCaseNo & "-" & txtCLMVersion & "(2).jpg"
        lRet = ShellExecute(Me.hwnd, "Open", sFile, &H0&, &H0&, SW_RESTORE)
    ElseIf Index = 2 Then
        sFile = DocPath & "ScannedDoc\" & txtCaseNo & "\Invoices\IN" & txtCaseNo & "-" & txtCLMVersion & "(3).jpg"
        lRet = ShellExecute(Me.hwnd, "Open", sFile, &H0&, &H0&, SW_RESTORE)
    ElseIf Index = 3 Then
        sFile = DocPath & "ScannedDoc\" & txtCaseNo & "\Invoices\IN" & txtCaseNo & "-" & txtCLMVersion & "(4).jpg"
        lRet = ShellExecute(Me.hwnd, "Open", sFile, &H0&, &H0&, SW_RESTORE)
    ElseIf Index = 4 Then
        sFile = DocPath & "ScannedDoc\" & txtCaseNo & "\Invoices\IN" & txtCaseNo & "-" & txtCLMVersion & "(5).jpg"
        lRet = ShellExecute(Me.hwnd, "Open", sFile, &H0&, &H0&, SW_RESTORE)
    ElseIf Index = 5 Then
        sFile = DocPath & "ScannedDoc\" & txtCaseNo & "\Invoices\IN" & txtCaseNo & "-" & txtCLMVersion & "(6).jpg"
        lRet = ShellExecute(Me.hwnd, "Open", sFile, &H0&, &H0&, SW_RESTORE)
    ElseIf Index = 6 Then
        sFile = DocPath & "ScannedDoc\" & txtCaseNo & "\Invoices\IN" & txtCaseNo & "-" & txtCLMVersion & "(7).jpg"
        lRet = ShellExecute(Me.hwnd, "Open", sFile, &H0&, &H0&, SW_RESTORE)
    ElseIf Index = 7 Then
        sFile = DocPath & "ScannedDoc\" & txtCaseNo & "\Invoices\IN" & txtCaseNo & "-" & txtCLMVersion & "(8).jpg"
        lRet = ShellExecute(Me.hwnd, "Open", sFile, &H0&, &H0&, SW_RESTORE)
    ElseIf Index = 8 Then
        sFile = DocPath & "ScannedDoc\" & txtCaseNo & "\Invoices\IN" & txtCaseNo & "-" & txtCLMVersion & "(9).jpg"
        lRet = ShellExecute(Me.hwnd, "Open", sFile, &H0&, &H0&, SW_RESTORE)
    ElseIf Index = 9 Then
        sFile = DocPath & "ScannedDoc\" & txtCaseNo & "\Invoices\IN" & txtCaseNo & "-" & txtCLMVersion & "(10).jpg"
        lRet = ShellExecute(Me.hwnd, "Open", sFile, &H0&, &H0&, SW_RESTORE)
    End If

End Sub

Private Sub SSTab1_Click(PreviousTab As Integer)
    'If SSTab1.Tab = 2 Then
    '    'Call ShowINVRemarks
    '    medixmasterconnMaint.Open "File Name=" & BOSPath & "Conn\MediConnHISMaint.UDL"
    '        Call LGListing
    '    medixmasterconnMaint.Close
    '    Set medixmasterconnMaint = Nothing
        
    If SSTab1.Tab = 4 Then
        Call LoadDoc
    End If
End Sub

Private Sub LoadDoc()
On Error Resume Next
Dim B As String, f As Integer, d As Integer, e As Integer
    Call addTheFiles
    Exit Sub
End Sub

Sub addTheFiles()
Dim B As String
List1.Clear

PictureFileName = DocPath & "ScannedDoc\" & txtCaseNo & ""
'\Invoices"
B = Dir(PictureFileName & "\*.*")
'b = Dir(Dir1.Path & "\*.*")
If B <> "" Then
    If Right(B, 4) = ".jpg" Then
        List1.AddItem B
    End If
    If Right(B, 4) = ".tif" Then
        List1.AddItem B
    End If
    'If Right(b, 4) = ".art" Then
    '    List1.AddItem b
    'End If
    If Right(B, 4) = ".GIF" Then
        List1.AddItem B
    End If
Else
    Exit Sub
End If
Do: DoEvents
    B = Dir
    If B <> "" Then
        If Right(B, 4) = ".jpg" Then
            List1.AddItem B
        End If
        If Right(B, 4) = ".tif" Then
            List1.AddItem B
        End If
        'If Right(b, 4) = ".art" Then
        '    List1.AddItem b
        'End If
        If Right(B, 4) = ".GIF" Then
            List1.AddItem B
        End If
    Else
        Exit Do
    End If
Loop
'List1.Height = Me.Height - 2520 - 350
End Sub
Private Sub loadform()
On Error Resume Next
Dim PictureFileName As String
    PictureFileName = DocPath & "ScannedDoc\" & txtCaseNo & "\Invoices\IN" & txtCaseNo & "-" & txtCLMVersion & "(1).jpg"
    Picture1(0).Picture = LoadPicture()
    Picture1(0).Picture = LoadPicture(PictureFileName)
    
    If Picture1(0).Picture = 0 Then
       Picture1(0).Visible = False
       Frame1(1).Width = 10920
       Label1(21).Visible = True
       Label1(1).Visible = False
       Label1(21) = "No Document Found!"
    Else
        Picture1(0).Visible = True
        Label1(1).Visible = True
        Label1(1) = "IN" & txtCaseNo & "-" & txtCLMVersion & "(1)"
        Label1(21).Visible = False
        
    End If
       
    PictureFileName = DocPath & "ScannedDoc\" & txtCaseNo & "\Invoices\IN" & txtCaseNo & "-" & txtCLMVersion & "(2).jpg"
    Picture1(1).Picture = LoadPicture()
    Picture1(1).Picture = LoadPicture(PictureFileName)
    If Picture1(1).Picture = 0 Then
       Picture1(1).Visible = False
       Label1(2).Visible = False
    Else
        Picture1(1).Visible = True
        Frame1(1).Width = 2280
        Label1(2).Visible = True
        Label1(2) = "IN" & txtCaseNo & "-" & txtCLMVersion & "(2)"
    End If
    
    PictureFileName = DocPath & "ScannedDoc\" & txtCaseNo & "\Invoices\IN" & txtCaseNo & "-" & txtCLMVersion & "(3).jpg"
    Picture1(2).Picture = LoadPicture()
    Picture1(2).Picture = LoadPicture(PictureFileName)
    If Picture1(2).Picture = 0 Then
        Picture1(2).Visible = False
        Label1(3).Visible = False
    Else
        Picture1(2).Visible = True
        Frame1(1).Width = 3360
        Label1(3).Visible = True
        Label1(3) = "IN" & txtCaseNo & "-" & txtCLMVersion & "(3)"
    End If
    
    PictureFileName = DocPath & "ScannedDoc\" & txtCaseNo & "\Invoices\IN" & txtCaseNo & "-" & txtCLMVersion & "(4).jpg"
    Picture1(3).Picture = LoadPicture()
    Picture1(3).Picture = LoadPicture(PictureFileName)
    If Picture1(3).Picture = 0 Then
        Picture1(3).Visible = False
        Label1(9).Visible = False
    Else
        Picture1(3).Visible = True
        Frame1(1).Width = 4440
        Label1(9).Visible = True
        Label1(9) = "IN" & txtCaseNo & "-" & txtCLMVersion & "(4)"
    End If
    
    PictureFileName = DocPath & "ScannedDoc\" & txtCaseNo & "\Invoices\IN" & txtCaseNo & "-" & txtCLMVersion & "(5).jpg"
    Picture1(4).Picture = LoadPicture()
    Picture1(4).Picture = LoadPicture(PictureFileName)
    If Picture1(4).Picture = 0 Then
        Picture1(4).Visible = False
        Label1(16).Visible = False
    Else
        Picture1(4).Visible = True
        Frame1(1).Width = 5520
        Label1(16).Visible = True
        Label1(16) = "IN" & txtCaseNo & "-" & txtCLMVersion & "(5)"
    End If

    PictureFileName = DocPath & "ScannedDoc\" & txtCaseNo & "\Invoices\IN" & txtCaseNo & "-" & txtCLMVersion & "(6).jpg"
    Picture1(5).Picture = LoadPicture()
    Picture1(5).Picture = LoadPicture(PictureFileName)
    If Picture1(5).Picture = 0 Then
        Picture1(5).Visible = False
        Label1(18).Visible = False
    Else
        Picture1(5).Visible = True
        Frame1(1).Width = 6600
        Label1(18).Visible = True
        Label1(18) = "IN" & txtCaseNo & "-" & txtCLMVersion & "(6)"
    End If


    PictureFileName = DocPath & "ScannedDoc\" & txtCaseNo & "\Invoices\IN" & txtCaseNo & "-" & txtCLMVersion & "(7).jpg"
    Picture1(6).Picture = LoadPicture()
    Picture1(6).Picture = LoadPicture(PictureFileName)
    If Picture1(6).Picture = 0 Then
        Picture1(6).Visible = False
        Label1(78).Visible = False
    Else
        Picture1(6).Visible = True
        Frame1(1).Width = 7680
        Label1(78).Visible = True
        Label1(78) = "IN" & txtCaseNo & "-" & txtCLMVersion & "(7)"
    End If
    
    PictureFileName = DocPath & "ScannedDoc\" & txtCaseNo & "\Invoices\IN" & txtCaseNo & "-" & txtCLMVersion & "(8).jpg"
    Picture1(7).Picture = LoadPicture()
    Picture1(7).Picture = LoadPicture(PictureFileName)
    If Picture1(7).Picture = 0 Then
        Picture1(7).Visible = False
        Label1(77).Visible = False
    Else
        Picture1(7).Visible = True
        Frame1(1).Width = 8760
        Label1(77).Visible = True
        Label1(77) = "IN" & txtCaseNo & "-" & txtCLMVersion & "(8)"
    End If
    
    PictureFileName = DocPath & "ScannedDoc\" & txtCaseNo & "\Invoices\IN" & txtCaseNo & "-" & txtCLMVersion & "(9).jpg"
    Picture1(8).Picture = LoadPicture()
    Picture1(8).Picture = LoadPicture(PictureFileName)
    If Picture1(8).Picture = 0 Then
        Picture1(8).Visible = False
        Label1(67).Visible = False
    Else
        Picture1(8).Visible = True
        Frame1(1).Width = 9840
        Label1(67).Visible = True
        Label1(67) = "IN" & txtCaseNo & "-" & txtCLMVersion & "(9)"
    End If

    PictureFileName = DocPath & "ScannedDoc\" & txtCaseNo & "\Invoices\IN" & txtCaseNo & "-" & txtCLMVersion & "(10).jpg"
    Picture1(9).Picture = LoadPicture()
    Picture1(9).Picture = LoadPicture(PictureFileName)
    If Picture1(9).Picture = 0 Then
        Picture1(9).Visible = False
        Label1(22).Visible = False
    Else
        Picture1(9).Visible = True
        Frame1(1).Width = 10920
        Label1(22).Visible = True
        Label1(22) = "IN" & txtCaseNo & "-" & txtCLMVersion & "(10)"
    End If

End Sub

Private Sub Text1_KeyPress(KeyAscii As Integer)
    KeyAscii = KeyNumber(KeyAscii)
End Sub

Private Sub Text2_KeyPress(KeyAscii As Integer)
    KeyAscii = KeyNumber(KeyAscii)
End Sub

Private Sub Text4_KeyPress(KeyAscii As Integer)
    KeyAscii = KeyNumber(KeyAscii)
End Sub

Private Sub Text5_KeyPress(KeyAscii As Integer)
    KeyAscii = KeyNumber(KeyAscii)
End Sub

Private Sub TxtAccess_KeyPress(KeyAscii As Integer)
    KeyAscii = KeyNumber(KeyAscii)
End Sub

Private Sub txtAnaCons_Change()
Dim a, B, C As Double
    
    If IsNumeric(txtAnaFees) = True Then
        a = txtAnaFees
    Else
        a = 0
    End If
    If IsNumeric(txtAnaCons) = True Then
        B = txtAnaCons
    Else
        B = 0
    End If
    If IsNumeric(txtAnaWard) = True Then
        C = txtAnaWard
    Else
        C = 0
    End If
   
    A7 = CDbl(a) + CDbl(B) + CDbl(C)
End Sub

Private Sub txtAnaCons_KeyPress(KeyAscii As Integer)
    KeyAscii = KeyNumber(KeyAscii)
End Sub

Private Sub txtAnaCons_LostFocus()
Dim a, B, C As Double
    
    If IsNumeric(txtAnaFees) = True Then
        a = txtAnaFees
    Else
        a = 0
    End If
    If IsNumeric(txtAnaCons) = True Then
        B = txtAnaCons
    Else
        B = 0
    End If
    If IsNumeric(txtAnaWard) = True Then
        C = txtAnaWard
    Else
        C = 0
    End If
   
    A7 = CDbl(a) + CDbl(B) + CDbl(C)
End Sub

Private Sub txtAnaFees_Change()
Dim a, B, C As Double
    
    If IsNumeric(txtAnaFees) = True Then
        a = txtAnaFees
    Else
        a = 0
    End If
    If IsNumeric(txtAnaCons) = True Then
        B = txtAnaCons
    Else
        B = 0
    End If
    If IsNumeric(txtAnaWard) = True Then
        C = txtAnaWard
    Else
        C = 0
    End If
   
    A7 = CDbl(a) + CDbl(B) + CDbl(C)

End Sub

Private Sub txtAnaFees_KeyPress(KeyAscii As Integer)
    KeyAscii = KeyNumber(KeyAscii)
End Sub

Private Sub txtAnaFees_LostFocus()
Dim a, B, C As Double
    
    If IsNumeric(txtAnaFees) = True Then
        a = txtAnaFees
    Else
        a = 0
    End If
    If IsNumeric(txtAnaCons) = True Then
        B = txtAnaCons
    Else
        B = 0
    End If
    If IsNumeric(txtAnaWard) = True Then
        C = txtAnaWard
    Else
        C = 0
    End If
   
    A7 = CDbl(a) + CDbl(B) + CDbl(C)

End Sub

Private Sub txtAnaWard_Change()
Dim a, B, C As Double
    
    If IsNumeric(txtAnaFees) = True Then
        a = txtAnaFees
    Else
        a = 0
    End If
    If IsNumeric(txtAnaCons) = True Then
        B = txtAnaCons
    Else
        B = 0
    End If
    If IsNumeric(txtAnaWard) = True Then
        C = txtAnaWard
    Else
        C = 0
    End If
   
    A7 = CDbl(a) + CDbl(B) + CDbl(C)
End Sub

Private Sub txtAnaWard_KeyPress(KeyAscii As Integer)
    KeyAscii = KeyNumber(KeyAscii)
End Sub

Private Sub txtAnaWard_LostFocus()
Dim a, B, C As Double
    
    If IsNumeric(txtAnaFees) = True Then
        a = txtAnaFees
    Else
        a = 0
    End If
    If IsNumeric(txtAnaCons) = True Then
        B = txtAnaCons
    Else
        B = 0
    End If
    If IsNumeric(txtAnaWard) = True Then
        C = txtAnaWard
    Else
        C = 0
    End If
   
    A7 = CDbl(a) + CDbl(B) + CDbl(C)
End Sub

Private Sub txtCaseNo_GotFocus()
    cmdShow.Default = True
    CmdUpdate.Visible = False
    CmdEdit.Visible = True
End Sub

Private Sub TxtCaseNo_LostFocus()
    txtCaseNo = ChkStr(txtCaseNo)
    cmdShow.Default = False
End Sub

Private Sub txtDayCash_KeyPress(KeyAscii As Integer)
    KeyAscii = KeyNumber(KeyAscii)
End Sub

Private Sub txtDIRDate_LostFocus()
    If txtDIRDate <> "__/__/____" Then
        If IsDate(txtDIRDate) = False Then
            MsgBox "Invalid date format!", vbCritical
        End If
    End If
End Sub

Private Sub txtFolDate_LostFocus()
    If txtFolDate <> "__/__/____" Then
        If IsDate(txtFolDate) = False Then
            MsgBox "Invalid date format!", vbCritical
        End If
    End If
End Sub

Private Sub txtGrossRB_Change()
    Call getTotal
End Sub

Private Sub txtGrossRB_LostFocus()
    Call getTotal
End Sub


Private Sub txtHOSName_LostFocus()
    txtHOSName = ChkStr(txtHOSName)
End Sub

Private Sub txtICUTax_LostFocus()
Dim a, B, C As Double
    
    If IsNumeric(txtRBTax) = True Then
        a = txtRBTax
    Else
        a = 0
    End If
    If IsNumeric(txtICUTax) = True Then
        B = txtICUTax
    Else
        B = 0
    End If
    If IsNumeric(txtLodgerTax) = True Then
        C = txtLodgerTax
    Else
        C = 0
    End If
    A19 = CDbl(a) + CDbl(B) + CDbl(C)
End Sub

Private Sub txtIndex_KeyPress(KeyAscii As Integer)
    KeyAscii = KeyNumber(KeyAscii)
End Sub

Private Sub txtInHosFees_Change()
Dim a, B, C As Double
    
    If IsNumeric(txtInHosFees) = True Then
        a = txtInHosFees
    Else
        a = 0
    End If
    If IsNumeric(txtInHosWardFees) = True Then
        B = txtInHosWardFees
    Else
        B = 0
    End If
    
    If IsNumeric(txtProcFees) = True Then
        C = txtProcFees
    Else
        C = 0
    End If
    
    A9 = CDbl(a) + CDbl(B) + CDbl(C)
End Sub

Private Sub txtInHosFees_KeyPress(KeyAscii As Integer)
    KeyAscii = KeyNumber(KeyAscii)
End Sub

Private Sub txtInHosFees_LostFocus()
Dim a, B, C As Double
    
    If IsNumeric(txtInHosFees) = True Then
        a = txtInHosFees
    Else
        a = 0
    End If
    If IsNumeric(txtInHosWardFees) = True Then
        B = txtInHosWardFees
    Else
        B = 0
    End If
    
    If IsNumeric(txtProcFees) = True Then
        C = txtProcFees
    Else
        C = 0
    End If
    
    A9 = CDbl(a) + CDbl(B) + CDbl(C)
End Sub

Private Sub txtInHosWardFees_Change()
Dim a, B, C As Double
    
    If IsNumeric(txtInHosFees) = True Then
        a = txtInHosFees
    Else
        a = 0
    End If
    If IsNumeric(txtInHosWardFees) = True Then
        B = txtInHosWardFees
    Else
        B = 0
    End If
    
    If IsNumeric(txtProcFees) = True Then
        C = txtProcFees
    Else
        C = 0
    End If
    
    A9 = CDbl(a) + CDbl(B) + CDbl(C)
End Sub

Private Sub txtInHosWardFees_KeyPress(KeyAscii As Integer)
    KeyAscii = KeyNumber(KeyAscii)
End Sub

Private Sub txtInHosWardFees_LostFocus()
Dim a, B, C As Double
    
    If IsNumeric(txtInHosFees) = True Then
        a = txtInHosFees
    Else
        a = 0
    End If
    If IsNumeric(txtInHosWardFees) = True Then
        B = txtInHosWardFees
    Else
        B = 0
    End If
    
    If IsNumeric(txtProcFees) = True Then
        C = txtProcFees
    Else
        C = 0
    End If
    
    A9 = CDbl(a) + CDbl(B) + CDbl(C)
End Sub

Private Sub txtInPhyVisit_KeyPress(KeyAscii As Integer)
    KeyAscii = KeyNumber(KeyAscii)
End Sub



Private Sub txtINVAmount_Change()
Dim TempInvAmt As Double
Dim TempInvExcess As Double
Dim TempInvPayAmt As Double
    
    If Mid(cboINVPayType, 1, 1) = "H" Then

        lblPayAmt = 0
        
        TempInvAmt = IIf(IsNumeric(txtINVAmount), txtINVAmount, 0)
        
        TempInvExcess = IIf(IsNumeric(txtInvExcess), txtInvExcess, 0)
    
        TempInvPayAmt = IIf(IsNumeric(lblPayAmt), lblPayAmt, 0)
        
        lblPayAmt = CDbl(TempInvAmt) - CDbl(TempInvExcess)
    End If
End Sub

Private Sub txtInvDate_LostFocus()
   If checking = True And IsDate(txtInvDate) = True Then
     FrmeOP.Enabled = True
     FrmeHos.Enabled = True
     FrmeSurgical.Enabled = True
     FrmeNS.Enabled = True
     FrmeOthers.Enabled = True
     FrmeNotCover.Enabled = True
     
     If Trim(Mid(ChkStr(cboINVStayType), 1, 1)) = "F" Then 'Post treatment / fol up
        txtFolDate.mask = ""
        txtFolDate = txtInvDate
        txtFolDate.mask = "##/##/####"
        
        If IsDate(txtDisChargeDate) = True Then
            If DateDiff("d", txtDisChargeDate, txtInvDate) < 0 Then 'ealier than discharge date
                MsgBox "Post Treatment:" & vbNewLine & "  Invoice date should not be earlier than discharge date!", vbExclamation
            ElseIf Trim(Mid(ChkStr(cboINVStayType), 1, 1)) = "F" And Trim(Mid(ChkStr(cboINVPayType), 1, 1)) = "C" Then
                If DateDiff("d", txtDisChargeDate, txtInvDate) > 31 Then ' later more than 31 days
                    MsgBox "Over 31 days upon discharged!", vbCritical
                End If
            ElseIf Trim(Mid(ChkStr(cboINVStayType), 1, 1)) = "F" And Trim(Mid(ChkStr(cboINVPayType), 1, 1)) = "R" Then
                If DateDiff("d", txtDisChargeDate, txtInvDate) > 31 Then ' later more than 31 days
                    MsgBox "Over 31 days upon discharged!" & vbNewLine & "  All invoices would not be covered anymore!", vbCritical
                    FrmeOP.Enabled = False
                    FrmeHos.Enabled = False
                    FrmeSurgical.Enabled = False
                    FrmeNS.Enabled = False
                    FrmeOthers.Enabled = False
                End If
            End If
        End If
     ElseIf Trim(Mid(ChkStr(cboINVStayType), 1, 1)) = "P" Then 'pre treatment
         If DateDiff("d", txtAdmissionDate, txtInvDate) > 0 Then 'later than AdmissionDate date
            MsgBox "Pre-Hospitalisation :" & vbNewLine & "  Invoice date should not be later than admission date!", vbExclamation
         ElseIf Trim(Mid(ChkStr(cboINVStayType), 1, 1)) = "P" And Trim(Mid(ChkStr(cboINVPayType), 1, 1)) = "C" Then
            If (CInt(DateDiff("d", txtAdmissionDate, txtInvDate)) < -60) Then            ' Earlier more than 60 days
                MsgBox "More than 60 days before admitted!", vbCritical
            End If
         ElseIf Trim(Mid(ChkStr(cboINVStayType), 1, 1)) = "P" And Trim(Mid(ChkStr(cboINVPayType), 1, 1)) = "R" Then
            If (CInt(DateDiff("d", txtAdmissionDate, txtInvDate)) < -60) Then            ' Earlier more than 60 days
                MsgBox "More than 60 days before admitted!" & vbNewLine & "  All invoices would not be covered anymore!", vbCritical
                FrmeOP.Enabled = False
                FrmeHos.Enabled = False
                FrmeSurgical.Enabled = False
                FrmeNS.Enabled = False
                FrmeOthers.Enabled = False
            End If
         End If
      End If
      
    End If
    If txtInvDate <> "__/__/____" Then
        If IsDate(txtInvDate) = False Then
            MsgBox "Invalid date format!", vbCritical
            txtInvDate.SetFocus
        Else
            txtDIRDate = txtInvDate
        End If
    End If
End Sub

Private Sub cmdcancel_Click()
    Call EnableEntries(False)
    CmdAdd.Enabled = True
    CmdUpdate.Visible = False
    CmdEdit.Visible = True
    CmdTotal.Enabled = True
End Sub

Private Sub cmdExit_Click()
    Unload Me
    Set frmInvoice = Nothing
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Unload Me
    Set frmInvoice = Nothing
End Sub

Private Sub CmdSearch_Click()
    frmCASList.Source = 6
    frmCASList.Show vbModal
End Sub

Public Sub cmdShow_Click()
    Dim ADOrstCAS As New ADODB.Recordset
    Dim PolStatus As New ADODB.Recordset
    Dim sql As String
    Dim temp As String
    temp = txtCaseNo
    Call ClearEntries
    txtCaseNo = temp
    Dim PLNSOF As Boolean
    
    Screen.MousePointer = vbHourglass
    SSTab1.Tab = 0
    cmdShow.Enabled = False
   ' medixmasterconn.Open "File Name=" & BOSPath & "Conn\mediconnect.UDL"
    'medixmasterconnMaint.Open "File Name=" & BOSPath & "Conn\MediConnHISMaint.UDL"
    'medixmasterconnWS.Open "File Name=" & BOSPath & "Conn\mediconnectWS.UDL"

    ScanCode = txtCaseNo
    If Mid(ScanCode, 1, 2) = "LO" Then
        txtCaseNo = ""
        Call SearchForm(ScanCode)
    Else
        If Trim(txtCaseNo) <> "" Then
        Load frmHSSDetails
        frmHSSDetails.Visible = False
        Call ClearHSS
        frmHSSDetails.Visible = False
        sql = "Exec SearchCasAdjustment '" & Trim(txtCaseNo) & "'"
        Set ADOrstCAS = New ADODB.Recordset
        ADOrstCAS.Open sql, medixmasterconn, adOpenKeyset, adLockOptimistic
        
        If Not ADOrstCAS.EOF Then
            With ADOrstCAS
                CmdAdd.Enabled = True
                CmdClear.Enabled = True
                CmdTotal.Enabled = True
                
                If !PLNSOF = "Y" Then
                    PLNSOF = True
                Else
                    PLNSOF = False
                End If
                HLTCode = !HLTCode
                If Len(!PLNCode) Then
                    PLNCode = !PLNCode
                End If
                If Len(!INSCode) Then
                    INSCode = !INSCode
                End If
                If Len(!PAYCode) Then
                    PAYCode = !PAYCode
                End If
                txtCASNumber = !CasNumber
                txtMBMNumber = !MBMNumber
                If Len(!MBMName) Then
                    txtMBMName = !MBMName
                End If
                If Len(!MBMPolicyNo) Then
                    txtMBMPolNo = !MBMPolicyNo
                End If
                'If Trim(!PayTypeCode) <> "" Then
                '    txtType = !PayTypeCode
                'End If
                If IsNull(!MBMCName) = True Then 'Pateint = Principal
                    txtPatientName = !MBMName
                    txtDOB = !MBMDOB
                Else
                    txtPatientName = !MBMCName
                    If Len(!MBMCDOB) Then
                        txtDOB = !MBMCDOB
                    End If
                End If
                If Trim(!CASHOSCode) <> "" Then
                    txtHospital = GetHosName(!CASHOSCode)
                End If
                'If IsDate(txtDOB) Then
                '    If Year(Date) - Year(txtDOB) > 15 Then
                '        'lblMore15Y.Visible = True
                '        A17.Enabled = False
                '        txtLodgerAmt.Enabled = False
                '        txtLodgerDay.Enabled = False
                '    Else
                '        lblMore15Y.Visible = False
                '    End If
                'End If
                sqlPolStatus = " select pstatuscode ,pstatusDesc from PolicyStatus where PStatusCode ='" & Trim(!MBMStatus) & "'"
                PolStatus.Open sqlPolStatus, medixmasterconn, adOpenKeyset, adLockOptimistic
                If PolStatus.recordCount Then
                    txtPolStatus = PolStatus!pstatusDesc
                End If
                
                PolStatus.Close
                Set PolStatus = Nothing
    
                If ChkStr(!CASStatus) = "O" Then
                    lblCASStatus = "Open"
                ElseIf ChkStr(!CASStatus) = "C" Then
                    lblCASStatus = "Close"
                ElseIf ChkStr(!CASStatus) = "R" Then
                    lblCASStatus = "Repudiated"
                End If
                
                'If Trim(!ClaimNatureCode) <> "" Then
                '    txtClaimNature = !ClaimNatureCode
                'End If
                
                If Trim(!ModalityCode) <> "" Then
                    txtModal = !ModalityCode
                End If
                If Trim(!MBMPayorEXPDate) <> "" Then
                    txtExpDate = !MBMPayorEXPDate
                End If
                If Trim(!MBMPayorEffDate) <> "" Then
                    txtEffDate = !MBMPayorEffDate
                End If
                
                If Len(!CASDateAdmitted) Then
                    txtAdmissionDate = !CASDateAdmitted
                End If
                If Len(Trim(!CASDischargeDate)) <> 0 Then
                    txtDisChargeDate = !CASDischargeDate
                End If
                If !MBMBTBG = "Y" Then
                    MBMBTBGStatus = True
                Else
                    MBMBTBGStatus = False
                End If
                
                TmpCASHOSName = IIf(Len(!HosName), !HosName, "")
                frmMMAEntry.CalSource = 1
                If frmMMAEntry.Visible = False Then
                    Load frmMMAEntry
                End If
                
                frmMMAEntry.txtPROCINVDone1 = IIf(Len(!CASOProcInvestigationDone1), !CASOProcInvestigationDone1, "")
                frmMMAEntry.txtPROCINVDone2 = IIf(Len(!CASOProcInvestigationDone2), !CASOProcInvestigationDone2, "")
                frmMMAEntry.txtPROCINVDone3 = IIf(Len(!CASOProcInvestigationDone3), !CASOProcInvestigationDone3, "")
                frmMMAEntry.txtINVCode1 = IIf(Len(!CASInvCode1), !CASInvCode1, "")
                frmMMAEntry.txtINVCode2 = IIf(Len(!CASInvCode2), !CASInvCode2, "")
                frmMMAEntry.txtINVCode3 = IIf(Len(!CASInvCode3), !CASInvCode3, "")
                frmMMAEntry.txtMEDSURDone1 = IIf(Len(!CASOMedSurMgmt1), !CASOMedSurMgmt1, "")
                frmMMAEntry.txtMEDSURDone2 = IIf(Len(!CASOMedSurMgmt2), !CASOMedSurMgmt2, "")
                frmMMAEntry.txtMEDSURDone3 = IIf(Len(!CASOMedSurMgmt3), !CASOMedSurMgmt3, "")
                frmMMAEntry.txtATTDoc(0) = IIf(Len(!CASAttendDoctor1), !CASAttendDoctor1, "")
                frmMMAEntry.txtATTDoc(1) = IIf(Len(!CASAttendDoctor2), !CASAttendDoctor2, "")
                frmMMAEntry.txtATTDoc(2) = IIf(Len(!CASAttendDoctor3), !CASAttendDoctor3, "")
                frmMMAEntry.txtATTDoc(3) = IIf(Len(!CASAttendDoctor4), !CASAttendDoctor4, "")
                frmMMAEntry.txtATTDoc(4) = IIf(Len(!CASAttendDoctor5), !CASAttendDoctor5, "")
                frmMMAEntry.txtAnaesDoc(0) = IIf(Len(!CASAnaesName1), !CASAnaesName1, "")
                frmMMAEntry.txtAnaesDoc(1) = IIf(Len(!CASAnaesName2), !CASAnaesName2, "")
                frmMMAEntry.txtAnaesDoc(2) = IIf(Len(!CASAnaesName3), !CASAnaesName3, "")
                
                'Surgical Proc
                frmMMAEntry.txtProcedure1 = IIf(Len(!CASOPROCRequired1), !CASOPROCRequired1, "")
                frmMMAEntry.txtProcedure2 = IIf(Len(!CASOPROCRequired2), !CASOPROCRequired2, "")
                frmMMAEntry.txtProcedure3 = IIf(Len(!CASOPROCRequired3), !CASOPROCRequired3, "")
                frmMMAEntry.txtProcedure4 = IIf(Len(!CASOPROCRequired4), !CASOPROCRequired4, "")
                frmMMAEntry.txtProcedure5 = IIf(Len(!CASOPROCRequired5), !CASOPROCRequired5, "")
                frmMMAEntry.txtProcedure6 = IIf(Len(!CASOPROCRequired6), !CASOPROCRequired6, "")
                frmMMAEntry.txtProcCode1 = IIf(Len(!CASCptCode1), !CASCptCode1, "")
                frmMMAEntry.txtProcCode2 = IIf(Len(!CASCptCode2), !CASCptCode2, "")
                frmMMAEntry.txtProcCode3 = IIf(Len(!CASCptCode3), !CASCptCode3, "")
                frmMMAEntry.txtProcCode4 = IIf(Len(!CASCptCode4), !CASCptCode4, "")
                frmMMAEntry.txtProcCode5 = IIf(Len(!CASCptCode5), !CASCptCode5, "")
                frmMMAEntry.txtProcCode6 = IIf(Len(!CASCptCode6), !CASCptCode6, "")
                
                If Len(!CASCptCode1) Then
                    Call SearchMMARate(!CASCptCode1)
                    frmMMAEntry.txtSMMARate1 = IIf(Len(MMASurgFees), MMASurgFees, 0)
                    frmMMAEntry.txtAnaesMMARate1 = IIf(Len(MMAAnaeFees), MMAAnaeFees, 0)
                    frmMMAEntry.txtS13Rate1 = IIf(Len(MMASurgFees13th), MMASurgFees13th, 0)
                    frmMMAEntry.txtAnae13Rate1 = IIf(Len(MMAAnaeFees13th), MMAAnaeFees13th, 0)
                    frmMMAEntry.txtS13MinRate1 = IIf(Len(MMASurgFees13thMin), MMASurgFees13thMin, 0)
                End If
                
                If Len(!CASCptCode2) Then
                    Call SearchMMARate(!CASCptCode2)
                    frmMMAEntry.txtSMMARate2 = IIf(Len(MMASurgFees), MMASurgFees, 0)
                    frmMMAEntry.txtAnaesMMARate2 = IIf(Len(MMAAnaeFees), MMAAnaeFees, 0)
                    frmMMAEntry.txtS13Rate2 = IIf(Len(MMASurgFees13th), MMASurgFees13th, 0)
                    frmMMAEntry.txtAnae13Rate2 = IIf(Len(MMAAnaeFees13th), MMAAnaeFees13th, 0)
                    frmMMAEntry.txtS13MinRate2 = IIf(Len(MMASurgFees13thMin), MMASurgFees13thMin, 0)
                    
                End If
                
                If Len(!CASCptCode3) Then
                    Call SearchMMARate(!CASCptCode3)
                    frmMMAEntry.txtSMMARate3 = IIf(Len(MMASurgFees), MMASurgFees, 0)
                    frmMMAEntry.txtAnaesMMARate3 = IIf(Len(MMAAnaeFees), MMAAnaeFees, 0)
                    frmMMAEntry.txtS13Rate3 = IIf(Len(MMASurgFees13th), MMASurgFees13th, 0)
                    frmMMAEntry.txtAnae13Rate3 = IIf(Len(MMAAnaeFees13th), MMAAnaeFees13th, 0)
                    frmMMAEntry.txtS13MinRate3 = IIf(Len(MMASurgFees13thMin), MMASurgFees13thMin, 0)
                    
                End If
                
                If Len(!CASCptCode4) Then
                    Call SearchMMARate(!CASCptCode4)
                    frmMMAEntry.txtSMMARate4 = IIf(Len(MMASurgFees), MMASurgFees, 0)
                    frmMMAEntry.txtAnaesMMARate4 = IIf(Len(MMAAnaeFees), MMAAnaeFees, 0)
                    frmMMAEntry.txtS13Rate4 = IIf(Len(MMASurgFees13th), MMASurgFees13th, 0)
                    frmMMAEntry.txtAnae13Rate4 = IIf(Len(MMAAnaeFees13th), MMAAnaeFees13th, 0)
                    frmMMAEntry.txtS13MinRate4 = IIf(Len(MMASurgFees13thMin), MMASurgFees13thMin, 0)
                End If
                
                If Len(!CASCptCode5) Then
                    Call SearchMMARate(!CASCptCode5)
                    frmMMAEntry.txtSMMARate5 = IIf(Len(MMASurgFees), MMASurgFees, 0)
                    frmMMAEntry.txtAnaesMMARate5 = IIf(Len(MMAAnaeFees), MMAAnaeFees, 0)
                    frmMMAEntry.txtS13Rate5 = IIf(Len(MMASurgFees13th), MMASurgFees13th, 0)
                    frmMMAEntry.txtAnae13Rate5 = IIf(Len(MMAAnaeFees13th), MMAAnaeFees13th, 0)
                    frmMMAEntry.txtS13MinRate5 = IIf(Len(MMASurgFees13thMin), MMASurgFees13thMin, 0)
                End If
                
                If Len(!CASCptCode6) Then
                    Call SearchMMARate(!CASCptCode6)
                    frmMMAEntry.txtSMMARate6 = IIf(Len(MMASurgFees), MMASurgFees, 0)
                    frmMMAEntry.txtAnaesMMARate6 = IIf(Len(MMAAnaeFees), MMAAnaeFees, 0)
                    frmMMAEntry.txtS13Rate6 = IIf(Len(MMASurgFees13th), MMASurgFees13th, 0)
                    frmMMAEntry.txtAnae13Rate6 = IIf(Len(MMAAnaeFees13th), MMAAnaeFees13th, 0)
                    frmMMAEntry.txtS13MinRate6 = IIf(Len(MMASurgFees13thMin), MMASurgFees13thMin, 0)
                End If
                
                frmMMAEntry.txtSActFee1 = IIf(Len(!CASProcActSurgFees1), !CASProcActSurgFees1, "")
                frmMMAEntry.txtSActFee2 = IIf(Len(!CASProcActSurgFees2), !CASProcActSurgFees2, "")
                frmMMAEntry.txtSActFee3 = IIf(Len(!CASProcActSurgFees3), !CASProcActSurgFees3, "")
                frmMMAEntry.txtSActFee4 = IIf(Len(!CASProcActSurgFees4), !CASProcActSurgFees4, "")
                frmMMAEntry.txtSActFee5 = IIf(Len(!CASProcActSurgFees5), !CASProcActSurgFees5, "")
                frmMMAEntry.txtSActFee6 = IIf(Len(!CASProcActSurgFees6), !CASProcActSurgFees6, "")
                
                frmMMAEntry.txtAnaesActFees1 = IIf(Len(!CASProcActAnaeFees1), !CASProcActAnaeFees1, "")
                frmMMAEntry.txtAnaesActFees2 = IIf(Len(!CASProcActAnaeFees2), !CASProcActAnaeFees2, "")
                frmMMAEntry.txtAnaesActFees3 = IIf(Len(!CASProcActAnaeFees3), !CASProcActAnaeFees3, "")
                frmMMAEntry.txtAnaesActFees4 = IIf(Len(!CASProcActAnaeFees4), !CASProcActAnaeFees4, "")
                frmMMAEntry.txtAnaesActFees5 = IIf(Len(!CASProcActAnaeFees5), !CASProcActAnaeFees5, "")
                frmMMAEntry.txtAnaesActFees6 = IIf(Len(!CASProcActAnaeFees6), !CASProcActAnaeFees6, "")
                'Non-Surgical Proc
                frmMMAEntry.txtNSProcedure1 = IIf(Len(!CASONonSurgPROCRequired1), !CASONonSurgPROCRequired1, "")
                frmMMAEntry.txtNSProcedure2 = IIf(Len(!CASONonSurgPROCRequired2), !CASONonSurgPROCRequired2, "")
                frmMMAEntry.txtNSProcedure3 = IIf(Len(!CASONonSurgPROCRequired3), !CASONonSurgPROCRequired3, "")
                frmMMAEntry.txtNSProcedure4 = IIf(Len(!CASONonSurgPROCRequired4), !CASONonSurgPROCRequired4, "")
                frmMMAEntry.txtNSProcedure5 = IIf(Len(!CASONonSurgPROCRequired5), !CASONonSurgPROCRequired5, "")
                frmMMAEntry.txtNSProcedure6 = IIf(Len(!CASONonSurgPROCRequired6), !CASONonSurgPROCRequired6, "")
                
                frmMMAEntry.txtNSProcCode1 = IIf(Len(!CASNonSurgCptCode1), !CASNonSurgCptCode1, "")
                frmMMAEntry.txtNSProcCode2 = IIf(Len(!CASNonSurgCptCode2), !CASNonSurgCptCode2, "")
                frmMMAEntry.txtNSProcCode3 = IIf(Len(!CASNonSurgCptCode3), !CASNonSurgCptCode3, "")
                frmMMAEntry.txtNSProcCode4 = IIf(Len(!CASNonSurgCptCode4), !CASNonSurgCptCode4, "")
                frmMMAEntry.txtNSProcCode5 = IIf(Len(!CASNonSurgCptCode5), !CASNonSurgCptCode5, "")
                frmMMAEntry.txtNSProcCode6 = IIf(Len(!CASNonSurgCptCode6), !CASNonSurgCptCode6, "")
                
                If Len(!CASNonSurgCptCode1) Then
                    Call SearchMMARate(!CASNonSurgCptCode1)
                    frmMMAEntry.txtNSMMARate1 = IIf(Len(MMASurgFees), MMASurgFees, 0)
                    frmMMAEntry.txtNS13Rate1 = IIf(Len(MMASurgFees13th), MMASurgFees13th, 0)
                    frmMMAEntry.txtNS13MinRate1 = IIf(Len(MMASurgFees13thMin), MMASurgFees13thMin, 0)
                    
                End If

                If Len(!CASNonSurgCptCode2) Then
                    Call SearchMMARate(!CASNonSurgCptCode2)
                    frmMMAEntry.txtNSMMARate2 = IIf(Len(MMASurgFees), MMASurgFees, 0)
                    frmMMAEntry.txtNS13MinRate2 = IIf(Len(MMASurgFees13thMin), MMASurgFees13thMin, 0)
                End If
                
                If Len(!CASNonSurgCptCode3) Then
                    Call SearchMMARate(!CASNonSurgCptCode3)
                    frmMMAEntry.txtNSMMARate3 = IIf(Len(MMASurgFees), MMASurgFees, 0)
                    frmMMAEntry.txtNS13MinRate3 = IIf(Len(MMASurgFees13thMin), MMASurgFees13thMin, 0)
                End If
                
                If Len(!CASNonSurgCptCode4) Then
                    Call SearchMMARate(!CASNonSurgCptCode4)
                    frmMMAEntry.txtNSMMARate4 = IIf(Len(MMASurgFees), MMASurgFees, 0)
                    frmMMAEntry.txtNS13MinRate4 = IIf(Len(MMASurgFees13thMin), MMASurgFees13thMin, 0)
                End If
                
                If Len(!CASNonSurgCptCode5) Then
                    Call SearchMMARate(!CASNonSurgCptCode5)
                    frmMMAEntry.txtNSMMARate5 = IIf(Len(MMASurgFees), MMASurgFees, 0)
                    frmMMAEntry.txtNS13MinRate5 = IIf(Len(MMASurgFees13thMin), MMASurgFees13thMin, 0)
                End If
                
                If Len(!CASNonSurgCptCode6) Then
                    Call SearchMMARate(!CASNonSurgCptCode6)
                    frmMMAEntry.txtNSMMARate6 = IIf(Len(MMASurgFees), MMASurgFees, 0)
                    frmMMAEntry.txtNS13MinRate6 = IIf(Len(MMASurgFees13thMin), MMASurgFees13thMin, 0)
                End If
                
                frmMMAEntry.txtNSActFee1 = IIf(Len(!CASProcActNonSurgFees1), !CASProcActNonSurgFees1, "")
                frmMMAEntry.txtNSActFee2 = IIf(Len(!CASProcActNonSurgFees2), !CASProcActNonSurgFees2, "")
                frmMMAEntry.txtNSActFee3 = IIf(Len(!CASProcActNonSurgFees3), !CASProcActNonSurgFees3, "")
                frmMMAEntry.txtNSActFee4 = IIf(Len(!CASProcActNonSurgFees4), !CASProcActNonSurgFees4, "")
                frmMMAEntry.txtNSActFee5 = IIf(Len(!CASProcActNonSurgFees5), !CASProcActNonSurgFees5, "")
                frmMMAEntry.txtNSActFee6 = IIf(Len(!CASProcActNonSurgFees6), !CASProcActNonSurgFees6, "")
                
                frmMMAEntry.LblCasNUmber = txtCaseNo
                
                If !PAYChkHosp = "Y" Then
                    tmpHospChk = "Y"
                Else
                    tmpHospChk = ""
                End If
            End With
            frmMMAEntry.CalSource = 0
            Call ShowInvoice
            Call getVersion
            'Call LGListing
            If PLNSOF = True Then
                MsgBox "Please refer schedule of fee! ", vbCritical
            End If
                        
            Else
                MsgBox "Record not found! ", vbCritical
            End If
            ADOrstCAS.Close
            Set ADOrstCAS = Nothing
            
            If (Trim(PAYCode) = "IB") Then
                Label103.Caption = "Maternity"
                Label21.Caption = "Psychiatric"
            Else
                Label103.Caption = "Home Nursing Care"
                Label21.Caption = "Outpatient Physiotherapy Treatment"
            End If
            
            
        Else
            MsgBox "Case Number cannot be empty! ", vbCritical
        End If

    End If
    
 '   medixmasterconn.Close
 '   Set medixmasterconn = Nothing
  '  medixmasterconnMaint.Close
 '   Set medixmasterconnMaint = Nothing
  '  medixmasterconnWS.Close
  '  Set medixmasterconnWS = Nothing
    cmdShow.Enabled = True
    Screen.MousePointer = Default
End Sub

Private Sub getVersion()
    Dim Version As New ADODB.Recordset
    Dim strsql As String
    strsql = "select casNumber, max(claimversion) AS lastVersionNo from claim where casNUmber ='" & Trim(txtCaseNo) & "' " & _
            " and ClaimStatus ='C' " & _
            "group by casNumber"
    Version.Open strsql, medixmasterconnWS, adOpenKeyset, adLockOptimistic
    If Not Version.EOF Then
        lblLastVersion = Version!lastVersionNo + 1
    Else
        lblLastVersion = 1
    End If
    Version.Close
    Set Version = Nothing
End Sub

Public Function GetHosName(HOSCode As String) As String
    Dim HOSSql As String
    Dim rst3 As New ADODB.Recordset
    HOSSql = " select HOSCode, HOSNAme  from HOS where HOSCOde='" & Trim(HOSCode) & "'"
    
    rst3.Open HOSSql, medixmasterconn, adOpenKeyset, adLockOptimistic
    If rst3.recordCount Then
        GetHosName = rst3!HosName
    Else
        MsgBox "Hospital Name not found!", vbCritical + vbOKOnly
    End If
    rst3.Close
    Set rst3 = Nothing
End Function

Public Function ShowINVRemarks()
    Dim sql As String
    Dim rst As New ADODB.Recordset
    
    txtRemarks = ""
   ' medixmasterconnWS.Open "File Name=" & BOSPath & "Conn\mediconnectWS.UDL"
    
    sql = " select CASNumber, Remarks  from ClaimInvoiceRemarks where CASNumber ='" & Trim(txtCaseNo) & "'"
    rst.Open sql, medixmasterconnWS, adOpenKeyset, adLockOptimistic
    
    If rst.recordCount Then
        If Len(rst!Remarks) Then
            txtRemarks = rst!Remarks
        End If
    End If
    rst.Close
    Set rst = Nothing
    
  '  medixmasterconnWS.Close
   ' Set medixmasterconnWS = Nothing
End Function

Public Function SaveINVRemarks()
    Dim sql As String
    Dim rst As New ADODB.Recordset
    
    sql = " select CASNumber, Remarks  from ClaimInvoiceRemarks where CASNumber ='" & Trim(txtCaseNo) & "'"
    rst.Open sql, medixmasterconnWS, adOpenKeyset, adLockOptimistic
    
    If rst.recordCount = 0 Then
        rst.AddNew
    End If
    
    rst!CasNumber = txtCaseNo
    If Len(txtRemarks) Then
        rst!Remarks = txtRemarks
    Else
        rst!Remarks = ""
    End If
    rst.Update
    
    rst.Close
    Set rst = Nothing
End Function

Private Sub ShowInvoice()
    Dim Invoice As New ADODB.Recordset
    Dim INVList As ListItem
    Dim strsql As String
    Dim strsqlHOS As String
    Dim HOS As New ADODB.Recordset
    
    Bordxstatus = True
    
    strsql = "Exec SearchCASInvoice_08 '" & Trim(txtCaseNo) & "'"
    Set Invoice = New ADODB.Recordset
    Invoice.Open strsql, medixmasterconnWS, adOpenKeyset, adLockOptimistic
    
    If Not Invoice.EOF Then
    
        lstInvoice.ListItems.Clear
        Do While Not Invoice.EOF
            Set INVList = lstInvoice.ListItems.Add(, , Invoice!INVIndex)
            With INVList
                If Trim(Invoice!INVBordxed) <> "" Then
                    .SubItems(1) = Invoice!INVBordxed
                End If
                
                If Invoice!INVBordxed = "False" And Invoice!INVWSed = "True" Then
                    Bordxstatus = False
                End If
                
                If Trim(Invoice!INVWSed) <> "" Then
                    .SubItems(2) = Invoice!INVWSed
                End If
                
                If Trim(Invoice!INVWSversion) <> "" Then
                    .SubItems(3) = Invoice!INVWSversion
                End If
                
                If Trim(Invoice!INVNo) <> "" Then
                    .SubItems(4) = Invoice!INVNo
                End If
                If Trim(Invoice!InvDate) <> "" Then
                    .SubItems(5) = Invoice!InvDate
                End If
                If Trim(Invoice!INVAmount) <> "" Then
                    .SubItems(6) = Invoice!INVAmount
                End If
                If Trim(Invoice!InvPayTo) <> "" Then
                    .SubItems(7) = Invoice!InvPayTo
                End If
                 
                strsqlHOS = "select HOsCode, HOSName from HOS Where HOSCode = '" & Trim(Invoice!INVPayee) & "'"
                HOS.Open strsqlHOS, medixmasterconn, adOpenKeyset, adLockOptimistic
                If HOS.recordCount Then
                    .SubItems(8) = HOS!HosName
                Else
                    .SubItems(8) = "Member"
                End If
                HOS.Close
                Set HOS = Nothing
                
                If Trim(Invoice!ClaimType) <> "" Then
                    .SubItems(9) = Invoice!ClaimType
                End If
                               
                If Trim(Invoice!INVFolUpDate) <> "" Then
                    .SubItems(10) = Invoice!INVFolUpDate
                End If
                
                If Trim(Invoice!CLMBoardDay) <> "" Then
                    .SubItems(11) = Invoice!CLMBoardDay
                End If
                
                If Trim(Invoice!CLMRB) <> "" Then
                    .SubItems(12) = Invoice!CLMRB
                End If
                If Trim(Invoice!CLMICUDay) <> "" Then
                    .SubItems(13) = Invoice!CLMICUDay
                End If
                If Trim(Invoice!CLMICU) <> "" Then
                    .SubItems(14) = Invoice!CLMICU
                End If
                If Trim(Invoice!CLMServices) <> "" Then
                    .SubItems(15) = Invoice!CLMServices
                End If
                If Trim(Invoice!CLMTheatre) <> "" Then
                    .SubItems(16) = Invoice!CLMTheatre
                End If
                If Trim(Invoice!CLMPreHSAmt) <> "" Then
                    .SubItems(17) = Invoice!CLMPreHSAmt
                End If
                If Trim(Invoice!CLMPhysician) <> "" Then
                    .SubItems(18) = Invoice!CLMPhysician
                End If
                If Trim(Invoice!CLMPost) <> "" Then
                    .SubItems(19) = Invoice!CLMPost
                End If
                If Trim(Invoice!CLMMedical) <> "" Then
                    .SubItems(20) = Invoice!CLMMedical
                End If
                If Trim(Invoice!CLMPhone) <> "" Then
                    .SubItems(21) = Invoice!CLMPhone
                End If
                If Trim(Invoice!CLMLodgerday) <> "" Then
                    .SubItems(22) = Invoice!CLMLodgerday
                End If
                If Trim(Invoice!CLMLodger) <> "" Then
                    .SubItems(23) = Invoice!CLMLodger
                End If
                If Trim(Invoice!CLMAdmin) <> "" Then
                    .SubItems(24) = Invoice!CLMAdmin
                End If
                If Trim(Invoice!CLMTax) <> "" Then
                    .SubItems(25) = Invoice!CLMTax
                End If
                If Trim(Invoice!CLMOthers1) <> "" Then
                    .SubItems(26) = Invoice!CLMOthers1
                End If
                If Trim(Invoice!CLMCashDay) <> "" Then
                    .SubItems(27) = Invoice!CLMCashDay
                End If
                If Trim(Invoice!CLMCash) <> "" Then
                    .SubItems(28) = Invoice!CLMCash
                End If
                If Trim(Invoice!CLMPreXRay) <> "" Then
                    .SubItems(29) = Invoice!CLMPreXRay
                End If
                If Trim(Invoice!CLMPreLab) <> "" Then
                    .SubItems(30) = Invoice!CLMPreLab
                End If
                If Trim(Invoice!CLMPreoTHERS) <> "" Then
                    .SubItems(31) = Invoice!CLMPreoTHERS
                End If
                If Trim(Invoice!CLMSurgicalAmt) <> "" Then
                    .SubItems(32) = Invoice!CLMSurgicalAmt
                End If
                If Trim(Invoice!CLMAna) <> "" Then
                    .SubItems(33) = Invoice!CLMAna
                End If
                If Trim(Invoice!CLMEmergency) <> "" Then
                    .SubItems(34) = Invoice!CLMEmergency
                End If
                If Trim(Invoice!CLMAmbulance) <> "" Then
                    .SubItems(35) = Invoice!CLMAmbulance
                End If
                If Trim(Invoice!CLMOutpatient) <> "" Then
                    .SubItems(36) = Invoice!CLMOutpatient
                End If
                If Trim(Invoice!CLMMonthly) <> "" Then
                    .SubItems(37) = Invoice!CLMMonthly
                End If
                If Trim(Invoice!PreH60) <> "" Then
                    .SubItems(38) = Invoice!PreH60
                End If
                If Trim(Invoice!PreMedic) <> "" Then
                    .SubItems(40) = Invoice!PreMedic
                End If
                If Trim(Invoice!PreGP) <> "" Then
                    .SubItems(39) = Invoice!PreGP
                End If
                If Trim(Invoice!PreSpec) <> "" Then
                    .SubItems(42) = Invoice!PreSpec
                End If
                If Trim(Invoice!PostHosp) <> "" Then
                    .SubItems(43) = Invoice!PostHosp
                End If
                If Trim(Invoice!PreMR) <> "" Then
                    .SubItems(41) = Invoice!PreMR
                End If
                If Trim(Invoice!PostNotSame) <> "" Then
                    .SubItems(44) = Invoice!PostNotSame
                End If
                If Trim(Invoice!Others) <> "" Then
                    .SubItems(45) = Invoice!Others
                End If
                If Trim(Invoice!INVPaidByMem) <> "" Then
                    .SubItems(46) = Invoice!INVPaidByMem
                End If
                
                If Trim(Invoice!CLMSurgPhy) <> "" Then
                    .SubItems(47) = Invoice!CLMSurgPhy
                End If
                If Trim(Invoice!CLMSurgPost) <> "" Then
                    .SubItems(48) = Invoice!CLMSurgPost
                End If
                If Trim(Invoice!CLMOther1Desc) <> "" Then
                    .SubItems(49) = Invoice!CLMOther1Desc
                End If
                If Trim(Invoice!CLMOthers2Desc) <> "" Then
                    .SubItems(50) = Invoice!CLMOthers2Desc
                End If
                If Trim(Invoice!CLMOthers2) <> "" Then
                    .SubItems(51) = Invoice!CLMOthers2
                End If
                If Trim(Invoice!CLMCashAmt) <> "" Then
                    .SubItems(52) = Invoice!CLMCashAmt
                End If
                If Trim(Invoice!CLMGuardianAmt) <> "" Then
                    .SubItems(53) = Invoice!CLMGuardianAmt
                End If
                If Trim(Invoice!CLMBoardAmt) <> "" Then
                    .SubItems(54) = Invoice!CLMBoardAmt
                End If
                If Trim(Invoice!CLMICUAmt) <> "" Then
                    .SubItems(55) = Invoice!CLMICUAmt
                End If
                If Trim(Invoice!CLMKidneyMnth) <> "" Then
                    .SubItems(56) = Invoice!CLMKidneyMnth
                End If
                If Trim(Invoice!CLMSurgicalFee) <> "" Then
                    .SubItems(57) = Invoice!CLMSurgicalFee
                End If
                If Trim(Invoice!CLMNSPreHos) <> "" Then
                    .SubItems(58) = Invoice!CLMNSPreHos
                End If
                
                If Trim(Invoice!CLMINHospHyVisitDay) <> "" Then
                    .SubItems(59) = Invoice!CLMINHospHyVisitDay
                End If
                
                If Trim(Invoice!CLMOrgan) <> "" Then
                    .SubItems(60) = Invoice!CLMOrgan
                End If
                
                If Trim(Invoice!CLMRBTax) <> "" Then
                    .SubItems(61) = Invoice!CLMRBTax
                End If

                If Trim(Invoice!CLMICUTax) <> "" Then
                    .SubItems(62) = Invoice!CLMICUTax
                End If

                If Trim(Invoice!CLMLodgerTax) <> "" Then
                    .SubItems(63) = Invoice!CLMLodgerTax
                End If
                
                If Trim(Invoice!INVStayType) <> "" Then
                    .SubItems(64) = Invoice!INVStayType
                End If
                                
                If Trim(Invoice!INVPAYType) <> "" Then
                    .SubItems(65) = Invoice!INVPAYType
                End If
                If Trim(Invoice!INVEnteredBy) <> "" Then
                    .SubItems(66) = Invoice!INVEnteredBy
                End If
                If Trim(Invoice!INVEnteredDate) <> "" Then
                    .SubItems(67) = Invoice!INVEnteredDate
                End If
                If Trim(Invoice!CLMMRI) <> "" Then
                    .SubItems(68) = Invoice!CLMMRI
                End If
                If Trim(Invoice!CLMMXRay) <> "" Then
                    .SubItems(69) = Invoice!CLMMXRay
                End If
                If Trim(Invoice!CLMMLab) <> "" Then
                    .SubItems(70) = Invoice!CLMMLab
                End If
                If Trim(Invoice!CLMMPreOthers) <> "" Then
                    .SubItems(71) = Invoice!CLMMPreOthers
                End If
                If Trim(Invoice!CLMMPreHosAmt) <> "" Then
                    .SubItems(72) = Invoice!CLMMPreHosAmt
                End If
                If Trim(Invoice!CLMPreHosp) <> "" Then
                    .SubItems(73) = Invoice!CLMPreHosp
                End If
                If Trim(Invoice!CLMPostTreat) <> "" Then
                    .SubItems(74) = Invoice!CLMPostTreat
                End If
                If Trim(Invoice!CLMAccPostAmt) <> "" Then
                    .SubItems(75) = Invoice!CLMAccPostAmt
                End If
                If Trim(Invoice!CLMNewRBTax) <> "" Then
                    .SubItems(76) = Invoice!CLMNewRBTax
                End If
                If Trim(Invoice!CLMNewICUTax) <> "" Then
                    .SubItems(77) = Invoice!CLMNewICUTax
                End If
                If Trim(Invoice!CLMNewLodgerTax) <> "" Then
                    .SubItems(78) = Invoice!CLMNewLodgerTax
                End If
                If Trim(Invoice!CLMTotalBoardAmt) <> "" Then
                    .SubItems(79) = Invoice!CLMTotalBoardAmt
                End If
                
                If Trim(Invoice!CLMTotalICUAmt) <> "" Then
                    .SubItems(80) = Invoice!CLMTotalICUAmt
                End If
                
                If Trim(Invoice!CLMTotalGAllowAmt) <> "" Then
                    .SubItems(81) = Invoice!CLMTotalGAllowAmt
                End If
                
                If Trim(Invoice!CLMSurgicalWard) <> "" Then
                    .SubItems(82) = Invoice!CLMSurgicalWard
                End If
                
                If Trim(Invoice!CLMAnaeFees) <> "" Then
                    .SubItems(83) = Invoice!CLMAnaeFees
                End If
                
                If Trim(Invoice!CLMAnaeConsultation) <> "" Then
                    .SubItems(84) = Invoice!CLMAnaeConsultation
                End If
                
                If Trim(Invoice!CLMAnaeWard) <> "" Then
                    .SubItems(85) = Invoice!CLMAnaeWard
                End If
                
                If Trim(Invoice!CLMInHosFees) <> "" Then
                    .SubItems(86) = Invoice!CLMInHosFees
                End If
                
                If Trim(Invoice!CLMInHosWard) <> "" Then
                    .SubItems(87) = Invoice!CLMInHosWard
                End If
                
                
                If Trim(Invoice!CLMMedRptMBM) <> "" Then
                    .SubItems(88) = Invoice!CLMMedRptMBM
                End If
                
                If Trim(Invoice!CLMTelChargesMBM) <> "" Then
                    .SubItems(89) = Invoice!CLMTelChargesMBM
                End If
                
                If Trim(Invoice!CLMGuaAllowMBM) <> "" Then
                    .SubItems(90) = Invoice!CLMGuaAllowMBM
                End If
                
                If Trim(Invoice!CLMAdminRegMBM) <> "" Then
                    .SubItems(91) = Invoice!CLMAdminRegMBM
                End If
                
                If Trim(Invoice!CLMOthersMBM) <> "" Then
                    .SubItems(92) = Invoice!CLMOthersMBM
                End If
                
                If Trim(Invoice!CLMGrossRB) <> "" Then
                    .SubItems(93) = Invoice!CLMGrossRB
                End If
                
                If Trim(Invoice!CLMOthersNotCoPay) <> "" Then
                    .SubItems(94) = Invoice!CLMOthersNotCoPay
                End If
                
                If Trim(Invoice!CLMOthersDescNoCOPay) <> "" Then
                    .SubItems(95) = Invoice!CLMOthersDescNoCOPay
                End If
                
                If Trim(Invoice!CLMProceduresFee) <> "" Then
                    .SubItems(96) = Invoice!CLMProceduresFee
                End If

                If Trim(Invoice!CLMINVRecDate) <> "" Then
                    .SubItems(97) = Invoice!CLMINVRecDate
                End If
                
                If Trim(Invoice!CLMINVRecMode) <> "" Then
                    .SubItems(98) = Invoice!CLMINVRecMode
                End If

                If Trim(Invoice!CLMHomeNCAmt) <> "" Then
                    .SubItems(99) = Invoice!CLMHomeNCAmt
                End If
                
                If Trim(Invoice!CLMUHCNo) <> "" Then
                    .SubItems(100) = Invoice!CLMUHCNo
                End If

                If Trim(Invoice!InvExcessAmt) <> "" Then
                    .SubItems(101) = Invoice!InvExcessAmt
                End If
                
                If Trim(Invoice!InvPayAmt) <> "" Then
                    .SubItems(102) = Invoice!InvPayAmt
                End If

                If Trim(Invoice!ReimOtherRemarks) <> "" Then
                    .SubItems(103) = Invoice!ReimOtherRemarks
                End If
                
                .SubItems(104) = IIf(IsNumeric(Invoice!CLMPharmacyHSS), Invoice!CLMPharmacyHSS, "")
                .SubItems(105) = IIf(IsNumeric(Invoice!CLMMedSuppliesHSS), Invoice!CLMMedSuppliesHSS, "")
                .SubItems(106) = IIf(IsNumeric(Invoice!CLMNursingCareHSS), Invoice!CLMNursingCareHSS, "")
                .SubItems(107) = IIf(IsNumeric(Invoice!CLMNursingSvrHSS), Invoice!CLMNursingSvrHSS, "")
                .SubItems(108) = IIf(IsNumeric(Invoice!CLMLabHSS), Invoice!CLMLabHSS, "")
                .SubItems(109) = IIf(IsNumeric(Invoice!CLMRadiologiHSS), Invoice!CLMRadiologiHSS, "")
                .SubItems(110) = IIf(IsNumeric(Invoice!CLMOTSuppHSS), Invoice!CLMOTSuppHSS, "")
                .SubItems(111) = IIf(IsNumeric(Invoice!CLMEquipHSS), Invoice!CLMEquipHSS, "")
                .SubItems(112) = IIf(IsNumeric(Invoice!CLMMiscHSS), Invoice!CLMMiscHSS, "")
                .SubItems(113) = IIf(IsNumeric(Invoice!INVInitialActAmt), Invoice!INVInitialActAmt, "")
                .SubItems(114) = IIf(Len(Invoice!INVHospName), Invoice!INVHospName, "")
                .SubItems(115) = IIf(Len(Invoice!CLMCancerAmt), Invoice!CLMCancerAmt, "")
                                                                    
            End With
            Invoice.MoveNext
        Loop
    End If
    Invoice.Close
    Set Invoice = Nothing
        
End Sub

' ## ********************** start Button
Private Sub CmdTotal_Click()
    Dim ForLoop, ListCount As Integer
    TotalClick = True
    CmdEdit.Enabled = False
    CmdAdd.Enabled = True
    CmdRemove.Enabled = False
    ListCount = lstInvoice.ListItems.Count
    txtInvNo = ""
    txtFolDate.Visible = False
    txtFolDate2.Visible = True
    txtFolDate2 = ""
    txtInvDate.Visible = False
    txtINVDate2.Visible = True
    txtINVDate2 = ""
    
    txtOthersDesc1 = ""
    txtOthersDesc2 = ""
    txtINVAmount = 0
    RoomAmount = "-"
    ICUAmount = "-"
    txtLodgerAmt = "-"
    txtRMCash = "-"
    
    txtINVAmount = 0
    TxtAccess = 0
    ' Hospital
    DB1 = 0
    A1 = 0
    ICU1 = 0
    A2 = 0
    A3 = 0
    A4 = 0
    ' Surgical
    txtXray = 0
    txtPreLab = 0
    txtPreOthers = 0
    A5 = 0
    Text1 = 0
    
    A24 = 0
    
    TxtMXRay = 0
    TxtMLab = 0
    TxtMPreOthers = 0
    A5B = 0
    Text2 = 0
    
    txtSurgicalFee = 0
    txtPhysician = 0
    txtPostTreat = 0
    A6 = 0
    A7 = 0
    
    ' Non Surgical
    A8 = 0
    A9 = 0
    A10 = 0
    
    ' Out- pAtient
    A11 = 0
    txtACC = 0
    txtPTreat = 0
    A12 = 0
    A13 = 0
    txtKidneyMonths = 0
    A14 = 0
    A26 = 0
    
    ' Others
    A15 = 0
    A16 = 0
    txtLodgerDay = 0
    A17 = 0
    A18 = 0
    A19 = 0
    A20 = 0
    A20B = 0
    txtDayCash = 0
    A21 = 0
    A22 = 0
    A25 = 0
    
    ' Not COver
    txtPreHos60 = 0
    txtPreGP = 0
    txtPreMEdic = 0
    txtPreMR = 0
    txtPreSpec1 = 0
    txtPostGHosp = 0
    txtpostNotSame = 0
    txtOthers = 0
    
    For ForLoop = 1 To ListCount
        With lstInvoice.ListItems(ForLoop)
            If Trim(.SubItems(4)) <> "" Then
                If Trim(txtInvNo) <> "" Then
                    txtInvNo = txtInvNo & "," & .SubItems(4)
                Else
                    txtInvNo = .SubItems(4)
                End If
            End If
            If Trim(.SubItems(10)) <> "" Then
                If Trim(txtFolDate2) <> "" Then
                    txtFolDate2 = txtFolDate2 & "," & .SubItems(10)
                Else
                    txtFolDate2 = .SubItems(10)
                End If
            End If
            
            If Trim(.SubItems(5)) <> "" Then
                If Trim(txtINVDate2) <> "" Then
                    txtINVDate2 = txtINVDate2 & "," & .SubItems(5)
                Else
                    txtINVDate2 = .SubItems(5)
                End If
            End If
            
            txtINVAmount = CDbl(txtINVAmount) + CDbl(IIf(IsNumeric(.SubItems(6)), .SubItems(6), 0))
            TxtAccess = CDbl(TxtAccess) + CDbl(IIf(IsNumeric(.SubItems(46)), .SubItems(46), 0))
            
            ' Hospital
            DB1 = CDbl(DB1) + CDbl(IIf(IsNumeric(.SubItems(11)), .SubItems(11), 0))
            A1 = CDbl(A1) + CDbl(IIf(IsNumeric(.SubItems(54)), .SubItems(54), 0))
            ICU1 = CDbl(ICU1) + CDbl(IIf(IsNumeric(.SubItems(13)), .SubItems(13), 0))
            A2 = CDbl(A2) + CDbl(IIf(IsNumeric(.SubItems(55)), .SubItems(55), 0))
            A3 = CDbl(A3) + CDbl(IIf(IsNumeric(.SubItems(15)), .SubItems(15), 0))
            A4 = CDbl(A4) + CDbl(IIf(IsNumeric(.SubItems(16)), .SubItems(16), 0))
            ' Surgical
            txtXray = CDbl(txtXray) + CDbl(IIf(IsNumeric(.SubItems(29)), .SubItems(29), 0))
            txtPreLab = CDbl(txtPreLab) + CDbl(IIf(IsNumeric(.SubItems(30)), .SubItems(30), 0))
            txtPreOthers = CDbl(txtPreOthers) + CDbl(IIf(IsNumeric(.SubItems(31)), .SubItems(31), 0))
            A5 = CDbl(A5) + CDbl(IIf(IsNumeric(.SubItems(17)), .SubItems(17), 0))
            
            TxtMXRay = CDbl(TxtMXRay) + CDbl(IIf(IsNumeric(.SubItems(69)), .SubItems(69), 0))
            TxtMLab = CDbl(TxtMLab) + CDbl(IIf(IsNumeric(.SubItems(70)), .SubItems(70), 0))
            TxtMPreOthers = CDbl(TxtMPreOthers) + CDbl(IIf(IsNumeric(.SubItems(71)), .SubItems(71), 0))
            A5B = CDbl(A5B) + CDbl(IIf(IsNumeric(.SubItems(72)), .SubItems(72), 0))
            
            txtSurgicalFee = CDbl(txtSurgicalFee) + CDbl(IIf(IsNumeric(.SubItems(57)), .SubItems(57), 0))
            txtPhysician = CDbl(txtPhysician) + CDbl(IIf(IsNumeric(.SubItems(47)), .SubItems(47), 0))
            txtPostTreat = CDbl(txtPostTreat) + CDbl(IIf(IsNumeric(.SubItems(48)), .SubItems(48), 0))
            A6 = CDbl(A6) + CDbl(IIf(IsNumeric(.SubItems(32)), .SubItems(32), 0))
            A7 = CDbl(A7) + CDbl(IIf(IsNumeric(.SubItems(33)), .SubItems(33), 0))
            
            ' Non Surgical
            A8 = CDbl(A8) + CDbl(IIf(IsNumeric(.SubItems(58)), .SubItems(58), 0))
            A9 = CDbl(A9) + CDbl(IIf(IsNumeric(.SubItems(18)), .SubItems(18), 0))
            A10 = CDbl(A10) + CDbl(IIf(IsNumeric(.SubItems(19)), .SubItems(19), 0))
            
            ' Out- pAtient
            A11 = CDbl(A11) + CDbl(IIf(IsNumeric(.SubItems(34)), .SubItems(34), 0))
            A12 = CDbl(A12) + CDbl(IIf(IsNumeric(.SubItems(35)), .SubItems(35), 0))
            A13 = CDbl(A13) + CDbl(IIf(IsNumeric(.SubItems(36)), .SubItems(36), 0))
            txtKidneyMonths = CDbl(txtKidneyMonths) + CDbl(IIf(IsNumeric(.SubItems(56)), .SubItems(56), 0))
            A14 = CDbl(A14) + CDbl(IIf(IsNumeric(.SubItems(37)), .SubItems(37), 0))
            
            ' Others
            A15 = CDbl(A15) + CDbl(IIf(IsNumeric(.SubItems(20)), .SubItems(20), 0))
            A16 = CDbl(A16) + CDbl(IIf(IsNumeric(.SubItems(21)), .SubItems(21), 0))
            txtLodgerDay = CDbl(txtLodgerDay) + CDbl(IIf(IsNumeric(.SubItems(22)), .SubItems(22), 0))
            A17 = CDbl(A17) + CDbl(IIf(IsNumeric(.SubItems(53)), .SubItems(53), 0))
            A18 = CDbl(A18) + CDbl(IIf(IsNumeric(.SubItems(24)), .SubItems(24), 0))
            A19 = CDbl(A19) + CDbl(IIf(IsNumeric(.SubItems(25)), .SubItems(25), 0))
            A20 = CDbl(A20) + CDbl(IIf(IsNumeric(.SubItems(26)), .SubItems(26), 0))
            If Trim(.SubItems(49)) <> "" Then
                If Trim(txtOthersDesc1) <> "" Then
                    txtOthersDesc1 = txtOthersDesc1 & "," & .SubItems(49)
                Else
                    txtOthersDesc1 = .SubItems(49)
                End If
            End If
            A20B = CDbl(A20B) + CDbl(IIf(IsNumeric(.SubItems(51)), .SubItems(51), 0))
            
            If Trim(.SubItems(50)) <> "" Then
                If Trim(txtOthersDesc2) <> "" Then
                    txtOthersDesc2 = txtOthersDesc2 & "," & .SubItems(50)
                Else
                    txtOthersDesc2 = .SubItems(50)
                End If
            End If
            
            txtDayCash = CDbl(txtDayCash) + CDbl(IIf(IsNumeric(.SubItems(27)), .SubItems(27), 0))
            A21 = CDbl(A21) + CDbl(IIf(IsNumeric(.SubItems(52)), .SubItems(52), 0))
            A22 = CDbl(A22) + CDbl(IIf(IsNumeric(.SubItems(60)), .SubItems(60), 0))
            A25 = CDbl(A25) + CDbl(IIf(IsNumeric(.SubItems(94)), .SubItems(94), 0))
            A26 = CDbl(A26) + CDbl(IIf(IsNumeric(.SubItems(99)), .SubItems(99), 0))
            ' Not COver
            txtPreHos60 = CDbl(txtPreHos60) + CDbl(IIf(IsNumeric(.SubItems(38)), .SubItems(38), 0))
            txtPreGP = CDbl(txtPreGP) + CDbl(IIf(IsNumeric(.SubItems(39)), .SubItems(39), 0))
            txtPreMEdic = CDbl(txtPreMEdic) + CDbl(IIf(IsNumeric(.SubItems(40)), .SubItems(40), 0))
            txtPreMR = CDbl(txtPreMR) + CDbl(IIf(IsNumeric(.SubItems(41)), .SubItems(41), 0))
            txtPreSpec1 = CDbl(txtPreSpec1) + CDbl(IIf(IsNumeric(.SubItems(42)), .SubItems(42), 0))
            txtPostGHosp = CDbl(txtPostGHosp) + CDbl(IIf(IsNumeric(.SubItems(43)), .SubItems(43), 0))
            txtpostNotSame = CDbl(txtpostNotSame) + CDbl(IIf(IsNumeric(.SubItems(44)), .SubItems(44), 0))
            txtOthers = CDbl(txtOthers) + CDbl(IIf(IsNumeric(.SubItems(45)), .SubItems(45), 0))
            
        End With
    Next
    TotalClick = False
End Sub

Private Function getTotal() As Double
    Dim ctl  As Control
    txtINVAmount = 0
    For Each ctl In Me.Controls
        If Mid(ctl.Name, 1, 1) = "A" And Len(ctl.Name) <= 4 Then
            If IsNumeric(ctl.Text) = True Then
                txtINVAmount = txtINVAmount + CDbl(ctl.Text)
            End If
            
        End If
    Next
    If IsNumeric(A1) = False Then
        A1 = 0
    End If
    If IsNumeric(A2) = False Then
        A2 = 0
    End If
    If IsNumeric(ART) = False Then
        ART = 0
    End If
    If IsNumeric(AIT) = False Then
        AIT = 0
    End If
    If IsNumeric(ALT) = False Then
        ALT = 0
    End If
    If IsNumeric(A17) = False Then
        A17 = 0
    End If
    
    txtINVAmount = CDbl(txtINVAmount) - (CDbl(A1) + CDbl(A2) + CDbl(ART) + CDbl(AIT) + CDbl(ALT) + CDbl(A17))
End Function

Private Sub lstInvoice_Click()
    Call clearInvoice
    Load frmHSSDetails
    frmHSSDetails.Visible = False
    Call ClearHSS
      
      If lstInvoice.ListItems.Count > 0 Then
            CmdRemove.Enabled = True
            CmdEdit.Enabled = True
            CmdEdit.Visible = True
            CmdUpdate.Visible = False
            txtCLMVersion.Visible = True
            Label19(1).Visible = True

          With lstInvoice.SelectedItem
            txtIndex = .Text
            intIndex = .Text
            If Trim(.SubItems(1)) = "" Then
                lblBordx = "No"
            Else
                lblBordx = .SubItems(1)
            End If
            
            If Trim(.SubItems(2)) = "" Then
                lblWs = "No"
            Else
                lblWs = .SubItems(2)
            End If
            
            If Trim(.SubItems(3)) = "" Then
                lblVersion = 1
                txtCLMVersion = 1
            Else
                lblVersion = .SubItems(3)
                txtCLMVersion = .SubItems(3)
            End If
            
            If Trim(.SubItems(4)) <> "" Then
                txtInvNo = .SubItems(4)
            End If
            
            If IsDate(.SubItems(5)) = True Then
                txtInvDate = Format(.SubItems(5), "dd/mm/yyyy")
            End If
            
            txtINVAmount = .SubItems(6)
            cboINVPayType = .SubItems(7)
            cboHosCode = .SubItems(8)
            CboPayType = .SubItems(9)
            txtFolDate.mask = ""
            txtFolDate = .SubItems(10)
            txtFolDate.mask = "##/##/####"
            TxtAccess = .SubItems(46)
            ' Hospital
            DB1 = .SubItems(11)
            RoomAmount = .SubItems(12)
            A1 = .SubItems(54)
            ICU1 = .SubItems(13)
            ICUAmount = .SubItems(14)
            A2 = .SubItems(55)
            A2B = .SubItems(80)
            A3 = .SubItems(15)
            
            A4 = .SubItems(16)
            ' Surgical
            txtXray = .SubItems(29)
            txtPreLab = .SubItems(30)
            txtPreOthers = .SubItems(31)
            A5 = .SubItems(17) 'Pre hosp
            A24 = .SubItems(73)
            
            TxtMXRay = .SubItems(69)
            TxtMLab = .SubItems(70)
            TxtMPreOthers = .SubItems(71)
            A5B = .SubItems(72) 'Pre hosp 2
            
            A6 = .SubItems(32) ' Surgical
            txtSurgicalFee = .SubItems(57)
            txtPhysician = .SubItems(47)
            txtPostTreat = .SubItems(48)
            
            A7 = .SubItems(33) ' ANa
            
            'Non -Surgical
            A8 = .SubItems(58)  ' Pre Hosp
            A9 = .SubItems(18)
            A10 = .SubItems(19)
            txtInPhyVisit = .SubItems(59)
            
            ' out patient
            A11 = .SubItems(75)
            txtACC = .SubItems(34)
            txtPTreat = .SubItems(74)
            A12 = .SubItems(35)
            A13 = .SubItems(36)
            txtKidneyMonths = .SubItems(56)
            A14 = .SubItems(37)
                            
            ' OThers
            A22 = .SubItems(60)
            
            A15 = .SubItems(20)  'MR
            A16 = .SubItems(21)   ' phone
            txtLodgerDay = .SubItems(22)  'lodger
            txtLodgerAmt = .SubItems(23)
            A17 = .SubItems(53)
            ALB = .SubItems(81)
            A18 = .SubItems(24) ' admin
            A19 = .SubItems(25) 'tax
            A20 = .SubItems(26) ' other1
            
            txtOthersDesc1 = .SubItems(49)
            A20B = .SubItems(51) ' other2
            txtOthersDesc2 = .SubItems(50)
            txtDayCash = .SubItems(27)
            txtRMCash = .SubItems(28)
            A21 = .SubItems(52)                'Cash allowance
            
            ' Reimb not payable
            txtPreHos60 = .SubItems(38)
            txtPreMEdic = .SubItems(40)
            txtPreGP = .SubItems(39)
            txtPreSpec1 = .SubItems(42)
            txtPostGHosp = .SubItems(43)
            txtPreMR = .SubItems(41)
            txtpostNotSame = .SubItems(44)
            txtOthers = .SubItems(45)
            txtRBTax = .SubItems(61)
            txtICUTax = .SubItems(62)
            txtLodgerTax = .SubItems(63)
            cboINVStayType = .SubItems(64)
            txtINVPayType = .SubItems(65)
            A23 = .SubItems(68)

            
            ART = .SubItems(76)
            AIT = .SubItems(77)
            ALT = .SubItems(78)
            A1B = .SubItems(79)
            A2B = .SubItems(80)
            txtSurgWard = .SubItems(82)
            txtAnaFees = .SubItems(83)
            txtAnaCons = .SubItems(84)
            txtAnaWard = .SubItems(85)
            txtInHosFees = .SubItems(86)
            txtInHosWardFees = .SubItems(87)
            If Trim(.SubItems(88)) = "Y" Then
                chkMedRptMBM.Value = 1
            Else
                chkMedRptMBM.Value = 0
            End If
            If Trim(.SubItems(89)) = "Y" Then
                chkTelChargesMBM.Value = 1
            Else
                chkTelChargesMBM.Value = 0
            End If
            If Trim(.SubItems(90)) = "Y" Then
                chkGuaAllowMBM.Value = 1
            Else
                chkGuaAllowMBM.Value = 0
            End If
            If Trim(.SubItems(91)) = "Y" Then
                chkAdminRegMBM.Value = 1
            Else
                chkAdminRegMBM.Value = 0
            End If
            If Trim(.SubItems(92)) = "Y" Then
                chkOthersMBM.Value = 1
            Else
                chkOthersMBM.Value = 0
            End If
            
            txtGrossRB = .SubItems(93)
            A25 = .SubItems(94)
            txtOhersNoCoPay = .SubItems(95)
            txtProcFees = .SubItems(96)
            txtDIRDate.mask = ""
            txtDIRDate = .SubItems(97)
            txtDIRDate.mask = "##/##/####"
            cboDRM = .SubItems(98)
            A26 = .SubItems(99)
            txtUHCNo = .SubItems(100)
            txtInvExcess = .SubItems(101)
            lblPayAmt = .SubItems(102)
            txtOthReimNotCovered = .SubItems(103)
            
            frmHSSDetails.txtInvPharmacy = .SubItems(104)
            frmHSSDetails.txtMedSupp = .SubItems(105)
            frmHSSDetails.txtNursingCare = .SubItems(106)
            frmHSSDetails.txtNursingServ = .SubItems(107)
            frmHSSDetails.txtLab = .SubItems(108)
            frmHSSDetails.txtRadiology = .SubItems(109)
            frmHSSDetails.txtOTSupp = .SubItems(110)
            frmHSSDetails.txtEquipChrg = .SubItems(111)
            frmHSSDetails.txtMisc = .SubItems(112)
            txtInitialActAmt = IIf(Len(.SubItems(113)), .SubItems(113), 0)
            txtHOSName = IIf(Len(.SubItems(114)), .SubItems(114), "")
            A27 = IIf(Len(.SubItems(115)), .SubItems(115), 0)
            cboMaternityStatus = IIf(Len(.SubItems(116)), .SubItems(116), "")
            Call getTotal
        End With
        EnableEntries (False)
        'Frame9.Enabled = False
        CmdEdit.Enabled = True
    End If
End Sub

Private Sub cmdAdd_Click()
    EditFlag = False
    If Trim(txtCaseNo) <> "" Then
    
        'If Bordxstatus = False And lstInvoice.listitems.Count <> "0" Then
        '    MsgBox "Please generate Bordx for the claim!", vbCritical + vbOKOnly
        'Else
            checking = False
            Call clearInvoice
            CmdUpdate.Visible = True
            CmdEdit.Visible = False
            CmdAdd.Enabled = False
            CmdTotal.Enabled = False
            If Mid(txtCaseNo, 1, 2) = "IB" Then
                EnableEntriesIB (True)
                cmdHSS.Enabled = False
            Else
                EnableEntries (True)
                cmdHSS.Enabled = True
            End If
            checking = True
            cboINVPayType.SetFocus
            TxtAccess.Enabled = False
        If Trim(Mid(PLNCode, 1, 1)) = "2" Or Trim(Mid(PLNCode, 1, 1)) = "3" Or Trim(Mid(PLNCode, 1, 1)) = "4" Or Trim(Mid(PLNCode, 1, 1)) = "5" Then
            ART.Enabled = True
            AIT.Enabled = True
            ALT.Enabled = True
            txtRBTax.Enabled = False
            txtICUTax.Enabled = False
            txtLodgerTax.Enabled = False
            A19.Enabled = False
        Else
            ART.Enabled = False
            AIT.Enabled = False
            ALT.Enabled = False
            txtRBTax.Enabled = True
            txtICUTax.Enabled = True
            txtLodgerTax.Enabled = True
        End If
    
        If Trim(txtModal) = "S" Then
            FrmeSurgical.Enabled = True
            FrmeNS.Enabled = False
        ElseIf Trim(txtModal) = "N" Then
            FrmeSurgical.Enabled = False
            FrmeNS.Enabled = True
        Else
            FrmeSurgical.Enabled = True
            FrmeNS.Enabled = True
        End If
        
        If MBMBTBGStatus = True Then
            Call EnableEntriesChkBox(True)
        Else
            Call EnableEntriesChkBox(False)
        End If
        
        'End If
        
    Else
        MsgBox "Please select the case number first!", vbCritical + vbOKOnly
    End If
    
End Sub

Private Sub cmdEdit_Click()
    If lblBordx = "False" Or lblBordx = "No" Then
        EditFlag = True
        CmdUpdate.Visible = True
        CmdEdit.Visible = False
        CmdTotal.Enabled = False
        CmdAdd.Enabled = False
        CmdRemove.Enabled = False
        CmdClear.Enabled = True
        If Mid(txtCaseNo, 1, 2) = "IB" Then
            EnableEntriesIB (True)
            cmdHSS.Enabled = False
        Else
            EnableEntries (True)
            cmdHSS.Enabled = True
        End If
        TxtAccess.Enabled = False
        If Trim(Mid(PLNCode, 1, 1)) = "2" Or Trim(Mid(PLNCode, 1, 1) = "3") Or Trim(Mid(PLNCode, 1, 1) = "4") Or Trim(Mid(PLNCode, 1, 1) = "5") Then
            ART.Enabled = True
            AIT.Enabled = True
            ALT.Enabled = True
            txtRBTax.Enabled = False
            txtICUTax.Enabled = False
            txtLodgerTax.Enabled = False
            A19.Enabled = False
        Else
            ART.Enabled = False
            AIT.Enabled = False
            ALT.Enabled = False
            txtRBTax.Enabled = True
            txtICUTax.Enabled = True
            txtLodgerTax.Enabled = True
        End If
    
        If Trim(txtModal) = "S" Then
            FrmeSurgical.Enabled = True
            FrmeNS.Enabled = False
        ElseIf Trim(txtModal) = "N" Then
            FrmeSurgical.Enabled = False
            FrmeNS.Enabled = True
        Else
            FrmeSurgical.Enabled = True
            FrmeNS.Enabled = True
        End If

        If MBMBTBGStatus = True Then
            Call EnableEntriesChkBox(True)
        Else
            Call EnableEntriesChkBox(False)
        End If

    Else
        MsgBox "Invoice has been bordx!", vbCritical
    End If
End Sub

Private Sub cmdUpdate_Click()
  '  medixmasterconn.Open "File Name=" & BOSPath & "Conn\mediconnect.UDL"
  '  medixmasterconnMaint.Open "File Name=" & BOSPath & "Conn\MediConnHISMaint.UDL"
  '  medixmasterconnWS.Open "File Name=" & BOSPath & "Conn\mediconnectWS.UDL"
        Call SaveClaim
        
        
  '  medixmasterconn.Close
 '   Set medixmasterconn = Nothing
  '  medixmasterconnMaint.Close
  '  Set medixmasterconnMaint = Nothing
  '  medixmasterconnWS.Close
  '  Set medixmasterconnWS = Nothing
        
End Sub

Private Sub CmdRemove_Click()
On Error GoTo errorRemove
    
    Dim ForLoop As Integer
    Dim beginTrans As Boolean
    Dim ListCount As Integer
    Dim strsql, strActual, strNot As String
    Dim RecordSaved
    Dim RemoveInvoice  As New ADODB.Recordset
    Dim RemoveActualInvoice  As New ADODB.Recordset
    Dim RemoveNotCover  As New ADODB.Recordset
    
   ' medixmasterconnWS.Open "File Name=" & BOSPath & "Conn\mediconnectWS.UDL"
    
    If lstInvoice.SelectedItem.SubItems(2) = False Then
        RecordSaved = MsgBox("Are you sure?", vbYesNo + vbExclamation)
          If RecordSaved = vbYes Then
              ListCount = lstInvoice.ListItems.Count
              For ForLoop = 1 To ListCount
                  If lstInvoice.ListItems(ForLoop).Selected = True Then
                      lstInvoice.ListItems.Remove (ForLoop)
                      Exit For
                  End If
              Next ForLoop
              
            strsql = "delete  from ClaimInvoice where INVIndex =" & intIndex
            medixmasterconnWS.beginTrans
            beginTrans = True
            RemoveInvoice.Open strsql, medixmasterconnWS, adOpenKeyset, adLockOptimistic
            strActual = "delete from ClaimInvoiceActual where INVIndex =" & intIndex
            RemoveActualInvoice.Open strActual, medixmasterconnWS, adOpenKeyset, adLockOptimistic
            strNot = "delete from ClaimInvoiceNotCover where INVIndex =" & intIndex
            RemoveNotCover.Open strNot, medixmasterconnWS, adOpenKeyset, adLockOptimistic
            
            medixmasterconnWS.commitTrans
          End If
          checking = False
          Call clearInvoice
          checking = True
    Else
        MsgBox "Selected Invoice cannot be deleted!" & vbNewLine & "It has been used in Worksheet!", vbCritical
        
    End If
  '  medixmasterconnWS.Close
  '  Set medixmasterconnWS = Nothing
    Exit Sub
errorRemove:
    If beginTrans = True Then
        medixmasterconnWS.RollbackTrans
    End If
    MsgBox Err.Description
End Sub

Private Sub UpdateInvList()
    Dim INVList As ListItem
    If EditFlag = True Then
        Set INVList = lstInvoice.SelectedItem
        INVList.Text = intIndex
    Else
        Set INVList = lstInvoice.ListItems.Add(, , intIndex)
        INVList.Text = intIndex
    End If
    
    With INVList
        .SubItems(1) = lblBordx
        
        .SubItems(44) = cboINVPayType
        .SubItems(45) = lblLastVersion
        
        
        .SubItems(1) = IIf(txtInvDate <> "__/__/____", txtInvDate, "")
        .SubItems(2) = IIf(IsNumeric(txtINVAmount) = True, txtINVAmount, 0)
        .SubItems(3) = IIf(Trim(cboHosCode) <> "", cboHosCode, "")
        .SubItems(4) = ChkStr(Mid(Trim(CboPayType), 1, IIf(InStr(1, Trim(CboPayType), "-") > 0, InStr(1, Trim(CboPayType), "-") - 1, 2)))
        .SubItems(6) = IIf(IsNumeric(TxtAccess) = True, TxtAccess, 0)
        .SubItems(7) = IIf(IsNumeric(DB1) = True, DB1, "")
        .SubItems(8) = IIf(IsNumeric(RoomAmount) = True, RoomAmount, "")
        .SubItems(9) = IIf(IsNumeric(ICU1) = True, ICU1, "")
        .SubItems(10) = IIf(IsNumeric(ICUAmount) = True, ICUAmount, "")
        .SubItems(11) = IIf(IsNumeric(A3) = True, A3, "")
        .SubItems(12) = IIf(IsNumeric(A4) = True, A4, "")
        .SubItems(13) = IIf(IsNumeric(A8) = True, A8, "")
        .SubItems(14) = IIf(IsNumeric(A9) = True, A9, "")
        .SubItems(15) = IIf(IsNumeric(A10) = True, A10, "")
        .SubItems(16) = IIf(IsNumeric(A15) = True, A15, "")
        .SubItems(17) = IIf(IsNumeric(A16) = True, A16, "")
        .SubItems(18) = IIf(IsNumeric(txtLodgerDay) = True, txtLodgerDay, "")
        .SubItems(19) = IIf(IsNumeric(txtLodgerAmt) = True, txtLodgerAmt, "")
        .SubItems(20) = IIf(IsNumeric(A18) = True, A18, "")
        .SubItems(21) = IIf(IsNumeric(A19) = True, A19, "")
        .SubItems(22) = IIf(IsNumeric(A20) = True, A20, "")
        .SubItems(23) = IIf(IsNumeric(txtDayCash) = True, txtDayCash, "")
        .SubItems(24) = IIf(IsNumeric(txtRMCash) = True, txtRMCash, "")
        .SubItems(25) = IIf(IsNumeric(txtXray) = True, txtXray, "")
        .SubItems(26) = IIf(IsNumeric(txtPreLab) = True, txtPreLab, "")
        .SubItems(27) = IIf(IsNumeric(txtPreOthers) = True, txtPreOthers, "")
        .SubItems(28) = IIf(IsNumeric(A6) = True, A6, "")
        .SubItems(29) = IIf(IsNumeric(A7) = True, A7, "")
        .SubItems(30) = IIf(IsNumeric(A11) = True, A11, "")
        .SubItems(31) = IIf(IsNumeric(A12) = True, A12, "")
        .SubItems(32) = IIf(IsNumeric(A13) = True, A13, "")
        .SubItems(33) = IIf(IsNumeric(A14) = True, A14, "")
        .SubItems(34) = IIf(IsNumeric(txtPreHos60) = True, txtPreHos60, "")
        .SubItems(35) = IIf(IsNumeric(txtPreMEdic) = True, txtPreMEdic, "")
        .SubItems(36) = IIf(IsNumeric(txtPreGP) = True, txtPreGP, "")
        .SubItems(37) = IIf(IsNumeric(txtPreSpec1) = True, txtPreSpec1, "")
        .SubItems(38) = IIf(IsNumeric(txtPostGHosp) = True, txtPostGHosp, "")
        .SubItems(39) = IIf(IsNumeric(txtPreMR) = True, txtPreMR, "")
        .SubItems(40) = IIf(IsNumeric(txtpostNotSame) = True, txtpostNotSame, "")
        .SubItems(41) = IIf(IsNumeric(txtOthers) = True, txtOthers, "")
        
    End With
    EditFlag = False
End Sub

Private Sub ICU1_lostfocus()
    If IsNumeric(ICU1) = True And IsNumeric(ICUAmount) = True Then
        A2 = CDbl(ICU1) * CDbl(ICUAmount)
    Else
        A2 = 0
    End If
End Sub

Private Sub ICU1_KeyPress(KeyAscii As Integer)
    KeyAscii = KeyNumber(KeyAscii)
End Sub

Private Sub ICUAmount_KeyPress(KeyAscii As Integer)
    KeyAscii = KeyNumber(KeyAscii)
End Sub

Private Sub txtINVDate2_GotFocus()
    txtINVDate2.SelStart = 1
End Sub

Private Sub txtInvExcess_Change()
Dim TempInvAmt As Double
Dim TempInvExcess As Double
Dim TempInvPayAmt As Double
    
    If Mid(cboINVPayType, 1, 1) = "H" Then
    
        lblPayAmt = 0
        
        TempInvAmt = IIf(IsNumeric(txtINVAmount), txtINVAmount, 0)
        
        TempInvExcess = IIf(IsNumeric(txtInvExcess), txtInvExcess, 0)
    
        TempInvPayAmt = IIf(IsNumeric(lblPayAmt), lblPayAmt, 0)
        
        lblPayAmt = CDbl(TempInvAmt) - CDbl(TempInvExcess)
    End If
End Sub

Private Sub txtInvExcess_LostFocus()
Dim TempInvAmt As Double
Dim TempInvExcess As Double
Dim TempInvPayAmt As Double
    
    lblPayAmt = 0
    
    TempInvAmt = IIf(IsNumeric(txtINVAmount), txtINVAmount, 0)
    
    TempInvExcess = IIf(IsNumeric(txtInvExcess), txtInvExcess, 0)

    TempInvPayAmt = IIf(IsNumeric(lblPayAmt), lblPayAmt, 0)
    
    lblPayAmt = CDbl(TempInvAmt) - CDbl(TempInvExcess)
End Sub



Private Sub txtKidneyMonths_KeyPress(KeyAscii As Integer)
    KeyAscii = KeyNumber(KeyAscii)
End Sub

Private Sub txtLodgerAmt_KeyPress(KeyAscii As Integer)
    KeyAscii = KeyNumber(KeyAscii)
End Sub

Private Sub txtLodgerDay_KeyPress(KeyAscii As Integer)
    KeyAscii = KeyNumber(KeyAscii)
End Sub

Private Sub txtLodgerTax_LostFocus()
Dim a, B, C As Double
    
    If IsNumeric(txtRBTax) = True Then
        a = txtRBTax
    Else
        a = 0
    End If
    If IsNumeric(txtICUTax) = True Then
        B = txtICUTax
    Else
        B = 0
    End If
    If IsNumeric(txtLodgerTax) = True Then
        C = txtLodgerTax
    Else
        C = 0
    End If
    A19 = CDbl(a) + CDbl(B) + CDbl(C)
End Sub

Private Sub TxtMLab_Change()
Dim a, B, C As Double
    
    If IsNumeric(TxtMXRay) = True Then
        a = TxtMXRay
    Else
        a = 0
    End If
    If IsNumeric(TxtMLab) = True Then
        B = TxtMLab
    Else
        B = 0
    End If
    If IsNumeric(TxtMPreOthers) = True Then
        C = TxtMPreOthers
    Else
        C = 0
    End If
    A5B = CDbl(a) + CDbl(B) + CDbl(C)
End Sub

Private Sub TxtMLab_LostFocus()
Dim a, B, C As Double
    
    If IsNumeric(TxtMXRay) = True Then
        a = TxtMXRay
    Else
        a = 0
    End If
    If IsNumeric(TxtMLab) = True Then
        B = TxtMLab
    Else
        B = 0
    End If
    If IsNumeric(TxtMPreOthers) = True Then
        C = TxtMPreOthers
    Else
        C = 0
    End If
    A5B = CDbl(a) + CDbl(B) + CDbl(C)
End Sub

Private Sub TxtMPreOthers_Change()
Dim a, B, C As Double
    
    If IsNumeric(TxtMXRay) = True Then
        a = TxtMXRay
    Else
        a = 0
    End If
    If IsNumeric(TxtMLab) = True Then
        B = TxtMLab
    Else
        B = 0
    End If
    If IsNumeric(TxtMPreOthers) = True Then
        C = TxtMPreOthers
    Else
        C = 0
    End If
    A5B = CDbl(a) + CDbl(B) + CDbl(C)

End Sub

Private Sub TxtMPreOthers_LostFocus()
Dim a, B, C As Double
    
    If IsNumeric(TxtMXRay) = True Then
        a = TxtMXRay
    Else
        a = 0
    End If
    If IsNumeric(TxtMLab) = True Then
        B = TxtMLab
    Else
        B = 0
    End If
    If IsNumeric(TxtMPreOthers) = True Then
        C = TxtMPreOthers
    Else
        C = 0
    End If
    A5B = CDbl(a) + CDbl(B) + CDbl(C)
End Sub

Private Sub TxtMXRay_Change()
    Dim a, B, C As Double
    
    If IsNumeric(TxtMXRay) = True Then
        a = TxtMXRay
    Else
        a = 0
    End If
    If IsNumeric(TxtMLab) = True Then
        B = TxtMLab
    Else
        B = 0
    End If
    If IsNumeric(TxtMPreOthers) = True Then
        C = TxtMPreOthers
    Else
        C = 0
    End If
    A5B = CDbl(a) + CDbl(B) + CDbl(C)

End Sub

Private Sub TxtMXRay_KeyPress(KeyAscii As Integer)
    KeyAscii = KeyNumber(KeyAscii)
End Sub

Private Sub TxtMLab_KeyPress(KeyAscii As Integer)
    KeyAscii = KeyNumber(KeyAscii)
End Sub

Private Sub TxtMPreOthers_KeyPress(KeyAscii As Integer)
    KeyAscii = KeyNumber(KeyAscii)
End Sub

Private Sub A5B_KeyPress(KeyAscii As Integer)
    KeyAscii = KeyNumber(KeyAscii)
End Sub

Private Sub txtACC_LostFocus()
    Dim a, B As Double
    
    If IsNumeric(txtACC) = True Then
        a = txtACC
    Else
        a = 0
    End If
    If IsNumeric(txtPTreat) = True Then
        B = txtPTreat
    Else
        B = 0
    End If
    
    A11 = CDbl(a) + CDbl(B)
End Sub


Private Sub txtPhysician_Change()
    Dim a, B, C, d As Double
    If TotalClick = False Then
        If IsNumeric(txtSurgicalFee) = True Then
            a = txtSurgicalFee
        Else
            a = 0
        End If
        If IsNumeric(txtPhysician) = True Then
            B = txtPhysician
        Else
            B = 0
        End If
        If IsNumeric(txtPostTreat) = True Then
            C = txtPostTreat
        Else
            C = 0
        End If
        
        If IsNumeric(txtSurgWard) = True Then
            d = txtSurgWard
        Else
            d = 0
        End If
    
       
        A6 = CDbl(a) + CDbl(B) + CDbl(C) + CDbl(d)
    End If
End Sub

Private Sub txtPostTreat_Change()
    Dim a, B, C, d As Double
    If TotalClick = False Then
        If IsNumeric(txtSurgicalFee) = True Then
            a = txtSurgicalFee
        Else
            a = 0
        End If
        If IsNumeric(txtPhysician) = True Then
            B = txtPhysician
        Else
            B = 0
        End If
        If IsNumeric(txtPostTreat) = True Then
            C = txtPostTreat
        Else
            C = 0
        End If
            If IsNumeric(txtSurgWard) = True Then
             d = txtSurgWard
         Else
             d = 0
         End If
        
         A6 = CDbl(a) + CDbl(B) + CDbl(C) + CDbl(d)
    End If
End Sub

Private Sub txtProcFees_Change()
Dim a, B, C As Double
    
    If IsNumeric(txtInHosFees) = True Then
        a = txtInHosFees
    Else
        a = 0
    End If
    If IsNumeric(txtInHosWardFees) = True Then
        B = txtInHosWardFees
    Else
        B = 0
    End If
    
    If IsNumeric(txtProcFees) = True Then
        C = txtProcFees
    Else
        C = 0
    End If
    
    A9 = CDbl(a) + CDbl(B) + CDbl(C)

End Sub

Private Sub txtProcFees_KeyPress(KeyAscii As Integer)
    KeyAscii = KeyNumber(KeyAscii)
End Sub

Private Sub txtProcFees_LostFocus()
Dim a, B, C As Double
    
    If IsNumeric(txtInHosFees) = True Then
        a = txtInHosFees
    Else
        a = 0
    End If
    If IsNumeric(txtInHosWardFees) = True Then
        B = txtInHosWardFees
    Else
        B = 0
    End If
    
    If IsNumeric(txtProcFees) = True Then
        C = txtProcFees
    Else
        C = 0
    End If
    
    A9 = CDbl(a) + CDbl(B) + CDbl(C)

End Sub

Private Sub txtPTreat_LostFocus()
    Dim a, B As Double
    
    If IsNumeric(txtACC) = True Then
        a = txtACC
    Else
        a = 0
    End If
    If IsNumeric(txtPTreat) = True Then
        B = txtPTreat
    Else
        B = 0
    End If
    
    A11 = CDbl(a) + CDbl(B)
End Sub

Private Sub TxtMXRay_LostFocus()
    Dim a, B, C As Double
    
    If IsNumeric(TxtMXRay) = True Then
        a = TxtMXRay
    Else
        a = 0
    End If
    If IsNumeric(TxtMLab) = True Then
        B = TxtMLab
    Else
        B = 0
    End If
    If IsNumeric(TxtMPreOthers) = True Then
        C = TxtMPreOthers
    Else
        C = 0
    End If
    A5B = CDbl(a) + CDbl(B) + CDbl(C)
End Sub

Private Sub txtOthers_KeyPress(KeyAscii As Integer)
    KeyAscii = KeyNumber(KeyAscii)
End Sub

Private Sub txtPhysician_KeyPress(KeyAscii As Integer)
    KeyAscii = KeyNumber(KeyAscii)
End Sub

Private Sub txtPostGHosp_KeyPress(KeyAscii As Integer)
    KeyAscii = KeyNumber(KeyAscii)
End Sub

Private Sub txtpostNotSame_KeyPress(KeyAscii As Integer)
    KeyAscii = KeyNumber(KeyAscii)
End Sub

Private Sub txtPostTreat_KeyPress(KeyAscii As Integer)
    KeyAscii = KeyNumber(KeyAscii)
End Sub

Private Sub txtPreGP_KeyPress(KeyAscii As Integer)
    KeyAscii = KeyNumber(KeyAscii)
End Sub

Private Sub txtPreHos60_KeyPress(KeyAscii As Integer)
    KeyAscii = KeyNumber(KeyAscii)
End Sub

Private Sub A24_KeyPress(KeyAscii As Integer)
KeyAscii = KeyNumber(KeyAscii)
End Sub

Private Sub txtPreLab_KeyPress(KeyAscii As Integer)
    KeyAscii = KeyNumber(KeyAscii)
End Sub

Private Sub txtPreMEdic_KeyPress(KeyAscii As Integer)
    KeyAscii = KeyNumber(KeyAscii)
End Sub

Private Sub txtPreMR_KeyPress(KeyAscii As Integer)
    KeyAscii = KeyNumber(KeyAscii)
End Sub

Private Sub txtPreOthers_KeyPress(KeyAscii As Integer)
    KeyAscii = KeyNumber(KeyAscii)
End Sub

Private Sub txtPreSpec1_KeyPress(KeyAscii As Integer)
    KeyAscii = KeyNumber(KeyAscii)
End Sub

Private Sub A5_LostFocus()
    Dim a, B, C, d As Double
    
    If IsNumeric(A5) = True Then
        a = A5
    Else
        a = 0
    End If
    If IsNumeric(A24) = True Then
        B = A24
    Else
        B = 0
    End If
    If IsNumeric(A6) = True Then
        C = A6
    Else
        C = 0
    End If
    If IsNumeric(A7) = True Then
        d = A7
    Else
        d = 0
    End If
    Text1 = CDbl(a) + CDbl(B) + CDbl(C) + CDbl(d)
    
    Call getTotal
End Sub

Private Sub A24_LostFocus()
    Dim a, B, C, d As Double
    
    If IsNumeric(A5) = True Then
        a = A5
    Else
        a = 0
    End If
    If IsNumeric(A24) = True Then
        B = A24
    Else
        B = 0
    End If
    If IsNumeric(A6) = True Then
        C = A6
    Else
        C = 0
    End If
    If IsNumeric(A7) = True Then
        d = A7
    Else
        d = 0
    End If
    Text1 = CDbl(a) + CDbl(B) + CDbl(C) + CDbl(d)
    Call getTotal
End Sub

Private Sub A6_LostFocus()
    Dim a, B, C, d As Double
    Text1 = 0
    If IsNumeric(A5) = True Then
        a = A5
    Else
        a = 0
    End If
    If IsNumeric(A24) = True Then
        B = A24
    Else
        B = 0
    End If
    If IsNumeric(A6) = True Then
        C = A6
    Else
        C = 0
    End If
    If IsNumeric(A7) = True Then
        d = A7
    Else
        d = 0
    End If
    Text1 = CDbl(a) + CDbl(B) + CDbl(C) + CDbl(d)
    
    Call getTotal
End Sub

Private Sub A7_LostFocus()
    Dim a, B, C, d As Double
    
    If IsNumeric(A5) = True Then
        a = A5
    Else
        a = 0
    End If
    If IsNumeric(A24) = True Then
        B = A24
    Else
        B = 0
    End If
    If IsNumeric(A6) = True Then
        C = A6
    Else
        C = 0
    End If
    If IsNumeric(A7) = True Then
        d = A7
    Else
        d = 0
    End If
    Text1 = CDbl(a) + CDbl(B) + CDbl(C) + CDbl(d)
    
    Call getTotal
End Sub

Private Sub A10_LostFocus()
    Dim a, B, C, d As Double
    
    If IsNumeric(A5B) = True Then
        a = A5B
    Else
        a = 0
    End If
    If IsNumeric(A8) = True Then
        B = A8
    Else
        B = 0
    End If
    If IsNumeric(A9) = True Then
        C = A9
    Else
        C = 0
    End If
    If IsNumeric(A10) = True Then
        d = A10
    Else
        d = 0
    End If
    Text2 = CDbl(a) + CDbl(B) + CDbl(C) + CDbl(d)
    
    Call getTotal
End Sub

Private Sub A9_LostFocus()
    Dim a, B, C, d As Double
    
    If IsNumeric(A5B) = True Then
        a = A5B
    Else
        a = 0
    End If
    If IsNumeric(A8) = True Then
        B = A8
    Else
        B = 0
    End If
    If IsNumeric(A9) = True Then
        C = A9
    Else
        C = 0
    End If
    If IsNumeric(A10) = True Then
        d = A10
    Else
        d = 0
    End If
    Text2 = CDbl(a) + CDbl(B) + CDbl(C) + CDbl(d)
    
    Call getTotal
End Sub

Private Sub A8_LostFocus()
    Dim a, B, C, d As Double
    
    If IsNumeric(A5B) = True Then
        a = A5B
    Else
        a = 0
    End If
    If IsNumeric(A8) = True Then
        B = A8
    Else
        B = 0
    End If
    If IsNumeric(A9) = True Then
        C = A9
    Else
        C = 0
    End If
    If IsNumeric(A10) = True Then
        d = A10
    Else
        d = 0
    End If
    Text2 = CDbl(a) + CDbl(B) + CDbl(C) + CDbl(d)
    
    Call getTotal
End Sub

Private Sub A5B_LostFocus()
    Dim a, B, C, d As Double
    
    If IsNumeric(A5B) = True Then
        a = A5B
    Else
        a = 0
    End If
    If IsNumeric(A8) = True Then
        B = A8
    Else
        B = 0
    End If
    If IsNumeric(A9) = True Then
        C = A9
    Else
        C = 0
    End If
    If IsNumeric(A10) = True Then
        d = A10
    Else
        d = 0
    End If
    Text2 = CDbl(a) + CDbl(B) + CDbl(C) + CDbl(d)
    
    Call getTotal
End Sub

Private Sub txtRBTax_LostFocus()
Dim a, B, C As Double
    
    If IsNumeric(txtRBTax) = True Then
        a = txtRBTax
    Else
        a = 0
    End If
    If IsNumeric(txtICUTax) = True Then
        B = txtICUTax
    Else
        B = 0
    End If
    If IsNumeric(txtLodgerTax) = True Then
        C = txtLodgerTax
    Else
        C = 0
    End If
    A19 = CDbl(a) + CDbl(B) + CDbl(C)
    
End Sub

Private Sub txtRMCash_KeyPress(KeyAscii As Integer)
    KeyAscii = KeyNumber(KeyAscii)
End Sub

Private Sub txtSurgicalFee_Change()
    Dim a, B, C, d As Double
    If TotalClick = False Then
        If IsNumeric(txtSurgicalFee) = True Then
            a = txtSurgicalFee
        Else
            a = 0
        End If
        If IsNumeric(txtPhysician) = True Then
            B = txtPhysician
        Else
            B = 0
        End If
        If IsNumeric(txtPostTreat) = True Then
            C = txtPostTreat
        Else
            C = 0
        End If
       
        If IsNumeric(txtSurgWard) = True Then
            d = txtSurgWard
        Else
            d = 0
        End If
    
        A6 = CDbl(a) + CDbl(B) + CDbl(C) + CDbl(d)
    End If
End Sub

Private Sub txtSurgicalFee_KeyPress(KeyAscii As Integer)
    KeyAscii = KeyNumber(KeyAscii)
End Sub

Private Sub txtSurgWard_Change()
Dim a, B, C, d As Double
    If TotalClick = False Then
        If IsNumeric(txtSurgicalFee) = True Then
            a = txtSurgicalFee
        Else
            a = 0
        End If
        If IsNumeric(txtPhysician) = True Then
            B = txtPhysician
        Else
            B = 0
        End If
        If IsNumeric(txtPostTreat) = True Then
            C = txtPostTreat
        Else
            C = 0
        End If
        
        If IsNumeric(txtSurgWard) = True Then
            d = txtSurgWard
        Else
            d = 0
        End If
        
        A6 = CDbl(a) + CDbl(B) + CDbl(C) + CDbl(d)
    End If
End Sub

Private Sub txtSurgWard_KeyPress(KeyAscii As Integer)
    KeyAscii = KeyNumber(KeyAscii)
End Sub

Private Sub txtSurgWard_LostFocus()
Dim a, B, C, d As Double
    
    If IsNumeric(txtSurgicalFee) = True Then
        a = txtSurgicalFee
    Else
        a = 0
    End If
    If IsNumeric(txtPhysician) = True Then
        B = txtPhysician
    Else
        B = 0
    End If
    If IsNumeric(txtPostTreat) = True Then
        C = txtPostTreat
    Else
        C = 0
    End If
   
    If IsNumeric(txtSurgWard) = True Then
        d = txtSurgWard
    Else
        d = 0
    End If
   
    A6 = CDbl(a) + CDbl(B) + CDbl(C) + CDbl(d)
End Sub

Private Sub txtXray_KeyPress(KeyAscii As Integer)
    KeyAscii = KeyNumber(KeyAscii)
End Sub


Private Sub txtXray_lostfocus()
    Dim a, B, C As Double
    
    If IsNumeric(txtXray) = True Then
        a = txtXray
    Else
        a = 0
    End If
    If IsNumeric(txtPreLab) = True Then
        B = txtPreLab
    Else
        B = 0
    End If
    If IsNumeric(txtPreOthers) = True Then
        C = txtPreOthers
    Else
        C = 0
    End If
    A5 = CDbl(a) + CDbl(B) + CDbl(C)
End Sub

Private Sub txtPreLab_lostfocus()
    Dim a, B, C As Double
    
    If IsNumeric(txtXray) = True Then
        a = txtXray
    Else
        a = 0
    End If
    If IsNumeric(txtPreLab) = True Then
        B = txtPreLab
    Else
        B = 0
    End If
    If IsNumeric(txtPreOthers) = True Then
        C = txtPreOthers
    Else
        C = 0
    End If
    A5 = CDbl(a) + CDbl(B) + CDbl(C)
End Sub

Private Sub txtPreOthers_lostfocus()
    Dim a, B, C As Double
    
    If IsNumeric(txtXray) = True Then
        a = txtXray
    Else
        a = 0
    End If
    If IsNumeric(txtPreLab) = True Then
        B = txtPreLab
    Else
        B = 0
    End If
    If IsNumeric(txtPreOthers) = True Then
        C = txtPreOthers
    Else
        C = 0
    End If
    A5 = CDbl(a) + CDbl(B) + CDbl(C)
End Sub


Private Sub txtSurgicalFee_lostfocus()
    Dim a, B, C, d As Double
    
    If IsNumeric(txtSurgicalFee) = True Then
        a = txtSurgicalFee
    Else
        a = 0
    End If
    If IsNumeric(txtPhysician) = True Then
        B = txtPhysician
    Else
        B = 0
    End If
    If IsNumeric(txtPostTreat) = True Then
        C = txtPostTreat
    Else
        C = 0
    End If
   
    If IsNumeric(txtSurgWard) = True Then
        d = txtSurgWard
    Else
        d = 0
    End If

    A6 = CDbl(a) + CDbl(B) + CDbl(C) + CDbl(d)

End Sub

Private Sub txtPhysician_lostfocus()
    Dim a, B, C, d As Double
    
    If IsNumeric(txtSurgicalFee) = True Then
        a = txtSurgicalFee
    Else
        a = 0
    End If
    If IsNumeric(txtPhysician) = True Then
        B = txtPhysician
    Else
        B = 0
    End If
    If IsNumeric(txtPostTreat) = True Then
        C = txtPostTreat
    Else
        C = 0
    End If
    
    If IsNumeric(txtSurgWard) = True Then
        d = txtSurgWard
    Else
        d = 0
    End If

   
    A6 = CDbl(a) + CDbl(B) + CDbl(C) + CDbl(d)
End Sub

Private Sub txtPostTreat_lostfocus()
    Dim a, B, C, d As Double
    
    If IsNumeric(txtSurgicalFee) = True Then
        a = txtSurgicalFee
    Else
        a = 0
    End If
    If IsNumeric(txtPhysician) = True Then
        B = txtPhysician
    Else
        B = 0
    End If
    If IsNumeric(txtPostTreat) = True Then
        C = txtPostTreat
    Else
        C = 0
    End If
   
    If IsNumeric(txtSurgWard) = True Then
        d = txtSurgWard
    Else
        d = 0
    End If
   
    A6 = CDbl(a) + CDbl(B) + CDbl(C) + CDbl(d)
End Sub

Private Sub txtLodgerDay_lostfocus()
    If IsNumeric(txtLodgerDay) = True And _
     IsNumeric(txtLodgerAmt) = True Then
        A17 = CDbl(txtLodgerDay) * CDbl(txtLodgerAmt)
    Else
        A17 = 0
    End If
End Sub

Private Sub txtLodgerAmt_lostfocus()
    If IsNumeric(txtLodgerDay) = True And _
     IsNumeric(txtLodgerAmt) = True Then
        A17 = CDbl(txtLodgerDay) * CDbl(txtLodgerAmt)
    Else
        A17 = 0
    End If
End Sub

Private Sub txtRMCash_LostFocus()
    If IsNumeric(txtRMCash) = True And _
    IsNumeric(txtDayCash) = True Then
        A21 = CDbl(txtRMCash) * CDbl(txtDayCash)
    Else
        A21 = 0
    End If
End Sub

Private Sub txtDayCash_lostfocus()
    If IsNumeric(txtRMCash) = True And _
    IsNumeric(txtDayCash) = True Then
        A21 = CDbl(txtRMCash) * CDbl(txtDayCash)
    Else
        A21 = 0
    End If
End Sub

' ***************** SAve Records
Private Sub SaveClaim()
'On Error GoTo ErrorSave
Dim startTrans As Boolean
Dim Invoice As New ADODB.Recordset
Dim strsql As String
Dim RecordSaved
Dim strsqlHOS  As String
Dim HOS As New ADODB.Recordset
Dim LGNo As String
Dim LGVer As String

    If Trim(txtCaseNo) = "" Then
        MsgBox "Please select a record before proceeding! ", vbCritical
        txtCaseNo.SetFocus
    ElseIf Trim(txtMBMName) = "" Then
        MsgBox "Please select a record before proceeding! ", vbCritical
        txtCaseNo.SetFocus
    ElseIf txtDisChargeDate = "__/__/____" Or txtDisChargeDate = "" Then
        MsgBox "Please Enter Date of Discharge! ", vbCritical
    ElseIf Trim(cboINVPayType) = "" Then
        MsgBox "Please select who to be paid (Member / Hospital )! ", vbCritical
        cboINVPayType.SetFocus
    ElseIf Trim(cboINVStayType) = "" Then
        MsgBox "Please select a Stay Type! ", vbCritical
        cboINVStayType.SetFocus
    ElseIf Trim(cboINVPayType) = "" Then
        MsgBox "Please select a Pay Type! ", vbCritical
        cboINVPayType.SetFocus
    
    ElseIf Trim(cboHosCode) = "" Then
        MsgBox "Please key in the person/hospital to be paid! ", vbCritical
        cboHosCode.SetFocus
    ElseIf Trim(txtInvNo) = "" Then
        MsgBox "Please key in Invoice No.! ", vbCritical
        txtInvNo.SetFocus
    ElseIf txtInvDate = "__/__/____" Or txtINVDate2 = "__/__/____" Then
        MsgBox "Please key in Invoice date! ", vbCritical
    ElseIf IsDate(txtInvDate) = False Then
        MsgBox "Please key in correct Invoice date! ", vbCritical
        txtInvDate.SetFocus
    ElseIf txtDIRDate = "__/__/____" And Mid(cboINVPayType, 1, 1) = "M" Then
        MsgBox "Please key in Invoice Receipt date! ", vbCritical
        txtDIRDate.SetFocus
    ElseIf cboDRM = "" And Mid(cboINVPayType, 1, 1) = "M" Then
        MsgBox "Please key in Invoice Receipt Mode! ", vbCritical
        cboDRM.SetFocus
    ElseIf txtInvExcess = "" And Mid(Trim(cboINVPayType), 1, 1) = "H" Then
        MsgBox "Please key in Excess Amount! ", vbCritical
        txtInvExcess.SetFocus
    ElseIf txtInitialActAmt = "" Then
        MsgBox "Please key in Initial Invoice Amount! ", vbCritical
        txtInitialActAmt.SetFocus
    ElseIf (cboHosCode = "DENTAL CLINIC" Or cboHosCode = "DIALYSIS CENTRE" Or cboHosCode = "FOREIGN HOSPITAL" Or cboHosCode = "GENERAL PRACTITIONER" Or cboHosCode = "GOVERNMENT HOSPITAL" Or cboHosCode = "NON-PANEL HOSPITAL" Or cboHosCode = "SPECIALIST CLINIC") And Trim(txtHOSName) = "" Then
        MsgBox "Please Enter Hospital Name or Clinic Name!", vbOKOnly + vbInformation
        txtHOSName.Enabled = True
        txtHOSName.SetFocus
    ElseIf Trim(Mid(cboINVPayType, 1, 1)) = "H" And Len(TmpCASHOSName) <> Len(cboHosCode) Then
       RecordSave = MsgBox("Different Hospital. Do you want to Proceed? ", vbYesNo + vbCritical)
        If RecordSave = vbNo Then
            cboHosCode.SetFocus
            Exit Sub
        Else
            GoTo SaveFunction
        End If
        
    Else
    
        Screen.MousePointer = vbHourglass
SaveFunction:
        strsql = "SELECt Top 100 INVIndex, InvNo " & _
            " FROM ClaimInvoice where CASNUmber ='" & RemStr(txtCaseNo) & "' and Invno='" & Trim(txtInvNo) & "'"
        Invoice.Open strsql, medixmasterconnWS, adOpenKeyset, adLockOptimistic
        If Invoice.recordCount And EditFlag = False Then
            MsgBox "Current Invoice No has been used before! PLease re-verify the invoice.", vbExclamation
            txtInvNo.SetFocus
        Else

            medixmasterconnWS.beginTrans
            startTrans = True
            Call getTotal
            Call SaveInvoice
            Call saveActual
            Call saveActualOthers
            Call saveNotCover
            Call SaveINVRemarks
            If Trim(lblWs) = True Then
                Call saveLimit
                Call SaveClaimDate
            End If
            Call SaveCASOthers
            ' commit changes
            medixmasterconnWS.commitTrans
            startTrans = False
            
            If PAYCode = "AX" Then
                strsqlHOS = "select HOSDiscInd,HOSName from HOS where HOSName = '" & Trim(cboHosCode) & "'"
                HOS.Open strsqlHOS, medixmasterconn, adOpenKeyset, adLockOptimistic
                    If HOS.recordCount Then
                        If Trim(HOS!HOSDiscInd) = "Y" Then
                            MsgBox "Please ensure hospital discount is given!", vbOKOnly + vbInformation
                        End If
                    End If
                HOS.Close
                Set HOS = Nothing
            End If
            
            CmdUpdate.Visible = False
            CmdEdit.Visible = True
            CmdTotal.Enabled = True
            CmdAdd.Enabled = True
            CmdRemove.Enabled = True

            CmdClear.Enabled = True
            Call ShowInvoice
            checking = False
            Call clearInvoice
            Call ClearHSS
            checking = True
            EnableEntries (False)
            CmdTotal.Enabled = True
            EditFlag = False
            
            'LGNo = Trim(Mid(lstLGListing.SelectedItem.SubItems(1), 1, IIf(InStr(1, lstLGListing.SelectedItem.SubItems(1), "-") = 0, 10, InStr(1, lstLGListing.SelectedItem.SubItems(1), "-") - 1)))
            'LGVer = Trim(Mid(lstLGListing.SelectedItem.SubItems(1), InStr(1, lstLGListing.SelectedItem.SubItems(1), "-") + 1, Len(lstLGListing.SelectedItem.SubItems(1)) - InStr(1, lstLGListing.SelectedItem.SubItems(1), "-")))

            'strsql = " update CASLG set " & _
            '         " LGIndicator = 'P' " & _
            '         " Where CASNumber = '" & Trim(LGNo) & "' and LGNo '" & Trim(LGVer) & "'"
            'medixmasterconn.Execute strsql


        End If
        Screen.MousePointer = vbDefault
    End If
    Exit Sub
'ErrorSave:
  '  Screen.MousePointer = vbDefault
  '  If startTrans = True Then
    '    medixmasterconnWS.RollbackTrans
    'End If
   ' MsgBox Err.Description
End Sub

Private Sub SaveClaimDate()
Dim tmpSql As String
    
    tmpSql = "Update Claim set CLMLastDocDate  = '" & Format(txtDIRDate, "yyyy/mm/dd") & "' Where CASNumber = '" & Trim(txtCaseNo) & "' And ClaimVersion = '" & Trim(lblVersion) & "'"
    medixmasterconnWS.Execute tmpSql

End Sub

Private Sub SaveInvoice()
    Dim ClaimInvoice  As New ADODB.Recordset
    Dim HOS As New ADODB.Recordset
    
    Dim strsql As String
    Dim strsqlHOS As String
    
    Dim INVPayee As String
    Dim InvDate As String
    Dim EnteredDate As String
    Dim FollowupDate As String
    Dim DIRDate As String
    Dim INVAmount As Double
    Dim MBMPaidAmount As Double
    Dim ExcAmt As Double
    Dim PayAmt As Double
    
    InvDate = txtInvDate
    
    EnteredDate = GetServerTime
    
    If IsDate(txtFolDate) Then
        FollowupDate = txtFolDate
    End If
    
    If IsDate(txtDIRDate) Then
        DIRDate = txtDIRDate
    End If
    
    If IsNumeric(txtINVAmount) = True Then
        INVAmount = CDbl(txtINVAmount)
    Else
        INVAmount = 0
    End If
    
    If IsNumeric(txtInvExcess) = True Then
        MBMPaidAmount = CDbl(txtInvExcess)
    Else
        MBMPaidAmount = 0
    End If
    
    
    If Len(txtUHCNo) Then
        CLMUHCNo = txtUHCNo
    Else
        CLMUHCNo = ""
    End If
    
    HOspitals = ""
    
     
    ExcAmt = IIf(IsNumeric(txtInvExcess), txtInvExcess, 0)
    
    PayAmt = IIf(IsNumeric(lblPayAmt), lblPayAmt, 0)
    
    
        If ChkStr(Mid(cboINVPayType, 1, IIf(InStr(1, cboINVPayType, "-") > 0, InStr(1, cboINVPayType, "-") - 1, 2))) = "H" Then
            HOspitals = Replace(cboHosCode, "'", "''")
            strsqlHOS = "select Top 10 HosCode, HosName from HOS where HosName = '" & Trim(HOspitals) & "'"
            HOS.Open strsqlHOS, medixmasterconn, adOpenKeyset, adLockOptimistic
            
            If HOS.recordCount Then
                INVPayee = HOS!HOSCode
            End If
            HOS.Close
            Set HOS = Nothing
        Else
            INVPayee = ChkStr(cboHosCode)
        End If
    
    If EditFlag = False Then
        
        If IsDate(FollowupDate) And IsDate(DIRDate) Then
            strsql = "EXEC Insert_ClaimInvoice_2013 '" & ChkStr(lblLastVersion) & "', " & 0 & " , " & 0 & ", " & _
                     "'" & ChkStr(txtInvNo) & "', '" & ChkStr(txtCaseNo) & "' , '" & Format(InvDate, "YYYY/MM/DD") & "', " & _
                     "'" & Format(EnteredDate, "YYYY/MM/DD") & "', '" & Logon & "' , '" & INVPayee & "', " & _
                     "'" & Mid(cboINVStayType, 1, 1) & "', '" & Mid(txtINVPayType, 1, 1) & "' , '" & ChkStr(Mid(cboINVPayType, 1, IIf(InStr(1, cboINVPayType, "-") > 0, InStr(1, cboINVPayType, "-") - 1, 2))) & "', " & _
                     "'" & Mid(cboDRM, 1, 2) & "', '" & Format(FollowupDate, "YYYY/MM/DD") & "','" & Format(DIRDate, "YYYY/MM/DD") & "', " & _
                     "" & INVAmount & ", " & MBMPaidAmount & ", '" & CLMUHCNo & "', " & ExcAmt & ", " & PayAmt & ", " & txtInitialActAmt & ", '" & Trim(cboRoomType) & "'"
        
        ElseIf IsDate(FollowupDate) And IsDate(DIRDate) = False Then
            strsql = "EXEC Insert_ClaimInvoice_WODIR '" & ChkStr(lblLastVersion) & "', " & 0 & " , " & 0 & ", " & _
                     "'" & ChkStr(txtInvNo) & "', '" & ChkStr(txtCaseNo) & "' , '" & Format(InvDate, "YYYY/MM/DD") & "', " & _
                     "'" & Format(EnteredDate, "YYYY/MM/DD") & "', '" & Logon & "' , '" & INVPayee & "', " & _
                     "'" & Mid(cboINVStayType, 1, 1) & "', '" & Mid(txtINVPayType, 1, 1) & "' , '" & ChkStr(Mid(cboINVPayType, 1, IIf(InStr(1, cboINVPayType, "-") > 0, InStr(1, cboINVPayType, "-") - 1, 2))) & "', " & _
                     "'" & Mid(cboDRM, 1, 2) & "', '" & Format(FollowupDate, "YYYY/MM/DD") & "', " & _
                     "" & INVAmount & ", " & MBMPaidAmount & ", '" & CLMUHCNo & "' , " & ExcAmt & ", " & PayAmt & " , " & txtInitialActAmt & ", '" & Trim(cboRoomType) & "'"
        
        ElseIf IsDate(FollowupDate) = False And IsDate(DIRDate) Then
            strsql = "EXEC Insert_ClaimInvoice_WOFolDate '" & ChkStr(lblLastVersion) & "', " & 0 & " , " & 0 & ", " & _
                     "'" & ChkStr(txtInvNo) & "', '" & ChkStr(txtCaseNo) & "' , '" & Format(InvDate, "YYYY/MM/DD") & "', " & _
                     "'" & Format(EnteredDate, "YYYY/MM/DD") & "', '" & Logon & "' , '" & INVPayee & "', " & _
                     "'" & Mid(cboINVStayType, 1, 1) & "', '" & Mid(txtINVPayType, 1, 1) & "' , '" & ChkStr(Mid(cboINVPayType, 1, IIf(InStr(1, cboINVPayType, "-") > 0, InStr(1, cboINVPayType, "-") - 1, 2))) & "', " & _
                     "'" & Mid(cboDRM, 1, 2) & "', '" & Format(DIRDate, "YYYY/MM/DD") & "', " & _
                     "" & INVAmount & ", " & MBMPaidAmount & ", '" & CLMUHCNo & "', " & ExcAmt & ", " & PayAmt & ", " & txtInitialActAmt & ", '" & Trim(cboRoomType) & "'"
                 
        ElseIf IsDate(FollowupDate) = False And IsDate(DIRDate) = False Then
            strsql = "EXEC Insert_ClaimInvoice_WO2013 '" & ChkStr(lblLastVersion) & "', " & 0 & " , " & 0 & ", " & _
                     "'" & ChkStr(txtInvNo) & "', '" & ChkStr(txtCaseNo) & "' , '" & Format(InvDate, "YYYY/MM/DD") & "', " & _
                     "'" & Format(EnteredDate, "YYYY/MM/DD") & "', '" & Logon & "' , '" & INVPayee & "', " & _
                     "'" & Mid(cboINVStayType, 1, 1) & "', '" & Mid(txtINVPayType, 1, 1) & "' , '" & ChkStr(Mid(cboINVPayType, 1, IIf(InStr(1, cboINVPayType, "-") > 0, InStr(1, cboINVPayType, "-") - 1, 2))) & "', " & _
                     "'" & Mid(cboDRM, 1, 2) & "', " & _
                     "" & INVAmount & ", " & MBMPaidAmount & ", '" & CLMUHCNo & "', " & ExcAmt & ", " & PayAmt & ", " & txtInitialActAmt & ", '" & Trim(cboRoomType) & "'"
                 
        End If
        
    
    Else
        If IsDate(FollowupDate) And IsDate(DIRDate) Then
            strsql = "EXEC update_ClaimInvoice2013 '" & ChkStr(txtInvNo) & "', '" & ChkStr(txtCaseNo) & "' , '" & Format(InvDate, "YYYY/MM/DD") & "', " & _
                     "'" & INVPayee & "', '" & Mid(cboINVStayType, 1, 1) & "', '" & Mid(txtINVPayType, 1, 1) & "' , '" & ChkStr(Mid(cboINVPayType, 1, IIf(InStr(1, cboINVPayType, "-") > 0, InStr(1, cboINVPayType, "-") - 1, 2))) & "', " & _
                     "'" & Mid(cboDRM, 1, 2) & "', '" & Format(FollowupDate, "YYYY/MM/DD") & "','" & Format(DIRDate, "YYYY/MM/DD") & "', " & _
                     "" & INVAmount & ", " & MBMPaidAmount & ", '" & CLMUHCNo & "', " & ExcAmt & ", " & PayAmt & ", " & txtInitialActAmt & ", '" & Trim(cboRoomType) & "', '" & txtIndex & "'"
        
        ElseIf IsDate(FollowupDate) And IsDate(DIRDate) = False Then
            strsql = "EXEC update_ClaimInvoice_WODIR2013 '" & ChkStr(txtInvNo) & "', '" & ChkStr(txtCaseNo) & "' , '" & Format(InvDate, "YYYY/MM/DD") & "', " & _
                     "'" & INVPayee & "', '" & Mid(cboINVStayType, 1, 1) & "', '" & Mid(txtINVPayType, 1, 1) & "' , '" & ChkStr(Mid(cboINVPayType, 1, IIf(InStr(1, cboINVPayType, "-") > 0, InStr(1, cboINVPayType, "-") - 1, 2))) & "', " & _
                     "'" & Mid(cboDRM, 1, 2) & "', '" & Format(FollowupDate, "YYYY/MM/DD") & "', " & _
                     "" & INVAmount & ", " & MBMPaidAmount & ", '" & CLMUHCNo & "', " & ExcAmt & ", " & PayAmt & " , " & txtInitialActAmt & ", '" & Trim(cboRoomType) & "', '" & txtIndex & "'"
        
        ElseIf IsDate(FollowupDate) = False And IsDate(DIRDate) Then
            strsql = "EXEC update_ClaimInvoice_WOFolDate2013 '" & ChkStr(txtInvNo) & "', '" & ChkStr(txtCaseNo) & "' , '" & Format(InvDate, "YYYY/MM/DD") & "', " & _
                     "'" & INVPayee & "', '" & Mid(cboINVStayType, 1, 1) & "', '" & Mid(txtINVPayType, 1, 1) & "' , '" & ChkStr(Mid(cboINVPayType, 1, IIf(InStr(1, cboINVPayType, "-") > 0, InStr(1, cboINVPayType, "-") - 1, 2))) & "', " & _
                     "'" & Mid(cboDRM, 1, 2) & "', '" & Format(DIRDate, "YYYY/MM/DD") & "', " & _
                     "" & INVAmount & ", " & MBMPaidAmount & ", '" & CLMUHCNo & "', " & ExcAmt & ", " & PayAmt & ", " & txtInitialActAmt & ", '" & Trim(cboRoomType) & "', '" & txtIndex & "'"
                 
        ElseIf IsDate(FollowupDate) = False And IsDate(DIRDate) = False Then
            strsql = "EXEC update_ClaimInvoice_WO2013 '" & ChkStr(txtInvNo) & "', '" & ChkStr(txtCaseNo) & "' , '" & Format(InvDate, "YYYY/MM/DD") & "', " & _
                     "'" & INVPayee & "', '" & Mid(cboINVStayType, 1, 1) & "', '" & Mid(txtINVPayType, 1, 1) & "' , '" & ChkStr(Mid(cboINVPayType, 1, IIf(InStr(1, cboINVPayType, "-") > 0, InStr(1, cboINVPayType, "-") - 1, 2))) & "', " & _
                     "'" & Mid(cboDRM, 1, 2) & "'," & INVAmount & ", " & MBMPaidAmount & ", '" & CLMUHCNo & "', " & ExcAmt & ", " & PayAmt & ", '" & Trim(cboRoomType) & "', " & txtInitialActAmt & ", '" & txtIndex & "'"
        End If
    End If
    
    medixmasterconnWS.Execute strsql
    
    If EditFlag = False Then
        Dim ClaimInvoiceLastIndex  As New ADODB.Recordset
        ClaimInvoiceLastIndex.Open "Select Top 1 INVINdex from ClaimInvoice Order by INVIndex DESC", medixmasterconnWS, adOpenKeyset, adLockOptimistic
        If Not ClaimInvoiceLastIndex.EOF Then
            intIndex = ClaimInvoiceLastIndex!INVIndex
        Else
            intIndex = 1
        End If
        ClaimInvoiceLastIndex.Close
        Set ClaimInvoiceLastIndex = Nothing
    Else
        intIndex = txtIndex
    End If
End Sub

Private Sub saveNotCover()
    Dim ClaimNotCover  As New ADODB.Recordset
    Dim strsql As String
    Dim PreH60 As Double
    Dim PreGP As Double
    Dim PreMedic As Double
    Dim PreMR As Double
    Dim PreSpec As Double
    Dim PostHosp As Double
    Dim PostNotSame  As Double
    Dim Others As Double
    Dim ReimOtherRemarks As String
    Dim tmpHospName As String
    
    PreH60 = IIf(IsNumeric(txtPreHos60) = True, txtPreHos60, 0)
    PreGP = IIf(IsNumeric(txtPreGP) = True, txtPreGP, 0)
    PreMedic = IIf(IsNumeric(txtPreMEdic), txtPreMEdic, 0)
    PreMR = IIf(IsNumeric(txtPreMR), txtPreMR, 0)
    PreSpec = IIf(IsNumeric(txtPreSpec1), txtPreSpec1, 0)
    PostHosp = IIf(IsNumeric(txtPostGHosp) = True, txtPostGHosp, 0)
    PostNotSame = IIf(IsNumeric(txtpostNotSame), txtpostNotSame, 0)
    Others = IIf(IsNumeric(txtOthers), txtOthers, 0)
    ReimOtherRemarks = IIf(Len(txtOthReimNotCovered), txtOthReimNotCovered, "")
    If Trim(txtHOSName) <> "" Then
        tmpHospName = Replace(txtHOSName, "'", "''")
    Else
        tmpHospName = ""
    End If
    
    If EditFlag = False Then
    
        strsql = "EXEC Insert_ClaimInvoiceNotCov08 " & intIndex & ", " & PreH60 & " , " & PreGP & ", " & _
                 "" & PreMedic & ", " & PreMR & " , " & PreSpec & ", " & _
                 "" & PostHosp & ", " & PostNotSame & " , " & Others & " ,  '" & ReimOtherRemarks & "', '" & tmpHospName & "'"
        
    Else
        strsql = "EXEC Update_ClaimInvoiceNotCov08 " & PreH60 & " , " & PreGP & ", " & _
                 "" & PreMedic & ", " & PreMR & " , " & PreSpec & ", " & _
                 "" & PostHosp & ", " & PostNotSame & " , " & Others & " ,  '" & ReimOtherRemarks & "', '" & tmpHospName & "', " & intIndex & ""
    End If
    
    medixmasterconnWS.Execute strsql

End Sub

Private Sub saveActual()
    Dim ClaimActual As New ADODB.Recordset
    Dim strsql As String
    Dim CAS As New ADODB.Recordset
    Dim strsqlCAS As String
    Dim MBM As New ADODB.Recordset
    Dim strsqlMBM As String
    Dim strsqlPLN As String
    Dim strsqlCri2 As String
    Dim strsqlTbl As String
    Dim PLN As New ADODB.Recordset
    Dim cmd As New ADODB.Command
    Dim Prm1 As ADODB.Parameter
    Dim Prm2 As ADODB.Parameter
    Dim Prm3 As ADODB.Parameter
    Dim cmdPLN As New ADODB.Command
    Dim cmdCLM As New ADODB.Command
    
    If A5 = "" Then
        A5 = 0
    End If
    
    If A24 = "" Then
        A24 = 0
    End If
    
    If A6 = "" Then
        A6 = 0
    End If
    
    If A7 = "" Then
        A7 = 0
    End If
    
    If A5B = "" Then
        A5B = 0
    End If
    
    If A8 = "" Then
        A8 = 0
    End If
    
    If A9 = "" Then
        A9 = 0
    End If
    
    If A10 = "" Then
        A10 = 0
    End If
    
    
    sqlstrCAS = "Exec SearchCasAdjustment '" & Trim(txtCaseNo) & "'"
        
    Set CAS = New ADODB.Recordset
    CAS.Open sqlstrCAS, medixmasterconn, adOpenKeyset, adLockOptimistic
    
    If Not CAS.EOF Then
    
        strsqlMBM = " Select MBMNumber, PLNCode "
        strsqlCri = "And MBMNumber = '" & Trim(CAS!MBMNumber) & "'"
        
        Set cmd.ActiveConnection = medixmasterconnMaint
        cmd.CommandText = "SearchMBMCrossRef"
        cmd.CommandType = adCmdStoredProc
        cmd.CommandTimeout = 300
        Set Prm1 = cmd.CreateParameter("strsqlMBM", adVarChar, adParamInput, 255, strsqlMBM)
        cmd.Parameters.Append Prm1
        Set Prm2 = cmd.CreateParameter("strsqlCri", adVarChar, adParamInput, 255, strsqlCri)
        cmd.Parameters.Append Prm2
        MBM.CursorLocation = adUseClient
        MBM.CursorType = adOpenForwardOnly
        MBM.CacheSize = 100
        Set MBM = cmd.Execute
        
        strsqlPLN = "Select PLNCode,NewSMPlan"
        strsqlCri2 = " PLNCode =  '" & Trim(MBM!PLNCode) & "'"
        strsqlTbl = "PLN"
        
        Set cmdPLN.ActiveConnection = medixmasterconn
        cmdPLN.CommandText = "SearchHISTableRec"
        cmdPLN.CommandType = adCmdStoredProc
        cmdPLN.CommandTimeout = 300
        Set Prm1 = cmdPLN.CreateParameter("SelFields", adVarChar, adParamInput, 255, strsqlPLN)
        cmdPLN.Parameters.Append Prm1
        Set Prm2 = cmdPLN.CreateParameter("TableName", adVarChar, adParamInput, 255, strsqlTbl)
        cmdPLN.Parameters.Append Prm2
        Set Prm3 = cmdPLN.CreateParameter("SelCriteria", adVarChar, adParamInput, 255, strsqlCri2)
        cmdPLN.Parameters.Append Prm3
        
        PLN.CursorLocation = adUseClient
        PLN.CursorType = adOpenForwardOnly
        PLN.CacheSize = 100
        Set PLN = cmdPLN.Execute

        If Not PLN.EOF Then
            If PLN!NewSMPlan = "Y" Then
                NewSMPlan = "Y"
            Else
                NewSMPlan = "N"
            End If
        End If
    
    PLN.Close
    Set PLN = Nothing
    
    End If
    If Trim(Mid(MBM!PLNCode, 1, 1)) = "2" Or Trim(Mid(MBM!PLNCode, 1, 1)) = "3" Or Trim(Mid(MBM!PLNCode, 1, 1)) = "4" Or Trim(Mid(MBM!PLNCode, 1, 1)) = "5" Or NewSMPlan = "Y" Then
        Text1 = CDbl(A5) + CDbl(A24) + CDbl(A6) + CDbl(A7)
    Else
        Text1 = 0
    End If
    
    If Trim(Mid(MBM!PLNCode, 1, 1)) = "2" Or Trim(Mid(MBM!PLNCode, 1, 1)) = "3" Or Trim(Mid(MBM!PLNCode, 1, 1)) = "4" Or Trim(Mid(MBM!PLNCode, 1, 1)) = "5" Or NewSMPlan = "Y" Then
        Text2 = CDbl(A5B) + CDbl(A8) + CDbl(A9) + CDbl(A10)
    Else
        Text2 = 0
    End If
    sql = "Exec SearchCasAdjustment '" & Trim(txtCaseNo) & "'"
    
    If EditFlag = False Then
        strsql = "EXEC Insert_ClaimInvActual2010 '" & ChkStr(intIndex) & "', " & IIf(IsNumeric(Text1) = True, Text1, 0) & " , " & IIf(IsNumeric(A1) = True, A1, 0) & ",  " & IIf(IsNumeric(txtGrossRB) = True, txtGrossRB, 0) & "," & _
                    "" & IIf(IsNumeric(A1B) = True, A1B, 0) & ", " & IIf(IsNumeric(A2) = True, A2, 0) & "," & IIf(IsNumeric(A2B) = True, A2B, 0) & ", " & IIf(IsNumeric(A3) = True, A3, 0) & ", " & IIf(IsNumeric(A4) = True, A4, 0) & ", " & _
                    "" & IIf(IsNumeric(A24) = True, A24, 0) & ", " & IIf(IsNumeric(A5) = True, A5, 0) & ", " & IIf(IsNumeric(A5B) = True, A5B, 0) & ", " & IIf(IsNumeric(A6) = True, A6, 0) & ", " & IIf(IsNumeric(A7) = True, A7, 0) & ", " & _
                    "" & IIf(IsNumeric(A8) = True, A8, 0) & ", " & IIf(IsNumeric(A9) = True, A9, 0) & ", " & IIf(IsNumeric(A10) = True, A10, 0) & ", " & IIf(IsNumeric(txtACC) = True, txtACC, 0) & ", " & IIf(IsNumeric(txtPTreat), txtPTreat, 0) & ", " & IIf(IsNumeric(A11) = True, A11, 0) & ", " & _
                    "" & IIf(IsNumeric(A12) = True, A12, 0) & ", " & IIf(IsNumeric(A13) = True, A13, 0) & ", " & IIf(IsNumeric(A14) = True, A14, 0) & ", " & IIf(IsNumeric(A15) = True, A15, 0) & " , " & IIf(IsNumeric(A16) = True, A16, 0) & ", " & IIf(IsNumeric(A17) = True, A17, 0) & ", " & _
                    "" & IIf(IsNumeric(ALB) = True, ALB, 0) & ", " & IIf(IsNumeric(A18) = True, A18, 0) & ", " & IIf(IsNumeric(txtRBTax) = True, txtRBTax, 0) & "," & IIf(IsNumeric(txtICUTax) = True, txtICUTax, 0) & ", " & IIf(IsNumeric(txtLodgerTax) = True, txtLodgerTax, 0) & ", " & IIf(IsNumeric(A19) = True, A19, 0) & "," & IIf(IsNumeric(A20) = True, A20, 0) & ",  " & IIf(IsNumeric(Text2) = True, Text2, 0) & "," & IIf(IsNumeric(A25) = True, A25, 0) & ",'" & ChkStr(txtOhersNoCoPay) & "', " & _
                    "" & IIf(IsNumeric(A20B) = True, A20B, 0) & ", '" & ChkStr(txtOthersDesc1) & "',  '" & ChkStr(txtOthersDesc2) & "'," & IIf(IsNumeric(A21) = True, A21, 0) & ", " & IIf(IsNumeric(A22) = True, A22, 0) & ", " & IIf(IsNumeric(A23) = True, A23, 0) & "," & IIf(IsNumeric(ART) = True, ART, 0) & " ," & IIf(IsNumeric(AIT) = True, AIT, 0) & "," & IIf(IsNumeric(ALT) = True, ALT, 0) & "," & IIf(IsNumeric(A26) = True, A26, 0) & "," & IIf(IsNumeric(A27) = True, A27, 0) & ""
        medixmasterconnWS.Execute strsql
    Else
    
        strsql = "EXEC Update_ClaimInvActual2010 " & IIf(IsNumeric(Text1) = True, Text1, 0) & " , " & IIf(IsNumeric(A1) = True, A1, 0) & ",  " & IIf(IsNumeric(txtGrossRB) = True, txtGrossRB, 0) & "," & _
                    "" & IIf(IsNumeric(A1B) = True, A1B, 0) & ", " & IIf(IsNumeric(A2) = True, A2, 0) & "," & IIf(IsNumeric(A2B) = True, A2B, 0) & ", " & IIf(IsNumeric(A3) = True, A3, 0) & ", " & IIf(IsNumeric(A4) = True, A4, 0) & ", " & _
                    "" & IIf(IsNumeric(A24) = True, A24, 0) & ", " & IIf(IsNumeric(A5) = True, A5, 0) & ", " & IIf(IsNumeric(A5B) = True, A5B, 0) & ", " & IIf(IsNumeric(A6) = True, A6, 0) & ", " & IIf(IsNumeric(A7) = True, A7, 0) & ", " & _
                    "" & IIf(IsNumeric(A8) = True, A8, 0) & ", " & IIf(IsNumeric(A9) = True, A9, 0) & ", " & IIf(IsNumeric(A10) = True, A10, 0) & ", " & IIf(IsNumeric(txtACC) = True, txtACC, 0) & ", " & IIf(IsNumeric(txtPTreat), txtPTreat, 0) & ", " & IIf(IsNumeric(A11) = True, A11, 0) & ", " & _
                    "" & IIf(IsNumeric(A12) = True, A12, 0) & ", " & IIf(IsNumeric(A13) = True, A13, 0) & ", " & IIf(IsNumeric(A14) = True, A14, 0) & ", " & IIf(IsNumeric(A15) = True, A15, 0) & " , " & IIf(IsNumeric(A16) = True, A16, 0) & ", " & IIf(IsNumeric(A17) = True, A17, 0) & ", " & _
                    "" & IIf(IsNumeric(ALB) = True, ALB, 0) & ", " & IIf(IsNumeric(A18) = True, A18, 0) & ", " & IIf(IsNumeric(txtRBTax) = True, txtRBTax, 0) & "," & IIf(IsNumeric(txtICUTax) = True, txtICUTax, 0) & ", " & IIf(IsNumeric(txtLodgerTax) = True, txtLodgerTax, 0) & ", " & IIf(IsNumeric(A19) = True, A19, 0) & "," & IIf(IsNumeric(A20) = True, A20, 0) & ",  " & IIf(IsNumeric(Text2) = True, Text2, 0) & "," & IIf(IsNumeric(A25) = True, A25, 0) & ",'" & ChkStr(txtOhersNoCoPay) & "', " & _
                    "" & IIf(IsNumeric(A20B) = True, A20B, 0) & ", '" & ChkStr(txtOthersDesc1) & "',  '" & ChkStr(txtOthersDesc2) & "'," & IIf(IsNumeric(A21) = True, A21, 0) & ", " & IIf(IsNumeric(A22) = True, A22, 0) & ", " & IIf(IsNumeric(A23) = True, A23, 0) & "," & IIf(IsNumeric(ART) = True, ART, 0) & " ," & IIf(IsNumeric(AIT) = True, AIT, 0) & "," & IIf(IsNumeric(ALT) = True, ALT, 0) & "," & IIf(IsNumeric(A26) = True, A26, 0) & "," & IIf(IsNumeric(A27) = True, A27, 0) & ", '" & ChkStr(intIndex) & "'"
        medixmasterconnWS.Execute strsql
    End If
    CAS.Close
    Set CAS = Nothing
    MBM.Close
    Set MBM = Nothing
End Sub


Private Sub saveActualOthers()
    Dim ClaimActualOthers  As New ADODB.Recordset
    Dim strsql  As String
    Dim MedRptMBM As String
    Dim TelChargesMBM As String
    Dim GuaAllowMBM As String
    Dim AdminRegMBM As String
    Dim OthersMBM As String
    Dim CLMBoardDay, CLMRB, CLMICUDay, CLMICU As Double
    Dim CLMPreXRay, CLMPreLab, CLMPreoTHERS, CLMMXRay As Double
    Dim CLMMLab, CLMMPreOthers, CLMSurgicalFee, CLMSurgPhy As Double
    Dim CLMSurgPost, CLMINHospHyVisitDay, CLMKidneyMnth As Double
    Dim CLMLodger, CLMLodgerday, CLMCashDay, CLMCash As Double
    Dim CLMSurgicalWard, CLMAnaeFees, CLMAnaeConsultation As Double
    Dim CLMAnaeWard, CLMInHosFees, CLMInHosWard, CLMProceduresFee As Double
    Dim CLMPharmacyHSS, CLMMedSuppliesHSS, CLMNursingCareHSS, CLMNursingSvrHSS As Double
    Dim CLMLabHSS, CLMRadiologiHSS, CLMOTSuppHSS, CLMEquipHSS, CLMMiscHSS As Double
    
    CLMBoardDay = IIf(IsNumeric(DB1), DB1, 0)
    CLMRB = IIf(IsNumeric(RoomAmount), RoomAmount, 0)
    CLMICUDay = IIf(IsNumeric(ICU1), ICU1, 0)
    CLMICU = IIf(IsNumeric(ICUAmount), ICUAmount, 0)
    CLMPreXRay = IIf(IsNumeric(txtXray), txtXray, 0)
    CLMPreLab = IIf(IsNumeric(txtPreLab), txtPreLab, 0)
    CLMPreoTHERS = IIf(IsNumeric(txtPreOthers), txtPreOthers, 0)
    CLMMXRay = IIf(IsNumeric(TxtMXRay), TxtMXRay, 0)
    CLMMLab = IIf(IsNumeric(TxtMLab), TxtMLab, 0)
    CLMMPreOthers = IIf(IsNumeric(TxtMPreOthers), TxtMPreOthers, 0)
    CLMSurgicalFee = IIf(IsNumeric(txtSurgicalFee), txtSurgicalFee, 0)
    CLMSurgPhy = IIf(IsNumeric(txtPhysician), txtPhysician, 0)
    CLMSurgPost = IIf(IsNumeric(txtPostTreat), txtPostTreat, 0)
    CLMINHospHyVisitDay = IIf(IsNumeric(txtInPhyVisit), txtInPhyVisit, 0)
    CLMKidneyMnth = IIf(IsNumeric(txtKidneyMonths), txtKidneyMonths, 0)
    CLMLodger = IIf(IsNumeric(txtLodgerAmt), txtLodgerAmt, 0)
    CLMLodgerday = IIf(IsNumeric(txtLodgerDay), txtLodgerDay, 0)
    CLMCashDay = IIf(IsNumeric(txtDayCash), txtDayCash, 0)
    CLMCash = IIf(IsNumeric(txtRMCash), txtRMCash, 0)
    CLMSurgicalWard = IIf(IsNumeric(txtSurgWard), txtSurgWard, 0)
    CLMAnaeFees = IIf(IsNumeric(txtAnaFees), txtAnaFees, 0)
    CLMAnaeConsultation = IIf(IsNumeric(txtAnaCons), txtAnaCons, 0)
    CLMAnaeWard = IIf(IsNumeric(txtAnaWard), txtAnaWard, 0)
    CLMInHosFees = IIf(IsNumeric(txtInHosFees), txtInHosFees, 0)
    CLMInHosWard = IIf(IsNumeric(txtInHosWardFees), txtInHosWardFees, 0)
    CLMProceduresFee = IIf(IsNumeric(txtProcFees), txtProcFees, 0)
        
    If chkMedRptMBM.Value = 1 Then
        MedRptMBM = "Y"
    Else
        MedRptMBM = "N"
    End If
    
    If chkTelChargesMBM.Value = 1 Then
        TelChargesMBM = "Y"
    Else
        TelChargesMBM = "N"
    End If
    
    If chkGuaAllowMBM.Value = 1 Then
        GuaAllowMBM = "Y"
    Else
        GuaAllowMBM = "N"
    End If
    
    If chkAdminRegMBM.Value = 1 Then
        AdminRegMBM = "Y"
    Else
        AdminRegMBM = "N"
    End If
    
    If chkOthersMBM.Value = 1 Then
        OthersMBM = "Y"
    Else
        OthersMBM = "N"
    End If

    CLMPharmacyHSS = IIf(IsNumeric(frmHSSDetails.txtInvPharmacy), frmHSSDetails.txtInvPharmacy, 0)
    CLMMedSuppliesHSS = IIf(IsNumeric(frmHSSDetails.txtMedSupp), frmHSSDetails.txtMedSupp, 0)
    CLMNursingCareHSS = IIf(IsNumeric(frmHSSDetails.txtNursingCare), frmHSSDetails.txtNursingCare, 0)
    CLMNursingSvrHSS = IIf(IsNumeric(frmHSSDetails.txtNursingServ), frmHSSDetails.txtNursingServ, 0)
    CLMLabHSS = IIf(IsNumeric(frmHSSDetails.txtLab), frmHSSDetails.txtLab, 0)
    CLMRadiologiHSS = IIf(IsNumeric(frmHSSDetails.txtRadiology), frmHSSDetails.txtRadiology, 0)
    CLMOTSuppHSS = IIf(IsNumeric(frmHSSDetails.txtOTSupp), frmHSSDetails.txtOTSupp, 0)
    CLMEquipHSS = IIf(IsNumeric(frmHSSDetails.txtEquipChrg), frmHSSDetails.txtEquipChrg, 0)
    CLMMiscHSS = IIf(IsNumeric(frmHSSDetails.txtMisc), frmHSSDetails.txtMisc, 0)
    
    If EditFlag = False Then
        strsql = "EXEC Insert_ClaimInvoiceActOthers_New " & intIndex & ", " & CLMBoardDay & " , " & CLMRB & ", " & _
                 "" & CLMICUDay & ", " & CLMICU & " , " & CLMPreXRay & ", " & _
                 "" & CLMPreLab & ", " & CLMPreoTHERS & " , " & CLMMXRay & ", " & _
                 "" & CLMMLab & ", " & CLMMPreOthers & " , " & CLMSurgicalFee & ", " & _
                 "" & CLMSurgPhy & ", " & CLMSurgPost & ", " & CLMINHospHyVisitDay & ", " & _
                 "" & CLMKidneyMnth & ", " & CLMLodger & ", " & CLMLodgerday & ", " & _
                 "" & CLMCashDay & ", " & CLMCash & ", " & CLMSurgicalWard & ", " & CLMAnaeFees & ", " & _
                 "" & CLMAnaeConsultation & ", " & CLMAnaeWard & ", " & CLMInHosFees & ", " & CLMInHosWard & ",  " & CLMProceduresFee & ", " & _
                 "'" & MedRptMBM & "', '" & TelChargesMBM & "', '" & GuaAllowMBM & "', '" & AdminRegMBM & "' , '" & OthersMBM & "', " & _
                 "" & CLMPharmacyHSS & ", " & CLMMedSuppliesHSS & ", " & CLMNursingCareHSS & ", " & _
                 "" & CLMNursingSvrHSS & ", " & CLMLabHSS & ", " & CLMRadiologiHSS & ", " & CLMOTSuppHSS & "," & _
                 "" & CLMEquipHSS & ", " & CLMMiscHSS & ""
    Else
        strsql = "EXEC Update_ClaimInvoiceActOthers_New " & CLMBoardDay & " , " & CLMRB & ", " & _
                 "" & CLMICUDay & ", " & CLMICU & " , " & CLMPreXRay & ", " & _
                 "" & CLMPreLab & ", " & CLMPreoTHERS & " , " & CLMMXRay & ", " & _
                 "" & CLMMLab & ", " & CLMMPreOthers & " , " & CLMSurgicalFee & ", " & _
                 "" & CLMSurgPhy & ", " & CLMSurgPost & ", " & CLMINHospHyVisitDay & ", " & _
                 "" & CLMKidneyMnth & ", " & CLMLodger & ", " & CLMLodgerday & ", " & _
                 "" & CLMCashDay & ", " & CLMCash & ", " & CLMSurgicalWard & ", " & CLMAnaeFees & ", " & _
                 "" & CLMAnaeConsultation & ", " & CLMAnaeWard & ", " & CLMInHosFees & ", " & CLMInHosWard & ",  " & CLMProceduresFee & ", " & _
                 "'" & MedRptMBM & "', '" & TelChargesMBM & "', '" & GuaAllowMBM & "', '" & AdminRegMBM & "' , '" & OthersMBM & "', " & _
                 "" & CLMPharmacyHSS & ", " & CLMMedSuppliesHSS & ", " & CLMNursingCareHSS & ", " & _
                 "" & CLMNursingSvrHSS & ", " & CLMLabHSS & ", " & CLMRadiologiHSS & ", " & CLMOTSuppHSS & "," & _
                 "" & CLMEquipHSS & ", " & CLMMiscHSS & ", " & intIndex & ""
                 
                 
    End If
    medixmasterconnWS.Execute strsql
End Sub
Private Sub saveLimit()
Dim strsqlCLMLmt As String
Dim CLMLmt As New ADODB.Recordset
Dim strsqlSameLmt As String
Dim SameLmt As New ADODB.Recordset
Dim strsqlActualLmt As String
Dim ActualLmt As New ADODB.Recordset
Dim BalBoardDays, BalICUDays, BalCashDays, BalGurdianDays As Double
Dim BalINHOSPhyDays, BalKIDDCDays As Double
Dim BalSurgicalAmt, BalCLMAna, BalOrgan As Double
    strsqlCLMLmt = "Select CLMBoardDays, CLMICUDays, CLMCashDays, " & _
                   "CLMGurdianDays, CLMINHOSPhyDays, CLMKIDDCDays, " & _
                   "CLMSurgicalAmt, CLMAna, CLMOrgan, " & _
                   "CLMBoardAmt, CLMICUAmt, CLMTax, CLMGuardianAmt, CLMMonthly  from ClaimLimit " & _
                   "Where CASNumber = '" & Trim(txtCaseNo) & "' And CLMVersion = '" & Trim(lblVersion) & "'"
    CLMLmt.Open strsqlCLMLmt, medixmasterconnWS, adOpenKeyset, adLockOptimistic
    
    If CLMLmt.recordCount Then
        strsqlSameLmt = "Select CLMBoardDays, CLMICUDays, CLMCashDays, " & _
                        "CLMSurgicalAmt, CLMAna, CLMOrgan, " & _
                        "CLMGurdianDays, CLMINHOSPhyDays, CLMKIDDCDays from ClaimSameLimit " & _
                        "Where CASNumber = '" & Trim(txtCaseNo) & "'"
        SameLmt.Open strsqlSameLmt, medixmasterconnWS, adOpenKeyset, adLockOptimistic
        If SameLmt.recordCount Then
            'SameLmt.AddNew
        'End If
        Call display_benefit
            If CDbl(SameLmt!CLMBoardDays) <> 0 Or SameLmt!CLMBoardDays <> "" Then
                If IsNumeric(CLMLmt!CLMBoardDays) = False Then
                    CLMLmt!CLMBoardDays = 0
                End If
                BalBoardDays = CInt(CLMLmt!CLMBoardDays) - CInt(DB1)
                CLMLmt!CLMBoardDays = DB1
                'Need to check is Full
                If IsNumeric(RBlimit) Then
                    CLMLmt!CLMBoardAmt = Format(DB1 * RBlimit, "#,##0.00")
                Else
                    CLMLmt!CLMBoardAmt = Format(DB1, "#,##0.00")
                End If
            End If
            
            If CDbl(SameLmt!CLMICUDays) <> 0 Then
                BalICUDays = CDbl(CLMLmt!CLMICUDays) - CDbl(ICU1)
                CLMLmt!CLMICUDays = ICU1
                If IsNumeric(ICUlimit) Then
                    CLMLmt!CLMICUAmt = Format(ICU1 * ICUlimit, "#,##0.00")
                Else
                    CLMLmt!CLMICUAmt = ICUlimit
                End If
            End If
            If CDbl(SameLmt!CLMCashDays) <> 0 Then
                If Len(CLMLmt!CLMCashDays) And CLMLmt!CLMCashDays <> "" Then
                    BalCashDays = IIf(Len(CLMLmt!CLMCashDays), CDbl(CLMLmt!CLMCashDays), 0) - IIf(Len((txtDayCash)), CDbl(txtDayCash), 0)
                Else
                    BalCashDays = 0
                End If
                'BalCashDays = IIf(Len(CLMLmt!CLMCashDays), CDbl(CLMLmt!CLMCashDays), 0) - IIf(Len((txtDayCash)), CDbl(txtDayCash), 0)
                CLMLmt!CLMCashDays = txtDayCash
                CLMLmt!CLMTax = Format(txtDayCash * IIf(Len(CashLimit), CashLimit, 0), "#,##0.00")
            End If
            
            If (Len(SameLmt!CLMGurdianDays)) Then
              If CDbl(SameLmt!CLMGurdianDays) <> 0 Then
                If IsNumeric(CLMLmt!CLMGurdianDays) Then
                    BalGurdianDays = CDbl(CLMLmt!CLMGurdianDays) - IIf(Len(txtLodgerDay), CDbl(txtLodgerDay), 0)
                Else
                    BalGurdianDays = 0
                End If
                    
                CLMLmt!CLMGurdianDays = txtLodgerDay
                CLMLmt!CLMGuardianAmt = Format(IIf(Len(txtLodgerDay), txtLodgerDay, 0) * (IIf(Len(GuardianLimit), GuardianLimit, 0)), "#,##0.00")
              End If
            End If
            
            If Len(CLMLmt!CLMINHOSPhyDays) Then
            If IsNumeric(CLMLmt!CLMINHOSPhyDays) = False Then
                CLMLmt!CLMINHOSPhyDays = 0
            End If
            Else
            If IsNumeric(CLMLmt!CLMINHOSPhyDays) = False Then
                CLMLmt!CLMINHOSPhyDays = 0
            End If
            End If
            If Len(SameLmt!CLMINHOSPhyDays) Then
            If CDbl(SameLmt!CLMINHOSPhyDays) <> 0 Then
                BalINHOSPhyDays = CDbl(CLMLmt!CLMINHOSPhyDays) - IIf(Len(txtInPhyVisit), CDbl(txtInPhyVisit), 0)
                CLMLmt!CLMINHOSPhyDays = IIf(Len(txtInPhyVisit), txtInPhyVisit, 0)
            End If
            End If
            
            If Len(SameLmt!CLMKIDDCDays) Then
            If CDbl(SameLmt!CLMKIDDCDays) <> 0 Then
                If IsNumeric(CLMLmt!CLMKIDDCDays) = False Then
                    CLMLmt!CLMKIDDCDays = 0
                End If
                
                BalKIDDCDays = CDbl(CLMLmt!CLMKIDDCDays) - IIf(Len(txtKidneyMonths), CDbl(txtKidneyMonths), 0)
                CLMLmt!CLMKIDDCDays = IIf(Len(txtKidneyMonths), CDbl(txtKidneyMonths), 0)
                CLMLmt!CLMMonthly = Format(KidneyLimit * IIf(Len(txtKidneyMonths), CDbl(txtKidneyMonths), 0), "#,##0.00")
            End If
            End If
        CLMLmt.Update
        CLMLmt.Close
        Set CLMLmt = Nothing
        
        'strsqlActualLmt = "Select CLMSurgicalAmt, CLMAna, CLMOrgan from ClaimActual " & _
        '                "Where CASNumber = '" & Trim(txtCaseNo) & "' And CLMVersion = '" & Trim(lblVersion) & "'"
        'ActualLmt.Open strsqlActualLmt, medixmasterconnWS, adOpenKeyset, adLockOptimistic
        
        'If ActualLmt.recordCount Then
        'If IsNumeric(SameLmt!CLMSurgicalAmt) = True Then
        '    If CDbl(SameLmt!CLMSurgicalAmt) <> 0 Then
        '        If IsNumeric(ActualLmt!CLMSurgicalAmt) = False Then
        '            ActualLmt!CLMSurgicalAmt = 0
        '        End If
        '            BalSurgicalAmt = CDbl(ActualLmt!CLMSurgicalAmt) - CDbl(A6)
        '
        '        If IsNumeric(BalSurgicalAmt) Then
        '            BalSurgicalAmt = Round(BalSurgicalAmt, 2)
        '        End If
        '    End If
        '        ActualLmt!CLMSurgicalAmt = A6
        'End If
        
       ' If IsNumeric(SameLmt!CLMAna) = True Then
       '     If CDbl(SameLmt!CLMAna) <> 0 Then
       '         If IsNumeric(ActualLmt!CLMAna) = False Then
       '             ActualLmt!CLMAna = 0
       '         End If
       '             BalCLMAna = CDbl(ActualLmt!CLMAna) - CDbl(A7)
       '     End If
       '        ActualLmt!CLMAna = A7
       ' End If
        
       ' If IsNumeric(SameLmt!CLMOrgan) = True Then
       '     If CDbl(SameLmt!CLMOrgan) <> 0 Then
       '         If IsNumeric(ActualLmt!CLMOrgan) = False Then
       '             ActualLmt!CLMOrgan = 0
       '         End If
       '             BalOrgan = CDbl(ActualLmt!CLMOrgan) - CDbl(A22)
       '     End If
       '     ActualLmt!CLMOrgan = A22
       ' End If
       '     ActualLmt.Update
       '     ActualLmt.Close
       '     Set ActualLmt = Nothing
       ' End If
            If IsNumeric(SameLmt!CLMBoardDays) = True Then
                If CDbl(SameLmt!CLMBoardDays) <> 0 Then
                    If BalBoardDays >= 0 Then
                        SameLmt!CLMBoardDays = CDbl(SameLmt!CLMBoardDays) + CDbl(BalBoardDays)
                    Else
                        SameLmt!CLMBoardDays = CDbl(SameLmt!CLMBoardDays) - (CDbl(BalBoardDays) * -1)
                    End If
                End If
            End If
            If IsNumeric(SameLmt!CLMICUDays) = True Then
                If CDbl(SameLmt!CLMICUDays) <> 0 Then
                    If BalICUDays >= 0 Then
                        SameLmt!CLMICUDays = CDbl(SameLmt!CLMICUDays) + CDbl(BalICUDays)
                    Else
                        SameLmt!CLMICUDays = CDbl(SameLmt!CLMICUDays) - (CDbl(BalICUDays) * -1)
                    End If
                End If
            End If
            
            If IsNumeric(SameLmt!CLMCashDays) = True Then
                If CDbl(SameLmt!CLMCashDays) <> 0 Then
                    If BalCashDays >= 0 Then
                        SameLmt!CLMCashDays = CDbl(SameLmt!CLMCashDays) + CDbl(BalCashDays)
                    Else
                        SameLmt!CLMCashDays = CDbl(SameLmt!CLMCashDays) - (CDbl(BalCashDays) * -1)
                    End If
                End If
            End If
            If IsNumeric(SameLmt!CLMGurdianDays) = True Then
                If CDbl(SameLmt!CLMGurdianDays) <> 0 Then
                    If BalGurdianDays >= 0 Then
                        SameLmt!CLMGurdianDays = CDbl(SameLmt!CLMGurdianDays) + CDbl(BalGurdianDays)
                    Else
                        SameLmt!CLMGurdianDays = CDbl(SameLmt!CLMGurdianDays) - (CDbl(BalGurdianDays) * -1)
                    End If
                End If
            End If
            
            If IsNumeric(SameLmt!CLMINHOSPhyDays) = True Then
                If CDbl(SameLmt!CLMINHOSPhyDays) <> 0 Then
                    If BalINHOSPhyDays >= 0 Then
                        SameLmt!CLMINHOSPhyDays = CDbl(SameLmt!CLMINHOSPhyDays) + CDbl(BalINHOSPhyDays)
                    Else
                        SameLmt!CLMINHOSPhyDays = CDbl(SameLmt!CLMINHOSPhyDays) - (CDbl(BalINHOSPhyDays) * -1)
                    End If
                End If
            End If
            If IsNumeric(SameLmt!CLMKIDDCDays) = True Then
                If CDbl(SameLmt!CLMKIDDCDays) <> 0 Then
                    If BalKIDDCDays >= 0 Then
                        SameLmt!CLMKIDDCDays = CDbl(SameLmt!CLMKIDDCDays) + CDbl(BalKIDDCDays)
                    Else
                        SameLmt!CLMKIDDCDays = CDbl(SameLmt!CLMKIDDCDays) - (CDbl(BalKIDDCDays) * -1)
                    End If
                End If
            End If
            'If IsNumeric(SameLmt!CLMSurgicalAmt) = True Then
            '    If CDbl(SameLmt!CLMSurgicalAmt) <> 0 Then
            '        If BalSurgicalAmt >= 0 Then
            '            SameLmt!CLMSurgicalAmt = CDbl(SameLmt!CLMSurgicalAmt) + CDbl(BalSurgicalAmt)
            '        Else
            '            SameLmt!CLMSurgicalAmt = CDbl(SameLmt!CLMSurgicalAmt) - (CDbl(BalSurgicalAmt) * -1)
            '            If CDbl(SameLmt!CLMSurgicalAmt) < 0 Then
            '                SameLmt!CLMSurgicalAmt = 0
            '            End If
            '        End If
            '    End If
            'End If
            
            'If IsNumeric(SameLmt!CLMAna) = True Then
            'If CDbl(SameLmt!CLMAna) <> 0 Then
            '    If BalCLMAna >= 0 Then
            '        SameLmt!CLMAna = CDbl(SameLmt!CLMAna) + CDbl(BalCLMAna)
            '    Else
            '        SameLmt!CLMAna = CDbl(SameLmt!CLMAna) - (CDbl(BalCLMAna) * -1)
            '            If CDbl(SameLmt!CLMAna) < 0 Then
            '                SameLmt!CLMAna = 0
            '            End If
            '    End If
            'End If
            'End If
            
            'If IsNumeric(SameLmt!CLMOrgan) = True Then
            '    If SameLmt!CLMOrgan = "-" Then
            '        SameLmt!CLMOrgan = "0"
            '    End If
            'End If
            'If IsNumeric(SameLmt!CLMOrgan) = True Then
            'If CDbl(SameLmt!CLMOrgan) <> 0 Then
            '    If BalOrgan >= 0 Then
            '        SameLmt!CLMOrgan = CDbl(SameLmt!CLMOrgan) + CDbl(BalOrgan)
            '    Else
            '        SameLmt!CLMOrgan = CDbl(SameLmt!CLMOrgan) - (CDbl(BalOrgan) * -1)
            '            If CDbl(SameLmt!CLMOrgan) < 0 Then
            '                SameLmt!CLMOrgan = 0
            '            End If
            '    End If
            'End If
            'End If
            SameLmt.Update
            
        Else
                CLMLmt!CLMBoardDays = DB1
                CLMLmt!CLMICUDays = ICU1
                CLMLmt!CLMCashDays = txtDayCash
                CLMLmt!CLMGurdianDays = txtLodgerDay
                CLMLmt!CLMINHOSPhyDays = txtInPhyVisit
        CLMLmt.Update
        CLMLmt.Close
        Set CLMLmt = Nothing
        
        End If
            
            SameLmt.Close
            Set SameLmt = Nothing
    End If
End Sub
Public Sub display_benefit()
    Dim BNF As New ADODB.Recordset
    Dim strsql As String
    Dim counter As Integer
    Dim i  As Integer
        
        strsql = "Select BNFCode , BNFNoOFDays, BNFClaimableAmount , BNFRMPerDis  , BNFEffDate from BNF  " & _
                " where PLnCode = '" & PLNCode & "' and INSCode = '" & Trim(INSCode) & "' and HLTCode = '" & Trim(HLTCode) & "' "
        strsql = strsql & " order by BNFCOde "
        BNF.Open strsql, medixmasterconn, adOpenKeyset, adLockOptimistic
        Do Until BNF.EOF
             With BNF
                For counter = 1 To 23 Step 1
                    If !BNFClaimAbleAmount <> "" Then
                        If !BNFCode = "101" Then
                            RBlimit = !BNFClaimAbleAmount
                        ElseIf !BNFCode = "102" Then
                            ICUlimit = !BNFClaimAbleAmount
                        ElseIf !BNFCode = "601" Then
                            CashLimit = Format(!BNFClaimAbleAmount, "#,#.00")
                            CashDays = !BNFNoOfDays
                        ElseIf !BNFCode = "701" Then
                            GuardianLimit = !BNFClaimAbleAmount
                            GuardianDays = !BNFNoOfDays
                        End If
                    End If

                    If !BNFNoOfDays <> "" And !BNFNoOfDays <> 0 Then
                        If !BNFCode = "302" Then
                            InHosPhycisianDays = !BNFNoOfDays
                        ElseIf !BNFCode = "404" Then
                            KidneyLimit = !BNFRMperDis
                            KidneyDCDays = !BNFNoOfDays
                        End If
                    End If
                    Exit For
            Next
            .MoveNext
         End With
    Loop
    
    BNF.Close
    Set BNF = Nothing
End Sub

' key press
Private Sub CboPayType_KeyPress(KeyAscii As Integer)
Dim NewText As String
Dim ValidCount As Integer
Dim ValidValue As String
Dim i As Integer
    
    If KeyAscii >= 32 And KeyAscii <> 127 Then
    NewText = LCase(Left(CboPayType.Text, CboPayType.SelStart) + Chr(KeyAscii) + Mid(CboPayType.Text, CboPayType.SelStart + CboPayType.SelLength + 1))

    ValidCount = 0
    ValidValue = ""

    For i = 0 To CboPayType.ListCount - 1
 
        If NewText = LCase(Left(CboPayType.List(i), Len(NewText))) Then
                ValidCount = ValidCount + 1
                ValidValue = CboPayType.List(i)
        End If
        Next
        If ValidCount <= 1 Then KeyAscii = 0
             If ValidCount = 1 Then
               CboPayType.Text = ValidValue
               CboPayType.SelStart = 0
               CboPayType.SelLength = Len(ValidValue)
             End If
        End If
End Sub


Private Sub cboINVPayType_KeyPress(KeyAscii As Integer)
Dim NewText As String
Dim ValidCount As Integer
Dim ValidValue As String
Dim i As Integer
    
    If KeyAscii >= 32 And KeyAscii <> 127 Then
    NewText = LCase(Left(cboINVPayType.Text, cboINVPayType.SelStart) + Chr(KeyAscii) + Mid(cboINVPayType.Text, cboINVPayType.SelStart + cboINVPayType.SelLength + 1))

    ValidCount = 0
    ValidValue = ""

    For i = 0 To cboINVPayType.ListCount - 1
 
        If NewText = LCase(Left(cboINVPayType.List(i), Len(NewText))) Then
                ValidCount = ValidCount + 1
                ValidValue = cboINVPayType.List(i)
        End If
        Next
        If ValidCount <= 1 Then KeyAscii = 0
             If ValidCount = 1 Then
               cboINVPayType.Text = ValidValue
               cboINVPayType.SelStart = 0
               cboINVPayType.SelLength = Len(ValidValue)
             End If
        End If
End Sub


'*************************** Diable Non Integer Value
Private Sub DB1_KeyPress(KeyAscii As Integer)
    KeyAscii = KeyNumber(KeyAscii)
End Sub


Private Sub RoomAmount_KeyPress(KeyAscii As Integer)
    KeyAscii = KeyNumber(KeyAscii)
End Sub


Private Sub A1_KeyPress(KeyAscii As Integer)
    KeyAscii = KeyNumber(KeyAscii)
End Sub

Private Sub A2_KeyPress(KeyAscii As Integer)
    KeyAscii = KeyNumber(KeyAscii)
End Sub

Private Sub A3_KeyPress(KeyAscii As Integer)
    KeyAscii = KeyNumber(KeyAscii)
End Sub

Private Sub A4_KeyPress(KeyAscii As Integer)
    KeyAscii = KeyNumber(KeyAscii)
End Sub

Private Sub A5_KeyPress(KeyAscii As Integer)
    KeyAscii = KeyNumber(KeyAscii)
End Sub

Private Sub A6_KeyPress(KeyAscii As Integer)
    KeyAscii = KeyNumber(KeyAscii)
End Sub

Private Sub A7_KeyPress(KeyAscii As Integer)
    KeyAscii = KeyNumber(KeyAscii)
End Sub

Private Sub A8_KeyPress(KeyAscii As Integer)
    KeyAscii = KeyNumber(KeyAscii)
End Sub

Private Sub A9_KeyPress(KeyAscii As Integer)
    KeyAscii = KeyNumber(KeyAscii)
End Sub


Private Sub A10_KeyPress(KeyAscii As Integer)
    KeyAscii = KeyNumber(KeyAscii)
End Sub

Private Sub A11_KeyPress(KeyAscii As Integer)
    KeyAscii = KeyNumber(KeyAscii)
End Sub

Private Sub txtACC_KeyPress(KeyAscii As Integer)
    KeyAscii = KeyNumber(KeyAscii)
End Sub

Private Sub txtPTreat_KeyPress(KeyAscii As Integer)
    KeyAscii = KeyNumber(KeyAscii)
End Sub

Private Sub A12_KeyPress(KeyAscii As Integer)
    KeyAscii = KeyNumber(KeyAscii)
End Sub

Private Sub A13_KeyPress(KeyAscii As Integer)
    KeyAscii = KeyNumber(KeyAscii)
End Sub

Private Sub A14_KeyPress(KeyAscii As Integer)
    KeyAscii = KeyNumber(KeyAscii)
End Sub

Private Sub A15_KeyPress(KeyAscii As Integer)
    KeyAscii = KeyNumber(KeyAscii)
End Sub

Private Sub A16_KeyPress(KeyAscii As Integer)
    KeyAscii = KeyNumber(KeyAscii)
End Sub

Private Sub A17_KeyPress(KeyAscii As Integer)
    KeyAscii = KeyNumber(KeyAscii)
End Sub

Private Sub A18_KeyPress(KeyAscii As Integer)
    KeyAscii = KeyNumber(KeyAscii)
End Sub

Private Sub A19_KeyPress(KeyAscii As Integer)
    KeyAscii = KeyNumber(KeyAscii)
End Sub

Private Sub A20_KeyPress(KeyAscii As Integer)
    KeyAscii = KeyNumber(KeyAscii)
End Sub

Private Sub A21_KeyPress(KeyAscii As Integer)
    KeyAscii = KeyNumber(KeyAscii)
End Sub
Private Sub funcword()
On Error GoTo openerror
    ' make sure only one word application launched
    txtSRVPath = crdPath

    Set objword = Nothing
    If objword Is Nothing Then
            Set objword = New Word.Application
            objword.Visible = True
            Set objdoc = objword.Documents.Add(txtSRVPath & "\template\" & "\FolREPLetter.DOT")
    ElseIf Not objdoc Is Nothing Then
          If MsgBox("Previous document " & objdoc & " not closed. Press OK to save and close it now or cancel to abort", _
            vbExclamation + vbOKCancel, DialogBoxTitle) = vbOK Then
            objdoc.Save
            objdoc.Close
          ElseIf vbCancel Then
            Exit Sub
          End If
            Set objdoc = Nothing
            Set objdoc = objword.Documents.Add(txtSRVPath & "\template\" & "FolREPLetter.DOT")
    End If
    bookmark ("FolREPLetter.DOT")
    Exit Sub
openerror:
    Me.Show
    If Err.Number = "5356" Then
        MsgBox Err.Description & ". Please close the previous document first using Word or save it to other name", vbExclamation + vbOKOnly, DialogBoxTitle
    ElseIf Err.Number = "462" Then
        Set objword = New Word.Application
        objword.Visible = True
    ElseIf Err.Number = "4198" Then
        Err.Clear
    Else
        MsgBox Err.Description & Err.Number, vbExclamation + vbOKOnly, DialogBoxTitle
    End If
End Sub

Private Function bookmark(Optional strAppend As String) As Integer
Dim strsqlPay As String
Dim PAY As New ADODB.Recordset
Dim strsqlRep As String
Dim REP As New ADODB.Recordset
Dim HOS As New ADODB.Recordset
Dim strsqlHOS As String
Dim strsqlLG As String
Dim LG As New ADODB.Recordset
Dim REPCat As New ADODB.Recordset
Dim strsqlREPCat As String

    strsqlPay = "Select * from PAY where PAYCode = '" & Trim(Mid(txtCaseNo, 1, 2)) & "'"
    PAY.Open strsqlPay, medixmasterconn, adOpenKeyset, adLockOptimistic
    
    HOspitals = Replace(cboHosCode, "'", "''")
    strsqlHOS = "select * from HOS where HOSName = '" & Trim(HOspitals) & "'"
    HOS.Open strsqlHOS, medixmasterconn, adOpenKeyset, adLockOptimistic
    
        bookmark = False
        If txtCaseNo <> "" Then
            objdoc.Bookmarks("CASENo").Range.Text = txtCaseNo
        End If

        If Trim(txtPatientName) <> "" Then
          objdoc.Bookmarks("patientname").Range.Text = txtPatientName
        Else
          objdoc.Bookmarks("patientname").Range.Text = " Please key in patient name here!"
        End If
        If txtMBMNumber <> "" Then
            objdoc.Bookmarks("mbmnumber").Range.Text = txtMBMNumber
        End If
        If cboHosCode <> "" Then
            objdoc.Bookmarks("hospital").Range.Text = Trim(cboHosCode)
        End If
        If txtMBMPolNo <> "" Then
            objdoc.Bookmarks("policynumber").Range.Text = txtMBMPolNo
        End If
        If txtMBMName <> "" Then
            objdoc.Bookmarks("InsuredName").Range.Text = txtMBMName
        End If
                
        If txtPreHos60 <> 0 And txtPreHos60 <> "" Then
            objdoc.Bookmarks("PreHos60").Range.Text = "Pre-Hospitalisation More Than 60 Days" & "      " & "RM" & txtPreHos60
        End If
        
        If txtPreGP <> 0 And txtPreGP <> "" Then
            objdoc.Bookmarks("PreHosGPFee").Range.Text = "Pre-Hospitalisation More Than 60 Days" & "      " & "RM" & txtPreGP
        End If
        
        If txtPreMEdic <> 0 And txtPreMEdic <> "" Then
            objdoc.Bookmarks("PreHospMedicine").Range.Text = "Medicine" & "      " & "RM" & txtPreMEdic
        End If
                
        If txtPreMR <> 0 And txtPreMR <> "" Then
            objdoc.Bookmarks("InsuredName").Range.Text = "Medical Report" & "      " & "RM" & txtPreMR
        End If
        
        If txtPreSpec1 <> 0 And txtPreSpec1 <> "" Then
            objdoc.Bookmarks("PreHospSPFees").Range.Text = "Specialist" & "      " & "RM" & txtPreSpec1
        End If
        
        If txtPostGHosp <> 0 And txtPostGHosp <> "" Then
            objdoc.Bookmarks("PostHosp31Days").Range.Text = "Post Hospitalisation More 31 days" & "      " & "RM" & txtPostGHosp
        End If
        
        If txtpostNotSame <> 0 And txtpostNotSame <> "" Then
            objdoc.Bookmarks("NotSamePhy").Range.Text = "Different Physician" & "      " & "RM" & txtpostNotSame
        End If
            
        If txtOthers <> 0 And txtOthers <> "" Then
            objdoc.Bookmarks("Others").Range.Text = Others & "      " & "RM" & txtOthers
        End If
        If txtPayCode <> "" Then
            objdoc.Bookmarks("Payor").Range.Text = PAY!PAYCompanyName
        End If
        
        bookmark = True

    HOS.Close
    Set HOS = Nothing
End Function

Private Sub SaveCASOthers()
Dim sqlstrCASOthers As String
Dim Other As New ADODB.Recordset
Dim sqlstrCASOthersTwo As String
Dim OtherTwo As New ADODB.Recordset
Dim strAppendCase As String
Dim strAppend As String
Dim strsqlCASOthers As String
Dim MEDSURDone1, MEDSURDone2, MEDSURDone3 As String
Dim PROCINVDone1, PROCINVDone2, PROCINVDone3 As String
Dim InvCode1, InvCode2, InvCode3 As String
Dim PROCRequired1, PROCRequired2, PROCRequired3, PROCRequired4, PROCRequired5, PROCRequired6 As String
Dim CPT1, CPT2, CPT3, CPT4, CPT5, CPT6 As String
Dim ActSurgFees1, ActSurgFees2, ActSurgFees3, ActSurgFees4, ActSurgFees5, ActSurgFees6 As Double
Dim ActAnaesFees1, ActAnaesFees2, ActAnaesFees3, ActAnaesFees4, ActAnaesFees5, ActAnaesFees6 As Double
Dim NSPROCRequired1, NSPROCRequired2, NSPROCRequired3, NSPROCRequired4, NSPROCRequired5, NSPROCRequired6 As String
Dim NSCPT1, NSCPT2, NSCPT3, NSCPT4, NSCPT5, NSCPT6 As String
Dim ActNonSurgFees1, ActNonSurgFees2, ActNonSurgFees3, ActNonSurgFees4, ActNonSurgFees5, ActNonSurgFees6 As Double
Dim TotalProc As Integer

Dim AttDr1, AttDr2, AttDr3, AttDr4, AttDr5 As String
Dim AnaesName1, AnaesName2, AnaesName3 As String
        
        TotalProc = 0
    
        'If txtPROCRequired1 <> "" Then
        '    TotalProc = CDbl(TotalProc) + 1
        'End If

        'If txtPROCRequired2 <> "" Then
        '    TotalProc = CDbl(TotalProc) + 1
        'End If

        
        'If txtPROCRequired3 <> "" Then
        '    TotalProc = CDbl(TotalProc) + 1
        'End If
        'If txtPROCRequired4 <> "" Then
        '    TotalProc = CDbl(TotalProc) + 1
        'End If
        
        'If txtPROCRequired5 <> "" Then
        '    TotalProc = CDbl(TotalProc) + 1
        'End If
            
        'If txtPROCRequired6 <> "" Then
        '    TotalProc = CDbl(TotalProc) + 1
        'End If
            
        MEDSURDone1 = ChkStr(Replace(frmMMAEntry.txtMEDSURDone1, "'", "''"))
        MEDSURDone2 = ChkStr(Replace(frmMMAEntry.txtMEDSURDone2, "'", "''"))
        MEDSURDone3 = ChkStr(Replace(frmMMAEntry.txtMEDSURDone3, "'", "''"))
        PROCINVDone1 = ChkStr(Replace(frmMMAEntry.txtPROCINVDone1, "'", "''"))
        PROCINVDone2 = ChkStr(Replace(frmMMAEntry.txtPROCINVDone2, "'", "''"))
        PROCINVDone3 = ChkStr(Replace(frmMMAEntry.txtPROCINVDone3, "'", "''"))
        PROCRequired1 = ChkStr(Replace(frmMMAEntry.txtProcedure1, "'", "''"))
        RROCRequired2 = ChkStr(Replace(frmMMAEntry.txtProcedure2, "'", "''"))
        PROCRequired3 = ChkStr(Replace(frmMMAEntry.txtProcedure3, "'", "''"))
        PROCRequired4 = ChkStr(Replace(frmMMAEntry.txtProcedure4, "'", "''"))
        PROCRequired5 = ChkStr(Replace(frmMMAEntry.txtProcedure5, "'", "''"))
        PROCRequired6 = ChkStr(Replace(frmMMAEntry.txtProcedure6, "'", "''"))
        
        CPT1 = IIf(Len(frmMMAEntry.txtProcCode1), frmMMAEntry.txtProcCode1, "")
        CPT2 = IIf(Len(frmMMAEntry.txtProcCode2), frmMMAEntry.txtProcCode2, "")
        CPT3 = IIf(Len(frmMMAEntry.txtProcCode3), frmMMAEntry.txtProcCode3, "")
        CPT4 = IIf(Len(frmMMAEntry.txtProcCode4), frmMMAEntry.txtProcCode4, "")
        CPT5 = IIf(Len(frmMMAEntry.txtProcCode5), frmMMAEntry.txtProcCode5, "")
        CPT6 = IIf(Len(frmMMAEntry.txtProcCode6), frmMMAEntry.txtProcCode6, "")
        
        ActSurgFees1 = IIf(Len(frmMMAEntry.txtSActFee1), frmMMAEntry.txtSActFee1, 0)
        ActSurgFees2 = IIf(Len(frmMMAEntry.txtSActFee2), frmMMAEntry.txtSActFee2, 0)
        ActSurgFees3 = IIf(Len(frmMMAEntry.txtSActFee3), frmMMAEntry.txtSActFee3, 0)
        ActSurgFees4 = IIf(Len(frmMMAEntry.txtSActFee4), frmMMAEntry.txtSActFee4, 0)
        ActSurgFees5 = IIf(Len(frmMMAEntry.txtSActFee5), frmMMAEntry.txtSActFee5, 0)
        ActSurgFees6 = IIf(Len(frmMMAEntry.txtSActFee6), frmMMAEntry.txtSActFee6, 0)
        
        ActAnaesFees1 = IIf(Len(frmMMAEntry.txtAnaesActFees1), frmMMAEntry.txtAnaesActFees1, 0)
        ActAnaesFees2 = IIf(Len(frmMMAEntry.txtAnaesActFees2), frmMMAEntry.txtAnaesActFees2, 0)
        ActAnaesFees3 = IIf(Len(frmMMAEntry.txtAnaesActFees3), frmMMAEntry.txtAnaesActFees3, 0)
        ActAnaesFees4 = IIf(Len(frmMMAEntry.txtAnaesActFees4), frmMMAEntry.txtAnaesActFees4, 0)
        ActAnaesFees5 = IIf(Len(frmMMAEntry.txtAnaesActFees5), frmMMAEntry.txtAnaesActFees5, 0)
        ActAnaesFees6 = IIf(Len(frmMMAEntry.txtAnaesActFees6), frmMMAEntry.txtAnaesActFees6, 0)
        
        InvCode1 = IIf(Len(frmMMAEntry.txtINVCode1), frmMMAEntry.txtINVCode1, "")
        InvCode2 = IIf(Len(frmMMAEntry.txtINVCode2), frmMMAEntry.txtINVCode2, "")
        InvCode3 = IIf(Len(frmMMAEntry.txtINVCode3), frmMMAEntry.txtINVCode3, "")
        
        NSPROCRequired1 = ChkStr(Replace(frmMMAEntry.txtNSProcedure1, "'", "''"))
        NSPROCRequired2 = ChkStr(Replace(frmMMAEntry.txtNSProcedure2, "'", "''"))
        NSPROCRequired3 = ChkStr(Replace(frmMMAEntry.txtNSProcedure3, "'", "''"))
        NSPROCRequired4 = ChkStr(Replace(frmMMAEntry.txtNSProcedure4, "'", "''"))
        NSPROCRequired5 = ChkStr(Replace(frmMMAEntry.txtNSProcedure5, "'", "''"))
        NSPROCRequired6 = ChkStr(Replace(frmMMAEntry.txtNSProcedure6, "'", "''"))
        
        NSCPT1 = IIf(Len(frmMMAEntry.txtNSProcCode1), frmMMAEntry.txtNSProcCode1, "")
        NSCPT2 = IIf(Len(frmMMAEntry.txtNSProcCode2), frmMMAEntry.txtNSProcCode2, "")
        NSCPT3 = IIf(Len(frmMMAEntry.txtNSProcCode3), frmMMAEntry.txtNSProcCode3, "")
        NSCPT4 = IIf(Len(frmMMAEntry.txtNSProcCode4), frmMMAEntry.txtNSProcCode4, "")
        NSCPT5 = IIf(Len(frmMMAEntry.txtNSProcCode5), frmMMAEntry.txtNSProcCode5, "")
        NSCPT6 = IIf(Len(frmMMAEntry.txtNSProcCode6), frmMMAEntry.txtNSProcCode6, "")
        
        ActNonSurgFees1 = IIf(Len(frmMMAEntry.txtNSActFee1), frmMMAEntry.txtNSActFee1, 0)
        ActNonSurgFees2 = IIf(Len(frmMMAEntry.txtNSActFee2), frmMMAEntry.txtNSActFee2, 0)
        ActNonSurgFees3 = IIf(Len(frmMMAEntry.txtNSActFee3), frmMMAEntry.txtNSActFee3, 0)
        ActNonSurgFees4 = IIf(Len(frmMMAEntry.txtNSActFee4), frmMMAEntry.txtNSActFee4, 0)
        ActNonSurgFees5 = IIf(Len(frmMMAEntry.txtNSActFee5), frmMMAEntry.txtNSActFee5, 0)
        ActNonSurgFees6 = IIf(Len(frmMMAEntry.txtNSActFee6), frmMMAEntry.txtNSActFee6, 0)
        'ER = Replace(txtATTDoc(0), "'", "''")
        AttDr1 = IIf(Len(frmMMAEntry.txtATTDoc(0)), Replace(frmMMAEntry.txtATTDoc(0), "'", "''"), "")
        AttDr2 = IIf(Len(frmMMAEntry.txtATTDoc(1)), Replace(frmMMAEntry.txtATTDoc(1), "'", "''"), "")
        AttDr3 = IIf(Len(frmMMAEntry.txtATTDoc(2)), Replace(frmMMAEntry.txtATTDoc(2), "'", "''"), "")
        AttDr4 = IIf(Len(frmMMAEntry.txtATTDoc(3)), Replace(frmMMAEntry.txtATTDoc(3), "'", "''"), "")
        AttDr5 = IIf(Len(frmMMAEntry.txtATTDoc(4)), Replace(frmMMAEntry.txtATTDoc(4), "'", "''"), "")
        AnaesName1 = IIf(Len(frmMMAEntry.txtAnaesDoc(0)), Replace(frmMMAEntry.txtAnaesDoc(0), "'", "''"), "")
        AnaesName2 = IIf(Len(frmMMAEntry.txtAnaesDoc(1)), Replace(frmMMAEntry.txtAnaesDoc(1), "'", "''"), "")
        AnaesName3 = IIf(Len(frmMMAEntry.txtAnaesDoc(2)), Replace(frmMMAEntry.txtAnaesDoc(2), "'", "''"), "")
        
        strsqlCASOthers = "'" & ChkStr(txtCaseNo) & "', " & _
                        "'" & MEDSURDone1 & "', '" & MEDSURDone2 & "', '" & MEDSURDone3 & "', " & _
                        "'" & PROCINVDone1 & "', '" & PROCINVDone2 & "', '" & PROCINVDone3 & "', " & _
                        "'" & InvCode1 & "', '" & InvCode2 & "', '" & InvCode3 & "', " & _
                        "'" & PROCRequired1 & "', '" & RROCRequired2 & "', '" & PROCRequired3 & "', " & _
                        "'" & PROCRequired4 & "', '" & PROCRequired5 & "','" & PROCRequired6 & "', " & _
                        "'" & CPT1 & "','" & CPT2 & "', '" & CPT3 & "', '" & CPT4 & "', '" & CPT5 & "' ,'" & CPT6 & "' , " & _
                        "" & ActSurgFees1 & ", " & ActSurgFees2 & ", " & ActSurgFees3 & ", " & ActSurgFees4 & ", " & ActSurgFees5 & "," & ActSurgFees6 & ", " & _
                        "" & ActAnaesFees1 & ", " & ActAnaesFees2 & ", " & ActAnaesFees3 & ", " & ActAnaesFees4 & ", " & ActAnaesFees5 & ", " & ActAnaesFees6 & "," & _
                        "'" & NSPROCRequired1 & "','" & NSPROCRequired2 & "', '" & NSPROCRequired3 & "', '" & NSPROCRequired4 & "', '" & NSPROCRequired5 & "' ,'" & NSPROCRequired6 & "' , " & _
                        "'" & NSCPT1 & "','" & NSCPT2 & "', '" & NSCPT3 & "', '" & NSCPT4 & "', '" & NSCPT5 & "','" & NSCPT6 & "', " & _
                        "" & ActNonSurgFees1 & ", " & ActNonSurgFees2 & ", " & ActNonSurgFees3 & ", " & ActNonSurgFees4 & ", " & ActNonSurgFees5 & ",  " & ActNonSurgFees6 & ""
        strsqlCASOthers = "EXEC Case_AdjInvCASOthers " + strsqlCASOthers

        medixmasterconn.Execute strsqlCASOthers
        
        sqlstrCASOthersTwo = "'" & ChkStr(txtCaseNo) & "', " & _
                        "'" & AttDr1 & "', '" & AttDr2 & "', '" & AttDr3 & "', " & _
                        "'" & AttDr4 & "', '" & AttDr5 & "', '" & AnaesName1 & "', " & _
                        "'" & AnaesName2 & "', '" & AnaesName3 & "', " & TotalProc & ""
                            
        sqlstrCASOthersTwo = "EXEC Case_AdjInvCASOthersTwo " + sqlstrCASOthersTwo

        medixmasterconn.Execute sqlstrCASOthersTwo
End Sub

Private Sub SearchMMARate(CPTCode As String)
Dim strsql As String
Dim MMA As New ADODB.Recordset
    MMASurgFees = 0
    MMAAnaeFees = 0
    MMASurgFees13th = 0
    MMAAnaeFees13th = 0
    MMASurgFees13thMin = 0
    
    strsqlMMA = "Select SurgeonsFees, AnaesthetistsFees,Surgeon13th,Anesthetist13th,Surgeon13thMin  from VIEW_MMA_Schedule where  OPCSCode = '" & Trim(CPTCode) & "'"
    MMA.Open strsqlMMA, medixmasterconnMaint, adOpenKeyset, adLockOptimistic
        Do While Not MMA.EOF
            MMASurgFees = IIf(Len(MMA!SurgeonsFees), MMA!SurgeonsFees, 0)
            MMAAnaeFees = IIf(Len(MMA!AnaesthetistsFees), MMA!AnaesthetistsFees, 0)
            MMASurgFees13th = IIf(Len(MMA!Surgeon13th), MMA!Surgeon13th, 0)
            MMAAnaeFees13th = IIf(Len(MMA!Anesthetist13th), MMA!Anesthetist13th, 0)
            MMASurgFees13thMin = IIf(Len(MMA!Surgeon13thMin), MMA!Surgeon13thMin, 0)
        MMA.MoveNext
        Loop
    MMA.Close
    Set MMA = Nothing
End Sub

Private Sub ClearHSS()
    frmHSSDetails.txtInvPharmacy = ""
    frmHSSDetails.txtMedSupp = ""
    frmHSSDetails.txtNursingCare = ""
    frmHSSDetails.txtNursingServ = ""
    frmHSSDetails.txtLab = ""
    frmHSSDetails.txtRadiology = ""
    frmHSSDetails.txtOTSupp = ""
    frmHSSDetails.txtEquipChrg = ""
    frmHSSDetails.txtMisc = ""
    frmHSSDetails.txtHSS = ""
End Sub

Private Sub LGListing()
    Dim LG As New ADODB.Recordset
    Dim LGList As ListItem
    Dim sql As String
    Dim cmd As New ADODB.Command
    Dim Prm1 As ADODB.Parameter
    Dim Prm2 As ADODB.Parameter
    Dim strAppendCase As String
    Dim strAppend As String
    
        lstLGListing.ListItems.Clear
        
        'sql = " Select CASeNo, LGNo, LGDate, " & _
            " LGAmt, LGIssuedBy, LGRemark  From CASLG " & _
            " where CASENo ='" & Trim(txtCaseNo) & "'"
    
        'LG.Open sql, medixmasterconnMaint, adOpenKeyset, adLockOptimistic
        strAppendCase = "7"
        strAppend = " AND CASENo ='" & Trim(txtCaseNo) & "'"
        Set cmd.ActiveConnection = medixmasterconnMaint
        cmd.CommandText = "Case_Adjustment"
        cmd.CommandType = adCmdStoredProc
        cmd.CommandTimeout = 300
        Set Prm1 = cmd.CreateParameter("StrAppend", adVarChar, adParamInput, 2000, strAppend)
        cmd.Parameters.Append Prm1
        Set Prm2 = cmd.CreateParameter("StrAppendCase", adVarChar, adParamInput, 2000, strAppendCase)
        cmd.Parameters.Append Prm2
        LG.CursorLocation = adUseClient
        LG.CursorType = adOpenForwardOnly
        LG.CacheSize = 100
        Set LG = cmd.Execute
        
        If Not LG.EOF Then
            Do Until LG.EOF
                'If LG!LGStatus = "B" Then
                '    Set LGList = lstLGListing.ListItems.Add(, , Trim("B" & LG!CaseNo & "-" & LG!LGNo & Mid(LG!LGPurpose, 1, 1)))
                'Else
                    Set LGList = lstLGListing.ListItems.Add(, , "")
                'End If
                With LG
                    
                    If LG!LGStatus = "B" Then
                        LGList.SubItems(1) = Trim("B" & LG!CaseNo & "-" & LG!LGNo & Mid(LG!LGPurpose, 1, 1))
                    Else
                        LGList.SubItems(1) = Trim("L" & LG!CaseNo & "-" & LG!LGNo & Mid(LG!LGPurpose, 1, 1))
                    End If
                    
                    If Len(!LGDate) Then
                        LGList.SubItems(2) = Format(!LGDate, "dd/mm/yyyy")
                    End If
                    If Len(!LGAmt) Then
                        LGList.SubItems(3) = Format(!LGAmt, "#,##0.00")
                    End If
                    If Len(cboPatientList) Then
                        LGList.SubItems(4) = Trim(txtPatientName)
                    End If
                    If Len(PatientName) Then
                        LGList.SubItems(5) = Trim(PatientName)
                    End If
                    If Len(txtDateAdmitted) Then
                       LGList.SubItems(6) = Format(txtDateAdmitted, "dd/mm/yyyy")
                    End If
                    If Len(txtHOSCode) Then
                        LGList.SubItems(7) = Trim(txtHOSCode)
                    End If
                    If Len(!LGIssuedBy) Then
                        LGList.SubItems(8) = Trim(!LGIssuedBy)
                    End If
                    If Len(!LGRemark) Then
                        LGList.SubItems(9) = Trim(!LGRemark)
                    End If
                    If Len(!CaseNo) Then
                        LGList.SubItems(10) = Trim(!CaseNo & "-" & !LGNo)
                    End If
                    If Len(!LGUserTime) Then
                        LGList.SubItems(11) = Trim(!LGUserTime)
                    End If
                    If Len(!LGSystemTime) Then
                        LGList.SubItems(12) = Trim(!LGSystemTime)
                    End If
                    If Len(!LGPurpose) Then
                        LGList.SubItems(13) = Trim(!LGPurpose)
                    End If
                    
                    
                    .MoveNext
                End With
                Loop
                LG.Close
                Set LG = Nothing
            End If
End Sub

