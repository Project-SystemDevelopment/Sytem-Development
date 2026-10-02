VERSION 5.00
Begin VB.Form frmMMAEntryGST 
   Caption         =   "Procedures/OPCS"
   ClientHeight    =   7020
   ClientLeft      =   60
   ClientTop       =   510
   ClientWidth     =   11310
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   ScaleHeight     =   7020
   ScaleWidth      =   11310
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame4 
      Height          =   3375
      Left            =   120
      TabIndex        =   32
      Top             =   3000
      Width           =   11175
      Begin VB.Frame FrameS14Rate 
         Enabled         =   0   'False
         Height          =   2535
         Left            =   5160
         TabIndex        =   214
         Top             =   360
         Width           =   855
         Begin VB.TextBox txtS14Rate8 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   120
            MaxLength       =   8
            TabIndex        =   235
            Top             =   2160
            Width           =   615
         End
         Begin VB.TextBox txtS14Rate7 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   120
            MaxLength       =   8
            TabIndex        =   234
            Top             =   1920
            Width           =   615
         End
         Begin VB.TextBox txtS14Rate6 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   120
            MaxLength       =   8
            TabIndex        =   221
            Top             =   1680
            Width           =   615
         End
         Begin VB.TextBox txtS14Rate5 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   120
            MaxLength       =   8
            TabIndex        =   220
            Top             =   1440
            Width           =   615
         End
         Begin VB.TextBox txtS14Rate4 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   120
            MaxLength       =   8
            TabIndex        =   219
            Top             =   1200
            Width           =   615
         End
         Begin VB.TextBox txtS14Rate3 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   120
            MaxLength       =   8
            TabIndex        =   218
            Top             =   960
            Width           =   615
         End
         Begin VB.TextBox txtS14Rate2 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   120
            MaxLength       =   8
            TabIndex        =   217
            Top             =   720
            Width           =   615
         End
         Begin VB.TextBox Text27 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   120
            TabIndex        =   216
            Top             =   2880
            Width           =   615
         End
         Begin VB.TextBox txtS14Rate1 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   120
            MaxLength       =   8
            TabIndex        =   215
            Top             =   480
            Width           =   615
         End
         Begin VB.Label Label90 
            Alignment       =   2  'Center
            Caption         =   "13th New"
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
            Index           =   16
            Left            =   80
            TabIndex        =   222
            Top             =   110
            Width           =   735
         End
         Begin VB.Line Line1 
            Index           =   7
            X1              =   75
            X2              =   720
            Y1              =   2835
            Y2              =   2835
         End
      End
      Begin VB.Frame FrameSActFee 
         Height          =   2535
         Left            =   4320
         TabIndex        =   166
         Top             =   360
         Width           =   855
         Begin VB.TextBox txtSActFee8 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   120
            MaxLength       =   8
            TabIndex        =   237
            Top             =   2160
            Width           =   615
         End
         Begin VB.TextBox txtSActFee7 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   120
            MaxLength       =   8
            TabIndex        =   236
            Top             =   1920
            Width           =   615
         End
         Begin VB.TextBox txtSActFee1 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   120
            MaxLength       =   8
            TabIndex        =   173
            Top             =   480
            Width           =   615
         End
         Begin VB.TextBox txtTotalSActFee 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            Enabled         =   0   'False
            Height          =   285
            Left            =   120
            TabIndex        =   172
            Top             =   3000
            Width           =   615
         End
         Begin VB.TextBox txtSActFee2 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   120
            MaxLength       =   8
            TabIndex        =   171
            Top             =   720
            Width           =   615
         End
         Begin VB.TextBox txtSActFee3 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   120
            MaxLength       =   8
            TabIndex        =   170
            Top             =   960
            Width           =   615
         End
         Begin VB.TextBox txtSActFee4 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   120
            MaxLength       =   8
            TabIndex        =   169
            Top             =   1200
            Width           =   615
         End
         Begin VB.TextBox txtSActFee5 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   120
            MaxLength       =   8
            TabIndex        =   168
            Top             =   1440
            Width           =   615
         End
         Begin VB.TextBox txtSActFee6 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   120
            MaxLength       =   8
            TabIndex        =   167
            Top             =   1680
            Width           =   615
         End
         Begin VB.Line Line2 
            X1              =   120
            X2              =   720
            Y1              =   2955
            Y2              =   2955
         End
         Begin VB.Label Label90 
            Alignment       =   2  'Center
            Caption         =   "Actual "
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
            TabIndex        =   174
            Top             =   120
            Width           =   615
         End
      End
      Begin VB.Frame Frame14Rate 
         Enabled         =   0   'False
         Height          =   2535
         Left            =   8640
         TabIndex        =   223
         Top             =   360
         Width           =   855
         Begin VB.TextBox txtAnae14Rate8 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   120
            MaxLength       =   8
            TabIndex        =   245
            Top             =   2160
            Width           =   615
         End
         Begin VB.TextBox txtAnae14Rate7 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   120
            MaxLength       =   8
            TabIndex        =   244
            Top             =   1920
            Width           =   615
         End
         Begin VB.TextBox txtAnae14Rate6 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   120
            MaxLength       =   8
            TabIndex        =   230
            Top             =   1680
            Width           =   615
         End
         Begin VB.TextBox txtAnae14Rate5 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   120
            MaxLength       =   8
            TabIndex        =   229
            Top             =   1440
            Width           =   615
         End
         Begin VB.TextBox txtAnae14Rate4 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   120
            MaxLength       =   8
            TabIndex        =   228
            Top             =   1200
            Width           =   615
         End
         Begin VB.TextBox txtAnae14Rate3 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   120
            MaxLength       =   8
            TabIndex        =   227
            Top             =   960
            Width           =   615
         End
         Begin VB.TextBox txtAnae14Rate2 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   120
            MaxLength       =   8
            TabIndex        =   226
            Top             =   720
            Width           =   615
         End
         Begin VB.TextBox txtAnae14Rate1 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   120
            MaxLength       =   8
            TabIndex        =   225
            Top             =   480
            Width           =   615
         End
         Begin VB.TextBox Text7 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   120
            TabIndex        =   224
            Top             =   2640
            Width           =   615
         End
         Begin VB.Label Label90 
            Alignment       =   2  'Center
            Caption         =   "13th New"
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
            Index           =   13
            Left            =   80
            TabIndex        =   231
            Top             =   120
            Width           =   735
         End
         Begin VB.Line Line3 
            Index           =   2
            X1              =   75
            X2              =   720
            Y1              =   2760
            Y2              =   2760
         End
      End
      Begin VB.Frame FrameNS14Rate 
         Enabled         =   0   'False
         Height          =   2535
         Left            =   5160
         TabIndex        =   205
         Top             =   360
         Width           =   855
         Begin VB.TextBox txtNS14Rate8 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   120
            MaxLength       =   8
            TabIndex        =   257
            Top             =   2160
            Width           =   615
         End
         Begin VB.TextBox txtNS14Rate7 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   120
            MaxLength       =   8
            TabIndex        =   256
            Top             =   1920
            Width           =   615
         End
         Begin VB.TextBox txtNS14Rate1 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   120
            MaxLength       =   8
            TabIndex        =   212
            Top             =   480
            Width           =   615
         End
         Begin VB.TextBox Text30 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   120
            TabIndex        =   211
            Top             =   2760
            Width           =   615
         End
         Begin VB.TextBox txtNS14Rate2 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   120
            MaxLength       =   8
            TabIndex        =   210
            Top             =   720
            Width           =   615
         End
         Begin VB.TextBox txtNS14Rate3 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   120
            MaxLength       =   8
            TabIndex        =   209
            Top             =   960
            Width           =   615
         End
         Begin VB.TextBox txtNS14Rate4 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   120
            MaxLength       =   8
            TabIndex        =   208
            Top             =   1200
            Width           =   615
         End
         Begin VB.TextBox txtNS14Rate5 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   120
            MaxLength       =   8
            TabIndex        =   207
            Top             =   1440
            Width           =   615
         End
         Begin VB.TextBox txtNS14Rate6 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   120
            MaxLength       =   8
            TabIndex        =   206
            Top             =   1680
            Width           =   615
         End
         Begin VB.Line Line1 
            Index           =   8
            X1              =   75
            X2              =   720
            Y1              =   2715
            Y2              =   2715
         End
         Begin VB.Label Label90 
            Alignment       =   2  'Center
            Caption         =   "13th New"
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
            Index           =   17
            Left            =   80
            TabIndex        =   213
            Top             =   110
            Width           =   735
         End
      End
      Begin VB.Frame FrameS14MinRate 
         Enabled         =   0   'False
         Height          =   2535
         Left            =   6000
         TabIndex        =   196
         Top             =   360
         Width           =   855
         Begin VB.TextBox txtS14MinRate8 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   120
            MaxLength       =   8
            TabIndex        =   239
            Top             =   2160
            Width           =   615
         End
         Begin VB.TextBox txtS14MinRate7 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   120
            MaxLength       =   8
            TabIndex        =   238
            Top             =   1920
            Width           =   615
         End
         Begin VB.TextBox txtS14MinRate6 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   120
            MaxLength       =   8
            TabIndex        =   203
            Top             =   1680
            Width           =   615
         End
         Begin VB.TextBox txtS14MinRate5 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   120
            MaxLength       =   8
            TabIndex        =   202
            Top             =   1440
            Width           =   615
         End
         Begin VB.TextBox txtS14MinRate4 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   120
            MaxLength       =   8
            TabIndex        =   201
            Top             =   1200
            Width           =   615
         End
         Begin VB.TextBox txtS14MinRate3 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   120
            MaxLength       =   8
            TabIndex        =   200
            Top             =   960
            Width           =   615
         End
         Begin VB.TextBox txtS14MinRate2 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   120
            MaxLength       =   8
            TabIndex        =   199
            Top             =   720
            Width           =   615
         End
         Begin VB.TextBox Text13 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   120
            TabIndex        =   198
            Top             =   2760
            Width           =   615
         End
         Begin VB.TextBox txtS14MinRate1 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   120
            MaxLength       =   8
            TabIndex        =   197
            Top             =   480
            Width           =   615
         End
         Begin VB.Label Label90 
            Alignment       =   2  'Center
            Caption         =   "13th New (GST)"
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
            Index           =   14
            Left            =   75
            TabIndex        =   204
            Top             =   110
            Width           =   735
         End
         Begin VB.Line Line1 
            Index           =   5
            X1              =   75
            X2              =   720
            Y1              =   2715
            Y2              =   2715
         End
      End
      Begin VB.Frame FrameNS14RateMin 
         Enabled         =   0   'False
         Height          =   2535
         Left            =   6000
         TabIndex        =   187
         Top             =   360
         Width           =   855
         Begin VB.TextBox txtNS14MinRate8 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   120
            MaxLength       =   8
            TabIndex        =   255
            Top             =   2160
            Width           =   615
         End
         Begin VB.TextBox txtNS14MinRate7 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   120
            MaxLength       =   8
            TabIndex        =   254
            Top             =   1920
            Width           =   615
         End
         Begin VB.TextBox txtNS14MinRate6 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   120
            MaxLength       =   8
            TabIndex        =   194
            Top             =   1680
            Width           =   615
         End
         Begin VB.TextBox txtNS14MinRate5 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   120
            MaxLength       =   8
            TabIndex        =   193
            Top             =   1440
            Width           =   615
         End
         Begin VB.TextBox txtNS14MinRate4 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   120
            MaxLength       =   8
            TabIndex        =   192
            Top             =   1200
            Width           =   615
         End
         Begin VB.TextBox txtNS14MinRate3 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   120
            MaxLength       =   8
            TabIndex        =   191
            Top             =   960
            Width           =   615
         End
         Begin VB.TextBox txtNS14MinRate2 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   120
            MaxLength       =   8
            TabIndex        =   190
            Top             =   720
            Width           =   615
         End
         Begin VB.TextBox Text20 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   240
            TabIndex        =   189
            Top             =   2880
            Width           =   615
         End
         Begin VB.TextBox txtNS14MinRate1 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   120
            MaxLength       =   8
            TabIndex        =   188
            Top             =   480
            Width           =   615
         End
         Begin VB.Label Label90 
            Alignment       =   2  'Center
            Caption         =   "13th Sch (GST)"
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
            Index           =   15
            Left            =   75
            TabIndex        =   195
            Top             =   110
            Width           =   735
         End
         Begin VB.Line Line1 
            Index           =   6
            X1              =   195
            X2              =   840
            Y1              =   2835
            Y2              =   2835
         End
      End
      Begin VB.ComboBox cboHospOT 
         Height          =   315
         Left            =   1080
         TabIndex        =   186
         Top             =   2955
         Width           =   1815
      End
      Begin VB.Frame FrameS13Rate 
         Enabled         =   0   'False
         Height          =   2535
         Left            =   6840
         TabIndex        =   157
         Top             =   360
         Width           =   855
         Begin VB.TextBox txtS13Rate8 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   120
            MaxLength       =   8
            TabIndex        =   241
            Top             =   2160
            Width           =   615
         End
         Begin VB.TextBox txtS13Rate7 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   120
            MaxLength       =   8
            TabIndex        =   240
            Top             =   1920
            Width           =   615
         End
         Begin VB.TextBox txtS13Rate6 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   120
            MaxLength       =   8
            TabIndex        =   164
            Top             =   1680
            Width           =   615
         End
         Begin VB.TextBox txtS13Rate5 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   120
            MaxLength       =   8
            TabIndex        =   163
            Top             =   1440
            Width           =   615
         End
         Begin VB.TextBox txtS13Rate4 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   120
            MaxLength       =   8
            TabIndex        =   162
            Top             =   1200
            Width           =   615
         End
         Begin VB.TextBox txtS13Rate3 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   120
            MaxLength       =   8
            TabIndex        =   161
            Top             =   960
            Width           =   615
         End
         Begin VB.TextBox txtS13Rate2 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   120
            MaxLength       =   8
            TabIndex        =   160
            Top             =   720
            Width           =   615
         End
         Begin VB.TextBox txtTotalS13Rate 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   120
            TabIndex        =   159
            Top             =   3000
            Width           =   615
         End
         Begin VB.TextBox txtS13Rate1 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   120
            MaxLength       =   8
            TabIndex        =   158
            Top             =   480
            Width           =   615
         End
         Begin VB.Label Label90 
            Alignment       =   2  'Center
            Caption         =   "13th Old"
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
            Index           =   8
            Left            =   80
            TabIndex        =   165
            Top             =   105
            Width           =   735
         End
         Begin VB.Line Line1 
            Index           =   1
            X1              =   75
            X2              =   720
            Y1              =   2955
            Y2              =   2955
         End
      End
      Begin VB.Frame FrameAnaesAct 
         Height          =   2535
         Left            =   7800
         TabIndex        =   148
         Top             =   360
         Width           =   855
         Begin VB.TextBox txtAnaesActFees8 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   120
            MaxLength       =   8
            TabIndex        =   243
            Top             =   2160
            Width           =   615
         End
         Begin VB.TextBox txtAnaesActFees7 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   120
            MaxLength       =   8
            TabIndex        =   242
            Top             =   1920
            Width           =   615
         End
         Begin VB.TextBox txtAnaesActFees6 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   120
            MaxLength       =   8
            TabIndex        =   155
            Top             =   1680
            Width           =   615
         End
         Begin VB.TextBox txtAnaesActFees5 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   120
            MaxLength       =   8
            TabIndex        =   154
            Top             =   1440
            Width           =   615
         End
         Begin VB.TextBox txtAnaesActFees4 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   120
            MaxLength       =   8
            TabIndex        =   153
            Top             =   1200
            Width           =   615
         End
         Begin VB.TextBox txtAnaesActFees3 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   120
            MaxLength       =   8
            TabIndex        =   152
            Top             =   960
            Width           =   615
         End
         Begin VB.TextBox txtAnaesActFees2 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   120
            MaxLength       =   8
            TabIndex        =   151
            Top             =   720
            Width           =   615
         End
         Begin VB.TextBox txtTotalAnaesActFees 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            Enabled         =   0   'False
            Height          =   285
            Left            =   120
            TabIndex        =   150
            Top             =   2880
            Width           =   615
         End
         Begin VB.TextBox txtAnaesActFees1 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   120
            MaxLength       =   8
            TabIndex        =   149
            Top             =   480
            Width           =   615
         End
         Begin VB.Label Label90 
            Alignment       =   2  'Center
            Caption         =   "Actual"
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
            Left            =   120
            TabIndex        =   156
            Top             =   120
            Width           =   615
         End
         Begin VB.Line Line4 
            X1              =   120
            X2              =   765
            Y1              =   2880
            Y2              =   2880
         End
      End
      Begin VB.Frame FrameNSActFee 
         Height          =   2535
         Left            =   4320
         TabIndex        =   139
         Top             =   360
         Width           =   855
         Begin VB.TextBox txtNSActFee8 
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   120
            MaxLength       =   8
            TabIndex        =   259
            Top             =   2160
            Width           =   615
         End
         Begin VB.TextBox txtNSActFee7 
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   120
            MaxLength       =   8
            TabIndex        =   258
            Top             =   1920
            Width           =   615
         End
         Begin VB.TextBox txtTotalNSActFee 
            Appearance      =   0  'Flat
            Enabled         =   0   'False
            Height          =   285
            Left            =   120
            TabIndex        =   146
            Top             =   2760
            Width           =   615
         End
         Begin VB.TextBox txtNSActFee1 
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   120
            MaxLength       =   8
            TabIndex        =   145
            Top             =   480
            Width           =   615
         End
         Begin VB.TextBox txtNSActFee2 
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   120
            MaxLength       =   8
            TabIndex        =   144
            Top             =   720
            Width           =   615
         End
         Begin VB.TextBox txtNSActFee3 
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   120
            MaxLength       =   8
            TabIndex        =   143
            Top             =   960
            Width           =   615
         End
         Begin VB.TextBox txtNSActFee4 
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   120
            MaxLength       =   8
            TabIndex        =   142
            Top             =   1200
            Width           =   615
         End
         Begin VB.TextBox txtNSActFee5 
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   120
            MaxLength       =   8
            TabIndex        =   141
            Top             =   1440
            Width           =   615
         End
         Begin VB.TextBox txtNSActFee6 
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   120
            MaxLength       =   8
            TabIndex        =   140
            Top             =   1700
            Width           =   615
         End
         Begin VB.Line Line6 
            X1              =   75
            X2              =   720
            Y1              =   2715
            Y2              =   2715
         End
         Begin VB.Label Label90 
            Alignment       =   2  'Center
            Caption         =   "Actual"
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
            Left            =   120
            TabIndex        =   147
            Top             =   120
            Width           =   615
         End
      End
      Begin VB.Frame FrameNS13Rate 
         Enabled         =   0   'False
         Height          =   2535
         Left            =   6840
         TabIndex        =   130
         Top             =   360
         Width           =   855
         Begin VB.TextBox txtNS13Rate8 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   120
            MaxLength       =   8
            TabIndex        =   253
            Top             =   2160
            Width           =   615
         End
         Begin VB.TextBox txtNS13Rate7 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   120
            MaxLength       =   8
            TabIndex        =   252
            Top             =   1920
            Width           =   615
         End
         Begin VB.TextBox txtNS13Rate1 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   120
            MaxLength       =   8
            TabIndex        =   137
            Top             =   480
            Width           =   615
         End
         Begin VB.TextBox txtTotalNS13Rate 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   120
            TabIndex        =   136
            Top             =   3000
            Width           =   615
         End
         Begin VB.TextBox txtNS13Rate2 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   120
            MaxLength       =   8
            TabIndex        =   135
            Top             =   720
            Width           =   615
         End
         Begin VB.TextBox txtNS13Rate3 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   120
            MaxLength       =   8
            TabIndex        =   134
            Top             =   960
            Width           =   615
         End
         Begin VB.TextBox txtNS13Rate4 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   120
            MaxLength       =   8
            TabIndex        =   133
            Top             =   1200
            Width           =   615
         End
         Begin VB.TextBox txtNS13Rate5 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   120
            MaxLength       =   8
            TabIndex        =   132
            Top             =   1440
            Width           =   615
         End
         Begin VB.TextBox txtNS13Rate6 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   120
            MaxLength       =   8
            TabIndex        =   131
            Top             =   1680
            Width           =   615
         End
         Begin VB.Line Line1 
            Index           =   2
            X1              =   75
            X2              =   720
            Y1              =   2955
            Y2              =   2955
         End
         Begin VB.Label Label90 
            Alignment       =   2  'Center
            Caption         =   "13th Sch"
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
            Index           =   10
            Left            =   80
            TabIndex        =   138
            Top             =   110
            Width           =   735
         End
      End
      Begin VB.Frame FrameNSMMARate 
         Enabled         =   0   'False
         Height          =   2415
         Left            =   5040
         TabIndex        =   121
         Top             =   3360
         Width           =   855
         Begin VB.TextBox txtTotalNSMMARate 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   120
            TabIndex        =   128
            Top             =   2040
            Width           =   615
         End
         Begin VB.TextBox txtNSMMARate1 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   120
            MaxLength       =   8
            TabIndex        =   127
            Top             =   480
            Width           =   615
         End
         Begin VB.TextBox txtNSMMARate2 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   120
            MaxLength       =   8
            TabIndex        =   126
            Top             =   720
            Width           =   615
         End
         Begin VB.TextBox txtNSMMARate3 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   120
            MaxLength       =   8
            TabIndex        =   125
            Top             =   960
            Width           =   615
         End
         Begin VB.TextBox txtNSMMARate4 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   120
            MaxLength       =   8
            TabIndex        =   124
            Top             =   1200
            Width           =   615
         End
         Begin VB.TextBox txtNSMMARate5 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   120
            MaxLength       =   8
            TabIndex        =   123
            Top             =   1440
            Width           =   615
         End
         Begin VB.TextBox txtNSMMARate6 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   120
            MaxLength       =   8
            TabIndex        =   122
            Top             =   1700
            Width           =   615
         End
         Begin VB.Line Line5 
            X1              =   75
            X2              =   720
            Y1              =   1995
            Y2              =   1995
         End
         Begin VB.Label Label90 
            Alignment       =   2  'Center
            Caption         =   "MMA"
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
            Index           =   7
            Left            =   120
            TabIndex        =   129
            Top             =   120
            Width           =   615
         End
      End
      Begin VB.Frame FrameSProc 
         Caption         =   "Procedure"
         Height          =   2535
         Left            =   120
         TabIndex        =   114
         Top             =   360
         Width           =   2775
         Begin VB.TextBox txtProcedure8 
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   120
            MaxLength       =   200
            TabIndex        =   251
            Top             =   2160
            Width           =   2535
         End
         Begin VB.TextBox txtProcedure7 
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   120
            MaxLength       =   200
            TabIndex        =   250
            Top             =   1920
            Width           =   2535
         End
         Begin VB.TextBox txtProcedure1 
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   120
            MaxLength       =   200
            TabIndex        =   120
            Top             =   480
            Width           =   2535
         End
         Begin VB.TextBox txtProcedure2 
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   120
            MaxLength       =   200
            TabIndex        =   119
            Top             =   720
            Width           =   2535
         End
         Begin VB.TextBox txtProcedure3 
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   120
            MaxLength       =   200
            TabIndex        =   118
            Top             =   960
            Width           =   2535
         End
         Begin VB.TextBox txtProcedure4 
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   120
            MaxLength       =   200
            TabIndex        =   117
            Top             =   1200
            Width           =   2535
         End
         Begin VB.TextBox txtProcedure5 
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   120
            MaxLength       =   200
            TabIndex        =   116
            Top             =   1440
            Width           =   2535
         End
         Begin VB.TextBox txtProcedure6 
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   120
            MaxLength       =   200
            TabIndex        =   115
            Top             =   1680
            Width           =   2535
         End
      End
      Begin VB.Frame FrameSOPCS 
         Caption         =   "OPCS"
         Height          =   2535
         Left            =   2880
         TabIndex        =   100
         Top             =   360
         Width           =   1335
         Begin VB.CommandButton Command2 
            Caption         =   ">"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   270
            Left            =   1020
            TabIndex        =   249
            Top             =   2190
            Width           =   255
         End
         Begin VB.TextBox txtProcCode8 
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   120
            MaxLength       =   6
            TabIndex        =   248
            Top             =   2160
            Width           =   855
         End
         Begin VB.CommandButton Command1 
            Caption         =   ">"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   270
            Left            =   1020
            TabIndex        =   247
            Top             =   1950
            Width           =   255
         End
         Begin VB.TextBox txtProcCode7 
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   120
            MaxLength       =   6
            TabIndex        =   246
            Top             =   1920
            Width           =   855
         End
         Begin VB.TextBox txtProcCode1 
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   120
            MaxLength       =   6
            TabIndex        =   113
            Top             =   480
            Width           =   855
         End
         Begin VB.TextBox txtProcCode2 
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   120
            MaxLength       =   6
            TabIndex        =   112
            Top             =   720
            Width           =   855
         End
         Begin VB.TextBox txtProcCode3 
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   120
            MaxLength       =   6
            TabIndex        =   111
            Top             =   960
            Width           =   855
         End
         Begin VB.TextBox txtProcCode4 
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   120
            MaxLength       =   6
            TabIndex        =   110
            Top             =   1200
            Width           =   855
         End
         Begin VB.TextBox txtProcCode5 
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   120
            MaxLength       =   6
            TabIndex        =   109
            Top             =   1440
            Width           =   855
         End
         Begin VB.CommandButton cmdOPCS 
            Caption         =   "Search"
            Height          =   255
            Index           =   1
            Left            =   480
            TabIndex        =   108
            Top             =   180
            Width           =   735
         End
         Begin VB.TextBox txtProcCode6 
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   120
            MaxLength       =   6
            TabIndex        =   107
            Top             =   1680
            Width           =   855
         End
         Begin VB.CommandButton cmdOPCSSSearch1 
            Caption         =   ">"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   270
            Left            =   1020
            TabIndex        =   106
            Top             =   460
            Width           =   255
         End
         Begin VB.CommandButton cmdOPCSSSearch2 
            Caption         =   ">"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   270
            Left            =   1020
            TabIndex        =   105
            Top             =   715
            Width           =   255
         End
         Begin VB.CommandButton cmdOPCSSSearch3 
            Caption         =   ">"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   270
            Left            =   1020
            TabIndex        =   104
            Top             =   955
            Width           =   255
         End
         Begin VB.CommandButton cmdOPCSSSearch4 
            Caption         =   ">"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   270
            Left            =   1020
            TabIndex        =   103
            Top             =   1210
            Width           =   255
         End
         Begin VB.CommandButton cmdOPCSSSearch5 
            Caption         =   ">"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   270
            Left            =   1020
            TabIndex        =   102
            Top             =   1475
            Width           =   255
         End
         Begin VB.CommandButton cmdOPCSSSearch6 
            Caption         =   ">"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   270
            Left            =   1020
            TabIndex        =   101
            Top             =   1705
            Width           =   255
         End
      End
      Begin VB.Frame Frame13Rate 
         Enabled         =   0   'False
         Height          =   2535
         Left            =   9480
         TabIndex        =   91
         Top             =   360
         Width           =   855
         Begin VB.TextBox txtAnae13Rate8 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   120
            MaxLength       =   8
            TabIndex        =   233
            Top             =   2160
            Width           =   615
         End
         Begin VB.TextBox txtAnae13Rate7 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   120
            MaxLength       =   8
            TabIndex        =   232
            Top             =   1920
            Width           =   615
         End
         Begin VB.TextBox txtTotalAnae13Rate 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   240
            TabIndex        =   98
            Top             =   2880
            Width           =   615
         End
         Begin VB.TextBox txtAnae13Rate1 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   120
            MaxLength       =   8
            TabIndex        =   97
            Top             =   480
            Width           =   615
         End
         Begin VB.TextBox txtAnae13Rate2 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   120
            MaxLength       =   8
            TabIndex        =   96
            Top             =   720
            Width           =   615
         End
         Begin VB.TextBox txtAnae13Rate3 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   120
            MaxLength       =   8
            TabIndex        =   95
            Top             =   960
            Width           =   615
         End
         Begin VB.TextBox txtAnae13Rate4 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   120
            MaxLength       =   8
            TabIndex        =   94
            Top             =   1200
            Width           =   615
         End
         Begin VB.TextBox txtAnae13Rate5 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   120
            MaxLength       =   8
            TabIndex        =   93
            Top             =   1440
            Width           =   615
         End
         Begin VB.TextBox txtAnae13Rate6 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   120
            MaxLength       =   8
            TabIndex        =   92
            Top             =   1680
            Width           =   615
         End
         Begin VB.Line Line3 
            Index           =   1
            X1              =   75
            X2              =   720
            Y1              =   2760
            Y2              =   2760
         End
         Begin VB.Label Label90 
            Alignment       =   2  'Center
            Caption         =   "13th Old"
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
            Index           =   9
            Left            =   80
            TabIndex        =   99
            Top             =   120
            Width           =   735
         End
      End
      Begin VB.Frame FrameAnaesMMA 
         Enabled         =   0   'False
         Height          =   2415
         Left            =   8280
         TabIndex        =   82
         Top             =   3240
         Width           =   855
         Begin VB.TextBox txtAnaesMMARate6 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   120
            MaxLength       =   8
            TabIndex        =   89
            Top             =   1680
            Width           =   615
         End
         Begin VB.TextBox txtAnaesMMARate5 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   120
            MaxLength       =   8
            TabIndex        =   88
            Top             =   1440
            Width           =   615
         End
         Begin VB.TextBox txtAnaesMMARate4 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   120
            MaxLength       =   8
            TabIndex        =   87
            Top             =   1200
            Width           =   615
         End
         Begin VB.TextBox txtAnaesMMARate3 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   120
            MaxLength       =   8
            TabIndex        =   86
            Top             =   960
            Width           =   615
         End
         Begin VB.TextBox txtAnaesMMARate2 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   120
            MaxLength       =   8
            TabIndex        =   85
            Top             =   720
            Width           =   615
         End
         Begin VB.TextBox txtAnaesMMARate1 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   120
            MaxLength       =   8
            TabIndex        =   84
            Top             =   480
            Width           =   615
         End
         Begin VB.TextBox txtTotalAnaesMMARate 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   120
            TabIndex        =   83
            Top             =   2040
            Width           =   615
         End
         Begin VB.Label Label90 
            Alignment       =   2  'Center
            Caption         =   "MMA"
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
            TabIndex        =   90
            Top             =   120
            Width           =   615
         End
         Begin VB.Line Line3 
            Index           =   0
            X1              =   75
            X2              =   720
            Y1              =   1995
            Y2              =   1995
         End
      End
      Begin VB.Frame FrameSMMARate 
         Enabled         =   0   'False
         Height          =   2415
         Left            =   5280
         TabIndex        =   73
         Top             =   3240
         Width           =   855
         Begin VB.TextBox txtSMMARate1 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   120
            MaxLength       =   8
            TabIndex        =   80
            Top             =   480
            Width           =   615
         End
         Begin VB.TextBox txtSTotalMMARate1 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   120
            TabIndex        =   79
            Top             =   2040
            Width           =   615
         End
         Begin VB.TextBox txtSMMARate2 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   120
            MaxLength       =   8
            TabIndex        =   78
            Top             =   720
            Width           =   615
         End
         Begin VB.TextBox txtSMMARate3 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   120
            MaxLength       =   8
            TabIndex        =   77
            Top             =   960
            Width           =   615
         End
         Begin VB.TextBox txtSMMARate4 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   120
            MaxLength       =   8
            TabIndex        =   76
            Top             =   1200
            Width           =   615
         End
         Begin VB.TextBox txtSMMARate5 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   120
            MaxLength       =   8
            TabIndex        =   75
            Top             =   1440
            Width           =   615
         End
         Begin VB.TextBox txtSMMARate6 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   120
            MaxLength       =   8
            TabIndex        =   74
            Top             =   1680
            Width           =   615
         End
         Begin VB.Line Line1 
            Index           =   0
            X1              =   75
            X2              =   720
            Y1              =   1995
            Y2              =   1995
         End
         Begin VB.Label Label90 
            Alignment       =   2  'Center
            Caption         =   "MMA"
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
            Left            =   120
            TabIndex        =   81
            Top             =   120
            Width           =   615
         End
      End
      Begin VB.Frame FrameNSProc 
         Height          =   2535
         Left            =   120
         TabIndex        =   65
         Top             =   360
         Width           =   2775
         Begin VB.TextBox txtNSProcedure8 
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   120
            MaxLength       =   200
            TabIndex        =   263
            Top             =   2160
            Width           =   2535
         End
         Begin VB.TextBox txtNSProcedure7 
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   120
            MaxLength       =   200
            TabIndex        =   262
            Top             =   1920
            Width           =   2535
         End
         Begin VB.TextBox txtNSProcedure1 
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   120
            MaxLength       =   200
            TabIndex        =   71
            Top             =   480
            Width           =   2535
         End
         Begin VB.TextBox txtNSProcedure2 
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   120
            MaxLength       =   200
            TabIndex        =   70
            Top             =   720
            Width           =   2535
         End
         Begin VB.TextBox txtNSProcedure3 
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   120
            MaxLength       =   200
            TabIndex        =   69
            Top             =   960
            Width           =   2535
         End
         Begin VB.TextBox txtNSProcedure4 
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   120
            MaxLength       =   200
            TabIndex        =   68
            Top             =   1200
            Width           =   2535
         End
         Begin VB.TextBox txtNSProcedure5 
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   120
            MaxLength       =   200
            TabIndex        =   67
            Top             =   1440
            Width           =   2535
         End
         Begin VB.TextBox txtNSProcedure6 
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   120
            MaxLength       =   200
            TabIndex        =   66
            Top             =   1700
            Width           =   2535
         End
         Begin VB.Label Label90 
            Caption         =   "Procedures"
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
            Index           =   1
            Left            =   120
            TabIndex        =   72
            Top             =   120
            Width           =   2175
         End
      End
      Begin VB.Frame FrameNSOPCS 
         Caption         =   "OPCS"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2535
         Left            =   2880
         TabIndex        =   51
         Top             =   360
         Width           =   1335
         Begin VB.TextBox txtNSProcCode8 
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   120
            MaxLength       =   6
            TabIndex        =   261
            Top             =   2160
            Width           =   855
         End
         Begin VB.TextBox txtNSProcCode7 
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   120
            MaxLength       =   6
            TabIndex        =   260
            Top             =   1920
            Width           =   855
         End
         Begin VB.TextBox txtNSProcCode1 
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   120
            MaxLength       =   6
            TabIndex        =   64
            Top             =   480
            Width           =   855
         End
         Begin VB.TextBox txtNSProcCode2 
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   120
            MaxLength       =   6
            TabIndex        =   63
            Top             =   720
            Width           =   855
         End
         Begin VB.TextBox txtNSProcCode3 
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   120
            MaxLength       =   6
            TabIndex        =   62
            Top             =   960
            Width           =   855
         End
         Begin VB.TextBox txtNSProcCode4 
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   120
            MaxLength       =   6
            TabIndex        =   61
            Top             =   1200
            Width           =   855
         End
         Begin VB.TextBox txtNSProcCode5 
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   120
            MaxLength       =   6
            TabIndex        =   60
            Top             =   1440
            Width           =   855
         End
         Begin VB.CommandButton cmdOPCS 
            Caption         =   "Search"
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
            Left            =   480
            TabIndex        =   59
            Top             =   200
            Width           =   735
         End
         Begin VB.TextBox txtNSProcCode6 
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   120
            MaxLength       =   6
            TabIndex        =   58
            Top             =   1700
            Width           =   855
         End
         Begin VB.CommandButton cmdOPCSNSSearch1 
            Caption         =   ">"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   270
            Left            =   1020
            TabIndex        =   57
            Top             =   470
            Width           =   255
         End
         Begin VB.CommandButton cmdOPCSNSSearch2 
            Caption         =   ">"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   270
            Left            =   1020
            TabIndex        =   56
            Top             =   720
            Width           =   255
         End
         Begin VB.CommandButton cmdOPCSNSSearch3 
            Caption         =   ">"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   270
            Left            =   1020
            TabIndex        =   55
            Top             =   960
            Width           =   255
         End
         Begin VB.CommandButton cmdOPCSNSSearch4 
            Caption         =   ">"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   270
            Left            =   1020
            TabIndex        =   54
            Top             =   1215
            Width           =   255
         End
         Begin VB.CommandButton cmdOPCSNSSearch5 
            Caption         =   ">"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   270
            Left            =   1020
            TabIndex        =   53
            Top             =   1470
            Width           =   255
         End
         Begin VB.CommandButton cmdOPCSNSSearch6 
            Caption         =   ">"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   270
            Left            =   1020
            TabIndex        =   52
            Top             =   1720
            Width           =   255
         End
      End
      Begin VB.Frame FrameS13MinRate 
         Enabled         =   0   'False
         Height          =   2415
         Left            =   6840
         TabIndex        =   42
         Top             =   3360
         Width           =   855
         Begin VB.TextBox txtS13MinRate6 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   120
            MaxLength       =   8
            TabIndex        =   49
            Top             =   1680
            Width           =   615
         End
         Begin VB.TextBox txtS13MinRate5 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   120
            MaxLength       =   8
            TabIndex        =   48
            Top             =   1440
            Width           =   615
         End
         Begin VB.TextBox txtS13MinRate4 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   120
            MaxLength       =   8
            TabIndex        =   47
            Top             =   1200
            Width           =   615
         End
         Begin VB.TextBox txtS13MinRate3 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   120
            MaxLength       =   8
            TabIndex        =   46
            Top             =   960
            Width           =   615
         End
         Begin VB.TextBox txtS13MinRate2 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   120
            MaxLength       =   8
            TabIndex        =   45
            Top             =   720
            Width           =   615
         End
         Begin VB.TextBox txtTotalS13MinRate 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   120
            TabIndex        =   44
            Top             =   2040
            Width           =   615
         End
         Begin VB.TextBox txtS13MinRate1 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   120
            MaxLength       =   8
            TabIndex        =   43
            Top             =   480
            Width           =   615
         End
         Begin VB.Label Label90 
            Alignment       =   2  'Center
            Caption         =   "13th Old (Min)"
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
            Index           =   11
            Left            =   80
            TabIndex        =   50
            Top             =   110
            Width           =   735
         End
         Begin VB.Line Line1 
            Index           =   3
            X1              =   75
            X2              =   720
            Y1              =   1995
            Y2              =   1995
         End
      End
      Begin VB.Frame FrameNS13RateMin 
         Enabled         =   0   'False
         Height          =   2415
         Left            =   6840
         TabIndex        =   33
         Top             =   3240
         Width           =   855
         Begin VB.TextBox txtNS13MinRate1 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   120
            MaxLength       =   8
            TabIndex        =   40
            Top             =   480
            Width           =   615
         End
         Begin VB.TextBox txtTotalNS13MinRate 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   120
            TabIndex        =   39
            Top             =   2040
            Width           =   615
         End
         Begin VB.TextBox txtNS13MinRate2 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   120
            MaxLength       =   8
            TabIndex        =   38
            Top             =   720
            Width           =   615
         End
         Begin VB.TextBox txtNS13MinRate3 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   120
            MaxLength       =   8
            TabIndex        =   37
            Top             =   960
            Width           =   615
         End
         Begin VB.TextBox txtNS13MinRate4 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   120
            MaxLength       =   8
            TabIndex        =   36
            Top             =   1200
            Width           =   615
         End
         Begin VB.TextBox txtNS13MinRate5 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   120
            MaxLength       =   8
            TabIndex        =   35
            Top             =   1440
            Width           =   615
         End
         Begin VB.TextBox txtNS13MinRate6 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   120
            MaxLength       =   8
            TabIndex        =   34
            Top             =   1680
            Width           =   615
         End
         Begin VB.Line Line1 
            Index           =   4
            X1              =   75
            X2              =   720
            Y1              =   1995
            Y2              =   1995
         End
         Begin VB.Label Label90 
            Alignment       =   2  'Center
            Caption         =   "13th Sch (Min)"
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
            Index           =   12
            Left            =   75
            TabIndex        =   41
            Top             =   110
            Width           =   735
         End
      End
      Begin VB.Label Label8 
         Caption         =   "Hospital OT"
         Height          =   255
         Left            =   120
         TabIndex        =   185
         Top             =   3000
         Width           =   1335
      End
      Begin VB.Label Label1 
         Caption         =   "Total"
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
         Left            =   3720
         TabIndex        =   178
         Top             =   3000
         Visible         =   0   'False
         Width           =   615
      End
      Begin VB.Label Label5 
         Caption         =   "Surgical"
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
         Left            =   5520
         TabIndex        =   177
         Top             =   120
         Width           =   2415
      End
      Begin VB.Label Label7 
         Caption         =   "Surgical"
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
         Left            =   120
         TabIndex        =   175
         Top             =   120
         Width           =   4215
      End
      Begin VB.Label Label6 
         Caption         =   "Anaes"
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
         Left            =   8640
         TabIndex        =   176
         Top             =   120
         Width           =   1815
      End
   End
   Begin VB.Frame Frame20 
      Height          =   1275
      Left            =   5160
      TabIndex        =   23
      Top             =   1800
      Width           =   6135
      Begin VB.TextBox txtPROCINVDone1 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   120
         MaxLength       =   150
         TabIndex        =   30
         Top             =   405
         Width           =   3855
      End
      Begin VB.TextBox txtPROCINVDone2 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   120
         MaxLength       =   150
         TabIndex        =   29
         Top             =   660
         Width           =   3855
      End
      Begin VB.TextBox txtPROCINVDone3 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   120
         MaxLength       =   150
         TabIndex        =   28
         Top             =   915
         Width           =   3855
      End
      Begin VB.TextBox txtINVCode1 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   3960
         MaxLength       =   150
         TabIndex        =   27
         Top             =   405
         Width           =   1335
      End
      Begin VB.TextBox txtINVCode2 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   3960
         MaxLength       =   150
         TabIndex        =   26
         Top             =   660
         Width           =   1335
      End
      Begin VB.TextBox txtINVCode3 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   3960
         MaxLength       =   150
         TabIndex        =   25
         Top             =   915
         Width           =   1335
      End
      Begin VB.CommandButton cmdInvestigation 
         Caption         =   "Search"
         Height          =   255
         Left            =   4560
         TabIndex        =   24
         Top             =   120
         Width           =   735
      End
      Begin VB.Label Label90 
         Caption         =   "Investigation"
         Height          =   495
         Index           =   0
         Left            =   120
         TabIndex        =   31
         Top             =   120
         Width           =   2655
      End
   End
   Begin VB.Frame Frame5 
      Height          =   1635
      Left            =   5160
      TabIndex        =   16
      Top             =   240
      Width           =   6135
      Begin VB.TextBox txtAnaesDoc 
         Appearance      =   0  'Flat
         Height          =   285
         Index           =   2
         Left            =   1395
         MaxLength       =   30
         TabIndex        =   19
         Top             =   750
         Width           =   3960
      End
      Begin VB.TextBox txtAnaesDoc 
         Appearance      =   0  'Flat
         Height          =   285
         Index           =   1
         Left            =   1395
         MaxLength       =   30
         TabIndex        =   18
         Top             =   495
         Width           =   3960
      End
      Begin VB.TextBox txtAnaesDoc 
         Appearance      =   0  'Flat
         Height          =   285
         Index           =   0
         Left            =   1395
         MaxLength       =   30
         TabIndex        =   17
         Top             =   240
         Width           =   3960
      End
      Begin VB.Label Label2 
         Caption         =   "Anaes Name 3"
         Height          =   255
         Index           =   7
         Left            =   120
         TabIndex        =   22
         Top             =   795
         Width           =   1095
      End
      Begin VB.Label Label2 
         Caption         =   "Anaes Name 2"
         Height          =   255
         Index           =   6
         Left            =   120
         TabIndex        =   21
         Top             =   540
         Width           =   1095
      End
      Begin VB.Label Label2 
         Caption         =   "Anaes Name 1"
         Height          =   255
         Index           =   5
         Left            =   120
         TabIndex        =   20
         Top             =   270
         Width           =   1095
      End
   End
   Begin VB.Frame Frame21 
      Height          =   1275
      Left            =   120
      TabIndex        =   11
      Top             =   1800
      Width           =   4995
      Begin VB.TextBox txtMEDSURDone1 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   120
         MaxLength       =   150
         TabIndex        =   14
         Top             =   405
         Width           =   4815
      End
      Begin VB.TextBox txtMEDSURDone2 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   120
         MaxLength       =   150
         TabIndex        =   13
         Top             =   660
         Width           =   4815
      End
      Begin VB.TextBox txtMEDSURDone3 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   120
         MaxLength       =   150
         TabIndex        =   12
         Top             =   915
         Width           =   4815
      End
      Begin VB.Label Label95 
         Caption         =   "Medical Management"
         Height          =   495
         Left            =   120
         TabIndex        =   15
         Top             =   120
         Width           =   1935
      End
   End
   Begin VB.Frame Frame2 
      Height          =   1635
      Left            =   120
      TabIndex        =   0
      Top             =   240
      Width           =   5010
      Begin VB.TextBox txtATTDoc 
         Appearance      =   0  'Flat
         Height          =   285
         Index           =   0
         Left            =   1440
         MaxLength       =   100
         TabIndex        =   5
         Top             =   165
         Width           =   3480
      End
      Begin VB.TextBox txtATTDoc 
         Appearance      =   0  'Flat
         Height          =   285
         Index           =   1
         Left            =   1440
         MaxLength       =   30
         TabIndex        =   4
         Top             =   420
         Width           =   3480
      End
      Begin VB.TextBox txtATTDoc 
         Appearance      =   0  'Flat
         Height          =   285
         Index           =   2
         Left            =   1440
         MaxLength       =   30
         TabIndex        =   3
         Top             =   675
         Width           =   3480
      End
      Begin VB.TextBox txtATTDoc 
         Appearance      =   0  'Flat
         Height          =   285
         Index           =   3
         Left            =   1440
         MaxLength       =   30
         TabIndex        =   2
         Top             =   915
         Width           =   3480
      End
      Begin VB.TextBox txtATTDoc 
         Appearance      =   0  'Flat
         Height          =   285
         Index           =   4
         Left            =   1440
         MaxLength       =   30
         TabIndex        =   1
         Top             =   1170
         Width           =   3480
      End
      Begin VB.Label Label2 
         Caption         =   "Attending Doc 1"
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   10
         Top             =   225
         Width           =   1335
      End
      Begin VB.Label Label2 
         Caption         =   "Attending Doc 2"
         Height          =   255
         Index           =   1
         Left            =   120
         TabIndex        =   9
         Top             =   495
         Width           =   1335
      End
      Begin VB.Label Label2 
         Caption         =   "Attending Doc 3"
         Height          =   255
         Index           =   2
         Left            =   120
         TabIndex        =   8
         Top             =   765
         Width           =   1335
      End
      Begin VB.Label Label2 
         Caption         =   "Attending Doc 4"
         Height          =   255
         Index           =   3
         Left            =   120
         TabIndex        =   7
         Top             =   1005
         Width           =   1335
      End
      Begin VB.Label Label2 
         Caption         =   "Attending Doc 5"
         Height          =   255
         Index           =   4
         Left            =   120
         TabIndex        =   6
         Top             =   1245
         Width           =   1335
      End
   End
   Begin VB.Frame Frame1 
      Height          =   615
      Left            =   120
      TabIndex        =   179
      Top             =   6360
      Width           =   11175
      Begin VB.CommandButton cmdExit 
         Caption         =   "&Exit"
         Height          =   375
         Left            =   10320
         TabIndex        =   181
         Top             =   150
         Width           =   735
      End
      Begin VB.CommandButton cmdAdd 
         Caption         =   "Save"
         Height          =   375
         Left            =   9600
         TabIndex        =   180
         Top             =   150
         Width           =   735
      End
      Begin VB.Label Label4 
         Caption         =   "Case Number"
         Height          =   255
         Left            =   120
         TabIndex        =   183
         Top             =   240
         Width           =   1215
      End
      Begin VB.Label lblCasNumber 
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
         Left            =   1200
         TabIndex        =   182
         Top             =   240
         Width           =   1575
      End
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00404040&
      Caption         =   "Surgical Procedures Entry"
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
      TabIndex        =   184
      Top             =   0
      Width           =   11205
   End
End
Attribute VB_Name = "frmMMAEntryGST"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public Source As Integer
Public CalSource As Integer
Dim MMASurgFees, MMAAnaeFees As Double
Dim MMASurgFees13thMin, MMASurgFees13th, MMAAnaeFees13th As Double
Dim MMADesc, Sch13 As String
Dim BodyMMA, Body13 As String
Dim PageMMA, Page13 As String
Dim tmpNSFee As Double

Private Sub cmdAdd_Click()
    If Source = 0 Then
        frmInvoiceGST.txtSurgicalFee = txtTotalSActFee
        frmInvoiceGST.txtAnaFees = txtTotalAnaesActFees
    ElseIf Source = 1 Then
        frmInvoiceGST.txtProcFees = txtTotalNSActFee
    End If
End Sub

Private Sub CmdExit_Click()
    frmMMAEntryGST.Visible = False
End Sub

Private Sub cmdInvestigation_Click()
    frmSearchInv.Source = 3
    frmSearchInv.Show vbModal
End Sub

Private Sub cmdOPCS_Click(a As Integer)
    If a = 1 Then
        frmSearchCPT.Source = 1
    Else
        frmSearchCPT.Source = 0
    End If
        frmSearchCPT.Show vbModal
End Sub

Private Sub cmdOPCSNSSearch1_Click()
    Call SearchMMADesc(txtNSProcCode1)
    Load frmMMA13thSchDesc
    frmMMA13thSchDesc.Show
    frmMMA13thSchDesc.lblMMACode = txtNSProcCode1
    frmMMA13thSchDesc.lblMMADesc = MMADesc
    frmMMA13thSchDesc.lbl13thSchDesc = Sch13
    frmMMA13thSchDesc.lblBodyMMA = BodyMMA
    frmMMA13thSchDesc.lblPageMMA = PageMMA
    frmMMA13thSchDesc.lblBody13th = Body13
    frmMMA13thSchDesc.lblPage13th = Page13
    frmMMA13thSchDesc.lblSurgMMA = MMASurgFees
    frmMMA13thSchDesc.lblAnaeMMA = MMAAnaeFees
    frmMMA13thSchDesc.lblSurg13thMin = MMASurgFees13thMin
    frmMMA13thSchDesc.lblSurg13thMax = MMASurgFees13th
    frmMMA13thSchDesc.lblAnae13th = MMAAnaeFees13th
End Sub

Private Sub cmdOPCSNSSearch2_Click()
    Call SearchMMADesc(txtNSProcCode2)
    Load frmMMA13thSchDesc
    frmMMA13thSchDesc.Show
    frmMMA13thSchDesc.lblMMACode = txtNSProcCode2
    frmMMA13thSchDesc.lblMMADesc = MMADesc
    frmMMA13thSchDesc.lbl13thSchDesc = Sch13
    frmMMA13thSchDesc.lblBodyMMA = BodyMMA
    frmMMA13thSchDesc.lblPageMMA = PageMMA
    frmMMA13thSchDesc.lblBody13th = Body13
    frmMMA13thSchDesc.lblPage13th = Page13
    frmMMA13thSchDesc.lblSurgMMA = MMASurgFees
    frmMMA13thSchDesc.lblAnaeMMA = MMAAnaeFees
    frmMMA13thSchDesc.lblSurg13thMin = MMASurgFees13thMin
    frmMMA13thSchDesc.lblSurg13thMax = MMASurgFees13th
    frmMMA13thSchDesc.lblAnae13th = MMAAnaeFees13th
End Sub

Private Sub cmdOPCSNSSearch3_Click()
    Call SearchMMADesc(txtNSProcCode3)
    Load frmMMA13thSchDesc
    frmMMA13thSchDesc.Show
    frmMMA13thSchDesc.lblMMACode = txtNSProcCode3
    frmMMA13thSchDesc.lblMMADesc = MMADesc
    frmMMA13thSchDesc.lbl13thSchDesc = Sch13
    frmMMA13thSchDesc.lblBodyMMA = BodyMMA
    frmMMA13thSchDesc.lblPageMMA = PageMMA
    frmMMA13thSchDesc.lblBody13th = Body13
    frmMMA13thSchDesc.lblPage13th = Page13
    frmMMA13thSchDesc.lblSurgMMA = MMASurgFees
    frmMMA13thSchDesc.lblAnaeMMA = MMAAnaeFees
    frmMMA13thSchDesc.lblSurg13thMin = MMASurgFees13thMin
    frmMMA13thSchDesc.lblSurg13thMax = MMASurgFees13th
    frmMMA13thSchDesc.lblAnae13th = MMAAnaeFees13th
End Sub

Private Sub cmdOPCSNSSearch4_Click()
    Call SearchMMADesc(txtNSProcCode4)
    Load frmMMA13thSchDesc
    frmMMA13thSchDesc.Show
    frmMMA13thSchDesc.lblMMACode = txtNSProcCode4
    frmMMA13thSchDesc.lblMMADesc = MMADesc
    frmMMA13thSchDesc.lbl13thSchDesc = Sch13
    frmMMA13thSchDesc.lblBodyMMA = BodyMMA
    frmMMA13thSchDesc.lblPageMMA = PageMMA
    frmMMA13thSchDesc.lblBody13th = Body13
    frmMMA13thSchDesc.lblPage13th = Page13
    frmMMA13thSchDesc.lblSurgMMA = MMASurgFees
    frmMMA13thSchDesc.lblAnaeMMA = MMAAnaeFees
    frmMMA13thSchDesc.lblSurg13thMin = MMASurgFees13thMin
    frmMMA13thSchDesc.lblSurg13thMax = MMASurgFees13th
    frmMMA13thSchDesc.lblAnae13th = MMAAnaeFees13th
End Sub

Private Sub cmdOPCSNSSearch5_Click()
    Call SearchMMADesc(txtNSProcCode5)
    Load frmMMA13thSchDesc
    frmMMA13thSchDesc.Show
    frmMMA13thSchDesc.lblMMACode = txtNSProcCode5
    frmMMA13thSchDesc.lblMMADesc = MMADesc
    frmMMA13thSchDesc.lbl13thSchDesc = Sch13
    frmMMA13thSchDesc.lblBodyMMA = BodyMMA
    frmMMA13thSchDesc.lblPageMMA = PageMMA
    frmMMA13thSchDesc.lblBody13th = Body13
    frmMMA13thSchDesc.lblPage13th = Page13
    frmMMA13thSchDesc.lblSurgMMA = MMASurgFees
    frmMMA13thSchDesc.lblAnaeMMA = MMAAnaeFees
    frmMMA13thSchDesc.lblSurg13thMin = MMASurgFees13thMin
    frmMMA13thSchDesc.lblSurg13thMax = MMASurgFees13th
    frmMMA13thSchDesc.lblAnae13th = MMAAnaeFees13th
End Sub

Private Sub cmdOPCSNSSearch6_Click()
    Call SearchMMADesc(txtNSProcCode6)
    Load frmMMA13thSchDesc
    frmMMA13thSchDesc.Show
    frmMMA13thSchDesc.lblMMACode = txtNSProcCode6
    frmMMA13thSchDesc.lblMMADesc = MMADesc
    frmMMA13thSchDesc.lbl13thSchDesc = Sch13
    frmMMA13thSchDesc.lblBodyMMA = BodyMMA
    frmMMA13thSchDesc.lblPageMMA = PageMMA
    frmMMA13thSchDesc.lblBody13th = Body13
    frmMMA13thSchDesc.lblPage13th = Page13
    frmMMA13thSchDesc.lblSurgMMA = MMASurgFees
    frmMMA13thSchDesc.lblAnaeMMA = MMAAnaeFees
    frmMMA13thSchDesc.lblSurg13thMin = MMASurgFees13thMin
    frmMMA13thSchDesc.lblSurg13thMax = MMASurgFees13th
    frmMMA13thSchDesc.lblAnae13th = MMAAnaeFees13th
End Sub

Private Sub cmdOPCSSSearch1_Click()
    Call SearchMMADesc(txtProcCode1)
    Load frmMMA13thSchDesc
    frmMMA13thSchDesc.Show
    frmMMA13thSchDesc.lblMMACode = txtProcCode1
    frmMMA13thSchDesc.lblMMADesc = MMADesc
    frmMMA13thSchDesc.lbl13thSchDesc = Sch13
    frmMMA13thSchDesc.lblBodyMMA = BodyMMA
    frmMMA13thSchDesc.lblPageMMA = PageMMA
    frmMMA13thSchDesc.lblBody13th = Body13
    frmMMA13thSchDesc.lblPage13th = Page13
    frmMMA13thSchDesc.lblSurgMMA = MMASurgFees
    frmMMA13thSchDesc.lblAnaeMMA = MMAAnaeFees
    frmMMA13thSchDesc.lblSurg13thMin = MMASurgFees13thMin
    frmMMA13thSchDesc.lblSurg13thMax = MMASurgFees13th
    frmMMA13thSchDesc.lblAnae13th = MMAAnaeFees13th
End Sub

Private Sub cmdOPCSSSearch2_Click()
    Call SearchMMADesc(txtProcCode2)
    Load frmMMA13thSchDesc
    frmMMA13thSchDesc.Show
    frmMMA13thSchDesc.lblMMACode = txtProcCode2
    frmMMA13thSchDesc.lblMMADesc = MMADesc
    frmMMA13thSchDesc.lbl13thSchDesc = Sch13
    frmMMA13thSchDesc.lblBodyMMA = BodyMMA
    frmMMA13thSchDesc.lblPageMMA = PageMMA
    frmMMA13thSchDesc.lblBody13th = Body13
    frmMMA13thSchDesc.lblPage13th = Page13
    frmMMA13thSchDesc.lblSurgMMA = MMASurgFees
    frmMMA13thSchDesc.lblAnaeMMA = MMAAnaeFees
    frmMMA13thSchDesc.lblSurg13thMin = MMASurgFees13thMin
    frmMMA13thSchDesc.lblSurg13thMax = MMASurgFees13th
    frmMMA13thSchDesc.lblAnae13th = MMAAnaeFees13th
End Sub

Private Sub cmdOPCSSSearch3_Click()
    Call SearchMMADesc(txtProcCode3)
    Load frmMMA13thSchDesc
    frmMMA13thSchDesc.Show
    frmMMA13thSchDesc.lblMMACode = txtProcCode3
    frmMMA13thSchDesc.lblMMADesc = MMADesc
    frmMMA13thSchDesc.lbl13thSchDesc = Sch13
    frmMMA13thSchDesc.lblBodyMMA = BodyMMA
    frmMMA13thSchDesc.lblPageMMA = PageMMA
    frmMMA13thSchDesc.lblBody13th = Body13
    frmMMA13thSchDesc.lblPage13th = Page13
    frmMMA13thSchDesc.lblSurgMMA = MMASurgFees
    frmMMA13thSchDesc.lblAnaeMMA = MMAAnaeFees
    frmMMA13thSchDesc.lblSurg13thMin = MMASurgFees13thMin
    frmMMA13thSchDesc.lblSurg13thMax = MMASurgFees13th
    frmMMA13thSchDesc.lblAnae13th = MMAAnaeFees13th
End Sub

Private Sub cmdOPCSSSearch4_Click()
    Call SearchMMADesc(txtProcCode4)
    Load frmMMA13thSchDesc
    frmMMA13thSchDesc.Show
    frmMMA13thSchDesc.lblMMACode = txtProcCode4
    frmMMA13thSchDesc.lblMMADesc = MMADesc
    frmMMA13thSchDesc.lbl13thSchDesc = Sch13
    frmMMA13thSchDesc.lblBodyMMA = BodyMMA
    frmMMA13thSchDesc.lblPageMMA = PageMMA
    frmMMA13thSchDesc.lblBody13th = Body13
    frmMMA13thSchDesc.lblPage13th = Page13
    frmMMA13thSchDesc.lblSurgMMA = MMASurgFees
    frmMMA13thSchDesc.lblAnaeMMA = MMAAnaeFees
    frmMMA13thSchDesc.lblSurg13thMin = MMASurgFees13thMin
    frmMMA13thSchDesc.lblSurg13thMax = MMASurgFees13th
    frmMMA13thSchDesc.lblAnae13th = MMAAnaeFees13th
End Sub

Private Sub cmdOPCSSSearch5_Click()
    Call SearchMMADesc(txtProcCode5)
    Load frmMMA13thSchDesc
    frmMMA13thSchDesc.Show
    frmMMA13thSchDesc.lblMMACode = txtProcCode5
    frmMMA13thSchDesc.lblMMADesc = MMADesc
    frmMMA13thSchDesc.lbl13thSchDesc = Sch13
    frmMMA13thSchDesc.lblBodyMMA = BodyMMA
    frmMMA13thSchDesc.lblPageMMA = PageMMA
    frmMMA13thSchDesc.lblBody13th = Body13
    frmMMA13thSchDesc.lblPage13th = Page13
    frmMMA13thSchDesc.lblSurgMMA = MMASurgFees
    frmMMA13thSchDesc.lblAnaeMMA = MMAAnaeFees
    frmMMA13thSchDesc.lblSurg13thMin = MMASurgFees13thMin
    frmMMA13thSchDesc.lblSurg13thMax = MMASurgFees13th
    frmMMA13thSchDesc.lblAnae13th = MMAAnaeFees13th
End Sub

Private Sub cmdOPCSSSearch6_Click()
    Call SearchMMADesc(txtProcCode6)
    Load frmMMA13thSchDesc
    frmMMA13thSchDesc.Show
    frmMMA13thSchDesc.lblMMACode = txtProcCode6
    frmMMA13thSchDesc.lblMMADesc = MMADesc
    frmMMA13thSchDesc.lbl13thSchDesc = Sch13
    frmMMA13thSchDesc.lblBodyMMA = BodyMMA
    frmMMA13thSchDesc.lblPageMMA = PageMMA
    frmMMA13thSchDesc.lblBody13th = Body13
    frmMMA13thSchDesc.lblPage13th = Page13
    frmMMA13thSchDesc.lblSurgMMA = MMASurgFees
    frmMMA13thSchDesc.lblAnaeMMA = MMAAnaeFees
    frmMMA13thSchDesc.lblSurg13thMin = MMASurgFees13thMin
    frmMMA13thSchDesc.lblSurg13thMax = MMASurgFees13th
    frmMMA13thSchDesc.lblAnae13th = MMAAnaeFees13th
End Sub

Private Sub Form_GotFocus()
    If Source = 0 Then
        frmMMAEntryGST.FrameNSProc.Visible = False
        frmMMAEntryGST.FrameNSOPCS.Visible = False
        frmMMAEntryGST.FrameNSMMARate.Visible = False
        frmMMAEntryGST.FrameNSActFee.Visible = False
        frmMMAEntryGST.FrameNS13Rate.Visible = False
        frmMMAEntryGST.FrameNS13RateMin.Visible = False
        frmMMAEntryGST.FrameNS14Rate.Visible = False
        frmMMAEntryGST.FrameNS14RateMin.Visible = False
        
        frmMMAEntryGST.FrameSProc.Visible = True
        frmMMAEntryGST.FrameSOPCS.Visible = True
        frmMMAEntryGST.FrameSMMARate.Visible = True
        frmMMAEntryGST.FrameSActFee.Visible = True
        frmMMAEntryGST.FrameS13Rate.Visible = True
        frmMMAEntryGST.FrameS14Rate.Visible = True
        frmMMAEntryGST.FrameAnaesMMA.Visible = True
        frmMMAEntryGST.FrameAnaesAct.Visible = True
        frmMMAEntryGST.Frame13Rate.Visible = True
        frmMMAEntryGST.Frame14Rate.Visible = True
        frmMMAEntryGST.Label6.Visible = True
        frmMMAEntryGST.FrameS13MinRate.Visible = True
        frmMMAEntryGST.FrameS14MinRate.Visible = True
        frmMMAEntryGST.Label3.Caption = "Surgical Procedures Entry"
    Else
        frmMMAEntryGST.FrameNSProc.Visible = True
        frmMMAEntryGST.FrameNSOPCS.Visible = True
        frmMMAEntryGST.FrameNSMMARate.Visible = True
        frmMMAEntryGST.FrameNSActFee.Visible = True
        frmMMAEntryGST.FrameNS13Rate.Visible = True
        frmMMAEntryGST.FrameNS13RateMin.Visible = True
        frmMMAEntryGST.FrameNS14Rate.Visible = True
        frmMMAEntryGST.FrameNS14RateMin.Visible = True
        
        frmMMAEntryGST.FrameSProc.Visible = False
        frmMMAEntryGST.FrameSOPCS.Visible = False
        frmMMAEntryGST.FrameSMMARate.Visible = False
        frmMMAEntryGST.FrameS13Rate.Visible = False
        frmMMAEntryGST.FrameS14Rate.Visible = False
        frmMMAEntryGST.FrameSActFee.Visible = False
        frmMMAEntryGST.FrameAnaesMMA.Visible = False
        frmMMAEntryGST.FrameAnaesAct.Visible = False
        frmMMAEntryGST.Frame13Rate.Visible = False
        frmMMAEntryGST.Frame14Rate.Visible = False
        frmMMAEntryGST.Label6.Visible = False
        frmMMAEntryGST.FrameS13MinRate.Visible = False
        frmMMAEntryGST.FrameS14MinRate.Visible = False
        frmMMAEntryGST.Label3.Caption = "Non-Surgical Procedures Entry"
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

Private Sub Form_Load()
    If Source = 0 Then
        frmMMAEntryGST.FrameNSProc.Visible = False
        frmMMAEntryGST.FrameNSOPCS.Visible = False
        frmMMAEntryGST.FrameNSMMARate.Visible = False
        frmMMAEntryGST.FrameNSActFee.Visible = False
        frmMMAEntryGST.FrameSProc.Visible = True
        frmMMAEntryGST.FrameSOPCS.Visible = True
        frmMMAEntryGST.FrameSMMARate.Visible = True
        frmMMAEntryGST.FrameSActFee.Visible = True
        frmMMAEntryGST.FrameAnaesMMA.Visible = True
        frmMMAEntryGST.FrameAnaesAct.Visible = True
        frmMMAEntryGST.Label3.Caption = "Surgical Procedures Entry"
    Else
        frmMMAEntryGST.FrameNSProc.Visible = True
        frmMMAEntryGST.FrameNSOPCS.Visible = True
        frmMMAEntryGST.FrameNSMMARate.Visible = True
        frmMMAEntryGST.FrameNSActFee.Visible = True
        frmMMAEntryGST.FrameSProc.Visible = False
        frmMMAEntryGST.FrameSOPCS.Visible = False
        frmMMAEntryGST.FrameSMMARate.Visible = False
        frmMMAEntryGST.FrameSActFee.Visible = False
        frmMMAEntryGST.FrameAnaesMMA.Visible = False
        frmMMAEntryGST.FrameAnaesAct.Visible = False
        frmMMAEntryGST.Label3.Caption = "Non-Surgical Procedures Entry"
    End If
    
    frmMMAEntryGST.cboHospOT.AddItem "Y - Yes"
    frmMMAEntryGST.cboHospOT.AddItem "N - No"
End Sub



Private Sub txtAna13Rate2_Change()
Dim a, B, C, d, e, f, g As Double

    a = IIf(IsNumeric(txtAnae13Rate1), txtAnae13Rate1, 0)
    B = IIf(IsNumeric(txtAnae13Rate2), txtAnae13Rate2, 0)
    C = IIf(IsNumeric(txtAnae13Rate3), txtAnae13Rate3, 0)
    d = IIf(IsNumeric(txtAnae13Rate4), txtAnae13Rate4, 0)
    e = IIf(IsNumeric(txtAnae13Rate5), txtAnae13Rate5, 0)
    g = IIf(IsNumeric(txtAnae13Rate6), txtAnae13Rate6, 0)
    
    f = CDbl(a) + CDbl(B) + CDbl(C) + CDbl(d) + CDbl(e) + CDbl(g)
End Sub

Private Sub txtAna13Rate3_Change()
Dim a, B, C, d, e, f, g As Double

    a = IIf(IsNumeric(txtAnae13Rate1), txtAnae13Rate1, 0)
    B = IIf(IsNumeric(txtAnae13Rate2), txtAnae13Rate2, 0)
    C = IIf(IsNumeric(txtAnae13Rate3), txtAnae13Rate3, 0)
    d = IIf(IsNumeric(txtAnae13Rate4), txtAnae13Rate4, 0)
    e = IIf(IsNumeric(txtAnae13Rate5), txtAnae13Rate5, 0)
    g = IIf(IsNumeric(txtAnae13Rate6), txtAnae13Rate6, 0)
    
    f = CDbl(a) + CDbl(B) + CDbl(C) + CDbl(d) + CDbl(e) + CDbl(g)
End Sub

Private Sub txtAna13Rate4_Change()
Dim a, B, C, d, e, f, g As Double

    a = IIf(IsNumeric(txtAnae13Rate1), txtAnae13Rate1, 0)
    B = IIf(IsNumeric(txtAnae13Rate2), txtAnae13Rate2, 0)
    C = IIf(IsNumeric(txtAnae13Rate3), txtAnae13Rate3, 0)
    d = IIf(IsNumeric(txtAnae13Rate4), txtAnae13Rate4, 0)
    e = IIf(IsNumeric(txtAnae13Rate5), txtAnae13Rate5, 0)
    g = IIf(IsNumeric(txtAnae13Rate6), txtAnae13Rate6, 0)
    
    f = CDbl(a) + CDbl(B) + CDbl(C) + CDbl(d) + CDbl(e) + CDbl(g)
End Sub

Private Sub txtAna13Rate5_Change()
Dim a, B, C, d, e, f, g As Double

    a = IIf(IsNumeric(txtAnae13Rate1), txtAnae13Rate1, 0)
    B = IIf(IsNumeric(txtAnae13Rate2), txtAnae13Rate2, 0)
    C = IIf(IsNumeric(txtAnae13Rate3), txtAnae13Rate3, 0)
    d = IIf(IsNumeric(txtAnae13Rate4), txtAnae13Rate4, 0)
    e = IIf(IsNumeric(txtAnae13Rate5), txtAnae13Rate5, 0)
    g = IIf(IsNumeric(txtAnae13Rate6), txtAnae13Rate6, 0)
    
    f = CDbl(a) + CDbl(B) + CDbl(C) + CDbl(d) + CDbl(e) + CDbl(g)
End Sub

Private Sub txtAna13Rate6_Change()
Dim a, B, C, d, e, f, g As Double

    a = IIf(IsNumeric(txtAnae13Rate1), txtAnae13Rate1, 0)
    B = IIf(IsNumeric(txtAnae13Rate2), txtAnae13Rate2, 0)
    C = IIf(IsNumeric(txtAnae13Rate3), txtAnae13Rate3, 0)
    d = IIf(IsNumeric(txtAnae13Rate4), txtAnae13Rate4, 0)
    e = IIf(IsNumeric(txtAnae13Rate5), txtAnae13Rate5, 0)
    g = IIf(IsNumeric(txtAnae13Rate6), txtAnae13Rate6, 0)
    
    f = CDbl(a) + CDbl(B) + CDbl(C) + CDbl(d) + CDbl(e) + CDbl(g)
End Sub

Private Sub Text12_Change()

End Sub

Private Sub txtAnae13Rate1_Change()
Dim a, B, C, d, e, f, g As Double

    a = IIf(IsNumeric(txtAnae13Rate1), txtAnae13Rate1, 0)
    B = IIf(IsNumeric(txtAnae13Rate2), txtAnae13Rate2, 0)
    C = IIf(IsNumeric(txtAnae13Rate3), txtAnae13Rate3, 0)
    d = IIf(IsNumeric(txtAnae13Rate4), txtAnae13Rate4, 0)
    e = IIf(IsNumeric(txtAnae13Rate5), txtAnae13Rate5, 0)
    g = IIf(IsNumeric(txtAnae13Rate6), txtAnae13Rate6, 0)
    
    f = CDbl(a) + CDbl(B) + CDbl(C) + CDbl(d) + CDbl(e) + CDbl(g)
    txtTotalAnae13Rate = f
    
    tmpNSFee = IIf(IsNumeric(txtAnaesActFees1), txtAnaesActFees1, 0)
    
    If CDbl(tmpNSFee) > CDbl(a) Then
        txtAnaesActFees1.ForeColor = &HFF&
    Else
        txtAnaesActFees1.ForeColor = &H80000008
    End If
    

End Sub

Private Sub txtAnae13Rate2_Change()
Dim a, B, C, d, e, f, g As Double

    a = IIf(IsNumeric(txtAnae13Rate1), txtAnae13Rate1, 0)
    B = IIf(IsNumeric(txtAnae13Rate2), txtAnae13Rate2, 0)
    C = IIf(IsNumeric(txtAnae13Rate3), txtAnae13Rate3, 0)
    d = IIf(IsNumeric(txtAnae13Rate4), txtAnae13Rate4, 0)
    e = IIf(IsNumeric(txtAnae13Rate5), txtAnae13Rate5, 0)
    g = IIf(IsNumeric(txtAnae13Rate6), txtAnae13Rate6, 0)
    
    f = CDbl(a) + CDbl(B) + CDbl(C) + CDbl(d) + CDbl(e) + CDbl(g)
    txtTotalAnae13Rate = f
    
    tmpNSFee = IIf(IsNumeric(txtAnaesActFees2), txtAnaesActFees2, 0)
    
    If CDbl(tmpNSFee) > CDbl(B) Then
        txtAnaesActFees2.ForeColor = &HFF&
    Else
        txtAnaesActFees2.ForeColor = &H80000008
    End If
    
End Sub

Private Sub txtAnae13Rate3_Change()
Dim a, B, C, d, e, f, g As Double

    a = IIf(IsNumeric(txtAnae13Rate1), txtAnae13Rate1, 0)
    B = IIf(IsNumeric(txtAnae13Rate2), txtAnae13Rate2, 0)
    C = IIf(IsNumeric(txtAnae13Rate3), txtAnae13Rate3, 0)
    d = IIf(IsNumeric(txtAnae13Rate4), txtAnae13Rate4, 0)
    e = IIf(IsNumeric(txtAnae13Rate5), txtAnae13Rate5, 0)
    g = IIf(IsNumeric(txtAnae13Rate6), txtAnae13Rate6, 0)
    
    f = CDbl(a) + CDbl(B) + CDbl(C) + CDbl(d) + CDbl(e) + CDbl(g)
    txtTotalAnae13Rate = f
    
    tmpNSFee = IIf(IsNumeric(txtAnaesActFees3), txtAnaesActFees3, 0)
    
    If CDbl(tmpNSFee) > CDbl(C) Then
        txtAnaesActFees3.ForeColor = &HFF&
    Else
        txtAnaesActFees3.ForeColor = &H80000008
    End If
    
End Sub

Private Sub txtAnae13Rate4_Change()
Dim a, B, C, d, e, f, g As Double

    a = IIf(IsNumeric(txtAnae13Rate1), txtAnae13Rate1, 0)
    B = IIf(IsNumeric(txtAnae13Rate2), txtAnae13Rate2, 0)
    C = IIf(IsNumeric(txtAnae13Rate3), txtAnae13Rate3, 0)
    d = IIf(IsNumeric(txtAnae13Rate4), txtAnae13Rate4, 0)
    e = IIf(IsNumeric(txtAnae13Rate5), txtAnae13Rate5, 0)
    g = IIf(IsNumeric(txtAnae13Rate6), txtAnae13Rate6, 0)
    
    f = CDbl(a) + CDbl(B) + CDbl(C) + CDbl(d) + CDbl(e) + CDbl(g)
    txtTotalAnae13Rate = f
    
    tmpNSFee = IIf(IsNumeric(txtAnaesActFees4), txtAnaesActFees4, 0)
    
    If CDbl(tmpNSFee) > CDbl(d) Then
        txtAnaesActFees4.ForeColor = &HFF&
    Else
        txtAnaesActFees4.ForeColor = &H80000008
    End If
    
End Sub

Private Sub txtAnae13Rate5_Change()
Dim a, B, C, d, e, f, g As Double

    a = IIf(IsNumeric(txtAnae13Rate1), txtAnae13Rate1, 0)
    B = IIf(IsNumeric(txtAnae13Rate2), txtAnae13Rate2, 0)
    C = IIf(IsNumeric(txtAnae13Rate3), txtAnae13Rate3, 0)
    d = IIf(IsNumeric(txtAnae13Rate4), txtAnae13Rate4, 0)
    e = IIf(IsNumeric(txtAnae13Rate5), txtAnae13Rate5, 0)
    g = IIf(IsNumeric(txtAnae13Rate6), txtAnae13Rate6, 0)
    
    f = CDbl(a) + CDbl(B) + CDbl(C) + CDbl(d) + CDbl(e) + CDbl(g)
    txtTotalAnae13Rate = f
    
    tmpNSFee = IIf(IsNumeric(txtAnaesActFees5), txtAnaesActFees5, 0)
    
    If CDbl(tmpNSFee) > CDbl(e) Then
        txtAnaesActFees5.ForeColor = &HFF&
    Else
        txtAnaesActFees5.ForeColor = &H80000008
    End If
    
End Sub

Private Sub txtAnae13Rate6_Change()
Dim a, B, C, d, e, f, g As Double

    a = IIf(IsNumeric(txtAnae13Rate1), txtAnae13Rate1, 0)
    B = IIf(IsNumeric(txtAnae13Rate2), txtAnae13Rate2, 0)
    C = IIf(IsNumeric(txtAnae13Rate3), txtAnae13Rate3, 0)
    d = IIf(IsNumeric(txtAnae13Rate4), txtAnae13Rate4, 0)
    e = IIf(IsNumeric(txtAnae13Rate5), txtAnae13Rate5, 0)
    g = IIf(IsNumeric(txtAnae13Rate6), txtAnae13Rate6, 0)
    
    f = CDbl(a) + CDbl(B) + CDbl(C) + CDbl(d) + CDbl(e) + CDbl(g)
    txtTotalAnae13Rate = f
    
    tmpNSFee = IIf(IsNumeric(txtAnaesActFees6), txtAnaesActFees6, 0)
    
    If CDbl(tmpNSFee) > CDbl(g) Then
        txtAnaesActFees6.ForeColor = &HFF&
    Else
        txtAnaesActFees6.ForeColor = &H80000008
    End If
    
End Sub

Private Sub txtAnaesActFees1_Change()
Dim a, B, C, d, e, f, g, h, i As Double

    a = IIf(IsNumeric(txtAnaesActFees1), txtAnaesActFees1, 0)
    B = IIf(IsNumeric(txtAnaesActFees2), txtAnaesActFees2, 0)
    C = IIf(IsNumeric(txtAnaesActFees3), txtAnaesActFees3, 0)
    d = IIf(IsNumeric(txtAnaesActFees4), txtAnaesActFees4, 0)
    e = IIf(IsNumeric(txtAnaesActFees5), txtAnaesActFees5, 0)
    g = IIf(IsNumeric(txtAnaesActFees6), txtAnaesActFees6, 0)
    h = IIf(IsNumeric(txtAnaesActFees7), txtAnaesActFees7, 0)
    i = IIf(IsNumeric(txtAnaesActFees8), txtAnaesActFees8, 0)
    
    f = CDbl(a) + CDbl(B) + CDbl(C) + CDbl(d) + CDbl(e) + CDbl(g) + CDbl(h) + CDbl(i)
    txtTotalAnaesActFees = f

    tmpNSFee = IIf(IsNumeric(txtAnae13Rate1), txtAnae13Rate1, 0)
    
    If CDbl(a) > CDbl(tmpNSFee) Then
        txtAnaesActFees1.ForeColor = &HFF&
    Else
        txtAnaesActFees1.ForeColor = &H80000008
    End If
End Sub

Private Sub txtAnaesActFees1_KeyPress(KeyAscii As Integer)
    KeyAscii = KeyNumber(KeyAscii)
End Sub

Private Sub txtAnaesActFees2_Change()
Dim a, B, C, d, e, f, g, h, i As Double

    a = IIf(IsNumeric(txtAnaesActFees1), txtAnaesActFees1, 0)
    B = IIf(IsNumeric(txtAnaesActFees2), txtAnaesActFees2, 0)
    C = IIf(IsNumeric(txtAnaesActFees3), txtAnaesActFees3, 0)
    d = IIf(IsNumeric(txtAnaesActFees4), txtAnaesActFees4, 0)
    e = IIf(IsNumeric(txtAnaesActFees5), txtAnaesActFees5, 0)
    g = IIf(IsNumeric(txtAnaesActFees6), txtAnaesActFees6, 0)
    h = IIf(IsNumeric(txtAnaesActFees7), txtAnaesActFees7, 0)
    i = IIf(IsNumeric(txtAnaesActFees8), txtAnaesActFees8, 0)
    
    f = CDbl(a) + CDbl(B) + CDbl(C) + CDbl(d) + CDbl(e) + CDbl(g) + CDbl(h) + CDbl(i)
    txtTotalAnaesActFees = f

    tmpNSFee = IIf(IsNumeric(txtAnae13Rate2), txtAnae13Rate2, 0)
    
    If CDbl(a) > CDbl(tmpNSFee) Then
        txtAnaesActFees2.ForeColor = &HFF&
    Else
        txtAnaesActFees2.ForeColor = &H80000008
    End If
    
End Sub

Private Sub txtAnaesActFees2_KeyPress(KeyAscii As Integer)
    KeyAscii = KeyNumber(KeyAscii)
End Sub

Private Sub txtAnaesActFees3_Change()
Dim a, B, C, d, e, f, g, h, i As Double

    a = IIf(IsNumeric(txtAnaesActFees1), txtAnaesActFees1, 0)
    B = IIf(IsNumeric(txtAnaesActFees2), txtAnaesActFees2, 0)
    C = IIf(IsNumeric(txtAnaesActFees3), txtAnaesActFees3, 0)
    d = IIf(IsNumeric(txtAnaesActFees4), txtAnaesActFees4, 0)
    e = IIf(IsNumeric(txtAnaesActFees5), txtAnaesActFees5, 0)
    g = IIf(IsNumeric(txtAnaesActFees6), txtAnaesActFees6, 0)
    h = IIf(IsNumeric(txtAnaesActFees7), txtAnaesActFees7, 0)
    i = IIf(IsNumeric(txtAnaesActFees8), txtAnaesActFees8, 0)
    
    f = CDbl(a) + CDbl(B) + CDbl(C) + CDbl(d) + CDbl(e) + CDbl(g) + CDbl(h) + CDbl(i)
    txtTotalAnaesActFees = f

    tmpNSFee = IIf(IsNumeric(txtAnae13Rate3), txtAnae13Rate3, 0)
    
    If CDbl(a) > CDbl(tmpNSFee) Then
        txtAnaesActFees3.ForeColor = &HFF&
    Else
        txtAnaesActFees3.ForeColor = &H80000008
    End If
End Sub

Private Sub txtAnaesActFees3_KeyPress(KeyAscii As Integer)
    KeyAscii = KeyNumber(KeyAscii)
End Sub

Private Sub txtAnaesActFees4_Change()
Dim a, B, C, d, e, f, g, h, i As Double

    a = IIf(IsNumeric(txtAnaesActFees1), txtAnaesActFees1, 0)
    B = IIf(IsNumeric(txtAnaesActFees2), txtAnaesActFees2, 0)
    C = IIf(IsNumeric(txtAnaesActFees3), txtAnaesActFees3, 0)
    d = IIf(IsNumeric(txtAnaesActFees4), txtAnaesActFees4, 0)
    e = IIf(IsNumeric(txtAnaesActFees5), txtAnaesActFees5, 0)
    g = IIf(IsNumeric(txtAnaesActFees6), txtAnaesActFees6, 0)
    h = IIf(IsNumeric(txtAnaesActFees7), txtAnaesActFees7, 0)
    i = IIf(IsNumeric(txtAnaesActFees8), txtAnaesActFees8, 0)
    
    f = CDbl(a) + CDbl(B) + CDbl(C) + CDbl(d) + CDbl(e) + CDbl(g) + CDbl(h) + CDbl(i)
    txtTotalAnaesActFees = f

    tmpNSFee = IIf(IsNumeric(txtAnae13Rate4), txtAnae13Rate4, 0)
    
    If CDbl(a) > CDbl(tmpNSFee) Then
        txtAnaesActFees4.ForeColor = &HFF&
    Else
        txtAnaesActFees4.ForeColor = &H80000008
    End If
    
End Sub

Private Sub txtAnaesActFees4_KeyPress(KeyAscii As Integer)
    KeyAscii = KeyNumber(KeyAscii)
End Sub

Private Sub txtAnaesActFees5_Change()
Dim a, B, C, d, e, f, g, h, i As Double

    a = IIf(IsNumeric(txtAnaesActFees1), txtAnaesActFees1, 0)
    B = IIf(IsNumeric(txtAnaesActFees2), txtAnaesActFees2, 0)
    C = IIf(IsNumeric(txtAnaesActFees3), txtAnaesActFees3, 0)
    d = IIf(IsNumeric(txtAnaesActFees4), txtAnaesActFees4, 0)
    e = IIf(IsNumeric(txtAnaesActFees5), txtAnaesActFees5, 0)
    g = IIf(IsNumeric(txtAnaesActFees6), txtAnaesActFees6, 0)
    h = IIf(IsNumeric(txtAnaesActFees7), txtAnaesActFees7, 0)
    i = IIf(IsNumeric(txtAnaesActFees8), txtAnaesActFees8, 0)
    
    f = CDbl(a) + CDbl(B) + CDbl(C) + CDbl(d) + CDbl(e) + CDbl(g) + CDbl(h) + CDbl(i)
    txtTotalAnaesActFees = f

    tmpNSFee = IIf(IsNumeric(txtAnae13Rate5), txtAnae13Rate5, 0)
    
    If CDbl(a) > CDbl(tmpNSFee) Then
        txtAnaesActFees5.ForeColor = &HFF&
    Else
        txtAnaesActFees5.ForeColor = &H80000008
    End If
    
End Sub

Private Sub txtAnaesActFees5_KeyPress(KeyAscii As Integer)
    KeyAscii = KeyNumber(KeyAscii)
End Sub

Private Sub txtAnaesActFees6_Change()
Dim a, B, C, d, e, f, g, h, i As Double

    a = IIf(IsNumeric(txtAnaesActFees1), txtAnaesActFees1, 0)
    B = IIf(IsNumeric(txtAnaesActFees2), txtAnaesActFees2, 0)
    C = IIf(IsNumeric(txtAnaesActFees3), txtAnaesActFees3, 0)
    d = IIf(IsNumeric(txtAnaesActFees4), txtAnaesActFees4, 0)
    e = IIf(IsNumeric(txtAnaesActFees5), txtAnaesActFees5, 0)
    g = IIf(IsNumeric(txtAnaesActFees6), txtAnaesActFees6, 0)
    h = IIf(IsNumeric(txtAnaesActFees7), txtAnaesActFees7, 0)
    i = IIf(IsNumeric(txtAnaesActFees8), txtAnaesActFees8, 0)
    
    f = CDbl(a) + CDbl(B) + CDbl(C) + CDbl(d) + CDbl(e) + CDbl(g) + CDbl(h) + CDbl(i)
    txtTotalAnaesActFees = f

    tmpNSFee = IIf(IsNumeric(txtAnae13Rate6), txtAnae13Rate6, 0)
    
    If CDbl(a) > CDbl(tmpNSFee) Then
        txtAnaesActFees6.ForeColor = &HFF&
    Else
        txtAnaesActFees6.ForeColor = &H80000008
    End If

End Sub

Private Sub txtAnaesActFees6_KeyPress(KeyAscii As Integer)
    KeyAscii = KeyNumber(KeyAscii)
End Sub

Private Sub txtAnaesActFees7_Change()
Dim a, B, C, d, e, f, g, h, i As Double

    a = IIf(IsNumeric(txtAnaesActFees1), txtAnaesActFees1, 0)
    B = IIf(IsNumeric(txtAnaesActFees2), txtAnaesActFees2, 0)
    C = IIf(IsNumeric(txtAnaesActFees3), txtAnaesActFees3, 0)
    d = IIf(IsNumeric(txtAnaesActFees4), txtAnaesActFees4, 0)
    e = IIf(IsNumeric(txtAnaesActFees5), txtAnaesActFees5, 0)
    g = IIf(IsNumeric(txtAnaesActFees6), txtAnaesActFees6, 0)
    h = IIf(IsNumeric(txtAnaesActFees7), txtAnaesActFees7, 0)
    i = IIf(IsNumeric(txtAnaesActFees8), txtAnaesActFees8, 0)
    
    f = CDbl(a) + CDbl(B) + CDbl(C) + CDbl(d) + CDbl(e) + CDbl(g) + CDbl(h) + CDbl(i)
    txtTotalAnaesActFees = f

    tmpNSFee = IIf(IsNumeric(txtAnae13Rate7), txtAnae13Rate7, 0)
    
    If CDbl(a) > CDbl(tmpNSFee) Then
        txtAnaesActFees7.ForeColor = &HFF&
    Else
        txtAnaesActFees7.ForeColor = &H80000008
    End If
End Sub

Private Sub txtAnaesActFees8_Change()
Dim a, B, C, d, e, f, g, h, i As Double

    a = IIf(IsNumeric(txtAnaesActFees1), txtAnaesActFees1, 0)
    B = IIf(IsNumeric(txtAnaesActFees2), txtAnaesActFees2, 0)
    C = IIf(IsNumeric(txtAnaesActFees3), txtAnaesActFees3, 0)
    d = IIf(IsNumeric(txtAnaesActFees4), txtAnaesActFees4, 0)
    e = IIf(IsNumeric(txtAnaesActFees5), txtAnaesActFees5, 0)
    g = IIf(IsNumeric(txtAnaesActFees6), txtAnaesActFees6, 0)
    h = IIf(IsNumeric(txtAnaesActFees7), txtAnaesActFees7, 0)
    i = IIf(IsNumeric(txtAnaesActFees8), txtAnaesActFees8, 0)
    
    f = CDbl(a) + CDbl(B) + CDbl(C) + CDbl(d) + CDbl(e) + CDbl(g) + CDbl(h) + CDbl(i)
    txtTotalAnaesActFees = f

    tmpNSFee = IIf(IsNumeric(txtAnae13Rate8), txtAnae13Rate8, 0)
    
    If CDbl(a) > CDbl(tmpNSFee) Then
        txtAnaesActFees8.ForeColor = &HFF&
    Else
        txtAnaesActFees8.ForeColor = &H80000008
    End If
End Sub

Private Sub txtAnaesMMARate1_Change()
Dim a, B, C, d, e, f, g As Double

    a = IIf(IsNumeric(txtAnaesMMARate1), txtAnaesMMARate1, 0)
    B = IIf(IsNumeric(txtAnaesMMARate2), txtAnaesMMARate2, 0)
    C = IIf(IsNumeric(txtAnaesMMARate3), txtAnaesMMARate3, 0)
    d = IIf(IsNumeric(txtAnaesMMARate4), txtAnaesMMARate4, 0)
    e = IIf(IsNumeric(txtAnaesMMARate5), txtAnaesMMARate5, 0)
    g = IIf(IsNumeric(txtAnaesMMARate6), txtAnaesMMARate6, 0)
    
    f = CDbl(a) + CDbl(B) + CDbl(C) + CDbl(d) + CDbl(e) + CDbl(g)
    txtTotalAnaesMMARate = f
End Sub

Private Sub txtAnaesMMARate1_KeyPress(KeyAscii As Integer)
    KeyAscii = KeyNumber(KeyAscii)
End Sub

Private Sub txtAnaesMMARate2_Change()
Dim a, B, C, d, e, f, g As Double

    a = IIf(IsNumeric(txtAnaesMMARate1), txtAnaesMMARate1, 0)
    B = IIf(IsNumeric(txtAnaesMMARate2), txtAnaesMMARate2, 0)
    C = IIf(IsNumeric(txtAnaesMMARate3), txtAnaesMMARate3, 0)
    d = IIf(IsNumeric(txtAnaesMMARate4), txtAnaesMMARate4, 0)
    e = IIf(IsNumeric(txtAnaesMMARate5), txtAnaesMMARate5, 0)
    g = IIf(IsNumeric(txtAnaesMMARate6), txtAnaesMMARate6, 0)
    
    f = CDbl(a) + CDbl(B) + CDbl(C) + CDbl(d) + CDbl(e) + CDbl(g)
    txtTotalAnaesMMARate = f
End Sub

Private Sub txtAnaesMMARate2_KeyPress(KeyAscii As Integer)
    KeyAscii = KeyNumber(KeyAscii)
End Sub

Private Sub txtAnaesMMARate3_Change()
Dim a, B, C, d, e, f, g As Double

    a = IIf(IsNumeric(txtAnaesMMARate1), txtAnaesMMARate1, 0)
    B = IIf(IsNumeric(txtAnaesMMARate2), txtAnaesMMARate2, 0)
    C = IIf(IsNumeric(txtAnaesMMARate3), txtAnaesMMARate3, 0)
    d = IIf(IsNumeric(txtAnaesMMARate4), txtAnaesMMARate4, 0)
    e = IIf(IsNumeric(txtAnaesMMARate5), txtAnaesMMARate5, 0)
    g = IIf(IsNumeric(txtAnaesMMARate6), txtAnaesMMARate6, 0)
    
    f = CDbl(a) + CDbl(B) + CDbl(C) + CDbl(d) + CDbl(e) + CDbl(g)
    txtTotalAnaesMMARate = f
End Sub

Private Sub txtAnaesMMARate3_KeyPress(KeyAscii As Integer)
    KeyAscii = KeyNumber(KeyAscii)
End Sub

Private Sub txtAnaesMMARate4_Change()
Dim a, B, C, d, e, f, g As Double

    a = IIf(IsNumeric(txtAnaesMMARate1), txtAnaesMMARate1, 0)
    B = IIf(IsNumeric(txtAnaesMMARate2), txtAnaesMMARate2, 0)
    C = IIf(IsNumeric(txtAnaesMMARate3), txtAnaesMMARate3, 0)
    d = IIf(IsNumeric(txtAnaesMMARate4), txtAnaesMMARate4, 0)
    e = IIf(IsNumeric(txtAnaesMMARate5), txtAnaesMMARate5, 0)
    g = IIf(IsNumeric(txtAnaesMMARate6), txtAnaesMMARate6, 0)
    
    f = CDbl(a) + CDbl(B) + CDbl(C) + CDbl(d) + CDbl(e) + CDbl(g)
    txtTotalAnaesMMARate = f
End Sub

Private Sub txtAnaesMMARate4_KeyPress(KeyAscii As Integer)
    KeyAscii = KeyNumber(KeyAscii)
End Sub

Private Sub txtAnaesMMARate5_Change()
Dim a, B, C, d, e, f, g As Double

    a = IIf(IsNumeric(txtAnaesMMARate1), txtAnaesMMARate1, 0)
    B = IIf(IsNumeric(txtAnaesMMARate2), txtAnaesMMARate2, 0)
    C = IIf(IsNumeric(txtAnaesMMARate3), txtAnaesMMARate3, 0)
    d = IIf(IsNumeric(txtAnaesMMARate4), txtAnaesMMARate4, 0)
    e = IIf(IsNumeric(txtAnaesMMARate5), txtAnaesMMARate5, 0)
    g = IIf(IsNumeric(txtAnaesMMARate6), txtAnaesMMARate6, 0)
    
    f = CDbl(a) + CDbl(B) + CDbl(C) + CDbl(d) + CDbl(e) + CDbl(g)
    txtTotalAnaesMMARate = f
End Sub

Private Sub txtAnaesMMARate5_KeyPress(KeyAscii As Integer)
    KeyAscii = KeyNumber(KeyAscii)
End Sub

Private Sub txtAnaesMMARate6_Change()
Dim a, B, C, d, e, f, g As Double

    a = IIf(IsNumeric(txtAnaesMMARate1), txtAnaesMMARate1, 0)
    B = IIf(IsNumeric(txtAnaesMMARate2), txtAnaesMMARate2, 0)
    C = IIf(IsNumeric(txtAnaesMMARate3), txtAnaesMMARate3, 0)
    d = IIf(IsNumeric(txtAnaesMMARate4), txtAnaesMMARate4, 0)
    e = IIf(IsNumeric(txtAnaesMMARate5), txtAnaesMMARate5, 0)
    g = IIf(IsNumeric(txtAnaesMMARate6), txtAnaesMMARate6, 0)
    
    f = CDbl(a) + CDbl(B) + CDbl(C) + CDbl(d) + CDbl(e) + CDbl(g)
    txtTotalAnaesMMARate = f
End Sub

Private Sub txtNS13MinRate1_Change()
Dim a, B, C, d, e, f, g As Double

    a = IIf(IsNumeric(txtNS13MinRate1), txtNS13MinRate1, 0)
    B = IIf(IsNumeric(txtNS13MinRate2), txtNS13MinRate2, 0)
    C = IIf(IsNumeric(txtNS13MinRate3), txtNS13MinRate3, 0)
    d = IIf(IsNumeric(txtNS13MinRate4), txtNS13MinRate4, 0)
    e = IIf(IsNumeric(txtNS13MinRate5), txtNS13MinRate5, 0)
    g = IIf(IsNumeric(txtNS13MinRate6), txtNS13MinRate6, 0)
    
    f = CDbl(a) + CDbl(B) + CDbl(C) + CDbl(d) + CDbl(e) + CDbl(g)
    txtTotalNS13MinRate = f

End Sub

Private Sub txtNS13MinRate2_Change()
Dim a, B, C, d, e, f, g As Double

    a = IIf(IsNumeric(txtNS13MinRate1), txtNS13MinRate1, 0)
    B = IIf(IsNumeric(txtNS13MinRate2), txtNS13MinRate2, 0)
    C = IIf(IsNumeric(txtNS13MinRate3), txtNS13MinRate3, 0)
    d = IIf(IsNumeric(txtNS13MinRate4), txtNS13MinRate4, 0)
    e = IIf(IsNumeric(txtNS13MinRate5), txtNS13MinRate5, 0)
    g = IIf(IsNumeric(txtNS13MinRate6), txtNS13MinRate6, 0)
    
    f = CDbl(a) + CDbl(B) + CDbl(C) + CDbl(d) + CDbl(e) + CDbl(g)
    txtTotalNS13MinRate = f

End Sub

Private Sub txtNS13MinRate3_Change()
Dim a, B, C, d, e, f, g As Double

    a = IIf(IsNumeric(txtNS13MinRate1), txtNS13MinRate1, 0)
    B = IIf(IsNumeric(txtNS13MinRate2), txtNS13MinRate2, 0)
    C = IIf(IsNumeric(txtNS13MinRate3), txtNS13MinRate3, 0)
    d = IIf(IsNumeric(txtNS13MinRate4), txtNS13MinRate4, 0)
    e = IIf(IsNumeric(txtNS13MinRate5), txtNS13MinRate5, 0)
    g = IIf(IsNumeric(txtNS13MinRate6), txtNS13MinRate6, 0)
    
    f = CDbl(a) + CDbl(B) + CDbl(C) + CDbl(d) + CDbl(e) + CDbl(g)
    txtTotalNS13MinRate = f

End Sub

Private Sub txtNS13MinRate4_Change()
Dim a, B, C, d, e, f, g As Double

    a = IIf(IsNumeric(txtNS13MinRate1), txtNS13MinRate1, 0)
    B = IIf(IsNumeric(txtNS13MinRate2), txtNS13MinRate2, 0)
    C = IIf(IsNumeric(txtNS13MinRate3), txtNS13MinRate3, 0)
    d = IIf(IsNumeric(txtNS13MinRate4), txtNS13MinRate4, 0)
    e = IIf(IsNumeric(txtNS13MinRate5), txtNS13MinRate5, 0)
    g = IIf(IsNumeric(txtNS13MinRate6), txtNS13MinRate6, 0)
    
    f = CDbl(a) + CDbl(B) + CDbl(C) + CDbl(d) + CDbl(e) + CDbl(g)
    txtTotalNS13MinRate = f

End Sub

Private Sub txtNS13MinRate5_Change()
Dim a, B, C, d, e, f, g As Double

    a = IIf(IsNumeric(txtNS13MinRate1), txtNS13MinRate1, 0)
    B = IIf(IsNumeric(txtNS13MinRate2), txtNS13MinRate2, 0)
    C = IIf(IsNumeric(txtNS13MinRate3), txtNS13MinRate3, 0)
    d = IIf(IsNumeric(txtNS13MinRate4), txtNS13MinRate4, 0)
    e = IIf(IsNumeric(txtNS13MinRate5), txtNS13MinRate5, 0)
    g = IIf(IsNumeric(txtNS13MinRate6), txtNS13MinRate6, 0)
    
    f = CDbl(a) + CDbl(B) + CDbl(C) + CDbl(d) + CDbl(e) + CDbl(g)
    txtTotalNS13MinRate = f

End Sub

Private Sub txtNS13MinRate6_Change()
Dim a, B, C, d, e, f, g As Double

    a = IIf(IsNumeric(txtNS13MinRate1), txtNS13MinRate1, 0)
    B = IIf(IsNumeric(txtNS13MinRate2), txtNS13MinRate2, 0)
    C = IIf(IsNumeric(txtNS13MinRate3), txtNS13MinRate3, 0)
    d = IIf(IsNumeric(txtNS13MinRate4), txtNS13MinRate4, 0)
    e = IIf(IsNumeric(txtNS13MinRate5), txtNS13MinRate5, 0)
    g = IIf(IsNumeric(txtNS13MinRate6), txtNS13MinRate6, 0)
    
    f = CDbl(a) + CDbl(B) + CDbl(C) + CDbl(d) + CDbl(e) + CDbl(g)
    txtTotalNS13MinRate = f

End Sub

Private Sub txtNS13Rate1_Change()
Dim a, B, C, d, e, f, g As Double

    a = IIf(IsNumeric(txtNS13Rate1), txtNS13Rate1, 0)
    B = IIf(IsNumeric(txtNS13Rate2), txtNS13Rate2, 0)
    C = IIf(IsNumeric(txtNS13Rate3), txtNS13Rate3, 0)
    d = IIf(IsNumeric(txtNS13Rate4), txtNS13Rate4, 0)
    e = IIf(IsNumeric(txtNS13Rate5), txtNS13Rate5, 0)
    g = IIf(IsNumeric(txtNS13Rate6), txtNS13Rate6, 0)
    
    f = CDbl(a) + CDbl(B) + CDbl(C) + CDbl(d) + CDbl(e) + CDbl(g)
    txtTotalNS13Rate = f

    tmpNSFee = IIf(IsNumeric(txtNSActFee1), txtNSActFee1, 0)
    
    If CDbl(tmpNSFee) > CDbl(a) Then
        txtNSActFee1.ForeColor = &HFF&
    Else
        txtNSActFee1.ForeColor = &H80000008
    End If

End Sub

Private Sub txtNS13Rate2_Change()
Dim a, B, C, d, e, f, g As Double

    a = IIf(IsNumeric(txtNS13Rate1), txtNS13Rate1, 0)
    B = IIf(IsNumeric(txtNS13Rate2), txtNS13Rate2, 0)
    C = IIf(IsNumeric(txtNS13Rate3), txtNS13Rate3, 0)
    d = IIf(IsNumeric(txtNS13Rate4), txtNS13Rate4, 0)
    e = IIf(IsNumeric(txtNS13Rate5), txtNS13Rate5, 0)
    g = IIf(IsNumeric(txtNS13Rate6), txtNS13Rate6, 0)
    
    f = CDbl(a) + CDbl(B) + CDbl(C) + CDbl(d) + CDbl(e) + CDbl(g)
    txtTotalNS13Rate = f

    tmpNSFee = IIf(IsNumeric(txtNSActFee2), txtNSActFee2, 0)
    
    If CDbl(tmpNSFee) > CDbl(B) Then
        txtNSActFee2.ForeColor = &HFF&
    Else
        txtNSActFee2.ForeColor = &H80000008
    End If

End Sub

Private Sub txtNS13Rate3_Change()
Dim a, B, C, d, e, f, g As Double

    a = IIf(IsNumeric(txtNS13Rate1), txtNS13Rate1, 0)
    B = IIf(IsNumeric(txtNS13Rate2), txtNS13Rate2, 0)
    C = IIf(IsNumeric(txtNS13Rate3), txtNS13Rate3, 0)
    d = IIf(IsNumeric(txtNS13Rate4), txtNS13Rate4, 0)
    e = IIf(IsNumeric(txtNS13Rate5), txtNS13Rate5, 0)
    g = IIf(IsNumeric(txtNS13Rate6), txtNS13Rate6, 0)
    
    f = CDbl(a) + CDbl(B) + CDbl(C) + CDbl(d) + CDbl(e) + CDbl(g)
    txtTotalNS13Rate = f

    tmpNSFee = IIf(IsNumeric(txtNSActFee3), txtNSActFee3, 0)
    
    If CDbl(tmpNSFee) > CDbl(C) Then
        txtNSActFee3.ForeColor = &HFF&
    Else
        txtNSActFee3.ForeColor = &H80000008
    End If

End Sub

Private Sub txtNS13Rate4_Change()
Dim a, B, C, d, e, f, g As Double

    a = IIf(IsNumeric(txtNS13Rate1), txtNS13Rate1, 0)
    B = IIf(IsNumeric(txtNS13Rate2), txtNS13Rate2, 0)
    C = IIf(IsNumeric(txtNS13Rate3), txtNS13Rate3, 0)
    d = IIf(IsNumeric(txtNS13Rate4), txtNS13Rate4, 0)
    e = IIf(IsNumeric(txtNS13Rate5), txtNS13Rate5, 0)
    g = IIf(IsNumeric(txtNS13Rate6), txtNS13Rate6, 0)
    
    f = CDbl(a) + CDbl(B) + CDbl(C) + CDbl(d) + CDbl(e) + CDbl(g)
    txtTotalNS13Rate = f

    tmpNSFee = IIf(IsNumeric(txtNSActFee4), txtNSActFee4, 0)
    
    If CDbl(tmpNSFee) > CDbl(d) Then
        txtNSActFee4.ForeColor = &HFF&
    Else
        txtNSActFee4.ForeColor = &H80000008
    End If

End Sub

Private Sub txtNS13Rate5_Change()
Dim a, B, C, d, e, f, g As Double

    a = IIf(IsNumeric(txtNS13Rate1), txtNS13Rate1, 0)
    B = IIf(IsNumeric(txtNS13Rate2), txtNS13Rate2, 0)
    C = IIf(IsNumeric(txtNS13Rate3), txtNS13Rate3, 0)
    d = IIf(IsNumeric(txtNS13Rate4), txtNS13Rate4, 0)
    e = IIf(IsNumeric(txtNS13Rate5), txtNS13Rate5, 0)
    g = IIf(IsNumeric(txtNS13Rate6), txtNS13Rate6, 0)
    
    f = CDbl(a) + CDbl(B) + CDbl(C) + CDbl(d) + CDbl(e) + CDbl(g)
    txtTotalNS13Rate = f

    tmpNSFee = IIf(IsNumeric(txtNSActFee5), txtNSActFee5, 0)
    
    If CDbl(tmpNSFee) > CDbl(e) Then
        txtNSActFee5.ForeColor = &HFF&
    Else
        txtNSActFee5.ForeColor = &H80000008
    End If

End Sub

Private Sub txtNS13Rate6_Change()
Dim a, B, C, d, e, f, g As Double

    a = IIf(IsNumeric(txtNS13Rate1), txtNS13Rate1, 0)
    B = IIf(IsNumeric(txtNS13Rate2), txtNS13Rate2, 0)
    C = IIf(IsNumeric(txtNS13Rate3), txtNS13Rate3, 0)
    d = IIf(IsNumeric(txtNS13Rate4), txtNS13Rate4, 0)
    e = IIf(IsNumeric(txtNS13Rate5), txtNS13Rate5, 0)
    g = IIf(IsNumeric(txtNS13Rate6), txtNS13Rate6, 0)
    
    f = CDbl(a) + CDbl(B) + CDbl(C) + CDbl(d) + CDbl(e) + CDbl(g)
    txtTotalNS13Rate = f

    tmpNSFee = IIf(IsNumeric(txtNSActFee6), txtNSActFee6, 0)
    
    If CDbl(tmpNSFee) > CDbl(g) Then
        txtNSActFee6.ForeColor = &HFF&
    Else
        txtNSActFee6.ForeColor = &H80000008
    End If
End Sub

Private Sub txtNSActFee1_Change()
Dim a, B, C, d, e, f, g, h, i As Double
Dim tmpNSFee As Double

    a = IIf(IsNumeric(txtNSActFee1), txtNSActFee1, 0)
    B = IIf(IsNumeric(txtNSActFee2), txtNSActFee2, 0)
    C = IIf(IsNumeric(txtNSActFee3), txtNSActFee3, 0)
    d = IIf(IsNumeric(txtNSActFee4), txtNSActFee4, 0)
    e = IIf(IsNumeric(txtNSActFee5), txtNSActFee5, 0)
    g = IIf(IsNumeric(txtNSActFee6), txtNSActFee6, 0)
    h = IIf(IsNumeric(txtNSActFee7), txtNSActFee7, 0)
    i = IIf(IsNumeric(txtNSActFee8), txtNSActFee8, 0)
    
    f = CDbl(a) + CDbl(B) + CDbl(C) + CDbl(d) + CDbl(e) + CDbl(g) + CDbl(h) + CDbl(i)
    txtTotalNSActFee = f
    
    tmpNSFee = IIf(IsNumeric(txtNS13Rate1), txtNS13Rate1, 0)
    
    If CDbl(a) > CDbl(tmpNSFee) Then
        txtNSActFee1.ForeColor = &HFF&
    Else
        txtNSActFee1.ForeColor = &H80000008
    End If
End Sub

Private Sub txtNSActFee1_KeyPress(KeyAscii As Integer)
    KeyAscii = KeyNumber(KeyAscii)
End Sub

Private Sub txtNSActFee2_Change()
Dim a, B, C, d, e, f, g, h, i As Double
Dim tmpNSFee As Double

    a = IIf(IsNumeric(txtNSActFee1), txtNSActFee1, 0)
    B = IIf(IsNumeric(txtNSActFee2), txtNSActFee2, 0)
    C = IIf(IsNumeric(txtNSActFee3), txtNSActFee3, 0)
    d = IIf(IsNumeric(txtNSActFee4), txtNSActFee4, 0)
    e = IIf(IsNumeric(txtNSActFee5), txtNSActFee5, 0)
    g = IIf(IsNumeric(txtNSActFee6), txtNSActFee6, 0)
    h = IIf(IsNumeric(txtNSActFee7), txtNSActFee7, 0)
    i = IIf(IsNumeric(txtNSActFee8), txtNSActFee8, 0)
    
    f = CDbl(a) + CDbl(B) + CDbl(C) + CDbl(d) + CDbl(e) + CDbl(g) + CDbl(h) + CDbl(i)
    txtTotalNSActFee = f
    
    tmpNSFee = IIf(IsNumeric(txtNS13Rate2), txtNS13Rate2, 0)
    
    If CDbl(a) > CDbl(tmpNSFee) Then
        txtNSActFee2.ForeColor = &HFF&
    Else
        txtNSActFee2.ForeColor = &H80000008
    End If
    
End Sub

Private Sub txtNSActFee2_KeyPress(KeyAscii As Integer)
    KeyAscii = KeyNumber(KeyAscii)
End Sub

Private Sub txtNSActFee3_Change()
Dim a, B, C, d, e, f, g, h, i As Double
Dim tmpNSFee As Double

    a = IIf(IsNumeric(txtNSActFee1), txtNSActFee1, 0)
    B = IIf(IsNumeric(txtNSActFee2), txtNSActFee2, 0)
    C = IIf(IsNumeric(txtNSActFee3), txtNSActFee3, 0)
    d = IIf(IsNumeric(txtNSActFee4), txtNSActFee4, 0)
    e = IIf(IsNumeric(txtNSActFee5), txtNSActFee5, 0)
    g = IIf(IsNumeric(txtNSActFee6), txtNSActFee6, 0)
    h = IIf(IsNumeric(txtNSActFee7), txtNSActFee7, 0)
    i = IIf(IsNumeric(txtNSActFee8), txtNSActFee8, 0)
    
    f = CDbl(a) + CDbl(B) + CDbl(C) + CDbl(d) + CDbl(e) + CDbl(g) + CDbl(h) + CDbl(i)
    txtTotalNSActFee = f
    
    tmpNSFee = IIf(IsNumeric(txtNS13Rate3), txtNS13Rate3, 0)
    
    If CDbl(a) > CDbl(tmpNSFee) Then
        txtNSActFee3.ForeColor = &HFF&
    Else
        txtNSActFee3.ForeColor = &H80000008
    End If
End Sub

Private Sub txtNSActFee3_KeyPress(KeyAscii As Integer)
    KeyAscii = KeyNumber(KeyAscii)
End Sub

Private Sub txtNSActFee4_Change()
Dim a, B, C, d, e, f, g, h, i As Double
Dim tmpNSFee As Double

    a = IIf(IsNumeric(txtNSActFee1), txtNSActFee1, 0)
    B = IIf(IsNumeric(txtNSActFee2), txtNSActFee2, 0)
    C = IIf(IsNumeric(txtNSActFee3), txtNSActFee3, 0)
    d = IIf(IsNumeric(txtNSActFee4), txtNSActFee4, 0)
    e = IIf(IsNumeric(txtNSActFee5), txtNSActFee5, 0)
    g = IIf(IsNumeric(txtNSActFee6), txtNSActFee6, 0)
    h = IIf(IsNumeric(txtNSActFee7), txtNSActFee7, 0)
    i = IIf(IsNumeric(txtNSActFee8), txtNSActFee8, 0)
    
    f = CDbl(a) + CDbl(B) + CDbl(C) + CDbl(d) + CDbl(e) + CDbl(g) + CDbl(h) + CDbl(i)
    txtTotalNSActFee = f
    
    tmpNSFee = IIf(IsNumeric(txtNS13Rate4), txtNS13Rate4, 0)
    
    If CDbl(a) > CDbl(tmpNSFee) Then
        txtNSActFee4.ForeColor = &HFF&
    Else
        txtNSActFee4.ForeColor = &H80000008
    End If
End Sub

Private Sub txtNSActFee4_KeyPress(KeyAscii As Integer)
    KeyAscii = KeyNumber(KeyAscii)
End Sub

Private Sub txtNSActFee5_Change()
Dim a, B, C, d, e, f, g, h, i As Double
Dim tmpNSFee As Double

    a = IIf(IsNumeric(txtNSActFee1), txtNSActFee1, 0)
    B = IIf(IsNumeric(txtNSActFee2), txtNSActFee2, 0)
    C = IIf(IsNumeric(txtNSActFee3), txtNSActFee3, 0)
    d = IIf(IsNumeric(txtNSActFee4), txtNSActFee4, 0)
    e = IIf(IsNumeric(txtNSActFee5), txtNSActFee5, 0)
    g = IIf(IsNumeric(txtNSActFee6), txtNSActFee6, 0)
    h = IIf(IsNumeric(txtNSActFee7), txtNSActFee7, 0)
    i = IIf(IsNumeric(txtNSActFee8), txtNSActFee8, 0)
    
    f = CDbl(a) + CDbl(B) + CDbl(C) + CDbl(d) + CDbl(e) + CDbl(g) + CDbl(h) + CDbl(i)
    txtTotalNSActFee = f
    
    tmpNSFee = IIf(IsNumeric(txtNS13Rate5), txtNS13Rate5, 0)
    
    If CDbl(a) > CDbl(tmpNSFee) Then
        txtNSActFee5.ForeColor = &HFF&
    Else
        txtNSActFee5.ForeColor = &H80000008
    End If
    
End Sub

Private Sub txtNSActFee5_KeyPress(KeyAscii As Integer)
    KeyAscii = KeyNumber(KeyAscii)
End Sub

Private Sub txtNSActFee6_Change()
Dim a, B, C, d, e, f, g, h, i As Double
Dim tmpNSFee As Double

    a = IIf(IsNumeric(txtNSActFee1), txtNSActFee1, 0)
    B = IIf(IsNumeric(txtNSActFee2), txtNSActFee2, 0)
    C = IIf(IsNumeric(txtNSActFee3), txtNSActFee3, 0)
    d = IIf(IsNumeric(txtNSActFee4), txtNSActFee4, 0)
    e = IIf(IsNumeric(txtNSActFee5), txtNSActFee5, 0)
    g = IIf(IsNumeric(txtNSActFee6), txtNSActFee6, 0)
    h = IIf(IsNumeric(txtNSActFee7), txtNSActFee7, 0)
    i = IIf(IsNumeric(txtNSActFee8), txtNSActFee8, 0)
    
    f = CDbl(a) + CDbl(B) + CDbl(C) + CDbl(d) + CDbl(e) + CDbl(g) + CDbl(h) + CDbl(i)
    txtTotalNSActFee = f
    
    tmpNSFee = IIf(IsNumeric(txtNS13Rate6), txtNS13Rate6, 0)
    
    If CDbl(a) > CDbl(tmpNSFee) Then
        txtNSActFee6.ForeColor = &HFF&
    Else
        txtNSActFee6.ForeColor = &H80000008
    End If
    
End Sub

Private Sub txtNSActFee6_KeyPress(KeyAscii As Integer)
    KeyAscii = KeyNumber(KeyAscii)
End Sub

Private Sub txtNSActFee7_Change()
Dim a, B, C, d, e, f, g, h, i As Double
Dim tmpNSFee As Double

    a = IIf(IsNumeric(txtNSActFee1), txtNSActFee1, 0)
    B = IIf(IsNumeric(txtNSActFee2), txtNSActFee2, 0)
    C = IIf(IsNumeric(txtNSActFee3), txtNSActFee3, 0)
    d = IIf(IsNumeric(txtNSActFee4), txtNSActFee4, 0)
    e = IIf(IsNumeric(txtNSActFee5), txtNSActFee5, 0)
    g = IIf(IsNumeric(txtNSActFee6), txtNSActFee6, 0)
    h = IIf(IsNumeric(txtNSActFee7), txtNSActFee7, 0)
    i = IIf(IsNumeric(txtNSActFee8), txtNSActFee8, 0)
    
    f = CDbl(a) + CDbl(B) + CDbl(C) + CDbl(d) + CDbl(e) + CDbl(g) + CDbl(h) + CDbl(i)
    txtTotalNSActFee = f
    
    tmpNSFee = IIf(IsNumeric(txtNS13Rate7), txtNS13Rate7, 0)
    
    If CDbl(a) > CDbl(tmpNSFee) Then
        txtNSActFee7.ForeColor = &HFF&
    Else
        txtNSActFee7.ForeColor = &H80000008
    End If
End Sub

Private Sub txtNSActFee8_Change()
Dim a, B, C, d, e, f, g, h, i As Double
Dim tmpNSFee As Double

    a = IIf(IsNumeric(txtNSActFee1), txtNSActFee1, 0)
    B = IIf(IsNumeric(txtNSActFee2), txtNSActFee2, 0)
    C = IIf(IsNumeric(txtNSActFee3), txtNSActFee3, 0)
    d = IIf(IsNumeric(txtNSActFee4), txtNSActFee4, 0)
    e = IIf(IsNumeric(txtNSActFee5), txtNSActFee5, 0)
    g = IIf(IsNumeric(txtNSActFee6), txtNSActFee6, 0)
    h = IIf(IsNumeric(txtNSActFee7), txtNSActFee7, 0)
    i = IIf(IsNumeric(txtNSActFee8), txtNSActFee8, 0)
    
    f = CDbl(a) + CDbl(B) + CDbl(C) + CDbl(d) + CDbl(e) + CDbl(g) + CDbl(h) + CDbl(i)
    txtTotalNSActFee = f
    
    tmpNSFee = IIf(IsNumeric(txtNS13Rate8), txtNS13Rate8, 0)
    
    If CDbl(a) > CDbl(tmpNSFee) Then
        txtNSActFee8.ForeColor = &HFF&
    Else
        txtNSActFee8.ForeColor = &H80000008
    End If
End Sub

Private Sub txtNSMMARate1_Change()
Dim a, B, C, d, e, f, g As Double

    a = IIf(IsNumeric(txtNSMMARate1), txtNSMMARate1, 0)
    B = IIf(IsNumeric(txtNSMMARate2), txtNSMMARate2, 0)
    C = IIf(IsNumeric(txtNSMMARate3), txtNSMMARate3, 0)
    d = IIf(IsNumeric(txtNSMMARate4), txtNSMMARate4, 0)
    e = IIf(IsNumeric(txtNSMMARate5), txtNSMMARate5, 0)
    g = IIf(IsNumeric(txtNSMMARate6), txtNSMMARate6, 0)
    
    f = CDbl(a) + CDbl(B) + CDbl(C) + CDbl(d) + CDbl(e) + CDbl(g)
    txtTotalNSMMARate = f
End Sub

Private Sub txtNSMMARate1_KeyPress(KeyAscii As Integer)
    KeyAscii = KeyNumber(KeyAscii)
End Sub

Private Sub txtNSMMARate2_Change()
Dim a, B, C, d, e, f, g As Double

    a = IIf(IsNumeric(txtNSMMARate1), txtNSMMARate1, 0)
    B = IIf(IsNumeric(txtNSMMARate2), txtNSMMARate2, 0)
    C = IIf(IsNumeric(txtNSMMARate3), txtNSMMARate3, 0)
    d = IIf(IsNumeric(txtNSMMARate4), txtNSMMARate4, 0)
    e = IIf(IsNumeric(txtNSMMARate5), txtNSMMARate5, 0)
    g = IIf(IsNumeric(txtNSMMARate6), txtNSMMARate6, 0)
    
    f = CDbl(a) + CDbl(B) + CDbl(C) + CDbl(d) + CDbl(e) + CDbl(g)
    txtTotalNSMMARate = f
End Sub

Private Sub txtNSMMARate2_KeyPress(KeyAscii As Integer)
    KeyAscii = KeyNumber(KeyAscii)
End Sub

Private Sub txtNSMMARate3_Change()
Dim a, B, C, d, e, f, g As Double

    a = IIf(IsNumeric(txtNSMMARate1), txtNSMMARate1, 0)
    B = IIf(IsNumeric(txtNSMMARate2), txtNSMMARate2, 0)
    C = IIf(IsNumeric(txtNSMMARate3), txtNSMMARate3, 0)
    d = IIf(IsNumeric(txtNSMMARate4), txtNSMMARate4, 0)
    e = IIf(IsNumeric(txtNSMMARate5), txtNSMMARate5, 0)
    g = IIf(IsNumeric(txtNSMMARate6), txtNSMMARate6, 0)
    
    f = CDbl(a) + CDbl(B) + CDbl(C) + CDbl(d) + CDbl(e) + CDbl(g)
    txtTotalNSMMARate = f
End Sub

Private Sub txtNSMMARate3_KeyPress(KeyAscii As Integer)
    KeyAscii = KeyNumber(KeyAscii)
End Sub

Private Sub txtNSMMARate4_Change()
Dim a, B, C, d, e, f, g As Double

    a = IIf(IsNumeric(txtNSMMARate1), txtNSMMARate1, 0)
    B = IIf(IsNumeric(txtNSMMARate2), txtNSMMARate2, 0)
    C = IIf(IsNumeric(txtNSMMARate3), txtNSMMARate3, 0)
    d = IIf(IsNumeric(txtNSMMARate4), txtNSMMARate4, 0)
    e = IIf(IsNumeric(txtNSMMARate5), txtNSMMARate5, 0)
    g = IIf(IsNumeric(txtNSMMARate6), txtNSMMARate6, 0)
    
    f = CDbl(a) + CDbl(B) + CDbl(C) + CDbl(d) + CDbl(e) + CDbl(g)
    txtTotalNSMMARate = f
End Sub

Private Sub txtNSMMARate4_KeyPress(KeyAscii As Integer)
    KeyAscii = KeyNumber(KeyAscii)
End Sub

Private Sub txtNSMMARate5_Change()
Dim a, B, C, d, e, f, g As Double

    a = IIf(IsNumeric(txtNSMMARate1), txtNSMMARate1, 0)
    B = IIf(IsNumeric(txtNSMMARate2), txtNSMMARate2, 0)
    C = IIf(IsNumeric(txtNSMMARate3), txtNSMMARate3, 0)
    d = IIf(IsNumeric(txtNSMMARate4), txtNSMMARate4, 0)
    e = IIf(IsNumeric(txtNSMMARate5), txtNSMMARate5, 0)
    g = IIf(IsNumeric(txtNSMMARate6), txtNSMMARate6, 0)
    
    f = CDbl(a) + CDbl(B) + CDbl(C) + CDbl(d) + CDbl(e) + CDbl(g)
    txtTotalNSMMARate = f
End Sub

Private Sub txtNSMMARate5_KeyPress(KeyAscii As Integer)
    KeyAscii = KeyNumber(KeyAscii)
End Sub




Private Sub txtNSMMARate6_Change()
Dim a, B, C, d, e, f, g As Double

    a = IIf(IsNumeric(txtNSMMARate1), txtNSMMARate1, 0)
    B = IIf(IsNumeric(txtNSMMARate2), txtNSMMARate2, 0)
    C = IIf(IsNumeric(txtNSMMARate3), txtNSMMARate3, 0)
    d = IIf(IsNumeric(txtNSMMARate4), txtNSMMARate4, 0)
    e = IIf(IsNumeric(txtNSMMARate5), txtNSMMARate5, 0)
    g = IIf(IsNumeric(txtNSMMARate6), txtNSMMARate6, 0)
    
    f = CDbl(a) + CDbl(B) + CDbl(C) + CDbl(d) + CDbl(e) + CDbl(g)
    txtTotalNSMMARate = f
End Sub

Private Sub txtNSProcCode1_Change()
    If Len(Trim(txtNSProcCode1)) Then
        If CalSource = 0 Then
            Call SearchMMARateCal(txtNSProcCode1)
        Else
            Call SearchMMARate(txtNSProcCode1)
        End If
        frmMMAEntryGST.txtNSMMARate1 = IIf(Len(MMASurgFees), MMASurgFees, 0)
        frmMMAEntryGST.txtNS13Rate1 = IIf(Len(MMASurgFees13th), MMASurgFees13th, 0)
        frmMMAEntryGST.txtNS13MinRate1 = IIf(Len(MMASurgFees13thMin), MMASurgFees13thMin, 0)
        
        Call Search14thRate(txtNSProcCode1)
        frmMMAEntryGST.txtNS14Rate1 = IIf(Len(MMASurgFees13th), MMASurgFees13th, 0)
        frmMMAEntryGST.txtNS14MinRate1 = IIf(Len(MMASurgFees13thMin), MMASurgFees13thMin, 0)
    Else
        frmMMAEntryGST.txtNSMMARate1 = 0
        frmMMAEntryGST.txtNS13Rate1 = 0
        frmMMAEntryGST.txtNS13MinRate1 = 0
    End If
End Sub

Private Sub txtNSProcCode2_Change()
    If Len(Trim(txtNSProcCode2)) Then
        If CalSource = 0 Then
            Call SearchMMARateCal(txtNSProcCode2)
        Else
            Call SearchMMARate(txtNSProcCode2)
        End If
        frmMMAEntryGST.txtNSMMARate2 = IIf(Len(MMASurgFees), MMASurgFees, 0)
        frmMMAEntryGST.txtNS13Rate2 = IIf(Len(MMASurgFees13th), MMASurgFees13th, 0)
        frmMMAEntryGST.txtNS13MinRate2 = IIf(Len(MMASurgFees13thMin), MMASurgFees13thMin, 0)
        
        Call Search14thRate(txtNSProcCode2)
        frmMMAEntryGST.txtNS14Rate2 = IIf(Len(MMASurgFees13th), MMASurgFees13th, 0)
        frmMMAEntryGST.txtNS14MinRate2 = IIf(Len(MMASurgFees13thMin), MMASurgFees13thMin, 0)
        
    Else
        frmMMAEntryGST.txtNSMMARate2 = 0
        frmMMAEntryGST.txtNS13Rate2 = 0
        frmMMAEntryGST.txtNS13MinRate2 = 0
    End If
End Sub

Private Sub txtNSProcCode3_Change()
    If Len(Trim(txtNSProcCode3)) Then
        If CalSource = 0 Then
            Call SearchMMARateCal(txtNSProcCode3)
        Else
            Call SearchMMARate(txtNSProcCode3)
        End If
        frmMMAEntryGST.txtNSMMARate3 = IIf(Len(MMASurgFees), MMASurgFees, 0)
        frmMMAEntryGST.txtNS13Rate3 = IIf(Len(MMASurgFees13th), MMASurgFees13th, 0)
        frmMMAEntryGST.txtNS13MinRate3 = IIf(Len(MMASurgFees13thMin), MMASurgFees13thMin, 0)
        
        Call Search14thRate(txtNSProcCode3)
        frmMMAEntryGST.txtNS14Rate3 = IIf(Len(MMASurgFees13th), MMASurgFees13th, 0)
        frmMMAEntryGST.txtNS14MinRate3 = IIf(Len(MMASurgFees13thMin), MMASurgFees13thMin, 0)
        
    Else
        frmMMAEntryGST.txtNSMMARate3 = 0
        frmMMAEntryGST.txtNS13Rate3 = 0
        frmMMAEntryGST.txtNS13MinRate3 = 0
        
    End If
End Sub

Private Sub txtNSProcCode4_Change()
    If Len(Trim(txtNSProcCode4)) Then
        If CalSource = 0 Then
            Call SearchMMARateCal(txtNSProcCode4)
        Else
            Call SearchMMARate(txtNSProcCode4)
        End If
        frmMMAEntryGST.txtNSMMARate4 = IIf(Len(MMASurgFees), MMASurgFees, 0)
        frmMMAEntryGST.txtNS13Rate4 = IIf(Len(MMASurgFees13th), MMASurgFees13th, 0)
        frmMMAEntryGST.txtNS13MinRate4 = IIf(Len(MMASurgFees13thMin), MMASurgFees13thMin, 0)
        
        Call Search14thRate(txtNSProcCode4)
        frmMMAEntryGST.txtNS14Rate4 = IIf(Len(MMASurgFees13th), MMASurgFees13th, 0)
        frmMMAEntryGST.txtNS14MinRate4 = IIf(Len(MMASurgFees13thMin), MMASurgFees13thMin, 0)
        
    Else
        frmMMAEntryGST.txtNSMMARate4 = 0
        frmMMAEntryGST.txtNS13Rate4 = 0
        frmMMAEntryGST.txtNS13MinRate4 = 0
    End If
End Sub

Private Sub txtNSProcCode5_Change()
    If Len(Trim(txtNSProcCode5)) Then
        If CalSource = 0 Then
            Call SearchMMARateCal(txtNSProcCode5)
        Else
            Call SearchMMARate(txtNSProcCode5)
        End If
        frmMMAEntryGST.txtNSMMARate5 = IIf(Len(MMASurgFees), MMASurgFees, 0)
        frmMMAEntryGST.txtNS13Rate5 = IIf(Len(MMASurgFees13th), MMASurgFees13th, 0)
        frmMMAEntryGST.txtNS13MinRate5 = IIf(Len(MMASurgFees13thMin), MMASurgFees13thMin, 0)
        
        Call Search14thRate(txtNSProcCode5)
        frmMMAEntryGST.txtNS14Rate5 = IIf(Len(MMASurgFees13th), MMASurgFees13th, 0)
        frmMMAEntryGST.txtNS14MinRate5 = IIf(Len(MMASurgFees13thMin), MMASurgFees13thMin, 0)
        
    Else
        frmMMAEntryGST.txtNSMMARate5 = 0
        frmMMAEntryGST.txtNS13Rate5 = 0
        frmMMAEntryGST.txtNS13MinRate5 = 0
        
    End If
End Sub

Private Sub txtNSProcCode6_Change()
    If Len(Trim(txtNSProcCode6)) Then
        If CalSource = 0 Then
            Call SearchMMARateCal(txtNSProcCode6)
        Else
            Call SearchMMARate(txtNSProcCode6)
        End If
        frmMMAEntryGST.txtNSMMARate6 = IIf(Len(MMASurgFees), MMASurgFees, 0)
        frmMMAEntryGST.txtNS13Rate6 = IIf(Len(MMASurgFees13th), MMASurgFees13th, 0)
        frmMMAEntryGST.txtNS13MinRate6 = IIf(Len(MMASurgFees13thMin), MMASurgFees13thMin, 0)
        
        Call Search14thRate(txtNSProcCode6)
        frmMMAEntryGST.txtNS14Rate6 = IIf(Len(MMASurgFees13th), MMASurgFees13th, 0)
        frmMMAEntryGST.txtNS14MinRate6 = IIf(Len(MMASurgFees13thMin), MMASurgFees13thMin, 0)
        
    Else
        frmMMAEntryGST.txtNSMMARate6 = 0
        frmMMAEntryGST.txtNS13Rate6 = 0
        frmMMAEntryGST.txtNS13MinRate6 = 0
    End If

End Sub

Private Sub txtProcCode1_Change()
    If Len(Trim(txtProcCode1)) Then
        If CalSource = 0 Then
            Call SearchMMARateCal(txtProcCode1)
        Else
            Call SearchMMARate(txtProcCode1)
        End If
        frmMMAEntryGST.txtSMMARate1 = IIf(Len(MMASurgFees), MMASurgFees, 0)
        frmMMAEntryGST.txtAnaesMMARate1 = IIf(Len(MMAAnaeFees), MMAAnaeFees, 0)
        
        frmMMAEntryGST.txtS13Rate1 = IIf(Len(MMASurgFees13th), MMASurgFees13th, 0)
        frmMMAEntryGST.txtAnae13Rate1 = IIf(Len(MMAAnaeFees13th), MMAAnaeFees13th, 0)
        
        frmMMAEntryGST.txtS13MinRate1 = IIf(Len(MMASurgFees13thMin), MMASurgFees13thMin, 0)
        
        Call Search14thRate(txtProcCode1)
        
        frmMMAEntryGST.txtAnae14Rate1 = IIf(Len(MMAAnaeFees13th), MMAAnaeFees13th, 0)
        
        frmMMAEntryGST.txtS14Rate1 = IIf(Len(MMASurgFees13th), MMASurgFees13th, 0)
        frmMMAEntryGST.txtS14MinRate1 = IIf(Len(MMASurgFees13th), MMASurgFees13th * 1.06, 0)
    Else
        frmMMAEntryGST.txtSMMARate1 = 0
        frmMMAEntryGST.txtAnaesMMARate1 = 0
        frmMMAEntryGST.txtS13Rate1 = 0
        frmMMAEntryGST.txtAnae13Rate1 = 0
        frmMMAEntryGST.txtS13MinRate1 = 0
    End If

End Sub

Private Sub txtProcCode2_Change()
    If Len(Trim(txtProcCode2)) Then
        If CalSource = 0 Then
            Call SearchMMARateCal(txtProcCode2)
        Else
            Call SearchMMARate(txtProcCode2)
        End If
        frmMMAEntryGST.txtSMMARate2 = IIf(Len(MMASurgFees), MMASurgFees, 0)
        frmMMAEntryGST.txtAnaesMMARate2 = IIf(Len(MMAAnaeFees), MMAAnaeFees, 0)
        frmMMAEntryGST.txtS13Rate2 = IIf(Len(MMASurgFees13th), MMASurgFees13th, 0)
        frmMMAEntryGST.txtAnae13Rate2 = IIf(Len(MMAAnaeFees13th), MMAAnaeFees13th, 0)
        frmMMAEntryGST.txtS13MinRate2 = IIf(Len(MMASurgFees13thMin), MMASurgFees13thMin, 0)
        
        Call Search14thRate(txtProcCode2)
        
        frmMMAEntryGST.txtAnae14Rate2 = IIf(Len(MMAAnaeFees13th), MMAAnaeFees13th, 0)
        frmMMAEntryGST.txtS14Rate2 = IIf(Len(MMASurgFees13th), MMASurgFees13th, 0)
        frmMMAEntryGST.txtS14MinRate2 = IIf(Len(MMASurgFees13th), MMASurgFees13th * 1.06, 0)
    Else
        frmMMAEntryGST.txtSMMARate2 = 0
        frmMMAEntryGST.txtAnaesMMARate2 = 0
        frmMMAEntryGST.txtS13Rate2 = 0
        frmMMAEntryGST.txtAnae13Rate2 = 0
        frmMMAEntryGST.txtS13MinRate2 = 0
        
    End If
End Sub

Private Sub txtProcCode3_Change()
    If Len(Trim(txtProcCode3)) Then
        If CalSource = 0 Then
            Call SearchMMARateCal(txtProcCode3)
        Else
            Call SearchMMARate(txtProcCode3)
        End If
        frmMMAEntryGST.txtSMMARate3 = IIf(Len(MMASurgFees), MMASurgFees, 0)
        frmMMAEntryGST.txtAnaesMMARate3 = IIf(Len(MMAAnaeFees), MMAAnaeFees, 0)
        
        frmMMAEntryGST.txtS13Rate3 = IIf(Len(MMASurgFees13th), MMASurgFees13th, 0)
        frmMMAEntryGST.txtAnae13Rate3 = IIf(Len(MMAAnaeFees13th), MMAAnaeFees13th, 0)
        
        frmMMAEntryGST.txtS13MinRate3 = IIf(Len(MMASurgFees13thMin), MMASurgFees13thMin, 0)
        
        Call Search14thRate(txtProcCode3)
        
        frmMMAEntryGST.txtAnae14Rate3 = IIf(Len(MMAAnaeFees13th), MMAAnaeFees13th, 0)
        frmMMAEntryGST.txtS14Rate3 = IIf(Len(MMASurgFees13th), MMASurgFees13th, 0)
        frmMMAEntryGST.txtS14MinRate3 = IIf(Len(MMASurgFees13th), MMASurgFees13th * 1.06, 0)
    Else
        frmMMAEntryGST.txtSMMARate3 = 0
        frmMMAEntryGST.txtAnaesMMARate3 = 0
        
        frmMMAEntryGST.txtS13Rate3 = 0
        frmMMAEntryGST.txtAnae13Rate3 = 0
        
        frmMMAEntryGST.txtS13MinRate3 = 0
    End If

End Sub

Private Sub txtProcCode4_Change()
    If Len(Trim(txtProcCode4)) Then
        If CalSource = 0 Then
            Call SearchMMARateCal(txtProcCode4)
        Else
            Call SearchMMARate(txtProcCode4)
        End If
        frmMMAEntryGST.txtSMMARate4 = IIf(Len(MMASurgFees), MMASurgFees, 0)
        frmMMAEntryGST.txtAnaesMMARate4 = IIf(Len(MMAAnaeFees), MMAAnaeFees, 0)
        
        frmMMAEntryGST.txtS13Rate4 = IIf(Len(MMASurgFees13th), MMASurgFees13th, 0)
        frmMMAEntryGST.txtAnae13Rate4 = IIf(Len(MMAAnaeFees13th), MMAAnaeFees13th, 0)
        
        frmMMAEntryGST.txtS13MinRate4 = IIf(Len(MMASurgFees13thMin), MMASurgFees13thMin, 0)
        
        Call Search14thRate(txtProcCode4)
        
        frmMMAEntryGST.txtAnae14Rate4 = IIf(Len(MMAAnaeFees13th), MMAAnaeFees13th, 0)
        frmMMAEntryGST.txtS14Rate4 = IIf(Len(MMASurgFees13th), MMASurgFees13th, 0)
        frmMMAEntryGST.txtS14MinRate4 = IIf(Len(MMASurgFees13th), MMASurgFees13th * 1.06, 0)
    Else
        frmMMAEntryGST.txtSMMARate4 = 0
        frmMMAEntryGST.txtAnaesMMARate4 = 0
        
        frmMMAEntryGST.txtS13Rate4 = 0
        frmMMAEntryGST.txtAnae13Rate4 = 0
        
        frmMMAEntryGST.txtS13MinRate4 = 0
        
    End If
End Sub

Private Sub txtProcCode5_Change()
    If Len(Trim(txtProcCode5)) Then
        If CalSource = 0 Then
            Call SearchMMARateCal(txtProcCode5)
        Else
            Call SearchMMARate(txtProcCode5)
        End If
        frmMMAEntryGST.txtSMMARate5 = IIf(Len(MMASurgFees), MMASurgFees, 0)
        frmMMAEntryGST.txtAnaesMMARate5 = IIf(Len(MMAAnaeFees), MMAAnaeFees, 0)
        frmMMAEntryGST.txtS13Rate5 = IIf(Len(MMASurgFees13th), MMASurgFees13th, 0)
        frmMMAEntryGST.txtAnae13Rate5 = IIf(Len(MMAAnaeFees13th), MMAAnaeFees13th, 0)
        frmMMAEntryGST.txtS13MinRate5 = IIf(Len(MMASurgFees13thMin), MMASurgFees13thMin, 0)
        
        
        Call Search14thRate(txtProcCode5)
        
        frmMMAEntryGST.txtAnae14Rate5 = IIf(Len(MMAAnaeFees13th), MMAAnaeFees13th, 0)
        frmMMAEntryGST.txtS14Rate5 = IIf(Len(MMASurgFees13th), MMASurgFees13th, 0)
        frmMMAEntryGST.txtS14MinRate5 = IIf(Len(MMASurgFees13th), MMASurgFees13th * 1.06, 0)
    Else
        frmMMAEntryGST.txtSMMARate5 = 0
        frmMMAEntryGST.txtAnaesMMARate5 = 0
        frmMMAEntryGST.txtS13Rate5 = 0
        frmMMAEntryGST.txtAnae13Rate5 = 0
        frmMMAEntryGST.txtS13MinRate5 = 0
    End If
End Sub

Private Sub txtProcCode6_Change()
    If Len(Trim(txtProcCode6)) Then
        If CalSource = 0 Then
            Call SearchMMARateCal(txtProcCode6)
        Else
            Call SearchMMARate(txtProcCode6)
        End If
        frmMMAEntryGST.txtSMMARate6 = IIf(Len(MMASurgFees), MMASurgFees, 0)
        frmMMAEntryGST.txtAnaesMMARate6 = IIf(Len(MMAAnaeFees), MMAAnaeFees, 0)
        
        frmMMAEntryGST.txtS13Rate6 = IIf(Len(MMASurgFees13th), MMASurgFees13th, 0)
        frmMMAEntryGST.txtAnae13Rate6 = IIf(Len(MMAAnaeFees13th), MMAAnaeFees13th, 0)
        
        frmMMAEntryGST.txtS13MinRate6 = IIf(Len(MMASurgFees13thMin), MMASurgFees13thMin, 0)
        
        Call Search14thRate(txtProcCode6)
        
        frmMMAEntryGST.txtAnae14Rate6 = IIf(Len(MMAAnaeFees13th), MMAAnaeFees13th, 0)
        frmMMAEntryGST.txtS14Rate6 = IIf(Len(MMASurgFees13th), MMASurgFees13th, 0)
        frmMMAEntryGST.txtS14MinRate6 = IIf(Len(MMASurgFees13th), MMASurgFees13th * 1.06, 0)
    Else
        frmMMAEntryGST.txtSMMARate6 = 0
        frmMMAEntryGST.txtAnaesMMARate6 = 0
        
        frmMMAEntryGST.txtS13Rate6 = 0
        frmMMAEntryGST.txtAnae13Rate6 = 0
        
        frmMMAEntryGST.txtS13MinRate6 = 0
    End If
End Sub
Private Function Search14thRate(OPCSCode As String)
Dim strsqlMMA As String
Dim MMA As New ADODB.Recordset
    MMASurgFees = 0
    MMAAnaeFees = 0
    MMASurgFees13th = 0
    MMAAnaeFees13th = 0
    MMASurgFees13thMin = 0
    
    'medixmasterconnMaint.Open "File Name=" & BOSPath & "Conn\MediConnHISMaint.UDL"

    strsqlMMA = "Select  Surgeon13th,Anesthetist13th,Surgeon13thMin  from MMA14Schedule where  OPCSCode = '" & Trim(OPCSCode) & "'"
    MMA.Open strsqlMMA, medixmasterconnMaint, adOpenForwardOnly, adLockBatchOptimistic
        Do While Not MMA.EOF
            MMASurgFees13th = IIf(Len(MMA!Surgeon13th), MMA!Surgeon13th, 0)
            MMAAnaeFees13th = IIf(Len(MMA!Anesthetist13th), MMA!Anesthetist13th, 0)
            MMASurgFees13thMin = IIf(Len(MMA!Surgeon13thMin), MMA!Surgeon13thMin, 0)
        MMA.MoveNext
        Loop
    MMA.Close
    Set MMA = Nothing
    'medixmasterconnMaint.Close
    'Set medixmasterconnMaint = Nothing
End Function

Private Sub txtProcCode7_Change()
    If Len(Trim(txtProcCode7)) Then
        If CalSource = 0 Then
            Call SearchMMARateCal(txtProcCode7)
        Else
            Call SearchMMARate(txtProcCode7)
        End If
        'frmMMAEntryGST.txtSMMARate7 = IIf(Len(MMASurgFees), MMASurgFees, 0)
        'frmMMAEntryGST.txtAnaesMMARate7 = IIf(Len(MMAAnaeFees), MMAAnaeFees, 0)
        
        frmMMAEntryGST.txtS13Rate7 = IIf(Len(MMASurgFees13th), MMASurgFees13th, 0)
        frmMMAEntryGST.txtAnae13Rate7 = IIf(Len(MMAAnaeFees13th), MMAAnaeFees13th, 0)
        
        'frmMMAEntryGST.txtS13MinRate7 = IIf(Len(MMASurgFees13thMin), MMASurgFees13thMin, 0)
        
        Call Search14thRate(txtProcCode7)
        
        frmMMAEntryGST.txtAnae14Rate7 = IIf(Len(MMAAnaeFees13th), MMAAnaeFees13th, 0)
        
        frmMMAEntryGST.txtS14Rate7 = IIf(Len(MMASurgFees13th), MMASurgFees13th, 0)
        frmMMAEntryGST.txtS14MinRate7 = IIf(Len(MMASurgFees13th), MMASurgFees13th * 1.06, 0)
    Else
        'frmMMAEntryGST.txtSMMARate7 = 0
        'frmMMAEntryGST.txtAnaesMMARate7 = 0
        frmMMAEntryGST.txtS13Rate7 = 0
        frmMMAEntryGST.txtAnae13Rate7 = 0
        'frmMMAEntryGST.txtS13MinRate7 = 0
    End If
End Sub

Private Sub txtProcCode8_Change()
    If Len(Trim(txtProcCode8)) Then
        If CalSource = 0 Then
            Call SearchMMARateCal(txtProcCode8)
        Else
            Call SearchMMARate(txtProcCode8)
        End If
        'frmMMAEntryGST.txtSMMARate8 = IIf(Len(MMASurgFees), MMASurgFees, 0)
        'frmMMAEntryGST.txtAnaesMMARate8 = IIf(Len(MMAAnaeFees), MMAAnaeFees, 0)
        
        frmMMAEntryGST.txtS13Rate8 = IIf(Len(MMASurgFees13th), MMASurgFees13th, 0)
        frmMMAEntryGST.txtAnae13Rate8 = IIf(Len(MMAAnaeFees13th), MMAAnaeFees13th, 0)
        
        'frmMMAEntryGST.txtS13MinRate8 = IIf(Len(MMASurgFees13thMin), MMASurgFees13thMin, 0)
        
        Call Search14thRate(txtProcCode8)
        
        frmMMAEntryGST.txtAnae14Rate8 = IIf(Len(MMAAnaeFees13th), MMAAnaeFees13th, 0)
        
        frmMMAEntryGST.txtS14Rate8 = IIf(Len(MMASurgFees13th), MMASurgFees13th, 0)
        frmMMAEntryGST.txtS14MinRate8 = IIf(Len(MMASurgFees13th), MMASurgFees13th * 1.06, 0)
    Else
        'frmMMAEntryGST.txtSMMARate8 = 0
        'frmMMAEntryGST.txtAnaesMMARate8 = 0
        frmMMAEntryGST.txtS13Rate8 = 0
        frmMMAEntryGST.txtAnae13Rate8 = 0
        'frmMMAEntryGST.txtS13MinRate8 = 0
    End If
End Sub

Private Sub txtS13MinRate1_Change()
Dim a, B, C, d, e, f, g As Double

    a = IIf(IsNumeric(txtS13MinRate1), txtS13MinRate1, 0)
    B = IIf(IsNumeric(txtS13MinRate2), txtS13MinRate2, 0)
    C = IIf(IsNumeric(txtS13MinRate3), txtS13MinRate3, 0)
    d = IIf(IsNumeric(txtS13MinRate4), txtS13MinRate4, 0)
    e = IIf(IsNumeric(txtS13MinRate5), txtS13MinRate5, 0)
    g = IIf(IsNumeric(txtS13MinRate6), txtS13MinRate6, 0)
    
    f = CDbl(a) + CDbl(B) + CDbl(C) + CDbl(d) + CDbl(e) + CDbl(g)
    txtTotalS13MinRate = f

End Sub

Private Sub txtS13MinRate2_Change()
Dim a, B, C, d, e, f, g As Double

    a = IIf(IsNumeric(txtS13MinRate1), txtS13MinRate1, 0)
    B = IIf(IsNumeric(txtS13MinRate2), txtS13MinRate2, 0)
    C = IIf(IsNumeric(txtS13MinRate3), txtS13MinRate3, 0)
    d = IIf(IsNumeric(txtS13MinRate4), txtS13MinRate4, 0)
    e = IIf(IsNumeric(txtS13MinRate5), txtS13MinRate5, 0)
    g = IIf(IsNumeric(txtS13MinRate6), txtS13MinRate6, 0)
    
    f = CDbl(a) + CDbl(B) + CDbl(C) + CDbl(d) + CDbl(e) + CDbl(g)
    txtTotalS13MinRate = f


End Sub

Private Sub txtS13MinRate3_Change()
Dim a, B, C, d, e, f, g As Double

    a = IIf(IsNumeric(txtS13MinRate1), txtS13MinRate1, 0)
    B = IIf(IsNumeric(txtS13MinRate2), txtS13MinRate2, 0)
    C = IIf(IsNumeric(txtS13MinRate3), txtS13MinRate3, 0)
    d = IIf(IsNumeric(txtS13MinRate4), txtS13MinRate4, 0)
    e = IIf(IsNumeric(txtS13MinRate5), txtS13MinRate5, 0)
    g = IIf(IsNumeric(txtS13MinRate6), txtS13MinRate6, 0)
    
    f = CDbl(a) + CDbl(B) + CDbl(C) + CDbl(d) + CDbl(e) + CDbl(g)
    txtTotalS13MinRate = f


End Sub

Private Sub txtS13MinRate4_Change()
Dim a, B, C, d, e, f, g As Double

    a = IIf(IsNumeric(txtS13MinRate1), txtS13MinRate1, 0)
    B = IIf(IsNumeric(txtS13MinRate2), txtS13MinRate2, 0)
    C = IIf(IsNumeric(txtS13MinRate3), txtS13MinRate3, 0)
    d = IIf(IsNumeric(txtS13MinRate4), txtS13MinRate4, 0)
    e = IIf(IsNumeric(txtS13MinRate5), txtS13MinRate5, 0)
    g = IIf(IsNumeric(txtS13MinRate6), txtS13MinRate6, 0)
    
    f = CDbl(a) + CDbl(B) + CDbl(C) + CDbl(d) + CDbl(e) + CDbl(g)
    txtTotalS13MinRate = f


End Sub

Private Sub txtS13MinRate5_Change()
Dim a, B, C, d, e, f, g As Double

    a = IIf(IsNumeric(txtS13MinRate1), txtS13MinRate1, 0)
    B = IIf(IsNumeric(txtS13MinRate2), txtS13MinRate2, 0)
    C = IIf(IsNumeric(txtS13MinRate3), txtS13MinRate3, 0)
    d = IIf(IsNumeric(txtS13MinRate4), txtS13MinRate4, 0)
    e = IIf(IsNumeric(txtS13MinRate5), txtS13MinRate5, 0)
    g = IIf(IsNumeric(txtS13MinRate6), txtS13MinRate6, 0)
    
    f = CDbl(a) + CDbl(B) + CDbl(C) + CDbl(d) + CDbl(e) + CDbl(g)
    txtTotalS13MinRate = f


End Sub

Private Sub txtS13MinRate6_Change()
Dim a, B, C, d, e, f, g As Double

    a = IIf(IsNumeric(txtS13MinRate1), txtS13MinRate1, 0)
    B = IIf(IsNumeric(txtS13MinRate2), txtS13MinRate2, 0)
    C = IIf(IsNumeric(txtS13MinRate3), txtS13MinRate3, 0)
    d = IIf(IsNumeric(txtS13MinRate4), txtS13MinRate4, 0)
    e = IIf(IsNumeric(txtS13MinRate5), txtS13MinRate5, 0)
    g = IIf(IsNumeric(txtS13MinRate6), txtS13MinRate6, 0)
    
    f = CDbl(a) + CDbl(B) + CDbl(C) + CDbl(d) + CDbl(e) + CDbl(g)
    txtTotalS13MinRate = f
End Sub

Private Sub txtS13Rate1_Change()
Dim a, B, C, d, e, f, g As Double

    a = IIf(IsNumeric(txtS13Rate1), txtS13Rate1, 0)
    B = IIf(IsNumeric(txtS13Rate2), txtS13Rate2, 0)
    C = IIf(IsNumeric(txtS13Rate3), txtS13Rate3, 0)
    d = IIf(IsNumeric(txtS13Rate4), txtS13Rate4, 0)
    e = IIf(IsNumeric(txtS13Rate5), txtS13Rate5, 0)
    g = IIf(IsNumeric(txtS13Rate6), txtS13Rate6, 0)
    
    f = CDbl(a) + CDbl(B) + CDbl(C) + CDbl(d) + CDbl(e) + CDbl(g)
    txtTotalS13Rate = f

    tmpNSFee = IIf(IsNumeric(txtSActFee1), txtSActFee1, 0)
    
    If CDbl(tmpNSFee) > CDbl(a) Then
        txtSActFee1.ForeColor = &HFF&
    Else
        txtSActFee1.ForeColor = &H80000008
    End If

End Sub

Private Sub txtS13Rate2_Change()
Dim a, B, C, d, e, f, g As Double

    a = IIf(IsNumeric(txtS13Rate1), txtS13Rate1, 0)
    B = IIf(IsNumeric(txtS13Rate2), txtS13Rate2, 0)
    C = IIf(IsNumeric(txtS13Rate3), txtS13Rate3, 0)
    d = IIf(IsNumeric(txtS13Rate4), txtS13Rate4, 0)
    e = IIf(IsNumeric(txtS13Rate5), txtS13Rate5, 0)
    g = IIf(IsNumeric(txtS13Rate6), txtS13Rate6, 0)
    
    f = CDbl(a) + CDbl(B) + CDbl(C) + CDbl(d) + CDbl(e) + CDbl(g)
    txtTotalS13Rate = f
    
    tmpNSFee = IIf(IsNumeric(txtSActFee2), txtSActFee2, 0)
    
    If CDbl(tmpNSFee) > CDbl(B) Then
        txtSActFee2.ForeColor = &HFF&
    Else
        txtSActFee2.ForeColor = &H80000008
    End If


End Sub

Private Sub txtS13Rate3_Change()
Dim a, B, C, d, e, f, g As Double

    a = IIf(IsNumeric(txtS13Rate1), txtS13Rate1, 0)
    B = IIf(IsNumeric(txtS13Rate2), txtS13Rate2, 0)
    C = IIf(IsNumeric(txtS13Rate3), txtS13Rate3, 0)
    d = IIf(IsNumeric(txtS13Rate4), txtS13Rate4, 0)
    e = IIf(IsNumeric(txtS13Rate5), txtS13Rate5, 0)
    g = IIf(IsNumeric(txtS13Rate6), txtS13Rate6, 0)
    
    f = CDbl(a) + CDbl(B) + CDbl(C) + CDbl(d) + CDbl(e) + CDbl(g)
    txtTotalS13Rate = f
    
    tmpNSFee = IIf(IsNumeric(txtSActFee3), txtSActFee3, 0)
    
    If CDbl(tmpNSFee) > CDbl(C) Then
        txtSActFee3.ForeColor = &HFF&
    Else
        txtSActFee3.ForeColor = &H80000008
    End If
    
End Sub

Private Sub txtS13Rate4_Change()
Dim a, B, C, d, e, f, g As Double

    a = IIf(IsNumeric(txtS13Rate1), txtS13Rate1, 0)
    B = IIf(IsNumeric(txtS13Rate2), txtS13Rate2, 0)
    C = IIf(IsNumeric(txtS13Rate3), txtS13Rate3, 0)
    d = IIf(IsNumeric(txtS13Rate4), txtS13Rate4, 0)
    e = IIf(IsNumeric(txtS13Rate5), txtS13Rate5, 0)
    g = IIf(IsNumeric(txtS13Rate6), txtS13Rate6, 0)
    
    f = CDbl(a) + CDbl(B) + CDbl(C) + CDbl(d) + CDbl(e) + CDbl(g)
    txtTotalS13Rate = f
    
    tmpNSFee = IIf(IsNumeric(txtSActFee4), txtSActFee4, 0)
    
    If CDbl(tmpNSFee) > CDbl(d) Then
        txtSActFee4.ForeColor = &HFF&
    Else
        txtSActFee4.ForeColor = &H80000008
    End If
    
End Sub

Private Sub txtS13Rate5_Change()
Dim a, B, C, d, e, f, g As Double

    a = IIf(IsNumeric(txtS13Rate1), txtS13Rate1, 0)
    B = IIf(IsNumeric(txtS13Rate2), txtS13Rate2, 0)
    C = IIf(IsNumeric(txtS13Rate3), txtS13Rate3, 0)
    d = IIf(IsNumeric(txtS13Rate4), txtS13Rate4, 0)
    e = IIf(IsNumeric(txtS13Rate5), txtS13Rate5, 0)
    g = IIf(IsNumeric(txtS13Rate6), txtS13Rate6, 0)
    
    f = CDbl(a) + CDbl(B) + CDbl(C) + CDbl(d) + CDbl(e) + CDbl(g)
    txtTotalS13Rate = f
    
    tmpNSFee = IIf(IsNumeric(txtSActFee5), txtSActFee5, 0)
    
    If CDbl(tmpNSFee) > CDbl(e) Then
        txtSActFee5.ForeColor = &HFF&
    Else
        txtSActFee5.ForeColor = &H80000008
    End If
    
End Sub

Private Sub txtS13Rate6_Change()
Dim a, B, C, d, e, f, g As Double

    a = IIf(IsNumeric(txtS13Rate1), txtS13Rate1, 0)
    B = IIf(IsNumeric(txtS13Rate2), txtS13Rate2, 0)
    C = IIf(IsNumeric(txtS13Rate3), txtS13Rate3, 0)
    d = IIf(IsNumeric(txtS13Rate4), txtS13Rate4, 0)
    e = IIf(IsNumeric(txtS13Rate5), txtS13Rate5, 0)
    g = IIf(IsNumeric(txtS13Rate6), txtS13Rate6, 0)
    
    f = CDbl(a) + CDbl(B) + CDbl(C) + CDbl(d) + CDbl(e) + CDbl(g)
    txtTotalS13Rate = f
    
    tmpNSFee = IIf(IsNumeric(txtSActFee6), txtSActFee6, 0)
    
    If CDbl(tmpNSFee) > CDbl(g) Then
        txtSActFee6.ForeColor = &HFF&
    Else
        txtSActFee6.ForeColor = &H80000008
    End If
    
End Sub

Private Sub txtSActFee1_Change()
Dim a, B, C, d, e, f, g, h, i As Double

    a = IIf(IsNumeric(txtSActFee1), txtSActFee1, 0)
    B = IIf(IsNumeric(txtSActFee2), txtSActFee2, 0)
    C = IIf(IsNumeric(txtSActFee3), txtSActFee3, 0)
    d = IIf(IsNumeric(txtSActFee4), txtSActFee4, 0)
    e = IIf(IsNumeric(txtSActFee5), txtSActFee5, 0)
    g = IIf(IsNumeric(txtSActFee6), txtSActFee6, 0)
    h = IIf(IsNumeric(txtSActFee7), txtSActFee7, 0)
    i = IIf(IsNumeric(txtSActFee8), txtSActFee8, 0)
    
    f = CDbl(a) + CDbl(B) + CDbl(C) + CDbl(d) + CDbl(e) + CDbl(g) + CDbl(h) + CDbl(i)
    txtTotalSActFee = f
    
    tmpNSFee = IIf(IsNumeric(txtS13Rate1), txtS13Rate1, 0)
    
    If CDbl(a) > CDbl(tmpNSFee) Then
        txtSActFee1.ForeColor = &HFF&
    Else
        txtSActFee1.ForeColor = &H80000008
    End If
    
End Sub

Private Sub txtSActFee1_KeyPress(KeyAscii As Integer)
    KeyAscii = KeyNumber(KeyAscii)
End Sub

Private Sub txtSActFee2_Change()
Dim a, B, C, d, e, f, g, h, i As Double

    a = IIf(IsNumeric(txtSActFee1), txtSActFee1, 0)
    B = IIf(IsNumeric(txtSActFee2), txtSActFee2, 0)
    C = IIf(IsNumeric(txtSActFee3), txtSActFee3, 0)
    d = IIf(IsNumeric(txtSActFee4), txtSActFee4, 0)
    e = IIf(IsNumeric(txtSActFee5), txtSActFee5, 0)
    g = IIf(IsNumeric(txtSActFee6), txtSActFee6, 0)
    h = IIf(IsNumeric(txtSActFee7), txtSActFee7, 0)
    i = IIf(IsNumeric(txtSActFee8), txtSActFee8, 0)
    
    f = CDbl(a) + CDbl(B) + CDbl(C) + CDbl(d) + CDbl(e) + CDbl(g) + CDbl(h) + CDbl(i)
    txtTotalSActFee = f
    
    tmpNSFee = IIf(IsNumeric(txtS13Rate2), txtS13Rate2, 0)
    
    If CDbl(B) > CDbl(tmpNSFee) Then
        txtSActFee2.ForeColor = &HFF&
    Else
        txtSActFee2.ForeColor = &H80000008
    End If
    
End Sub

Private Sub txtSActFee2_KeyPress(KeyAscii As Integer)
    KeyAscii = KeyNumber(KeyAscii)
End Sub

Private Sub txtSActFee3_Change()
Dim a, B, C, d, e, f, g, h, i As Double

    a = IIf(IsNumeric(txtSActFee1), txtSActFee1, 0)
    B = IIf(IsNumeric(txtSActFee2), txtSActFee2, 0)
    C = IIf(IsNumeric(txtSActFee3), txtSActFee3, 0)
    d = IIf(IsNumeric(txtSActFee4), txtSActFee4, 0)
    e = IIf(IsNumeric(txtSActFee5), txtSActFee5, 0)
    g = IIf(IsNumeric(txtSActFee6), txtSActFee6, 0)
    h = IIf(IsNumeric(txtSActFee7), txtSActFee7, 0)
    i = IIf(IsNumeric(txtSActFee8), txtSActFee8, 0)
    
    f = CDbl(a) + CDbl(B) + CDbl(C) + CDbl(d) + CDbl(e) + CDbl(g) + CDbl(h) + CDbl(i)
    txtTotalSActFee = f
    
    tmpNSFee = IIf(IsNumeric(txtS13Rate3), txtS13Rate3, 0)
    
    If CDbl(C) > CDbl(tmpNSFee) Then
        txtSActFee3.ForeColor = &HFF&
    Else
        txtSActFee3.ForeColor = &H80000008
    End If
    
End Sub

Private Sub txtSActFee3_KeyPress(KeyAscii As Integer)
    KeyAscii = KeyNumber(KeyAscii)
End Sub

Private Sub txtSActFee4_Change()
Dim a, B, C, d, e, f, g, h, i As Double

    a = IIf(IsNumeric(txtSActFee1), txtSActFee1, 0)
    B = IIf(IsNumeric(txtSActFee2), txtSActFee2, 0)
    C = IIf(IsNumeric(txtSActFee3), txtSActFee3, 0)
    d = IIf(IsNumeric(txtSActFee4), txtSActFee4, 0)
    e = IIf(IsNumeric(txtSActFee5), txtSActFee5, 0)
    g = IIf(IsNumeric(txtSActFee6), txtSActFee6, 0)
    h = IIf(IsNumeric(txtSActFee7), txtSActFee7, 0)
    i = IIf(IsNumeric(txtSActFee8), txtSActFee8, 0)
    
    
    f = CDbl(a) + CDbl(B) + CDbl(C) + CDbl(d) + CDbl(e) + CDbl(g) + CDbl(h) + CDbl(i)
    txtTotalSActFee = f
    
    tmpNSFee = IIf(IsNumeric(txtS13Rate4), txtS13Rate4, 0)
    
    If CDbl(d) > CDbl(tmpNSFee) Then
        txtSActFee4.ForeColor = &HFF&
    Else
        txtSActFee4.ForeColor = &H80000008
    End If
    
End Sub

Private Sub txtSActFee4_KeyPress(KeyAscii As Integer)
    KeyAscii = KeyNumber(KeyAscii)
End Sub

Private Sub txtSActFee5_Change()
Dim a, B, C, d, e, f, g, h, i As Double

    a = IIf(IsNumeric(txtSActFee1), txtSActFee1, 0)
    B = IIf(IsNumeric(txtSActFee2), txtSActFee2, 0)
    C = IIf(IsNumeric(txtSActFee3), txtSActFee3, 0)
    d = IIf(IsNumeric(txtSActFee4), txtSActFee4, 0)
    e = IIf(IsNumeric(txtSActFee5), txtSActFee5, 0)
    g = IIf(IsNumeric(txtSActFee6), txtSActFee6, 0)
    h = IIf(IsNumeric(txtSActFee7), txtSActFee7, 0)
    i = IIf(IsNumeric(txtSActFee8), txtSActFee8, 0)
    
    f = CDbl(a) + CDbl(B) + CDbl(C) + CDbl(d) + CDbl(e) + CDbl(g) + CDbl(h) + CDbl(i)
    txtTotalSActFee = f
    
    tmpNSFee = IIf(IsNumeric(txtS13Rate5), txtS13Rate5, 0)
    
    If CDbl(e) > CDbl(tmpNSFee) Then
        txtSActFee5.ForeColor = &HFF&
    Else
        txtSActFee5.ForeColor = &H80000008
    End If
    
End Sub

Private Sub txtSActFee5_KeyPress(KeyAscii As Integer)
    KeyAscii = KeyNumber(KeyAscii)
End Sub

Private Sub txtSActFee6_Change()
Dim a, B, C, d, e, f, g, h, i As Double

    a = IIf(IsNumeric(txtSActFee1), txtSActFee1, 0)
    B = IIf(IsNumeric(txtSActFee2), txtSActFee2, 0)
    C = IIf(IsNumeric(txtSActFee3), txtSActFee3, 0)
    d = IIf(IsNumeric(txtSActFee4), txtSActFee4, 0)
    e = IIf(IsNumeric(txtSActFee5), txtSActFee5, 0)
    g = IIf(IsNumeric(txtSActFee6), txtSActFee6, 0)
    h = IIf(IsNumeric(txtSActFee7), txtSActFee7, 0)
    i = IIf(IsNumeric(txtSActFee8), txtSActFee8, 0)
    
    f = CDbl(a) + CDbl(B) + CDbl(C) + CDbl(d) + CDbl(e) + CDbl(g) + CDbl(h) + CDbl(i)
    txtTotalSActFee = f
    
    tmpNSFee = IIf(IsNumeric(txtS13Rate6), txtS13Rate6, 0)
    
    If CDbl(g) > CDbl(tmpNSFee) Then
        txtSActFee6.ForeColor = &HFF&
    Else
        txtSActFee6.ForeColor = &H80000008
    End If
    
End Sub

Private Sub txtSActFee7_Change()
Dim a, B, C, d, e, f, g, h, i As Double

    a = IIf(IsNumeric(txtSActFee1), txtSActFee1, 0)
    B = IIf(IsNumeric(txtSActFee2), txtSActFee2, 0)
    C = IIf(IsNumeric(txtSActFee3), txtSActFee3, 0)
    d = IIf(IsNumeric(txtSActFee4), txtSActFee4, 0)
    e = IIf(IsNumeric(txtSActFee5), txtSActFee5, 0)
    g = IIf(IsNumeric(txtSActFee6), txtSActFee6, 0)
    h = IIf(IsNumeric(txtSActFee7), txtSActFee7, 0)
    i = IIf(IsNumeric(txtSActFee8), txtSActFee8, 0)
    
    f = CDbl(a) + CDbl(B) + CDbl(C) + CDbl(d) + CDbl(e) + CDbl(g) + CDbl(h) + CDbl(i)
    txtTotalSActFee = f
    
    tmpNSFee = IIf(IsNumeric(txtS13Rate7), txtS13Rate7, 0)
    
    If CDbl(g) > CDbl(tmpNSFee) Then
        txtSActFee7.ForeColor = &HFF&
    Else
        txtSActFee7.ForeColor = &H80000008
    End If
End Sub

Private Sub txtSActFee8_Change()
Dim a, B, C, d, e, f, g, h, i As Double

    a = IIf(IsNumeric(txtSActFee1), txtSActFee1, 0)
    B = IIf(IsNumeric(txtSActFee2), txtSActFee2, 0)
    C = IIf(IsNumeric(txtSActFee3), txtSActFee3, 0)
    d = IIf(IsNumeric(txtSActFee4), txtSActFee4, 0)
    e = IIf(IsNumeric(txtSActFee5), txtSActFee5, 0)
    g = IIf(IsNumeric(txtSActFee6), txtSActFee6, 0)
    h = IIf(IsNumeric(txtSActFee7), txtSActFee7, 0)
    i = IIf(IsNumeric(txtSActFee8), txtSActFee8, 0)
    
    f = CDbl(a) + CDbl(B) + CDbl(C) + CDbl(d) + CDbl(e) + CDbl(g) + CDbl(h) + CDbl(i)
    txtTotalSActFee = f
    
    tmpNSFee = IIf(IsNumeric(txtS13Rate8), txtS13Rate8, 0)
    
    If CDbl(g) > CDbl(tmpNSFee) Then
        txtSActFee8.ForeColor = &HFF&
    Else
        txtSActFee8.ForeColor = &H80000008
    End If
End Sub

Private Sub txtSMMARate1_Change()
Dim a, B, C, d, e, f, g As Double

    a = IIf(IsNumeric(txtSMMARate1), txtSMMARate1, 0)
    B = IIf(IsNumeric(txtSMMARate2), txtSMMARate2, 0)
    C = IIf(IsNumeric(txtSMMARate3), txtSMMARate3, 0)
    d = IIf(IsNumeric(txtSMMARate4), txtSMMARate4, 0)
    e = IIf(IsNumeric(txtSMMARate5), txtSMMARate5, 0)
    g = IIf(IsNumeric(txtSMMARate6), txtSMMARate6, 0)
    
    f = CDbl(a) + CDbl(B) + CDbl(C) + CDbl(d) + CDbl(e) + CDbl(g)
    txtSTotalMMARate1 = f
End Sub

Private Sub txtSMMARate1_KeyPress(KeyAscii As Integer)
    KeyAscii = KeyNumber(KeyAscii)
End Sub

Private Sub txtSMMARate2_Change()
Dim a, B, C, d, e, f, g As Double

    a = IIf(IsNumeric(txtSMMARate1), txtSMMARate1, 0)
    B = IIf(IsNumeric(txtSMMARate2), txtSMMARate2, 0)
    C = IIf(IsNumeric(txtSMMARate3), txtSMMARate3, 0)
    d = IIf(IsNumeric(txtSMMARate4), txtSMMARate4, 0)
    e = IIf(IsNumeric(txtSMMARate5), txtSMMARate5, 0)
    g = IIf(IsNumeric(txtSMMARate6), txtSMMARate6, 0)
    
    f = CDbl(a) + CDbl(B) + CDbl(C) + CDbl(d) + CDbl(e) + CDbl(g)
    txtSTotalMMARate1 = f
End Sub

Private Sub txtSMMARate2_KeyPress(KeyAscii As Integer)
    KeyAscii = KeyNumber(KeyAscii)
End Sub

Private Sub txtSMMARate3_Change()
Dim a, B, C, d, e, f, g As Double

    a = IIf(IsNumeric(txtSMMARate1), txtSMMARate1, 0)
    B = IIf(IsNumeric(txtSMMARate2), txtSMMARate2, 0)
    C = IIf(IsNumeric(txtSMMARate3), txtSMMARate3, 0)
    d = IIf(IsNumeric(txtSMMARate4), txtSMMARate4, 0)
    e = IIf(IsNumeric(txtSMMARate5), txtSMMARate5, 0)
    g = IIf(IsNumeric(txtSMMARate6), txtSMMARate6, 0)
    
    f = CDbl(a) + CDbl(B) + CDbl(C) + CDbl(d) + CDbl(e) + CDbl(g)
    txtSTotalMMARate1 = f
End Sub

Private Sub txtSMMARate3_KeyPress(KeyAscii As Integer)
    KeyAscii = KeyNumber(KeyAscii)
End Sub

Private Sub txtSMMARate4_Change()
Dim a, B, C, d, e, f, g As Double

    a = IIf(IsNumeric(txtSMMARate1), txtSMMARate1, 0)
    B = IIf(IsNumeric(txtSMMARate2), txtSMMARate2, 0)
    C = IIf(IsNumeric(txtSMMARate3), txtSMMARate3, 0)
    d = IIf(IsNumeric(txtSMMARate4), txtSMMARate4, 0)
    e = IIf(IsNumeric(txtSMMARate5), txtSMMARate5, 0)
    g = IIf(IsNumeric(txtSMMARate6), txtSMMARate6, 0)
    
    f = CDbl(a) + CDbl(B) + CDbl(C) + CDbl(d) + CDbl(e) + CDbl(g)
    txtSTotalMMARate1 = f
End Sub

Private Sub txtSMMARate4_KeyPress(KeyAscii As Integer)
    KeyAscii = KeyNumber(KeyAscii)
End Sub

Private Sub txtSMMARate5_Change()
Dim a, B, C, d, e, f, g As Double

    a = IIf(IsNumeric(txtSMMARate1), txtSMMARate1, 0)
    B = IIf(IsNumeric(txtSMMARate2), txtSMMARate2, 0)
    C = IIf(IsNumeric(txtSMMARate3), txtSMMARate3, 0)
    d = IIf(IsNumeric(txtSMMARate4), txtSMMARate4, 0)
    e = IIf(IsNumeric(txtSMMARate5), txtSMMARate5, 0)
    g = IIf(IsNumeric(txtSMMARate6), txtSMMARate6, 0)
    
    f = CDbl(a) + CDbl(B) + CDbl(C) + CDbl(d) + CDbl(e) + CDbl(g)
    txtSTotalMMARate1 = f
End Sub

Private Sub txtSMMARate5_KeyPress(KeyAscii As Integer)
    KeyAscii = KeyNumber(KeyAscii)
End Sub

Private Sub SearchMMARate(CPTCode As String)
Dim strsqlMMA As String
Dim MMA As New ADODB.Recordset
    
    MMASurgFees = 0
    MMAAnaeFees = 0
    MMASurgFees13th = 0
    MMAAnaeFees13th = 0
    MMASurgFees13thMin = 0
    strsqlMMA = ""
    strsqlMMA = "Select SurgeonsFees, AnaesthetistsFees,Surgeon13th,Anesthetist13th,Surgeon13thMin  from VIEW_MMA_Schedule where  OPCSCode = '" & Trim(CPTCode) & "'"
    MMA.Open strsqlMMA, medixmasterconnMaint, adOpenForwardOnly, adLockOptimistic
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

Private Sub SearchMMADesc(CPTCode As String)
Dim strsqlMMA As String
Dim MMA As New ADODB.Recordset

    MMADesc = ""
    Sch13 = ""
   ' medixmasterconnMaint.Open "File Name=" & BOSPath & "Conn\MediConnHISMaint.UDL"
    strsqlMMA = "Select OPCSCode, OPCSNarrative, ProcedureDescription,ProcMainBPCode,BodyParts,ProcPageNo,Page,  " & _
                "SurgeonsFees, AnaesthetistsFees, Surgeon13thMin,Surgeon13th,Anesthetist13th  " & _
                "from VIEW_MMA_Schedule where  OPCSCode = '" & Trim(CPTCode) & "'"
    MMA.Open strsqlMMA, medixmasterconnMaint, adOpenForwardOnly, adLockOptimistic
        Do While Not MMA.EOF
            MMADesc = IIf(Len(MMA!OPCSNarrative), MMA!OPCSNarrative, "")
            Sch13 = IIf(Len(MMA!ProcedureDescription), MMA!ProcedureDescription, "")
            BodyMMA = IIf(Len(MMA!ProcMainBPCode), MMA!ProcMainBPCode, "")
            Body13 = IIf(Len(MMA!BodyParts), MMA!BodyParts, "")
            PageMMA = IIf(Len(MMA!ProcPageNo), MMA!ProcPageNo, "")
            Page13 = IIf(Len(MMA!Page), MMA!Page, "")
            MMASurgFees = IIf(Len(MMA!SurgeonsFees), MMA!SurgeonsFees, 0)
            MMAAnaeFees = IIf(Len(MMA!AnaesthetistsFees), MMA!AnaesthetistsFees, 0)
            MMASurgFees13thMin = IIf(Len(MMA!Surgeon13thMin), MMA!Surgeon13thMin, 0)
            MMASurgFees13th = IIf(Len(MMA!Surgeon13th), MMA!Surgeon13th, 0)
            MMAAnaeFees13th = IIf(Len(MMA!Anesthetist13th), MMA!Anesthetist13th, 0)
        MMA.MoveNext
        Loop
    MMA.Close
    Set MMA = Nothing
 '   medixmasterconnMaint.Close
  '  Set medixmasterconnMaint = Nothing
End Sub


Private Sub SearchMMARateCal(CPTCode As String)
Dim strsqlMMA As String
Dim MMA As New ADODB.Recordset
    
    MMASurgFees = 0
    MMAAnaeFees = 0
    MMASurgFees13th = 0
    MMAAnaeFees13th = 0
    MMASurgFees13thMin = 0
    
   ' medixmasterconnMaint.Open "File Name=" & BOSPath & "Conn\MediConnHISMaint.UDL"

    strsqlMMA = "Select SurgeonsFees, AnaesthetistsFees,Surgeon13th,Anesthetist13th,Surgeon13thMin  from VIEW_MMA_Schedule where  OPCSCode = '" & Trim(CPTCode) & "'"
    MMA.Open strsqlMMA, medixmasterconnMaint, adOpenForwardOnly, adLockOptimistic
        Do While Not MMA.EOF
            MMASurgFees = IIf(Len(MMA!SurgeonsFees), MMA!SurgeonsFees, 0)
            MMAAnaeFees = IIf(Len(MMA!AnaesthetistsFees), MMA!AnaesthetistsFees, 0)
            MMASurgFees13th = IIf(Len(MMA!Surgeon13th), MMA!Surgeon13th, 0)
            MMAAnaeFees13th = IIf(Len(MMA!Anesthetist13th), MMA!Anesthetist13th, 0)
            MMASurgFees13thMin = IIf(Len(MMA!Surgeon13thMin), MMA!Surgeon13thMin, 0)
        MMA.MoveNext
        Loop
        'Else
        '    MMASurgFees = 0
        '    MMAAnaeFees = 0
        '    MMASurgFees13th = 0
        '    MMAAnaeFees13th = 0
        'End If
    MMA.Close
    Set MMA = Nothing
   ' medixmasterconnMaint.Close
   ' Set medixmasterconnMaint = Nothing

End Sub

Private Sub txtSMMARate6_Change()
Dim a, B, C, d, e, f, g As Double

    a = IIf(IsNumeric(txtSMMARate1), txtSMMARate1, 0)
    B = IIf(IsNumeric(txtSMMARate2), txtSMMARate2, 0)
    C = IIf(IsNumeric(txtSMMARate3), txtSMMARate3, 0)
    d = IIf(IsNumeric(txtSMMARate4), txtSMMARate4, 0)
    e = IIf(IsNumeric(txtSMMARate5), txtSMMARate5, 0)
    g = IIf(IsNumeric(txtSMMARate6), txtSMMARate6, 0)
    
    f = CDbl(a) + CDbl(B) + CDbl(C) + CDbl(d) + CDbl(e) + CDbl(g)
    txtSTotalMMARate1 = f
End Sub



