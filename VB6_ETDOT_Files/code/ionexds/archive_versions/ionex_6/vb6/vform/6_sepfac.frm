VERSION 5.00
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "Threed32.ocx"
Begin VB.Form frmSeparationFactors 
   Appearance      =   0  'Flat
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Input Separation Factors (Alpha's)"
   ClientHeight    =   6675
   ClientLeft      =   870
   ClientTop       =   1230
   ClientWidth     =   9105
   ControlBox      =   0   'False
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   8.25
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ForeColor       =   &H80000008&
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   6675
   ScaleWidth      =   9105
   Begin Threed.SSPanel panelDescription 
      Height          =   1575
      Left            =   3600
      TabIndex        =   145
      Top             =   4800
      Width           =   3015
      _Version        =   65536
      _ExtentX        =   5318
      _ExtentY        =   2778
      _StockProps     =   15
      Caption         =   " "
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Begin VB.Shape Shape1 
         Height          =   675
         Left            =   840
         Top             =   720
         Width           =   1335
      End
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00800000&
         Caption         =   "i"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   1740
         TabIndex        =   170
         Top             =   780
         Width           =   315
      End
      Begin VB.Label Label3 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00808000&
         Caption         =   "j"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   1740
         TabIndex        =   169
         Top             =   1080
         Width           =   300
      End
      Begin VB.Label Label4 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "Alpha"
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   900
         TabIndex        =   168
         Top             =   960
         Width           =   795
      End
      Begin VB.Label lblDescription 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Values in Above Table Represent Separation Factors of the Ions in the System"
         ForeColor       =   &H80000008&
         Height          =   615
         Left            =   360
         TabIndex        =   167
         Top             =   120
         Width           =   2535
      End
   End
   Begin Threed.SSPanel panelIons 
      Height          =   2055
      Left            =   720
      TabIndex        =   144
      Top             =   4560
      Width           =   2535
      _Version        =   65536
      _ExtentX        =   4471
      _ExtentY        =   3625
      _StockProps     =   15
      Caption         =   " "
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Begin VB.Label lblIonNumber 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "10"
         ForeColor       =   &H80000008&
         Height          =   192
         Index           =   10
         Left            =   0
         TabIndex        =   166
         Top             =   1800
         Width           =   312
      End
      Begin VB.Label lblIonNumber 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "9"
         ForeColor       =   &H80000008&
         Height          =   192
         Index           =   9
         Left            =   0
         TabIndex        =   165
         Top             =   1620
         Width           =   312
      End
      Begin VB.Label lblIonNumber 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "8"
         ForeColor       =   &H80000008&
         Height          =   192
         Index           =   8
         Left            =   0
         TabIndex        =   164
         Top             =   1440
         Width           =   312
      End
      Begin VB.Label lblIonNumber 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "7"
         ForeColor       =   &H80000008&
         Height          =   192
         Index           =   7
         Left            =   0
         TabIndex        =   163
         Top             =   1260
         Width           =   312
      End
      Begin VB.Label lblIonNumber 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "6"
         ForeColor       =   &H80000008&
         Height          =   192
         Index           =   6
         Left            =   0
         TabIndex        =   162
         Top             =   1080
         Width           =   312
      End
      Begin VB.Label lblIonNumber 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "5"
         ForeColor       =   &H80000008&
         Height          =   192
         Index           =   5
         Left            =   0
         TabIndex        =   161
         Top             =   900
         Width           =   312
      End
      Begin VB.Label lblIonNumber 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "4"
         ForeColor       =   &H80000008&
         Height          =   192
         Index           =   4
         Left            =   0
         TabIndex        =   160
         Top             =   720
         Width           =   312
      End
      Begin VB.Label lblIonNumber 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "3"
         ForeColor       =   &H80000008&
         Height          =   192
         Index           =   3
         Left            =   0
         TabIndex        =   159
         Top             =   540
         Width           =   312
      End
      Begin VB.Label lblIonNumber 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "2"
         ForeColor       =   &H80000008&
         Height          =   192
         Index           =   2
         Left            =   0
         TabIndex        =   158
         Top             =   360
         Width           =   312
      End
      Begin VB.Label lblIonNumber 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "1"
         ForeColor       =   &H80000008&
         Height          =   192
         Index           =   1
         Left            =   0
         TabIndex        =   157
         Top             =   180
         Width           =   312
      End
      Begin VB.Label lblIonName 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   192
         Index           =   10
         Left            =   300
         TabIndex        =   156
         Top             =   1800
         Width           =   2292
      End
      Begin VB.Label lblIonName 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   192
         Index           =   9
         Left            =   300
         TabIndex        =   155
         Top             =   1620
         Width           =   2292
      End
      Begin VB.Label lblIonName 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   192
         Index           =   8
         Left            =   300
         TabIndex        =   154
         Top             =   1440
         Width           =   2292
      End
      Begin VB.Label lblIonName 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   192
         Index           =   7
         Left            =   300
         TabIndex        =   153
         Top             =   1260
         Width           =   2292
      End
      Begin VB.Label lblIonName 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   192
         Index           =   6
         Left            =   300
         TabIndex        =   152
         Top             =   1080
         Width           =   2292
      End
      Begin VB.Label lblIonName 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   192
         Index           =   5
         Left            =   300
         TabIndex        =   151
         Top             =   900
         Width           =   2292
      End
      Begin VB.Label lblIonName 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   192
         Index           =   4
         Left            =   300
         TabIndex        =   150
         Top             =   720
         Width           =   2292
      End
      Begin VB.Label lblIonName 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   192
         Index           =   3
         Left            =   300
         TabIndex        =   149
         Top             =   540
         Width           =   2292
      End
      Begin VB.Label lblIonName 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   192
         Index           =   2
         Left            =   300
         TabIndex        =   148
         Top             =   360
         Width           =   2292
      End
      Begin VB.Label lblIonName 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   192
         Index           =   1
         Left            =   300
         TabIndex        =   147
         Top             =   180
         Width           =   2292
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Ion Name"
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   240
         TabIndex        =   146
         Top             =   0
         Width           =   2295
      End
   End
   Begin Threed.SSOption optSeparationFactors 
      Height          =   255
      Index           =   1
      Left            =   1200
      TabIndex        =   124
      Top             =   4200
      Width           =   255
      _Version        =   65536
      _ExtentX        =   450
      _ExtentY        =   450
      _StockProps     =   78
      Caption         =   "SSOption1"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.TextBox txtSeparationFactorsRow3 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      Height          =   288
      Index           =   4
      Left            =   3240
      TabIndex        =   23
      Text            =   "1"
      Top             =   1320
      Visible         =   0   'False
      Width           =   732
   End
   Begin VB.CommandButton cmdCancel 
      Appearance      =   0  'Flat
      Caption         =   "&Cancel"
      Height          =   372
      Left            =   8040
      TabIndex        =   102
      Top             =   5760
      Width           =   975
   End
   Begin VB.CommandButton cmdOK 
      Appearance      =   0  'Flat
      Caption         =   "&OK"
      Height          =   372
      Left            =   6840
      TabIndex        =   123
      Top             =   5760
      Width           =   975
   End
   Begin VB.TextBox txtSeparationFactorsRow10 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      Enabled         =   0   'False
      Height          =   288
      Index           =   10
      Left            =   7920
      TabIndex        =   99
      Text            =   "1.00"
      Top             =   3840
      Visible         =   0   'False
      Width           =   732
   End
   Begin VB.TextBox txtSeparationFactorsRow10 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      Height          =   288
      Index           =   9
      Left            =   7140
      TabIndex        =   98
      Text            =   "1"
      Top             =   3840
      Visible         =   0   'False
      Width           =   732
   End
   Begin VB.TextBox txtSeparationFactorsRow10 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      Height          =   288
      Index           =   8
      Left            =   6360
      TabIndex        =   97
      Text            =   "1"
      Top             =   3840
      Visible         =   0   'False
      Width           =   732
   End
   Begin VB.TextBox txtSeparationFactorsRow10 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      Height          =   288
      Index           =   7
      Left            =   5580
      TabIndex        =   96
      Text            =   "1"
      Top             =   3840
      Visible         =   0   'False
      Width           =   732
   End
   Begin VB.TextBox txtSeparationFactorsRow10 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      Height          =   288
      Index           =   6
      Left            =   4800
      TabIndex        =   95
      Text            =   "1"
      Top             =   3840
      Visible         =   0   'False
      Width           =   732
   End
   Begin VB.TextBox txtSeparationFactorsRow10 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      Height          =   288
      Index           =   5
      Left            =   4020
      TabIndex        =   94
      Text            =   "1"
      Top             =   3840
      Visible         =   0   'False
      Width           =   732
   End
   Begin VB.TextBox txtSeparationFactorsRow10 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      Height          =   288
      Index           =   4
      Left            =   3240
      TabIndex        =   93
      Text            =   "1"
      Top             =   3840
      Visible         =   0   'False
      Width           =   732
   End
   Begin VB.TextBox txtSeparationFactorsRow10 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      Height          =   288
      Index           =   3
      Left            =   2460
      TabIndex        =   92
      Text            =   "1"
      Top             =   3840
      Visible         =   0   'False
      Width           =   732
   End
   Begin VB.TextBox txtSeparationFactorsRow10 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      Height          =   288
      Index           =   2
      Left            =   1680
      TabIndex        =   91
      Text            =   "1"
      Top             =   3840
      Visible         =   0   'False
      Width           =   732
   End
   Begin VB.TextBox txtSeparationFactorsRow10 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      Height          =   288
      Index           =   1
      Left            =   900
      TabIndex        =   90
      Text            =   "1"
      Top             =   3840
      Visible         =   0   'False
      Width           =   732
   End
   Begin VB.TextBox txtSeparationFactorsRow9 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      Height          =   288
      Index           =   10
      Left            =   7920
      TabIndex        =   89
      Text            =   "1"
      Top             =   3480
      Visible         =   0   'False
      Width           =   732
   End
   Begin VB.TextBox txtSeparationFactorsRow9 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      Enabled         =   0   'False
      Height          =   288
      Index           =   9
      Left            =   7140
      TabIndex        =   88
      Text            =   "1.00"
      Top             =   3480
      Visible         =   0   'False
      Width           =   732
   End
   Begin VB.TextBox txtSeparationFactorsRow9 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      Height          =   288
      Index           =   8
      Left            =   6360
      TabIndex        =   87
      Text            =   "1"
      Top             =   3480
      Visible         =   0   'False
      Width           =   732
   End
   Begin VB.TextBox txtSeparationFactorsRow9 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      Height          =   288
      Index           =   7
      Left            =   5580
      TabIndex        =   86
      Text            =   "1"
      Top             =   3480
      Visible         =   0   'False
      Width           =   732
   End
   Begin VB.TextBox txtSeparationFactorsRow9 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      Height          =   288
      Index           =   6
      Left            =   4800
      TabIndex        =   85
      Text            =   "1"
      Top             =   3480
      Visible         =   0   'False
      Width           =   732
   End
   Begin VB.TextBox txtSeparationFactorsRow9 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      Height          =   288
      Index           =   5
      Left            =   4020
      TabIndex        =   84
      Text            =   "1"
      Top             =   3480
      Visible         =   0   'False
      Width           =   732
   End
   Begin VB.TextBox txtSeparationFactorsRow9 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      Height          =   288
      Index           =   4
      Left            =   3240
      TabIndex        =   83
      Text            =   "1"
      Top             =   3480
      Visible         =   0   'False
      Width           =   732
   End
   Begin VB.TextBox txtSeparationFactorsRow9 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      Height          =   288
      Index           =   3
      Left            =   2460
      TabIndex        =   82
      Text            =   "1"
      Top             =   3480
      Visible         =   0   'False
      Width           =   732
   End
   Begin VB.TextBox txtSeparationFactorsRow9 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      Height          =   288
      Index           =   2
      Left            =   1680
      TabIndex        =   81
      Text            =   "1"
      Top             =   3480
      Visible         =   0   'False
      Width           =   732
   End
   Begin VB.TextBox txtSeparationFactorsRow9 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      Height          =   288
      Index           =   1
      Left            =   900
      TabIndex        =   80
      Text            =   "1"
      Top             =   3480
      Visible         =   0   'False
      Width           =   732
   End
   Begin VB.TextBox txtSeparationFactorsRow8 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      Height          =   288
      Index           =   10
      Left            =   7920
      TabIndex        =   79
      Text            =   "1"
      Top             =   3120
      Visible         =   0   'False
      Width           =   732
   End
   Begin VB.TextBox txtSeparationFactorsRow8 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      Height          =   288
      Index           =   9
      Left            =   7140
      TabIndex        =   78
      Text            =   "1"
      Top             =   3120
      Visible         =   0   'False
      Width           =   732
   End
   Begin VB.TextBox txtSeparationFactorsRow8 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      Enabled         =   0   'False
      Height          =   288
      Index           =   8
      Left            =   6360
      TabIndex        =   77
      Text            =   "1.00"
      Top             =   3120
      Visible         =   0   'False
      Width           =   732
   End
   Begin VB.TextBox txtSeparationFactorsRow8 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      Height          =   288
      Index           =   7
      Left            =   5580
      TabIndex        =   76
      Text            =   "1"
      Top             =   3120
      Visible         =   0   'False
      Width           =   732
   End
   Begin VB.TextBox txtSeparationFactorsRow8 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      Height          =   288
      Index           =   6
      Left            =   4800
      TabIndex        =   75
      Text            =   "1"
      Top             =   3120
      Visible         =   0   'False
      Width           =   732
   End
   Begin VB.TextBox txtSeparationFactorsRow8 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      Height          =   288
      Index           =   5
      Left            =   4020
      TabIndex        =   74
      Text            =   "1"
      Top             =   3120
      Visible         =   0   'False
      Width           =   732
   End
   Begin VB.TextBox txtSeparationFactorsRow8 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      Height          =   288
      Index           =   4
      Left            =   3240
      TabIndex        =   73
      Text            =   "1"
      Top             =   3120
      Visible         =   0   'False
      Width           =   732
   End
   Begin VB.TextBox txtSeparationFactorsRow8 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      Height          =   288
      Index           =   3
      Left            =   2460
      TabIndex        =   72
      Text            =   "1"
      Top             =   3120
      Visible         =   0   'False
      Width           =   732
   End
   Begin VB.TextBox txtSeparationFactorsRow8 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      Height          =   288
      Index           =   2
      Left            =   1680
      TabIndex        =   71
      Text            =   "1"
      Top             =   3120
      Visible         =   0   'False
      Width           =   732
   End
   Begin VB.TextBox txtSeparationFactorsRow8 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      Height          =   288
      Index           =   1
      Left            =   900
      TabIndex        =   70
      Text            =   "1"
      Top             =   3120
      Visible         =   0   'False
      Width           =   732
   End
   Begin VB.TextBox txtSeparationFactorsRow7 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      Height          =   288
      Index           =   10
      Left            =   7920
      TabIndex        =   69
      Text            =   "1"
      Top             =   2760
      Visible         =   0   'False
      Width           =   732
   End
   Begin VB.TextBox txtSeparationFactorsRow7 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      Height          =   288
      Index           =   9
      Left            =   7140
      TabIndex        =   68
      Text            =   "1"
      Top             =   2760
      Visible         =   0   'False
      Width           =   732
   End
   Begin VB.TextBox txtSeparationFactorsRow7 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      Height          =   288
      Index           =   8
      Left            =   6360
      TabIndex        =   67
      Text            =   "1"
      Top             =   2760
      Visible         =   0   'False
      Width           =   732
   End
   Begin VB.TextBox txtSeparationFactorsRow7 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      Enabled         =   0   'False
      Height          =   288
      Index           =   7
      Left            =   5580
      TabIndex        =   66
      Text            =   "1.00"
      Top             =   2760
      Visible         =   0   'False
      Width           =   732
   End
   Begin VB.TextBox txtSeparationFactorsRow7 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      Height          =   288
      Index           =   6
      Left            =   4800
      TabIndex        =   65
      Text            =   "1"
      Top             =   2760
      Visible         =   0   'False
      Width           =   732
   End
   Begin VB.TextBox txtSeparationFactorsRow7 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      Height          =   288
      Index           =   5
      Left            =   4020
      TabIndex        =   64
      Text            =   "1"
      Top             =   2760
      Visible         =   0   'False
      Width           =   732
   End
   Begin VB.TextBox txtSeparationFactorsRow7 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      Height          =   288
      Index           =   4
      Left            =   3240
      TabIndex        =   63
      Text            =   "1"
      Top             =   2760
      Visible         =   0   'False
      Width           =   732
   End
   Begin VB.TextBox txtSeparationFactorsRow7 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      Height          =   288
      Index           =   3
      Left            =   2460
      TabIndex        =   62
      Text            =   "1"
      Top             =   2760
      Visible         =   0   'False
      Width           =   732
   End
   Begin VB.TextBox txtSeparationFactorsRow7 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      Height          =   288
      Index           =   2
      Left            =   1680
      TabIndex        =   61
      Text            =   "1"
      Top             =   2760
      Visible         =   0   'False
      Width           =   732
   End
   Begin VB.TextBox txtSeparationFactorsRow7 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      Height          =   288
      Index           =   1
      Left            =   900
      TabIndex        =   60
      Text            =   "1"
      Top             =   2760
      Visible         =   0   'False
      Width           =   732
   End
   Begin VB.TextBox txtSeparationFactorsRow6 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      Height          =   288
      Index           =   10
      Left            =   7920
      TabIndex        =   59
      Text            =   "1"
      Top             =   2400
      Visible         =   0   'False
      Width           =   732
   End
   Begin VB.TextBox txtSeparationFactorsRow6 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      Height          =   288
      Index           =   9
      Left            =   7140
      TabIndex        =   58
      Text            =   "1"
      Top             =   2400
      Visible         =   0   'False
      Width           =   732
   End
   Begin VB.TextBox txtSeparationFactorsRow6 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      Height          =   288
      Index           =   8
      Left            =   6360
      TabIndex        =   57
      Text            =   "1"
      Top             =   2400
      Visible         =   0   'False
      Width           =   732
   End
   Begin VB.TextBox txtSeparationFactorsRow6 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      Height          =   288
      Index           =   7
      Left            =   5580
      TabIndex        =   56
      Text            =   "1"
      Top             =   2400
      Visible         =   0   'False
      Width           =   732
   End
   Begin VB.TextBox txtSeparationFactorsRow6 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      Enabled         =   0   'False
      Height          =   288
      Index           =   6
      Left            =   4800
      TabIndex        =   55
      Text            =   "1.00"
      Top             =   2400
      Visible         =   0   'False
      Width           =   732
   End
   Begin VB.TextBox txtSeparationFactorsRow6 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      Height          =   288
      Index           =   5
      Left            =   4020
      TabIndex        =   54
      Text            =   "1"
      Top             =   2400
      Visible         =   0   'False
      Width           =   732
   End
   Begin VB.TextBox txtSeparationFactorsRow6 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      Height          =   288
      Index           =   4
      Left            =   3240
      TabIndex        =   53
      Text            =   "1"
      Top             =   2400
      Visible         =   0   'False
      Width           =   732
   End
   Begin VB.TextBox txtSeparationFactorsRow6 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      Height          =   288
      Index           =   3
      Left            =   2460
      TabIndex        =   52
      Text            =   "1"
      Top             =   2400
      Visible         =   0   'False
      Width           =   732
   End
   Begin VB.TextBox txtSeparationFactorsRow6 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      Height          =   288
      Index           =   2
      Left            =   1680
      TabIndex        =   51
      Text            =   "1"
      Top             =   2400
      Visible         =   0   'False
      Width           =   732
   End
   Begin VB.TextBox txtSeparationFactorsRow6 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      Height          =   288
      Index           =   1
      Left            =   900
      TabIndex        =   50
      Text            =   "1"
      Top             =   2400
      Visible         =   0   'False
      Width           =   732
   End
   Begin VB.TextBox txtSeparationFactorsRow5 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      Height          =   288
      Index           =   10
      Left            =   7920
      TabIndex        =   49
      Text            =   "1"
      Top             =   2040
      Visible         =   0   'False
      Width           =   732
   End
   Begin VB.TextBox txtSeparationFactorsRow5 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      Height          =   288
      Index           =   9
      Left            =   7140
      TabIndex        =   48
      Text            =   "1"
      Top             =   2040
      Visible         =   0   'False
      Width           =   732
   End
   Begin VB.TextBox txtSeparationFactorsRow5 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      Height          =   288
      Index           =   8
      Left            =   6360
      TabIndex        =   47
      Text            =   "1"
      Top             =   2040
      Visible         =   0   'False
      Width           =   732
   End
   Begin VB.TextBox txtSeparationFactorsRow5 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      Height          =   288
      Index           =   7
      Left            =   5580
      TabIndex        =   46
      Text            =   "1"
      Top             =   2040
      Visible         =   0   'False
      Width           =   732
   End
   Begin VB.TextBox txtSeparationFactorsRow5 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      Height          =   288
      Index           =   6
      Left            =   4800
      TabIndex        =   45
      Text            =   "1"
      Top             =   2040
      Visible         =   0   'False
      Width           =   732
   End
   Begin VB.TextBox txtSeparationFactorsRow5 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      Enabled         =   0   'False
      Height          =   288
      Index           =   5
      Left            =   4020
      TabIndex        =   44
      Text            =   "1.00"
      Top             =   2040
      Visible         =   0   'False
      Width           =   732
   End
   Begin VB.TextBox txtSeparationFactorsRow5 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      Height          =   288
      Index           =   4
      Left            =   3240
      TabIndex        =   43
      Text            =   "1"
      Top             =   2040
      Visible         =   0   'False
      Width           =   732
   End
   Begin VB.TextBox txtSeparationFactorsRow5 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      Height          =   288
      Index           =   3
      Left            =   2460
      TabIndex        =   42
      Text            =   "1"
      Top             =   2040
      Visible         =   0   'False
      Width           =   732
   End
   Begin VB.TextBox txtSeparationFactorsRow5 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      Height          =   288
      Index           =   2
      Left            =   1680
      TabIndex        =   41
      Text            =   "1"
      Top             =   2040
      Visible         =   0   'False
      Width           =   732
   End
   Begin VB.TextBox txtSeparationFactorsRow5 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      Height          =   288
      Index           =   1
      Left            =   900
      TabIndex        =   40
      Text            =   "1"
      Top             =   2040
      Visible         =   0   'False
      Width           =   732
   End
   Begin VB.TextBox txtSeparationFactorsRow4 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      Height          =   288
      Index           =   10
      Left            =   7920
      TabIndex        =   39
      Text            =   "1"
      Top             =   1680
      Visible         =   0   'False
      Width           =   732
   End
   Begin VB.TextBox txtSeparationFactorsRow4 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      Height          =   288
      Index           =   9
      Left            =   7140
      TabIndex        =   38
      Text            =   "1"
      Top             =   1680
      Visible         =   0   'False
      Width           =   732
   End
   Begin VB.TextBox txtSeparationFactorsRow4 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      Height          =   288
      Index           =   8
      Left            =   6360
      TabIndex        =   37
      Text            =   "1"
      Top             =   1680
      Visible         =   0   'False
      Width           =   732
   End
   Begin VB.TextBox txtSeparationFactorsRow4 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      Height          =   288
      Index           =   7
      Left            =   5580
      TabIndex        =   36
      Text            =   "1"
      Top             =   1680
      Visible         =   0   'False
      Width           =   732
   End
   Begin VB.TextBox txtSeparationFactorsRow4 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      Height          =   288
      Index           =   6
      Left            =   4800
      TabIndex        =   35
      Text            =   "1"
      Top             =   1680
      Visible         =   0   'False
      Width           =   732
   End
   Begin VB.TextBox txtSeparationFactorsRow4 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      Height          =   288
      Index           =   5
      Left            =   4020
      TabIndex        =   34
      Text            =   "1"
      Top             =   1680
      Visible         =   0   'False
      Width           =   732
   End
   Begin VB.TextBox txtSeparationFactorsRow4 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      Enabled         =   0   'False
      Height          =   288
      Index           =   4
      Left            =   3240
      TabIndex        =   33
      Text            =   "1.00"
      Top             =   1680
      Visible         =   0   'False
      Width           =   732
   End
   Begin VB.TextBox txtSeparationFactorsRow4 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      Height          =   288
      Index           =   3
      Left            =   2460
      TabIndex        =   32
      Text            =   "1"
      Top             =   1680
      Visible         =   0   'False
      Width           =   732
   End
   Begin VB.TextBox txtSeparationFactorsRow4 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      Height          =   288
      Index           =   2
      Left            =   1680
      TabIndex        =   31
      Text            =   "1"
      Top             =   1680
      Visible         =   0   'False
      Width           =   732
   End
   Begin VB.TextBox txtSeparationFactorsRow4 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      Height          =   288
      Index           =   1
      Left            =   900
      TabIndex        =   30
      Text            =   "1"
      Top             =   1680
      Visible         =   0   'False
      Width           =   732
   End
   Begin VB.TextBox txtSeparationFactorsRow3 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      Height          =   288
      Index           =   10
      Left            =   7920
      TabIndex        =   29
      Text            =   "1"
      Top             =   1320
      Visible         =   0   'False
      Width           =   732
   End
   Begin VB.TextBox txtSeparationFactorsRow3 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      Height          =   288
      Index           =   9
      Left            =   7140
      TabIndex        =   28
      Text            =   "1"
      Top             =   1320
      Visible         =   0   'False
      Width           =   732
   End
   Begin VB.TextBox txtSeparationFactorsRow3 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      Height          =   288
      Index           =   8
      Left            =   6360
      TabIndex        =   27
      Text            =   "1"
      Top             =   1320
      Visible         =   0   'False
      Width           =   732
   End
   Begin VB.TextBox txtSeparationFactorsRow3 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      Height          =   288
      Index           =   7
      Left            =   5580
      TabIndex        =   26
      Text            =   "1"
      Top             =   1320
      Visible         =   0   'False
      Width           =   732
   End
   Begin VB.TextBox txtSeparationFactorsRow3 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      Height          =   288
      Index           =   6
      Left            =   4800
      TabIndex        =   25
      Text            =   "1"
      Top             =   1320
      Visible         =   0   'False
      Width           =   732
   End
   Begin VB.TextBox txtSeparationFactorsRow3 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      Height          =   288
      Index           =   5
      Left            =   4020
      TabIndex        =   24
      Text            =   "1"
      Top             =   1320
      Visible         =   0   'False
      Width           =   732
   End
   Begin VB.TextBox txtSeparationFactorsRow3 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      Enabled         =   0   'False
      Height          =   288
      Index           =   3
      Left            =   2460
      TabIndex        =   22
      Text            =   "1.00"
      Top             =   1320
      Visible         =   0   'False
      Width           =   732
   End
   Begin VB.TextBox txtSeparationFactorsRow3 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      Height          =   288
      Index           =   2
      Left            =   1680
      TabIndex        =   21
      Text            =   "1"
      Top             =   1320
      Visible         =   0   'False
      Width           =   732
   End
   Begin VB.TextBox txtSeparationFactorsRow3 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      Height          =   288
      Index           =   1
      Left            =   900
      TabIndex        =   20
      Text            =   "1"
      Top             =   1320
      Visible         =   0   'False
      Width           =   732
   End
   Begin VB.TextBox txtSeparationFactorsRow2 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      Height          =   288
      Index           =   10
      Left            =   7920
      TabIndex        =   19
      Text            =   "1"
      Top             =   960
      Visible         =   0   'False
      Width           =   732
   End
   Begin VB.TextBox txtSeparationFactorsRow2 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      Height          =   288
      Index           =   9
      Left            =   7140
      TabIndex        =   18
      Text            =   "1"
      Top             =   960
      Visible         =   0   'False
      Width           =   732
   End
   Begin VB.TextBox txtSeparationFactorsRow2 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      Height          =   288
      Index           =   8
      Left            =   6360
      TabIndex        =   17
      Text            =   "1"
      Top             =   960
      Visible         =   0   'False
      Width           =   732
   End
   Begin VB.TextBox txtSeparationFactorsRow2 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      Height          =   288
      Index           =   7
      Left            =   5580
      TabIndex        =   16
      Text            =   "1"
      Top             =   960
      Visible         =   0   'False
      Width           =   732
   End
   Begin VB.TextBox txtSeparationFactorsRow2 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      Height          =   288
      Index           =   6
      Left            =   4800
      TabIndex        =   15
      Text            =   "1"
      Top             =   960
      Visible         =   0   'False
      Width           =   732
   End
   Begin VB.TextBox txtSeparationFactorsRow2 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      Height          =   288
      Index           =   5
      Left            =   4020
      TabIndex        =   14
      Text            =   "1"
      Top             =   960
      Visible         =   0   'False
      Width           =   732
   End
   Begin VB.TextBox txtSeparationFactorsRow2 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      Height          =   288
      Index           =   4
      Left            =   3240
      TabIndex        =   13
      Text            =   "1"
      Top             =   960
      Visible         =   0   'False
      Width           =   732
   End
   Begin VB.TextBox txtSeparationFactorsRow2 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      Height          =   288
      Index           =   3
      Left            =   2460
      TabIndex        =   12
      Text            =   "1"
      Top             =   960
      Visible         =   0   'False
      Width           =   732
   End
   Begin VB.TextBox txtSeparationFactorsRow2 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      Enabled         =   0   'False
      Height          =   288
      Index           =   2
      Left            =   1680
      TabIndex        =   11
      Text            =   "1.00"
      Top             =   960
      Visible         =   0   'False
      Width           =   732
   End
   Begin VB.TextBox txtSeparationFactorsRow2 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      Height          =   288
      Index           =   1
      Left            =   900
      TabIndex        =   10
      Text            =   "1"
      Top             =   960
      Visible         =   0   'False
      Width           =   732
   End
   Begin VB.TextBox txtSeparationFactorsRow1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      Height          =   288
      Index           =   10
      Left            =   7920
      TabIndex        =   9
      Text            =   "1"
      Top             =   600
      Visible         =   0   'False
      Width           =   732
   End
   Begin VB.TextBox txtSeparationFactorsRow1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      Height          =   288
      Index           =   9
      Left            =   7140
      TabIndex        =   8
      Text            =   "1"
      Top             =   600
      Visible         =   0   'False
      Width           =   732
   End
   Begin VB.TextBox txtSeparationFactorsRow1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      Height          =   288
      Index           =   8
      Left            =   6360
      TabIndex        =   7
      Text            =   "1"
      Top             =   600
      Visible         =   0   'False
      Width           =   732
   End
   Begin VB.TextBox txtSeparationFactorsRow1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      Height          =   288
      Index           =   7
      Left            =   5580
      TabIndex        =   6
      Text            =   "1"
      Top             =   600
      Visible         =   0   'False
      Width           =   732
   End
   Begin VB.TextBox txtSeparationFactorsRow1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      Height          =   288
      Index           =   6
      Left            =   4800
      TabIndex        =   5
      Text            =   "1"
      Top             =   600
      Visible         =   0   'False
      Width           =   732
   End
   Begin VB.TextBox txtSeparationFactorsRow1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      Height          =   288
      Index           =   5
      Left            =   4020
      TabIndex        =   4
      Text            =   "1"
      Top             =   600
      Visible         =   0   'False
      Width           =   732
   End
   Begin VB.TextBox txtSeparationFactorsRow1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      Height          =   288
      Index           =   4
      Left            =   3240
      TabIndex        =   3
      Text            =   "1"
      Top             =   600
      Visible         =   0   'False
      Width           =   732
   End
   Begin VB.TextBox txtSeparationFactorsRow1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      Height          =   288
      Index           =   3
      Left            =   2460
      TabIndex        =   2
      Text            =   "1"
      Top             =   600
      Visible         =   0   'False
      Width           =   732
   End
   Begin VB.TextBox txtSeparationFactorsRow1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      Height          =   288
      Index           =   2
      Left            =   1680
      TabIndex        =   1
      Text            =   "1"
      Top             =   600
      Visible         =   0   'False
      Width           =   732
   End
   Begin VB.TextBox txtSeparationFactorsRow1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      Enabled         =   0   'False
      Height          =   288
      Index           =   1
      Left            =   900
      TabIndex        =   0
      Text            =   "1.00"
      Top             =   600
      Visible         =   0   'False
      Width           =   732
   End
   Begin Threed.SSOption optSeparationFactors 
      Height          =   255
      Index           =   2
      Left            =   1920
      TabIndex        =   125
      Top             =   4200
      Width           =   255
      _Version        =   65536
      _ExtentX        =   450
      _ExtentY        =   450
      _StockProps     =   78
      Caption         =   "SSOption1"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin Threed.SSOption optSeparationFactors 
      Height          =   255
      Index           =   3
      Left            =   2760
      TabIndex        =   126
      Top             =   4200
      Width           =   255
      _Version        =   65536
      _ExtentX        =   450
      _ExtentY        =   450
      _StockProps     =   78
      Caption         =   "SSOption1"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin Threed.SSOption optSeparationFactors 
      Height          =   255
      Index           =   4
      Left            =   3480
      TabIndex        =   127
      Top             =   4200
      Width           =   255
      _Version        =   65536
      _ExtentX        =   450
      _ExtentY        =   450
      _StockProps     =   78
      Caption         =   "SSOption1"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin Threed.SSOption optSeparationFactors 
      Height          =   255
      Index           =   5
      Left            =   4320
      TabIndex        =   128
      Top             =   4200
      Width           =   255
      _Version        =   65536
      _ExtentX        =   450
      _ExtentY        =   450
      _StockProps     =   78
      Caption         =   "SSOption1"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin Threed.SSOption optSeparationFactors 
      Height          =   255
      Index           =   6
      Left            =   5040
      TabIndex        =   129
      Top             =   4200
      Width           =   255
      _Version        =   65536
      _ExtentX        =   450
      _ExtentY        =   450
      _StockProps     =   78
      Caption         =   "SSOption1"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin Threed.SSOption optSeparationFactors 
      Height          =   255
      Index           =   7
      Left            =   5880
      TabIndex        =   130
      Top             =   4200
      Width           =   255
      _Version        =   65536
      _ExtentX        =   450
      _ExtentY        =   450
      _StockProps     =   78
      Caption         =   "SSOption1"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin Threed.SSOption optSeparationFactors 
      Height          =   255
      Index           =   8
      Left            =   6600
      TabIndex        =   131
      Top             =   4200
      Width           =   255
      _Version        =   65536
      _ExtentX        =   450
      _ExtentY        =   450
      _StockProps     =   78
      Caption         =   "SSOption1"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin Threed.SSOption optSeparationFactors 
      Height          =   255
      Index           =   9
      Left            =   7320
      TabIndex        =   132
      Top             =   4200
      Width           =   255
      _Version        =   65536
      _ExtentX        =   450
      _ExtentY        =   450
      _StockProps     =   78
      Caption         =   "SSOption1"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin Threed.SSOption optSeparationFactors 
      Height          =   255
      Index           =   10
      Left            =   8160
      TabIndex        =   133
      Top             =   4200
      Width           =   255
      _Version        =   65536
      _ExtentX        =   450
      _ExtentY        =   450
      _StockProps     =   78
      Caption         =   "SSOption1"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin Threed.SSOption optSeparationFactors 
      Height          =   255
      Index           =   11
      Left            =   8760
      TabIndex        =   134
      Top             =   600
      Width           =   255
      _Version        =   65536
      _ExtentX        =   450
      _ExtentY        =   450
      _StockProps     =   78
      Caption         =   "SSOption1"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin Threed.SSOption optSeparationFactors 
      Height          =   255
      Index           =   12
      Left            =   8760
      TabIndex        =   135
      Top             =   960
      Width           =   255
      _Version        =   65536
      _ExtentX        =   450
      _ExtentY        =   450
      _StockProps     =   78
      Caption         =   "SSOption1"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin Threed.SSOption optSeparationFactors 
      Height          =   255
      Index           =   13
      Left            =   8760
      TabIndex        =   136
      Top             =   1320
      Width           =   255
      _Version        =   65536
      _ExtentX        =   450
      _ExtentY        =   450
      _StockProps     =   78
      Caption         =   "SSOption1"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin Threed.SSOption optSeparationFactors 
      Height          =   255
      Index           =   14
      Left            =   8760
      TabIndex        =   137
      Top             =   1680
      Width           =   255
      _Version        =   65536
      _ExtentX        =   450
      _ExtentY        =   450
      _StockProps     =   78
      Caption         =   "SSOption1"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin Threed.SSOption optSeparationFactors 
      Height          =   255
      Index           =   15
      Left            =   8760
      TabIndex        =   138
      Top             =   2040
      Width           =   255
      _Version        =   65536
      _ExtentX        =   450
      _ExtentY        =   450
      _StockProps     =   78
      Caption         =   "SSOption1"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin Threed.SSOption optSeparationFactors 
      Height          =   255
      Index           =   16
      Left            =   8760
      TabIndex        =   139
      Top             =   2400
      Width           =   255
      _Version        =   65536
      _ExtentX        =   450
      _ExtentY        =   450
      _StockProps     =   78
      Caption         =   "SSOption1"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin Threed.SSOption optSeparationFactors 
      Height          =   255
      Index           =   17
      Left            =   8760
      TabIndex        =   140
      Top             =   2760
      Width           =   255
      _Version        =   65536
      _ExtentX        =   450
      _ExtentY        =   450
      _StockProps     =   78
      Caption         =   "SSOption1"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin Threed.SSOption optSeparationFactors 
      Height          =   255
      Index           =   18
      Left            =   8760
      TabIndex        =   141
      Top             =   3120
      Width           =   255
      _Version        =   65536
      _ExtentX        =   450
      _ExtentY        =   450
      _StockProps     =   78
      Caption         =   "SSOption1"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin Threed.SSOption optSeparationFactors 
      Height          =   255
      Index           =   19
      Left            =   8760
      TabIndex        =   142
      Top             =   3480
      Width           =   255
      _Version        =   65536
      _ExtentX        =   450
      _ExtentY        =   450
      _StockProps     =   78
      Caption         =   "SSOption1"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin Threed.SSOption optSeparationFactors 
      Height          =   255
      Index           =   20
      Left            =   8760
      TabIndex        =   143
      Top             =   3840
      Width           =   255
      _Version        =   65536
      _ExtentX        =   450
      _ExtentY        =   450
      _StockProps     =   78
      Caption         =   "SSOption1"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Label lblheaderj 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00808000&
      Caption         =   "j"
      ForeColor       =   &H00FFFFFF&
      Height          =   252
      Left            =   120
      TabIndex        =   122
      Top             =   2280
      Width           =   276
   End
   Begin VB.Label lblheaderi 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00800000&
      Caption         =   "i"
      ForeColor       =   &H00FFFFFF&
      Height          =   252
      Left            =   4560
      TabIndex        =   121
      Top             =   60
      Width           =   312
   End
   Begin VB.Label lblj 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H00808000&
      Caption         =   "10"
      ForeColor       =   &H00FFFFFF&
      Height          =   192
      Index           =   10
      Left            =   540
      TabIndex        =   120
      Top             =   3900
      Visible         =   0   'False
      Width           =   276
   End
   Begin VB.Label lblj 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H00808000&
      Caption         =   "9"
      ForeColor       =   &H00FFFFFF&
      Height          =   192
      Index           =   9
      Left            =   540
      TabIndex        =   119
      Top             =   3540
      Visible         =   0   'False
      Width           =   276
   End
   Begin VB.Label lblj 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H00808000&
      Caption         =   "8"
      ForeColor       =   &H00FFFFFF&
      Height          =   192
      Index           =   8
      Left            =   540
      TabIndex        =   118
      Top             =   3180
      Visible         =   0   'False
      Width           =   276
   End
   Begin VB.Label lblj 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H00808000&
      Caption         =   "7"
      ForeColor       =   &H00FFFFFF&
      Height          =   192
      Index           =   7
      Left            =   540
      TabIndex        =   117
      Top             =   2820
      Visible         =   0   'False
      Width           =   276
   End
   Begin VB.Label lblj 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H00808000&
      Caption         =   "6"
      ForeColor       =   &H00FFFFFF&
      Height          =   192
      Index           =   6
      Left            =   540
      TabIndex        =   116
      Top             =   2460
      Visible         =   0   'False
      Width           =   276
   End
   Begin VB.Label lblj 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H00808000&
      Caption         =   "5"
      ForeColor       =   &H00FFFFFF&
      Height          =   192
      Index           =   5
      Left            =   540
      TabIndex        =   115
      Top             =   2100
      Visible         =   0   'False
      Width           =   276
   End
   Begin VB.Label lblj 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H00808000&
      Caption         =   "4"
      ForeColor       =   &H00FFFFFF&
      Height          =   192
      Index           =   4
      Left            =   540
      TabIndex        =   114
      Top             =   1740
      Visible         =   0   'False
      Width           =   276
   End
   Begin VB.Label lblj 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H00808000&
      Caption         =   "3"
      ForeColor       =   &H00FFFFFF&
      Height          =   192
      Index           =   3
      Left            =   540
      TabIndex        =   113
      Top             =   1380
      Visible         =   0   'False
      Width           =   276
   End
   Begin VB.Label lblj 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H00808000&
      Caption         =   "2"
      ForeColor       =   &H00FFFFFF&
      Height          =   192
      Index           =   2
      Left            =   540
      TabIndex        =   112
      Top             =   1020
      Visible         =   0   'False
      Width           =   276
   End
   Begin VB.Label lblj 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H00808000&
      Caption         =   "1"
      ForeColor       =   &H00FFFFFF&
      Height          =   192
      Index           =   1
      Left            =   540
      TabIndex        =   111
      Top             =   660
      Visible         =   0   'False
      Width           =   276
   End
   Begin VB.Label lbli 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00800000&
      Caption         =   "10"
      ForeColor       =   &H00FFFFFF&
      Height          =   192
      Index           =   10
      Left            =   7920
      TabIndex        =   110
      Top             =   360
      Visible         =   0   'False
      Width           =   732
   End
   Begin VB.Label lbli 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00800000&
      Caption         =   "9"
      ForeColor       =   &H00FFFFFF&
      Height          =   192
      Index           =   9
      Left            =   7140
      TabIndex        =   109
      Top             =   360
      Visible         =   0   'False
      Width           =   732
   End
   Begin VB.Label lbli 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00800000&
      Caption         =   "8"
      ForeColor       =   &H00FFFFFF&
      Height          =   192
      Index           =   8
      Left            =   6360
      TabIndex        =   108
      Top             =   360
      Visible         =   0   'False
      Width           =   732
   End
   Begin VB.Label lbli 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00800000&
      Caption         =   "7"
      ForeColor       =   &H00FFFFFF&
      Height          =   192
      Index           =   7
      Left            =   5580
      TabIndex        =   107
      Top             =   360
      Visible         =   0   'False
      Width           =   732
   End
   Begin VB.Label lbli 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00800000&
      Caption         =   "6"
      ForeColor       =   &H00FFFFFF&
      Height          =   192
      Index           =   6
      Left            =   4800
      TabIndex        =   106
      Top             =   360
      Visible         =   0   'False
      Width           =   732
   End
   Begin VB.Label lbli 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00800000&
      Caption         =   "5"
      ForeColor       =   &H00FFFFFF&
      Height          =   192
      Index           =   5
      Left            =   4020
      TabIndex        =   105
      Top             =   360
      Visible         =   0   'False
      Width           =   732
   End
   Begin VB.Label lbli 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00800000&
      Caption         =   "4"
      ForeColor       =   &H00FFFFFF&
      Height          =   192
      Index           =   4
      Left            =   3240
      TabIndex        =   104
      Top             =   360
      Visible         =   0   'False
      Width           =   732
   End
   Begin VB.Label lbli 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00800000&
      Caption         =   "3"
      ForeColor       =   &H00FFFFFF&
      Height          =   192
      Index           =   3
      Left            =   2460
      TabIndex        =   103
      Top             =   360
      Visible         =   0   'False
      Width           =   732
   End
   Begin VB.Label lbli 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00800000&
      Caption         =   "2"
      ForeColor       =   &H00FFFFFF&
      Height          =   192
      Index           =   2
      Left            =   1680
      TabIndex        =   101
      Top             =   360
      Visible         =   0   'False
      Width           =   732
   End
   Begin VB.Label lbli 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00800000&
      Caption         =   "1"
      ForeColor       =   &H00FFFFFF&
      Height          =   192
      Index           =   1
      Left            =   900
      TabIndex        =   100
      Top             =   360
      Visible         =   0   'False
      Width           =   732
   End
End
Attribute VB_Name = "frmSeparationFactors"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

Dim Temp_Text As String

Private Sub cmdCancel_Click()
    Dim i As Integer

    optSeparationFactors(SeparationFactorInput.Value) = False
    Select Case OldOptionButtonSeparationFactors
       Case 1 To 10
          SeparationFactorInput.Row = False
       Case 11 To 20
          SeparationFactorInput.Row = True
    End Select
    SeparationFactorInput.Value = OldOptionButtonSeparationFactors

    For i = 1 To MAX_CHEMICAL
        OneDimSeparationFactors(i) = OldOneDimSeparationFactors(i)
    Next i

    Call CalculateSeparationFactors
    Call ShowSeparationFactors

    optSeparationFactors(SeparationFactorInput.Value).Value = True

    frmSeparationFactors.Hide

End Sub

Private Sub cmdOK_Click()
    Dim i As Integer, j As Integer
    Dim ValueToDisplay As Double

    For i = 1 To NumberOfIons
        If AddingCation Or EditingCation Then
           NowProj.cation(i).SeparationFactor = OneDimSeparationFactors(i)
        ElseIf AddingAnion Or EditingAnion Then
           NowProj.Anion(i).SeparationFactor = OneDimSeparationFactors(i)
        End If
    Next i

    'Set lblAlpha label boxes appropriately on frmAddComponent
    If SeparationFactorInput.Value <> OldOptionButtonSeparationFactors Then
       If SeparationFactorInput.Row = True Then
          Select Case OldOptionButtonSeparationFactors
             Case 1 To 10   'Were Entering Down a Column Before but now entering across a row
                frmAddComponent!lblAlpha(1).Caption = frmAddComponent!lblAlpha(2).Caption
          End Select
          If AddingCation Or EditingCation Then
             frmAddComponent!lblAlpha(2).Caption = Trim$(NowProj.cation(SeparationFactorInput.Value - 10).Name)
          ElseIf AddingAnion Or EditingAnion Then
             frmAddComponent!lblAlpha(2).Caption = Trim$(NowProj.Anion(SeparationFactorInput.Value - 10).Name)
          End If
       Else
          Select Case OldOptionButtonSeparationFactors
             Case 11 To 20   'Were Entering Across a Row Before but Now entering down a column
                frmAddComponent!lblAlpha(2).Caption = frmAddComponent!lblAlpha(1).Caption
          End Select
          If AddingCation Or EditingCation Then
             frmAddComponent!lblAlpha(1).Caption = Trim$(NowProj.cation(SeparationFactorInput.Value).Name)
          ElseIf AddingAnion Or EditingAnion Then
             frmAddComponent!lblAlpha(1).Caption = Trim$(NowProj.Anion(SeparationFactorInput.Value).Name)
          End If
       End If
    End If

    ValueToDisplay = OneDimSeparationFactors(NumberOfIonToEdit)
    frmAddComponent!txtAlphaValue.Text = Format$(ValueToDisplay, GetTheFormat(ValueToDisplay))

    frmSeparationFactors.Hide

End Sub

Private Sub DisableAllTextBoxes()
    Dim i As Integer

    For i = 1 To 10
        txtSeparationFactorsRow1(i).Enabled = False
        txtSeparationFactorsRow2(i).Enabled = False
        txtSeparationFactorsRow3(i).Enabled = False
        txtSeparationFactorsRow4(i).Enabled = False
        txtSeparationFactorsRow5(i).Enabled = False
        txtSeparationFactorsRow6(i).Enabled = False
        txtSeparationFactorsRow7(i).Enabled = False
        txtSeparationFactorsRow8(i).Enabled = False
        txtSeparationFactorsRow9(i).Enabled = False
        txtSeparationFactorsRow10(i).Enabled = False
    Next i

End Sub

Private Sub DisplayDownAColumn(Column As Integer)

         Select Case NumberOfIons
           Case 1
              txtSeparationFactorsRow1(Column).Text = Format$(Ion_Array(1).SeparationFactor, GetTheFormat(Ion_Array(1).SeparationFactor))

           Case 2
              txtSeparationFactorsRow1(Column).Text = Format$(Ion_Array(1).SeparationFactor, GetTheFormat(Ion_Array(1).SeparationFactor))
              txtSeparationFactorsRow2(Column).Text = Format$(Ion_Array(2).SeparationFactor, GetTheFormat(Ion_Array(2).SeparationFactor))

           Case 3
              txtSeparationFactorsRow1(Column).Text = Format$(Ion_Array(1).SeparationFactor, GetTheFormat(Ion_Array(1).SeparationFactor))
              txtSeparationFactorsRow2(Column).Text = Format$(Ion_Array(2).SeparationFactor, GetTheFormat(Ion_Array(2).SeparationFactor))
              txtSeparationFactorsRow3(Column).Text = Format$(Ion_Array(3).SeparationFactor, GetTheFormat(Ion_Array(3).SeparationFactor))

           Case 4
              txtSeparationFactorsRow1(Column).Text = Format$(Ion_Array(1).SeparationFactor, GetTheFormat(Ion_Array(1).SeparationFactor))
              txtSeparationFactorsRow2(Column).Text = Format$(Ion_Array(2).SeparationFactor, GetTheFormat(Ion_Array(2).SeparationFactor))
              txtSeparationFactorsRow3(Column).Text = Format$(Ion_Array(3).SeparationFactor, GetTheFormat(Ion_Array(3).SeparationFactor))
              txtSeparationFactorsRow4(Column).Text = Format$(Ion_Array(4).SeparationFactor, GetTheFormat(Ion_Array(4).SeparationFactor))

           Case 5
              txtSeparationFactorsRow1(Column).Text = Format$(Ion_Array(1).SeparationFactor, GetTheFormat(Ion_Array(1).SeparationFactor))
              txtSeparationFactorsRow2(Column).Text = Format$(Ion_Array(2).SeparationFactor, GetTheFormat(Ion_Array(2).SeparationFactor))
              txtSeparationFactorsRow3(Column).Text = Format$(Ion_Array(3).SeparationFactor, GetTheFormat(Ion_Array(3).SeparationFactor))
              txtSeparationFactorsRow4(Column).Text = Format$(Ion_Array(4).SeparationFactor, GetTheFormat(Ion_Array(4).SeparationFactor))
              txtSeparationFactorsRow5(Column).Text = Format$(Ion_Array(5).SeparationFactor, GetTheFormat(Ion_Array(5).SeparationFactor))

           Case 6
              txtSeparationFactorsRow1(Column).Text = Format$(Ion_Array(1).SeparationFactor, GetTheFormat(Ion_Array(1).SeparationFactor))
              txtSeparationFactorsRow2(Column).Text = Format$(Ion_Array(2).SeparationFactor, GetTheFormat(Ion_Array(2).SeparationFactor))
              txtSeparationFactorsRow3(Column).Text = Format$(Ion_Array(3).SeparationFactor, GetTheFormat(Ion_Array(3).SeparationFactor))
              txtSeparationFactorsRow4(Column).Text = Format$(Ion_Array(4).SeparationFactor, GetTheFormat(Ion_Array(4).SeparationFactor))
              txtSeparationFactorsRow5(Column).Text = Format$(Ion_Array(5).SeparationFactor, GetTheFormat(Ion_Array(5).SeparationFactor))
              txtSeparationFactorsRow6(Column).Text = Format$(Ion_Array(6).SeparationFactor, GetTheFormat(Ion_Array(6).SeparationFactor))

           Case 7
              txtSeparationFactorsRow1(Column).Text = Format$(Ion_Array(1).SeparationFactor, GetTheFormat(Ion_Array(1).SeparationFactor))
              txtSeparationFactorsRow2(Column).Text = Format$(Ion_Array(2).SeparationFactor, GetTheFormat(Ion_Array(2).SeparationFactor))
              txtSeparationFactorsRow3(Column).Text = Format$(Ion_Array(3).SeparationFactor, GetTheFormat(Ion_Array(3).SeparationFactor))
              txtSeparationFactorsRow4(Column).Text = Format$(Ion_Array(4).SeparationFactor, GetTheFormat(Ion_Array(4).SeparationFactor))
              txtSeparationFactorsRow5(Column).Text = Format$(Ion_Array(5).SeparationFactor, GetTheFormat(Ion_Array(5).SeparationFactor))
              txtSeparationFactorsRow6(Column).Text = Format$(Ion_Array(6).SeparationFactor, GetTheFormat(Ion_Array(6).SeparationFactor))
              txtSeparationFactorsRow7(Column).Text = Format$(Ion_Array(7).SeparationFactor, GetTheFormat(Ion_Array(7).SeparationFactor))
           
           Case 8
              txtSeparationFactorsRow1(Column).Text = Format$(Ion_Array(1).SeparationFactor, GetTheFormat(Ion_Array(1).SeparationFactor))
              txtSeparationFactorsRow2(Column).Text = Format$(Ion_Array(2).SeparationFactor, GetTheFormat(Ion_Array(2).SeparationFactor))
              txtSeparationFactorsRow3(Column).Text = Format$(Ion_Array(3).SeparationFactor, GetTheFormat(Ion_Array(3).SeparationFactor))
              txtSeparationFactorsRow4(Column).Text = Format$(Ion_Array(4).SeparationFactor, GetTheFormat(Ion_Array(4).SeparationFactor))
              txtSeparationFactorsRow5(Column).Text = Format$(Ion_Array(5).SeparationFactor, GetTheFormat(Ion_Array(5).SeparationFactor))
              txtSeparationFactorsRow6(Column).Text = Format$(Ion_Array(6).SeparationFactor, GetTheFormat(Ion_Array(6).SeparationFactor))
              txtSeparationFactorsRow7(Column).Text = Format$(Ion_Array(7).SeparationFactor, GetTheFormat(Ion_Array(7).SeparationFactor))
              txtSeparationFactorsRow8(Column).Text = Format$(Ion_Array(8).SeparationFactor, GetTheFormat(Ion_Array(8).SeparationFactor))

           Case 9
              txtSeparationFactorsRow1(Column).Text = Format$(Ion_Array(1).SeparationFactor, GetTheFormat(Ion_Array(1).SeparationFactor))
              txtSeparationFactorsRow2(Column).Text = Format$(Ion_Array(2).SeparationFactor, GetTheFormat(Ion_Array(2).SeparationFactor))
              txtSeparationFactorsRow3(Column).Text = Format$(Ion_Array(3).SeparationFactor, GetTheFormat(Ion_Array(3).SeparationFactor))
              txtSeparationFactorsRow4(Column).Text = Format$(Ion_Array(4).SeparationFactor, GetTheFormat(Ion_Array(4).SeparationFactor))
              txtSeparationFactorsRow5(Column).Text = Format$(Ion_Array(5).SeparationFactor, GetTheFormat(Ion_Array(5).SeparationFactor))
              txtSeparationFactorsRow6(Column).Text = Format$(Ion_Array(6).SeparationFactor, GetTheFormat(Ion_Array(6).SeparationFactor))
              txtSeparationFactorsRow7(Column).Text = Format$(Ion_Array(7).SeparationFactor, GetTheFormat(Ion_Array(7).SeparationFactor))
              txtSeparationFactorsRow8(Column).Text = Format$(Ion_Array(8).SeparationFactor, GetTheFormat(Ion_Array(8).SeparationFactor))
              txtSeparationFactorsRow9(Column).Text = Format$(Ion_Array(9).SeparationFactor, GetTheFormat(Ion_Array(9).SeparationFactor))

           Case 10
              txtSeparationFactorsRow1(Column).Text = Format$(Ion_Array(1).SeparationFactor, GetTheFormat(Ion_Array(1).SeparationFactor))
              txtSeparationFactorsRow2(Column).Text = Format$(Ion_Array(2).SeparationFactor, GetTheFormat(Ion_Array(2).SeparationFactor))
              txtSeparationFactorsRow3(Column).Text = Format$(Ion_Array(3).SeparationFactor, GetTheFormat(Ion_Array(3).SeparationFactor))
              txtSeparationFactorsRow4(Column).Text = Format$(Ion_Array(4).SeparationFactor, GetTheFormat(Ion_Array(4).SeparationFactor))
              txtSeparationFactorsRow5(Column).Text = Format$(Ion_Array(5).SeparationFactor, GetTheFormat(Ion_Array(5).SeparationFactor))
              txtSeparationFactorsRow6(Column).Text = Format$(Ion_Array(6).SeparationFactor, GetTheFormat(Ion_Array(6).SeparationFactor))
              txtSeparationFactorsRow7(Column).Text = Format$(Ion_Array(7).SeparationFactor, GetTheFormat(Ion_Array(7).SeparationFactor))
              txtSeparationFactorsRow8(Column).Text = Format$(Ion_Array(8).SeparationFactor, GetTheFormat(Ion_Array(8).SeparationFactor))
              txtSeparationFactorsRow9(Column).Text = Format$(Ion_Array(9).SeparationFactor, GetTheFormat(Ion_Array(9).SeparationFactor))
              txtSeparationFactorsRow10(Column).Text = Format$(Ion_Array(10).SeparationFactor, GetTheFormat(Ion_Array(10).SeparationFactor))

          End Select

End Sub

Private Sub EnableTextBoxesForInput()
    Dim i As Integer, j As Integer
    
    For i = 1 To 20
        Select Case optSeparationFactors(i).Value
           Case True
              Select Case i
                 Case 1 To 10
                    txtSeparationFactorsRow1(i).Enabled = True
                    txtSeparationFactorsRow2(i).Enabled = True
                    txtSeparationFactorsRow3(i).Enabled = True
                    txtSeparationFactorsRow4(i).Enabled = True
                    txtSeparationFactorsRow5(i).Enabled = True
                    txtSeparationFactorsRow6(i).Enabled = True
                    txtSeparationFactorsRow7(i).Enabled = True
                    txtSeparationFactorsRow8(i).Enabled = True
                    txtSeparationFactorsRow9(i).Enabled = True
                    txtSeparationFactorsRow10(i).Enabled = True
                    OneDimSeparationFactors(1) = CDbl(txtSeparationFactorsRow1(i).Text)
                    OneDimSeparationFactors(2) = CDbl(txtSeparationFactorsRow2(i).Text)
                    OneDimSeparationFactors(3) = CDbl(txtSeparationFactorsRow3(i).Text)
                    OneDimSeparationFactors(4) = CDbl(txtSeparationFactorsRow4(i).Text)
                    OneDimSeparationFactors(5) = CDbl(txtSeparationFactorsRow5(i).Text)
                    OneDimSeparationFactors(6) = CDbl(txtSeparationFactorsRow6(i).Text)
                    OneDimSeparationFactors(7) = CDbl(txtSeparationFactorsRow7(i).Text)
                    OneDimSeparationFactors(8) = CDbl(txtSeparationFactorsRow8(i).Text)
                    OneDimSeparationFactors(9) = CDbl(txtSeparationFactorsRow9(i).Text)
                    OneDimSeparationFactors(10) = CDbl(txtSeparationFactorsRow10(i).Text)
                 Case 11
                    For j = 1 To 10
                        txtSeparationFactorsRow1(j).Enabled = True
                        OneDimSeparationFactors(j) = CDbl(txtSeparationFactorsRow1(j).Text)
                    Next j
                 Case 12
                    For j = 1 To 10
                        txtSeparationFactorsRow2(j).Enabled = True
                        OneDimSeparationFactors(j) = CDbl(txtSeparationFactorsRow2(j).Text)
                    Next j
                 Case 13
                    For j = 1 To 10
                        txtSeparationFactorsRow3(j).Enabled = True
                        OneDimSeparationFactors(j) = CDbl(txtSeparationFactorsRow3(j).Text)
                    Next j
                 Case 14
                    For j = 1 To 10
                        txtSeparationFactorsRow4(j).Enabled = True
                        OneDimSeparationFactors(j) = CDbl(txtSeparationFactorsRow4(j).Text)
                    Next j
                 Case 15
                    For j = 1 To 10
                        txtSeparationFactorsRow5(j).Enabled = True
                        OneDimSeparationFactors(j) = CDbl(txtSeparationFactorsRow5(j).Text)
                    Next j
                 Case 16
                    For j = 1 To 10
                        txtSeparationFactorsRow6(j).Enabled = True
                        OneDimSeparationFactors(j) = CDbl(txtSeparationFactorsRow6(j).Text)
                    Next j
                 Case 17
                    For j = 1 To 10
                        txtSeparationFactorsRow7(j).Enabled = True
                        OneDimSeparationFactors(j) = CDbl(txtSeparationFactorsRow7(j).Text)
                    Next j
                 Case 18
                    For j = 1 To 10
                        txtSeparationFactorsRow8(j).Enabled = True
                        OneDimSeparationFactors(j) = CDbl(txtSeparationFactorsRow8(j).Text)
                    Next j
                 Case 19
                    For j = 1 To 10
                        txtSeparationFactorsRow9(j).Enabled = True
                        OneDimSeparationFactors(j) = CDbl(txtSeparationFactorsRow9(j).Text)
                    Next j
                 Case 20
                    For j = 1 To 10
                        txtSeparationFactorsRow10(j).Enabled = True
                        OneDimSeparationFactors(j) = CDbl(txtSeparationFactorsRow10(j).Text)
                    Next j
              End Select
        End Select
    Next i

    'Disable text boxes for ions onto themselves (i.e. 1,1; 2,2; 3,3 etc.)
    For i = 1 To 10
        txtSeparationFactorsRow1(1).Enabled = False
        txtSeparationFactorsRow2(2).Enabled = False
        txtSeparationFactorsRow3(3).Enabled = False
        txtSeparationFactorsRow4(4).Enabled = False
        txtSeparationFactorsRow5(5).Enabled = False
        txtSeparationFactorsRow6(6).Enabled = False
        txtSeparationFactorsRow7(7).Enabled = False
        txtSeparationFactorsRow8(8).Enabled = False
        txtSeparationFactorsRow9(9).Enabled = False
        txtSeparationFactorsRow10(10).Enabled = False
    Next i

End Sub

Private Sub Form_Activate()
    Dim i As Integer, j As Integer
    Dim i_left As Integer, j_top As Integer
    Dim OptButtonLeft As Integer, OptButtonTop As Integer
    Dim RightOfPanel As Integer, RightOfOptionButton As Integer
    Dim WidthcmdOKandcmdCancel As Integer
    Dim PositionLeft As Integer
    Dim WidthOfPanelIonsAndDescription As Integer

    Call DisableAllTextBoxes

    For i = 1 To 10
        txtSeparationFactorsRow1(i).visible = False
        txtSeparationFactorsRow2(i).visible = False
        txtSeparationFactorsRow3(i).visible = False
        txtSeparationFactorsRow4(i).visible = False
        txtSeparationFactorsRow5(i).visible = False
        txtSeparationFactorsRow6(i).visible = False
        txtSeparationFactorsRow7(i).visible = False
        txtSeparationFactorsRow8(i).visible = False
        txtSeparationFactorsRow9(i).visible = False
        txtSeparationFactorsRow10(i).visible = False
        lbli(i).visible = False
        lblj(i).visible = False
        optSeparationFactors(i).visible = False
        optSeparationFactors(i + 10).visible = False
        lblIonName(i).visible = False
        lblIonNumber(i).visible = False
    Next i

    i_left = (lbli(NumberOfIons).left + lbli(NumberOfIons).width - lbli(1).left) / 2 - lblheaderi.width / 2
    i_left = i_left + lbli(1).left
    lblheaderi.left = i_left

    j_top = (lblj(NumberOfIons).top + lblj(NumberOfIons).height - lblj(1).top) / 2 - lblheaderj.height / 2
    j_top = j_top + lblj(1).top
    lblheaderj.top = j_top

    OptButtonLeft = txtSeparationFactorsRow1(NumberOfIons).left + txtSeparationFactorsRow1(NumberOfIons).width + 120
    panelIons.left = txtSeparationFactorsRow1(1).left

    Select Case NumberOfIons
       Case 1
          OptButtonTop = txtSeparationFactorsRow1(1).top + txtSeparationFactorsRow1(1).height + 60
       Case 2
          OptButtonTop = txtSeparationFactorsRow2(1).top + txtSeparationFactorsRow2(1).height + 60
       Case 3
          OptButtonTop = txtSeparationFactorsRow3(1).top + txtSeparationFactorsRow3(1).height + 60
       Case 4
          OptButtonTop = txtSeparationFactorsRow4(1).top + txtSeparationFactorsRow4(1).height + 60
       Case 5
          OptButtonTop = txtSeparationFactorsRow5(1).top + txtSeparationFactorsRow5(1).height + 60
       Case 6
          OptButtonTop = txtSeparationFactorsRow6(1).top + txtSeparationFactorsRow6(1).height + 60
       Case 7
          OptButtonTop = txtSeparationFactorsRow7(1).top + txtSeparationFactorsRow7(1).height + 60
       Case 8
          OptButtonTop = txtSeparationFactorsRow8(1).top + txtSeparationFactorsRow8(1).height + 60
       Case 9
          OptButtonTop = txtSeparationFactorsRow9(1).top + txtSeparationFactorsRow9(1).height + 60
       Case 10
          OptButtonTop = txtSeparationFactorsRow10(1).top + txtSeparationFactorsRow10(1).height + 60
    End Select

    panelIons.top = OptButtonTop + optSeparationFactors(1).height + 60

    For i = 1 To NumberOfIons
        lbli(i).visible = True
        lblj(i).visible = True
        optSeparationFactors(i).top = OptButtonTop
        optSeparationFactors(i + 10).left = OptButtonLeft
        optSeparationFactors(i).visible = True
        optSeparationFactors(i + 10).visible = True
        lblIonNumber(i).visible = True
        lblIonName(i).visible = True
        lblIonName(i).Caption = UCase$(Trim$(NowProj.cation(i).Name))

        Select Case NumberOfIons
           Case 1
              txtSeparationFactorsRow1(i).visible = True

           Case 2
              txtSeparationFactorsRow1(i).visible = True
              txtSeparationFactorsRow2(i).visible = True

           Case 3
              txtSeparationFactorsRow1(i).visible = True
              txtSeparationFactorsRow2(i).visible = True
              txtSeparationFactorsRow3(i).visible = True

           Case 4
              txtSeparationFactorsRow1(i).visible = True
              txtSeparationFactorsRow2(i).visible = True
              txtSeparationFactorsRow3(i).visible = True
              txtSeparationFactorsRow4(i).visible = True

           Case 5
              txtSeparationFactorsRow1(i).visible = True
              txtSeparationFactorsRow2(i).visible = True
              txtSeparationFactorsRow3(i).visible = True
              txtSeparationFactorsRow4(i).visible = True
              txtSeparationFactorsRow5(i).visible = True

           Case 6
              txtSeparationFactorsRow1(i).visible = True
              txtSeparationFactorsRow2(i).visible = True
              txtSeparationFactorsRow3(i).visible = True
              txtSeparationFactorsRow4(i).visible = True
              txtSeparationFactorsRow5(i).visible = True
              txtSeparationFactorsRow6(i).visible = True

           Case 7
              txtSeparationFactorsRow1(i).visible = True
              txtSeparationFactorsRow2(i).visible = True
              txtSeparationFactorsRow3(i).visible = True
              txtSeparationFactorsRow4(i).visible = True
              txtSeparationFactorsRow5(i).visible = True
              txtSeparationFactorsRow6(i).visible = True
              txtSeparationFactorsRow7(i).visible = True
           
           Case 8
              txtSeparationFactorsRow1(i).visible = True
              txtSeparationFactorsRow2(i).visible = True
              txtSeparationFactorsRow3(i).visible = True
              txtSeparationFactorsRow4(i).visible = True
              txtSeparationFactorsRow5(i).visible = True
              txtSeparationFactorsRow6(i).visible = True
              txtSeparationFactorsRow7(i).visible = True
              txtSeparationFactorsRow8(i).visible = True

           Case 9
              txtSeparationFactorsRow1(i).visible = True
              txtSeparationFactorsRow2(i).visible = True
              txtSeparationFactorsRow3(i).visible = True
              txtSeparationFactorsRow4(i).visible = True
              txtSeparationFactorsRow5(i).visible = True
              txtSeparationFactorsRow6(i).visible = True
              txtSeparationFactorsRow7(i).visible = True
              txtSeparationFactorsRow8(i).visible = True
              txtSeparationFactorsRow9(i).visible = True

           Case 10
              txtSeparationFactorsRow1(i).visible = True
              txtSeparationFactorsRow2(i).visible = True
              txtSeparationFactorsRow3(i).visible = True
              txtSeparationFactorsRow4(i).visible = True
              txtSeparationFactorsRow5(i).visible = True
              txtSeparationFactorsRow6(i).visible = True
              txtSeparationFactorsRow7(i).visible = True
              txtSeparationFactorsRow8(i).visible = True
              txtSeparationFactorsRow9(i).visible = True
              txtSeparationFactorsRow10(i).visible = True

        End Select
    Next i

    panelIons.height = lblIonName(NumberOfIons).top + lblIonName(NumberOfIons).height + 60

    If NumberOfIons <= 5 Then
       panelDescription.left = panelIons.left
       panelDescription.top = panelIons.top + panelIons.height + 120
       RightOfPanel = panelDescription.left + panelDescription.width
       RightOfOptionButton = optSeparationFactors(11).left + optSeparationFactors(11).width
       If RightOfPanel > RightOfOptionButton Then
          frmSeparationFactors.width = RightOfPanel + 240
       Else
          frmSeparationFactors.width = RightOfOptionButton + 220
       End If

       panelIons.left = frmSeparationFactors.width / 2 - panelIons.width / 2
       panelDescription.left = panelIons.left
       WidthcmdOKandcmdCancel = cmdOK.width + 120 + cmdCancel.width
       cmdOK.left = frmSeparationFactors.width / 2 - WidthcmdOKandcmdCancel / 2
       cmdCancel.left = cmdOK.left + cmdOK.width + 120
       cmdOK.top = panelDescription.top + panelDescription.height + 300
       cmdCancel.top = cmdOK.top
       frmSeparationFactors.height = cmdOK.top + cmdOK.height + 360

       frmSeparationFactors.WindowState = 0

       'Position the form on the screen (Centered in Left Half of It)
       If WindowState = 0 Then
          'don't attempt if screen Minimized or Maximized
          If NumberOfIons = 5 Then
             PositionLeft = 240
          Else
             PositionLeft = ((Screen.width / 2 - frmIonExchangeMain.left) / 2) - frmSeparationFactors.width / 2
          End If
          Move (frmIonExchangeMain.left + PositionLeft), (Screen.height - frmSeparationFactors.height) / 2
       End If

    Else
       panelDescription.top = panelIons.top
       panelDescription.left = panelIons.left + panelIons.width + 180
       RightOfPanel = panelDescription.left + panelDescription.width
       RightOfOptionButton = optSeparationFactors(11).left + optSeparationFactors(11).width
       If RightOfPanel > RightOfOptionButton Then
          frmSeparationFactors.width = RightOfPanel + 240
       Else
          frmSeparationFactors.width = RightOfOptionButton + 220
       End If
       WidthOfPanelIonsAndDescription = panelIons.width + 180 + panelDescription.width
       panelIons.left = frmSeparationFactors.width / 2 - WidthOfPanelIonsAndDescription / 2
       panelDescription.left = panelIons.left + panelIons.width + 180
       WidthcmdOKandcmdCancel = cmdOK.width + 120 + cmdCancel.width
       cmdOK.left = frmSeparationFactors.width - WidthcmdOKandcmdCancel - 240
       cmdCancel.left = cmdOK.left + cmdOK.width + 120
       
       cmdOK.top = panelDescription.top + panelDescription.height + 180
       cmdCancel.top = cmdOK.top
       frmSeparationFactors.height = cmdOK.top + cmdOK.height + 360

       frmSeparationFactors.WindowState = 0

       'Position the form on the screen (Centered)
       If WindowState = 0 Then
          'don't attempt if screen Minimized or Maximized
          Move (Screen.width - frmSeparationFactors.width) / 2, (Screen.height - frmSeparationFactors.height) / 2
       End If

    End If

    Call CalculateSeparationFactors
    Call ShowSeparationFactors
    Call EnableTextBoxesForInput

End Sub

Private Sub Form_Load()

    'Center the form on the screen
    If WindowState = 0 Then
       'don't attempt if screen Minimized or Maximized
       Move (Screen.width - frmSeparationFactors.width) / 2, (Screen.height - frmSeparationFactors.height) / 2
    End If

End Sub

Private Sub optSeparationFactors_Click(index As Integer, Value As Integer)

    Select Case index
       Case 1 To 10
            SeparationFactorInput.Row = False
       Case 11 To 20
            SeparationFactorInput.Row = True
    End Select

    SeparationFactorInput.Value = index

    Call DisableAllTextBoxes
    Call EnableTextBoxesForInput
    Call CalculateSeparationFactors
    Call ShowSeparationFactors

End Sub

Private Sub ShowSeparationFactors()
    Dim i As Integer, j As Integer

       For i = 1 To NumberOfIons
           txtSeparationFactorsRow1(i).Text = Format$(TwoDimSeparationFactors(i, 1), GetTheFormat(TwoDimSeparationFactors(i, 1)))
           txtSeparationFactorsRow2(i).Text = Format$(TwoDimSeparationFactors(i, 2), GetTheFormat(TwoDimSeparationFactors(i, 2)))
           txtSeparationFactorsRow3(i).Text = Format$(TwoDimSeparationFactors(i, 3), GetTheFormat(TwoDimSeparationFactors(i, 3)))
           txtSeparationFactorsRow4(i).Text = Format$(TwoDimSeparationFactors(i, 4), GetTheFormat(TwoDimSeparationFactors(i, 4)))
           txtSeparationFactorsRow5(i).Text = Format$(TwoDimSeparationFactors(i, 5), GetTheFormat(TwoDimSeparationFactors(i, 5)))
           txtSeparationFactorsRow6(i).Text = Format$(TwoDimSeparationFactors(i, 6), GetTheFormat(TwoDimSeparationFactors(i, 6)))
           txtSeparationFactorsRow7(i).Text = Format$(TwoDimSeparationFactors(i, 7), GetTheFormat(TwoDimSeparationFactors(i, 7)))
           txtSeparationFactorsRow8(i).Text = Format$(TwoDimSeparationFactors(i, 8), GetTheFormat(TwoDimSeparationFactors(i, 8)))
           txtSeparationFactorsRow9(i).Text = Format$(TwoDimSeparationFactors(i, 9), GetTheFormat(TwoDimSeparationFactors(i, 9)))
           txtSeparationFactorsRow10(i).Text = Format$(TwoDimSeparationFactors(i, 10), GetTheFormat(TwoDimSeparationFactors(i, 10)))
       Next i
   

End Sub

Private Sub txtSeparationFactorsRow1_GotFocus(index As Integer)
    Call TextGetFocus(txtSeparationFactorsRow1(index), Temp_Text)
End Sub

Private Sub txtSeparationFactorsRow1_KeyPress(index As Integer, KeyAscii As Integer)

    If KeyAscii = 13 Then
       KeyAscii = 0
       SendKeys "{Tab}"
       Exit Sub
    End If

    Call NumberCheck(KeyAscii)

End Sub

Private Sub txtSeparationFactorsRow1_LostFocus(index As Integer)
    Dim ValueChanged As Integer, NewValue As Double, IsError As Integer
    Dim msg As String, CurrentUnits As Integer
    Dim OldValue As Double, ValueToDisplay As Double

    Call TextHandleError(IsError, txtSeparationFactorsRow1(index), Temp_Text)

    If Not IsError Then
       OldValue = CDbl(Temp_Text)
       NewValue = CDbl(txtSeparationFactorsRow1(index).Text)
       Call NumberChanged(ValueChanged, OldValue, NewValue)
       If ValueChanged Then
          If HaveValue(NewValue) Then
             If SeparationFactorInput.Row = True Then
                OneDimSeparationFactors(index) = NewValue
             Else
                OneDimSeparationFactors(1) = NewValue
             End If
             
             Call CalculateSeparationFactors
             Call ShowSeparationFactors
          Else
             txtSeparationFactorsRow1(index).Text = Temp_Text
             txtSeparationFactorsRow1(index).SetFocus
             Exit Sub
          End If
       End If
    End If

End Sub

Private Sub txtSeparationFactorsRow10_GotFocus(index As Integer)
    Call TextGetFocus(txtSeparationFactorsRow10(index), Temp_Text)
End Sub

Private Sub txtSeparationFactorsRow10_KeyPress(index As Integer, KeyAscii As Integer)

    If KeyAscii = 13 Then
       KeyAscii = 0
       SendKeys "{Tab}"
       Exit Sub
    End If

    Call NumberCheck(KeyAscii)

End Sub

Private Sub txtSeparationFactorsRow10_LostFocus(index As Integer)
    Dim ValueChanged As Integer, NewValue As Double, IsError As Integer
    Dim msg As String, CurrentUnits As Integer
    Dim OldValue As Double, ValueToDisplay As Double

    Call TextHandleError(IsError, txtSeparationFactorsRow10(index), Temp_Text)

    If Not IsError Then
       OldValue = CDbl(Temp_Text)
       NewValue = CDbl(txtSeparationFactorsRow10(index).Text)
       Call NumberChanged(ValueChanged, OldValue, NewValue)
       If ValueChanged Then
          If HaveValue(NewValue) Then
             If SeparationFactorInput.Row = True Then
                OneDimSeparationFactors(index) = NewValue
             Else
                OneDimSeparationFactors(10) = NewValue
             End If
           
             Call CalculateSeparationFactors
             Call ShowSeparationFactors
          Else
             txtSeparationFactorsRow10(index).Text = Temp_Text
             txtSeparationFactorsRow10(index).SetFocus
             Exit Sub
          End If
       End If
    End If

End Sub

Private Sub txtSeparationFactorsRow100_LostFocus(index As Integer)
    Dim ValueChanged As Integer, NewValue As Double, IsError As Integer
    Dim msg As String, CurrentUnits As Integer
    Dim OldValue As Double, ValueToDisplay As Double

    Call TextHandleError(IsError, txtSeparationFactorsRow10(index), Temp_Text)

    If Not IsError Then
       OldValue = CDbl(Temp_Text)
       NewValue = CDbl(txtSeparationFactorsRow10(index).Text)
       Call NumberChanged(ValueChanged, OldValue, NewValue)
       If ValueChanged Then
          If HaveValue(NewValue) Then
             OneDimSeparationFactors(index) = NewValue
             Call CalculateSeparationFactors
             Call ShowSeparationFactors
          Else
             txtSeparationFactorsRow10(index).Text = Temp_Text
             txtSeparationFactorsRow10(index).SetFocus
             Exit Sub
          End If
       End If
    End If

End Sub

Private Sub txtSeparationFactorsRow2_GotFocus(index As Integer)
    Call TextGetFocus(txtSeparationFactorsRow2(index), Temp_Text)
End Sub

Private Sub txtSeparationFactorsRow2_KeyPress(index As Integer, KeyAscii As Integer)

    If KeyAscii = 13 Then
       KeyAscii = 0
       SendKeys "{Tab}"
       Exit Sub
    End If

    Call NumberCheck(KeyAscii)

End Sub

Private Sub txtSeparationFactorsRow2_LostFocus(index As Integer)
    Dim ValueChanged As Integer, NewValue As Double, IsError As Integer
    Dim msg As String, CurrentUnits As Integer
    Dim OldValue As Double, ValueToDisplay As Double

    Call TextHandleError(IsError, txtSeparationFactorsRow2(index), Temp_Text)

    If Not IsError Then
       OldValue = CDbl(Temp_Text)
       NewValue = CDbl(txtSeparationFactorsRow2(index).Text)
       Call NumberChanged(ValueChanged, OldValue, NewValue)
       If ValueChanged Then
          If HaveValue(NewValue) Then
             If SeparationFactorInput.Row = True Then
                OneDimSeparationFactors(index) = NewValue
             Else
                OneDimSeparationFactors(2) = NewValue
             End If
             Call CalculateSeparationFactors
             Call ShowSeparationFactors
          Else
             txtSeparationFactorsRow2(index).Text = Temp_Text
             txtSeparationFactorsRow2(index).SetFocus
             Exit Sub
          End If
       End If
    End If

End Sub

Private Sub txtSeparationFactorsRow3_GotFocus(index As Integer)
    Call TextGetFocus(txtSeparationFactorsRow3(index), Temp_Text)
End Sub

Private Sub txtSeparationFactorsRow3_KeyPress(index As Integer, KeyAscii As Integer)

    If KeyAscii = 13 Then
       KeyAscii = 0
       SendKeys "{Tab}"
       Exit Sub
    End If

    Call NumberCheck(KeyAscii)

End Sub

Private Sub txtSeparationFactorsRow3_LostFocus(index As Integer)
    Dim ValueChanged As Integer, NewValue As Double, IsError As Integer
    Dim msg As String, CurrentUnits As Integer
    Dim OldValue As Double, ValueToDisplay As Double

    Call TextHandleError(IsError, txtSeparationFactorsRow3(index), Temp_Text)

    If Not IsError Then
       OldValue = CDbl(Temp_Text)
       NewValue = CDbl(txtSeparationFactorsRow3(index).Text)
       Call NumberChanged(ValueChanged, OldValue, NewValue)
       If ValueChanged Then
          If HaveValue(NewValue) Then
             If SeparationFactorInput.Row = True Then
                OneDimSeparationFactors(index) = NewValue
             Else
                OneDimSeparationFactors(3) = NewValue
             End If
             
             Call CalculateSeparationFactors
             Call ShowSeparationFactors
          Else
             txtSeparationFactorsRow3(index).Text = Temp_Text
             txtSeparationFactorsRow3(index).SetFocus
             Exit Sub
          End If
       End If
    End If

End Sub

Private Sub txtSeparationFactorsRow4_GotFocus(index As Integer)
    Call TextGetFocus(txtSeparationFactorsRow4(index), Temp_Text)
End Sub

Private Sub txtSeparationFactorsRow4_KeyPress(index As Integer, KeyAscii As Integer)

    If KeyAscii = 13 Then
       KeyAscii = 0
       SendKeys "{Tab}"
       Exit Sub
    End If

    Call NumberCheck(KeyAscii)

End Sub

Private Sub txtSeparationFactorsRow4_LostFocus(index As Integer)
    Dim ValueChanged As Integer, NewValue As Double, IsError As Integer
    Dim msg As String, CurrentUnits As Integer
    Dim OldValue As Double, ValueToDisplay As Double

    Call TextHandleError(IsError, txtSeparationFactorsRow4(index), Temp_Text)

    If Not IsError Then
       OldValue = CDbl(Temp_Text)
       NewValue = CDbl(txtSeparationFactorsRow4(index).Text)
       Call NumberChanged(ValueChanged, OldValue, NewValue)
       If ValueChanged Then
          If HaveValue(NewValue) Then
             If SeparationFactorInput.Row = True Then
                OneDimSeparationFactors(index) = NewValue
             Else
                OneDimSeparationFactors(4) = NewValue
             End If
           
             Call CalculateSeparationFactors
             Call ShowSeparationFactors
          Else
             txtSeparationFactorsRow4(index).Text = Temp_Text
             txtSeparationFactorsRow4(index).SetFocus
             Exit Sub
          End If
       End If
    End If

End Sub

Private Sub txtSeparationFactorsRow5_GotFocus(index As Integer)
    Call TextGetFocus(txtSeparationFactorsRow5(index), Temp_Text)
End Sub

Private Sub txtSeparationFactorsRow5_KeyPress(index As Integer, KeyAscii As Integer)

    If KeyAscii = 13 Then
       KeyAscii = 0
       SendKeys "{Tab}"
       Exit Sub
    End If

    Call NumberCheck(KeyAscii)

End Sub

Private Sub txtSeparationFactorsRow5_LostFocus(index As Integer)
    Dim ValueChanged As Integer, NewValue As Double, IsError As Integer
    Dim msg As String, CurrentUnits As Integer
    Dim OldValue As Double, ValueToDisplay As Double

    Call TextHandleError(IsError, txtSeparationFactorsRow5(index), Temp_Text)

    If Not IsError Then
       OldValue = CDbl(Temp_Text)
       NewValue = CDbl(txtSeparationFactorsRow5(index).Text)
       Call NumberChanged(ValueChanged, OldValue, NewValue)
       If ValueChanged Then
          If HaveValue(NewValue) Then
             If SeparationFactorInput.Row = True Then
                OneDimSeparationFactors(index) = NewValue
             Else
                OneDimSeparationFactors(5) = NewValue
             End If
             
             Call CalculateSeparationFactors
             Call ShowSeparationFactors
          Else
             txtSeparationFactorsRow5(index).Text = Temp_Text
             txtSeparationFactorsRow5(index).SetFocus
             Exit Sub
          End If
       End If
    End If

End Sub

Private Sub txtSeparationFactorsRow6_GotFocus(index As Integer)
    Call TextGetFocus(txtSeparationFactorsRow6(index), Temp_Text)
End Sub

Private Sub txtSeparationFactorsRow6_KeyPress(index As Integer, KeyAscii As Integer)

    If KeyAscii = 13 Then
       KeyAscii = 0
       SendKeys "{Tab}"
       Exit Sub
    End If

    Call NumberCheck(KeyAscii)

End Sub

Private Sub txtSeparationFactorsRow6_LostFocus(index As Integer)
    Dim ValueChanged As Integer, NewValue As Double, IsError As Integer
    Dim msg As String, CurrentUnits As Integer
    Dim OldValue As Double, ValueToDisplay As Double

    Call TextHandleError(IsError, txtSeparationFactorsRow6(index), Temp_Text)

    If Not IsError Then
       OldValue = CDbl(Temp_Text)
       NewValue = CDbl(txtSeparationFactorsRow6(index).Text)
       Call NumberChanged(ValueChanged, OldValue, NewValue)
       If ValueChanged Then
          If HaveValue(NewValue) Then
             If SeparationFactorInput.Row = True Then
                OneDimSeparationFactors(index) = NewValue
             Else
                OneDimSeparationFactors(6) = NewValue
             End If
          
             Call CalculateSeparationFactors
             Call ShowSeparationFactors
          Else
             txtSeparationFactorsRow6(index).Text = Temp_Text
             txtSeparationFactorsRow6(index).SetFocus
             Exit Sub
          End If
       End If
    End If

End Sub

Private Sub txtSeparationFactorsRow7_GotFocus(index As Integer)
    Call TextGetFocus(txtSeparationFactorsRow7(index), Temp_Text)
End Sub

Private Sub txtSeparationFactorsRow7_KeyPress(index As Integer, KeyAscii As Integer)

    If KeyAscii = 13 Then
       KeyAscii = 0
       SendKeys "{Tab}"
       Exit Sub
    End If

    Call NumberCheck(KeyAscii)

End Sub

Private Sub txtSeparationFactorsRow7_LostFocus(index As Integer)
    Dim ValueChanged As Integer, NewValue As Double, IsError As Integer
    Dim msg As String, CurrentUnits As Integer
    Dim OldValue As Double, ValueToDisplay As Double

    Call TextHandleError(IsError, txtSeparationFactorsRow7(index), Temp_Text)

    If Not IsError Then
       OldValue = CDbl(Temp_Text)
       NewValue = CDbl(txtSeparationFactorsRow7(index).Text)
       Call NumberChanged(ValueChanged, OldValue, NewValue)
       If ValueChanged Then
          If HaveValue(NewValue) Then
             If SeparationFactorInput.Row = True Then
                OneDimSeparationFactors(index) = NewValue
             Else
                OneDimSeparationFactors(7) = NewValue
             End If
          
             Call CalculateSeparationFactors
             Call ShowSeparationFactors
          Else
             txtSeparationFactorsRow7(index).Text = Temp_Text
             txtSeparationFactorsRow7(index).SetFocus
             Exit Sub
          End If
       End If
    End If

End Sub

Private Sub txtSeparationFactorsRow8_GotFocus(index As Integer)
    Call TextGetFocus(txtSeparationFactorsRow8(index), Temp_Text)
End Sub

Private Sub txtSeparationFactorsRow8_KeyPress(index As Integer, KeyAscii As Integer)

    If KeyAscii = 13 Then
       KeyAscii = 0
       SendKeys "{Tab}"
       Exit Sub
    End If

    Call NumberCheck(KeyAscii)

End Sub

Private Sub txtSeparationFactorsRow8_LostFocus(index As Integer)
    Dim ValueChanged As Integer, NewValue As Double, IsError As Integer
    Dim msg As String, CurrentUnits As Integer
    Dim OldValue As Double, ValueToDisplay As Double

    Call TextHandleError(IsError, txtSeparationFactorsRow8(index), Temp_Text)

    If Not IsError Then
       OldValue = CDbl(Temp_Text)
       NewValue = CDbl(txtSeparationFactorsRow8(index).Text)
       Call NumberChanged(ValueChanged, OldValue, NewValue)
       If ValueChanged Then
          If HaveValue(NewValue) Then
             If SeparationFactorInput.Row = True Then
                OneDimSeparationFactors(index) = NewValue
             Else
                OneDimSeparationFactors(8) = NewValue
             End If
           
             Call CalculateSeparationFactors
             Call ShowSeparationFactors
          Else
             txtSeparationFactorsRow8(index).Text = Temp_Text
             txtSeparationFactorsRow8(index).SetFocus
             Exit Sub
          End If
       End If
    End If

End Sub

Private Sub txtSeparationFactorsRow9_GotFocus(index As Integer)
    Call TextGetFocus(txtSeparationFactorsRow9(index), Temp_Text)
End Sub

Private Sub txtSeparationFactorsRow9_KeyPress(index As Integer, KeyAscii As Integer)

    If KeyAscii = 13 Then
       KeyAscii = 0
       SendKeys "{Tab}"
       Exit Sub
    End If

    Call NumberCheck(KeyAscii)

End Sub

Private Sub txtSeparationFactorsRow9_LostFocus(index As Integer)
    Dim ValueChanged As Integer, NewValue As Double, IsError As Integer
    Dim msg As String, CurrentUnits As Integer
    Dim OldValue As Double, ValueToDisplay As Double

    Call TextHandleError(IsError, txtSeparationFactorsRow9(index), Temp_Text)

    If Not IsError Then
       OldValue = CDbl(Temp_Text)
       NewValue = CDbl(txtSeparationFactorsRow9(index).Text)
       Call NumberChanged(ValueChanged, OldValue, NewValue)
       If ValueChanged Then
          If HaveValue(NewValue) Then
             If SeparationFactorInput.Row = True Then
                OneDimSeparationFactors(index) = NewValue
             Else
                OneDimSeparationFactors(9) = NewValue
             End If
           
             Call CalculateSeparationFactors
             Call ShowSeparationFactors
          Else
             txtSeparationFactorsRow9(index).Text = Temp_Text
             txtSeparationFactorsRow9(index).SetFocus
             Exit Sub
          End If
       End If
    End If

End Sub

