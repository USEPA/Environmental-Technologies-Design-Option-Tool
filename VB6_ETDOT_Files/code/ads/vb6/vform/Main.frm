VERSION 5.00
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "Threed32.ocx"
Object = "{B16553C3-06DB-101B-85B2-0000C009BE81}#1.0#0"; "Spin32.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmMain 
   Appearance      =   0  'Flat
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "{Caption set in Form_Load}"
   ClientHeight    =   6915
   ClientLeft      =   1740
   ClientTop       =   2160
   ClientWidth     =   9540
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
   Icon            =   "Main.frx":0000
   MaxButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   461
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   636
   Begin VB.PictureBox Picture1 
      Height          =   855
      Left            =   9720
      ScaleHeight     =   795
      ScaleWidth      =   1275
      TabIndex        =   97
      Top             =   5160
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.CommandButton Command4 
      Caption         =   "Print Screen"
      Height          =   330
      Left            =   1560
      TabIndex        =   98
      ToolTipText     =   "Click here to print current screen to selected printer"
      Top             =   6105
      Width           =   1455
   End
   Begin Threed.SSFrame SSFrame8 
      Height          =   2985
      Left            =   9330
      TabIndex        =   96
      Top             =   2160
      Visible         =   0   'False
      Width           =   4485
      _Version        =   65536
      _ExtentX        =   7911
      _ExtentY        =   5265
      _StockProps     =   14
      Caption         =   "Invisible - Do not delete!"
      ForeColor       =   255
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin Threed.SSFrame ssframe_Adsorbent 
      Height          =   2895
      Left            =   4710
      TabIndex        =   43
      Top             =   3540
      Width           =   4605
      _Version        =   65536
      _ExtentX        =   8123
      _ExtentY        =   5106
      _StockProps     =   14
      Caption         =   "Adsorbent Properties:"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Begin VB.CommandButton cmdViewDimensionless 
         Appearance      =   0  'Flat
         Caption         =   "D&imensionless Groups"
         Height          =   315
         Left            =   120
         TabIndex        =   84
         TabStop         =   0   'False
         Top             =   2310
         Width           =   2167
      End
      Begin VB.TextBox txtCarbon 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         Height          =   285
         Index           =   0
         Left            =   2040
         TabIndex        =   13
         Text            =   "Test"
         Top             =   630
         Width           =   2415
      End
      Begin VB.ComboBox txtCarbonUnits 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         Height          =   315
         Index           =   2
         Left            =   3300
         Style           =   2  'Dropdown List
         TabIndex        =   83
         TabStop         =   0   'False
         Top             =   1260
         Width           =   1200
      End
      Begin VB.ComboBox txtCarbonUnits 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         Height          =   315
         Index           =   1
         Left            =   3300
         Style           =   2  'Dropdown List
         TabIndex        =   82
         TabStop         =   0   'False
         Top             =   930
         Width           =   1200
      End
      Begin VB.TextBox txtCarbon 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         Height          =   285
         Index           =   3
         Left            =   2040
         TabIndex        =   16
         Text            =   "Test"
         Top             =   1620
         Width           =   1215
      End
      Begin VB.TextBox txtCarbon 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         Height          =   285
         Index           =   2
         Left            =   2040
         TabIndex        =   15
         Text            =   "Test"
         Top             =   1290
         Width           =   1215
      End
      Begin VB.TextBox txtCarbon 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         Height          =   285
         Index           =   1
         Left            =   2040
         TabIndex        =   14
         Text            =   "Test"
         Top             =   960
         Width           =   1215
      End
      Begin VB.CommandButton cmdCarbon 
         Appearance      =   0  'Flat
         Caption         =   "Adsorbe&nt Database"
         Height          =   315
         Left            =   120
         TabIndex        =   81
         TabStop         =   0   'False
         Top             =   240
         Width           =   4335
      End
      Begin VB.CommandButton cmdpolanyi 
         Appearance      =   0  'Flat
         Caption         =   "Polan&yi Parameters"
         Height          =   315
         Left            =   2340
         TabIndex        =   80
         TabStop         =   0   'False
         Top             =   2310
         Width           =   2115
      End
      Begin VB.TextBox txtCarbon 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         Height          =   285
         Index           =   4
         Left            =   2040
         TabIndex        =   17
         Text            =   "Test"
         Top             =   1950
         Width           =   1215
      End
      Begin VB.Label lblMiscUnits 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         Caption         =   "-"
         ForeColor       =   &H80000008&
         Height          =   315
         Index           =   0
         Left            =   3300
         TabIndex        =   91
         Top             =   1650
         Width           =   1200
      End
      Begin VB.Label lblCarbon 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         Caption         =   "Porosity"
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   3
         Left            =   240
         TabIndex        =   90
         Top             =   1650
         Width           =   1695
      End
      Begin VB.Label lblCarbon 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         Caption         =   "Particle Radius"
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   2
         Left            =   240
         TabIndex        =   89
         Top             =   1320
         Width           =   1695
      End
      Begin VB.Label lblCarbon 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         Caption         =   "Apparent Density"
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   1
         Left            =   240
         TabIndex        =   88
         Top             =   990
         Width           =   1695
      End
      Begin VB.Label lblCarbon 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         Caption         =   "Name"
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   0
         Left            =   240
         TabIndex        =   87
         Top             =   660
         Width           =   1695
      End
      Begin VB.Label lblCarbon 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         Caption         =   "Particle Shape Factor"
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   4
         Left            =   60
         TabIndex        =   86
         Top             =   1980
         Width           =   1875
      End
      Begin VB.Label lblMiscUnits 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         Caption         =   "-"
         ForeColor       =   &H80000008&
         Height          =   315
         Index           =   1
         Left            =   3300
         TabIndex        =   85
         Top             =   1980
         Width           =   1200
      End
   End
   Begin Threed.SSFrame ssframe_FixedBed 
      Height          =   3525
      Left            =   4710
      TabIndex        =   42
      Top             =   60
      Width           =   4605
      _Version        =   65536
      _ExtentX        =   8123
      _ExtentY        =   6218
      _StockProps     =   14
      Caption         =   "Fixed Bed Properties:"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Begin VB.CommandButton cmdNote 
         Height          =   675
         Index           =   0
         Left            =   120
         Picture         =   "Main.frx":030A
         Style           =   1  'Graphical
         TabIndex        =   93
         Top             =   1440
         Width           =   555
      End
      Begin VB.CommandButton cmdNote 
         Height          =   675
         Index           =   1
         Left            =   240
         Picture         =   "Main.frx":074C
         Style           =   1  'Graphical
         TabIndex        =   94
         Top             =   1560
         Width           =   555
      End
      Begin VB.ComboBox txtBedUnits 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         Height          =   315
         Index           =   4
         Left            =   3300
         Style           =   2  'Dropdown List
         TabIndex        =   62
         TabStop         =   0   'False
         Top             =   1830
         Width           =   1200
      End
      Begin VB.ComboBox txtBedUnits 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         Height          =   315
         Index           =   3
         Left            =   3300
         Style           =   2  'Dropdown List
         TabIndex        =   61
         TabStop         =   0   'False
         Top             =   1530
         Width           =   1200
      End
      Begin VB.ComboBox txtBedUnits 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         Height          =   315
         Index           =   2
         Left            =   3300
         Style           =   2  'Dropdown List
         TabIndex        =   60
         TabStop         =   0   'False
         Top             =   1230
         Width           =   1200
      End
      Begin VB.ComboBox txtBedUnits 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         Height          =   315
         Index           =   1
         Left            =   3300
         Style           =   2  'Dropdown List
         TabIndex        =   59
         TabStop         =   0   'False
         Top             =   930
         Width           =   1200
      End
      Begin VB.ComboBox txtBedUnits 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         Height          =   315
         Index           =   0
         Left            =   3300
         Style           =   2  'Dropdown List
         TabIndex        =   58
         TabStop         =   0   'False
         Top             =   630
         Width           =   1200
      End
      Begin VB.TextBox txtBedValue 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         Height          =   288
         HelpContextID   =   20
         Index           =   0
         Left            =   2040
         TabIndex        =   8
         Text            =   "Test"
         Top             =   660
         Width           =   1215
      End
      Begin VB.TextBox txtBedValue 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         Height          =   288
         HelpContextID   =   20
         Index           =   1
         Left            =   2040
         TabIndex        =   9
         Text            =   "Test"
         Top             =   960
         Width           =   1215
      End
      Begin VB.TextBox txtBedValue 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         Height          =   288
         HelpContextID   =   20
         Index           =   2
         Left            =   2040
         TabIndex        =   10
         Text            =   "Test"
         Top             =   1260
         Width           =   1215
      End
      Begin VB.TextBox txtBedValue 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         Height          =   288
         HelpContextID   =   20
         Index           =   3
         Left            =   2040
         TabIndex        =   11
         Text            =   "Test"
         Top             =   1560
         Width           =   1215
      End
      Begin VB.TextBox txtBedValue 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         Height          =   288
         HelpContextID   =   20
         Index           =   4
         Left            =   2040
         TabIndex        =   12
         Text            =   "Test"
         Top             =   1860
         Width           =   1215
      End
      Begin VB.CommandButton cmdAdsorberDB 
         Appearance      =   0  'Flat
         Caption         =   "Adsor&ber Database"
         Height          =   315
         Left            =   120
         TabIndex        =   57
         TabStop         =   0   'False
         Top             =   270
         Width           =   4335
      End
      Begin VB.Label lblBed 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         Caption         =   "Bed Length"
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   0
         Left            =   240
         TabIndex        =   79
         Top             =   690
         Width           =   1695
      End
      Begin VB.Label lblBed 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         Caption         =   "Bed Diameter"
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   1
         Left            =   240
         TabIndex        =   78
         Top             =   990
         Width           =   1695
      End
      Begin VB.Label lblBed 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         Caption         =   "Bed Mass"
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   2
         Left            =   360
         TabIndex        =   77
         Top             =   1290
         Width           =   1575
      End
      Begin VB.Label lblBed 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         Caption         =   "Flowrate"
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   3
         Left            =   240
         TabIndex        =   76
         Top             =   1590
         Width           =   1695
      End
      Begin VB.Label lblBed 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         Caption         =   "EBCT"
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   4
         Left            =   240
         TabIndex        =   75
         Top             =   1890
         Width           =   1710
      End
      Begin VB.Label lblBedDensityDisplay 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Test"
         ForeColor       =   &H00808080&
         Height          =   285
         Left            =   2040
         TabIndex        =   74
         Top             =   2175
         Width           =   1215
      End
      Begin VB.Label lblBed 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         Caption         =   "Bed Density"
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   5
         Left            =   240
         TabIndex        =   73
         Top             =   2190
         Width           =   1710
      End
      Begin VB.Label lblMiscUnits 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         Caption         =   "(g/mL)"
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   2
         Left            =   3300
         TabIndex        =   72
         Top             =   2190
         Width           =   1200
      End
      Begin VB.Label lblporosity 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Test"
         ForeColor       =   &H00808080&
         Height          =   285
         Left            =   2040
         TabIndex        =   71
         Top             =   2490
         Width           =   1215
      End
      Begin VB.Label lblInterstitialVelocity 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Test"
         ForeColor       =   &H00808080&
         Height          =   285
         Left            =   2040
         TabIndex        =   70
         Top             =   3090
         Width           =   1215
      End
      Begin VB.Label lblSuperficialVelocity 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Test"
         ForeColor       =   &H00808080&
         Height          =   285
         Left            =   2040
         TabIndex        =   69
         Top             =   2790
         Width           =   1215
      End
      Begin VB.Label lblMiscUnits 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         Caption         =   "(m/hr)"
         ForeColor       =   &H80000008&
         Height          =   220
         Index           =   3
         Left            =   3300
         TabIndex        =   68
         Top             =   3105
         Width           =   1200
      End
      Begin VB.Label lblMiscUnits 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         Caption         =   "(m/hr)"
         ForeColor       =   &H80000008&
         Height          =   315
         Index           =   4
         Left            =   3300
         TabIndex        =   67
         Top             =   2805
         Width           =   1200
      End
      Begin VB.Label lblMiscUnits 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         Caption         =   "-"
         ForeColor       =   &H80000008&
         Height          =   315
         Index           =   5
         Left            =   3300
         TabIndex        =   66
         Top             =   2505
         Width           =   1200
      End
      Begin VB.Label lblBed 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         Caption         =   "Interstitial Velocity"
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   7
         Left            =   240
         TabIndex        =   65
         Top             =   3105
         Width           =   1710
      End
      Begin VB.Label lblBed 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         Caption         =   "Superficial Velocity"
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   8
         Left            =   240
         TabIndex        =   64
         Top             =   2805
         Width           =   1710
      End
      Begin VB.Label lblBed 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         Caption         =   "Bed Porosity"
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   9
         Left            =   240
         TabIndex        =   63
         Top             =   2505
         Width           =   1710
      End
   End
   Begin Threed.SSFrame ssframe_PSDM 
      Height          =   3075
      Left            =   120
      TabIndex        =   41
      Top             =   3000
      Width           =   4605
      _Version        =   65536
      _ExtentX        =   8123
      _ExtentY        =   5424
      _StockProps     =   14
      Caption         =   "Simulation Parameters for PSDM Only:"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Begin VB.CommandButton cmdParamsPSDMInRoom 
         Appearance      =   0  'Flat
         Caption         =   "Edi&t Parameters for PSDMR Model"
         Height          =   315
         Left            =   150
         TabIndex        =   95
         TabStop         =   0   'False
         Top             =   2690
         Width           =   4305
      End
      Begin VB.TextBox txtTime 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         Height          =   285
         Index           =   0
         Left            =   2100
         TabIndex        =   2
         Text            =   "txtTime"
         Top             =   280
         Width           =   1215
      End
      Begin VB.TextBox txtTime 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         Height          =   285
         Index           =   1
         Left            =   2100
         TabIndex        =   3
         Text            =   "txtTime"
         Top             =   610
         Width           =   1215
      End
      Begin VB.TextBox txtTime 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         Height          =   285
         Index           =   2
         Left            =   2100
         TabIndex        =   4
         Text            =   "txtTime"
         Top             =   950
         Width           =   1215
      End
      Begin VB.TextBox txtNumberOfBeds 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   2880
         TabIndex        =   5
         Text            =   "Test"
         Top             =   1320
         Width           =   735
      End
      Begin VB.ComboBox txtTimeUnits 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         Height          =   315
         Index           =   0
         Left            =   3360
         Style           =   2  'Dropdown List
         TabIndex        =   46
         TabStop         =   0   'False
         Top             =   270
         Width           =   1035
      End
      Begin VB.ComboBox txtTimeUnits 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         Height          =   315
         Index           =   1
         Left            =   3360
         Style           =   2  'Dropdown List
         TabIndex        =   45
         TabStop         =   0   'False
         Top             =   600
         Width           =   1035
      End
      Begin VB.ComboBox txtTimeUnits 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         Height          =   315
         Index           =   2
         Left            =   3360
         Style           =   2  'Dropdown List
         TabIndex        =   44
         TabStop         =   0   'False
         Top             =   930
         Width           =   1035
      End
      Begin Threed.SSFrame Frame3D2 
         Height          =   960
         Left            =   120
         TabIndex        =   47
         Top             =   1650
         Width           =   4335
         _Version        =   65536
         _ExtentX        =   7646
         _ExtentY        =   1693
         _StockProps     =   14
         Caption         =   "Number of Collocation Points:"
         ForeColor       =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ShadowColor     =   1
         Begin VB.TextBox txtNPoint 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            Height          =   285
            Index           =   0
            Left            =   2760
            TabIndex        =   6
            Text            =   "Test"
            Top             =   250
            Width           =   735
         End
         Begin VB.TextBox txtNPoint 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            Height          =   285
            Index           =   1
            Left            =   2760
            TabIndex        =   7
            Text            =   "Test"
            Top             =   590
            Width           =   735
         End
         Begin Spin.SpinButton spnPoint 
            Height          =   285
            Index           =   1
            Left            =   3480
            TabIndex        =   48
            Top             =   590
            Width           =   255
            _Version        =   65536
            _ExtentX        =   2646
            _ExtentY        =   1323
            _StockProps     =   73
         End
         Begin Spin.SpinButton spnPoint 
            Height          =   285
            Index           =   0
            Left            =   3480
            TabIndex        =   49
            Top             =   250
            Width           =   255
            _Version        =   65536
            _ExtentX        =   2646
            _ExtentY        =   1323
            _StockProps     =   73
         End
         Begin VB.Label lblText 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H00C0C0C0&
            Caption         =   "Axial Direction"
            ForeColor       =   &H80000008&
            Height          =   255
            Index           =   0
            Left            =   540
            TabIndex        =   51
            Top             =   280
            Width           =   1875
         End
         Begin VB.Label lblText 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H00C0C0C0&
            Caption         =   "Radial Direction"
            ForeColor       =   &H80000008&
            Height          =   255
            Index           =   1
            Left            =   540
            TabIndex        =   50
            Top             =   600
            Width           =   1875
         End
      End
      Begin Spin.SpinButton spnNumberOfBeds 
         Height          =   285
         Left            =   3600
         TabIndex        =   52
         Top             =   1320
         Width           =   255
         _Version        =   65536
         _ExtentX        =   2646
         _ExtentY        =   1323
         _StockProps     =   73
      End
      Begin VB.Label lblTime 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         Caption         =   "Total Run Time"
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   0
         Left            =   180
         TabIndex        =   56
         Top             =   310
         Width           =   1875
      End
      Begin VB.Label lblTime 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         Caption         =   "First Point Displayed"
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   1
         Left            =   180
         TabIndex        =   55
         Top             =   630
         Width           =   1875
      End
      Begin VB.Label lblTime 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         Caption         =   "Time Step"
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   2
         Left            =   180
         TabIndex        =   54
         Top             =   980
         Width           =   1875
      End
      Begin VB.Label lblAxialElementsDesc 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Number of Axial Elements"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   240
         TabIndex        =   53
         Top             =   1350
         Width           =   2475
      End
   End
   Begin Threed.SSFrame ssframe_Component 
      Height          =   2145
      Left            =   120
      TabIndex        =   35
      Top             =   960
      Width           =   4605
      _Version        =   65536
      _ExtentX        =   8123
      _ExtentY        =   3784
      _StockProps     =   14
      Caption         =   "Component Properties:"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Begin VB.CommandButton cmdADEComponent 
         Appearance      =   0  'Flat
         Caption         =   "&Edit Properties"
         Height          =   315
         HelpContextID   =   20
         Index           =   2
         Left            =   2820
         TabIndex        =   40
         TabStop         =   0   'False
         Top             =   1670
         Width           =   1515
      End
      Begin VB.CommandButton cmdADEComponent 
         Appearance      =   0  'Flat
         Caption         =   "De&lete"
         Height          =   315
         HelpContextID   =   20
         Index           =   1
         Left            =   1800
         TabIndex        =   39
         TabStop         =   0   'False
         Top             =   1670
         Width           =   975
      End
      Begin VB.CommandButton cmdADEComponent 
         Appearance      =   0  'Flat
         Caption         =   "&Add"
         Height          =   315
         HelpContextID   =   20
         Index           =   0
         Left            =   240
         TabIndex        =   38
         TabStop         =   0   'False
         Top             =   1670
         Width           =   1035
      End
      Begin VB.ListBox lstComponents 
         Height          =   1035
         HelpContextID   =   20
         Left            =   120
         MultiSelect     =   1  'Simple
         TabIndex        =   37
         TabStop         =   0   'False
         Top             =   255
         Width           =   4335
      End
      Begin VB.ComboBox cboSelectCompo 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   240
         Style           =   2  'Dropdown List
         TabIndex        =   36
         TabStop         =   0   'False
         Top             =   1300
         Width           =   4095
      End
   End
   Begin Threed.SSFrame ssframe_Water 
      Height          =   945
      Left            =   120
      TabIndex        =   29
      Top             =   60
      Width           =   4605
      _Version        =   65536
      _ExtentX        =   8123
      _ExtentY        =   1667
      _StockProps     =   14
      Caption         =   "Water Properties:"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Begin VB.CommandButton cmdWaterCorrelations 
         Appearance      =   0  'Flat
         Caption         =   "Correlations"
         Height          =   615
         Left            =   2940
         TabIndex        =   30
         TabStop         =   0   'False
         Top             =   240
         Width           =   1515
      End
      Begin VB.TextBox txtWater 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         Height          =   288
         HelpContextID   =   20
         Index           =   1
         Left            =   1260
         TabIndex        =   0
         Text            =   "txtWater(1)"
         Top             =   270
         Width           =   1092
      End
      Begin VB.TextBox txtWater 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         Height          =   288
         HelpContextID   =   20
         Index           =   0
         Left            =   1260
         TabIndex        =   1
         Text            =   "txtWater(0)"
         Top             =   575
         Width           =   1092
      End
      Begin VB.Label lblWater 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Pressure"
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   1
         Left            =   30
         TabIndex        =   34
         Top             =   300
         Width           =   1125
      End
      Begin VB.Label lblWater 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Temperature"
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   0
         Left            =   30
         TabIndex        =   33
         Top             =   600
         Width           =   1125
      End
      Begin VB.Label lblWaterUnit 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "lblWaterUnit(1)"
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   1
         Left            =   2400
         TabIndex        =   32
         Top             =   300
         Width           =   495
      End
      Begin VB.Label lblWaterUnit 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "lblWaterUnit(0)"
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   0
         Left            =   2400
         TabIndex        =   31
         Top             =   600
         Width           =   495
      End
   End
   Begin Threed.SSPanel SSPanel2 
      Align           =   2  'Align Bottom
      Height          =   405
      Left            =   0
      TabIndex        =   18
      Top             =   6510
      Width           =   9540
      _Version        =   65536
      _ExtentX        =   16828
      _ExtentY        =   714
      _StockProps     =   15
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Begin Threed.SSPanel sspanel_Dirty 
         Height          =   285
         Left            =   60
         TabIndex        =   19
         Top             =   60
         Width           =   2115
         _Version        =   65536
         _ExtentX        =   3731
         _ExtentY        =   503
         _StockProps     =   15
         Caption         =   "sspanel_Dirty"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BevelOuter      =   1
      End
      Begin Threed.SSPanel sspanel_Status 
         Height          =   285
         Left            =   2220
         TabIndex        =   20
         Top             =   60
         Width           =   7185
         _Version        =   65536
         _ExtentX        =   12674
         _ExtentY        =   503
         _StockProps     =   15
         Caption         =   "sspanel_Status"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BevelOuter      =   1
      End
   End
   Begin Threed.SSFrame SSFrame4 
      Height          =   1035
      Left            =   9750
      TabIndex        =   92
      Top             =   270
      Visible         =   0   'False
      Width           =   3015
      _Version        =   65536
      _ExtentX        =   5318
      _ExtentY        =   1826
      _StockProps     =   14
      Caption         =   "Used -- Invisible"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Begin MSComDlg.CommonDialog CommonDialog1 
         Left            =   240
         Top             =   240
         _ExtentX        =   847
         _ExtentY        =   847
         _Version        =   393216
      End
   End
   Begin Threed.SSFrame SSFrame2 
      Height          =   3765
      Left            =   10770
      TabIndex        =   21
      Top             =   1710
      Visible         =   0   'False
      Width           =   4425
      _Version        =   65536
      _ExtentX        =   7805
      _ExtentY        =   6641
      _StockProps     =   14
      Caption         =   "Old -- Invisible"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Begin VB.ComboBox Combo1 
         Height          =   315
         Left            =   2580
         Style           =   2  'Dropdown List
         TabIndex        =   28
         Top             =   1410
         Width           =   1245
      End
      Begin Threed.SSFrame SSFrame1 
         Height          =   3015
         Left            =   240
         TabIndex        =   22
         Top             =   300
         Width           =   2115
         _Version        =   65536
         _ExtentX        =   3731
         _ExtentY        =   5318
         _StockProps     =   14
         Caption         =   "SSFrame1"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Begin VB.ListBox List1 
            Height          =   450
            Left            =   330
            TabIndex        =   25
            Top             =   2370
            Width           =   1245
         End
         Begin VB.CommandButton Command1 
            Caption         =   "Command1"
            Height          =   595
            Left            =   330
            TabIndex        =   24
            Top             =   510
            Width           =   1245
         End
         Begin VB.TextBox Text1 
            Height          =   525
            Left            =   300
            TabIndex        =   23
            Text            =   "Text1"
            Top             =   1080
            Width           =   1245
         End
         Begin VB.Label Label1 
            Caption         =   "Label1"
            Height          =   525
            Left            =   330
            TabIndex        =   26
            Top             =   1710
            Width           =   1245
         End
      End
      Begin Spin.SpinButton SpinButton1 
         Height          =   765
         Left            =   2520
         TabIndex        =   27
         Top             =   480
         Width           =   825
         _Version        =   65536
         _ExtentX        =   1455
         _ExtentY        =   1349
         _StockProps     =   73
         ForeColor       =   -2147483630
         BackColor       =   -2147483633
      End
   End
   Begin VB.Menu mnuFile 
      Caption         =   "&File"
      Begin VB.Menu mnuFileItem 
         Caption         =   "&New"
         Index           =   0
         Shortcut        =   ^N
      End
      Begin VB.Menu mnuFileItem 
         Caption         =   "&Open ..."
         Index           =   1
         Shortcut        =   ^O
      End
      Begin VB.Menu mnuFileItem 
         Caption         =   "&Save"
         Index           =   2
         Shortcut        =   ^S
      End
      Begin VB.Menu mnuFileItem 
         Caption         =   "Save &As ..."
         Index           =   3
      End
      Begin VB.Menu mnuFileItem 
         Caption         =   "-"
         Index           =   5
      End
      Begin VB.Menu mnuFileItem 
         Caption         =   "Select P&rinter"
         Index           =   6
      End
      Begin VB.Menu mnuFileItem 
         Caption         =   "&Print ..."
         Index           =   7
         Begin VB.Menu mnuPrintSubItem 
            Caption         =   "To &Printer"
            Index           =   0
         End
         Begin VB.Menu mnuPrintSubItem 
            Caption         =   "To a &File"
            Index           =   1
         End
      End
      Begin VB.Menu mnuFileItem 
         Caption         =   "-"
         Index           =   190
      End
      Begin VB.Menu mnuFileItem 
         Caption         =   "&1 Old File #1"
         Index           =   191
         Visible         =   0   'False
      End
      Begin VB.Menu mnuFileItem 
         Caption         =   "&2 Old File #2"
         Index           =   192
         Visible         =   0   'False
      End
      Begin VB.Menu mnuFileItem 
         Caption         =   "&3 Old File #3"
         Index           =   193
         Visible         =   0   'False
      End
      Begin VB.Menu mnuFileItem 
         Caption         =   "&4 Old File #4"
         Index           =   194
         Visible         =   0   'False
      End
      Begin VB.Menu mnuFileItem 
         Caption         =   "-"
         Index           =   199
         Visible         =   0   'False
      End
      Begin VB.Menu mnuFileItem 
         Caption         =   "E&xit"
         Index           =   200
      End
   End
   Begin VB.Menu mnuPhase 
      Caption         =   "&Phase"
      Begin VB.Menu mnuPhaseItem 
         Caption         =   "&Liquid Phase"
         Checked         =   -1  'True
         Index           =   0
      End
      Begin VB.Menu mnuPhaseItem 
         Caption         =   "&Gas Phase"
         Index           =   1
      End
   End
   Begin VB.Menu mnuRun 
      Caption         =   "&Run"
      Begin VB.Menu mnuRunItem 
         Caption         =   "&PSDM"
         Index           =   0
         Shortcut        =   +{F3}
      End
      Begin VB.Menu mnuRunItem 
         Caption         =   "&CPHSDM"
         Index           =   1
         Shortcut        =   +{F4}
      End
      Begin VB.Menu mnuRunItem 
         Caption         =   "&ECM"
         Index           =   2
         Shortcut        =   +{F5}
      End
      Begin VB.Menu mnuRunItem 
         Caption         =   "PSDMR in &Room"
         Index           =   10
         Shortcut        =   +{F6}
      End
      Begin VB.Menu mnuRunItem 
         Caption         =   "PSDMR &Alone"
         Index           =   20
         Shortcut        =   +{F7}
         Visible         =   0   'False
      End
   End
   Begin VB.Menu mnuDisplay 
      Caption         =   "Re&sults"
      Begin VB.Menu mnuResultsItem 
         Caption         =   "&PSDM Results"
         Index           =   0
         Shortcut        =   {F3}
      End
      Begin VB.Menu mnuResultsItem 
         Caption         =   "&CPHSDM Results"
         Index           =   1
         Shortcut        =   {F4}
      End
      Begin VB.Menu mnuResultsItem 
         Caption         =   "&ECM Results"
         Index           =   2
         Shortcut        =   {F5}
      End
      Begin VB.Menu mnuResultsItem 
         Caption         =   "Compare PSDM Results to &Data"
         Index           =   3
         Shortcut        =   {F6}
      End
      Begin VB.Menu mnuResultsItem 
         Caption         =   "Compare CPHSDM Results to D&ata"
         Index           =   4
         Shortcut        =   {F7}
      End
      Begin VB.Menu mnuResultsItem 
         Caption         =   "PSDM in &Room Results"
         Index           =   10
         Visible         =   0   'False
      End
   End
   Begin VB.Menu mnuOptions 
      Caption         =   "&Options"
      Begin VB.Menu mnuOptionsItem 
         Caption         =   "&Fouling of GAC"
         Index           =   0
      End
      Begin VB.Menu mnuOptionsItem 
         Caption         =   "&Influent concentrations"
         Index           =   1
      End
      Begin VB.Menu mnuOptionsItem 
         Caption         =   "&Effluent concentrations"
         Index           =   2
      End
   End
   Begin VB.Menu mnuDatabase 
      Caption         =   "&Databases"
      Begin VB.Menu mnuDatabasesItem 
         Caption         =   "&Adsorbent Database"
         Index           =   0
      End
      Begin VB.Menu mnuDatabasesItem 
         Caption         =   "&Isotherm Database"
         Index           =   1
      End
      Begin VB.Menu mnuDatabasesItem 
         Caption         =   "A&dsorber Database"
         Index           =   2
      End
   End
   Begin VB.Menu mnuHelp 
      Caption         =   "&Help"
      Begin VB.Menu mnuHelpItem 
         Caption         =   "&Search ... (INVISIBLE)"
         Enabled         =   0   'False
         Index           =   0
         Visible         =   0   'False
      End
      Begin VB.Menu mnuHelpItem 
         Caption         =   "&How to use Help ... (INVISIBLE)"
         Enabled         =   0   'False
         Index           =   1
         Visible         =   0   'False
      End
      Begin VB.Menu mnuHelpItem 
         Caption         =   "&Online Help ..."
         Enabled         =   0   'False
         Index           =   10
         Visible         =   0   'False
      End
      Begin VB.Menu mnuHelpItem 
         Caption         =   "Online Manual ..."
         Index           =   20
      End
      Begin VB.Menu mnuHelpItem 
         Caption         =   "Manual Printing Instructions ..."
         Index           =   22
      End
      Begin VB.Menu mnuHelpItem 
         Caption         =   "-"
         Index           =   30
      End
      Begin VB.Menu mnuHelpItem 
         Caption         =   "View Version History ..."
         Index           =   80
      End
      Begin VB.Menu mnuHelpItem 
         Caption         =   "View Disclaimer ..."
         Index           =   85
         Visible         =   0   'False
      End
      Begin VB.Menu mnuHelpItem 
         Caption         =   "&Technical Assistance Provided By ..."
         Index           =   90
         Visible         =   0   'False
      End
      Begin VB.Menu mnuHelpItem 
         Caption         =   "-"
         Index           =   98
      End
      Begin VB.Menu mnuHelpItem 
         Caption         =   "&About AdDesignS ..."
         Index           =   99
      End
   End
   Begin VB.Menu mnuOther 
      Caption         =   "Other"
      Index           =   5
      Visible         =   0   'False
      Begin VB.Menu mnuOtherItem 
         Caption         =   "&Technical Help"
         Index           =   1
      End
   End
   Begin VB.Menu mnuMTU 
      Caption         =   "&MTU Internal"
      Begin VB.Menu mnuMTUItem 
         Caption         =   "&Create alternate input file"
         Index           =   10
      End
      Begin VB.Menu mnuMTUItem 
         Caption         =   "&Keep temporary model files"
         Index           =   40
      End
      Begin VB.Menu mnuMTUItem 
         Caption         =   "&Make menu invisible"
         Index           =   50
      End
      Begin VB.Menu mnuMTUItem 
         Caption         =   "&Read me"
         Index           =   199
      End
   End
   Begin VB.Menu mnuUnused 
      Caption         =   "Unused"
      Visible         =   0   'False
      Begin VB.Menu mnuBatch 
         Caption         =   "&Batch Files"
      End
      Begin VB.Menu mnuBatchItem 
         Caption         =   "&PSDM"
         Index           =   0
         Visible         =   0   'False
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit






Const frmMain_declarations_end = True


Sub frmMain_Reset_DemoVersionDisablings()
  If (IsThisADemo() = True) Then
    mnuFileItem(0).Enabled = False
    ''''mnuFileItem(1).Enabled = False
    mnuFileItem(2).Enabled = False
    mnuFileItem(3).Enabled = False
    ''''mnuFileItem(191).Enabled = False
    ''''mnuFileItem(192).Enabled = False
    ''''mnuFileItem(193).Enabled = False
    ''''mnuFileItem(194).Enabled = False
    mnuPhaseItem(0).Enabled = False
    mnuPhaseItem(1).Enabled = False
    cmdADEComponent(0).Enabled = False
    cmdADEComponent(1).Enabled = False
  End If
End Sub


Sub Avoid_Weird_Focus_Problem()
  Call unitsys_control_MostRecent_Force_lostfocus
  frmMain.lstComponents.SetFocus
End Sub


Sub Populate_frmMain_Units()
'  'Fixed Bed Properties:
'  Call Populate_Length_Units(txtBedUnits(0), LENGTH_M)
'  Call Populate_Length_Units(txtBedUnits(1), LENGTH_M)
'  Call Populate_Mass_Units(txtBedUnits(2), MASS_KG)
'  Call Populate_Flowrate_Units(txtBedUnits(3), FLOW_M3_per_S)
'  Call Populate_Time_Units(txtBedUnits(4), TIME_MIN)
'  'time properties
'  Call Populate_Time_Units(txttimeunits(0), TIME_D)
'  Call Populate_Time_Units(txttimeunits(1), TIME_D)
'  Call Populate_Time_Units(txttimeunits(2), TIME_D)
'  'Adsorbent Properties:
'  Call Populate_Density_Units(txtCarbonUnits(1), APPARENT_DENSITY_G_per_ML)
'  Call Populate_Length_Units(txtCarbonUnits(2), LENGTH_M)

  'WATER/AIR PROPERTIES.
  Call unitsys_register(frmMain, lblWater(1), txtWater(1), Nothing, "", _
      "", "", "", "", 100#, False)
  Call unitsys_register(frmMain, lblWater(0), txtWater(0), Nothing, "", _
      "", "", "", "", 100#, False)
  'PSDM PARAMETERS.
  Call unitsys_register(frmMain, lblTime(0), txtTime(0), txtTimeUnits(0), "time", _
      "d", "min", "", "", 100#, True)
  Call unitsys_register(frmMain, lblTime(1), txtTime(1), txtTimeUnits(1), "time", _
      "d", "min", "", "", 100#, True)
  Call unitsys_register(frmMain, lblTime(2), txtTime(2), txtTimeUnits(2), "time", _
      "d", "min", "", "", 100#, True)
  Call unitsys_register(frmMain, lblAxialElementsDesc, txtNumberOfBeds, Nothing, "", _
      "", "", "0", "0", 100#, False)
  Call unitsys_register(frmMain, lblText(0), txtNPoint(0), Nothing, "", _
      "", "", "0", "0", 100#, False)
  Call unitsys_register(frmMain, lblText(1), txtNPoint(1), Nothing, "", _
      "", "", "0", "0", 100#, False)
  'BED PROPERTIES.
  Call unitsys_register(frmMain, lblBed(0), txtBedValue(0), txtBedUnits(0), "length", _
      "m", "m", "", "", 100#, True)
  Call unitsys_register(frmMain, lblBed(1), txtBedValue(1), txtBedUnits(1), "length", _
      "m", "m", "", "", 100#, True)
  Call unitsys_register(frmMain, lblBed(2), txtBedValue(2), txtBedUnits(2), "mass", _
      "kg", "kg", "", "", 100#, True)
  Call unitsys_register(frmMain, lblBed(3), txtBedValue(3), txtBedUnits(3), "flow_volumetric", _
      "m³/s", "m³/s", "", "", 100#, True)
  Call unitsys_register(frmMain, lblBed(4), txtBedValue(4), txtBedUnits(4), "time", _
      "s", "s", "", "", 100#, True)
  'ADSORBENT PROPERTIES.
  Call unitsys_register(frmMain, lblCarbon(1), txtCarbon(1), txtCarbonUnits(1), "density", _
      "g/mL", "g/mL", "", "", 100#, True)
  Call unitsys_register(frmMain, lblCarbon(2), txtCarbon(2), txtCarbonUnits(2), "length", _
      "m", "m", "", "", 100#, True)
  Call unitsys_register(frmMain, lblCarbon(3), txtCarbon(3), Nothing, "", _
      "", "", "", "", 100#, False)
  Call unitsys_register(frmMain, lblCarbon(4), txtCarbon(4), Nothing, "", _
      "", "", "", "", 100#, False)
End Sub


Private Sub cmdADEComponent_Click(Index As Integer)
Dim Raise_Dirty_Flag As Boolean
Dim temp As String
Dim RetVal As Integer
Dim N As Integer
Dim i As Integer
Dim J As Integer
  Select Case Index
    Case 0:   'ADD.
      Call frmCompoProp.frmCompoProp_Add(Raise_Dirty_Flag)
      If (Raise_Dirty_Flag) Then
        'RAISE DIRTY FLAG.
        Call DirtyStatus_Throw
      End If
    Case 1:   'DELETE.
      If (cboSelectCompo.ListIndex = -1) Or _
          (cboSelectCompo.ListIndex > Number_Component - 1) Then
        Call Show_Error("You must first select a component.")
        Exit Sub
      End If
      temp = Trim$(cboSelectCompo.List(cboSelectCompo.ListIndex))
      RetVal = MsgBox("Do you really want to delete component '" & _
          temp & "' ?", vbQuestion + vbYesNo, _
          AppName_For_Display_Short & " : Delete Component ?")
      If RetVal = vbNo Then Exit Sub
      N = cboSelectCompo.ListIndex + 1
      '
      ' DELETE COMPONENT FROM MAIN COMPONENT PROPERTIES DATA AREA.
      '
      For i = N To Number_Component - 1
        Component(i) = Component(i + 1)
        For J = 1 To 400
          C_Influent(i, J) = C_Influent(i + 1, J)
          C_Data_Points(i, J) = C_Data_Points(i + 1, J)
        Next J
      Next i
      Number_Component = Number_Component - 1
      '
      ' DELETE COMPONENT FROM ROOM PROPERTIES DATA AREA.
      '
      For i = N To RoomParams.COUNT_CONTAMINANT - 1
        RoomParams.ROOM_C0(i) = RoomParams.ROOM_C0(i + 1)
        RoomParams.ROOM_EMIT(i) = RoomParams.ROOM_EMIT(i + 1)
        RoomParams.ROOM_SS_VALUE(i) = RoomParams.ROOM_SS_VALUE(i + 1)
        RoomParams.INITIAL_ROOM_CONC(i) = RoomParams.INITIAL_ROOM_CONC(i + 1)
        RoomParams.RXN_RATE_CONSTANT(i) = RoomParams.RXN_RATE_CONSTANT(i + 1)
        RoomParams.RXN_PRODUCT(i) = RoomParams.RXN_PRODUCT(i + 1)
        RoomParams.RXN_RATIO(i) = RoomParams.RXN_RATIO(i + 1)
      Next i
      RoomParams.COUNT_CONTAMINANT = RoomParams.COUNT_CONTAMINANT - 1
      '
      ' RAISE DIRTY FLAG AND REFRESH MAIN WINDOW.
      '
      Call DirtyStatus_Throw
      Call frmMain_Refresh
    Case 2:   'EDIT.
      If (cboSelectCompo.ListIndex = -1) Or _
          (cboSelectCompo.ListIndex > Number_Component - 1) Then
        Call Show_Error("You must first select a component.")
        Exit Sub
      End If
      Call frmCompoProp.frmCompoProp_Edit( _
          Raise_Dirty_Flag, _
          cboSelectCompo.ListIndex + 1)
      If (Raise_Dirty_Flag) Then
        'RAISE DIRTY FLAG.
        Call DirtyStatus_Throw
      End If
  End Select
End Sub
Private Sub cmdAdsorberDB_Click()
Dim User_Transferred_Data As Boolean
Dim New_L As Double
Dim New_D As Double
Dim New_M As Double
Dim New_Q As Double
  Call frmEditAdsorber.frmEditAdsorber_QueryDatabase( _
      User_Transferred_Data)
  If (Not User_Transferred_Data) Then Exit Sub
  'TRANSFER DATA TO MAIN WINDOW.
  ''SET L,D,M,Q UNITS TO M, M, KG, and M^3/S RESPECTIVELY
  'txtBedUnits(0).ListIndex = 0
  'txtBedUnits(1).ListIndex = 0
  'txtBedUnits(2).ListIndex = 0
  'txtBedUnits(3).ListIndex = 0
  'TRANSFER PARAMETERS BACK TO MAIN SCREEN
  New_L = frmEditAdsorber_ReturnParameters.L
  New_D = frmEditAdsorber_ReturnParameters.D
  New_M = frmEditAdsorber_ReturnParameters.M
  New_Q = frmEditAdsorber_ReturnParameters.Q
  'txtBedValue(0).Text = Format_It(New_L, 3)
  'txtBedValue(1).Text = Format_It(New_D, 3)
  'txtBedValue(2).Text = Format_It(New_M, 2)
  'txtBedValue(3).Text = Format_It(New_Q, 3)
  Bed.length = New_L
  Bed.Diameter = New_D
  Bed.Weight = New_M
  Bed.Flowrate = New_Q
  'RAISE DIRTY FLAG.
  Call DirtyStatus_Throw
  'UPDATE WINDOW DISPLAY.
  Call frmMain_Refresh
  ''UPDATE SOME BED PROPERTY DISPLAYS:
  'Call Update_Bed_Density_Display
  'Call Update_Several_Bed_Properties(3)
End Sub
Private Sub cmdCarbon_Click()
Dim User_Transferred_Data As Boolean
  Call frmEditCarbon.frmEditCarbon_QueryDatabase( _
      User_Transferred_Data)
  If (Not User_Transferred_Data) Then Exit Sub
  '
  '    DATA WAS ALREADY TRANSFERRED IN THE SUB-WINDOW.
  '    NO FURTHER DATA TRANSFER IS REQUIRED.
  '    ONLY THE DIRTY FLAG AND THE WINDOW DISPLAY
  '    UPDATE NEED BE PERFORMED (SEE BELOW).
  '
  'RAISE DIRTY FLAG.
  Call DirtyStatus_Throw
  'UPDATE WINDOW DISPLAY.
  Call frmMain_Refresh
End Sub



Private Sub cmdNote_Click(Index As Integer)
Dim Temp_FileNote As String
Dim RaiseDirtyFlag As Boolean
  Temp_FileNote = FileNote
  Call frmFileNote.frmFileNote_Run( _
     Temp_FileNote, _
     RaiseDirtyFlag)
  If (RaiseDirtyFlag) Then
    FileNote = Temp_FileNote
    'THROW DIRTY FLAG.
    Call DirtyStatus_Throw
    'REFRESH WINDOW.
    Call frmMain_Refresh
  End If
End Sub
Private Sub cmdParamsPSDMInRoom_Click()
Dim Raise_Dirty_Flag As Boolean
  Call frmInputParamsPSDMInRoom. _
      frmInputParamsPSDMInRoom_Edit(Raise_Dirty_Flag)
  If (Raise_Dirty_Flag) Then
    'THROW DIRTY FLAG.
    Call DirtyStatus_Throw
  End If
End Sub
Private Sub cmdpolanyi_Click()
Dim Raise_Dirty_Flag As Boolean
  Call frmPolanyi.frmPolanyi_Edit(Me, Raise_Dirty_Flag)
  If (Raise_Dirty_Flag) Then
    'THROW DIRTY FLAG.
    Call DirtyStatus_Throw
  End If
End Sub
Private Sub cmdViewDimensionless_Click()
  frmDimensionless.Show 1
End Sub
Private Sub cmdWaterCorrelations_Click()
Dim Raise_Dirty_Flag As Boolean
  Call frmFluidProps.frmFluidProps_Edit(Raise_Dirty_Flag)
  If (Raise_Dirty_Flag) Then
    'THROW DIRTY FLAG.
    Call DirtyStatus_Throw
  End If
End Sub



Private Sub Command4_Click()
    Set Picture1.Picture = CaptureActiveWindow()
    PrintPictureToFitPage Printer, Picture1.Picture
    Printer.EndDoc
    ' Set focus back to form.
    Me.SetFocus
End Sub

Private Sub Form_Load()
Dim is_internal_mtu As Boolean
Dim TurnOff_ForPSDMInRoom As Boolean
  '
  ' MISC INITS.
  '
  sspanel_Dirty.Caption = ""
  sspanel_Status.Caption = ""
  Me.Caption = AppName_For_Display_Short
  Me.Width = 9600
  Me.Height = 7600
  Call CenterOnScreen(Me)
  CommonDialog1.Filename = App.Path & "\examples\*.dat"
  lblWaterUnit(0).Caption = "C"
  lblWaterUnit(1).Caption = "atm"
  cmdNote(1).Move cmdNote(0).Left, cmdNote(0).Top
  cmdNote(0).Visible = False
  cmdNote(1).Visible = False
  '
  ' CHECK FOR FILE THAT INDICATES THIS IS INTERNAL TO MTU:
  '
  is_internal_mtu = False
  If (check_internal_to_mtu()) Then is_internal_mtu = True
  mnuMTU.Visible = is_internal_mtu
  
  '///Modefication///Sinan///07/03/06, adding bouth the PSDM and the PSDM in room Models
  'for the Run menu.
  ' PSDM IN ROOM INITS.
  '
'  TurnOff_ForPSDMInRoom = False
'  If (Activate_PSDMInRoom = True) Then
'    TurnOff_ForPSDMInRoom = True
'  End If
'  If (is_internal_mtu = True) Then
'    TurnOff_ForPSDMInRoom = False
'  End If
'  mnuRunItem(0).Visible = Not TurnOff_ForPSDMInRoom
'  mnuRunItem(1).Visible = Not TurnOff_ForPSDMInRoom
'  mnuRunItem(2).Visible = Not TurnOff_ForPSDMInRoom
'  mnuRunItem(10).Visible = Activate_PSDMInRoom
'  mnuRunItem(20).Visible = Activate_PSDMInRoom
'  mnuResultsItem(1).Visible = Not TurnOff_ForPSDMInRoom
'  mnuResultsItem(2).Visible = Not TurnOff_ForPSDMInRoom
'  mnuResultsItem(4).Visible = Not TurnOff_ForPSDMInRoom
'  If (is_internal_mtu = True) Then
'    mnuRunItem(0).Caption = mnuRunItem(0).Caption & " (*)"
'    mnuRunItem(1).Caption = mnuRunItem(1).Caption & " (*)"
'    mnuRunItem(2).Caption = mnuRunItem(2).Caption & " (*)"
'    mnuResultsItem(1).Caption = mnuResultsItem(1).Caption & " (*)"
'    mnuResultsItem(2).Caption = mnuResultsItem(2).Caption & " (*)"
'    mnuResultsItem(4).Caption = mnuResultsItem(4).Caption & " (*)"
'  End If
  ''''mnuResultsItem(10).Visible = Activate_PSDMInRoom
  'cmdParamsPSDMInRoom.Visible = Activate_PSDMInRoom
  '///End of Modefication.///
  
  ' POPULATE UNITS INTO SCROLLBOX CONTROLS.
  '
  Call Populate_frmMain_Units
  '
  ' CREATE A NEW FILE IN MEMORY.
  '
  Call file_new
  '
  ' POPULATE LAST-FEW-FILES LIST.
  '
  Call OldFileList_Populate( _
      1, _
      frmMain.mnuFileItem(199), _
      frmMain.mnuFileItem(191), _
      frmMain.mnuFileItem(192), _
      frmMain.mnuFileItem(193), _
      frmMain.mnuFileItem(194))
End Sub
Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
  If (file_query_unload() = False) Then
    Cancel = True
  End If
End Sub
Private Sub Form_Unload(Cancel As Integer)
  Call frmMain_Close_All_Windows
  Call unitsys_unregister_all_on_form(Me)
End Sub


Private Sub lstComponents_Click()
Dim i As Integer
  'Debug.Print "lstComponents_Click"
  'For i = 0 To lstComponents.ListCount - 1
  '  Debug.Print "lst.selected(" & Trim$(Str$(i)) & ") = " & _
  '      Trim$(Str$(lstComponents.Selected(i)))
  'Next i
  For i = 0 To lstComponents.ListCount - 1
    Component(i + 1).Is_Selected_On_List = (lstComponents.Selected(i))
  Next i
End Sub


Private Sub mnuDatabasesItem_Click(Index As Integer)
  Select Case Index
    Case 0:         'ADSORBENT DB.
      Call frmEditCarbon.frmEditCarbon_EditDatabase
    Case 1:         'ISOTHERM DB.
      Call frmEditIsotherm.frmEditIsotherm_EditDatabase
    Case 2:         'ADSORBER DB.
      Call frmEditAdsorber.frmEditAdsorber_EditDatabase
  End Select
End Sub
Private Sub mnuFileItem_Click(Index As Integer)
  Select Case Index
    Case 0:      'New
      If (file_query_unload()) Then
        Call Avoid_Weird_Focus_Problem
        Call file_new
      End If
    Case 1:      'Open ...
      If (file_query_unload()) Then
        Call Avoid_Weird_Focus_Problem
        Call File_OpenAs("")
      End If
    Case 2:      'Save
      If (Filename = "") Then
        Call Avoid_Weird_Focus_Problem
        Call File_SaveAs("")
      Else
        Call Avoid_Weird_Focus_Problem
        Call File_SaveAs(Filename)
      End If
    Case 3:      'Save As ...
      Call Avoid_Weird_Focus_Problem
      Call File_SaveAs("")
    Case 6:       'Select Printer ...
      CommonDialog1.ShowPrinter
    'Case 85:      'Print ...
    '  frmPrint.Show 1
    Case 191 To 194:      'Last-few-files list
      If (file_query_unload()) Then
        If (mnuFileItem(Index).Visible) Then
          Call File_OpenAs(OldFiles(1, Index - 190))
        End If
      End If
    Case 200:     'Exit
      'NOTE: MDIForm_QueryUnload() TAKES CAKE OF THIS.
      'If we do it here, _two_ message boxes will pop up
      'when the user has data which needs saving !
      'If (file_query_unload()) Then
      '  Unload Me
      'End If
      Unload Me
  End Select
End Sub
Private Sub mnuHelpItem_Click(Index As Integer)
Dim fn_This As String
  Select Case Index
    Case 10:      'ONLINE HELP.
      'NOTE: We currently do NOT have the resources to
      'create an online help file for the program (1/7/98)
      'therefore no online help is available.
      Call Show_Message("Online help is currently unavailable.  " & _
          "Please refer to the printed manual or the Acrobat-format ADS.PDF file.")
      Exit Sub
      'Call LaunchFile_General("", MAIN_APP_PATH & "\help\ads.hlp")
    Case 20:      'ONLINE MANUAL.
      ''''fn_This = MAIN_APP_PATH & "\help\ads.pdf"
      fn_This = MAIN_APP_PATH & "\help\ads.doc"
      If (FileExists(fn_This) = False) Then
        Call Show_Message("The file `" & fn_This & "` is missing.")
        Exit Sub
      End If
      Call ShellExecute_LocalFile(fn_This)
      ''''Call LaunchFile_General("", fn_This)
    Case 22:      'MANUAL PRINTING INSTRUCTIONS.
      fn_This = Global_fpath_dir_CPAS & "\dbase\printing.txt"
      If (FileExists(fn_This) = False) Then
        Call Show_Message("The file `" & fn_This & "` is missing.")
        Exit Sub
      End If
      Call Launch_Notepad(fn_This)
    Case 80:
      fn_This = App.Path & "\dbase\readme.txt"
      If (FileExists(fn_This) = False) Then
        Call Show_Message("The file `" & fn_This & "` is missing.")
        Exit Sub
      End If
      Call Launch_Notepad(fn_This)
    Case 85:    'VIEW DISCLAIMER.
      'SHOW THE DISCLAIMER WINDOW.
      splash_mode = 101
      splash_button_pressed = 0
      frmSplash.Show 1
    Case 90:    'TECH ASSISTANCE.
      frmAbout2.Show 1
    Case 99:    'ABOUT.
      frmAbout.Show 1
  End Select
End Sub
Private Sub mnuMTUItem_Click(Index As Integer)
  Select Case Index
    Case 40:    'KEEP TEMPORARY MODEL FILES.
      mnuMTUItem(40).Checked = Not mnuMTUItem(40).Checked
  End Select
End Sub
Private Sub mnuOptionsItem_Click(Index As Integer)
Dim msg As String
Dim i As Integer
Dim J As Integer
Dim Raise_Dirty_Flag As Boolean
  Select Case Index
    Case 0:       'FOULING.
      ''''frmFouling.Show 1
      Call frmFouling.frmFouling_Go(Raise_Dirty_Flag)
      If (Raise_Dirty_Flag) Then
        'RAISE DIRTY FLAG.
        Call DirtyStatus_Throw
      End If
    
    Case 1:       'INFLUENT CONCENTRATIONS.
      '---- Options--Influent Concentrations
      '-- Setup global variables to make the call
      frmConcentrations_caption = "Influent Concentrations"
      frmConcentrations_Cunits = "mg/L"
      frmConcentrations_Tunits = "days"
      frmConcentrations_TimeOrderImportant = True
      frmConcentrations_NumPoints = Number_Influent_Points
      frmConcentrations_NumConcs = Number_Component
      For i = 1 To frmConcentrations_NumPoints
        frmConcentrations_Times(i) = T_Influent(i)
        For J = 1 To frmConcentrations_NumConcs
          frmConcentrations_Concs(J, i) = C_Influent(J, i)
        Next J
      Next i
      '-- Make the call
      frmVarConcentrations.Show 1
      '-- Reinitialize Last-Few-Files list after frmConcentrations is done
      'xaxaxaNC
      'Call LastFewFiles_InitializeList(LASTFEW_WHICHAPP_ADSIM, LASTFEW_ADSIM_frmPFPSDM)
      ''''Call LastFewFiles_InitializeList(LASTFEW_WHICHAPP_ADXDESIGNS, LASTFEW_ADXDESIGNS_frmPFPSDM)
      If (Not frmConcentrations_cancelled) Then
        Number_Influent_Points = frmConcentrations_NumPoints
        Number_Component = frmConcentrations_NumConcs
        For i = 1 To frmConcentrations_NumPoints
          T_Influent(i) = frmConcentrations_Times(i)
          For J = 1 To frmConcentrations_NumConcs
            C_Influent(J, i) = frmConcentrations_Concs(J, i)
          Next J
        Next i
      End If
    Case 2:       'EFFLUENT CONCENTRATIONS.
      '---- Options--Effluent Concentrations
      '-- Setup global variables to make the call
      frmConcentrations_caption = "Effluent Concentrations"
      frmConcentrations_Cunits = "C/C0"
      frmConcentrations_Tunits = "days"
      frmConcentrations_TimeOrderImportant = False
      frmConcentrations_NumPoints = NData_Points
      frmConcentrations_NumConcs = Number_Component
      For i = 1 To frmConcentrations_NumPoints
        frmConcentrations_Times(i) = T_Data_Points(i) * 24# * 60#
        For J = 1 To frmConcentrations_NumConcs
          frmConcentrations_Concs(J, i) = C_Data_Points(J, i)
        Next J
      Next i
      '-- Make the call
      frmVarConcentrations.Show 1
      '-- Reinitialize Last-Few-Files list after frmConcentrations is done
      'xaxaxaNC
      'Call LastFewFiles_InitializeList(LASTFEW_WHICHAPP_ADSIM, LASTFEW_ADSIM_frmPFPSDM)
      ''''Call LastFewFiles_InitializeList(LASTFEW_WHICHAPP_ADXDESIGNS, LASTFEW_ADXDESIGNS_frmPFPSDM)
      If (Not frmConcentrations_cancelled) Then
        NData_Points = frmConcentrations_NumPoints
        If NData_Points > 0 And mnuResultsItem(1).Enabled = True Then
          mnuResultsItem(4).Enabled = True
        End If
        If NData_Points > 0 And mnuResultsItem(0).Enabled = True Then
          mnuResultsItem(3).Enabled = True
        End If
        Number_Component = frmConcentrations_NumConcs
        For i = 1 To frmConcentrations_NumPoints
          T_Data_Points(i) = frmConcentrations_Times(i) / 24# / 60#
          For J = 1 To frmConcentrations_NumConcs
            C_Data_Points(J, i) = frmConcentrations_Concs(J, i)
          Next J
        Next i
      End If
  End Select
End Sub
Private Sub mnuPhaseItem_Click(Index As Integer)
Dim OldBedPhase As Integer
  OldBedPhase = Bed.Phase
  Select Case Index
    Case 0:     'LIQUID PHASE.
      Call chem_phase(0)
    Case 1:     'GAS PHASE.
      Call chem_phase(1)
  End Select
  If (Bed.Phase <> OldBedPhase) Then
    'THROW DIRTY FLAG AND REFRESH WINDOW.
    Call DirtyStatus_Throw
    Call frmMain_Refresh
  End If
End Sub
Private Sub mnuPrintSubItem_Click(Index As Integer)
  Select Case Index
    Case 0:     'PRINT-TO-PRINTER.
      Print_To_Printer = True
      frmPrintInputs.Show 1
    Case 1:     'PRINT-TO-FILE.
      Print_To_Printer = False
      frmPrintInputs.Show 1
  End Select
End Sub
Private Sub mnuResultsItem_Click(Index As Integer)
  Select Case Index
    Case 0:     'PSDM.
      frmModelPSDMResults.Show 1
    Case 1:     'CPHSDM.
      frmModelCPHSDMResults.Show 1
    Case 2:     'ECM.
      frmModelECMResults.Show 1
    Case 3:
      frmCompareData_WhichSet = frmCompareData_WhichSet_PSDM
      frmCompareData_caption = "Comparison of PSDM Results with Effluent Data"
      frmModelDataComparison.Show 1
    Case 4:
      frmCompareData_WhichSet = frmCompareData_WhichSet_CPHSDM
      frmCompareData_caption = "Comparison of CPHSDM Results with Effluent Data"
      frmModelDataComparison.Show 1
  End Select
End Sub
Private Sub mnuRunItem_Click(Index As Integer)
Dim i As Integer
Dim J As Integer
Dim Num_K_Reduction As Integer
  Select Case Index
    '
    '///////////////////////////////////////////////////////////////////////////////////////////////////
    '/////////    PSDM
    '///////////////////////////////////////////////////////////////////////////////////////////////////
    Case 0:
      If (Prepare_To_Run_PSDM() = False) Then
        Exit Sub
      End If
      'RUN THE MODEL.
      Call ModelPSDM_Go
    '
    '///////////////////////////////////////////////////////////////////////////////////////////////////
    '/////////    PSDMR in Room
    '///////////////////////////////////////////////////////////////////////////////////////////////////
    Case 10:
      If (Prepare_To_Run_PSDM_In_Room() = False) Then
        Exit Sub
      End If
      'RUN THE MODEL.
      Call ModelPSDMInRoom_Go(PSDMR_MODE_INROOM)
    '
    '///////////////////////////////////////////////////////////////////////////////////////////////////
    '/////////    PSDMR Alone
    '///////////////////////////////////////////////////////////////////////////////////////////////////
    Case 20:
      If (Prepare_To_Run_PSDM_In_Room() = False) Then
        Exit Sub
      End If
      'RUN THE MODEL.
      Call ModelPSDMInRoom_Go(PSDMR_MODE_ALONE)
    '
    '///////////////////////////////////////////////////////////////////////////////////////////////////
    '/////////    CPHSDM
    '///////////////////////////////////////////////////////////////////////////////////////////////////
    Case 1:
      '---- Make sure # fouling components is = 0.
      Num_K_Reduction = 0
      For i = 0 To lstComponents.ListCount - 1
        If (lstComponents.Selected(i)) Then
          If (Component(i + 1).K_Reduction) Then
            Num_K_Reduction = Num_K_Reduction + 1
          End If
        End If
      Next i
      If (Num_K_Reduction > 0) Then
        Call Show_Message( _
            "Warning: There are currently " & Trim$(Str$(Num_K_Reduction)) & _
            " components specified with fouling correlations.  The CPHSDM model " & _
            "does not use the fouling correlations and will ignore them.")
      End If
      Call AllModels_Verify_Selected_Components(MODELTYPE_CPHSDM)
      If (Number_Component_CPM = 0) Then
        Exit Sub   'ERROR MESSAGE HANDLED IN AllModels_... SUBROUTINE.
      End If
      'RUN THE MODEL.
      Call ModelCPHSDM_Go
    '
    '///////////////////////////////////////////////////////////////////////////////////////////////////
    '/////////    ECM
    '///////////////////////////////////////////////////////////////////////////////////////////////////
    Case 2:
      '---- Make sure # fouling components is = 0.
      Num_K_Reduction = 0
      For i = 0 To lstComponents.ListCount - 1
        If (lstComponents.Selected(i)) Then
          If (Component(i + 1).K_Reduction) Then
            Num_K_Reduction = Num_K_Reduction + 1
          End If
        End If
      Next i
      If (Num_K_Reduction > 0) Then
        Call Show_Message( _
            "Warning: There are currently " & Trim$(Str$(Num_K_Reduction)) & _
            " components specified with fouling correlations.  The ECM model " & _
            "does not use the fouling correlations and will ignore them.")
      End If
      Call AllModels_Verify_Selected_Components(MODELTYPE_ECM)
      If (Number_Component_ECM = 0) Then
        Exit Sub   'ERROR MESSAGE HANDLED IN AllModels_... SUBROUTINE.
      End If
      For i = 1 To Number_Component_ECM
        For J = i + 1 To Number_Component_ECM
          If (Trim$(Component(Component_Index_ECM(i)).Name) = _
              Trim$(Component(Component_Index_ECM(J)).Name)) Then
            Call Show_Error( _
                "Components " & Format$(Component_Index_ECM(i), "0") & _
                " and " & Format$(Component_Index_ECM(J), "0") & _
                " have the same name." & vbCrLf & _
                "Please change one before running the ECM.")
            Exit Sub
          End If
        Next J
      Next i
      'RUN THE MODEL.
      Call ModelECM_Go
  End Select
End Sub


Private Sub spnNumberOfBeds_SpinDown()
  If (Bed.NumberOfBeds > 1) Then
    Bed.NumberOfBeds = Bed.NumberOfBeds - 1
    'THROW DIRTY FLAG AND REFRESH WINDOW.
    Call DirtyStatus_Throw
    Call frmMain_Refresh
  End If
End Sub
Private Sub spnNumberOfBeds_SpinUp()
  If (Bed.NumberOfBeds < Maximum_Beds_In_Series) Then
    Bed.NumberOfBeds = Bed.NumberOfBeds + 1
    'THROW DIRTY FLAG AND REFRESH WINDOW.
    Call DirtyStatus_Throw
    Call frmMain_Refresh
  End If
End Sub


Private Sub spnPoint_SpinDown(Index As Integer)
  Select Case Index
    Case 0:     'AXIAL POINTS.
      If (MC > 1) Then
        MC = MC - 1
        'THROW DIRTY FLAG AND REFRESH WINDOW.
        Call DirtyStatus_Throw
        Call frmMain_Refresh
      End If
    Case 1:     'RADIAL POINTS.
      If (NC > 1) Then
        NC = NC - 1
        'THROW DIRTY FLAG AND REFRESH WINDOW.
        Call DirtyStatus_Throw
        Call frmMain_Refresh
      End If
  End Select
End Sub
Private Sub spnPoint_SpinUp(Index As Integer)
  Select Case Index
    Case 0:     'AXIAL POINTS.
      If (MC < Max_Axial_Collocation) Then
        MC = MC + 1
        'THROW DIRTY FLAG AND REFRESH WINDOW.
        Call DirtyStatus_Throw
        Call frmMain_Refresh
      End If
    Case 1:     'RADIAL POINTS.
      If (NC < Max_Radial_Collocation) Then
        NC = NC + 1
        'THROW DIRTY FLAG AND REFRESH WINDOW.
        Call DirtyStatus_Throw
        Call frmMain_Refresh
      End If
  End Select
End Sub


Private Sub txtBedUnits_Click(Index As Integer)
Dim Ctl As Control
Set Ctl = txtBedUnits(Index)
  Call unitsys_control_cbox_click(Ctl)
End Sub
Private Sub txtBedUnits_KeyPress(Index As Integer, KeyAscii As Integer)
  KeyAscii = Global_TextKeyPress(KeyAscii)
End Sub


Private Sub txtBedValue_GotFocus(Index As Integer)
Dim Ctl As Control
Set Ctl = txtBedValue(Index)
Dim StatusMessagePanel As String
  Call unitsys_control_txtx_gotfocus(Ctl)
  Select Case Index
    Case 0
      StatusMessagePanel = "Type in the bed length"
    Case 1
      StatusMessagePanel = "Type in the bed diameter"
    Case 2
      StatusMessagePanel = "Type in the mass of adsorbent in the bed"
    Case 3
      StatusMessagePanel = "Type in the inlet flowrate"
    Case 4
      StatusMessagePanel = "Type in the Empty Bed Contact Time (EBCT)"
  End Select
  Call GenericStatus_Set(StatusMessagePanel)
End Sub
Private Sub txtBedValue_KeyPress(Index As Integer, KeyAscii As Integer)
  KeyAscii = Global_NumericKeyPress(KeyAscii)
End Sub
Private Sub txtBedValue_LostFocus(Index As Integer)
Dim NewValue_Okay As Integer
Dim NewValue As Double
Dim Ctl As Control
Set Ctl = txtBedValue(Index)
Dim Val_Low As Double
Dim Val_High As Double
Dim Raise_Dirty_Flag As Boolean
Dim Too_Small As Integer
  'NOTE: LOW AND HIGH VALUES IN BASE UNITS
  If (Index = 4) Then
    Val_Low = 1E-20 * 60#
    Val_High = 1E+20 * 60#
  Else
    Val_Low = 1E-20
    Val_High = 1E+20
  End If
  NewValue_Okay = False
  If (unitsys_control_txtx_lostfocus_validate(Ctl, Val_Low, Val_High, NewValue, Raise_Dirty_Flag)) Then
    NewValue_Okay = True
  End If
  Call unitsys_control_txtx_lostfocus(Ctl, NewValue)
  Call GenericStatus_Set("")
  If (NewValue_Okay) Then
    If (IsThisADemo() = True) And (Raise_Dirty_Flag) Then
      ''''Call Demo_ShowError("Changing data values is not allowed in the demonstration version.")
      Raise_Dirty_Flag = False
    End If
    If (Raise_Dirty_Flag) Then
      'STORE TO MEMORY.
      Select Case Index
        Case 0:     'BED LENGTH.
          Call Check_Length(NewValue, Too_Small)
          Raise_Dirty_Flag = False
          If Not (Too_Small) Then
            Bed.length = NewValue
            Raise_Dirty_Flag = True
          End If
          'Call Update_Display
          Call Update_KP_Values
          'Call Update_Bed_Density_Display
          'Call Update_Several_Bed_Properties(1)
        Case 1:     'BED DIAMETER.
          Call Check_Diameter(NewValue, Too_Small)
          Raise_Dirty_Flag = False
          If Not (Too_Small) Then
            Bed.Diameter = NewValue
            Raise_Dirty_Flag = True
          End If
          'Call Update_Display
          Call Update_KP_Values
          'Call Update_Bed_Density_Display
          'Call Update_Several_Bed_Properties(3)
        Case 2:     'BED MASS.
          Call Check_Weight(NewValue, Too_Small)
          Raise_Dirty_Flag = False
          If Not (Too_Small) Then
            Bed.Weight = NewValue
            Raise_Dirty_Flag = True
          End If
          ''Call Update_Display
          Call Update_KP_Values
          'Call Update_Bed_Density_Display
          'Call Update_Several_Bed_Properties(1)
        Case 3:     'BED FLOW RATE.
          Bed.Flowrate = NewValue
          ''Call Update_Display       'Updates display of flowrate and EBCT.
          Call Update_KP_Values
          'Call Update_Several_Bed_Properties(2)
        Case 4:     'BED EBCT.
          Bed.Flowrate = Bed.length * PI * Bed.Diameter * Bed.Diameter / 4# / NewValue         'EBCT in sec
          ''Call Update_Display       'Updates display of flowrate and EBCT.
          Call Update_KP_Values
          'Call Update_Several_Bed_Properties(2)
      End Select
      If (Raise_Dirty_Flag) Then
        'THROW DIRTY FLAG.
        Call DirtyStatus_Throw
      End If
    End If
    'REFRESH WINDOW.
    Call frmMain_Refresh
  End If
End Sub


Private Sub txtCarbon_GotFocus(Index As Integer)
Dim Ctl As Control
Set Ctl = txtCarbon(Index)
Dim StatusMessagePanel As String
  If (Index = 0) Then
    Call Global_GotFocus(Ctl)
  Else
    Call unitsys_control_txtx_gotfocus(Ctl)
  End If
  Select Case Index
    Case 0
      StatusMessagePanel = "Type in the adsorbent name"
    Case 1
      StatusMessagePanel = "Type in the adsorbent density (that includes pore volume)"
    Case 2
      StatusMessagePanel = "Type in the average particle radius"
    Case 3
      StatusMessagePanel = "Type in the particle porosity"
    Case 4
      StatusMessagePanel = "Type in the particle shape factor (spheres=1.0)"
'    Case 4
'      StatusMessagePanel = " Type in the particle tortuosity"
  End Select
  Call GenericStatus_Set(StatusMessagePanel)
End Sub
Private Sub txtCarbon_KeyPress(Index As Integer, KeyAscii As Integer)
  If (Index = 0) Then
    KeyAscii = Global_TextKeyPress(KeyAscii)
  Else
    KeyAscii = Global_NumericKeyPress(KeyAscii)
  End If
End Sub
Private Sub txtCarbon_LostFocus(Index As Integer)
Dim NewValue_Okay As Integer
Dim NewValue As Double
Dim Ctl As Control
Set Ctl = txtCarbon(Index)
Dim Val_Low As Double
Dim Val_High As Double
Dim Raise_Dirty_Flag As Boolean
Dim Too_Small As Integer
  'HANDLE THE CARBON NAME TEXTBOX.
  If (Index = 0) Then
    If (Trim$(Ctl.Text) = "") Then
      Ctl.Text = Carbon.Name
      'Call Show_Error("You must enter a non-blank string for the carbon name.")
      'NOTE: SHOWING THIS ERROR MESSAGE MESSES UP THE
      'SUBSEQUENT GOTFOCUS IF THE USER HITS <Enter> OR <Tab>.
    Else
      If (Trim$(Carbon.Name) <> Trim$(Ctl.Text)) Then
        Carbon.Name = Trim$(Ctl.Text)
        'THROW DIRTY FLAG.
        Call DirtyStatus_Throw
      End If
    End If
    Call Global_LostFocus(Ctl)
    Call GenericStatus_Set("")
    Exit Sub
  End If
  'NOTE: LOW AND HIGH VALUES IN BASE UNITS
  Val_Low = 1E-20
  Val_High = 1E+20
  NewValue_Okay = False
  If (unitsys_control_txtx_lostfocus_validate(Ctl, Val_Low, Val_High, NewValue, Raise_Dirty_Flag)) Then
    NewValue_Okay = True
  End If
  Call unitsys_control_txtx_lostfocus(Ctl, NewValue)
  Call GenericStatus_Set("")
  If (NewValue_Okay) Then
    If (IsThisADemo() = True) And (Raise_Dirty_Flag) Then
      ''''Call Demo_ShowError("Changing data values is not allowed in the demonstration version.")
      Raise_Dirty_Flag = False
    End If
    If (Raise_Dirty_Flag) Then
      'STORE TO MEMORY.
      Select Case Index
        Case 1:     'APPARENT DENSITY.
          Call Check_Density(NewValue, Too_Small)
          Raise_Dirty_Flag = False
          If Not (Too_Small) Then
            Carbon.Density = NewValue
            Raise_Dirty_Flag = True
          End If
          'Call Update_Display
          Call Update_KP_Values
          'Call Update_Bed_Density_Display
          'Call Update_Several_Bed_Properties(1)
        Case 2:     'PARTICLE RADIUS.
          Carbon.ParticleRadius = NewValue
          'Call Update_Display
          Call Update_KP_Values
        Case 3:     'POROSITY.
          Carbon.Porosity = NewValue
          'Call Update_Display
          Call Update_KP_Values
        Case 4:     'PARTICLE SHAPE FACTOR.
          Carbon.ShapeFactor = NewValue
          'Call Update_Display
          Call Update_KP_Values
      End Select
      If (Raise_Dirty_Flag) Then
        'THROW DIRTY FLAG.
        Call DirtyStatus_Throw
      End If
    End If
    'REFRESH WINDOW.
    Call frmMain_Refresh
  End If
End Sub


Private Sub txtCarbonUnits_Click(Index As Integer)
Dim Ctl As Control
Set Ctl = txtCarbonUnits(Index)
  Call unitsys_control_cbox_click(Ctl)
End Sub
Private Sub txtCarbonUnits_KeyPress(Index As Integer, KeyAscii As Integer)
  KeyAscii = Global_TextKeyPress(KeyAscii)
End Sub


Private Sub txtNPoint_GotFocus(Index As Integer)
Dim Ctl As Control
Set Ctl = txtNPoint(Index)
Dim StatusMessagePanel As String
  Call unitsys_control_txtx_gotfocus(Ctl)
  Select Case Index
    Case 0:
      StatusMessagePanel = "Type in the number of collocation points in the axial direction (PSDM only)"
    Case 1:
      StatusMessagePanel = "Type in the number of collocation points in the radial direction (PSDM only)"
  End Select
  Call GenericStatus_Set(StatusMessagePanel)
End Sub
Private Sub txtNPoint_KeyPress(Index As Integer, KeyAscii As Integer)
  KeyAscii = Global_NumericKeyPress(KeyAscii)
End Sub
Private Sub txtNPoint_LostFocus(Index As Integer)
Dim NewValue_Okay As Integer
Dim NewValue As Double
Dim Ctl As Control
Set Ctl = txtNPoint(Index)
Dim Val_Low As Double
Dim Val_High As Double
Dim Raise_Dirty_Flag As Boolean
Dim Too_Small As Integer
  'NOTE: LOW AND HIGH VALUES IN BASE UNITS
  Select Case Index
    Case 0:   'AXIAL DIR.
      Val_Low = 1#
      Val_High = CDbl(Max_Axial_Collocation)
    Case 1:   'RADIAL DIR.
      Val_Low = 1#
      Val_High = CDbl(Max_Radial_Collocation)
  End Select
  NewValue_Okay = False
  If (unitsys_control_txtx_lostfocus_validate(Ctl, Val_Low, Val_High, NewValue, Raise_Dirty_Flag)) Then
    NewValue_Okay = True
  End If
  Call unitsys_control_txtx_lostfocus(Ctl, NewValue)
  Call GenericStatus_Set("")
  If (NewValue_Okay) Then
    If (IsThisADemo() = True) And (Raise_Dirty_Flag) Then
      ''''Call Demo_ShowError("Changing data values is not allowed in the demonstration version.")
      Raise_Dirty_Flag = False
    End If
    If (Raise_Dirty_Flag) Then
      'STORE TO MEMORY.
      Select Case Index
        Case 0:     'AXIAL DIR.
          MC = CInt(NewValue)
        Case 1:     'RADIAL DIR.
          NC = CInt(NewValue)
      End Select
      If (Raise_Dirty_Flag) Then
        'THROW DIRTY FLAG.
        Call DirtyStatus_Throw
      End If
    End If
    'REFRESH WINDOW.
    Call frmMain_Refresh
  End If
End Sub


Private Sub txtNumberOfBeds_GotFocus()
Dim Ctl As Control
Set Ctl = txtNumberOfBeds
Dim StatusMessagePanel As String
  Call unitsys_control_txtx_gotfocus(Ctl)
  StatusMessagePanel = "Type in the number of axial elements (PSDM only)"
  Call GenericStatus_Set(StatusMessagePanel)
End Sub
Private Sub txtNumberOfBeds_KeyPress(KeyAscii As Integer)
  KeyAscii = Global_NumericKeyPress(KeyAscii)
End Sub
Private Sub txtNumberOfBeds_LostFocus()
Dim NewValue_Okay As Integer
Dim NewValue As Double
Dim Ctl As Control
Set Ctl = txtNumberOfBeds
Dim Val_Low As Double
Dim Val_High As Double
Dim Raise_Dirty_Flag As Boolean
Dim Too_Small As Integer
  'NOTE: LOW AND HIGH VALUES IN BASE UNITS
  Val_Low = 1#
  Val_High = CDbl(Maximum_Beds_In_Series)
  NewValue_Okay = False
  If (unitsys_control_txtx_lostfocus_validate(Ctl, Val_Low, Val_High, NewValue, Raise_Dirty_Flag)) Then
    NewValue_Okay = True
  End If
  Call unitsys_control_txtx_lostfocus(Ctl, NewValue)
  Call GenericStatus_Set("")
  If (NewValue_Okay) Then
    If (IsThisADemo() = True) And (Raise_Dirty_Flag) Then
      ''''Call Demo_ShowError("Changing data values is not allowed in the demonstration version.")
      Raise_Dirty_Flag = False
    End If
    If (Raise_Dirty_Flag) Then
      'STORE TO MEMORY.
      Bed.NumberOfBeds = CInt(NewValue)
      If (Raise_Dirty_Flag) Then
        'THROW DIRTY FLAG.
        Call DirtyStatus_Throw
      End If
    End If
    'REFRESH WINDOW.
    Call frmMain_Refresh
  End If
End Sub


Private Sub txtTime_GotFocus(Index As Integer)
Dim Ctl As Control
Set Ctl = txtTime(Index)
Dim StatusMessagePanel As String
  Call unitsys_control_txtx_gotfocus(Ctl)
  Select Case Index
    Case 0
      StatusMessagePanel = "Type in the total run time of the fixed bed adsorber (PSDM only)"
    Case 1
      StatusMessagePanel = "Type in the time of the first point to be displayed (PSDM only)"
    Case 2
      StatusMessagePanel = "Type in the time step (PSDM only)"
  End Select
  Call GenericStatus_Set(StatusMessagePanel)
End Sub
Private Sub txtTime_KeyPress(Index As Integer, KeyAscii As Integer)
  KeyAscii = Global_NumericKeyPress(KeyAscii)
End Sub
Private Sub txtTime_LostFocus(Index As Integer)
Dim NewValue_Okay As Integer
Dim NewValue As Double
Dim Ctl As Control
Set Ctl = txtTime(Index)
Dim Val_Low As Double
Dim Val_High As Double
Dim Raise_Dirty_Flag As Boolean
Dim Too_Small As Integer
Dim ForceAbort As Integer
  'NOTE: LOW AND HIGH VALUES IN BASE UNITS
  Select Case Index
    Case 0, 2:      'TOTAL RUN TIME, TIME STEP.
      Val_Low = 1E-20
      Val_High = 1E+20
    Case 1:         'FIRST POINT DISPLAYED.
      Val_Low = 0#
      Val_High = 1E+20
  End Select
  NewValue_Okay = False
  If (unitsys_control_txtx_lostfocus_validate(Ctl, Val_Low, Val_High, NewValue, Raise_Dirty_Flag)) Then
    NewValue_Okay = True
  End If
  Call unitsys_control_txtx_lostfocus(Ctl, NewValue)
  Call GenericStatus_Set("")
  If (NewValue_Okay) Then
    If (IsThisADemo() = True) And (Raise_Dirty_Flag) Then
      ''''Call Demo_ShowError("Changing data values is not allowed in the demonstration version.")
      Raise_Dirty_Flag = False
    End If
    If (Raise_Dirty_Flag) Then
      'STORE TO MEMORY.
      Select Case Index
        Case 0:       'TOTAL RUN TIME.
          Call Check_Time_Parameters(0, NewValue, ForceAbort)
          If Not (ForceAbort) Then
            TimeP.End = NewValue
          End If
        Case 1:       'FIRST POINT DISPLAYED.
          Call Check_Time_Parameters(1, NewValue, ForceAbort)
          If Not (ForceAbort) Then
            TimeP.Init = NewValue
          End If
        Case 2:       'TIME STEP.
          Call Check_Time_Parameters(2, NewValue, ForceAbort)
          If Not (ForceAbort) Then
            TimeP.Step = NewValue
          End If
      End Select
      If (Raise_Dirty_Flag) Then
        'THROW DIRTY FLAG.
        Call DirtyStatus_Throw
      End If
    End If
    'REFRESH WINDOW.
    Call frmMain_Refresh
  End If
End Sub


Private Sub txttimeunits_Click(Index As Integer)
Dim Ctl As Control
Set Ctl = txtTimeUnits(Index)
  Call unitsys_control_cbox_click(Ctl)
End Sub
Private Sub txttimeunits_KeyPress(Index As Integer, KeyAscii As Integer)
  KeyAscii = Global_TextKeyPress(KeyAscii)
End Sub


Private Sub txtWater_GotFocus(Index As Integer)
Dim Ctl As Control
Set Ctl = txtTime(Index)
Dim StatusMessagePanel As String
  Call unitsys_control_txtx_gotfocus(txtWater(Index))
  Select Case Index
    Case 0:
      StatusMessagePanel = "Type in the Fluid Temperature"
    Case 1:
      StatusMessagePanel = "Type in the Fluid Pressure"
  End Select
  Call GenericStatus_Set(StatusMessagePanel)
End Sub
Private Sub txtWater_KeyPress(Index As Integer, KeyAscii As Integer)
  KeyAscii = Global_NumericKeyPress(KeyAscii)
End Sub
Private Sub txtWater_LostFocus(Index As Integer)
Dim NewValue_Okay As Integer
Dim NewValue As Double
Dim Ctl As Control
Set Ctl = txtWater(Index)
Dim Val_Low As Double
Dim Val_High As Double
Dim Raise_Dirty_Flag As Boolean
Dim Too_Small As Integer
Dim ForceAbort As Integer
  'NOTE: LOW AND HIGH VALUES IN BASE UNITS
  Select Case Index
    Case 0:         'TEMPERATURE (degC).
      Val_Low = 0.01
      Val_High = 100#
    Case 1:         'PRESSURE (atm).
      Val_Low = 0.001
      Val_High = 100#
  End Select
  NewValue_Okay = False
  If (unitsys_control_txtx_lostfocus_validate(Ctl, Val_Low, Val_High, NewValue, Raise_Dirty_Flag)) Then
    NewValue_Okay = True
  End If
  Call unitsys_control_txtx_lostfocus(Ctl, NewValue)
  Call GenericStatus_Set("")
  If (NewValue_Okay) Then
    If (IsThisADemo() = True) And (Raise_Dirty_Flag) Then
      ''''Call Demo_ShowError("Changing data values is not allowed in the demonstration version.")
      Raise_Dirty_Flag = False
    End If
    If (Raise_Dirty_Flag) Then
      'STORE TO MEMORY.
      Select Case Index
        Case 0:       'TEMPERATURE (degC).
          Bed.Temperature = NewValue
          'Call Update_Display_Water
        Case 1:       'PRESSURE (atm).
          Bed.Pressure = NewValue
          'Call Update_Display_Water
      End Select
      Call Update_FluidDensity(Bed.Temperature, Bed.Pressure, Bed.WaterDensity)
      Call Update_FluidViscosity(Bed.Temperature, Bed.WaterViscosity)
      If (Raise_Dirty_Flag) Then
        'THROW DIRTY FLAG.
        Call DirtyStatus_Throw
      End If
    End If
    'REFRESH WINDOW.
    Call frmMain_Refresh
  End If
End Sub






 


 
 
 

