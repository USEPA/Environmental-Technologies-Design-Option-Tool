VERSION 5.00
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "THREED32.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form frmIonExchangeMain 
   Appearance      =   0  'Flat
   BackColor       =   &H00C0C0C0&
   Caption         =   "Ion Exchange Design Software"
   ClientHeight    =   6630
   ClientLeft      =   2085
   ClientTop       =   2310
   ClientWidth     =   9510
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
   Icon            =   "6_Ionexmn.frx":0000
   LinkTopic       =   "Form1"
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   6630
   ScaleWidth      =   9510
   Begin Threed.SSFrame fraAbsorbentProperties 
      Height          =   2640
      Left            =   120
      TabIndex        =   7
      Top             =   3375
      Width           =   4695
      _Version        =   65536
      _ExtentX        =   8281
      _ExtentY        =   4657
      _StockProps     =   14
      Caption         =   "Resin Properties"
      ForeColor       =   -2147483646
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial Black"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Begin VB.CommandButton cmdAdsorbentDatabase 
         Appearance      =   0  'Flat
         Caption         =   "Resin Database"
         Enabled         =   0   'False
         Height          =   315
         Left            =   2400
         TabIndex        =   17
         Top             =   170
         Width           =   1575
      End
      Begin VB.TextBox txtAdsorbentProperties 
         Appearance      =   0  'Flat
         Height          =   285
         Index           =   1
         Left            =   1740
         TabIndex        =   16
         Top             =   900
         Width           =   1095
      End
      Begin VB.TextBox txtAdsorbentProperties 
         Appearance      =   0  'Flat
         Height          =   285
         Index           =   2
         Left            =   1740
         TabIndex        =   15
         Top             =   1230
         Width           =   1095
      End
      Begin VB.TextBox txtAdsorbentProperties 
         Appearance      =   0  'Flat
         Height          =   285
         Index           =   3
         Left            =   1740
         TabIndex        =   14
         Top             =   1560
         Width           =   1095
      End
      Begin VB.TextBox txtAdsorbentProperties 
         Appearance      =   0  'Flat
         Height          =   285
         Index           =   4
         Left            =   1740
         TabIndex        =   13
         Top             =   1890
         Width           =   1092
      End
      Begin VB.TextBox txtAdsorbentProperties 
         Appearance      =   0  'Flat
         Height          =   285
         Index           =   5
         Left            =   1740
         TabIndex        =   12
         Top             =   2220
         Width           =   1092
      End
      Begin VB.ComboBox cboAdsorbentPropertyUnits 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         Height          =   315
         Index           =   1
         Left            =   2940
         Style           =   2  'Dropdown List
         TabIndex        =   11
         TabStop         =   0   'False
         Top             =   900
         Width           =   1515
      End
      Begin VB.ComboBox cboAdsorbentPropertyUnits 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         Height          =   315
         Index           =   2
         Left            =   2940
         Style           =   2  'Dropdown List
         TabIndex        =   10
         TabStop         =   0   'False
         Top             =   1230
         Width           =   1515
      End
      Begin VB.ComboBox cboAdsorbentPropertyUnits 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         Height          =   315
         Index           =   5
         Left            =   2940
         Style           =   2  'Dropdown List
         TabIndex        =   9
         TabStop         =   0   'False
         Top             =   2220
         Width           =   1515
      End
      Begin VB.ComboBox cboAdsorbents 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   1740
         Style           =   2  'Dropdown List
         TabIndex        =   8
         Top             =   540
         Width           =   2712
      End
      Begin VB.Label lblAdsorbentProperties 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Name"
         ForeColor       =   &H80000008&
         Height          =   195
         Index           =   0
         Left            =   120
         TabIndex        =   25
         Top             =   600
         Width           =   1545
      End
      Begin VB.Label lblAdsorbentProperties 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Apparent Density"
         ForeColor       =   &H80000008&
         Height          =   195
         Index           =   1
         Left            =   120
         TabIndex        =   24
         Top             =   920
         Width           =   1545
      End
      Begin VB.Label lblAdsorbentProperties 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Particle Radius"
         ForeColor       =   &H80000008&
         Height          =   195
         Index           =   2
         Left            =   120
         TabIndex        =   23
         Top             =   1280
         Width           =   1545
      End
      Begin VB.Label lblAdsorbentProperties 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Particle Porosity"
         ForeColor       =   &H80000008&
         Height          =   195
         Index           =   3
         Left            =   120
         TabIndex        =   22
         Top             =   1610
         Width           =   1545
      End
      Begin VB.Label lblAdsorbentProperties 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Tortuosity"
         ForeColor       =   &H80000008&
         Height          =   195
         Index           =   4
         Left            =   120
         TabIndex        =   21
         Top             =   1940
         Width           =   1545
      End
      Begin VB.Label lblAdsorbentProperties 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Total Capacity"
         ForeColor       =   &H80000008&
         Height          =   195
         Index           =   5
         Left            =   120
         TabIndex        =   20
         Top             =   2280
         Width           =   1545
      End
      Begin VB.Label lblAdsorbentProperties 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "(-)"
         ForeColor       =   &H80000008&
         Height          =   195
         Index           =   6
         Left            =   2940
         TabIndex        =   19
         Top             =   1560
         Width           =   1095
      End
      Begin VB.Label lblAdsorbentProperties 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "(-)"
         ForeColor       =   &H80000008&
         Height          =   195
         Index           =   7
         Left            =   2940
         TabIndex        =   18
         Top             =   1890
         Width           =   1095
      End
   End
   Begin Threed.SSFrame fraIonInSystem 
      Height          =   2535
      Left            =   120
      TabIndex        =   75
      Top             =   950
      Width           =   4695
      _Version        =   65536
      _ExtentX        =   8281
      _ExtentY        =   4471
      _StockProps     =   14
      Caption         =   "Ions in System"
      ForeColor       =   -2147483646
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial Black"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Begin VB.ComboBox cboIons 
         Appearance      =   0  'Flat
         Height          =   315
         Index           =   2
         Left            =   3240
         Style           =   2  'Dropdown List
         TabIndex        =   94
         Top             =   1320
         Width           =   1335
      End
      Begin VB.CommandButton cmdEditProperties 
         Appearance      =   0  'Flat
         Caption         =   "Edit Properties"
         Height          =   252
         Left            =   3185
         TabIndex        =   93
         Top             =   1740
         Width           =   1425
      End
      Begin VB.CommandButton cmdAddDeleteIons 
         Appearance      =   0  'Flat
         Caption         =   "Remove Ion"
         Height          =   252
         Index           =   1
         Left            =   3185
         TabIndex        =   92
         Top             =   2040
         Width           =   1425
      End
      Begin VB.ListBox lstIons 
         Appearance      =   0  'Flat
         Height          =   1200
         Index           =   0
         Left            =   180
         TabIndex        =   81
         Top             =   600
         Width           =   1275
      End
      Begin VB.ListBox lstIons 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         Height          =   1200
         Index           =   1
         Left            =   1740
         MultiSelect     =   1  'Simple
         TabIndex        =   80
         Top             =   600
         Width           =   1275
      End
      Begin VB.CommandButton cmdAddDeleteIons 
         Appearance      =   0  'Flat
         Caption         =   "Add Cation"
         Height          =   252
         Index           =   0
         Left            =   3240
         TabIndex        =   79
         Top             =   240
         Width           =   1275
      End
      Begin VB.CommandButton cmdAddDeleteIons 
         Appearance      =   0  'Flat
         Caption         =   "Add Anion"
         Height          =   252
         Index           =   2
         Left            =   3240
         TabIndex        =   78
         Top             =   600
         Width           =   1275
      End
      Begin VB.ComboBox cboIons 
         Appearance      =   0  'Flat
         Height          =   315
         Index           =   0
         Left            =   195
         Style           =   2  'Dropdown List
         TabIndex        =   77
         Top             =   2040
         Width           =   1272
      End
      Begin VB.ComboBox cboIons 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         Height          =   315
         Index           =   1
         Left            =   1740
         Style           =   2  'Dropdown List
         TabIndex        =   76
         Top             =   2040
         Width           =   1272
      End
      Begin VB.Shape Shape3 
         Height          =   1425
         Left            =   3120
         Top             =   960
         Width           =   1515
      End
      Begin VB.Label lblKineticDimensionless 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Select Ion:"
         ForeColor       =   &H80000008&
         Height          =   195
         Index           =   10
         Left            =   3360
         TabIndex        =   95
         Top             =   1080
         Width           =   1035
      End
      Begin VB.Label lblCations 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Cations"
         ForeColor       =   &H80000008&
         Height          =   195
         Index           =   0
         Left            =   120
         TabIndex        =   85
         Top             =   240
         Width           =   1215
      End
      Begin VB.Label lblAnions 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Anions"
         Enabled         =   0   'False
         ForeColor       =   &H80000008&
         Height          =   195
         Index           =   0
         Left            =   1680
         TabIndex        =   84
         Top             =   240
         Width           =   1215
      End
      Begin VB.Shape Shape1 
         Height          =   1905
         Left            =   120
         Top             =   480
         Width           =   1395
      End
      Begin VB.Shape Shape2 
         Height          =   1905
         Left            =   1680
         Top             =   480
         Width           =   1395
      End
      Begin VB.Label lblCations 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Presaturant"
         ForeColor       =   &H80000008&
         Height          =   195
         Index           =   1
         Left            =   240
         TabIndex        =   83
         Top             =   1800
         Width           =   1275
      End
      Begin VB.Label lblAnions 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Presaturant"
         ForeColor       =   &H80000008&
         Height          =   195
         Index           =   1
         Left            =   1680
         TabIndex        =   82
         Top             =   1800
         Width           =   1275
      End
   End
   Begin MSComDlg.CommonDialog CMDialog1 
      Left            =   7920
      Top             =   6240
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin Threed.SSFrame fraDimensionlessGroups 
      Height          =   1815
      Left            =   7200
      TabIndex        =   44
      Top             =   3720
      Width           =   2175
      _Version        =   65536
      _ExtentX        =   3836
      _ExtentY        =   3201
      _StockProps     =   14
      Caption         =   "Dimensionless Groups"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Alignment       =   2
      Begin VB.Label lblKineticDimensionless 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Dgt"
         ForeColor       =   &H80000008&
         Height          =   195
         Index           =   4
         Left            =   120
         TabIndex        =   73
         Top             =   720
         Width           =   435
      End
      Begin VB.Label lblKineticDimensionless 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Edp"
         ForeColor       =   &H80000008&
         Height          =   195
         Index           =   5
         Left            =   120
         TabIndex        =   72
         Top             =   960
         Width           =   435
      End
      Begin VB.Label lblKineticDimensionless 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "St"
         ForeColor       =   &H80000008&
         Height          =   195
         Index           =   6
         Left            =   120
         TabIndex        =   71
         Top             =   1200
         Width           =   435
      End
      Begin VB.Label lblKineticDimensionlessValue 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   195
         Index           =   5
         Left            =   660
         TabIndex        =   70
         Top             =   720
         Width           =   1095
      End
      Begin VB.Label lblKineticDimensionlessValue 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   195
         Index           =   6
         Left            =   660
         TabIndex        =   69
         Top             =   960
         Width           =   1095
      End
      Begin VB.Label lblKineticDimensionlessValue 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   195
         Index           =   7
         Left            =   660
         TabIndex        =   68
         Top             =   1200
         Width           =   1095
      End
      Begin VB.Label lblKineticDimensionlessUnits 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "(-)"
         ForeColor       =   &H80000008&
         Height          =   195
         Index           =   5
         Left            =   1860
         TabIndex        =   67
         Top             =   720
         Width           =   255
      End
      Begin VB.Label lblKineticDimensionlessUnits 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "(-)"
         ForeColor       =   &H80000008&
         Height          =   195
         Index           =   6
         Left            =   1860
         TabIndex        =   66
         Top             =   960
         Width           =   255
      End
      Begin VB.Label lblKineticDimensionlessUnits 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "(-)"
         ForeColor       =   &H80000008&
         Height          =   195
         Index           =   7
         Left            =   1860
         TabIndex        =   65
         Top             =   1200
         Width           =   255
      End
      Begin VB.Label lblKineticDimensionlessUnits 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "(-)"
         ForeColor       =   &H80000008&
         Height          =   195
         Index           =   3
         Left            =   1860
         TabIndex        =   64
         Top             =   240
         Width           =   255
      End
      Begin VB.Label lblKineticDimensionlessValue 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   195
         Index           =   3
         Left            =   660
         TabIndex        =   63
         Top             =   240
         Width           =   1095
      End
      Begin VB.Label lblKineticDimensionless 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Dgs"
         ForeColor       =   &H80000008&
         Height          =   195
         Index           =   7
         Left            =   120
         TabIndex        =   62
         Top             =   240
         Width           =   435
      End
      Begin VB.Label lblKineticDimensionlessUnits 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "(-)"
         ForeColor       =   &H80000008&
         Height          =   195
         Index           =   4
         Left            =   1860
         TabIndex        =   61
         Top             =   480
         Width           =   255
      End
      Begin VB.Label lblKineticDimensionlessValue 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   195
         Index           =   4
         Left            =   660
         TabIndex        =   60
         Top             =   480
         Width           =   1095
      End
      Begin VB.Label lblKineticDimensionless 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Dgp"
         ForeColor       =   &H80000008&
         Height          =   195
         Index           =   8
         Left            =   120
         TabIndex        =   59
         Top             =   480
         Width           =   435
      End
      Begin VB.Label lblKineticDimensionlessUnits 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "(-)"
         ForeColor       =   &H80000008&
         Height          =   195
         Index           =   8
         Left            =   1860
         TabIndex        =   58
         Top             =   1440
         Width           =   255
      End
      Begin VB.Label lblKineticDimensionlessValue 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   195
         Index           =   8
         Left            =   660
         TabIndex        =   57
         Top             =   1440
         Width           =   1095
      End
      Begin VB.Label lblKineticDimensionless 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Bip"
         ForeColor       =   &H80000008&
         Height          =   195
         Index           =   9
         Left            =   120
         TabIndex        =   56
         Top             =   1440
         Width           =   435
      End
   End
   Begin Threed.SSFrame fraKinetic 
      Height          =   1815
      Left            =   4920
      TabIndex        =   43
      Top             =   3720
      Width           =   2295
      _Version        =   65536
      _ExtentX        =   4048
      _ExtentY        =   3201
      _StockProps     =   14
      Caption         =   "Kinetic Parameters"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Alignment       =   2
      Begin VB.Label lblKineticDimensionless 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "kf"
         ForeColor       =   &H80000008&
         Height          =   195
         Index           =   1
         Left            =   120
         TabIndex        =   55
         Top             =   360
         Width           =   315
      End
      Begin VB.Label lblKineticDimensionless 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Dl"
         ForeColor       =   &H80000008&
         Height          =   195
         Index           =   2
         Left            =   120
         TabIndex        =   54
         Top             =   720
         Width           =   315
      End
      Begin VB.Label lblKineticDimensionless 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Dp"
         ForeColor       =   &H80000008&
         Height          =   195
         Index           =   3
         Left            =   120
         TabIndex        =   53
         Top             =   1080
         Width           =   315
      End
      Begin VB.Label lblKineticDimensionlessValue 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   195
         Index           =   0
         Left            =   540
         TabIndex        =   52
         Top             =   360
         Width           =   1095
      End
      Begin VB.Label lblKineticDimensionlessValue 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   195
         Index           =   1
         Left            =   540
         TabIndex        =   51
         Top             =   720
         Width           =   1095
      End
      Begin VB.Label lblKineticDimensionlessValue 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   195
         Index           =   2
         Left            =   540
         TabIndex        =   50
         Top             =   1080
         Width           =   1095
      End
      Begin VB.Label lblKineticDimensionlessUnits 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "cm/s"
         ForeColor       =   &H80000008&
         Height          =   195
         Index           =   0
         Left            =   1680
         TabIndex        =   49
         Top             =   360
         Width           =   615
      End
      Begin VB.Label lblKineticDimensionlessUnits 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "cm2/s"
         ForeColor       =   &H80000008&
         Height          =   195
         Index           =   1
         Left            =   1680
         TabIndex        =   48
         Top             =   720
         Width           =   615
      End
      Begin VB.Label lblKineticDimensionlessUnits 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "cm2/s"
         ForeColor       =   &H80000008&
         Height          =   195
         Index           =   2
         Left            =   1680
         TabIndex        =   47
         Top             =   1080
         Width           =   615
      End
   End
   Begin Threed.SSFrame fraKineticDimensionless 
      Height          =   3495
      Left            =   4920
      TabIndex        =   42
      Top             =   2520
      Width           =   4455
      _Version        =   65536
      _ExtentX        =   7858
      _ExtentY        =   6165
      _StockProps     =   14
      Caption         =   "Kinetic Parameters and D'less Groups"
      ForeColor       =   -2147483646
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial Black"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Begin VB.CommandButton cmdInputKineticParameters 
         Appearance      =   0  'Flat
         Caption         =   "View Kinetic Parameters"
         Height          =   312
         Left            =   840
         TabIndex        =   74
         Top             =   3075
         Width           =   2655
      End
      Begin VB.ComboBox cboKinDimComponent 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   2280
         Style           =   2  'Dropdown List
         TabIndex        =   45
         Top             =   680
         Width           =   1695
      End
      Begin VB.Label lblKineticDimensionless 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Select Component:"
         ForeColor       =   &H80000008&
         Height          =   195
         Index           =   0
         Left            =   360
         TabIndex        =   46
         Top             =   720
         Width           =   1635
      End
   End
   Begin Threed.SSFrame fraBedData 
      Height          =   2535
      Left            =   4920
      TabIndex        =   26
      Top             =   0
      Width           =   4455
      _Version        =   65536
      _ExtentX        =   7858
      _ExtentY        =   4471
      _StockProps     =   14
      Caption         =   "Bed Data"
      ForeColor       =   -2147483646
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial Black"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Begin VB.TextBox txtBedData 
         Appearance      =   0  'Flat
         Height          =   288
         Index           =   0
         Left            =   2040
         TabIndex        =   36
         Top             =   360
         Width           =   1092
      End
      Begin VB.TextBox txtBedData 
         Appearance      =   0  'Flat
         Height          =   288
         Index           =   1
         Left            =   2040
         TabIndex        =   35
         Top             =   720
         Width           =   1092
      End
      Begin VB.TextBox txtBedData 
         Appearance      =   0  'Flat
         Height          =   288
         Index           =   2
         Left            =   2040
         TabIndex        =   34
         Top             =   1080
         Width           =   1092
      End
      Begin VB.TextBox txtBedData 
         Appearance      =   0  'Flat
         Height          =   288
         Index           =   3
         Left            =   2040
         TabIndex        =   33
         Top             =   1440
         Width           =   1092
      End
      Begin VB.TextBox txtBedData 
         Appearance      =   0  'Flat
         Height          =   288
         Index           =   4
         Left            =   2040
         TabIndex        =   32
         Top             =   1800
         Width           =   1092
      End
      Begin VB.ComboBox cboBedDataUnits 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         Height          =   315
         Index           =   0
         Left            =   3240
         Style           =   2  'Dropdown List
         TabIndex        =   31
         TabStop         =   0   'False
         Top             =   360
         Width           =   1095
      End
      Begin VB.ComboBox cboBedDataUnits 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         Height          =   315
         Index           =   1
         Left            =   3240
         Style           =   2  'Dropdown List
         TabIndex        =   30
         TabStop         =   0   'False
         Top             =   720
         Width           =   1095
      End
      Begin VB.ComboBox cboBedDataUnits 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         Height          =   315
         Index           =   2
         Left            =   3240
         Style           =   2  'Dropdown List
         TabIndex        =   29
         TabStop         =   0   'False
         Top             =   1080
         Width           =   1095
      End
      Begin VB.ComboBox cboBedDataUnits 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         Height          =   315
         Index           =   3
         Left            =   3240
         Style           =   2  'Dropdown List
         TabIndex        =   28
         TabStop         =   0   'False
         Top             =   1440
         Width           =   1095
      End
      Begin VB.ComboBox cboBedDataUnits 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         Height          =   315
         Index           =   4
         Left            =   3240
         Style           =   2  'Dropdown List
         TabIndex        =   27
         TabStop         =   0   'False
         Top             =   1800
         Width           =   1095
      End
      Begin VB.Label lblBedData 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Bed Length"
         ForeColor       =   &H80000008&
         Height          =   195
         Index           =   0
         Left            =   120
         TabIndex        =   41
         Top             =   420
         Width           =   1800
      End
      Begin VB.Label lblBedData 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Adsorber Diameter"
         ForeColor       =   &H80000008&
         Height          =   195
         Index           =   1
         Left            =   120
         TabIndex        =   40
         Top             =   780
         Width           =   1800
      End
      Begin VB.Label lblBedData 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Bed Mass"
         ForeColor       =   &H80000008&
         Height          =   195
         Index           =   2
         Left            =   120
         TabIndex        =   39
         Top             =   1140
         Width           =   1800
      End
      Begin VB.Label lblBedData 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Flow Rate"
         ForeColor       =   &H80000008&
         Height          =   195
         Index           =   3
         Left            =   120
         TabIndex        =   38
         Top             =   1500
         Width           =   1800
      End
      Begin VB.Label lblBedData 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "EBCT"
         ForeColor       =   &H80000008&
         Height          =   195
         Index           =   4
         Left            =   840
         TabIndex        =   37
         Top             =   1845
         Width           =   1095
      End
   End
   Begin Threed.SSFrame fraOperatingConditions 
      Height          =   975
      Left            =   120
      TabIndex        =   0
      Top             =   0
      Width           =   4695
      _Version        =   65536
      _ExtentX        =   8281
      _ExtentY        =   1720
      _StockProps     =   14
      Caption         =   "Operating Conditions"
      ForeColor       =   -2147483646
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial Black"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Begin VB.TextBox txtOperatingConditions 
         Appearance      =   0  'Flat
         Height          =   285
         Index           =   0
         Left            =   1920
         TabIndex        =   4
         Top             =   360
         Width           =   1095
      End
      Begin VB.TextBox txtOperatingConditions 
         Appearance      =   0  'Flat
         Height          =   285
         Index           =   1
         Left            =   1920
         TabIndex        =   3
         Top             =   675
         Width           =   1095
      End
      Begin VB.ComboBox cboOperatingConditionsUnits 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         Height          =   315
         Index           =   0
         Left            =   3120
         Style           =   2  'Dropdown List
         TabIndex        =   2
         TabStop         =   0   'False
         Top             =   360
         Width           =   1155
      End
      Begin VB.ComboBox cboOperatingConditionsUnits 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         Height          =   315
         Index           =   1
         Left            =   3120
         Style           =   2  'Dropdown List
         TabIndex        =   1
         TabStop         =   0   'False
         Top             =   675
         Width           =   1155
      End
      Begin VB.Label lblOperatingConditions 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Temperature"
         ForeColor       =   &H80000008&
         Height          =   195
         Index           =   1
         Left            =   240
         TabIndex        =   6
         Top             =   675
         Width           =   1530
      End
      Begin VB.Label lblOperatingConditions 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Pressure"
         ForeColor       =   &H80000008&
         Height          =   195
         Index           =   0
         Left            =   600
         TabIndex        =   5
         Top             =   420
         Width           =   1170
      End
   End
   Begin Threed.SSFrame SSFrame1 
      Height          =   1005
      Left            =   7920
      TabIndex        =   86
      Top             =   180
      Visible         =   0   'False
      Width           =   2985
      _Version        =   65536
      _ExtentX        =   5265
      _ExtentY        =   1773
      _StockProps     =   14
      Caption         =   "Invisible"
      ForeColor       =   -2147483630
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Begin VB.PictureBox Picture1_nolongerused 
         Height          =   345
         Left            =   450
         ScaleHeight     =   285
         ScaleWidth      =   14805
         TabIndex        =   87
         Top             =   540
         Width           =   14865
      End
      Begin MSComDlg.CommonDialog CommonDialog1 
         Left            =   120
         Top             =   330
         _ExtentX        =   847
         _ExtentY        =   847
         _Version        =   393216
      End
      Begin Threed.SSPanel panDirty_to_be_deleted 
         Height          =   285
         Left            =   900
         TabIndex        =   88
         Top             =   270
         Width           =   1845
         _Version        =   65536
         _ExtentX        =   3254
         _ExtentY        =   503
         _StockProps     =   15
         Caption         =   "Data Unchanged"
         ForeColor       =   -2147483630
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
   End
   Begin Threed.SSPanel SSPanel2 
      Align           =   2  'Align Bottom
      Height          =   405
      Left            =   0
      TabIndex        =   89
      Top             =   6225
      Width           =   9510
      _Version        =   65536
      _ExtentX        =   16775
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
         TabIndex        =   90
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
         TabIndex        =   91
         Top             =   60
         Width           =   5000
         _Version        =   65536
         _ExtentX        =   8819
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
   Begin VB.Menu mnuFileMenu 
      Caption         =   "&File"
      Begin VB.Menu mnuFile 
         Caption         =   "&New"
         Index           =   0
      End
      Begin VB.Menu mnuFile 
         Caption         =   "&Open"
         Index           =   1
      End
      Begin VB.Menu mnuFile 
         Caption         =   "&Save"
         Index           =   2
      End
      Begin VB.Menu mnuFile 
         Caption         =   "Save &As"
         Index           =   3
      End
      Begin VB.Menu mnuFile 
         Caption         =   "-"
         Index           =   4
      End
      Begin VB.Menu mnuFile 
         Caption         =   "Select P&rinter"
         Index           =   5
      End
      Begin VB.Menu mnuFile 
         Caption         =   "&Print"
         Index           =   6
         Begin VB.Menu mnuFilePrint 
            Caption         =   "To &Printer"
            Index           =   0
         End
         Begin VB.Menu mnuFilePrint 
            Caption         =   "To &File"
            Index           =   1
         End
      End
      Begin VB.Menu mnuFile 
         Caption         =   "-"
         Index           =   7
      End
      Begin VB.Menu mnuFile 
         Caption         =   "&1 Old File #1"
         Index           =   8
      End
      Begin VB.Menu mnuFile 
         Caption         =   "&2 Old File #2"
         Index           =   9
      End
      Begin VB.Menu mnuFile 
         Caption         =   "&3 Old File #3"
         Index           =   10
      End
      Begin VB.Menu mnuFile 
         Caption         =   "&4 Old File #4"
         Index           =   11
      End
      Begin VB.Menu mnuFile 
         Caption         =   "-"
         Index           =   12
      End
      Begin VB.Menu mnuFile 
         Caption         =   "E&xit"
         HelpContextID   =   200
         Index           =   13
      End
   End
   Begin VB.Menu mnuRunMenu 
      Caption         =   "&Run"
      Begin VB.Menu mnuRun 
         Caption         =   "&PFPDM"
         Index           =   0
      End
   End
   Begin VB.Menu mnuResultsMenu 
      Caption         =   "Re&sults"
      Begin VB.Menu mnuResults 
         Caption         =   "&PFPDM"
         Index           =   0
      End
      Begin VB.Menu mnuResults 
         Caption         =   "Compare to &Data"
         Index           =   1
      End
   End
   Begin VB.Menu mnuOptionsMenu 
      Caption         =   "&Options"
      Begin VB.Menu mnuOptions 
         Caption         =   "&Influent Concentrations"
         Index           =   0
      End
      Begin VB.Menu mnuOptions 
         Caption         =   "Set &Number of Beds"
         Index           =   1
      End
      Begin VB.Menu mnuOptions 
         Caption         =   "Set &Time Parameters"
         Index           =   2
      End
      Begin VB.Menu mnuOptions 
         Caption         =   "Set &Collocation Points"
         Index           =   3
      End
      Begin VB.Menu mnuOptions 
         Caption         =   "Set &Resin Phase Presaturant Conditions"
         Enabled         =   0   'False
         Index           =   4
      End
      Begin VB.Menu mnuOptions 
         Caption         =   "Set &EPS and DH0"
         Index           =   5
      End
   End
End
Attribute VB_Name = "frmIonExchangeMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

Dim Temp_Text As String
Dim IsError As Integer

Private Sub cboAdsorbentPropertyUnits_Click(Index As Integer)
    Dim ValueToDisplay As Double

    Select Case Index
       Case 1   'Apparent Density
            Select Case cboAdsorbentPropertyUnits(1).ListIndex
               Case APPARENT_DENSITY_G_per_ML    'g/ml
                    ValueToDisplay = NowProj.Resin.ApparentDensity
               Case APPARENT_DENSITY_KG_per_M3    'kg/m3
                    ValueToDisplay = NowProj.Resin.ApparentDensity * DensityConversionFactor(APPARENT_DENSITY_KG_per_M3)
               Case APPARENT_DENSITY_LB_per_FT3    'lb/ft3
                    ValueToDisplay = NowProj.Resin.ApparentDensity * DensityConversionFactor(APPARENT_DENSITY_LB_per_FT3)
               Case APPARENT_DENSITY_LB_per_GAL    'lb/gal
                    ValueToDisplay = NowProj.Resin.ApparentDensity * DensityConversionFactor(APPARENT_DENSITY_LB_per_GAL)
            End Select
            txtAdsorbentProperties(1).Text = Format$(ValueToDisplay, GetTheFormat(ValueToDisplay))

       Case 2   'Particle Radius
            Select Case cboAdsorbentPropertyUnits(2).ListIndex
               Case LENGTH_M    'm
                    ValueToDisplay = NowProj.Resin.ParticleRadius
               Case LENGTH_CM   'cm
                    ValueToDisplay = NowProj.Resin.ParticleRadius * LengthConversionFactor(LENGTH_CM)
               Case LENGTH_FT   'ft
                    ValueToDisplay = NowProj.Resin.ParticleRadius * LengthConversionFactor(LENGTH_FT)
               Case LENGTH_IN   'in
                    ValueToDisplay = NowProj.Resin.ParticleRadius * LengthConversionFactor(LENGTH_IN)
            End Select
            txtAdsorbentProperties(2).Text = Format$(ValueToDisplay, GetTheFormat(ValueToDisplay))

       Case 5   'Total Resin Capacity
            Select Case cboAdsorbentPropertyUnits(5).ListIndex
               Case RESIN_CAPACITY_MEQ_per_G   'meq/g resin
                    ValueToDisplay = NowProj.Resin.TotalCapacity
               Case RESIN_CAPACITY_MEQ_per_MLbed   'meq/ml bed
                    ValueToDisplay = NowProj.Resin.TotalCapacity * ResinCapacityConversionFactor(RESIN_CAPACITY_MEQ_per_MLbed)
               Case RESIN_CAPACITY_MEQ_per_MLresin   'meq/ml resin
                    ValueToDisplay = NowProj.Resin.TotalCapacity * ResinCapacityConversionFactor(RESIN_CAPACITY_MEQ_per_MLresin)
            End Select
            txtAdsorbentProperties(5).Text = Format$(ValueToDisplay, GetTheFormat(ValueToDisplay))

    End Select

End Sub

Private Sub cboAdsorbents_Click()
    Dim i As Integer
    Dim oldval As String
    
    If (cboAdsorbents.ListIndex <> Val(cboAdsorbents.Tag)) Then
    
      NowProj.Resin.Name = cboAdsorbents.List(cboAdsorbents.ListIndex)
  
      Select Case cboAdsorbents.ListIndex
         Case 0   'IRN-77
              'enable Cations List Box and Disable Anions list box for IRN-77
              cboIons(0).Enabled = True
              lstIons(0).Enabled = True
              lblCations(0).Enabled = True
              lblCations(1).Enabled = True
              cmdAddDeleteIons(0).Enabled = True
              cmdAddDeleteIons(1).Enabled = True
  
              cboIons(1).Enabled = False
              lstIons(1).Enabled = False
              lblAnions(0).Enabled = False
              lblAnions(1).Enabled = False
              cmdAddDeleteIons(2).Enabled = False
  
              cboIons(2).Clear
              cboKinDimComponent.Clear
              frmInputKineticParameters!cboIon.Clear
              'Load cations into Kinetic Parameters combo Box and onto frmInputKineticParameters
              If frmIonExchangeMain!cboIons(0).ListCount > 0 Then
                 For i = 0 To frmIonExchangeMain!cboIons(0).ListCount - 1
                     cboIons(2).AddItem frmIonExchangeMain!cboIons(0).List(i)
  '                   cboKinDimComponent.AddItem cboIons(0).List(i)
                     frmInputKineticParameters!cboIon.AddItem frmIonExchangeMain!cboIons(0).List(i)
                 Next i
                 cboIons(2).ListIndex = 0
  '               cboKinDimComponent.ListIndex = 0
                 frmInputKineticParameters!cboIon.ListIndex = 0
              End If
             Cations.Available = True
             Anions.Available = False
         Case 1   'IRN-78
              'disable Cations List Box and enable Anions list box for IRN-78
              cboIons(0).Enabled = False
              lstIons(0).Enabled = False
              lblCations(0).Enabled = False
              lblCations(1).Enabled = False
              cmdAddDeleteIons(0).Enabled = False
              cmdAddDeleteIons(1).Enabled = False
  
              cboIons(1).Enabled = True
              lstIons(1).Enabled = True
              lblAnions(0).Enabled = True
              lblAnions(1).Enabled = True
              cmdAddDeleteIons(2).Enabled = True
  
              cboIons(2).Clear
              cboKinDimComponent.Clear
              frmInputKineticParameters!cboIon.Clear
              'Load anions into Kinetic Parameters combo Box
              If frmIonExchangeMain!cboIons(1).ListCount > 0 Then
                 For i = 0 To frmIonExchangeMain!cboIons(1).ListCount - 1
                     cboIons(2).AddItem frmIonExchangeMain!cboIons(1).List(i)
  '                   cboKinDimComponent.AddItem cboIons(1).List(i)
                     frmInputKineticParameters!cboIon.AddItem frmIonExchangeMain!cboIons(1).List(i)
                 Next i
                 cboIons(2).ListIndex = 0
  '               cboKinDimComponent.ListIndex = 0
                 frmInputKineticParameters!cboIon.ListIndex = 0
              End If
  
              Cations.Available = False
              Anions.Available = True
  
         Case 2   'IRA-68
              'Enable both Anions and Cations for IRA-68
              cboIons(0).Enabled = True
              lstIons(0).Enabled = True
              lblCations(0).Enabled = True
              lblCations(1).Enabled = True
              cmdAddDeleteIons(0).Enabled = True
              cmdAddDeleteIons(1).Enabled = True
  
              cboIons(1).Enabled = True
              lstIons(1).Enabled = True
              lblAnions(0).Enabled = True
              lblAnions(1).Enabled = True
              cmdAddDeleteIons(2).Enabled = True
  
              cboIons(2).Clear
              cboKinDimComponent.Clear
              frmInputKineticParameters!cboIon.Clear
  
              If cboIons(0).ListCount > 0 Or cboIons(1).ListCount > 0 Then
                 'Load cations into Kinetic Parameters combo Box
                 For i = 0 To cboIons(0).ListCount - 1
                     cboIons(2).AddItem cboIons(0).List(i)
  '                   cboKinDimComponent.AddItem cboIons(0).List(i)
                     frmInputKineticParameters!cboIon.AddItem cboIons(0).List(i)
                 Next i
              
                 'Load anions into Kinetic Parameters combo Box
                 For i = 0 To cboIons(1).ListCount - 1
                     cboIons(2).AddItem cboIons(1).List(i)
  '                   cboKinDimComponent.AddItem cboIons(1).List(i)
                     frmInputKineticParameters!cboIon.AddItem cboIons(1).List(i)
                 Next i
                 cboIons(2).ListIndex = 0
  '               cboKinDimComponent.ListIndex = 0
                 frmInputKineticParameters!cboIon.ListIndex = 0
              End If
  
              Cations.Available = True
              Anions.Available = True
      End Select
      Call Local_DirtyStatus_Set( _
          True, True)
      Call frmIonExchangeMain_Refresh
    End If
    
End Sub

Private Sub cboBedDataUnits_Click(Index As Integer)
    Dim ValueToDisplay As Double

    Select Case Index
       Case 0   'Bed Length
            Select Case cboBedDataUnits(0).ListIndex
               Case LENGTH_M    'm
                    ValueToDisplay = NowProj.Bed.length
               Case LENGTH_CM   'cm
                    ValueToDisplay = NowProj.Bed.length * LengthConversionFactor(LENGTH_CM)
               Case LENGTH_FT   'ft
                    ValueToDisplay = NowProj.Bed.length * LengthConversionFactor(LENGTH_FT)
               Case LENGTH_IN   'in
                    ValueToDisplay = NowProj.Bed.length * LengthConversionFactor(LENGTH_IN)
            End Select
            txtBedData(0).Text = Format$(ValueToDisplay, GetTheFormat(ValueToDisplay))

       Case 1   'Bed Diameter
            Select Case cboBedDataUnits(1).ListIndex
               Case LENGTH_M    'm
                    ValueToDisplay = NowProj.Bed.Diameter
               Case LENGTH_CM   'cm
                    ValueToDisplay = NowProj.Bed.Diameter * LengthConversionFactor(LENGTH_CM)
               Case LENGTH_FT   'ft
                    ValueToDisplay = NowProj.Bed.Diameter * LengthConversionFactor(LENGTH_FT)
               Case LENGTH_IN   'in
                    ValueToDisplay = NowProj.Bed.Diameter * LengthConversionFactor(LENGTH_IN)
            End Select
            txtBedData(1).Text = Format$(ValueToDisplay, GetTheFormat(ValueToDisplay))

       Case 2   'Bed Mass
            Select Case cboBedDataUnits(2).ListIndex
               Case MASS_KG   'kg
                    ValueToDisplay = NowProj.Bed.Weight
               Case MASS_G    'g
                    ValueToDisplay = NowProj.Bed.Weight * MassConversionFactor(MASS_G)
               Case MASS_LB   'lb
                    ValueToDisplay = NowProj.Bed.Weight * MassConversionFactor(MASS_LB)
            End Select
            txtBedData(2).Text = Format$(ValueToDisplay, GetTheFormat(ValueToDisplay))

       Case 3   'Flow Rate
            Select Case cboBedDataUnits(3).ListIndex
               Case FLOW_M3_per_S     'm3/s
                    ValueToDisplay = NowProj.Bed.Flowrate.Value
               Case FLOW_M3_per_D     'm3/d
                    ValueToDisplay = NowProj.Bed.Flowrate.Value * FlowConversionFactor(FLOW_M3_per_D)
               Case FLOW_CM3_per_S    'cm3/s
                    ValueToDisplay = NowProj.Bed.Flowrate.Value * FlowConversionFactor(FLOW_CM3_per_S)
               Case FLOW_ML_per_MIN   'ml/min
                    ValueToDisplay = NowProj.Bed.Flowrate.Value * FlowConversionFactor(FLOW_ML_per_MIN)
               Case FLOW_FT3_per_S    'ft3/s
                    ValueToDisplay = NowProj.Bed.Flowrate.Value * FlowConversionFactor(FLOW_FT3_per_S)
               Case FLOW__FT3_per_D   'ft3/d
                    ValueToDisplay = NowProj.Bed.Flowrate.Value * FlowConversionFactor(FLOW__FT3_per_D)
               Case FLOW_GPM   'gpm
                    ValueToDisplay = NowProj.Bed.Flowrate.Value * FlowConversionFactor(FLOW_GPM)
               Case FLOW_GPD   'gpd
                    ValueToDisplay = NowProj.Bed.Flowrate.Value * FlowConversionFactor(FLOW_GPD)
               Case FLOW_MGD   'MGD
                    ValueToDisplay = NowProj.Bed.Flowrate.Value * FlowConversionFactor(FLOW_MGD)
            End Select
            txtBedData(3).Text = Format$(ValueToDisplay, GetTheFormat(ValueToDisplay))

       Case 4   'EBCT
            Select Case cboBedDataUnits(4).ListIndex
               Case TIME_MIN   'min
                    ValueToDisplay = NowProj.Bed.EBCT.Value
               Case TIME_S     's
                    ValueToDisplay = NowProj.Bed.EBCT.Value * TimeConversionFactor(TIME_S)
               Case TIME_HR    'hr
                    ValueToDisplay = NowProj.Bed.EBCT.Value * TimeConversionFactor(TIME_HR)
               Case TIME_D     'd
                    ValueToDisplay = NowProj.Bed.EBCT.Value * TimeConversionFactor(TIME_D)
            End Select
            txtBedData(4).Text = Format$(ValueToDisplay, GetTheFormat(ValueToDisplay))


    End Select

End Sub

Private Sub cboIons_Click(Index As Integer)
    Dim i As Integer

    If Index = 2 Then
       Exit Sub
    End If

    If Index = 0 Then      'Cations
       NowProj.PresaturantCation = cboIons(0).ListIndex + 1
    ElseIf Index = 1 Then  'anions
       NowProj.PresaturantAnion = cboIons(1).ListIndex + 1
    End If
    
    lstIons(Index).Clear
    For i = 0 To cboIons(Index).ListCount - 1
        If i <> cboIons(Index).ListIndex Then
           lstIons(Index).AddItem cboIons(Index).List(i)
        End If
    Next i

    For i = 0 To lstIons(Index).ListCount - 1
        lstIons(Index).Selected(i) = False
    Next i

    For i = 1 To MAX_CHEMICAL
        NowProj.Resin.PresaturantPercentage(i) = 0#
    Next i

    'Start Cations_Selected or Anions_Selected Arrays
    Select Case Index
       Case 0   'Cations
          NumSelectedCations = 1
          Cations_Selected(1) = NowProj.PresaturantCation
          NowProj.Resin.PresaturantPercentage(Cations_Selected(1)) = 100#
       Case 1   'Anions
          NumSelectedAnions = 1
          Anions_Selected(1) = NowProj.PresaturantAnion
          NowProj.Resin.PresaturantPercentage(Anions_Selected(1)) = 100#
    End Select

    cboKinDimComponent.Clear
    cboKinDimComponent.Enabled = False

    For i = 3 To 8
        lblKineticDimensionlessValue(i).Caption = ""
    Next i

    mnuRun(0).Enabled = False
    mnuOptions(4).Enabled = False

End Sub

Private Sub cboKinDimComponent_Click()
    Dim i As Integer, ListIndex As Integer
    Dim FoundAnion As Integer, FoundCation As Integer
    Dim ValueToDisplay As Double
    Dim NumberOfIonFound As Integer

    ListIndex = cboKinDimComponent.ListIndex

    'Display values for kf, Dl, and Dp on Main form

    If Cations.Available And Anions.Available Then   'May be editing either cations or anions
       FoundCation = False
       FoundAnion = False
       For i = 0 To cboKinDimComponent.ListCount - 1
           If NowProj.Cation(Cations_Selected(i + 1)).Name = cboKinDimComponent.List(ListIndex) Then
              FoundCation = True
              NumberOfIonFound = Cations_Selected(i + 1)
              Exit For
           End If

           If NowProj.Anion(Anions_Selected(i + 1)).Name = cboKinDimComponent.List(ListIndex) Then
              FoundAnion = True
              NumberOfIonFound = Anions_Selected(i + 1)
              Exit For
           End If
       Next i

       If FoundCation Then

          'Display kf
          ValueToDisplay = NowProj.Cation(NumberOfIonFound).Kinetic.IonicTransportCoefficient.Value
          lblKineticDimensionlessValue(0).Caption = Format$(ValueToDisplay, GetTheFormat(ValueToDisplay))

          'Display Dl
          ValueToDisplay = NowProj.Cation(NumberOfIonFound).Kinetic.LiquidDiffusivity.Value
          lblKineticDimensionlessValue(1).Caption = Format$(ValueToDisplay, GetTheFormat(ValueToDisplay))

          'Display Dp
          ValueToDisplay = NowProj.Cation(NumberOfIonFound).Kinetic.PoreDiffusivity.Value
          lblKineticDimensionlessValue(2).Caption = Format$(ValueToDisplay, GetTheFormat(ValueToDisplay))

          If Not NowProj.OKToGetCationDimensionless Then
             For i = 3 To 8
                 lblKineticDimensionlessValue(i).Caption = ""
             Next i
             Exit Sub
          End If

          'Display Dgs
          ValueToDisplay = NowProj.Cation(NumberOfIonFound).Dimensionless.SurfaceDistributionParameter
          lblKineticDimensionlessValue(3).Caption = Format$(ValueToDisplay, GetTheFormat(ValueToDisplay))

          'Display Dgp
          ValueToDisplay = NowProj.Cation(NumberOfIonFound).Dimensionless.PoreDistributionParameter
          lblKineticDimensionlessValue(4).Caption = Format$(ValueToDisplay, GetTheFormat(ValueToDisplay))

          'Display Dgt
          ValueToDisplay = NowProj.Cation(NumberOfIonFound).Dimensionless.TotalDistributionParameter
          lblKineticDimensionlessValue(5).Caption = Format$(ValueToDisplay, GetTheFormat(ValueToDisplay))

          'Display Edp
          ValueToDisplay = NowProj.Cation(NumberOfIonFound).Dimensionless.PoreDiffusionModulus
          lblKineticDimensionlessValue(6).Caption = Format$(ValueToDisplay, GetTheFormat(ValueToDisplay))
     
          'Display St
          ValueToDisplay = NowProj.Cation(NumberOfIonFound).Dimensionless.StantonNumber
          lblKineticDimensionlessValue(7).Caption = Format$(ValueToDisplay, GetTheFormat(ValueToDisplay))

          'Display Bip
          ValueToDisplay = NowProj.Cation(NumberOfIonFound).Dimensionless.PoreBiotNumber
          lblKineticDimensionlessValue(8).Caption = Format$(ValueToDisplay, GetTheFormat(ValueToDisplay))
       
       End If
       
       If FoundAnion Then
          'Display kf
          ValueToDisplay = NowProj.Anion(NumberOfIonFound).Kinetic.IonicTransportCoefficient.Value
          lblKineticDimensionlessValue(0).Caption = Format$(ValueToDisplay, GetTheFormat(ValueToDisplay))

          'Display Dl
          ValueToDisplay = NowProj.Anion(NumberOfIonFound).Kinetic.LiquidDiffusivity.Value
          lblKineticDimensionlessValue(1).Caption = Format$(ValueToDisplay, GetTheFormat(ValueToDisplay))

          'Display Dp
          ValueToDisplay = NowProj.Anion(NumberOfIonFound).Kinetic.PoreDiffusivity.Value
          lblKineticDimensionlessValue(2).Caption = Format$(ValueToDisplay, GetTheFormat(ValueToDisplay))

          If Not NowProj.OKToGetAnionDimensionless Then
             For i = 3 To 8
                 lblKineticDimensionlessValue(i).Caption = ""
             Next i
             Exit Sub
          End If

          'Display Dgs
          ValueToDisplay = NowProj.Anion(NumberOfIonFound).Dimensionless.SurfaceDistributionParameter
          lblKineticDimensionlessValue(3).Caption = Format$(ValueToDisplay, GetTheFormat(ValueToDisplay))

          'Display Dgp
          ValueToDisplay = NowProj.Anion(NumberOfIonFound).Dimensionless.PoreDistributionParameter
          lblKineticDimensionlessValue(4).Caption = Format$(ValueToDisplay, GetTheFormat(ValueToDisplay))

          'Display Dgt
          ValueToDisplay = NowProj.Anion(NumberOfIonFound).Dimensionless.TotalDistributionParameter
          lblKineticDimensionlessValue(5).Caption = Format$(ValueToDisplay, GetTheFormat(ValueToDisplay))

          'Display Edp
          ValueToDisplay = NowProj.Anion(NumberOfIonFound).Dimensionless.PoreDiffusionModulus
          lblKineticDimensionlessValue(6).Caption = Format$(ValueToDisplay, GetTheFormat(ValueToDisplay))
     
          'Display St
          ValueToDisplay = NowProj.Anion(NumberOfIonFound).Dimensionless.StantonNumber
          lblKineticDimensionlessValue(7).Caption = Format$(ValueToDisplay, GetTheFormat(ValueToDisplay))

          'Display Bip
          ValueToDisplay = NowProj.Anion(NumberOfIonFound).Dimensionless.PoreBiotNumber
          lblKineticDimensionlessValue(8).Caption = Format$(ValueToDisplay, GetTheFormat(ValueToDisplay))

       End If

    ElseIf Cations.Available Then  'Only cations in list
       FoundCation = False
       For i = 0 To frmIonExchangeMain!cboKinDimComponent.ListCount - 1
           If NowProj.Cation(Cations_Selected(i + 1)).Name = _
                frmIonExchangeMain!cboKinDimComponent.List(ListIndex) Then
              FoundCation = True
              NumberOfIonFound = Cations_Selected(i + 1)
              Exit For
           End If

       Next i
       

          'Display kf
          ValueToDisplay = NowProj.Cation(NumberOfIonFound).Kinetic.IonicTransportCoefficient.Value
          lblKineticDimensionlessValue(0).Caption = Format$(ValueToDisplay, GetTheFormat(ValueToDisplay))

          'Display Dl
          ValueToDisplay = NowProj.Cation(NumberOfIonFound).Kinetic.LiquidDiffusivity.Value
          lblKineticDimensionlessValue(1).Caption = Format$(ValueToDisplay, GetTheFormat(ValueToDisplay))

          'Display Dp
          ValueToDisplay = NowProj.Cation(NumberOfIonFound).Kinetic.PoreDiffusivity.Value
          lblKineticDimensionlessValue(2).Caption = Format$(ValueToDisplay, GetTheFormat(ValueToDisplay))

          If Not NowProj.OKToGetCationDimensionless Then
             For i = 3 To 8
                 lblKineticDimensionlessValue(i).Caption = ""
             Next i
             Exit Sub
          End If

          'Display Dgs
          ValueToDisplay = NowProj.Cation(NumberOfIonFound).Dimensionless.SurfaceDistributionParameter
          lblKineticDimensionlessValue(3).Caption = Format$(ValueToDisplay, GetTheFormat(ValueToDisplay))

          'Display Dgp
          ValueToDisplay = NowProj.Cation(NumberOfIonFound).Dimensionless.PoreDistributionParameter
          lblKineticDimensionlessValue(4).Caption = Format$(ValueToDisplay, GetTheFormat(ValueToDisplay))

          'Display Dgt
          ValueToDisplay = NowProj.Cation(NumberOfIonFound).Dimensionless.TotalDistributionParameter
          lblKineticDimensionlessValue(5).Caption = Format$(ValueToDisplay, GetTheFormat(ValueToDisplay))

          'Display Edp
          ValueToDisplay = NowProj.Cation(NumberOfIonFound).Dimensionless.PoreDiffusionModulus
          lblKineticDimensionlessValue(6).Caption = Format$(ValueToDisplay, GetTheFormat(ValueToDisplay))
     
          'Display St
          ValueToDisplay = NowProj.Cation(NumberOfIonFound).Dimensionless.StantonNumber
          lblKineticDimensionlessValue(7).Caption = Format$(ValueToDisplay, GetTheFormat(ValueToDisplay))

          'Display Bip
          ValueToDisplay = NowProj.Cation(NumberOfIonFound).Dimensionless.PoreBiotNumber
          lblKineticDimensionlessValue(8).Caption = Format$(ValueToDisplay, GetTheFormat(ValueToDisplay))

    ElseIf Anions.Available Then  'Only anions in list
       FoundAnion = False
       For i = 0 To cboKinDimComponent.ListCount - 1
           If NowProj.Anion(Anions_Selected(i + 1)).Name = cboKinDimComponent.List(ListIndex) Then
              FoundAnion = True
              NumberOfIonFound = Anions_Selected(i + 1)
              Exit For
           End If
       Next i
      

          'Display kf
          ValueToDisplay = NowProj.Anion(NumberOfIonFound).Kinetic.IonicTransportCoefficient.Value
          lblKineticDimensionlessValue(0).Caption = Format$(ValueToDisplay, GetTheFormat(ValueToDisplay))

          'Display Dl
          ValueToDisplay = NowProj.Anion(NumberOfIonFound).Kinetic.LiquidDiffusivity.Value
          lblKineticDimensionlessValue(1).Caption = Format$(ValueToDisplay, GetTheFormat(ValueToDisplay))

          'Display Dp
          ValueToDisplay = NowProj.Anion(NumberOfIonFound).Kinetic.PoreDiffusivity.Value
          lblKineticDimensionlessValue(2).Caption = Format$(ValueToDisplay, GetTheFormat(ValueToDisplay))

          If Not NowProj.OKToGetAnionDimensionless Then
             For i = 3 To 8
                 lblKineticDimensionlessValue(i).Caption = ""
             Next i
             Exit Sub
          End If

          'Display Dgs
          ValueToDisplay = NowProj.Anion(NumberOfIonFound).Dimensionless.SurfaceDistributionParameter
          lblKineticDimensionlessValue(3).Caption = Format$(ValueToDisplay, GetTheFormat(ValueToDisplay))

          'Display Dgp
          ValueToDisplay = NowProj.Anion(NumberOfIonFound).Dimensionless.PoreDistributionParameter
          lblKineticDimensionlessValue(4).Caption = Format$(ValueToDisplay, GetTheFormat(ValueToDisplay))

          'Display Dgt
          ValueToDisplay = NowProj.Anion(NumberOfIonFound).Dimensionless.TotalDistributionParameter
          lblKineticDimensionlessValue(5).Caption = Format$(ValueToDisplay, GetTheFormat(ValueToDisplay))

          'Display Edp
          ValueToDisplay = NowProj.Anion(NumberOfIonFound).Dimensionless.PoreDiffusionModulus
          lblKineticDimensionlessValue(6).Caption = Format$(ValueToDisplay, GetTheFormat(ValueToDisplay))
     
          'Display St
          ValueToDisplay = NowProj.Anion(NumberOfIonFound).Dimensionless.StantonNumber
          lblKineticDimensionlessValue(7).Caption = Format$(ValueToDisplay, GetTheFormat(ValueToDisplay))

          'Display Bip
          ValueToDisplay = NowProj.Anion(NumberOfIonFound).Dimensionless.PoreBiotNumber
          lblKineticDimensionlessValue(8).Caption = Format$(ValueToDisplay, GetTheFormat(ValueToDisplay))

    End If
    
End Sub

Private Sub cboOperatingConditionsUnits_Click(Index As Integer)
    Dim ValueToDisplay As Double

    Select Case Index
       Case 0   'Operating Pressure
            Select Case cboOperatingConditionsUnits(0).ListIndex
               Case PRESSURE_PA    'Pa
                    ValueToDisplay = NowProj.Operating.Pressure
               Case PRESSURE_KPA   'kPa
                    ValueToDisplay = NowProj.Operating.Pressure * PressureConversionFactor(PRESSURE_KPA)
               Case PRESSURE_BARS   'bars
                    ValueToDisplay = NowProj.Operating.Pressure * PressureConversionFactor(PRESSURE_BARS)
               Case PRESSURE_ATM   'atm
                    ValueToDisplay = NowProj.Operating.Pressure * PressureConversionFactor(PRESSURE_ATM)
               Case PRESSURE_PSI   'psi
                    ValueToDisplay = NowProj.Operating.Pressure * PressureConversionFactor(PRESSURE_PSI)
               Case PRESSURE_MMHG   'mm Hg
                    ValueToDisplay = NowProj.Operating.Pressure * PressureConversionFactor(PRESSURE_MMHG)
               Case PRESSURE_MH2O   'm H2O
                    ValueToDisplay = NowProj.Operating.Pressure * PressureConversionFactor(PRESSURE_MH2O)
               Case PRESSURE_FTH2O   'ft H2O
                    ValueToDisplay = NowProj.Operating.Pressure * PressureConversionFactor(PRESSURE_FTH2O)
               Case PRESSURE_INHG   'in. Hg
                    ValueToDisplay = NowProj.Operating.Pressure * PressureConversionFactor(PRESSURE_INHG)
            End Select
            txtOperatingConditions(0).Text = Format$(ValueToDisplay, GetTheFormat(ValueToDisplay))

       Case 1   'Operating Temperature
            Select Case cboOperatingConditionsUnits(1).ListIndex
               Case TEMPERATURE_K    'K
                    ValueToDisplay = NowProj.Operating.Temperature
               Case TEMPERATURE_C   'C
                    ValueToDisplay = TemperatureConversion(TEMPERATURE_C, NowProj.Operating.Temperature)
               Case TEMPERATURE_R   'R
                    ValueToDisplay = TemperatureConversion(TEMPERATURE_R, NowProj.Operating.Temperature)
               Case TEMPERATURE_F   'F
                    ValueToDisplay = TemperatureConversion(TEMPERATURE_F, NowProj.Operating.Temperature)
            End Select
            txtOperatingConditions(1).Text = Format$(ValueToDisplay, GetTheFormat(ValueToDisplay))

    End Select

End Sub

Private Sub cmdAddDeleteIons_Click(Index As Integer)
    Dim i As Integer
    Dim msg As String
    Dim ListIndex As Integer
    
    Select Case Index
       Case 0   'Add Cation

            If NowProj.NumberOfCations = MAX_CHEMICAL Then
               msg = "Adding another cation would exceed the maximum number of "
               msg = msg & "cations allowed for a simulation.  If you would like to "
               msg = msg + "add another cation, you must remove one of the "
               msg = msg + "current cations first."
               MsgBox msg, MB_ICONSTOP, "Too Many Cations"
            End If

            frmAddComponent.Caption = "Add Cation"
            frmAddComponent!lblValenceSign.Caption = "+"
            frmAddComponent!txtAddIon(0).Text = "Cation"
            frmAddComponent!txtAddIon(1).Text = Trim$(Str$(DefaultCation.MolecularWeight))
            frmAddComponent!txtAddIon(2).Text = Trim$(Str$(DefaultCation.InitialConcentration))
            frmAddComponent!lblValence.Caption = Trim$(Str$(CInt(DefaultCation.Valence)))
            If NowProj.NumberOfCations > 0 Then
               frmAddComponent!txtAlphaValue.Text = Trim$(Str$(DefaultCation.SeparationFactor))
            Else
               frmAddComponent!txtAlphaValue.Text = "1.00"
            End If

            SeparationFactorInput.Row = NowProj.CationSeparationFactorInput.Row
            SeparationFactorInput.Value = NowProj.CationSeparationFactorInput.Value
            If SeparationFactorInput.Row = True Then
               frmAddComponent!lblAlpha(1).Caption = frmAddComponent!txtAddIon(0).Text
            Else
               frmAddComponent!lblAlpha(2).Caption = frmAddComponent!txtAddIon(0).Text
            End If

            If NowProj.NumberOfCations > 0 Then
               If SeparationFactorInput.Row = True Then
                  frmAddComponent!lblAlpha(2).Caption = NowProj.Cation(SeparationFactorInput.Value - 10).Name
               Else
                  frmAddComponent!lblAlpha(1).Caption = NowProj.Cation(SeparationFactorInput.Value).Name
               End If
            Else
               If SeparationFactorInput.Row = True Then
                  frmAddComponent!lblAlpha(2).Caption = frmAddComponent!lblAlpha(1).Caption
               Else
                  frmAddComponent!lblAlpha(1).Caption = frmAddComponent!lblAlpha(2).Caption
               End If
            End If

            If Trim$(frmAddComponent!lblAlpha(1).Caption) = Trim$(frmAddComponent!lblAlpha(2).Caption) Then
               frmAddComponent!txtAlphaValue.Enabled = False
            Else
               frmAddComponent!txtAlphaValue.Enabled = True
            End If
                  
            For i = 0 To frmAddComponent!cboAnion.ListCount - 1
                If DefaultCation.Kinetic.NernstHaskellAnion.Ion_Name = frmAddComponent!cboAnion.List(i) Then
                   frmAddComponent!cboAnion.ListIndex = i
                End If
            Next i
            For i = 0 To frmAddComponent!cboCation.ListCount - 1
                If DefaultCation.Kinetic.NernstHaskellCation.Ion_Name = frmAddComponent!cboCation.List(i) Then
                   frmAddComponent!cboCation.ListIndex = i
                End If
            Next i

            ChangedIon = DefaultCation

            If NowProj.NumberOfCations = 0 Then
               ChangedIon.SeparationFactor = 1#
               frmAddComponent!txtAlphaValue.Enabled = False
            Else
               frmAddComponent!txtAlphaValue.Enabled = True
            End If

            'Generate click events on appropriate units
            ListIndex = frmAddComponent!cboAddIonUnits(0).ListIndex
            frmAddComponent!cboAddIonUnits(0).ListIndex = -1
            frmAddComponent!cboAddIonUnits(0).ListIndex = ListIndex

            ListIndex = frmAddComponent!cboAddIonUnits(1).ListIndex
            frmAddComponent!cboAddIonUnits(1).ListIndex = -1
            frmAddComponent!cboAddIonUnits(1).ListIndex = ListIndex

            AddingCation = True
            AddingAnion = False
            EditingCation = False
            EditingAnion = False
            NumberOfIons = NowProj.NumberOfCations + 1
''''            ReDim NowProj.Cation(1 To NumberOfIons)
            NumberOfIonToEdit = NumberOfIons
            If NumberOfIons > 1 Then
               frmAddComponent!cmdViewSeparationFactors.Enabled = True
            Else
               frmAddComponent!cmdViewSeparationFactors.Enabled = False
            End If

            NowProj.Cation(NumberOfIons).Name = "Cation"
            NowProj.Cation(NumberOfIons).SeparationFactor = DefaultCation.SeparationFactor

            For i = 1 To NumberOfIons
                OneDimSeparationFactors(i) = NowProj.Cation(i).SeparationFactor
            Next i

            frmAddComponent.Show 1

             NowProj.CationSeparationFactorInput.Row = SeparationFactorInput.Row
             NowProj.CationSeparationFactorInput.Value = SeparationFactorInput.Value

            AddingCation = False

       Case 1   'Remove Ion

       Case 2   'Add Anion

            If NowProj.NumberOfAnions = MAX_CHEMICAL Then
               msg = "Adding another Anion would exceed the maximum number of "
               msg = msg & "anions allowed for a simulation.  If you would like to "
               msg = msg + "add another anion, you must remove one of the "
               msg = msg + "current anions first."
               MsgBox msg, MB_ICONSTOP, "Too Many Anions"
            End If

            frmAddComponent.Caption = "Add Anion"
            frmAddComponent!lblValenceSign.Caption = "-"
            frmAddComponent!txtAddIon(0).Text = "Anion"
            frmAddComponent!txtAddIon(1).Text = Trim$(Str$(DefaultAnion.MolecularWeight))
            frmAddComponent!txtAddIon(2).Text = Trim$(Str$(DefaultAnion.InitialConcentration))
            frmAddComponent!lblValence.Caption = Trim$(Str$(CInt(DefaultAnion.Valence)))
            frmAddComponent!txtAlphaValue.Text = Trim$(Str$(DefaultAnion.SeparationFactor))

            SeparationFactorInput.Row = NowProj.AnionSeparationFactorInput.Row
            SeparationFactorInput.Value = NowProj.AnionSeparationFactorInput.Value
            If SeparationFactorInput.Row = True Then
               frmAddComponent!lblAlpha(1).Caption = frmAddComponent!txtAddIon(0).Text
            Else
               frmAddComponent!lblAlpha(2).Caption = frmAddComponent!txtAddIon(0).Text
            End If

            If NowProj.NumberOfAnions > 0 Then
               If SeparationFactorInput.Row = True Then
                  frmAddComponent!lblAlpha(2).Caption = NowProj.Anion(SeparationFactorInput.Value - 10).Name
               Else
                  frmAddComponent!lblAlpha(1).Caption = NowProj.Anion(SeparationFactorInput.Value).Name
               End If
            Else
               If SeparationFactorInput.Row = True Then
                  frmAddComponent!lblAlpha(2).Caption = frmAddComponent!lblAlpha(1).Caption
               Else
                  frmAddComponent!lblAlpha(1).Caption = frmAddComponent!lblAlpha(2).Caption
               End If
            End If

            If Trim$(frmAddComponent!lblAlpha(1).Caption) = Trim$(frmAddComponent!lblAlpha(2).Caption) Then
               frmAddComponent!txtAlphaValue.Enabled = False
            Else
               frmAddComponent!txtAlphaValue.Enabled = True
            End If

            For i = 0 To frmAddComponent!cboCation.ListCount - 1
                If DefaultAnion.Kinetic.NernstHaskellCation.Ion_Name = frmAddComponent!cboCation.List(i) Then
                   frmAddComponent!cboCation.ListIndex = i
                End If
            Next i
            For i = 0 To frmAddComponent!cboAnion.ListCount - 1
                If DefaultAnion.Kinetic.NernstHaskellAnion.Ion_Name = frmAddComponent!cboAnion.List(i) Then
                   frmAddComponent!cboAnion.ListIndex = i
                End If
            Next i

            ChangedIon = DefaultAnion
            
            If NowProj.NumberOfAnions = 0 Then
               ChangedIon.SeparationFactor = 1#
               frmAddComponent!txtAlphaValue.Enabled = False
            Else
               frmAddComponent!txtAlphaValue.Enabled = True
            End If

            'Generate click events on appropriate units
            ListIndex = frmAddComponent!cboAddIonUnits(0).ListIndex
            frmAddComponent!cboAddIonUnits(0).ListIndex = -1
            frmAddComponent!cboAddIonUnits(0).ListIndex = ListIndex

            ListIndex = frmAddComponent!cboAddIonUnits(1).ListIndex
            frmAddComponent!cboAddIonUnits(1).ListIndex = -1
            frmAddComponent!cboAddIonUnits(1).ListIndex = ListIndex

            AddingCation = False
            AddingAnion = True
            EditingCation = False
            EditingAnion = False
            NumberOfIons = NowProj.NumberOfAnions + 1
''''            ReDim NowProj.Anion(1 To NumberOfIons)
            NumberOfIonToEdit = NumberOfIons
            If NumberOfIons > 1 Then
               frmAddComponent!cmdViewSeparationFactors.Enabled = True
            Else
               frmAddComponent!cmdViewSeparationFactors.Enabled = False
            End If

            NowProj.Anion(NumberOfIons).Name = "Anion"
            NowProj.Anion(NumberOfIons).SeparationFactor = DefaultAnion.SeparationFactor

            For i = 1 To NumberOfIons
                OneDimSeparationFactors(i) = NowProj.Anion(i).SeparationFactor
            Next i

            frmAddComponent.Show 1

             NowProj.AnionSeparationFactorInput.Row = SeparationFactorInput.Row
             NowProj.AnionSeparationFactorInput.Value = SeparationFactorInput.Value

            AddingAnion = False

    End Select

End Sub



Private Sub cmdEditProperties_Click()
    Dim i As Integer
    Dim FoundCation As Integer, FoundAnion As Integer
    Dim ListIndex As Integer

    If Trim$(cboIons(2).List(cboIons(2).ListIndex)) = "" Then
      Call Show_Message00("Please select an ion", _
        vbInformation, _
        App.title)
      Exit Sub
    End If
    FoundCation = False
    FoundAnion = False
    NumberOfIonToEdit = 0
    For i = 1 To NowProj.NumberOfCations
        If Trim$(NowProj.Cation(i).Name) = Trim$(cboIons(2).List(cboIons(2).ListIndex)) Then
           NumberOfIonToEdit = i
           FoundCation = True
           Exit For
        End If
    Next i
    If Not FoundCation Then
       For i = 1 To NowProj.NumberOfAnions
           If Trim$(NowProj.Anion(i).Name) = Trim$(cboIons(2).List(cboIons(2).ListIndex)) Then
              NumberOfIonToEdit = i
              FoundAnion = True
              Exit For
           End If
       Next i
    End If

    If FoundCation = True Then
       frmAddComponent.Caption = "Edit Cation"
       frmAddComponent!lblValenceSign.Caption = "+"
       frmAddComponent!txtAddIon(0).Text = Trim$(NowProj.Cation(NumberOfIonToEdit).Name)
       frmAddComponent!txtAddIon(0).Enabled = False
       frmAddComponent!txtAddIon(1).Text = Trim$(Str$(NowProj.Cation(NumberOfIonToEdit).MolecularWeight))
       frmAddComponent!txtAddIon(2).Text = Trim$(Str$(NowProj.Cation(NumberOfIonToEdit).InitialConcentration))
       frmAddComponent!lblValence.Caption = Trim$(Str$(CInt(NowProj.Cation(NumberOfIonToEdit).Valence)))
       frmAddComponent!txtAlphaValue.Text = Trim$(Str$(NowProj.Cation(NumberOfIonToEdit).SeparationFactor))

        SeparationFactorInput.Row = NowProj.CationSeparationFactorInput.Row
        SeparationFactorInput.Value = NowProj.CationSeparationFactorInput.Value
        If SeparationFactorInput.Row = True Then
           frmAddComponent!lblAlpha(1).Caption = frmAddComponent!txtAddIon(0).Text
        Else
           frmAddComponent!lblAlpha(2).Caption = frmAddComponent!txtAddIon(0).Text
        End If
  
        If NowProj.NumberOfCations > 0 Then
           If SeparationFactorInput.Row = True Then
              frmAddComponent!lblAlpha(2).Caption = NowProj.Cation(SeparationFactorInput.Value - 10).Name
           Else
              frmAddComponent!lblAlpha(1).Caption = NowProj.Cation(SeparationFactorInput.Value).Name
           End If
        Else
           If SeparationFactorInput.Row = True Then
              frmAddComponent!lblAlpha(2).Caption = frmAddComponent!lblAlpha(1).Caption
           Else
              frmAddComponent!lblAlpha(1).Caption = frmAddComponent!lblAlpha(2).Caption
           End If
        End If
  
        If Trim$(frmAddComponent!lblAlpha(1).Caption) = Trim$(frmAddComponent!lblAlpha(2).Caption) Then
           frmAddComponent!txtAlphaValue.Enabled = False
        Else
           frmAddComponent!txtAlphaValue.Enabled = True
        End If

       For i = 0 To frmAddComponent!cboAnion.ListCount - 1
           If NowProj.Cation(NumberOfIonToEdit).Kinetic.NernstHaskellAnion.Ion_Name = _
              frmAddComponent!cboAnion.List(i) Then
              frmAddComponent!cboAnion.ListIndex = i
           End If
       Next i

        frmAddComponent.lblAddIonValue(0).Caption = _
            NowProj.Cation(NumberOfIonToEdit).Kinetic.NernstHaskellAnion.Valence
        frmAddComponent.lblAddIonValue(1).Caption = _
            NowProj.Cation(NumberOfIonToEdit).Kinetic.NernstHaskellAnion.LimitingIonicConductance

       For i = 0 To frmAddComponent!cboCation.ListCount - 1
           If NowProj.Cation(NumberOfIonToEdit).Kinetic.NernstHaskellCation.Ion_Name = _
              frmAddComponent!cboCation.List(i) Then
              frmAddComponent!cboCation.ListIndex = i
           End If
       Next i
       
        frmAddComponent.lblAddIonValue(2).Caption = _
            NowProj.Cation(NumberOfIonToEdit).Kinetic.NernstHaskellCation.Valence
        frmAddComponent.lblAddIonValue(3).Caption = _
            NowProj.Cation(NumberOfIonToEdit).Kinetic.NernstHaskellCation.LimitingIonicConductance
            
       ChangedIon = NowProj.Cation(NumberOfIonToEdit)

       'Generate click events on appropriate units
       ListIndex = frmAddComponent!cboAddIonUnits(0).ListIndex
       frmAddComponent!cboAddIonUnits(0).ListIndex = -1
       frmAddComponent!cboAddIonUnits(0).ListIndex = ListIndex

       ListIndex = frmAddComponent!cboAddIonUnits(1).ListIndex
       frmAddComponent!cboAddIonUnits(1).ListIndex = -1
       frmAddComponent!cboAddIonUnits(1).ListIndex = ListIndex

       AddingCation = False
       AddingAnion = False
       EditingCation = True
       EditingAnion = False
       NumberOfIons = NowProj.NumberOfCations
            If NumberOfIons > 1 Then
               frmAddComponent!cmdViewSeparationFactors.Enabled = True
            Else
               frmAddComponent!cmdViewSeparationFactors.Enabled = False
            End If

            For i = 1 To NumberOfIons
                OneDimSeparationFactors(i) = NowProj.Cation(i).SeparationFactor
            Next i

       frmAddComponent.Show 1

            NowProj.CationSeparationFactorInput.Row = SeparationFactorInput.Row
            NowProj.CationSeparationFactorInput.Value = SeparationFactorInput.Value

       EditingCation = False

    ElseIf FoundAnion = True Then
       frmAddComponent.Caption = "Edit Anion"
       frmAddComponent!lblValenceSign.Caption = "-"
       frmAddComponent!txtAddIon(0).Text = Trim$(NowProj.Anion(NumberOfIonToEdit).Name)
       frmAddComponent!txtAddIon(0).Enabled = False
       frmAddComponent!txtAddIon(1).Text = Trim$(Str$(NowProj.Anion(NumberOfIonToEdit).MolecularWeight))
       frmAddComponent!txtAddIon(2).Text = Trim$(Str$(NowProj.Anion(NumberOfIonToEdit).InitialConcentration))
       frmAddComponent!lblValence.Caption = Trim$(Str$(CInt(NowProj.Anion(NumberOfIonToEdit).Valence)))
       frmAddComponent!txtAlphaValue.Text = Trim$(Str$(NowProj.Anion(NumberOfIonToEdit).SeparationFactor))

            SeparationFactorInput.Row = NowProj.AnionSeparationFactorInput.Row
            SeparationFactorInput.Value = NowProj.AnionSeparationFactorInput.Value
            If SeparationFactorInput.Row = True Then
               frmAddComponent!lblAlpha(1).Caption = frmAddComponent!txtAddIon(0).Text
            Else
               frmAddComponent!lblAlpha(2).Caption = frmAddComponent!txtAddIon(0).Text
            End If

            If NowProj.NumberOfAnions > 0 Then
               If SeparationFactorInput.Row = True Then
                  frmAddComponent!lblAlpha(2).Caption = NowProj.Anion(SeparationFactorInput.Value - 10).Name
               Else
                  frmAddComponent!lblAlpha(1).Caption = NowProj.Anion(SeparationFactorInput.Value).Name
               End If
            Else
               If SeparationFactorInput.Row = True Then
                  frmAddComponent!lblAlpha(2).Caption = frmAddComponent!lblAlpha(1).Caption
               Else
                  frmAddComponent!lblAlpha(1).Caption = frmAddComponent!lblAlpha(2).Caption
               End If
            End If

            If Trim$(frmAddComponent!lblAlpha(1).Caption) = Trim$(frmAddComponent!lblAlpha(2).Caption) Then
               frmAddComponent!txtAlphaValue.Enabled = False
            Else
               frmAddComponent!txtAlphaValue.Enabled = True
            End If

       For i = 0 To frmAddComponent!cboAnion.ListCount - 1
           If NowProj.Anion(NumberOfIonToEdit).Kinetic.NernstHaskellAnion.Ion_Name = frmAddComponent!cboAnion.List(i) Then
              frmAddComponent!cboAnion.ListIndex = i
           End If
       Next i
       For i = 0 To frmAddComponent!cboCation.ListCount - 1
           If NowProj.Anion(NumberOfIonToEdit).Kinetic.NernstHaskellCation.Ion_Name = frmAddComponent!cboCation.List(i) Then
              frmAddComponent!cboCation.ListIndex = i
           End If
       Next i

       ChangedIon = NowProj.Anion(NumberOfIonToEdit)
       
       'Generate click events on appropriate units
       ListIndex = frmAddComponent!cboAddIonUnits(0).ListIndex
       frmAddComponent!cboAddIonUnits(0).ListIndex = -1
       frmAddComponent!cboAddIonUnits(0).ListIndex = ListIndex

       ListIndex = frmAddComponent!cboAddIonUnits(1).ListIndex
       frmAddComponent!cboAddIonUnits(1).ListIndex = -1
       frmAddComponent!cboAddIonUnits(1).ListIndex = ListIndex

       AddingCation = False
       AddingAnion = False
       EditingCation = False
       EditingAnion = True
       NumberOfIons = NowProj.NumberOfAnions
            If NumberOfIons > 1 Then
               frmAddComponent!cmdViewSeparationFactors.Enabled = True
            Else
               frmAddComponent!cmdViewSeparationFactors.Enabled = False
            End If

            For i = 1 To NumberOfIons
                OneDimSeparationFactors(i) = NowProj.Anion(i).SeparationFactor
            Next i
            
       frmAddComponent.Show 1

             NowProj.AnionSeparationFactorInput.Row = SeparationFactorInput.Row
             NowProj.AnionSeparationFactorInput.Value = SeparationFactorInput.Value

       EditingAnion = False

    End If

    frmAddComponent!txtAddIon(0).Enabled = True

End Sub

Private Sub cmdInputKineticParameters_Click()
    Dim i As Integer, ListIndex As Integer

'       frmInputKineticParameters!cboIon.ListIndex = -1
'       frmInputKineticParameters!cboIon.ListIndex = cboKinDimComponent.ListIndex

    If cmdAddDeleteIons(0).Enabled And cmdAddDeleteIons(2).Enabled Then   'Both Cations and Anions can be modified
       EditingCation = True
       EditingAnion = True
       For i = 1 To NowProj.NumberOfCations
           OldCationKineticParameters(i) = NowProj.Cation(i).Kinetic
       Next i

       For i = 1 To NowProj.NumberOfAnions
           OldAnionKineticParameters(i) = NowProj.Anion(i).Kinetic
       Next i
    ElseIf cmdAddDeleteIons(0).Enabled Then   'Only cations can be modified
       EditingCation = True
       EditingAnion = False

       For i = 1 To NowProj.NumberOfCations
           OldCationKineticParameters(i) = NowProj.Cation(i).Kinetic
       Next i

    ElseIf cmdAddDeleteIons(2).Enabled Then   'Only anions can be modified
       EditingCation = True
       EditingAnion = False

       For i = 1 To NowProj.NumberOfAnions
           OldAnionKineticParameters(i) = NowProj.Anion(i).Kinetic
       Next i

    End If

    ViewingKineticParametersForm = True

    ListIndex = frmInputKineticParameters!cboIon.ListIndex
    frmInputKineticParameters!cboIon.ListIndex = -1
    frmInputKineticParameters!cboIon.ListIndex = ListIndex

    frmInputKineticParameters.Show 1  'Modal
    ViewingKineticParametersForm = False

    EditingAnion = False
    EditingCation = False

End Sub

Private Sub Form_Load()

    Screen.MousePointer = 11
    'Set paths for program
   IonExchangePath = MAIN_APP_PATH
'    IonExchangePath = "d:\nasa\pfpdm\vbasic"
'    IonExchangePath = "c:\nasa\vbasic"
'    IonExchangePath = "i:\research\projects\nasa\models\mfb\pfpdm\vbasic_3"
'    ChDrive IonExchangePath
'    ChDir IonExchangePath
'    SaveAndLoadPath = CurDir$ & "\examples"

    frmIonExchangeMain.WindowState = 0
    frmIonExchangeMain.width = SCREEN_WIDTH_STANDARD
    frmIonExchangeMain.height = SCREEN_HEIGHT_STANDARD

    'Center the form on the screen
    If WindowState = 0 Then
       'don't attempt if screen Minimized or Maximized
       Move (Screen.width - frmIonExchangeMain.width) / 2, (Screen.height - frmIonExchangeMain.height) / 2
    End If

    frmIonExchangeMain.Caption = "Ion Exchange Simulation Software - untitled.iex"
    OldFileName$ = "untitled.iex"
    filename$ = ""
    
    Screen.MousePointer = 1
'    replaced with PopulateUnits code in refresh
'    Call LoadUnitsOperatingConditions
'    Call LoadUnitsBedData
'    Call LoadUnitsAdsorbentProperties
'    Call LoadUnitsAddIon
'    Call LoadUnitsKineticParameters
'    Call LoadUnitsTimeParameters
 
    NowProj.OKToGetCationDimensionless = False
    NowProj.OKToGetAnionDimensionless = False
    
'    Call InitializeAvailableIons
    Call InitializeIonExchangeParameters
    Call LoadNernstHaskellDatabases
    Call InitializeDefaultIonProperties
    Call InitializeSeparationFactorInfo
    Call InitializeTimeAndCollocationInfo
    ViewingKineticParametersForm = False
    NowProj.NumberOfCations = 0
    NowProj.NumberOfAnions = 0
    NowProj.VarInfluentFileCation = "NONE"
    NowProj.VarInfluentFileAnion = "NONE"
    frmIonExchangeMain!fraKineticDimensionless.Enabled = False
    frmIonExchangeMain!cboIons(2).Clear
    ClickGeneratedFromcboIon = False
    NumSelectedCations = 0
    NumSelectedAnions = 0
    'POPULATE OLD FILE LIST.
  
    Call OldFileList_Populate( _
        1, _
        frmIonExchangeMain.mnuFile(12), _
        frmIonExchangeMain.mnuFile(8), _
        frmIonExchangeMain.mnuFile(9), _
        frmIonExchangeMain.mnuFile(10), _
        frmIonExchangeMain.mnuFile(11))
        
    mnuRun(0).Enabled = False
    mnuResults(0).Enabled = False
    mnuResults(1).Enabled = False

    'Load forms
    Load frmInputKineticParameters
    Load frmAddComponent
    Load frmSeparationFactors
    Load frmOptionsInputParameters
    
    'temporarily enabling these menu options -- cannot run because
    ' DLL's aren't 32 bit
    frmIonExchangeMain!mnuResults(0).Enabled = True
    frmIonExchangeMain!mnuResults(1).Enabled = True
    '
'    Call file_new
    Call frmIonExchangeMain_Refresh
    Me.sspanel_Dirty = "Data Unchanged"
    Me.sspanel_Status = ""
    frmIonExchangeMain.Show
    

End Sub


Private Sub Form_Unload(Cancel As Integer)
    Unload frmInputKineticParameters
    Unload frmAddComponent
    Unload frmSeparationFactors
    End
End Sub

Private Sub GetAndDisplayEBCT()
    Dim CurrentUnits As Integer, ValueToDisplay As Double

    NowProj.Bed.EBCT.Value = EBCT(NowProj.Bed.length, NowProj.Bed.Diameter, NowProj.Bed.Flowrate.Value)
    CurrentUnits = cboBedDataUnits(4).ListIndex
    If CurrentUnits = 0 Then
       ValueToDisplay = NowProj.Bed.EBCT.Value
    Else
       ValueToDisplay = NowProj.Bed.EBCT.Value * TimeConversionFactor(CurrentUnits)
    End If
    txtBedData(4).Text = Format$(ValueToDisplay, GetTheFormat(ValueToDisplay))

End Sub

Private Sub GetAndDisplayFlowrate()
    Dim CurrentUnits As Integer, ValueToDisplay As Double

    NowProj.Bed.Flowrate.Value = Flowrate(NowProj.Bed.length, NowProj.Bed.Diameter, NowProj.Bed.EBCT.Value)
    CurrentUnits = cboBedDataUnits(3).ListIndex
    If CurrentUnits = 0 Then
       ValueToDisplay = NowProj.Bed.Flowrate.Value
    Else
       ValueToDisplay = NowProj.Bed.Flowrate.Value * FlowConversionFactor(CurrentUnits)
    End If
    txtBedData(3).Text = Format$(ValueToDisplay, GetTheFormat(ValueToDisplay))

End Sub




Private Sub lstIons_Click(Index As Integer)
    Dim i As Integer, j As Integer
    
    Select Case Index
       Case 0   'Cations
          EditingCation = True

          NumSelectedCations = 1
          Cations_Selected(1) = NowProj.PresaturantCation
  
          For i = 1 To lstIons(0).ListCount
              If lstIons(0).Selected(i - 1) Then
                 For j = 1 To NowProj.NumberOfCations
                     If Trim$(NowProj.Cation(j).Name) = Trim$(lstIons(0).List(i - 1)) Then
                        NumSelectedCations = NumSelectedCations + 1
                        Cations_Selected(NumSelectedCations) = j
                        Exit For
                     End If
                 Next j
              End If
          Next i

          If NumSelectedCations < 2 Then
             mnuRun(0).Enabled = False
             mnuOptions(4).Enabled = False
             cboKinDimComponent.Clear
             cboKinDimComponent.Enabled = False
             For i = 0 To 8
                 lblKineticDimensionlessValue(i).Caption = ""
             Next i
          Else
             mnuRun(0).Enabled = True
             mnuOptions(4).Enabled = True
             fraKineticDimensionless.Enabled = True
             cboKinDimComponent.Clear
             For i = 1 To NumSelectedCations
                 cboKinDimComponent.AddItem Trim$(NowProj.Cation(Cations_Selected(i)).Name)
             Next i
             cboKinDimComponent.Enabled = True

             Call CalculateSumEquivInitialConc
             For i = 1 To NumSelectedCations
                 NumberOfIonToEdit = Cations_Selected(i)
                 Call CalculateDimensionlessGroups
             Next i

'             cboKinDimComponent.ListIndex = -1
             cboKinDimComponent.ListIndex = 0

             'Set Presaturant back to 100 % of initial resin phase concentration
             For i = 1 To MAX_CHEMICAL
                 If i = NowProj.PresaturantCation Then
                    NowProj.Resin.PresaturantPercentage(i) = 100#
                 Else
                    NowProj.Resin.PresaturantPercentage(i) = 0#
                 End If
             Next i

             EditingCation = False
          End If

          Number_Component = NumSelectedCations

       Case 1   'Anions
          EditingAnion = True

          NumSelectedAnions = 1
          Anions_Selected(1) = NowProj.PresaturantAnion
  
          For i = 1 To lstIons(1).ListCount
              If lstIons(1).Selected(i - 1) Then
                 For j = 1 To NowProj.NumberOfAnions
                     If Trim$(NowProj.Anion(j).Name) = Trim$(lstIons(1).List(i - 1)) Then
                        NumSelectedAnions = NumSelectedAnions + 1
                        Anions_Selected(NumSelectedAnions) = j
                        Exit For
                     End If
                 Next j
              End If
          Next i

          If NumSelectedAnions < 2 Then
             mnuRun(0).Enabled = False
             mnuOptions(4).Enabled = False
             cboKinDimComponent.Clear
             cboKinDimComponent.Enabled = False
             For i = 0 To 8
                 lblKineticDimensionlessValue(i).Caption = ""
             Next i
          Else
             mnuRun(0).Enabled = True
             mnuOptions(4).Enabled = True
             fraKineticDimensionless.Enabled = True
             cboKinDimComponent.Clear
             For i = 1 To NumSelectedAnions
                 cboKinDimComponent.AddItem Trim$(NowProj.Anion(Anions_Selected(i)).Name)
             Next i
             cboKinDimComponent.Enabled = True

             Call CalculateSumEquivInitialConc
             For i = 1 To NumSelectedAnions
                 NumberOfIonToEdit = Anions_Selected(i)
                 Call CalculateDimensionlessGroups
             Next i

'             cboKinDimComponent.ListIndex = -1
             cboKinDimComponent.ListIndex = 0

             'Set Presaturant back to 100 % of initial resin phase concentration
             For i = 1 To MAX_CHEMICAL
                 If i = NowProj.PresaturantAnion Then
                    NowProj.Resin.PresaturantPercentage(i) = 100#
                 Else
                    NowProj.Resin.PresaturantPercentage(i) = 0#
                 End If
             Next i

             EditingAnion = False
          End If

          Number_Component = NumSelectedAnions

    End Select

End Sub

Private Sub mnuFile_Click(Index As Integer)
Dim frmPrint_DO_INPUTS As Boolean
Dim frmPrint_DO_OUTPUTS As Boolean
Dim frmPrint_DO_PLOTS As Boolean

    Select Case Index
       Case 0   'New
          If (file_query_unload()) Then
            Call file_new
          End If
       Case 1   'Open
'            ChDrive SaveAndLoadPath
'            ChDir SaveAndLoadPath
'            Call LoadIonExchange("")
'            SaveAndLoadPath = CurDir$
''''''            If (file_query_unload()) Then
''''''              Call File_OpenAs("")
''''''            End If
            If (file_query_unload()) Then
              Call File_OpenAs("")
            End If
            txtOperatingConditions(0).SetFocus
'            ChDrive IonExchangePath
'            ChDir IonExchangePath
          Case 2:      'Save
            If (Current_Filename = "") Then
              Call File_SaveAs("")
            Else
              Call File_Save(Current_Filename)
            End If
          Case 3:      'Save As ...
            Call File_SaveAs("")
          Case 6:      'Print inputs to Excel
            frmPrint_DO_INPUTS = True
            frmPrint_DO_OUTPUTS = True
            frmPrint_DO_PLOTS = False
'            frmPrint.Show 1
'       Case 2   'Save
'          ChDrive SaveAndLoadPath
'          ChDir SaveAndLoadPath
'          Call SaveIonExchange
'          SaveAndLoadPath = CurDir$
'          ChDir IonExchangePath
'          ChDrive IonExchangePath

'       Case 3   'Save As
'          ChDrive SaveAndLoadPath
'          ChDir SaveAndLoadPath
'          OldFileName$ = filename$
'          filename$ = ""
'          Call SaveIonExchange
'          SaveAndLoadPath = CurDir$
'          ChDir IonExchangePath
'          ChDrive IonExchangePath

       Case 5   'Select Printer
            On Error GoTo PrinterError
               CMDialog1.flags = PD_PRINTSETUP
'------Begin Modification Hokanson: 12-Aug2000
               CMDialog1.CancelError = True
'------End Modification Hokanson: 11-Aug2000
               CMDialog1.Action = 5
               Exit Sub
PrinterError:
               Resume ExitSelectPrint:

ExitSelectPrint:

       Case 6   'Print
       
       Case 8 To 12:      'Last-few-files list
          If (file_query_unload()) Then
            If (mnuFile(Index).visible) Then
'              Call LoadIonExchange(OldFiles(1, index - 7))
              Call File_OpenAs(OldFiles(1, Index - 7))
            End If
          End If

       Case 13   'Exit
            ChDir MAIN_APP_PATH
            Unload Me

    End Select

End Sub

Private Sub mnuFilePrint_Click(Index As Integer)

    Select Case Index
       Case 0   'Print to printer
          Call PrintIonExchange
       Case 1   'Print to file
          Call PrintIonExchangeToFile
    End Select

End Sub

Private Sub mnuOptions_Click(Index As Integer)
    Dim i As Integer, j As Integer
    
    Select Case Index
       Case 0   'Set Variable Influent Concentrations
          If Cations.Available And Anions.Available Then

          ElseIf Cations.Available Then
             For i = 1 To NowProj.NumberOfCations
                 Ion(i) = NowProj.Cation(i)
             Next i
              Total_NumberOfComponents = NowProj.NumberOfCations
             Call ReadVarInfluentConcs

          ElseIf Anions.Available Then
             For i = 1 To NowProj.NumberOfAnions
                 Ion(i) = NowProj.Anion(i)
             Next i
              Total_NumberOfComponents = NowProj.NumberOfAnions


          End If
             
          frmConcentrations.Show 1
       Case 1   'Set Number Of Beds
          frmOptionsInputParameters.Show 1
       Case 2   'Set Collocation Points
          frmOptionsInputParameters.Show 1
       Case 3   'Set Time Parameters
          frmOptionsInputParameters.Show 1
       Case 4   'Initial Resin Phase Concentrations
          frmResinPresaturantConditions.Show 1
'------Begin Modification Hokanson: 11-Aug2000
       Case 5   'Set EPS and DH0
          frmOptionsInputParameters.Show 1
'------End Modification Hokanson: 11-Aug2000
    End Select

End Sub

Private Sub mnuResults_Click(Index As Integer)
    Dim i As Integer, j As Integer

    Select Case Index
       Case 0   'PFPDM results
          frmbreak.Show 1
       Case 1   'Compare to Data
          frmPlantData.Show 1
    End Select
End Sub

Private Sub mnuRun_Click(Index As Integer)
    Dim i As Integer, j As Integer

    Select Case Index
       Case 0   'PFPDM
'          Call GetSelectedComponents(0)

          If Cations.Available And Anions.Available Then

          ElseIf Cations.Available Then
             NumSelectedComponents_PFPDM = NumSelectedCations
             'Place Presaturant Ion in Last Element of Array
             For i = 2 To NumSelectedCations
                 Component_Index_PFPDM(i - 1) = Cations_Selected(i)
             Next i
             Component_Index_PFPDM(NumSelectedCations) = Cations_Selected(1)
             For i = 1 To NowProj.NumberOfCations
                 Ion(i) = NowProj.Cation(i)
             Next i

             'Determine Alpha_Input Array to send to PFPDM
             For i = 1 To NowProj.NumberOfCations
                 OneDimSeparationFactors(i) = NowProj.Cation(i).SeparationFactor
             Next i
             SeparationFactorInput.Row = NowProj.CationSeparationFactorInput.Row
             SeparationFactorInput.Value = NowProj.CationSeparationFactorInput.Value
             NumberOfIons = NowProj.NumberOfCations
             Call CalculateSeparationFactors
             For i = 1 To NumSelectedCations
                 j = Cations_Selected(i)
                 AlphaInput(i) = TwoDimSeparationFactors(j, 1)
             Next i

             Call Call_PFPDM

          ElseIf Anions.Available Then

             NumSelectedComponents_PFPDM = NumSelectedAnions
             'Place Presaturant Ion in Last Element of Array
             For i = 2 To NumSelectedAnions
                 Component_Index_PFPDM(i - 1) = Anions_Selected(i)
             Next i
             Component_Index_PFPDM(NumSelectedAnions) = Anions_Selected(1)
             For i = 1 To NowProj.NumberOfAnions
                 Ion(i) = NowProj.Anion(i)
             Next i

             'Determine Alpha_Input Array to send to PFPDM
             For i = 1 To NowProj.NumberOfAnions
                 OneDimSeparationFactors(i) = NowProj.Anion(i).SeparationFactor
             Next i
             SeparationFactorInput.Row = NowProj.AnionSeparationFactorInput.Row
             SeparationFactorInput.Value = NowProj.AnionSeparationFactorInput.Value
             NumberOfIons = NowProj.NumberOfAnions
             Call CalculateSeparationFactors
             For i = 1 To NumSelectedAnions
                 j = Anions_Selected(i)
                 AlphaInput(i) = TwoDimSeparationFactors(j, 1)
             Next i

             Call Call_PFPDM

          End If

    End Select

End Sub

Private Sub txtAdsorbentProperties_GotFocus(Index As Integer)
    Call TextGetFocus(txtAdsorbentProperties(Index), Temp_Text)
End Sub

Private Sub txtAdsorbentProperties_KeyPress(Index As Integer, KeyAscii As Integer)
    
    If KeyAscii = 13 Then
       KeyAscii = 0
       SendKeys "{Tab}"
       Exit Sub
    End If

    If Index = 0 Then Exit Sub

    Call NumberCheck(KeyAscii)

End Sub

Private Sub txtAdsorbentProperties_LostFocus(Index As Integer)
    Dim ValueChanged As Integer, NewValue As Double
    Dim msg As String, CurrentUnits As Integer
    Dim OldValue As Double, ValueToDisplay As Double
    Dim i As Integer


    Call TextHandleError(IsError, txtAdsorbentProperties(Index), Temp_Text)

    If Not IsError Then
       Is_Dirty = True
       NewValue = CDbl(txtAdsorbentProperties(Index).Text)
       'Convert NewValue to Standard Units if Necessary
       Select Case Index
          Case 1   'Apparent Density
               OldValue = NowProj.Resin.ApparentDensity
               CurrentUnits = cboAdsorbentPropertyUnits(1).ListIndex
               If CurrentUnits <> 0 Then
                  NewValue = NewValue / DensityConversionFactor(CurrentUnits)
               End If
          Case 2   'Particle Radius
               OldValue = NowProj.Resin.ParticleRadius
               CurrentUnits = cboAdsorbentPropertyUnits(2).ListIndex
               If CurrentUnits <> 0 Then
                  NewValue = NewValue / LengthConversionFactor(CurrentUnits)
               End If
          Case 3   'Particle Porosity
               OldValue = NowProj.Resin.ParticlePorosity
          Case 4   'Tortuosity
               OldValue = NowProj.Resin.Tortuosity
          Case 5   'Total Resin Capacity
               OldValue = NowProj.Resin.TotalCapacity
               CurrentUnits = cboAdsorbentPropertyUnits(5).ListIndex
               If CurrentUnits <> 0 Then
                  NewValue = NewValue / ResinCapacityConversionFactor(CurrentUnits)
               End If
       End Select

       Select Case Index
          Case 1    'Apparent Density
             Call NumberChanged(ValueChanged, OldValue, NewValue)
             If ValueChanged Then
                If HaveValue(NewValue) Then
                   NowProj.Resin.ApparentDensity = NewValue
                   
                   Call CalculateBedPorosity
                   Call CalculateEffectiveContactTime
                   Call CalculateInterstitialVelocity
                   
                   Call UpdateKineticParametersAllIons
                   Call UpdateDimensionlessGroupAllIons

                Else
                   txtAdsorbentProperties(1).Text = Temp_Text
                   txtAdsorbentProperties(1).SetFocus
                   Exit Sub
                End If
             End If

          Case 2    'Particle Radius
             Call NumberChanged(ValueChanged, OldValue, NewValue)
             If ValueChanged Then
                If HaveValue(NewValue) Then
                   NowProj.Resin.ParticleRadius = NewValue
                   Call CalculateParticleDiameter

                   Call UpdateKineticParametersAllIons
                   Call UpdateDimensionlessGroupAllIons

                Else
                   txtAdsorbentProperties(2).Text = Temp_Text
                   txtAdsorbentProperties(2).SetFocus
                   Exit Sub
                End If
             End If

          Case 3    'Particle Porosity
             Call NumberChanged(ValueChanged, OldValue, NewValue)
             If ValueChanged Then
                If HaveValue(NewValue) Then
                   NowProj.Resin.ParticlePorosity = NewValue

                   Call UpdateDimensionlessGroupAllIons

                Else
                   txtAdsorbentProperties(3).Text = Temp_Text
                   txtAdsorbentProperties(3).SetFocus
                   Exit Sub
                End If
             End If
             
          Case 4    'Tortuosity
             Call NumberChanged(ValueChanged, OldValue, NewValue)
             If ValueChanged Then
                If HaveValue(NewValue) Then
                   NowProj.Resin.Tortuosity = NewValue

                   Call UpdateKineticParametersAllIons
                   Call UpdateDimensionlessGroupAllIons

                Else
                   txtAdsorbentProperties(4).Text = Temp_Text
                   txtAdsorbentProperties(4).SetFocus
                   Exit Sub
                End If
             End If

          Case 5    'Resin Capacity
             Call NumberChanged(ValueChanged, OldValue, NewValue)
             If ValueChanged Then
                If HaveValue(NewValue) Then
                   NowProj.Resin.TotalCapacity = NewValue

                   Call UpdateDimensionlessGroupAllIons

                Else
                   txtAdsorbentProperties(5).Text = Temp_Text
                   txtAdsorbentProperties(5).SetFocus
                   Exit Sub
                End If
             End If


       End Select

    End If
    Call Local_DirtyStatus_Set( _
          Is_Dirty, True)
    Call frmIonExchangeMain_Refresh
    
End Sub

Private Sub txtBedData_GotFocus(Index As Integer)
    Call TextGetFocus(txtBedData(Index), Temp_Text)
End Sub

Private Sub txtBedData_KeyPress(Index As Integer, KeyAscii As Integer)

    If KeyAscii = 13 Then
       KeyAscii = 0
       SendKeys "{Tab}"
       Exit Sub
    End If

    Call NumberCheck(KeyAscii)

End Sub

Private Sub txtBedData_LostFocus(Index As Integer)
    Dim ValueChanged As Integer, NewValue As Double
    Dim msg As String, CurrentUnits As Integer
    Dim OldValue As Double, ValueToDisplay As Double
    Dim i As Integer

    Call TextHandleError(IsError, txtBedData(Index), Temp_Text)

    If Not IsError Then
       Is_Dirty = True
       NewValue = CDbl(txtBedData(Index).Text)
       'Convert NewValue to Standard Units if Necessary
       Select Case Index
          Case 0   'Bed Length
               OldValue = NowProj.Bed.length
               CurrentUnits = cboBedDataUnits(0).ListIndex
               If CurrentUnits <> 0 Then
                  NewValue = NewValue / LengthConversionFactor(CurrentUnits)
               End If
          Case 1   'Bed Diameter
               OldValue = NowProj.Bed.Diameter
               CurrentUnits = cboBedDataUnits(1).ListIndex
               If CurrentUnits <> 0 Then
                  NewValue = NewValue / LengthConversionFactor(CurrentUnits)
               End If
          Case 2   'Bed Mass
               OldValue = NowProj.Bed.Weight
               CurrentUnits = cboBedDataUnits(2).ListIndex
               If CurrentUnits <> 0 Then
                  NewValue = NewValue / MassConversionFactor(CurrentUnits)
               End If
          Case 3   'Flowrate
               OldValue = NowProj.Bed.Flowrate.Value
               CurrentUnits = cboBedDataUnits(3).ListIndex
               If CurrentUnits <> 0 Then
                  NewValue = NewValue / FlowConversionFactor(CurrentUnits)
               End If
          Case 4   'EBCT
               OldValue = NowProj.Bed.EBCT.Value
               CurrentUnits = cboBedDataUnits(4).ListIndex
               If CurrentUnits <> 0 Then
                  NewValue = NewValue / TimeConversionFactor(CurrentUnits)
               End If
       End Select

       Select Case Index
          Case 0    'Bed Length
             Call NumberChanged(ValueChanged, OldValue, NewValue)
             If ValueChanged Then
                If HaveValue(NewValue) Then
                   NowProj.Bed.length = NewValue
                   If NowProj.Bed.Flowrate.UserInput Then
                      Call GetAndDisplayEBCT
                   Else
                      Call GetAndDisplayFlowrate
                   End If

                   Call CalculateBedVolume
                   Call CalculateBedDensity
                   Call CalculateBedPorosity
                   Call CalculateEffectiveContactTime
                   Call CalculateInterstitialVelocity
                   
                   Call UpdateKineticParametersAllIons
                   Call UpdateDimensionlessGroupAllIons

                Else
                   txtBedData(0).Text = Temp_Text
                   txtBedData(0).SetFocus
                   Exit Sub
                End If
             End If

          Case 1    'Bed Diameter
             Call NumberChanged(ValueChanged, OldValue, NewValue)
             If ValueChanged Then
                If HaveValue(NewValue) Then
                   NowProj.Bed.Diameter = NewValue
                   If NowProj.Bed.Flowrate.UserInput Then
                      Call GetAndDisplayEBCT
                   Else
                      Call GetAndDisplayFlowrate
                   End If

                   Call CalculateBedArea
                   Call CalculateBedVolume
                   Call CalculateBedDensity
                   Call CalculateBedPorosity
                   Call CalculateEffectiveContactTime
                   Call CalculateSuperficialVelocity
                   Call CalculateInterstitialVelocity

                   Call UpdateKineticParametersAllIons
                   Call UpdateDimensionlessGroupAllIons

                Else
                   txtBedData(1).Text = Temp_Text
                   txtBedData(1).SetFocus
                   Exit Sub
                End If
             End If

          Case 2    'Bed Mass
             Call NumberChanged(ValueChanged, OldValue, NewValue)
             If ValueChanged Then
                If HaveValue(NewValue) Then
                   NowProj.Bed.Weight = NewValue
                   Call CalculateBedDensity
                   Call CalculateBedPorosity
                   Call CalculateInterstitialVelocity

                   Call CalculateKineticParameters
                   Call UpdateDimensionlessGroupAllIons

                Else
                   txtBedData(2).Text = Temp_Text
                   txtBedData(2).SetFocus
                   Exit Sub
                End If
             End If

             
          Case 3    'Flowrate
             Call NumberChanged(ValueChanged, OldValue, NewValue)
             If ValueChanged Then
                If HaveValue(NewValue) Then
                   NowProj.Bed.Flowrate.Value = NewValue
                   NowProj.Bed.Flowrate.UserInput = True
                   NowProj.Bed.EBCT.UserInput = False
                   Call GetAndDisplayEBCT
                   Call CalculateEffectiveContactTime
                   Call CalculateSuperficialVelocity
                   Call CalculateInterstitialVelocity
                   
                   Call UpdateKineticParametersAllIons
                   Call UpdateDimensionlessGroupAllIons
                      
                Else
                   txtBedData(3).Text = Temp_Text
                   txtBedData(3).SetFocus
                   Exit Sub
                End If
             End If

          Case 4    'EBCT
             Call NumberChanged(ValueChanged, OldValue, NewValue)
             If ValueChanged Then
                If HaveValue(NewValue) Then
                   NowProj.Bed.EBCT.Value = NewValue
                   NowProj.Bed.EBCT.UserInput = True
                   NowProj.Bed.Flowrate.UserInput = False
                   Call GetAndDisplayFlowrate
                   Call CalculateEffectiveContactTime
                   Call CalculateSuperficialVelocity
                   Call CalculateInterstitialVelocity
                  
                   Call UpdateKineticParametersAllIons
                   Call UpdateDimensionlessGroupAllIons

                Else
                   txtBedData(4).Text = Temp_Text
                   txtBedData(4).SetFocus
                   Exit Sub
                End If
             End If


       End Select

    End If
    Call Local_DirtyStatus_Set( _
          Is_Dirty, True)
    Call frmIonExchangeMain_Refresh
End Sub

Private Sub txtOperatingConditions_GotFocus(Index As Integer)
    Call TextGetFocus(txtOperatingConditions(Index), Temp_Text)
End Sub

Private Sub txtOperatingConditions_KeyPress(Index As Integer, KeyAscii As Integer)
    
    If KeyAscii = 13 Then
       KeyAscii = 0
       SendKeys "{Tab}"
       Exit Sub
    End If

    Call NumberCheck(KeyAscii)

End Sub

Private Sub txtOperatingConditions_LostFocus(Index As Integer)
    Dim ValueChanged As Integer, NewValue As Double
    Dim msg As String, CurrentUnits As Integer
    Dim OldValue As Double, ValueToDisplay As Double
    Dim i As Integer

    
    Call TextHandleError(IsError, txtOperatingConditions(Index), Temp_Text)
    
    If Not IsError Then
       NewValue = CDbl(txtOperatingConditions(Index).Text)
       'Convert NewValue to Standard Units if Necessary
       Select Case Index
          Case 0   'Operating Pressure
               OldValue = NowProj.Operating.Pressure
               CurrentUnits = cboOperatingConditionsUnits(0).ListIndex
               If CurrentUnits <> 0 Then
                  NewValue = NewValue / PressureConversionFactor(CurrentUnits)
               End If
          Case 1   'Operating Temperature
               OldValue = NowProj.Operating.Temperature
               CurrentUnits = cboOperatingConditionsUnits(1).ListIndex
               If CurrentUnits <> 0 Then
                  NewValue = ReverseTemperatureConversion(CurrentUnits, NewValue)
               End If
       End Select

       Select Case Index
          Case 0    'Operating Pressure
             Call NumberChanged(ValueChanged, OldValue, NewValue)
             If ValueChanged Then
                If HaveValue(NewValue) Then
                    NowProj.Operating.Pressure = NewValue
                Else
                   txtOperatingConditions(0).Text = Temp_Text
                   txtOperatingConditions(0).SetFocus
                   Exit Sub
                End If
             End If

          Case 1    'Operating Temperature
             Call NumberChanged(ValueChanged, OldValue, NewValue)
             If ValueChanged Then
                If HaveValue(NewValue) Then
                    NowProj.Operating.Temperature = NewValue
                   Call CalculateLiquidDensity
                   Call CalculateLiquidViscosity

                Call UpdateKineticParametersAllIons
                Call UpdateDimensionlessGroupAllIons

                Else
                   txtOperatingConditions(1).Text = Temp_Text
                   txtOperatingConditions(1).SetFocus
                   Exit Sub
                End If
             End If

       End Select

    End If
    
    If (Is_Dirty) Then
      'THROW DIRTY FLAG.
      Call Local_DirtyStatus_Set( _
          Is_Dirty, True)
    End If
    'REFRESH WINDOW.
    Call frmIonExchangeMain_Refresh
End Sub



Sub Local_DirtyStatus_Set(DirtyFlag As Boolean, NewSetting As Boolean)
  Call Global_DirtyStatus_Set(Me, DirtyFlag, NewSetting)
End Sub

