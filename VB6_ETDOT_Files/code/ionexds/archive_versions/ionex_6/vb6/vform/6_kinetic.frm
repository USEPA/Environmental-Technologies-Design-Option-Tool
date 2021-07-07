VERSION 5.00
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "THREED32.OCX"
Begin VB.Form frmInputKineticParameters 
   Appearance      =   0  'Flat
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Kinetic Parameters"
   ClientHeight    =   6480
   ClientLeft      =   1365
   ClientTop       =   2175
   ClientWidth     =   8670
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
   ScaleHeight     =   6480
   ScaleWidth      =   8670
   Begin Threed.SSFrame fraPoreDiffusivityInput 
      Height          =   1095
      Left            =   4440
      TabIndex        =   101
      Top             =   5400
      Width           =   4215
      _Version        =   65536
      _ExtentX        =   7435
      _ExtentY        =   1931
      _StockProps     =   14
      Caption         =   "In Model, Use Pore Diffusivity from:"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Begin VB.TextBox txtPoreDiffusivityCorr 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         Height          =   288
         Left            =   1800
         TabIndex        =   113
         Top             =   360
         Width           =   972
      End
      Begin VB.TextBox txtPoreDiffusivityUser 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         Height          =   288
         Left            =   1800
         TabIndex        =   104
         Top             =   670
         Width           =   972
      End
      Begin VB.ComboBox cboPoreDiffusivityUnits 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         Height          =   315
         Index           =   0
         Left            =   2880
         Style           =   2  'Dropdown List
         TabIndex        =   103
         TabStop         =   0   'False
         Top             =   300
         Width           =   1152
      End
      Begin VB.ComboBox cboPoreDiffusivityUnits 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         Height          =   315
         Index           =   1
         Left            =   2880
         Style           =   2  'Dropdown List
         TabIndex        =   102
         TabStop         =   0   'False
         Top             =   670
         Width           =   1152
      End
      Begin Threed.SSOption optPoreDiffusivity 
         Height          =   375
         Index           =   0
         Left            =   360
         TabIndex        =   106
         Top             =   240
         Width           =   1215
         _Version        =   65536
         _ExtentX        =   2143
         _ExtentY        =   661
         _StockProps     =   78
         Caption         =   "Correlation"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.24
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin Threed.SSOption optPoreDiffusivity 
         Height          =   255
         Index           =   1
         Left            =   360
         TabIndex        =   107
         Top             =   670
         Width           =   1215
         _Version        =   65536
         _ExtentX        =   2143
         _ExtentY        =   450
         _StockProps     =   78
         Caption         =   "User Input"
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
      Begin VB.Label lblPoreDiffusivityUser 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   1800
         TabIndex        =   115
         Top             =   720
         Width           =   975
      End
      Begin VB.Label lblPoreDiffusivityCorr 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   1800
         TabIndex        =   105
         Top             =   360
         Width           =   975
      End
   End
   Begin Threed.SSFrame fraPoreDiffusivity 
      Height          =   975
      Left            =   4440
      TabIndex        =   91
      Top             =   4440
      Width           =   4215
      _Version        =   65536
      _ExtentX        =   7435
      _ExtentY        =   1720
      _StockProps     =   14
      Caption         =   "Pore Diffusivity, Dp"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Begin VB.Label lblPoreDiffusivityUnits 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "cm2/s"
         ForeColor       =   &H80000008&
         Height          =   195
         Index           =   2
         Left            =   3060
         TabIndex        =   100
         Top             =   660
         Width           =   975
      End
      Begin VB.Label lblPoreDiffusivityValue 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   195
         Index           =   2
         Left            =   1920
         TabIndex        =   99
         Top             =   660
         Width           =   975
      End
      Begin VB.Label lblPoreDiffusivity 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Pore Diffusivity"
         ForeColor       =   &H80000008&
         Height          =   195
         Index           =   2
         Left            =   240
         TabIndex        =   98
         Top             =   660
         Width           =   1545
      End
      Begin VB.Label lblPoreDiffusivity 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Liquid Diffusivity"
         ForeColor       =   &H80000008&
         Height          =   195
         Index           =   0
         Left            =   240
         TabIndex        =   97
         Top             =   240
         Width           =   1545
      End
      Begin VB.Label lblPoreDiffusivityValue 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   195
         Index           =   0
         Left            =   1920
         TabIndex        =   96
         Top             =   240
         Width           =   975
      End
      Begin VB.Label lblPoreDiffusivityUnits 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "cm2/s"
         ForeColor       =   &H80000008&
         Height          =   195
         Index           =   0
         Left            =   3060
         TabIndex        =   95
         Top             =   240
         Width           =   975
      End
      Begin VB.Label lblPoreDiffusivity 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Tortuosity"
         ForeColor       =   &H80000008&
         Height          =   195
         Index           =   1
         Left            =   240
         TabIndex        =   94
         Top             =   450
         Width           =   1545
      End
      Begin VB.Label lblPoreDiffusivityValue 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   195
         Index           =   1
         Left            =   1920
         TabIndex        =   93
         Top             =   450
         Width           =   975
      End
      Begin VB.Label lblPoreDiffusivityUnits 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "(-)"
         ForeColor       =   &H80000008&
         Height          =   195
         Index           =   1
         Left            =   3060
         TabIndex        =   92
         Top             =   450
         Width           =   975
      End
   End
   Begin Threed.SSFrame fraIonicTransport 
      Height          =   975
      Left            =   4440
      TabIndex        =   84
      Top             =   3480
      Width           =   4215
      _Version        =   65536
      _ExtentX        =   7435
      _ExtentY        =   1720
      _StockProps     =   14
      Caption         =   "In Model, Use Ionic Transport Coefficient from:"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Begin VB.TextBox txtIonicTransportCoeffCorr 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         Height          =   288
         Left            =   1920
         TabIndex        =   112
         Top             =   275
         Width           =   972
      End
      Begin VB.ComboBox cboIonicTransportUnits 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         Height          =   315
         Index           =   1
         Left            =   3000
         Style           =   2  'Dropdown List
         TabIndex        =   87
         TabStop         =   0   'False
         Top             =   600
         Width           =   1152
      End
      Begin VB.ComboBox cboIonicTransportUnits 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         Height          =   315
         Index           =   0
         Left            =   3000
         Style           =   2  'Dropdown List
         TabIndex        =   86
         TabStop         =   0   'False
         Top             =   240
         Width           =   1152
      End
      Begin VB.TextBox txtIonicTransCoeffUser 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         Height          =   288
         Left            =   1920
         TabIndex        =   85
         Top             =   600
         Width           =   972
      End
      Begin Threed.SSOption optIonicTransportCoeff 
         Height          =   375
         Index           =   0
         Left            =   240
         TabIndex        =   89
         Top             =   240
         Width           =   1215
         _Version        =   65536
         _ExtentX        =   2143
         _ExtentY        =   661
         _StockProps     =   78
         Caption         =   "Correlation"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.24
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin Threed.SSOption optIonicTransportCoeff 
         Height          =   255
         Index           =   1
         Left            =   240
         TabIndex        =   90
         Top             =   600
         Width           =   1215
         _Version        =   65536
         _ExtentX        =   2143
         _ExtentY        =   450
         _StockProps     =   78
         Caption         =   "User Input"
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
      Begin VB.Label lblIonicTransCoeffUser 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   1920
         TabIndex        =   114
         Top             =   600
         Width           =   975
      End
      Begin VB.Label lblIonicTransportCoeffCorr 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   1920
         TabIndex        =   88
         Top             =   300
         Width           =   975
      End
   End
   Begin Threed.SSFrame fraIonicTransportCoefficient 
      Height          =   3495
      Left            =   4440
      TabIndex        =   39
      Top             =   0
      Width           =   4215
      _Version        =   65536
      _ExtentX        =   7435
      _ExtentY        =   6165
      _StockProps     =   14
      Caption         =   "Ionic Transport coefficient, kf"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Begin VB.ComboBox cboIonicTransport 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   1920
         Style           =   2  'Dropdown List
         TabIndex        =   40
         Top             =   2820
         Width           =   2175
      End
      Begin VB.Label lblIonicTransportCoeffUnits 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "cm"
         ForeColor       =   &H80000008&
         Height          =   195
         Index           =   0
         Left            =   3060
         TabIndex        =   83
         Top             =   360
         Width           =   975
      End
      Begin VB.Label lblIonicTranportCoeffValue 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   195
         Index           =   0
         Left            =   1920
         TabIndex        =   82
         Top             =   360
         Width           =   975
      End
      Begin VB.Label lblIonicTransportCoeff 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Particle Diameter"
         ForeColor       =   &H80000008&
         Height          =   195
         Index           =   0
         Left            =   240
         TabIndex        =   81
         Top             =   360
         Width           =   1545
      End
      Begin VB.Label lblIonicTransportCoeffUnits 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "cm"
         ForeColor       =   &H80000008&
         Height          =   195
         Index           =   1
         Left            =   3060
         TabIndex        =   80
         Top             =   540
         Width           =   975
      End
      Begin VB.Label lblIonicTranportCoeffValue 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   195
         Index           =   1
         Left            =   1920
         TabIndex        =   79
         Top             =   540
         Width           =   975
      End
      Begin VB.Label lblIonicTransportCoeff 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Column Diameter"
         ForeColor       =   &H80000008&
         Height          =   195
         Index           =   1
         Left            =   240
         TabIndex        =   78
         Top             =   540
         Width           =   1545
      End
      Begin VB.Label lblIonicTransportCoeffUnits 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "cm2"
         ForeColor       =   &H80000008&
         Height          =   195
         Index           =   2
         Left            =   3060
         TabIndex        =   77
         Top             =   720
         Width           =   975
      End
      Begin VB.Label lblIonicTranportCoeffValue 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   195
         Index           =   2
         Left            =   1920
         TabIndex        =   76
         Top             =   720
         Width           =   975
      End
      Begin VB.Label lblIonicTransportCoeffUnits 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "cm3/s"
         ForeColor       =   &H80000008&
         Height          =   195
         Index           =   3
         Left            =   3060
         TabIndex        =   75
         Top             =   900
         Width           =   975
      End
      Begin VB.Label lblIonicTranportCoeffValue 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   195
         Index           =   3
         Left            =   1920
         TabIndex        =   74
         Top             =   900
         Width           =   975
      End
      Begin VB.Label lblIonicTransportCoeff 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Inlet Flow Rate"
         ForeColor       =   &H80000008&
         Height          =   195
         Index           =   3
         Left            =   240
         TabIndex        =   73
         Top             =   900
         Width           =   1545
      End
      Begin VB.Label lblIonicTransportCoeffUnits 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "cm/s"
         ForeColor       =   &H80000008&
         Height          =   195
         Index           =   4
         Left            =   3060
         TabIndex        =   72
         Top             =   1080
         Width           =   975
      End
      Begin VB.Label lblIonicTranportCoeffValue 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   195
         Index           =   4
         Left            =   1920
         TabIndex        =   71
         Top             =   1080
         Width           =   975
      End
      Begin VB.Label lblIonicTransportCoeff 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Superficial Vel."
         ForeColor       =   &H80000008&
         Height          =   195
         Index           =   4
         Left            =   240
         TabIndex        =   70
         Top             =   1080
         Width           =   1545
      End
      Begin VB.Label lblIonicTransportCoeffUnits 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "(-)"
         ForeColor       =   &H80000008&
         Height          =   195
         Index           =   5
         Left            =   3060
         TabIndex        =   69
         Top             =   1260
         Width           =   975
      End
      Begin VB.Label lblIonicTranportCoeffValue 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   195
         Index           =   5
         Left            =   1920
         TabIndex        =   68
         Top             =   1260
         Width           =   975
      End
      Begin VB.Label lblIonicTransportCoeff 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Bed Porosity"
         ForeColor       =   &H80000008&
         Height          =   195
         Index           =   5
         Left            =   240
         TabIndex        =   67
         Top             =   1260
         Width           =   1545
      End
      Begin VB.Label lblIonicTransportCoeffUnits 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "cm/s"
         ForeColor       =   &H80000008&
         Height          =   195
         Index           =   6
         Left            =   3060
         TabIndex        =   66
         Top             =   1440
         Width           =   975
      End
      Begin VB.Label lblIonicTranportCoeffValue 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   195
         Index           =   6
         Left            =   1920
         TabIndex        =   65
         Top             =   1440
         Width           =   975
      End
      Begin VB.Label lblIonicTransportCoeff 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Interstitial Vel."
         ForeColor       =   &H80000008&
         Height          =   195
         Index           =   6
         Left            =   240
         TabIndex        =   64
         Top             =   1440
         Width           =   1545
      End
      Begin VB.Label lblIonicTransportCoeffUnits 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "C"
         ForeColor       =   &H80000008&
         Height          =   195
         Index           =   7
         Left            =   3060
         TabIndex        =   63
         Top             =   1620
         Width           =   975
      End
      Begin VB.Label lblIonicTranportCoeffValue 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   195
         Index           =   7
         Left            =   1920
         TabIndex        =   62
         Top             =   1620
         Width           =   975
      End
      Begin VB.Label lblIonicTransportCoeff 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Temperature"
         ForeColor       =   &H80000008&
         Height          =   195
         Index           =   7
         Left            =   240
         TabIndex        =   61
         Top             =   1620
         Width           =   1545
      End
      Begin VB.Label lblIonicTransportCoeffUnits 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "g/cm3"
         ForeColor       =   &H80000008&
         Height          =   195
         Index           =   8
         Left            =   3060
         TabIndex        =   60
         Top             =   1800
         Width           =   975
      End
      Begin VB.Label lblIonicTranportCoeffValue 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   195
         Index           =   8
         Left            =   1920
         TabIndex        =   59
         Top             =   1800
         Width           =   975
      End
      Begin VB.Label lblIonicTransportCoeff 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Liquid Density"
         ForeColor       =   &H80000008&
         Height          =   195
         Index           =   8
         Left            =   240
         TabIndex        =   58
         Top             =   1800
         Width           =   1545
      End
      Begin VB.Label lblIonicTransportCoeff 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Liquid Viscosity"
         ForeColor       =   &H80000008&
         Height          =   195
         Index           =   9
         Left            =   240
         TabIndex        =   57
         Top             =   1980
         Width           =   1545
      End
      Begin VB.Label lblIonicTranportCoeffValue 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   195
         Index           =   9
         Left            =   1920
         TabIndex        =   56
         Top             =   1980
         Width           =   975
      End
      Begin VB.Label lblIonicTransportCoeffUnits 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "g/cm/s"
         ForeColor       =   &H80000008&
         Height          =   195
         Index           =   9
         Left            =   3060
         TabIndex        =   55
         Top             =   1980
         Width           =   975
      End
      Begin VB.Label lblIonicTransportCoeff 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Reynold's No."
         ForeColor       =   &H80000008&
         Height          =   195
         Index           =   10
         Left            =   240
         TabIndex        =   54
         Top             =   2160
         Width           =   1545
      End
      Begin VB.Label lblIonicTranportCoeffValue 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   195
         Index           =   10
         Left            =   1920
         TabIndex        =   53
         Top             =   2160
         Width           =   975
      End
      Begin VB.Label lblIonicTransportCoeffUnits 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "(-)"
         ForeColor       =   &H80000008&
         Height          =   195
         Index           =   10
         Left            =   3060
         TabIndex        =   52
         Top             =   2160
         Width           =   975
      End
      Begin VB.Label lblIonicTransportCoeff 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Liquid Diffusivity"
         ForeColor       =   &H80000008&
         Height          =   195
         Index           =   11
         Left            =   240
         TabIndex        =   51
         Top             =   2340
         Width           =   1545
      End
      Begin VB.Label lblIonicTranportCoeffValue 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   195
         Index           =   11
         Left            =   1920
         TabIndex        =   50
         Top             =   2340
         Width           =   975
      End
      Begin VB.Label lblIonicTransportCoeffUnits 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "cm2/s"
         ForeColor       =   &H80000008&
         Height          =   195
         Index           =   11
         Left            =   3060
         TabIndex        =   49
         Top             =   2340
         Width           =   975
      End
      Begin VB.Label lblIonicTransportCoeff 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Schmidt No."
         ForeColor       =   &H80000008&
         Height          =   195
         Index           =   12
         Left            =   240
         TabIndex        =   48
         Top             =   2520
         Width           =   1545
      End
      Begin VB.Label lblIonicTranportCoeffValue 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   195
         Index           =   12
         Left            =   1920
         TabIndex        =   47
         Top             =   2520
         Width           =   975
      End
      Begin VB.Label lblIonicTransportCoeffUnits 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "(-)"
         ForeColor       =   &H80000008&
         Height          =   195
         Index           =   12
         Left            =   3060
         TabIndex        =   46
         Top             =   2520
         Width           =   975
      End
      Begin VB.Label lblIonicTransportCoeff 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Column Area"
         ForeColor       =   &H80000008&
         Height          =   195
         Index           =   2
         Left            =   240
         TabIndex        =   45
         Top             =   720
         Width           =   1545
      End
      Begin VB.Label lblIonicTransportCoeffUnits 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "cm/s"
         ForeColor       =   &H80000008&
         Height          =   195
         Index           =   13
         Left            =   3060
         TabIndex        =   44
         Top             =   3180
         Width           =   975
      End
      Begin VB.Label lblIonicTranportCoeffValue 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   195
         Index           =   13
         Left            =   1920
         TabIndex        =   43
         Top             =   3180
         Width           =   975
      End
      Begin VB.Label lblIonicTransportCoeff 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Ionic Trans. Coeff."
         ForeColor       =   &H80000008&
         Height          =   195
         Index           =   13
         Left            =   240
         TabIndex        =   42
         Top             =   3180
         Width           =   1545
      End
      Begin VB.Label lblIonicTransportCoeff 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Correlation Used:"
         ForeColor       =   &H80000008&
         Height          =   195
         Index           =   14
         Left            =   300
         TabIndex        =   41
         Top             =   2880
         Width           =   1545
      End
   End
   Begin Threed.SSFrame fraLiquidDiffusivityInput 
      Height          =   1215
      Left            =   120
      TabIndex        =   34
      Top             =   4200
      Width           =   4215
      _Version        =   65536
      _ExtentX        =   7435
      _ExtentY        =   2143
      _StockProps     =   14
      Caption         =   "In Model, Use Liquid diffusivity from:"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Begin VB.TextBox txtLiquidDiffCorrelation 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         Height          =   288
         Left            =   1680
         TabIndex        =   108
         Top             =   360
         Width           =   972
      End
      Begin VB.TextBox txtLiquidDiffUserInput 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         Height          =   288
         Left            =   1680
         TabIndex        =   37
         Top             =   720
         Width           =   972
      End
      Begin VB.ComboBox cboLiquidDiffusivityUnits 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         Height          =   315
         Index           =   0
         Left            =   2760
         Style           =   2  'Dropdown List
         TabIndex        =   36
         TabStop         =   0   'False
         Top             =   360
         Width           =   1152
      End
      Begin VB.ComboBox cboLiquidDiffusivityUnits 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         Height          =   315
         Index           =   1
         Left            =   2760
         Style           =   2  'Dropdown List
         TabIndex        =   35
         TabStop         =   0   'False
         Top             =   720
         Width           =   1152
      End
      Begin Threed.SSOption optLiquidDiffusivity 
         Height          =   375
         Index           =   1
         Left            =   120
         TabIndex        =   38
         Top             =   720
         Width           =   1215
         _Version        =   65536
         _ExtentX        =   2143
         _ExtentY        =   661
         _StockProps     =   78
         Caption         =   "User Input"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.24
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin Threed.SSOption optLiquidDiffusivity 
         Height          =   375
         Index           =   0
         Left            =   120
         TabIndex        =   110
         Top             =   360
         Width           =   1215
         _Version        =   65536
         _ExtentX        =   2143
         _ExtentY        =   661
         _StockProps     =   78
         Caption         =   "Correlation"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.24
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin VB.Label lblLiquidDiffUserInput 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   285
         Left            =   1800
         TabIndex        =   111
         Top             =   720
         Width           =   375
      End
      Begin VB.Label lblLiquidDiffCorrelation 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   285
         Left            =   1680
         TabIndex        =   109
         Top             =   360
         Width           =   375
      End
   End
   Begin Threed.SSFrame fraLiquidDiffusivity 
      Height          =   3255
      Left            =   120
      TabIndex        =   7
      Top             =   960
      Width           =   4215
      _Version        =   65536
      _ExtentX        =   7435
      _ExtentY        =   5741
      _StockProps     =   14
      Caption         =   "Liquid Diffusivity, DI (Nernst-Haskell)"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Begin VB.Label lblKineticParameters 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Anion:"
         ForeColor       =   &H80000008&
         Height          =   195
         Index           =   0
         Left            =   300
         TabIndex        =   33
         Top             =   720
         Width           =   615
      End
      Begin VB.Label lblKineticParameters 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Cation:"
         ForeColor       =   &H80000008&
         Height          =   195
         Index           =   1
         Left            =   2220
         TabIndex        =   32
         Top             =   720
         Width           =   615
      End
      Begin VB.Shape Shape1 
         Height          =   1155
         Left            =   240
         Top             =   660
         Width           =   1935
      End
      Begin VB.Shape Shape2 
         Height          =   1155
         Left            =   2160
         Top             =   660
         Width           =   1935
      End
      Begin VB.Label lblLiquidDiffusivity 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Valence"
         ForeColor       =   &H80000008&
         Height          =   195
         Index           =   0
         Left            =   300
         TabIndex        =   31
         Top             =   1320
         Width           =   795
      End
      Begin VB.Label lblLiquidDiffusivity 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "L.I.C."
         ForeColor       =   &H80000008&
         Height          =   195
         Index           =   1
         Left            =   300
         TabIndex        =   30
         Top             =   1560
         Width           =   795
      End
      Begin VB.Label lblLiquidDiffusivityValue 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   195
         Index           =   0
         Left            =   1200
         TabIndex        =   29
         Top             =   1320
         Width           =   795
      End
      Begin VB.Label lblLiquidDiffusivityValue 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   195
         Index           =   1
         Left            =   1200
         TabIndex        =   28
         Top             =   1560
         Width           =   795
      End
      Begin VB.Label lblLiquidDiffusivity 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Valence"
         ForeColor       =   &H80000008&
         Height          =   195
         Index           =   2
         Left            =   2280
         TabIndex        =   27
         Top             =   1320
         Width           =   795
      End
      Begin VB.Label lblLiquidDiffusivity 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "L.I.C."
         ForeColor       =   &H80000008&
         Height          =   195
         Index           =   3
         Left            =   2280
         TabIndex        =   26
         Top             =   1560
         Width           =   795
      End
      Begin VB.Label lblLiquidDiffusivityValue 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   195
         Index           =   2
         Left            =   3180
         TabIndex        =   25
         Top             =   1320
         Width           =   795
      End
      Begin VB.Label lblLiquidDiffusivityValue 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   195
         Index           =   3
         Left            =   3180
         TabIndex        =   24
         Top             =   1560
         Width           =   795
      End
      Begin VB.Shape Shape3 
         Height          =   315
         Left            =   240
         Top             =   1800
         Width           =   3855
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "L.I.C. = Limiting Ionic Conductance"
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   360
         TabIndex        =   23
         Top             =   1860
         Width           =   3675
      End
      Begin VB.Label lblLiquidDiffusivity 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Temperature"
         ForeColor       =   &H80000008&
         Height          =   195
         Index           =   4
         Left            =   240
         TabIndex        =   22
         Top             =   2220
         Width           =   1545
      End
      Begin VB.Label lblLiquidDiffusivityValue 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   195
         Index           =   4
         Left            =   1920
         TabIndex        =   21
         Top             =   2220
         Width           =   975
      End
      Begin VB.Label lblLiquidDiffusivityUnits 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "K"
         ForeColor       =   &H80000008&
         Height          =   195
         Index           =   0
         Left            =   3060
         TabIndex        =   20
         Top             =   2220
         Width           =   975
      End
      Begin VB.Label lblLiquidDiffusivity 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Faraday's Const."
         ForeColor       =   &H80000008&
         Height          =   195
         Index           =   5
         Left            =   240
         TabIndex        =   19
         Top             =   2400
         Width           =   1545
      End
      Begin VB.Label lblLiquidDiffusivityValue 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   195
         Index           =   5
         Left            =   1920
         TabIndex        =   18
         Top             =   2400
         Width           =   975
      End
      Begin VB.Label lblLiquidDiffusivityUnits 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "cal/g/eq"
         ForeColor       =   &H80000008&
         Height          =   195
         Index           =   1
         Left            =   3060
         TabIndex        =   17
         Top             =   2400
         Width           =   975
      End
      Begin VB.Label lblLiquidDiffusivity 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Gas Constant"
         ForeColor       =   &H80000008&
         Height          =   195
         Index           =   6
         Left            =   240
         TabIndex        =   16
         Top             =   2580
         Width           =   1545
      End
      Begin VB.Label lblLiquidDiffusivityValue 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   195
         Index           =   6
         Left            =   1920
         TabIndex        =   15
         Top             =   2580
         Width           =   975
      End
      Begin VB.Label lblLiquidDiffusivityUnits 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "J/mol/K"
         ForeColor       =   &H80000008&
         Height          =   195
         Index           =   2
         Left            =   3060
         TabIndex        =   14
         Top             =   2580
         Width           =   975
      End
      Begin VB.Label lblLiquidDiffusivity 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Liquid Diffusivity"
         ForeColor       =   &H80000008&
         Height          =   195
         Index           =   7
         Left            =   240
         TabIndex        =   13
         Top             =   3000
         Width           =   1545
      End
      Begin VB.Label lblLiquidDiffusivityValue 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   195
         Index           =   7
         Left            =   1920
         TabIndex        =   12
         Top             =   3000
         Width           =   975
      End
      Begin VB.Label lblLiquidDiffusivityUnits 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "cm2/s"
         ForeColor       =   &H80000008&
         Height          =   195
         Index           =   3
         Left            =   3060
         TabIndex        =   11
         Top             =   3000
         Width           =   975
      End
      Begin VB.Shape Shape4 
         Height          =   795
         Left            =   240
         Top             =   2100
         Width           =   3855
      End
      Begin VB.Shape Shape5 
         Height          =   315
         Left            =   240
         Top             =   360
         Width           =   3855
      End
      Begin VB.Label Label3 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Parameters Needed for Correlation"
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   360
         TabIndex        =   10
         Top             =   400
         Width           =   3615
      End
      Begin VB.Label lblAnion 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   420
         TabIndex        =   9
         Top             =   960
         Width           =   1635
      End
      Begin VB.Label lblCation 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   2340
         TabIndex        =   8
         Top             =   960
         Width           =   1635
      End
   End
   Begin VB.CommandButton cmdKinetic 
      Appearance      =   0  'Flat
      Caption         =   "Pore Diff., Dp"
      Height          =   552
      Index           =   2
      Left            =   2880
      TabIndex        =   6
      Top             =   0
      Width           =   1476
   End
   Begin VB.CommandButton cmdKinetic 
      Appearance      =   0  'Flat
      Caption         =   "&Ionic Transport, kf"
      Height          =   552
      Index           =   1
      Left            =   1440
      TabIndex        =   5
      Top             =   0
      Width           =   1476
   End
   Begin VB.CommandButton cmdKinetic 
      Appearance      =   0  'Flat
      Caption         =   "&Liquid Diff., Dl"
      Height          =   555
      Index           =   0
      Left            =   0
      TabIndex        =   4
      Top             =   0
      Width           =   1476
   End
   Begin VB.ComboBox cboIon 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   1260
      Style           =   2  'Dropdown List
      TabIndex        =   2
      Top             =   620
      Width           =   2952
   End
   Begin VB.CommandButton cmdCancel 
      Appearance      =   0  'Flat
      Caption         =   "&Cancel"
      Height          =   372
      Left            =   2280
      TabIndex        =   1
      Top             =   5760
      Width           =   1095
   End
   Begin VB.CommandButton cmdOK 
      Appearance      =   0  'Flat
      Caption         =   "&OK"
      Height          =   372
      Left            =   1080
      TabIndex        =   0
      Top             =   5760
      Width           =   975
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Ion:"
      ForeColor       =   &H80000008&
      Height          =   192
      Left            =   120
      TabIndex        =   3
      Top             =   620
      Width           =   1032
   End
End
Attribute VB_Name = "frmInputKineticParameters"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

Dim Temp_Text As String
Dim IsError As Integer

Private Sub cboIon_Click()
    Dim FoundCation As Integer, FoundAnion As Integer
    Dim IonToSearchFor As String
    Dim i As Integer
    Dim ValueToDisplay As Double
    Dim ListIndex As Integer
   
    ClickGeneratedFromcboIon = True

    'Search for Current Ion in Lists of Anions and Cations
    FoundCation = False
    FoundAnion = False
    EditingCation = False
    EditingAnion = False
    NumberOfIonToEdit = 0
    IonToSearchFor = Trim$(cboIon.List(cboIon.ListIndex))
    For i = 1 To NowProj.NumberOfCations
        If Trim$(NowProj.Cation(i).Name) = IonToSearchFor Then
           NumberOfIonToEdit = i
           FoundCation = True
           Exit For
        End If
    Next i
    If Not FoundCation Then
       For i = 1 To NowProj.NumberOfAnions
           If Trim$(NowProj.Anion(i).Name) = IonToSearchFor Then
              NumberOfIonToEdit = i
              FoundAnion = True
              Exit For
           End If
       Next i
    End If

    If FoundCation Then

       EditingCation = True

       'Display NernstHaskell Cation and Anion Information on Screen for Selected Cation
       NernstHaskell.SelectedCation = NowProj.Cation(NumberOfIonToEdit).Kinetic.NernstHaskellCation
       lblCation.Caption = Trim$(NernstHaskell.SelectedCation.Ion_Name)
       lblLiquidDiffusivityValue(2).Caption = "+" & Format$(NernstHaskell.SelectedCation.Valence, "0")
       lblLiquidDiffusivityValue(3).Caption = Trim$(Str$(NernstHaskell.SelectedCation.LimitingIonicConductance))

       NernstHaskell.SelectedAnion = NowProj.Cation(NumberOfIonToEdit).Kinetic.NernstHaskellAnion
       lblAnion.Caption = Trim$(NernstHaskell.SelectedAnion.Ion_Name)
       lblLiquidDiffusivityValue(0).Caption = "+" & Format$(NernstHaskell.SelectedAnion.Valence, "0")
       lblLiquidDiffusivityValue(1).Caption = Trim$(Str$(NernstHaskell.SelectedAnion.LimitingIonicConductance))

       'Show current temperature in units of K on Kinetic parameters form
       frmInputKineticParameters!lblLiquidDiffusivityValue(4).Caption = Format$(NowProj.Operating.Temperature, "0.00")
       
       'Show Liquid Diffusivity Values for Current Contaminant in Appropriate places
       ValueToDisplay = NowProj.Cation(NumberOfIonToEdit).Kinetic.LiquidDiffusivityCorrelation
       frmInputKineticParameters!lblLiquidDiffusivityValue(7).Caption = Format$(ValueToDisplay, GetTheFormat(ValueToDisplay))

       ListIndex = frmInputKineticParameters!cboLiquidDiffusivityUnits(0).ListIndex
       frmInputKineticParameters!cboLiquidDiffusivityUnits(0).ListIndex = 0
       frmInputKineticParameters!lblLiquidDiffCorrelation.Caption = frmInputKineticParameters!lblLiquidDiffusivityValue(7).Caption
       frmInputKineticParameters!cboLiquidDiffusivityUnits(0).ListIndex = ListIndex

       If NowProj.Cation(NumberOfIonToEdit).Kinetic.LiquidDiffusivity.UserInput = False Then
          ListIndex = frmInputKineticParameters!cboLiquidDiffusivityUnits(1).ListIndex
          cboLiquidDiffusivityUnits(1).ListIndex = 0
          frmInputKineticParameters!txtLiquidDiffUserInput.Text = frmInputKineticParameters!lblLiquidDiffusivityValue(7).Caption
          frmInputKineticParameters!cboLiquidDiffusivityUnits(1).ListIndex = ListIndex
          frmInputKineticParameters!optLiquidDiffusivity(0).Value = True
       Else
          ListIndex = frmInputKineticParameters!cboLiquidDiffusivityUnits(1).ListIndex
          frmInputKineticParameters!cboLiquidDiffusivityUnits(1).ListIndex = 0
          ValueToDisplay = NowProj.Cation(NumberOfIonToEdit).Kinetic.LiquidDiffusivityUserInput
          frmInputKineticParameters!txtLiquidDiffUserInput.Text = Format$(ValueToDisplay, GetTheFormat(ValueToDisplay))
          frmInputKineticParameters!cboLiquidDiffusivityUnits(1).ListIndex = ListIndex
          frmInputKineticParameters!optLiquidDiffusivity(1).Value = True
       End If
       ValueToDisplay = NowProj.Cation(NumberOfIonToEdit).Kinetic.LiquidDiffusivity.Value
       frmInputKineticParameters!lblIonicTranportCoeffValue(11).Caption = Format$(ValueToDisplay, GetTheFormat(ValueToDisplay))
          
       'Load parameters for kf calculation into label boxes

       'Particle Diameter (cm)
       ValueToDisplay = NowProj.Resin.ParticleDiameter * 100
       frmInputKineticParameters!lblIonicTranportCoeffValue(0).Caption = Format$(ValueToDisplay, GetTheFormat(ValueToDisplay))

       'Column Diameter (cm)
       ValueToDisplay = NowProj.Bed.Diameter * 100
       frmInputKineticParameters!lblIonicTranportCoeffValue(1).Caption = Format$(ValueToDisplay, GetTheFormat(ValueToDisplay))

       'Column Area (cm2)
       ValueToDisplay = NowProj.Bed.Area * 100 ^ 2
       frmInputKineticParameters!lblIonicTranportCoeffValue(2).Caption = Format$(ValueToDisplay, GetTheFormat(ValueToDisplay))

       'Inlet Flow Rate (cm3/s)
       ValueToDisplay = NowProj.Bed.Flowrate.Value * 1000000#
       frmInputKineticParameters!lblIonicTranportCoeffValue(3).Caption = Format$(ValueToDisplay, GetTheFormat(ValueToDisplay))

       'Superficial Velocity (cm/s)
       ValueToDisplay = NowProj.Bed.SuperficialVelocity
       frmInputKineticParameters!lblIonicTranportCoeffValue(4).Caption = Format$(ValueToDisplay, GetTheFormat(ValueToDisplay))

       'Bed Porosity (-)
       ValueToDisplay = NowProj.Bed.Porosity
       frmInputKineticParameters!lblIonicTranportCoeffValue(5).Caption = Format$(ValueToDisplay, GetTheFormat(ValueToDisplay))

       'Interstitial Velocity (cm/s)
       ValueToDisplay = NowProj.Bed.InterstitialVelocity
       frmInputKineticParameters!lblIonicTranportCoeffValue(6).Caption = Format$(ValueToDisplay, GetTheFormat(ValueToDisplay))

       'Temperature (C)
       ValueToDisplay = NowProj.Operating.Temperature - 273.15
       frmInputKineticParameters!lblIonicTranportCoeffValue(7).Caption = Format$(ValueToDisplay, GetTheFormat(ValueToDisplay))

       'Liquid Density (g/cm3)
       ValueToDisplay = NowProj.Operating.LiquidDensity
       frmInputKineticParameters!lblIonicTranportCoeffValue(8).Caption = Format$(ValueToDisplay, GetTheFormat(ValueToDisplay))

       'Liquid Viscosity (g/cm/s)
       ValueToDisplay = NowProj.Operating.LiquidViscosity
       frmInputKineticParameters!lblIonicTranportCoeffValue(9).Caption = Format$(ValueToDisplay, GetTheFormat(ValueToDisplay))

       'Reynolds Number (-)
       ValueToDisplay = NowProj.Cation(NumberOfIonToEdit).Kinetic.ReynoldsNumber
       frmInputKineticParameters!lblIonicTranportCoeffValue(10).Caption = Format$(ValueToDisplay, GetTheFormat(ValueToDisplay))

       'Schmidt Number (-)
       ValueToDisplay = NowProj.Cation(NumberOfIonToEdit).Kinetic.SchmidtNumber
       frmInputKineticParameters!lblIonicTranportCoeffValue(12).Caption = Format$(ValueToDisplay, GetTheFormat(ValueToDisplay))

       'Ionic Transport Coefficient, kf (cm/s)
       ValueToDisplay = NowProj.Cation(NumberOfIonToEdit).Kinetic.IonicTransportCoeffCorrelation
       frmInputKineticParameters!lblIonicTranportCoeffValue(13).Caption = Format$(ValueToDisplay, GetTheFormat(ValueToDisplay))

       ListIndex = frmInputKineticParameters!cboIonicTransportUnits(0).ListIndex
       frmInputKineticParameters!cboIonicTransportUnits(0).ListIndex = 0
       frmInputKineticParameters!txtIonicTransportCoeffCorr.Text = frmInputKineticParameters!lblIonicTranportCoeffValue(13).Caption
       frmInputKineticParameters!cboIonicTransportUnits(0).ListIndex = ListIndex

       If NowProj.Cation(NumberOfIonToEdit).Kinetic.IonicTransportCoefficient.UserInput = False Then
          ListIndex = frmInputKineticParameters!cboIonicTransportUnits(1).ListIndex
          frmInputKineticParameters!cboIonicTransportUnits(1).ListIndex = 0
          frmInputKineticParameters!txtIonicTransCoeffUser.Text = frmInputKineticParameters!lblIonicTranportCoeffValue(13).Caption
          frmInputKineticParameters!cboIonicTransportUnits(1).ListIndex = ListIndex
          frmInputKineticParameters!optIonicTransportCoeff(0).Value = True
       Else
          ListIndex = frmInputKineticParameters!cboIonicTransportUnits(1).ListIndex
          frmInputKineticParameters!cboIonicTransportUnits(1).ListIndex = 0
          ValueToDisplay = NowProj.Cation(NumberOfIonToEdit).Kinetic.IonicTransportCoeffUserInput
          frmInputKineticParameters!txtIonicTransCoeffUser.Text = Format$(ValueToDisplay, GetTheFormat(ValueToDisplay))
          frmInputKineticParameters!cboIonicTransportUnits(1).ListIndex = ListIndex
          frmInputKineticParameters!optIonicTransportCoeff(1).Value = True
       End If

       'Pore Diffusivity information

       'Liquid Diffusivity
       ValueToDisplay = NowProj.Cation(NumberOfIonToEdit).Kinetic.LiquidDiffusivity.Value
       frmInputKineticParameters!lblPoreDiffusivityValue(0).Caption = Format$(ValueToDisplay, GetTheFormat(ValueToDisplay))

       'Tortuosity
       ValueToDisplay = NowProj.Resin.Tortuosity
       frmInputKineticParameters!lblPoreDiffusivityValue(1).Caption = Format$(ValueToDisplay, GetTheFormat(ValueToDisplay))

       'PoreDiffusivity
       
       ValueToDisplay = NowProj.Cation(NumberOfIonToEdit).Kinetic.PoreDiffusivityCorrelation
       frmInputKineticParameters!lblPoreDiffusivityValue(2).Caption = Format$(ValueToDisplay, GetTheFormat(ValueToDisplay))

       ListIndex = frmInputKineticParameters!cboPoreDiffusivityUnits(0).ListIndex
       frmInputKineticParameters!cboPoreDiffusivityUnits(0).ListIndex = 0
       frmInputKineticParameters!txtPoreDiffusivityCorr.Text = lblPoreDiffusivityValue(2).Caption
       frmInputKineticParameters!cboPoreDiffusivityUnits(0).ListIndex = ListIndex

       If NowProj.Cation(NumberOfIonToEdit).Kinetic.PoreDiffusivity.UserInput = False Then
          ListIndex = frmInputKineticParameters!cboPoreDiffusivityUnits(1).ListIndex
          frmInputKineticParameters!cboPoreDiffusivityUnits(1).ListIndex = 0
          frmInputKineticParameters!txtPoreDiffusivityUser.Text = frmInputKineticParameters!lblPoreDiffusivityValue(2).Caption
          frmInputKineticParameters!cboPoreDiffusivityUnits(1).ListIndex = ListIndex
          frmInputKineticParameters!optPoreDiffusivity(0).Value = True
       Else
          ListIndex = frmInputKineticParameters!cboPoreDiffusivityUnits(1).ListIndex
          frmInputKineticParameters!cboPoreDiffusivityUnits(1).ListIndex = 0
          ValueToDisplay = NowProj.Cation(NumberOfIonToEdit).Kinetic.PoreDiffusivityUserInput
          frmInputKineticParameters!txtPoreDiffusivityUser.Text = Format$(ValueToDisplay, GetTheFormat(ValueToDisplay))
          frmInputKineticParameters!cboPoreDiffusivityUnits(1).ListIndex = ListIndex
          frmInputKineticParameters!optPoreDiffusivity(1).Value = True
       End If

    ElseIf FoundAnion Then

       EditingAnion = True

       'Display NernstHaskell Cation and Anion Information on Screen for Selected Anion
       NernstHaskell.SelectedCation = NowProj.Anion(NumberOfIonToEdit).Kinetic.NernstHaskellCation
       lblCation.Caption = Trim$(NernstHaskell.SelectedCation.Ion_Name)
       lblLiquidDiffusivityValue(2).Caption = "+" & Format$(NernstHaskell.SelectedCation.Valence, "0")
       lblLiquidDiffusivityValue(3).Caption = Trim$(Str$(NernstHaskell.SelectedCation.LimitingIonicConductance))

       NernstHaskell.SelectedAnion = NowProj.Anion(NumberOfIonToEdit).Kinetic.NernstHaskellAnion
       lblAnion.Caption = Trim$(NernstHaskell.SelectedAnion.Ion_Name)
       lblLiquidDiffusivityValue(0).Caption = "+" & Format$(NernstHaskell.SelectedAnion.Valence, "0")
       lblLiquidDiffusivityValue(1).Caption = Trim$(Str$(NernstHaskell.SelectedAnion.LimitingIonicConductance))

       'Show current temperature in units of K on Kinetic parameters form
       frmInputKineticParameters!lblLiquidDiffusivityValue(4).Caption = Format$(NowProj.Operating.Temperature, "0.00")

       'Show Liquid Diffusivity Values for Current Contaminant in Appropriate places
       ValueToDisplay = NowProj.Anion(NumberOfIonToEdit).Kinetic.LiquidDiffusivityCorrelation
       frmInputKineticParameters!lblLiquidDiffusivityValue(7).Caption = Format$(ValueToDisplay, GetTheFormat(ValueToDisplay))

       ListIndex = frmInputKineticParameters!cboLiquidDiffusivityUnits(0).ListIndex
       frmInputKineticParameters!cboLiquidDiffusivityUnits(0).ListIndex = 0
       frmInputKineticParameters!lblLiquidDiffCorrelation.Caption = frmInputKineticParameters!lblLiquidDiffusivityValue(7).Caption
       frmInputKineticParameters!cboLiquidDiffusivityUnits(0).ListIndex = ListIndex

       If NowProj.Anion(NumberOfIonToEdit).Kinetic.LiquidDiffusivity.UserInput = False Then
          ListIndex = frmInputKineticParameters!cboLiquidDiffusivityUnits(1).ListIndex
          frmInputKineticParameters!cboLiquidDiffusivityUnits(1).ListIndex = 0
          frmInputKineticParameters!txtLiquidDiffUserInput.Text = frmInputKineticParameters!lblLiquidDiffusivityValue(7).Caption
          frmInputKineticParameters!cboLiquidDiffusivityUnits(1).ListIndex = ListIndex
          frmInputKineticParameters!optLiquidDiffusivity(0).Value = True
       Else
          ListIndex = frmInputKineticParameters!cboLiquidDiffusivityUnits(1).ListIndex
          frmInputKineticParameters!cboLiquidDiffusivityUnits(1).ListIndex = 0
          ValueToDisplay = NowProj.Anion(NumberOfIonToEdit).Kinetic.LiquidDiffusivityUserInput
          frmInputKineticParameters!txtLiquidDiffUserInput.Text = Format$(ValueToDisplay, GetTheFormat(ValueToDisplay))
          frmInputKineticParameters!cboLiquidDiffusivityUnits(1).ListIndex = ListIndex
          frmInputKineticParameters!optLiquidDiffusivity(1).Value = True
       End If
       ValueToDisplay = NowProj.Anion(NumberOfIonToEdit).Kinetic.LiquidDiffusivity.Value
       frmInputKineticParameters!lblIonicTranportCoeffValue(11).Caption = Format$(ValueToDisplay, GetTheFormat(ValueToDisplay))

       'Load parameters for kf calculation into label boxes

       'Particle Diameter (cm)
       ValueToDisplay = NowProj.Resin.ParticleDiameter * 100
       frmInputKineticParameters!lblIonicTranportCoeffValue(0).Caption = Format$(ValueToDisplay, GetTheFormat(ValueToDisplay))

       'Column Diameter (cm)
       ValueToDisplay = NowProj.Bed.Diameter * 100
       frmInputKineticParameters!lblIonicTranportCoeffValue(1).Caption = Format$(ValueToDisplay, GetTheFormat(ValueToDisplay))

       'Column Area (cm2)
       ValueToDisplay = NowProj.Bed.Area * 100 ^ 2
       frmInputKineticParameters!lblIonicTranportCoeffValue(2).Caption = Format$(ValueToDisplay, GetTheFormat(ValueToDisplay))

       'Inlet Flow Rate (cm3/s)
       ValueToDisplay = NowProj.Bed.Flowrate.Value * 1000000#
       frmInputKineticParameters!lblIonicTranportCoeffValue(3).Caption = Format$(ValueToDisplay, GetTheFormat(ValueToDisplay))

       'Superficial Velocity (cm/s)
       ValueToDisplay = NowProj.Bed.SuperficialVelocity
       frmInputKineticParameters!lblIonicTranportCoeffValue(4).Caption = Format$(ValueToDisplay, GetTheFormat(ValueToDisplay))

       'Bed Porosity (-)
       ValueToDisplay = NowProj.Bed.Porosity
       frmInputKineticParameters!lblIonicTranportCoeffValue(5).Caption = Format$(ValueToDisplay, GetTheFormat(ValueToDisplay))

       'Interstitial Velocity (cm/s)
       ValueToDisplay = NowProj.Bed.InterstitialVelocity
       frmInputKineticParameters!lblIonicTranportCoeffValue(6).Caption = Format$(ValueToDisplay, GetTheFormat(ValueToDisplay))

       'Temperature (C)
       ValueToDisplay = NowProj.Operating.Temperature - 273.15
       frmInputKineticParameters!lblIonicTranportCoeffValue(7).Caption = Format$(ValueToDisplay, GetTheFormat(ValueToDisplay))

       'Liquid Density (g/cm3)
       ValueToDisplay = NowProj.Operating.LiquidDensity
       frmInputKineticParameters!lblIonicTranportCoeffValue(8).Caption = Format$(ValueToDisplay, GetTheFormat(ValueToDisplay))

       'Liquid Viscosity (g/cm/s)
       ValueToDisplay = NowProj.Operating.LiquidViscosity
       frmInputKineticParameters!lblIonicTranportCoeffValue(9).Caption = Format$(ValueToDisplay, GetTheFormat(ValueToDisplay))

       'Reynolds Number (-)
       ValueToDisplay = NowProj.Anion(NumberOfIonToEdit).Kinetic.ReynoldsNumber
       frmInputKineticParameters!lblIonicTranportCoeffValue(10).Caption = Format$(ValueToDisplay, GetTheFormat(ValueToDisplay))

       'Schmidt Number (-)
       ValueToDisplay = NowProj.Anion(NumberOfIonToEdit).Kinetic.SchmidtNumber
       frmInputKineticParameters!lblIonicTranportCoeffValue(12).Caption = Format$(ValueToDisplay, GetTheFormat(ValueToDisplay))

       'Ionic Transport Coefficient, kf (cm/s)
       ValueToDisplay = NowProj.Anion(NumberOfIonToEdit).Kinetic.IonicTransportCoeffCorrelation
       frmInputKineticParameters!lblIonicTranportCoeffValue(13).Caption = Format$(ValueToDisplay, GetTheFormat(ValueToDisplay))

       ListIndex = frmInputKineticParameters!cboIonicTransportUnits(0).ListIndex
       frmInputKineticParameters!cboIonicTransportUnits(0).ListIndex = 0
       frmInputKineticParameters!txtIonicTransportCoeffCorr.Text = frmInputKineticParameters!lblIonicTranportCoeffValue(13).Caption
       frmInputKineticParameters!cboIonicTransportUnits(0).ListIndex = ListIndex

       If NowProj.Anion(NumberOfIonToEdit).Kinetic.IonicTransportCoefficient.UserInput = False Then
          ListIndex = frmInputKineticParameters!cboIonicTransportUnits(1).ListIndex
          frmInputKineticParameters!cboIonicTransportUnits(1).ListIndex = 0
          frmInputKineticParameters!txtIonicTransCoeffUser.Text = frmInputKineticParameters!lblIonicTranportCoeffValue(13).Caption
          frmInputKineticParameters!cboIonicTransportUnits(1).ListIndex = ListIndex
          frmInputKineticParameters!optIonicTransportCoeff(0).Value = True
       Else
          ListIndex = frmInputKineticParameters!cboIonicTransportUnits(1).ListIndex
          frmInputKineticParameters!cboIonicTransportUnits(1).ListIndex = 0
          ValueToDisplay = NowProj.Anion(NumberOfIonToEdit).Kinetic.IonicTransportCoeffUserInput
          frmInputKineticParameters!txtIonicTransCoeffUser.Text = Format$(ValueToDisplay, GetTheFormat(ValueToDisplay))
          frmInputKineticParameters!cboIonicTransportUnits(1).ListIndex = ListIndex
          frmInputKineticParameters!optIonicTransportCoeff(1).Value = True
       End If

       'Pore Diffusivity information

       'Liquid Diffusivity
       ValueToDisplay = NowProj.Anion(NumberOfIonToEdit).Kinetic.LiquidDiffusivity.Value
       lblPoreDiffusivityValue(0).Caption = Format$(ValueToDisplay, GetTheFormat(ValueToDisplay))

       'Tortuosity
       ValueToDisplay = NowProj.Resin.Tortuosity
       lblPoreDiffusivityValue(1).Caption = Format$(ValueToDisplay, GetTheFormat(ValueToDisplay))

       'Pore Diffusivity
       
       ValueToDisplay = NowProj.Anion(NumberOfIonToEdit).Kinetic.PoreDiffusivityCorrelation
       lblPoreDiffusivityValue(2).Caption = Format$(ValueToDisplay, GetTheFormat(ValueToDisplay))

       ListIndex = frmInputKineticParameters!cboPoreDiffusivityUnits(0).ListIndex
       frmInputKineticParameters!cboPoreDiffusivityUnits(0).ListIndex = 0
       frmInputKineticParameters!txtPoreDiffusivityCorr.Text = lblPoreDiffusivityValue(2).Caption
       frmInputKineticParameters!cboPoreDiffusivityUnits(0).ListIndex = ListIndex

       If NowProj.Anion(NumberOfIonToEdit).Kinetic.PoreDiffusivity.UserInput = False Then
          ListIndex = frmInputKineticParameters!cboPoreDiffusivityUnits(1).ListIndex
          frmInputKineticParameters!cboPoreDiffusivityUnits(1).ListIndex = 0
          frmInputKineticParameters!txtPoreDiffusivityUser.Text = frmInputKineticParameters!lblPoreDiffusivityValue(2).Caption
          frmInputKineticParameters!cboPoreDiffusivityUnits(1).ListIndex = ListIndex
          frmInputKineticParameters!optPoreDiffusivity(0).Value = True
       Else
          ListIndex = frmInputKineticParameters!cboPoreDiffusivityUnits(1).ListIndex
          frmInputKineticParameters!cboPoreDiffusivityUnits(1).ListIndex = 0
          ValueToDisplay = NowProj.Anion(NumberOfIonToEdit).Kinetic.PoreDiffusivityUserInput
          frmInputKineticParameters!txtPoreDiffusivityUser.Text = Format$(ValueToDisplay, GetTheFormat(ValueToDisplay))
          frmInputKineticParameters!cboPoreDiffusivityUnits(1).ListIndex = ListIndex
          frmInputKineticParameters!optPoreDiffusivity(1).Value = True
       End If

    End If

    If FoundCation Or FoundAnion Then
       If FoundCation Then
          If NumSelectedCations = 1 Then GoTo ExitSub
          For i = 1 To NumSelectedCations
             If Cations_Selected(i) = NumberOfIonToEdit Then
                Call CalculateDimensionlessGroups
'                frmIonExchangeMain!cboKinDimComponent.ListIndex = -1
'                frmIonExchangeMain!cboKinDimComponent.ListIndex = i - 1
             End If
          Next i
       End If

       If FoundAnion Then
          If NumSelectedAnions = 1 Then GoTo ExitSub
          For i = 1 To NumSelectedAnions
             If Anions_Selected(i) = NumberOfIonToEdit Then
                Call CalculateDimensionlessGroups
'                frmIonExchangeMain!cboKinDimComponent.ListIndex = -1
'                frmIonExchangeMain!cboKinDimComponent.ListIndex = i - 1
             End If
          Next i
       End If
    End If

ExitSub:
    ClickGeneratedFromcboIon = False

End Sub

Private Sub cboIonicTransport_Click()
    Dim i As Integer, ListIndex As Integer
    Dim PermanentIonToEdit As Integer

    Select Case cboIonicTransport.ListIndex
       Case 0   'Wildhagen correlation
          If NowProj.IonicTransportCoeffCorrName = IONIC_TRANSPORT_COEFFICIENT_1 Then Exit Sub
          NowProj.IonicTransportCoeffCorrName = IONIC_TRANSPORT_COEFFICIENT_1
       Case 1   'Gnielinski correlation
          If NowProj.IonicTransportCoeffCorrName = IONIC_TRANSPORT_COEFFICIENT_2 Then Exit Sub
          NowProj.IonicTransportCoeffCorrName = IONIC_TRANSPORT_COEFFICIENT_2
    End Select

    If (NowProj.NumberOfCations = 0) And (NowProj.NumberOfAnions = 0) Then Exit Sub

    AddingCation = True
    AddingAnion = False
    For i = 1 To NowProj.NumberOfCations
        NumberOfIonToEdit = i
        Call CalculateKineticParameters
    Next i
    AddingCation = False
    AddingAnion = True
    For i = 1 To NowProj.NumberOfAnions
        NumberOfIonToEdit = i
        Call CalculateKineticParameters
    Next i

    'Update dimensionless groups for all but currently selected ion (which will be updated in click event on cboIon)
    If NumSelectedCations > 0 Then
       PermanentIonToEdit = NumberOfIonToEdit
       AddingCation = True
       For i = 1 To NumSelectedCations
           NumberOfIonToEdit = Cations_Selected(i)
           If NumberOfIonToEdit <> PermanentIonToEdit Then
              Call CalculateDimensionlessGroups
           End If
       Next i
       AddingCation = False
       NumberOfIonToEdit = PermanentIonToEdit
    End If

    If NumSelectedAnions > 0 Then
       PermanentIonToEdit = NumberOfIonToEdit
       AddingAnion = True
       For i = 1 To NumSelectedAnions
           NumberOfIonToEdit = Anions_Selected(i)
           If NumberOfIonToEdit <> PermanentIonToEdit Then
              Call CalculateDimensionlessGroups
           End If
       Next i
       AddingAnion = False
       NumberOfIonToEdit = PermanentIonToEdit
    End If

    'Generate click on cboIon
    ListIndex = cboIon.ListIndex
    cboIon.ListIndex = -1
    cboIon.ListIndex = ListIndex

End Sub

Private Sub cboIonicTransportUnits_Click(Index As Integer)
    Dim ValueToConvert As Double
    Dim ValueToDisplay As Double
    Dim ListIndex As Integer

    If cboIon.ListCount = 0 Then Exit Sub

    Select Case Index
       Case 0   'Correlation
          If EditingCation Then
             ValueToConvert = NowProj.Cation(NumberOfIonToEdit).Kinetic.IonicTransportCoeffCorrelation
          ElseIf EditingAnion Then
             ValueToConvert = NowProj.Anion(NumberOfIonToEdit).Kinetic.IonicTransportCoeffCorrelation
          End If
          ListIndex = cboIonicTransportUnits(0).ListIndex
       Case 1   'User Input
          If EditingCation Then
             ValueToConvert = NowProj.Cation(NumberOfIonToEdit).Kinetic.IonicTransportCoeffUserInput
          ElseIf EditingAnion Then
             ValueToConvert = NowProj.Anion(NumberOfIonToEdit).Kinetic.IonicTransportCoeffUserInput
          End If
          ListIndex = cboIonicTransportUnits(1).ListIndex
    End Select

    Select Case ListIndex
       Case VELOCITY_CM_per_S     'cm/s
          ValueToDisplay = ValueToConvert
       Case VELOCITY_CM_per_MIN   'cm/min
          ValueToDisplay = ValueToConvert * VelocityConversionFactor(VELOCITY_CM_per_MIN)
       Case VELOCITY_M_per_S      'm/s
          ValueToDisplay = ValueToConvert * VelocityConversionFactor(VELOCITY_M_per_S)
       Case VELOCITY_M_per_MIN    'm/min
          ValueToDisplay = ValueToConvert * VelocityConversionFactor(VELOCITY_M_per_MIN)
       Case VELOCITY_M_per_HR     'm/hr
          ValueToDisplay = ValueToConvert * VelocityConversionFactor(VELOCITY_M_per_HR)
       Case VELOCITY_M_per_D      'm/d
          ValueToDisplay = ValueToConvert * VelocityConversionFactor(VELOCITY_M_per_D)
       Case VELOCITY_FT_per_S     'ft/s
          ValueToDisplay = ValueToConvert * VelocityConversionFactor(VELOCITY_FT_per_S)
       Case VELOCITY_FT_per_MIN   'ft/min
          ValueToDisplay = ValueToConvert * VelocityConversionFactor(VELOCITY_FT_per_MIN)
       Case VELOCITY_FT_per_HR    'ft/hr
          ValueToDisplay = ValueToConvert * VelocityConversionFactor(VELOCITY_FT_per_HR)
       Case VELOCITY_FT_per_D     'ft/d
          ValueToDisplay = ValueToConvert * VelocityConversionFactor(VELOCITY_FT_per_D)
       End Select

       Select Case Index
          Case 0   'Correlation
             txtIonicTransportCoeffCorr.Text = Format$(ValueToDisplay, GetTheFormat(ValueToDisplay))
          Case 1   'User Input
             txtIonicTransCoeffUser.Text = Format$(ValueToDisplay, GetTheFormat(ValueToDisplay))
       End Select

End Sub

Private Sub cboLiquidDiffusivityUnits_Click(Index As Integer)
    Dim ValueToConvert As Double
    Dim ValueToDisplay As Double
    Dim ListIndex As Integer

    If cboIon.ListCount = 0 Then Exit Sub

    Select Case Index
       Case 0   'Correlation
          If EditingCation Then
             ValueToConvert = NowProj.Cation(NumberOfIonToEdit).Kinetic.LiquidDiffusivityCorrelation
          ElseIf EditingAnion Then
             ValueToConvert = NowProj.Anion(NumberOfIonToEdit).Kinetic.LiquidDiffusivityCorrelation
          End If
          ListIndex = cboLiquidDiffusivityUnits(0).ListIndex
       Case 1   'User Input
          If EditingCation Then
             ValueToConvert = NowProj.Cation(NumberOfIonToEdit).Kinetic.LiquidDiffusivityUserInput
          ElseIf EditingAnion Then
             ValueToConvert = NowProj.Anion(NumberOfIonToEdit).Kinetic.LiquidDiffusivityUserInput
          End If
          ListIndex = cboLiquidDiffusivityUnits(1).ListIndex
    End Select

    Select Case ListIndex
       Case DIFFUSIVITY_CM2_per_S     'cm2/s
          ValueToDisplay = ValueToConvert
       Case DIFFUSIVITY_CM2_per_MIN   'cm2/min
          ValueToDisplay = ValueToConvert * DiffusivityConversionFactor(DIFFUSIVITY_CM2_per_MIN)
       Case DIFFUSIVITY_M2_per_S      'm2/s
          ValueToDisplay = ValueToConvert * DiffusivityConversionFactor(DIFFUSIVITY_M2_per_S)
       Case DIFFUSIVITY_M2_per_MIN    'm2/min
          ValueToDisplay = ValueToConvert * DiffusivityConversionFactor(DIFFUSIVITY_M2_per_MIN)
       Case DIFFUSIVITY_M2_per_HR     'm2/hr
          ValueToDisplay = ValueToConvert * DiffusivityConversionFactor(DIFFUSIVITY_M2_per_HR)
       Case DIFFUSIVITY_M2_per_D      'm2/d
          ValueToDisplay = ValueToConvert * DiffusivityConversionFactor(DIFFUSIVITY_M2_per_D)
       Case DIFFUSIVITY_FT2_per_S     'ft2/s
          ValueToDisplay = ValueToConvert * DiffusivityConversionFactor(DIFFUSIVITY_FT2_per_S)
       Case DIFFUSIVITY_FT2_per_MIN   'ft2/min
          ValueToDisplay = ValueToConvert * DiffusivityConversionFactor(DIFFUSIVITY_FT2_per_MIN)
       Case DIFFUSIVITY_FT2_per_HR    'ft2/hr
          ValueToDisplay = ValueToConvert * DiffusivityConversionFactor(DIFFUSIVITY_FT2_per_HR)
       Case DIFFUSIVITY_FT2_per_D     'ft2/d
          ValueToDisplay = ValueToConvert * DiffusivityConversionFactor(DIFFUSIVITY_FT2_per_D)
       End Select

       Select Case Index
          Case 0   'Correlation
             txtLiquidDiffCorrelation.Text = Format$(ValueToDisplay, GetTheFormat(ValueToDisplay))
          Case 1   'User Input
             txtLiquidDiffUserInput.Text = Format$(ValueToDisplay, GetTheFormat(ValueToDisplay))
       End Select
    
End Sub

Private Sub cboPoreDiffusivityUnits_Click(Index As Integer)
    Dim ValueToConvert As Double
    Dim ValueToDisplay As Double
    Dim ListIndex As Integer

    If cboIon.ListCount = 0 Then Exit Sub

    Select Case Index
       Case 0   'Correlation
          If EditingCation Then
             ValueToConvert = NowProj.Cation(NumberOfIonToEdit).Kinetic.PoreDiffusivityCorrelation
          ElseIf EditingAnion Then
             ValueToConvert = NowProj.Anion(NumberOfIonToEdit).Kinetic.PoreDiffusivityCorrelation
          End If
          ListIndex = cboPoreDiffusivityUnits(0).ListIndex
       Case 1   'User Input
          If EditingCation Then
             ValueToConvert = NowProj.Cation(NumberOfIonToEdit).Kinetic.PoreDiffusivityUserInput
          ElseIf EditingAnion Then
             ValueToConvert = NowProj.Anion(NumberOfIonToEdit).Kinetic.PoreDiffusivityUserInput
          End If
          ListIndex = cboPoreDiffusivityUnits(1).ListIndex
    End Select

    Select Case ListIndex
       Case DIFFUSIVITY_CM2_per_S     'cm2/s
          ValueToDisplay = ValueToConvert
       Case DIFFUSIVITY_CM2_per_MIN   'cm2/min
          ValueToDisplay = ValueToConvert * DiffusivityConversionFactor(DIFFUSIVITY_CM2_per_MIN)
       Case DIFFUSIVITY_M2_per_S      'm2/s
          ValueToDisplay = ValueToConvert * DiffusivityConversionFactor(DIFFUSIVITY_M2_per_S)
       Case DIFFUSIVITY_M2_per_MIN    'm2/min
          ValueToDisplay = ValueToConvert * DiffusivityConversionFactor(DIFFUSIVITY_M2_per_MIN)
       Case DIFFUSIVITY_M2_per_HR     'm2/hr
          ValueToDisplay = ValueToConvert * DiffusivityConversionFactor(DIFFUSIVITY_M2_per_HR)
       Case DIFFUSIVITY_M2_per_D      'm2/d
          ValueToDisplay = ValueToConvert * DiffusivityConversionFactor(DIFFUSIVITY_M2_per_D)
       Case DIFFUSIVITY_FT2_per_S     'ft2/s
          ValueToDisplay = ValueToConvert * DiffusivityConversionFactor(DIFFUSIVITY_FT2_per_S)
       Case DIFFUSIVITY_FT2_per_MIN   'ft2/min
          ValueToDisplay = ValueToConvert * DiffusivityConversionFactor(DIFFUSIVITY_FT2_per_MIN)
       Case DIFFUSIVITY_FT2_per_HR    'ft2/hr
          ValueToDisplay = ValueToConvert * DiffusivityConversionFactor(DIFFUSIVITY_FT2_per_HR)
       Case DIFFUSIVITY_FT2_per_D     'ft2/d
          ValueToDisplay = ValueToConvert * DiffusivityConversionFactor(DIFFUSIVITY_FT2_per_D)
       End Select

       Select Case Index
          Case 0   'Correlation
             txtPoreDiffusivityCorr.Text = Format$(ValueToDisplay, GetTheFormat(ValueToDisplay))
          Case 1   'User Input
             txtPoreDiffusivityUser.Text = Format$(ValueToDisplay, GetTheFormat(ValueToDisplay))
       End Select

End Sub

Private Sub cmdCancel_Click()
    Dim i As Integer
    Dim ListIndex As Integer

    If EditingCation And EditingAnion Then
       For i = 1 To NowProj.NumberOfCations
           NowProj.Cation(i).Kinetic = OldCationKineticParameters(i)
       Next i

       For i = 1 To NowProj.NumberOfAnions
           NowProj.Anion(i).Kinetic = OldAnionKineticParameters(i)
       Next i

    ElseIf EditingCation Then
       For i = 1 To NowProj.NumberOfCations
           NowProj.Cation(i).Kinetic = OldCationKineticParameters(i)
       Next i

    ElseIf EditingAnion Then
       For i = 1 To NowProj.NumberOfAnions
           NowProj.Anion(i).Kinetic = OldAnionKineticParameters(i)
       Next i

    End If

    Call UpdateDimensionlessGroupAllIons

    ListIndex = cboIon.ListIndex
    cboIon.ListIndex = -1
    cboIon.ListIndex = ListIndex

    frmInputKineticParameters.Hide

End Sub

Private Sub cmdKinetic_Click(Index As Integer)

    Select Case Index
       Case 0   'Liquid Diffusivity
            fraPoreDiffusivity.visible = False
            fraPoreDiffusivityInput.visible = False
            fraIonicTransportCoefficient.visible = False
            fraIonicTransport.visible = False
            fraLiquidDiffusivity.visible = True
            fraLiquidDiffusivityInput.visible = True
            cmdKinetic(0).Enabled = False
            cmdKinetic(1).Enabled = True
            cmdKinetic(2).Enabled = True

       Case 1   'Ionic Transport Coefficient
            fraLiquidDiffusivity.visible = False
            fraLiquidDiffusivityInput.visible = False
            fraPoreDiffusivity.visible = False
            fraPoreDiffusivityInput.visible = False
            fraIonicTransportCoefficient.visible = True
            fraIonicTransport.visible = True
            cmdKinetic(0).Enabled = True
            cmdKinetic(1).Enabled = False
            cmdKinetic(2).Enabled = True
            
       Case 2   'Pore Diffusivity
            fraLiquidDiffusivity.visible = False
            fraLiquidDiffusivityInput.visible = False
            fraIonicTransportCoefficient.visible = False
            fraIonicTransport.visible = False
            fraPoreDiffusivity.visible = True
            fraPoreDiffusivityInput.visible = True
            cmdKinetic(0).Enabled = True
            cmdKinetic(1).Enabled = True
            cmdKinetic(2).Enabled = False

    End Select
End Sub

Private Sub cmdOK_Click()
    Dim ListIndex As Integer

    frmInputKineticParameters.Hide

End Sub

Private Sub Form_Activate()

    'Set up form so Liquid Diffusivity information is visible and other info. isn't
    fraLiquidDiffusivity.visible = True

    fraLiquidDiffusivityInput.visible = True

    fraIonicTransportCoefficient.visible = False
    fraIonicTransportCoefficient.top = fraLiquidDiffusivity.top
    fraIonicTransportCoefficient.left = fraLiquidDiffusivity.left

    fraIonicTransport.visible = False
    fraIonicTransport.top = fraIonicTransportCoefficient.top + fraIonicTransportCoefficient.height + 120
    fraIonicTransport.left = fraIonicTransportCoefficient.left

    fraPoreDiffusivity.visible = False
    fraPoreDiffusivity.top = fraLiquidDiffusivity.top
    fraPoreDiffusivity.left = fraLiquidDiffusivity.left

    fraPoreDiffusivityInput.visible = False
    fraPoreDiffusivityInput.top = fraPoreDiffusivity.top + fraPoreDiffusivity.height + 120
    fraPoreDiffusivityInput.left = fraPoreDiffusivity.left

    cmdKinetic(0).Enabled = False
    cmdKinetic(1).Enabled = True
    cmdKinetic(2).Enabled = True

    Call frmInputKineticParameters_Refresh
    
End Sub

Private Sub Form_Load()
    Dim PositionLeft As Integer

    frmInputKineticParameters.WindowState = 0

    frmInputKineticParameters.width = 4440
    frmInputKineticParameters.height = 6800

    'Position the form on the screen (Centered on left half of it)
    If WindowState = 0 Then
       'don't attempt if screen Minimized or Maximized
       PositionLeft = ((Screen.width / 2 - frmIonExchangeMain.left) / 2) - frmInputKineticParameters.width / 2
       Move (frmIonExchangeMain.left + PositionLeft), (Screen.height - frmInputKineticParameters.height) / 2

    End If
    
'    Call Populate_frmInputKineticParameters_Units
    Call LoadLiquidDiffusivityParameters

    'Load Liquid Diffusivity Units with superscripts
    lblLiquidDiffusivityUnits(3).Caption = "cm²/s"
    lblIonicTransportCoeffUnits(2).Caption = "cm²"
    lblIonicTransportCoeffUnits(3).Caption = "cm³/s"
    lblIonicTransportCoeffUnits(8).Caption = "g/cm³"
    lblIonicTransportCoeffUnits(11).Caption = "cm²/s"
    lblPoreDiffusivityUnits(0).Caption = "cm²/s"
    lblPoreDiffusivityUnits(2).Caption = "cm²/s"

    cboIonicTransport.AddItem IONIC_TRANSPORT_COEFFICIENT_1
    cboIonicTransport.AddItem IONIC_TRANSPORT_COEFFICIENT_2
    
    Call frmInputKineticParameters_Refresh
    
End Sub

Private Sub LoadLiquidDiffusivityParameters()
    'Liquid Diffusivity Parameters
    'Nernst-Haskell Parameters

    NernstHaskell.FaradaysConstant = 96500#
    lblLiquidDiffusivityValue(5).Caption = Trim$(Str$(NernstHaskell.FaradaysConstant))

    NernstHaskell.GasConstant = 8.314
    lblLiquidDiffusivityValue(6).Caption = Trim$(Str$(NernstHaskell.GasConstant))

End Sub



Private Sub optIonicTransportCoeff_Click(Index As Integer, Value As Integer)
    Dim ValueToDisplay As Double
    Dim ListIndex As Integer

    Select Case Index
       Case 0   'Correlation
          txtIonicTransCoeffUser.Enabled = False
          txtIonicTransportCoeffCorr.Enabled = False
          If EditingCation Then
             NowProj.Cation(NumberOfIonToEdit).Kinetic.IonicTransportCoefficient.Value = NowProj.Cation(NumberOfIonToEdit).Kinetic.IonicTransportCoeffCorrelation
             NowProj.Cation(NumberOfIonToEdit).Kinetic.IonicTransportCoefficient.UserInput = False
          ElseIf EditingAnion Then
             NowProj.Anion(NumberOfIonToEdit).Kinetic.IonicTransportCoefficient.Value = NowProj.Anion(NumberOfIonToEdit).Kinetic.IonicTransportCoeffCorrelation
             NowProj.Anion(NumberOfIonToEdit).Kinetic.IonicTransportCoefficient.UserInput = False
          End If
       Case 1   'User Input
          txtIonicTransportCoeffCorr.Enabled = False
          txtIonicTransCoeffUser.Enabled = True
          If EditingCation Then
             NowProj.Cation(NumberOfIonToEdit).Kinetic.IonicTransportCoefficient.Value = NowProj.Cation(NumberOfIonToEdit).Kinetic.IonicTransportCoeffUserInput
             NowProj.Cation(NumberOfIonToEdit).Kinetic.IonicTransportCoefficient.UserInput = True
          ElseIf EditingAnion Then
             NowProj.Anion(NumberOfIonToEdit).Kinetic.IonicTransportCoefficient.Value = NowProj.Anion(NumberOfIonToEdit).Kinetic.IonicTransportCoeffUserInput
             NowProj.Anion(NumberOfIonToEdit).Kinetic.IonicTransportCoefficient.UserInput = True
          End If
    End Select
    Call CalculateKineticParameters

    If Not ClickGeneratedFromcboIon Then   'Generate click on cboIon
       ListIndex = cboIon.ListIndex
       cboIon.ListIndex = -1
       cboIon.ListIndex = ListIndex
    End If

End Sub

Private Sub optLiquidDiffusivity_Click(Index As Integer, Value As Integer)
    Dim ValueToDisplay As Double
    Dim ListIndex As Integer

    Select Case Index
       Case 0   'Correlation
          txtLiquidDiffUserInput.Enabled = False
          txtLiquidDiffCorrelation.Enabled = False
          If EditingCation Then
             NowProj.Cation(NumberOfIonToEdit).Kinetic.LiquidDiffusivity.Value = NowProj.Cation(NumberOfIonToEdit).Kinetic.LiquidDiffusivityCorrelation
             NowProj.Cation(NumberOfIonToEdit).Kinetic.LiquidDiffusivity.UserInput = False
          ElseIf EditingAnion Then
             NowProj.Anion(NumberOfIonToEdit).Kinetic.LiquidDiffusivity.Value = NowProj.Anion(NumberOfIonToEdit).Kinetic.LiquidDiffusivityCorrelation
             NowProj.Anion(NumberOfIonToEdit).Kinetic.LiquidDiffusivity.UserInput = False
          End If
       Case 1   'User Input
          txtLiquidDiffUserInput.Enabled = True
          txtLiquidDiffCorrelation.Enabled = False
          If EditingCation Then
             NowProj.Cation(NumberOfIonToEdit).Kinetic.LiquidDiffusivity.Value = NowProj.Cation(NumberOfIonToEdit).Kinetic.LiquidDiffusivityUserInput
             NowProj.Cation(NumberOfIonToEdit).Kinetic.LiquidDiffusivity.UserInput = True
          ElseIf EditingAnion Then
             NowProj.Anion(NumberOfIonToEdit).Kinetic.LiquidDiffusivity.Value = NowProj.Anion(NumberOfIonToEdit).Kinetic.LiquidDiffusivityUserInput
             NowProj.Anion(NumberOfIonToEdit).Kinetic.LiquidDiffusivity.UserInput = True
          End If
    End Select
    Call CalculateKineticParameters

    If Not ClickGeneratedFromcboIon Then   'Generate click on cboIon
       ListIndex = cboIon.ListIndex
       cboIon.ListIndex = -1
       cboIon.ListIndex = ListIndex
    End If

End Sub

Private Sub optPoreDiffusivity_Click(Index As Integer, Value As Integer)
    Dim ValueToDisplay As Double
    Dim ListIndex As Integer

    Select Case Index
       Case 0   'Correlation
          txtPoreDiffusivityUser.Enabled = False
          txtPoreDiffusivityCorr.Enabled = False
          If EditingCation Then
             NowProj.Cation(NumberOfIonToEdit).Kinetic.PoreDiffusivity.Value = NowProj.Cation(NumberOfIonToEdit).Kinetic.PoreDiffusivityCorrelation
             NowProj.Cation(NumberOfIonToEdit).Kinetic.PoreDiffusivity.UserInput = False
          ElseIf EditingAnion Then
             NowProj.Anion(NumberOfIonToEdit).Kinetic.PoreDiffusivity.Value = NowProj.Anion(NumberOfIonToEdit).Kinetic.PoreDiffusivityCorrelation
             NowProj.Anion(NumberOfIonToEdit).Kinetic.PoreDiffusivity.UserInput = False
          End If
       Case 1   'User Input
          txtPoreDiffusivityCorr.Enabled = False
          txtPoreDiffusivityUser.Enabled = True
          If EditingCation Then
             NowProj.Cation(NumberOfIonToEdit).Kinetic.PoreDiffusivity.Value = NowProj.Cation(NumberOfIonToEdit).Kinetic.PoreDiffusivityUserInput
             NowProj.Cation(NumberOfIonToEdit).Kinetic.PoreDiffusivity.UserInput = True
          ElseIf EditingAnion Then
             NowProj.Anion(NumberOfIonToEdit).Kinetic.PoreDiffusivity.Value = NowProj.Anion(NumberOfIonToEdit).Kinetic.PoreDiffusivityUserInput
             NowProj.Anion(NumberOfIonToEdit).Kinetic.PoreDiffusivity.UserInput = True
          End If
    End Select
    Call CalculateKineticParameters

    If Not ClickGeneratedFromcboIon Then   'Generate click on cboIon
       ListIndex = cboIon.ListIndex
       cboIon.ListIndex = -1
       cboIon.ListIndex = ListIndex
    End If

End Sub

Private Sub txtIonicTransCoeffUser_GotFocus()
    Call TextGetFocus(txtIonicTransCoeffUser, Temp_Text)
End Sub

Private Sub txtIonicTransCoeffUser_KeyPress(KeyAscii As Integer)

    If KeyAscii = 13 Then
       KeyAscii = 0
       SendKeys "{Tab}"
       Exit Sub
    End If

    Call NumberCheck(KeyAscii)

End Sub

Private Sub txtIonicTransCoeffUser_LostFocus()
    Dim ValueChanged As Integer, NewValue As Double
    Dim msg As String, CurrentUnits As Integer
    Dim OldValue As Double, ValueToDisplay As Double
    Dim i As Integer
    Dim ListIndex As Integer

    Call TextHandleError(IsError, txtIonicTransCoeffUser, Temp_Text)

    If Not IsError Then
       NewValue = CDbl(txtIonicTransCoeffUser.Text)

       If EditingCation Then
          OldValue = NowProj.Cation(NumberOfIonToEdit).Kinetic.IonicTransportCoeffUserInput
       ElseIf EditingAnion Then
          OldValue = NowProj.Anion(NumberOfIonToEdit).Kinetic.IonicTransportCoeffUserInput
       End If

       'Convert NewValue to Standard Units if Necessary
       CurrentUnits = cboIonicTransportUnits(1).ListIndex
       If CurrentUnits <> 0 Then
          NewValue = NewValue / VelocityConversionFactor(CurrentUnits)
       End If

       Call NumberChanged(ValueChanged, OldValue, NewValue)
       If ValueChanged Then
          If HaveValue(NewValue) Then
             If EditingCation Then
                NowProj.Cation(NumberOfIonToEdit).Kinetic.IonicTransportCoeffUserInput = NewValue
                NowProj.Cation(NumberOfIonToEdit).Kinetic.IonicTransportCoefficient.Value = NewValue
             ElseIf EditingAnion Then
                NowProj.Anion(NumberOfIonToEdit).Kinetic.IonicTransportCoeffUserInput = NewValue
                NowProj.Anion(NumberOfIonToEdit).Kinetic.IonicTransportCoefficient.Value = NewValue
             End If

             'Recalculate dimensionless groups and print on main form
             ListIndex = cboIon.ListIndex
             cboIon.ListIndex = -1
             cboIon.ListIndex = ListIndex

          Else
             txtIonicTransCoeffUser.Text = Temp_Text
             txtIonicTransCoeffUser.SetFocus
             Exit Sub
          End If
       End If
    End If

End Sub


Private Sub txtLiquidDiffUserInput_GotFocus()
    Call TextGetFocus(txtLiquidDiffUserInput, Temp_Text)
End Sub

Private Sub txtLiquidDiffUserInput_KeyPress(KeyAscii As Integer)

    If KeyAscii = 13 Then
       KeyAscii = 0
       SendKeys "{Tab}"
       Exit Sub
    End If

    Call NumberCheck(KeyAscii)

End Sub

Private Sub txtLiquidDiffUserInput_LostFocus()
    Dim ValueChanged As Integer, NewValue As Double
    Dim msg As String, CurrentUnits As Integer
    Dim OldValue As Double, ValueToDisplay As Double
    Dim i As Integer
    Dim ListIndex As Integer

    Call TextHandleError(IsError, txtLiquidDiffUserInput, Temp_Text)

    If Not IsError Then
       NewValue = CDbl(txtLiquidDiffUserInput.Text)

       If EditingCation Then
          OldValue = NowProj.Cation(NumberOfIonToEdit).Kinetic.LiquidDiffusivityUserInput
       ElseIf EditingAnion Then
          OldValue = NowProj.Anion(NumberOfIonToEdit).Kinetic.LiquidDiffusivityUserInput
       End If

       'Convert NewValue to Standard Units if Necessary
       CurrentUnits = cboLiquidDiffusivityUnits(1).ListIndex
       If CurrentUnits <> 0 Then
          NewValue = NewValue / DiffusivityConversionFactor(CurrentUnits)
       End If

       Call NumberChanged(ValueChanged, OldValue, NewValue)
       If ValueChanged Then
          If HaveValue(NewValue) Then
             If EditingCation Then
                NowProj.Cation(NumberOfIonToEdit).Kinetic.LiquidDiffusivityUserInput = NewValue
                NowProj.Cation(NumberOfIonToEdit).Kinetic.LiquidDiffusivity.Value = NewValue
             ElseIf EditingAnion Then
                NowProj.Anion(NumberOfIonToEdit).Kinetic.LiquidDiffusivityUserInput = NewValue
                NowProj.Anion(NumberOfIonToEdit).Kinetic.LiquidDiffusivity.Value = NewValue
             End If

             Call CalculateKineticParameters

             'Reprint values on frmInputKineticParameters based on new Liquid Diffusivity
             ListIndex = cboIon.ListIndex
             cboIon.ListIndex = -1
             cboIon.ListIndex = ListIndex

          Else
             txtLiquidDiffUserInput.Text = Temp_Text
             txtLiquidDiffUserInput.SetFocus
             Exit Sub
          End If
       End If
    End If

End Sub

Private Sub txtPoreDiffusivityUser_GotFocus()
    Call TextGetFocus(txtPoreDiffusivityUser, Temp_Text)
End Sub

Private Sub txtPoreDiffusivityUser_KeyPress(KeyAscii As Integer)

    If KeyAscii = 13 Then
       KeyAscii = 0
       SendKeys "{Tab}"
       Exit Sub
    End If

    Call NumberCheck(KeyAscii)

End Sub

Private Sub txtPoreDiffusivityUser_LostFocus()
    Dim ValueChanged As Integer, NewValue As Double
    Dim msg As String, CurrentUnits As Integer
    Dim OldValue As Double, ValueToDisplay As Double
    Dim i As Integer
    Dim ListIndex As Integer

    Call TextHandleError(IsError, txtPoreDiffusivityUser, Temp_Text)

    If Not IsError Then
       NewValue = CDbl(txtPoreDiffusivityUser.Text)

       If EditingCation Then
          OldValue = NowProj.Cation(NumberOfIonToEdit).Kinetic.PoreDiffusivityUserInput
       ElseIf EditingAnion Then
          OldValue = NowProj.Anion(NumberOfIonToEdit).Kinetic.PoreDiffusivityUserInput
       End If

       'Convert NewValue to Standard Units if Necessary
       CurrentUnits = cboPoreDiffusivityUnits(1).ListIndex
       If CurrentUnits <> 0 Then
          NewValue = NewValue / DiffusivityConversionFactor(CurrentUnits)
       End If

       Call NumberChanged(ValueChanged, OldValue, NewValue)
       If ValueChanged Then
          If HaveValue(NewValue) Then
             If EditingCation Then
                NowProj.Cation(NumberOfIonToEdit).Kinetic.PoreDiffusivityUserInput = NewValue
                NowProj.Cation(NumberOfIonToEdit).Kinetic.PoreDiffusivity.Value = NewValue
             ElseIf EditingAnion Then
                NowProj.Anion(NumberOfIonToEdit).Kinetic.PoreDiffusivityUserInput = NewValue
                NowProj.Anion(NumberOfIonToEdit).Kinetic.PoreDiffusivity.Value = NewValue
             End If

             'Recalculate dimensionless groups and print on main form
             ListIndex = cboIon.ListIndex
             cboIon.ListIndex = -1
             cboIon.ListIndex = ListIndex

          Else
             txtPoreDiffusivityUser.Text = Temp_Text
             txtPoreDiffusivityUser.SetFocus
             Exit Sub
          End If
       End If
    End If

End Sub

