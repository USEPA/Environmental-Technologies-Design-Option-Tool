VERSION 5.00
Begin VB.Form frmInputKineticParameters 
   Appearance      =   0  'Flat
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Kinetic Parameters"
   ClientHeight    =   6480
   ClientLeft      =   2430
   ClientTop       =   270
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
   Begin VB.CommandButton cmdKinetic 
      Appearance      =   0  'Flat
      Caption         =   "Pore Diff., Dp"
      Height          =   552
      Index           =   2
      Left            =   2880
      TabIndex        =   105
      Top             =   0
      Width           =   1476
   End
   Begin VB.CommandButton cmdKinetic 
      Appearance      =   0  'Flat
      Caption         =   "&Ionic Transport, kf"
      Height          =   552
      Index           =   1
      Left            =   1440
      TabIndex        =   104
      Top             =   0
      Width           =   1476
   End
   Begin VB.CommandButton cmdKinetic 
      Appearance      =   0  'Flat
      Caption         =   "&Liquid Diff., Dl"
      Height          =   552
      Index           =   0
      Left            =   0
      TabIndex        =   103
      Top             =   0
      Width           =   1476
   End
   Begin VB.PictureBox fraPoreDiffusivityInput 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H00000000&
      Height          =   1032
      Left            =   4440
      ScaleHeight     =   1005
      ScaleWidth      =   4065
      TabIndex        =   96
      Top             =   5340
      Width           =   4092
      Begin VB.ComboBox cboPoreDiffusivityUnits 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         Height          =   288
         Index           =   1
         Left            =   2820
         Style           =   2  'Dropdown List
         TabIndex        =   102
         TabStop         =   0   'False
         Top             =   660
         Width           =   1152
      End
      Begin VB.ComboBox cboPoreDiffusivityUnits 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         Height          =   288
         Index           =   0
         Left            =   2820
         Style           =   2  'Dropdown List
         TabIndex        =   100
         TabStop         =   0   'False
         Top             =   300
         Width           =   1152
      End
      Begin VB.TextBox txtPoreDiffusivityUser 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         Height          =   288
         Left            =   1740
         TabIndex        =   99
         Top             =   660
         Width           =   972
      End
      Begin VB.PictureBox optPoreDiffusivity 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   192
         Index           =   1
         Left            =   120
         ScaleHeight     =   165
         ScaleWidth      =   1485
         TabIndex        =   98
         TabStop         =   0   'False
         Top             =   720
         Width           =   1512
      End
      Begin VB.PictureBox optPoreDiffusivity 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   192
         Index           =   0
         Left            =   120
         ScaleHeight     =   165
         ScaleWidth      =   1485
         TabIndex        =   97
         Top             =   360
         Width           =   1512
      End
      Begin VB.Label lblPoreDiffusivityCorr 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   252
         Left            =   1740
         TabIndex        =   101
         Top             =   300
         Width           =   972
      End
   End
   Begin VB.PictureBox fraIonicTransport 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H00000000&
      Height          =   1032
      Left            =   4440
      ScaleHeight     =   1005
      ScaleWidth      =   4065
      TabIndex        =   89
      Top             =   3540
      Width           =   4092
      Begin VB.PictureBox optIonicTransportCoeff 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   192
         Index           =   0
         Left            =   120
         ScaleHeight     =   165
         ScaleWidth      =   1485
         TabIndex        =   94
         Top             =   360
         Width           =   1512
      End
      Begin VB.PictureBox optIonicTransportCoeff 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   192
         Index           =   1
         Left            =   120
         ScaleHeight     =   165
         ScaleWidth      =   1485
         TabIndex        =   93
         TabStop         =   0   'False
         Top             =   720
         Width           =   1512
      End
      Begin VB.TextBox txtIonicTransCoeffUser 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         Height          =   288
         Left            =   1740
         TabIndex        =   92
         Top             =   660
         Width           =   972
      End
      Begin VB.ComboBox cboIonicTransportUnits 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         Height          =   288
         Index           =   0
         Left            =   2820
         Style           =   2  'Dropdown List
         TabIndex        =   91
         TabStop         =   0   'False
         Top             =   300
         Width           =   1152
      End
      Begin VB.ComboBox cboIonicTransportUnits 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         Height          =   288
         Index           =   1
         Left            =   2820
         Style           =   2  'Dropdown List
         TabIndex        =   90
         TabStop         =   0   'False
         Top             =   660
         Width           =   1152
      End
      Begin VB.Label lblIonicTransportCoeffCorr 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   252
         Left            =   1740
         TabIndex        =   95
         Top             =   300
         Width           =   972
      End
   End
   Begin VB.PictureBox fraLiquidDiffusivityInput 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H00000000&
      Height          =   1032
      Left            =   120
      ScaleHeight     =   1005
      ScaleWidth      =   4065
      TabIndex        =   82
      Top             =   4500
      Width           =   4092
      Begin VB.ComboBox cboLiquidDiffusivityUnits 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         Height          =   288
         Index           =   1
         Left            =   2820
         Style           =   2  'Dropdown List
         TabIndex        =   88
         TabStop         =   0   'False
         Top             =   660
         Width           =   1152
      End
      Begin VB.ComboBox cboLiquidDiffusivityUnits 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         Height          =   288
         Index           =   0
         Left            =   2820
         Style           =   2  'Dropdown List
         TabIndex        =   87
         TabStop         =   0   'False
         Top             =   300
         Width           =   1152
      End
      Begin VB.TextBox txtLiquidDiffUserInput 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         Height          =   288
         Left            =   1740
         TabIndex        =   86
         Top             =   660
         Width           =   972
      End
      Begin VB.PictureBox optLiquidDiffusivity 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   192
         Index           =   1
         Left            =   120
         ScaleHeight     =   165
         ScaleWidth      =   1485
         TabIndex        =   84
         TabStop         =   0   'False
         Top             =   720
         Width           =   1512
      End
      Begin VB.PictureBox optLiquidDiffusivity 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   192
         Index           =   0
         Left            =   120
         ScaleHeight     =   165
         ScaleWidth      =   1485
         TabIndex        =   83
         Top             =   360
         Width           =   1512
      End
      Begin VB.Label lblLiquidDiffCorrelation 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   252
         Left            =   1740
         TabIndex        =   85
         Top             =   300
         Width           =   972
      End
   End
   Begin VB.PictureBox fraPoreDiffusivity 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H00000000&
      Height          =   1032
      Left            =   4440
      ScaleHeight     =   1005
      ScaleWidth      =   4065
      TabIndex        =   72
      Top             =   4620
      Width           =   4092
      Begin VB.Label lblPoreDiffusivityUnits 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "(-)"
         ForeColor       =   &H80000008&
         Height          =   192
         Index           =   1
         Left            =   2940
         TabIndex        =   81
         Top             =   480
         Width           =   972
      End
      Begin VB.Label lblPoreDiffusivityValue 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   192
         Index           =   1
         Left            =   1800
         TabIndex        =   80
         Top             =   480
         Width           =   972
      End
      Begin VB.Label lblPoreDiffusivity 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Tortuosity"
         ForeColor       =   &H80000008&
         Height          =   192
         Index           =   1
         Left            =   120
         TabIndex        =   79
         Top             =   480
         Width           =   1548
      End
      Begin VB.Label lblPoreDiffusivityUnits 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "cm2/s"
         ForeColor       =   &H80000008&
         Height          =   192
         Index           =   0
         Left            =   2940
         TabIndex        =   78
         Top             =   300
         Width           =   972
      End
      Begin VB.Label lblPoreDiffusivityValue 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   192
         Index           =   0
         Left            =   1800
         TabIndex        =   77
         Top             =   300
         Width           =   972
      End
      Begin VB.Label lblPoreDiffusivity 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Liquid Diffusivity"
         ForeColor       =   &H80000008&
         Height          =   192
         Index           =   0
         Left            =   120
         TabIndex        =   76
         Top             =   300
         Width           =   1548
      End
      Begin VB.Label lblPoreDiffusivity 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Pore Diffusivity"
         ForeColor       =   &H80000008&
         Height          =   192
         Index           =   2
         Left            =   120
         TabIndex        =   75
         Top             =   720
         Width           =   1548
      End
      Begin VB.Label lblPoreDiffusivityValue 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   192
         Index           =   2
         Left            =   1800
         TabIndex        =   74
         Top             =   720
         Width           =   972
      End
      Begin VB.Label lblPoreDiffusivityUnits 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "cm2/s"
         ForeColor       =   &H80000008&
         Height          =   192
         Index           =   2
         Left            =   2940
         TabIndex        =   73
         Top             =   720
         Width           =   972
      End
   End
   Begin VB.ComboBox cboIon 
      Appearance      =   0  'Flat
      Height          =   288
      Left            =   1260
      Style           =   2  'Dropdown List
      TabIndex        =   26
      Top             =   660
      Width           =   2952
   End
   Begin VB.PictureBox fraIonicTransportCoefficient 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H00000000&
      Height          =   3375
      Left            =   4440
      ScaleHeight     =   3345
      ScaleWidth      =   4065
      TabIndex        =   28
      Top             =   120
      Width           =   4095
      Begin VB.ComboBox cboIonicTransport 
         Appearance      =   0  'Flat
         Height          =   300
         Left            =   1800
         Style           =   2  'Dropdown List
         TabIndex        =   109
         Top             =   2760
         Width           =   2175
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
         Left            =   180
         TabIndex        =   108
         Top             =   2820
         Width           =   1545
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
         Left            =   120
         TabIndex        =   71
         Top             =   3120
         Width           =   1545
      End
      Begin VB.Label lblIonicTranportCoeffValue 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   195
         Index           =   13
         Left            =   1800
         TabIndex        =   70
         Top             =   3120
         Width           =   975
      End
      Begin VB.Label lblIonicTransportCoeffUnits 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "cm/s"
         ForeColor       =   &H80000008&
         Height          =   195
         Index           =   13
         Left            =   2940
         TabIndex        =   69
         Top             =   3120
         Width           =   975
      End
      Begin VB.Label lblIonicTransportCoeff 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Column Area"
         ForeColor       =   &H80000008&
         Height          =   192
         Index           =   2
         Left            =   120
         TabIndex        =   32
         Top             =   660
         Width           =   1548
      End
      Begin VB.Label lblIonicTransportCoeffUnits 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "(-)"
         ForeColor       =   &H80000008&
         Height          =   192
         Index           =   12
         Left            =   2940
         TabIndex        =   68
         Top             =   2460
         Width           =   972
      End
      Begin VB.Label lblIonicTranportCoeffValue 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   192
         Index           =   12
         Left            =   1800
         TabIndex        =   67
         Top             =   2460
         Width           =   972
      End
      Begin VB.Label lblIonicTransportCoeff 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Schmidt No."
         ForeColor       =   &H80000008&
         Height          =   192
         Index           =   12
         Left            =   120
         TabIndex        =   66
         Top             =   2460
         Width           =   1548
      End
      Begin VB.Label lblIonicTransportCoeffUnits 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "cm2/s"
         ForeColor       =   &H80000008&
         Height          =   192
         Index           =   11
         Left            =   2940
         TabIndex        =   65
         Top             =   2280
         Width           =   972
      End
      Begin VB.Label lblIonicTranportCoeffValue 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   192
         Index           =   11
         Left            =   1800
         TabIndex        =   64
         Top             =   2280
         Width           =   972
      End
      Begin VB.Label lblIonicTransportCoeff 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Liquid Diffusivity"
         ForeColor       =   &H80000008&
         Height          =   192
         Index           =   11
         Left            =   120
         TabIndex        =   63
         Top             =   2280
         Width           =   1548
      End
      Begin VB.Label lblIonicTransportCoeffUnits 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "(-)"
         ForeColor       =   &H80000008&
         Height          =   192
         Index           =   10
         Left            =   2940
         TabIndex        =   62
         Top             =   2100
         Width           =   972
      End
      Begin VB.Label lblIonicTranportCoeffValue 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   192
         Index           =   10
         Left            =   1800
         TabIndex        =   61
         Top             =   2100
         Width           =   972
      End
      Begin VB.Label lblIonicTransportCoeff 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Reynold's No."
         ForeColor       =   &H80000008&
         Height          =   192
         Index           =   10
         Left            =   120
         TabIndex        =   60
         Top             =   2100
         Width           =   1548
      End
      Begin VB.Label lblIonicTransportCoeffUnits 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "g/cm/s"
         ForeColor       =   &H80000008&
         Height          =   192
         Index           =   9
         Left            =   2940
         TabIndex        =   59
         Top             =   1920
         Width           =   972
      End
      Begin VB.Label lblIonicTranportCoeffValue 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   192
         Index           =   9
         Left            =   1800
         TabIndex        =   58
         Top             =   1920
         Width           =   972
      End
      Begin VB.Label lblIonicTransportCoeff 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Liquid Viscosity"
         ForeColor       =   &H80000008&
         Height          =   192
         Index           =   9
         Left            =   120
         TabIndex        =   57
         Top             =   1920
         Width           =   1548
      End
      Begin VB.Label lblIonicTransportCoeff 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Liquid Density"
         ForeColor       =   &H80000008&
         Height          =   192
         Index           =   8
         Left            =   120
         TabIndex        =   33
         Top             =   1740
         Width           =   1548
      End
      Begin VB.Label lblIonicTranportCoeffValue 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   192
         Index           =   8
         Left            =   1800
         TabIndex        =   34
         Top             =   1740
         Width           =   972
      End
      Begin VB.Label lblIonicTransportCoeffUnits 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "g/cm3"
         ForeColor       =   &H80000008&
         Height          =   192
         Index           =   8
         Left            =   2940
         TabIndex        =   35
         Top             =   1740
         Width           =   972
      End
      Begin VB.Label lblIonicTransportCoeff 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Temperature"
         ForeColor       =   &H80000008&
         Height          =   192
         Index           =   7
         Left            =   120
         TabIndex        =   36
         Top             =   1560
         Width           =   1548
      End
      Begin VB.Label lblIonicTranportCoeffValue 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   192
         Index           =   7
         Left            =   1800
         TabIndex        =   37
         Top             =   1560
         Width           =   972
      End
      Begin VB.Label lblIonicTransportCoeffUnits 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "C"
         ForeColor       =   &H80000008&
         Height          =   192
         Index           =   7
         Left            =   2940
         TabIndex        =   38
         Top             =   1560
         Width           =   972
      End
      Begin VB.Label lblIonicTransportCoeff 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Interstitial Vel."
         ForeColor       =   &H80000008&
         Height          =   192
         Index           =   6
         Left            =   120
         TabIndex        =   56
         Top             =   1380
         Width           =   1548
      End
      Begin VB.Label lblIonicTranportCoeffValue 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   192
         Index           =   6
         Left            =   1800
         TabIndex        =   55
         Top             =   1380
         Width           =   972
      End
      Begin VB.Label lblIonicTransportCoeffUnits 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "cm/s"
         ForeColor       =   &H80000008&
         Height          =   192
         Index           =   6
         Left            =   2940
         TabIndex        =   54
         Top             =   1380
         Width           =   972
      End
      Begin VB.Label lblIonicTransportCoeff 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Bed Porosity"
         ForeColor       =   &H80000008&
         Height          =   192
         Index           =   5
         Left            =   120
         TabIndex        =   53
         Top             =   1200
         Width           =   1548
      End
      Begin VB.Label lblIonicTranportCoeffValue 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   192
         Index           =   5
         Left            =   1800
         TabIndex        =   52
         Top             =   1200
         Width           =   972
      End
      Begin VB.Label lblIonicTransportCoeffUnits 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "(-)"
         ForeColor       =   &H80000008&
         Height          =   192
         Index           =   5
         Left            =   2940
         TabIndex        =   51
         Top             =   1200
         Width           =   972
      End
      Begin VB.Label lblIonicTransportCoeff 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Superficial Vel."
         ForeColor       =   &H80000008&
         Height          =   192
         Index           =   4
         Left            =   120
         TabIndex        =   50
         Top             =   1020
         Width           =   1548
      End
      Begin VB.Label lblIonicTranportCoeffValue 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   192
         Index           =   4
         Left            =   1800
         TabIndex        =   49
         Top             =   1020
         Width           =   972
      End
      Begin VB.Label lblIonicTransportCoeffUnits 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "cm/s"
         ForeColor       =   &H80000008&
         Height          =   192
         Index           =   4
         Left            =   2940
         TabIndex        =   48
         Top             =   1020
         Width           =   972
      End
      Begin VB.Label lblIonicTransportCoeff 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Inlet Flow Rate"
         ForeColor       =   &H80000008&
         Height          =   192
         Index           =   3
         Left            =   120
         TabIndex        =   47
         Top             =   840
         Width           =   1548
      End
      Begin VB.Label lblIonicTranportCoeffValue 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   192
         Index           =   3
         Left            =   1800
         TabIndex        =   46
         Top             =   840
         Width           =   972
      End
      Begin VB.Label lblIonicTransportCoeffUnits 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "cm3/s"
         ForeColor       =   &H80000008&
         Height          =   192
         Index           =   3
         Left            =   2940
         TabIndex        =   45
         Top             =   840
         Width           =   972
      End
      Begin VB.Label lblIonicTranportCoeffValue 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   192
         Index           =   2
         Left            =   1800
         TabIndex        =   43
         Top             =   660
         Width           =   972
      End
      Begin VB.Label lblIonicTransportCoeffUnits 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "cm2"
         ForeColor       =   &H80000008&
         Height          =   192
         Index           =   2
         Left            =   2940
         TabIndex        =   42
         Top             =   660
         Width           =   972
      End
      Begin VB.Label lblIonicTransportCoeff 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Column Diameter"
         ForeColor       =   &H80000008&
         Height          =   192
         Index           =   1
         Left            =   120
         TabIndex        =   41
         Top             =   480
         Width           =   1548
      End
      Begin VB.Label lblIonicTranportCoeffValue 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   192
         Index           =   1
         Left            =   1800
         TabIndex        =   40
         Top             =   480
         Width           =   972
      End
      Begin VB.Label lblIonicTransportCoeffUnits 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "cm"
         ForeColor       =   &H80000008&
         Height          =   192
         Index           =   1
         Left            =   2940
         TabIndex        =   39
         Top             =   480
         Width           =   972
      End
      Begin VB.Label lblIonicTransportCoeff 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Particle Diameter"
         ForeColor       =   &H80000008&
         Height          =   192
         Index           =   0
         Left            =   120
         TabIndex        =   31
         Top             =   300
         Width           =   1548
      End
      Begin VB.Label lblIonicTranportCoeffValue 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   192
         Index           =   0
         Left            =   1800
         TabIndex        =   30
         Top             =   300
         Width           =   972
      End
      Begin VB.Label lblIonicTransportCoeffUnits 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "cm"
         ForeColor       =   &H80000008&
         Height          =   192
         Index           =   0
         Left            =   2940
         TabIndex        =   29
         Top             =   300
         Width           =   972
      End
   End
   Begin VB.PictureBox fraLiquidDiffusivity 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H00000000&
      Height          =   3312
      Left            =   120
      ScaleHeight     =   3285
      ScaleWidth      =   4065
      TabIndex        =   2
      Top             =   1080
      Width           =   4092
      Begin VB.Label lblCation 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   252
         Left            =   2220
         TabIndex        =   107
         Top             =   960
         Width           =   1632
      End
      Begin VB.Label lblAnion 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   252
         Left            =   300
         TabIndex        =   106
         Top             =   960
         Width           =   1632
      End
      Begin VB.Label Label3 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Parameters Needed for Correlation"
         ForeColor       =   &H80000008&
         Height          =   192
         Left            =   240
         TabIndex        =   27
         Top             =   420
         Width           =   3612
      End
      Begin VB.Shape Shape5 
         Height          =   312
         Left            =   120
         Top             =   360
         Width           =   3852
      End
      Begin VB.Shape Shape4 
         Height          =   792
         Left            =   120
         Top             =   2100
         Width           =   3852
      End
      Begin VB.Label lblLiquidDiffusivityUnits 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "cm2/s"
         ForeColor       =   &H80000008&
         Height          =   192
         Index           =   3
         Left            =   2940
         TabIndex        =   14
         Top             =   3000
         Width           =   972
      End
      Begin VB.Label lblLiquidDiffusivityValue 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   192
         Index           =   7
         Left            =   1800
         TabIndex        =   25
         Top             =   3000
         Width           =   972
      End
      Begin VB.Label lblLiquidDiffusivity 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Liquid Diffusivity"
         ForeColor       =   &H80000008&
         Height          =   192
         Index           =   7
         Left            =   120
         TabIndex        =   24
         Top             =   3000
         Width           =   1548
      End
      Begin VB.Label lblLiquidDiffusivityUnits 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "J/mol/K"
         ForeColor       =   &H80000008&
         Height          =   192
         Index           =   2
         Left            =   2940
         TabIndex        =   23
         Top             =   2580
         Width           =   972
      End
      Begin VB.Label lblLiquidDiffusivityValue 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   192
         Index           =   6
         Left            =   1800
         TabIndex        =   22
         Top             =   2580
         Width           =   972
      End
      Begin VB.Label lblLiquidDiffusivity 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Gas Constant"
         ForeColor       =   &H80000008&
         Height          =   192
         Index           =   6
         Left            =   120
         TabIndex        =   21
         Top             =   2580
         Width           =   1548
      End
      Begin VB.Label lblLiquidDiffusivityUnits 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "cal/g/eq"
         ForeColor       =   &H80000008&
         Height          =   192
         Index           =   1
         Left            =   2940
         TabIndex        =   20
         Top             =   2400
         Width           =   972
      End
      Begin VB.Label lblLiquidDiffusivityValue 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   192
         Index           =   5
         Left            =   1800
         TabIndex        =   19
         Top             =   2400
         Width           =   972
      End
      Begin VB.Label lblLiquidDiffusivity 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Faraday's Const."
         ForeColor       =   &H80000008&
         Height          =   192
         Index           =   5
         Left            =   120
         TabIndex        =   18
         Top             =   2400
         Width           =   1548
      End
      Begin VB.Label lblLiquidDiffusivityUnits 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "K"
         ForeColor       =   &H80000008&
         Height          =   192
         Index           =   0
         Left            =   2940
         TabIndex        =   17
         Top             =   2220
         Width           =   972
      End
      Begin VB.Label lblLiquidDiffusivityValue 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   192
         Index           =   4
         Left            =   1800
         TabIndex        =   16
         Top             =   2220
         Width           =   972
      End
      Begin VB.Label lblLiquidDiffusivity 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Temperature"
         ForeColor       =   &H80000008&
         Height          =   192
         Index           =   4
         Left            =   120
         TabIndex        =   15
         Top             =   2220
         Width           =   1548
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "L.I.C. = Limiting Ionic Conductance"
         ForeColor       =   &H80000008&
         Height          =   192
         Left            =   240
         TabIndex        =   7
         Top             =   1860
         Width           =   3672
      End
      Begin VB.Shape Shape3 
         Height          =   312
         Left            =   120
         Top             =   1800
         Width           =   3852
      End
      Begin VB.Label lblLiquidDiffusivityValue 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   192
         Index           =   3
         Left            =   3060
         TabIndex        =   13
         Top             =   1560
         Width           =   792
      End
      Begin VB.Label lblLiquidDiffusivityValue 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   192
         Index           =   2
         Left            =   3060
         TabIndex        =   12
         Top             =   1320
         Width           =   792
      End
      Begin VB.Label lblLiquidDiffusivity 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "L.I.C."
         ForeColor       =   &H80000008&
         Height          =   192
         Index           =   3
         Left            =   2160
         TabIndex        =   11
         Top             =   1560
         Width           =   792
      End
      Begin VB.Label lblLiquidDiffusivity 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Valence"
         ForeColor       =   &H80000008&
         Height          =   192
         Index           =   2
         Left            =   2160
         TabIndex        =   10
         Top             =   1320
         Width           =   792
      End
      Begin VB.Label lblLiquidDiffusivityValue 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   192
         Index           =   1
         Left            =   1080
         TabIndex        =   9
         Top             =   1560
         Width           =   792
      End
      Begin VB.Label lblLiquidDiffusivityValue 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   192
         Index           =   0
         Left            =   1080
         TabIndex        =   8
         Top             =   1320
         Width           =   792
      End
      Begin VB.Label lblLiquidDiffusivity 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "L.I.C."
         ForeColor       =   &H80000008&
         Height          =   192
         Index           =   1
         Left            =   180
         TabIndex        =   6
         Top             =   1560
         Width           =   792
      End
      Begin VB.Label lblLiquidDiffusivity 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Valence"
         ForeColor       =   &H80000008&
         Height          =   192
         Index           =   0
         Left            =   180
         TabIndex        =   5
         Top             =   1320
         Width           =   792
      End
      Begin VB.Shape Shape2 
         Height          =   1152
         Left            =   2040
         Top             =   660
         Width           =   1932
      End
      Begin VB.Shape Shape1 
         Height          =   1152
         Left            =   120
         Top             =   660
         Width           =   1932
      End
      Begin VB.Label lblKineticParameters 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Cation:"
         ForeColor       =   &H80000008&
         Height          =   192
         Index           =   1
         Left            =   2100
         TabIndex        =   4
         Top             =   720
         Width           =   612
      End
      Begin VB.Label lblKineticParameters 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Anion:"
         ForeColor       =   &H80000008&
         Height          =   192
         Index           =   0
         Left            =   180
         TabIndex        =   3
         Top             =   720
         Width           =   612
      End
   End
   Begin VB.CommandButton cmdCancel 
      Appearance      =   0  'Flat
      Caption         =   "Cancel"
      Height          =   372
      Left            =   2160
      TabIndex        =   1
      Top             =   5700
      Width           =   732
   End
   Begin VB.CommandButton cmdOK 
      Appearance      =   0  'Flat
      Caption         =   "OK"
      Height          =   372
      Left            =   1200
      TabIndex        =   0
      Top             =   5700
      Width           =   732
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
      TabIndex        =   44
      Top             =   720
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
    For i = 1 To NumberOfCations
        If Trim$(Cation(i).Name) = IonToSearchFor Then
           NumberOfIonToEdit = i
           FoundCation = True
           Exit For
        End If
    Next i
    If Not FoundCation Then
       For i = 1 To NumberOfAnions
           If Trim$(Anion(i).Name) = IonToSearchFor Then
              NumberOfIonToEdit = i
              FoundAnion = True
              Exit For
           End If
       Next i
    End If

    If FoundCation Then

       EditingCation = True

       'Display NernstHaskell Cation and Anion Information on Screen for Selected Cation
       NernstHaskell.SelectedCation = Cation(NumberOfIonToEdit).Kinetic.NernstHaskellCation
       lblCation.Caption = Trim$(NernstHaskell.SelectedCation.Ion_Name)
       lblLiquidDiffusivityValue(2).Caption = "+" & Format$(NernstHaskell.SelectedCation.Valence, "0")
       lblLiquidDiffusivityValue(3).Caption = Trim$(Str$(NernstHaskell.SelectedCation.LimitingIonicConductance))

       NernstHaskell.SelectedAnion = Cation(NumberOfIonToEdit).Kinetic.NernstHaskellAnion
       lblAnion.Caption = Trim$(NernstHaskell.SelectedAnion.Ion_Name)
       lblLiquidDiffusivityValue(0).Caption = "+" & Format$(NernstHaskell.SelectedAnion.Valence, "0")
       lblLiquidDiffusivityValue(1).Caption = Trim$(Str$(NernstHaskell.SelectedAnion.LimitingIonicConductance))

       'Show current temperature in units of K on Kinetic parameters form
       frmInputKineticParameters!lblLiquidDiffusivityValue(4).Caption = Format$(Operating.Temperature, "0.00")
       
       'Show Liquid Diffusivity Values for Current Contaminant in Appropriate places
       ValueToDisplay = Cation(NumberOfIonToEdit).Kinetic.LiquidDiffusivityCorrelation
       frmInputKineticParameters!lblLiquidDiffusivityValue(7).Caption = Format$(ValueToDisplay, GetTheFormat(ValueToDisplay))

       ListIndex = cboLiquidDiffusivityUnits(0).ListIndex
       cboLiquidDiffusivityUnits(0).ListIndex = 0
       frmInputKineticParameters!lblLiquidDiffCorrelation.Caption = frmInputKineticParameters!lblLiquidDiffusivityValue(7).Caption
       cboLiquidDiffusivityUnits(0).ListIndex = ListIndex

       If Cation(NumberOfIonToEdit).Kinetic.LiquidDiffusivity.UserInput = False Then
          ListIndex = cboLiquidDiffusivityUnits(1).ListIndex
          cboLiquidDiffusivityUnits(1).ListIndex = 0
          frmInputKineticParameters!txtLiquidDiffUserInput.Text = frmInputKineticParameters!lblLiquidDiffusivityValue(7).Caption
          cboLiquidDiffusivityUnits(1).ListIndex = ListIndex
          optLiquidDiffusivity(0).Value = True
       Else
          ListIndex = cboLiquidDiffusivityUnits(1).ListIndex
          cboLiquidDiffusivityUnits(1).ListIndex = 0
          ValueToDisplay = Cation(NumberOfIonToEdit).Kinetic.LiquidDiffusivityUserInput
          frmInputKineticParameters!txtLiquidDiffUserInput.Text = Format$(ValueToDisplay, GetTheFormat(ValueToDisplay))
          cboLiquidDiffusivityUnits(1).ListIndex = ListIndex
          optLiquidDiffusivity(1).Value = True
       End If
       ValueToDisplay = Cation(NumberOfIonToEdit).Kinetic.LiquidDiffusivity.Value
       frmInputKineticParameters!lblIonicTranportCoeffValue(11).Caption = Format$(ValueToDisplay, GetTheFormat(ValueToDisplay))
          
       'Load parameters for kf calculation into label boxes

       'Particle Diameter (cm)
       ValueToDisplay = Resin.ParticleDiameter * 100
       frmInputKineticParameters!lblIonicTranportCoeffValue(0).Caption = Format$(ValueToDisplay, GetTheFormat(ValueToDisplay))

       'Column Diameter (cm)
       ValueToDisplay = Bed.Diameter * 100
       frmInputKineticParameters!lblIonicTranportCoeffValue(1).Caption = Format$(ValueToDisplay, GetTheFormat(ValueToDisplay))

       'Column Area (cm2)
       ValueToDisplay = Bed.Area * 100 ^ 2
       frmInputKineticParameters!lblIonicTranportCoeffValue(2).Caption = Format$(ValueToDisplay, GetTheFormat(ValueToDisplay))

       'Inlet Flow Rate (cm3/s)
       ValueToDisplay = Bed.FlowRate.Value * 1000000#
       frmInputKineticParameters!lblIonicTranportCoeffValue(3).Caption = Format$(ValueToDisplay, GetTheFormat(ValueToDisplay))

       'Superficial Velocity (cm/s)
       ValueToDisplay = Bed.SuperficialVelocity
       frmInputKineticParameters!lblIonicTranportCoeffValue(4).Caption = Format$(ValueToDisplay, GetTheFormat(ValueToDisplay))

       'Bed Porosity (-)
       ValueToDisplay = Bed.Porosity
       frmInputKineticParameters!lblIonicTranportCoeffValue(5).Caption = Format$(ValueToDisplay, GetTheFormat(ValueToDisplay))

       'Interstitial Velocity (cm/s)
       ValueToDisplay = Bed.InterstitialVelocity
       frmInputKineticParameters!lblIonicTranportCoeffValue(6).Caption = Format$(ValueToDisplay, GetTheFormat(ValueToDisplay))

       'Temperature (C)
       ValueToDisplay = Operating.Temperature - 273.15
       frmInputKineticParameters!lblIonicTranportCoeffValue(7).Caption = Format$(ValueToDisplay, GetTheFormat(ValueToDisplay))

       'Liquid Density (g/cm3)
       ValueToDisplay = Operating.LiquidDensity
       frmInputKineticParameters!lblIonicTranportCoeffValue(8).Caption = Format$(ValueToDisplay, GetTheFormat(ValueToDisplay))

       'Liquid Viscosity (g/cm/s)
       ValueToDisplay = Operating.LiquidViscosity
       frmInputKineticParameters!lblIonicTranportCoeffValue(9).Caption = Format$(ValueToDisplay, GetTheFormat(ValueToDisplay))

       'Reynolds Number (-)
       ValueToDisplay = Cation(NumberOfIonToEdit).Kinetic.ReynoldsNumber
       frmInputKineticParameters!lblIonicTranportCoeffValue(10).Caption = Format$(ValueToDisplay, GetTheFormat(ValueToDisplay))

       'Schmidt Number (-)
       ValueToDisplay = Cation(NumberOfIonToEdit).Kinetic.SchmidtNumber
       frmInputKineticParameters!lblIonicTranportCoeffValue(12).Caption = Format$(ValueToDisplay, GetTheFormat(ValueToDisplay))

       'Ionic Transport Coefficient, kf (cm/s)
       ValueToDisplay = Cation(NumberOfIonToEdit).Kinetic.IonicTransportCoeffCorrelation
       frmInputKineticParameters!lblIonicTranportCoeffValue(13).Caption = Format$(ValueToDisplay, GetTheFormat(ValueToDisplay))

       ListIndex = cboIonicTransportUnits(0).ListIndex
       cboIonicTransportUnits(0).ListIndex = 0
       frmInputKineticParameters!lblIonicTransportCoeffCorr.Caption = frmInputKineticParameters!lblIonicTranportCoeffValue(13).Caption
       cboIonicTransportUnits(0).ListIndex = ListIndex

       If Cation(NumberOfIonToEdit).Kinetic.IonicTransportCoefficient.UserInput = False Then
          ListIndex = cboIonicTransportUnits(1).ListIndex
          cboIonicTransportUnits(1).ListIndex = 0
          frmInputKineticParameters!txtIonicTransCoeffUser.Text = frmInputKineticParameters!lblIonicTranportCoeffValue(13).Caption
          cboIonicTransportUnits(1).ListIndex = ListIndex
          optIonicTransportCoeff(0).Value = True
       Else
          ListIndex = cboIonicTransportUnits(1).ListIndex
          cboIonicTransportUnits(1).ListIndex = 0
          ValueToDisplay = Cation(NumberOfIonToEdit).Kinetic.IonicTransportCoeffUserInput
          frmInputKineticParameters!txtIonicTransCoeffUser.Text = Format$(ValueToDisplay, GetTheFormat(ValueToDisplay))
          cboIonicTransportUnits(1).ListIndex = ListIndex
          optIonicTransportCoeff(1).Value = True
       End If

       'Pore Diffusivity information

       'Liquid Diffusivity
       ValueToDisplay = Cation(NumberOfIonToEdit).Kinetic.LiquidDiffusivity.Value
       lblPoreDiffusivityValue(0).Caption = Format$(ValueToDisplay, GetTheFormat(ValueToDisplay))

       'Tortuosity
       ValueToDisplay = Resin.Tortuosity
       lblPoreDiffusivityValue(1).Caption = Format$(ValueToDisplay, GetTheFormat(ValueToDisplay))

       'PoreDiffusivity
       
       ValueToDisplay = Cation(NumberOfIonToEdit).Kinetic.PoreDiffusivityCorrelation
       lblPoreDiffusivityValue(2).Caption = Format$(ValueToDisplay, GetTheFormat(ValueToDisplay))

       ListIndex = cboPoreDiffusivityUnits(0).ListIndex
       cboPoreDiffusivityUnits(0).ListIndex = 0
       frmInputKineticParameters!lblPoreDiffusivityCorr.Caption = lblPoreDiffusivityValue(2).Caption
       cboPoreDiffusivityUnits(0).ListIndex = ListIndex

       If Cation(NumberOfIonToEdit).Kinetic.PoreDiffusivity.UserInput = False Then
          ListIndex = cboPoreDiffusivityUnits(1).ListIndex
          cboPoreDiffusivityUnits(1).ListIndex = 0
          frmInputKineticParameters!txtPoreDiffusivityUser.Text = frmInputKineticParameters!lblPoreDiffusivityValue(2).Caption
          cboPoreDiffusivityUnits(1).ListIndex = ListIndex
          optPoreDiffusivity(0).Value = True
       Else
          ListIndex = cboPoreDiffusivityUnits(1).ListIndex
          cboPoreDiffusivityUnits(1).ListIndex = 0
          ValueToDisplay = Cation(NumberOfIonToEdit).Kinetic.PoreDiffusivityUserInput
          frmInputKineticParameters!txtPoreDiffusivityUser.Text = Format$(ValueToDisplay, GetTheFormat(ValueToDisplay))
          cboPoreDiffusivityUnits(1).ListIndex = ListIndex
          optPoreDiffusivity(1).Value = True
       End If

    ElseIf FoundAnion Then

       EditingAnion = True

       'Display NernstHaskell Cation and Anion Information on Screen for Selected Anion
       NernstHaskell.SelectedCation = Anion(NumberOfIonToEdit).Kinetic.NernstHaskellCation
       lblCation.Caption = Trim$(NernstHaskell.SelectedCation.Ion_Name)
       lblLiquidDiffusivityValue(2).Caption = "+" & Format$(NernstHaskell.SelectedCation.Valence, "0")
       lblLiquidDiffusivityValue(3).Caption = Trim$(Str$(NernstHaskell.SelectedCation.LimitingIonicConductance))

       NernstHaskell.SelectedAnion = Anion(NumberOfIonToEdit).Kinetic.NernstHaskellAnion
       lblAnion.Caption = Trim$(NernstHaskell.SelectedAnion.Ion_Name)
       lblLiquidDiffusivityValue(0).Caption = "+" & Format$(NernstHaskell.SelectedAnion.Valence, "0")
       lblLiquidDiffusivityValue(1).Caption = Trim$(Str$(NernstHaskell.SelectedAnion.LimitingIonicConductance))

       'Show current temperature in units of K on Kinetic parameters form
       frmInputKineticParameters!lblLiquidDiffusivityValue(4).Caption = Format$(Operating.Temperature, "0.00")

       'Show Liquid Diffusivity Values for Current Contaminant in Appropriate places
       ValueToDisplay = Anion(NumberOfIonToEdit).Kinetic.LiquidDiffusivityCorrelation
       frmInputKineticParameters!lblLiquidDiffusivityValue(7).Caption = Format$(ValueToDisplay, GetTheFormat(ValueToDisplay))

       ListIndex = cboLiquidDiffusivityUnits(0).ListIndex
       cboLiquidDiffusivityUnits(0).ListIndex = 0
       frmInputKineticParameters!lblLiquidDiffCorrelation.Caption = frmInputKineticParameters!lblLiquidDiffusivityValue(7).Caption
       cboLiquidDiffusivityUnits(0).ListIndex = ListIndex

       If Anion(NumberOfIonToEdit).Kinetic.LiquidDiffusivity.UserInput = False Then
          ListIndex = cboLiquidDiffusivityUnits(1).ListIndex
          cboLiquidDiffusivityUnits(1).ListIndex = 0
          frmInputKineticParameters!txtLiquidDiffUserInput.Text = frmInputKineticParameters!lblLiquidDiffusivityValue(7).Caption
          cboLiquidDiffusivityUnits(1).ListIndex = ListIndex
          optLiquidDiffusivity(0).Value = True
       Else
          ListIndex = cboLiquidDiffusivityUnits(1).ListIndex
          cboLiquidDiffusivityUnits(1).ListIndex = 0
          ValueToDisplay = Anion(NumberOfIonToEdit).Kinetic.LiquidDiffusivityUserInput
          frmInputKineticParameters!txtLiquidDiffUserInput.Text = Format$(ValueToDisplay, GetTheFormat(ValueToDisplay))
          cboLiquidDiffusivityUnits(1).ListIndex = ListIndex
          optLiquidDiffusivity(1).Value = True
       End If
       ValueToDisplay = Anion(NumberOfIonToEdit).Kinetic.LiquidDiffusivity.Value
       frmInputKineticParameters!lblIonicTranportCoeffValue(11).Caption = Format$(ValueToDisplay, GetTheFormat(ValueToDisplay))

       'Load parameters for kf calculation into label boxes

       'Particle Diameter (cm)
       ValueToDisplay = Resin.ParticleDiameter * 100
       frmInputKineticParameters!lblIonicTranportCoeffValue(0).Caption = Format$(ValueToDisplay, GetTheFormat(ValueToDisplay))

       'Column Diameter (cm)
       ValueToDisplay = Bed.Diameter * 100
       frmInputKineticParameters!lblIonicTranportCoeffValue(1).Caption = Format$(ValueToDisplay, GetTheFormat(ValueToDisplay))

       'Column Area (cm2)
       ValueToDisplay = Bed.Area * 100 ^ 2
       frmInputKineticParameters!lblIonicTranportCoeffValue(2).Caption = Format$(ValueToDisplay, GetTheFormat(ValueToDisplay))

       'Inlet Flow Rate (cm3/s)
       ValueToDisplay = Bed.FlowRate.Value * 1000000#
       frmInputKineticParameters!lblIonicTranportCoeffValue(3).Caption = Format$(ValueToDisplay, GetTheFormat(ValueToDisplay))

       'Superficial Velocity (cm/s)
       ValueToDisplay = Bed.SuperficialVelocity
       frmInputKineticParameters!lblIonicTranportCoeffValue(4).Caption = Format$(ValueToDisplay, GetTheFormat(ValueToDisplay))

       'Bed Porosity (-)
       ValueToDisplay = Bed.Porosity
       frmInputKineticParameters!lblIonicTranportCoeffValue(5).Caption = Format$(ValueToDisplay, GetTheFormat(ValueToDisplay))

       'Interstitial Velocity (cm/s)
       ValueToDisplay = Bed.InterstitialVelocity
       frmInputKineticParameters!lblIonicTranportCoeffValue(6).Caption = Format$(ValueToDisplay, GetTheFormat(ValueToDisplay))

       'Temperature (C)
       ValueToDisplay = Operating.Temperature - 273.15
       frmInputKineticParameters!lblIonicTranportCoeffValue(7).Caption = Format$(ValueToDisplay, GetTheFormat(ValueToDisplay))

       'Liquid Density (g/cm3)
       ValueToDisplay = Operating.LiquidDensity
       frmInputKineticParameters!lblIonicTranportCoeffValue(8).Caption = Format$(ValueToDisplay, GetTheFormat(ValueToDisplay))

       'Liquid Viscosity (g/cm/s)
       ValueToDisplay = Operating.LiquidViscosity
       frmInputKineticParameters!lblIonicTranportCoeffValue(9).Caption = Format$(ValueToDisplay, GetTheFormat(ValueToDisplay))

       'Reynolds Number (-)
       ValueToDisplay = Anion(NumberOfIonToEdit).Kinetic.ReynoldsNumber
       frmInputKineticParameters!lblIonicTranportCoeffValue(10).Caption = Format$(ValueToDisplay, GetTheFormat(ValueToDisplay))

       'Schmidt Number (-)
       ValueToDisplay = Anion(NumberOfIonToEdit).Kinetic.SchmidtNumber
       frmInputKineticParameters!lblIonicTranportCoeffValue(12).Caption = Format$(ValueToDisplay, GetTheFormat(ValueToDisplay))

       'Ionic Transport Coefficient, kf (cm/s)
       ValueToDisplay = Anion(NumberOfIonToEdit).Kinetic.IonicTransportCoeffCorrelation
       frmInputKineticParameters!lblIonicTranportCoeffValue(13).Caption = Format$(ValueToDisplay, GetTheFormat(ValueToDisplay))

       ListIndex = cboIonicTransportUnits(0).ListIndex
       cboIonicTransportUnits(0).ListIndex = 0
       frmInputKineticParameters!lblIonicTransportCoeffCorr.Caption = frmInputKineticParameters!lblIonicTranportCoeffValue(13).Caption
       cboIonicTransportUnits(0).ListIndex = ListIndex

       If Anion(NumberOfIonToEdit).Kinetic.IonicTransportCoefficient.UserInput = False Then
          ListIndex = cboIonicTransportUnits(1).ListIndex
          cboIonicTransportUnits(1).ListIndex = 0
          frmInputKineticParameters!txtIonicTransCoeffUser.Text = frmInputKineticParameters!lblIonicTranportCoeffValue(13).Caption
          cboIonicTransportUnits(1).ListIndex = ListIndex
          optIonicTransportCoeff(0).Value = True
       Else
          ListIndex = cboIonicTransportUnits(1).ListIndex
          cboIonicTransportUnits(1).ListIndex = 0
          ValueToDisplay = Anion(NumberOfIonToEdit).Kinetic.IonicTransportCoeffUserInput
          frmInputKineticParameters!txtIonicTransCoeffUser.Text = Format$(ValueToDisplay, GetTheFormat(ValueToDisplay))
          cboIonicTransportUnits(1).ListIndex = ListIndex
          optIonicTransportCoeff(1).Value = True
       End If

       'Pore Diffusivity information

       'Liquid Diffusivity
       ValueToDisplay = Anion(NumberOfIonToEdit).Kinetic.LiquidDiffusivity.Value
       lblPoreDiffusivityValue(0).Caption = Format$(ValueToDisplay, GetTheFormat(ValueToDisplay))

       'Tortuosity
       ValueToDisplay = Resin.Tortuosity
       lblPoreDiffusivityValue(1).Caption = Format$(ValueToDisplay, GetTheFormat(ValueToDisplay))

       'Pore Diffusivity
       
       ValueToDisplay = Anion(NumberOfIonToEdit).Kinetic.PoreDiffusivityCorrelation
       lblPoreDiffusivityValue(2).Caption = Format$(ValueToDisplay, GetTheFormat(ValueToDisplay))

       ListIndex = cboPoreDiffusivityUnits(0).ListIndex
       cboPoreDiffusivityUnits(0).ListIndex = 0
       frmInputKineticParameters!lblPoreDiffusivityCorr.Caption = lblPoreDiffusivityValue(2).Caption
       cboPoreDiffusivityUnits(0).ListIndex = ListIndex

       If Anion(NumberOfIonToEdit).Kinetic.PoreDiffusivity.UserInput = False Then
          ListIndex = cboPoreDiffusivityUnits(1).ListIndex
          cboPoreDiffusivityUnits(1).ListIndex = 0
          frmInputKineticParameters!txtPoreDiffusivityUser.Text = frmInputKineticParameters!lblPoreDiffusivityValue(2).Caption
          cboPoreDiffusivityUnits(1).ListIndex = ListIndex
          optPoreDiffusivity(0).Value = True
       Else
          ListIndex = cboPoreDiffusivityUnits(1).ListIndex
          cboPoreDiffusivityUnits(1).ListIndex = 0
          ValueToDisplay = Anion(NumberOfIonToEdit).Kinetic.PoreDiffusivityUserInput
          frmInputKineticParameters!txtPoreDiffusivityUser.Text = Format$(ValueToDisplay, GetTheFormat(ValueToDisplay))
          cboPoreDiffusivityUnits(1).ListIndex = ListIndex
          optPoreDiffusivity(1).Value = True
       End If

    End If

    If FoundCation Or FoundAnion Then
       If FoundCation Then
          If NumSelectedCations = 1 Then GoTo ExitSub
          For i = 1 To NumSelectedCations
             If Cations_Selected(i) = NumberOfIonToEdit Then
                Call CalculateDimensionlessGroups
                frmIonExchangeMain!cboKinDimComponent.ListIndex = -1
                frmIonExchangeMain!cboKinDimComponent.ListIndex = i - 1
             End If
          Next i
       End If

       If FoundAnion Then
          If NumSelectedAnions = 1 Then GoTo ExitSub
          For i = 1 To NumSelectedAnions
             If Anions_Selected(i) = NumberOfIonToEdit Then
                Call CalculateDimensionlessGroups
                frmIonExchangeMain!cboKinDimComponent.ListIndex = -1
                frmIonExchangeMain!cboKinDimComponent.ListIndex = i - 1
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
          If IonicTransportCoeffCorrName = IONIC_TRANSPORT_COEFFICIENT_1 Then Exit Sub
          IonicTransportCoeffCorrName = IONIC_TRANSPORT_COEFFICIENT_1
       Case 1   'Gnielinski correlation
          If IonicTransportCoeffCorrName = IONIC_TRANSPORT_COEFFICIENT_2 Then Exit Sub
          IonicTransportCoeffCorrName = IONIC_TRANSPORT_COEFFICIENT_2
    End Select

    If (NumberOfCations = 0) And (NumberOfAnions = 0) Then Exit Sub

    AddingCation = True
    AddingAnion = False
    For i = 1 To NumberOfCations
        NumberOfIonToEdit = i
        Call CalculateKineticParameters
    Next i
    AddingCation = False
    AddingAnion = True
    For i = 1 To NumberOfAnions
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
             ValueToConvert = Cation(NumberOfIonToEdit).Kinetic.IonicTransportCoeffCorrelation
          ElseIf EditingAnion Then
             ValueToConvert = Anion(NumberOfIonToEdit).Kinetic.IonicTransportCoeffCorrelation
          End If
          ListIndex = cboIonicTransportUnits(0).ListIndex
       Case 1   'User Input
          If EditingCation Then
             ValueToConvert = Cation(NumberOfIonToEdit).Kinetic.IonicTransportCoeffUserInput
          ElseIf EditingAnion Then
             ValueToConvert = Anion(NumberOfIonToEdit).Kinetic.IonicTransportCoeffUserInput
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
             lblIonicTransportCoeffCorr.Caption = Format$(ValueToDisplay, GetTheFormat(ValueToDisplay))
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
             ValueToConvert = Cation(NumberOfIonToEdit).Kinetic.LiquidDiffusivityCorrelation
          ElseIf EditingAnion Then
             ValueToConvert = Anion(NumberOfIonToEdit).Kinetic.LiquidDiffusivityCorrelation
          End If
          ListIndex = cboLiquidDiffusivityUnits(0).ListIndex
       Case 1   'User Input
          If EditingCation Then
             ValueToConvert = Cation(NumberOfIonToEdit).Kinetic.LiquidDiffusivityUserInput
          ElseIf EditingAnion Then
             ValueToConvert = Anion(NumberOfIonToEdit).Kinetic.LiquidDiffusivityUserInput
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
             lblLiquidDiffCorrelation.Caption = Format$(ValueToDisplay, GetTheFormat(ValueToDisplay))
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
             ValueToConvert = Cation(NumberOfIonToEdit).Kinetic.PoreDiffusivityCorrelation
          ElseIf EditingAnion Then
             ValueToConvert = Anion(NumberOfIonToEdit).Kinetic.PoreDiffusivityCorrelation
          End If
          ListIndex = cboPoreDiffusivityUnits(0).ListIndex
       Case 1   'User Input
          If EditingCation Then
             ValueToConvert = Cation(NumberOfIonToEdit).Kinetic.PoreDiffusivityUserInput
          ElseIf EditingAnion Then
             ValueToConvert = Anion(NumberOfIonToEdit).Kinetic.PoreDiffusivityUserInput
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
             lblPoreDiffusivityCorr.Caption = Format$(ValueToDisplay, GetTheFormat(ValueToDisplay))
          Case 1   'User Input
             txtPoreDiffusivityUser.Text = Format$(ValueToDisplay, GetTheFormat(ValueToDisplay))
       End Select

End Sub

Private Sub cmdCancel_Click()
    Dim i As Integer
    Dim ListIndex As Integer

    If EditingCation And EditingAnion Then
       For i = 1 To NumberOfCations
           Cation(i).Kinetic = OldCationKineticParameters(i)
       Next i

       For i = 1 To NumberOfAnions
           Anion(i).Kinetic = OldAnionKineticParameters(i)
       Next i

    ElseIf EditingCation Then
       For i = 1 To NumberOfCations
           Cation(i).Kinetic = OldCationKineticParameters(i)
       Next i

    ElseIf EditingAnion Then
       For i = 1 To NumberOfAnions
           Anion(i).Kinetic = OldAnionKineticParameters(i)
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

End Sub

Private Sub Form_Load()
    Dim PositionLeft As Integer

    frmInputKineticParameters.WindowState = 0

    frmInputKineticParameters.width = 4440
    frmInputKineticParameters.height = 6500

    'Position the form on the screen (Centered on left half of it)
    If WindowState = 0 Then
       'don't attempt if screen Minimized or Maximized
       PositionLeft = ((Screen.width / 2 - frmIonExchangeMain.left) / 2) - frmInputKineticParameters.width / 2
       Move (frmIonExchangeMain.left + PositionLeft), (Screen.height - frmInputKineticParameters.height) / 2

    End If

    Call LoadLiquidDiffusivityParameters

    'Load Liquid Diffusivity Units with superscripts
    lblLiquidDiffusivityUnits(3).Caption = "cm" & Chr$(178) & "/s"
    lblIonicTransportCoeffUnits(2).Caption = "cm" & Chr$(178)
    lblIonicTransportCoeffUnits(3).Caption = "cm" & Chr$(179) & "/s"
    lblIonicTransportCoeffUnits(8).Caption = "g/cm" & Chr$(179)
    lblIonicTransportCoeffUnits(11).Caption = "cm" & Chr$(178) & "/s"
    lblPoreDiffusivityUnits(0).Caption = "cm" & Chr$(178) & "/s"
    lblPoreDiffusivityUnits(2).Caption = "cm" & Chr$(178) & "/s"

    cboIonicTransport.AddItem IONIC_TRANSPORT_COEFFICIENT_1
    cboIonicTransport.AddItem IONIC_TRANSPORT_COEFFICIENT_2
    
    If IonicTransportCoeffCorrName = IONIC_TRANSPORT_COEFFICIENT_1 Then
       cboIonicTransport.ListIndex = 0
    ElseIf IonicTransportCoeffCorrName = IONIC_TRANSPORT_COEFFICIENT_2 Then
       cboIonicTransport.ListIndex = 1
    End If

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
          lblIonicTransportCoeffCorr.Enabled = True
          If EditingCation Then
             Cation(NumberOfIonToEdit).Kinetic.IonicTransportCoefficient.Value = Cation(NumberOfIonToEdit).Kinetic.IonicTransportCoeffCorrelation
             Cation(NumberOfIonToEdit).Kinetic.IonicTransportCoefficient.UserInput = False
          ElseIf EditingAnion Then
             Anion(NumberOfIonToEdit).Kinetic.IonicTransportCoefficient.Value = Anion(NumberOfIonToEdit).Kinetic.IonicTransportCoeffCorrelation
             Anion(NumberOfIonToEdit).Kinetic.IonicTransportCoefficient.UserInput = False
          End If
       Case 1   'User Input
          lblIonicTransportCoeffCorr.Enabled = False
          txtIonicTransCoeffUser.Enabled = True
          If EditingCation Then
             Cation(NumberOfIonToEdit).Kinetic.IonicTransportCoefficient.Value = Cation(NumberOfIonToEdit).Kinetic.IonicTransportCoeffUserInput
             Cation(NumberOfIonToEdit).Kinetic.IonicTransportCoefficient.UserInput = True
          ElseIf EditingAnion Then
             Anion(NumberOfIonToEdit).Kinetic.IonicTransportCoefficient.Value = Anion(NumberOfIonToEdit).Kinetic.IonicTransportCoeffUserInput
             Anion(NumberOfIonToEdit).Kinetic.IonicTransportCoefficient.UserInput = True
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
          lblLiquidDiffCorrelation.Enabled = True
          If EditingCation Then
             Cation(NumberOfIonToEdit).Kinetic.LiquidDiffusivity.Value = Cation(NumberOfIonToEdit).Kinetic.LiquidDiffusivityCorrelation
             Cation(NumberOfIonToEdit).Kinetic.LiquidDiffusivity.UserInput = False
          ElseIf EditingAnion Then
             Anion(NumberOfIonToEdit).Kinetic.LiquidDiffusivity.Value = Anion(NumberOfIonToEdit).Kinetic.LiquidDiffusivityCorrelation
             Anion(NumberOfIonToEdit).Kinetic.LiquidDiffusivity.UserInput = False
          End If
       Case 1   'User Input
          lblLiquidDiffCorrelation.Enabled = False
          txtLiquidDiffUserInput.Enabled = True
          If EditingCation Then
             Cation(NumberOfIonToEdit).Kinetic.LiquidDiffusivity.Value = Cation(NumberOfIonToEdit).Kinetic.LiquidDiffusivityUserInput
             Cation(NumberOfIonToEdit).Kinetic.LiquidDiffusivity.UserInput = True
          ElseIf EditingAnion Then
             Anion(NumberOfIonToEdit).Kinetic.LiquidDiffusivity.Value = Anion(NumberOfIonToEdit).Kinetic.LiquidDiffusivityUserInput
             Anion(NumberOfIonToEdit).Kinetic.LiquidDiffusivity.UserInput = True
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
          lblPoreDiffusivityCorr.Enabled = True
          If EditingCation Then
             Cation(NumberOfIonToEdit).Kinetic.PoreDiffusivity.Value = Cation(NumberOfIonToEdit).Kinetic.PoreDiffusivityCorrelation
             Cation(NumberOfIonToEdit).Kinetic.PoreDiffusivity.UserInput = False
          ElseIf EditingAnion Then
             Anion(NumberOfIonToEdit).Kinetic.PoreDiffusivity.Value = Anion(NumberOfIonToEdit).Kinetic.PoreDiffusivityCorrelation
             Anion(NumberOfIonToEdit).Kinetic.PoreDiffusivity.UserInput = False
          End If
       Case 1   'User Input
          lblPoreDiffusivityCorr.Enabled = False
          txtPoreDiffusivityUser.Enabled = True
          If EditingCation Then
             Cation(NumberOfIonToEdit).Kinetic.PoreDiffusivity.Value = Cation(NumberOfIonToEdit).Kinetic.PoreDiffusivityUserInput
             Cation(NumberOfIonToEdit).Kinetic.PoreDiffusivity.UserInput = True
          ElseIf EditingAnion Then
             Anion(NumberOfIonToEdit).Kinetic.PoreDiffusivity.Value = Anion(NumberOfIonToEdit).Kinetic.PoreDiffusivityUserInput
             Anion(NumberOfIonToEdit).Kinetic.PoreDiffusivity.UserInput = True
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
          OldValue = Cation(NumberOfIonToEdit).Kinetic.IonicTransportCoeffUserInput
       ElseIf EditingAnion Then
          OldValue = Anion(NumberOfIonToEdit).Kinetic.IonicTransportCoeffUserInput
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
                Cation(NumberOfIonToEdit).Kinetic.IonicTransportCoeffUserInput = NewValue
                Cation(NumberOfIonToEdit).Kinetic.IonicTransportCoefficient.Value = NewValue
             ElseIf EditingAnion Then
                Anion(NumberOfIonToEdit).Kinetic.IonicTransportCoeffUserInput = NewValue
                Anion(NumberOfIonToEdit).Kinetic.IonicTransportCoefficient.Value = NewValue
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
          OldValue = Cation(NumberOfIonToEdit).Kinetic.LiquidDiffusivityUserInput
       ElseIf EditingAnion Then
          OldValue = Anion(NumberOfIonToEdit).Kinetic.LiquidDiffusivityUserInput
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
                Cation(NumberOfIonToEdit).Kinetic.LiquidDiffusivityUserInput = NewValue
                Cation(NumberOfIonToEdit).Kinetic.LiquidDiffusivity.Value = NewValue
             ElseIf EditingAnion Then
                Anion(NumberOfIonToEdit).Kinetic.LiquidDiffusivityUserInput = NewValue
                Anion(NumberOfIonToEdit).Kinetic.LiquidDiffusivity.Value = NewValue
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
          OldValue = Cation(NumberOfIonToEdit).Kinetic.PoreDiffusivityUserInput
       ElseIf EditingAnion Then
          OldValue = Anion(NumberOfIonToEdit).Kinetic.PoreDiffusivityUserInput
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
                Cation(NumberOfIonToEdit).Kinetic.PoreDiffusivityUserInput = NewValue
                Cation(NumberOfIonToEdit).Kinetic.PoreDiffusivity.Value = NewValue
             ElseIf EditingAnion Then
                Anion(NumberOfIonToEdit).Kinetic.PoreDiffusivityUserInput = NewValue
                Anion(NumberOfIonToEdit).Kinetic.PoreDiffusivity.Value = NewValue
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

