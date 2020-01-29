VERSION 5.00
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "THREED32.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmPTADScreen2 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Packed Tower Aeration - Rating Mode"
   ClientHeight    =   6780
   ClientLeft      =   870
   ClientTop       =   2115
   ClientWidth     =   9480
   Icon            =   "PTADScreen2.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   ScaleHeight     =   6780
   ScaleWidth      =   9480
   Begin VB.CommandButton Command1 
      Appearance      =   0  'Flat
      Caption         =   "Concentration Results"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   960
      TabIndex        =   80
      TabStop         =   0   'False
      Top             =   6000
      Width           =   2655
   End
   Begin Threed.SSFrame Frame1 
      Height          =   1095
      Left            =   60
      TabIndex        =   11
      Top             =   120
      Width           =   4455
      _Version        =   65536
      _ExtentX        =   7858
      _ExtentY        =   1931
      _StockProps     =   14
      Caption         =   "Design Based On:"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Begin VB.ComboBox UnitsDesignBasis 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   0
         Left            =   3060
         Style           =   2  'Dropdown List
         TabIndex        =   18
         TabStop         =   0   'False
         Top             =   300
         Width           =   1275
      End
      Begin VB.ComboBox UnitsDesignBasis 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   1
         Left            =   3060
         Style           =   2  'Dropdown List
         TabIndex        =   17
         TabStop         =   0   'False
         Top             =   660
         Width           =   1275
      End
      Begin VB.Label lblDesignParameters 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
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
         Height          =   255
         Index           =   1
         Left            =   1740
         TabIndex        =   22
         Top             =   660
         Width           =   1215
      End
      Begin VB.Label lblDesignParameters 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
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
         Height          =   255
         Index           =   0
         Left            =   1740
         TabIndex        =   21
         Top             =   300
         Width           =   1215
      End
      Begin VB.Label lblDesignParametersLabel 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Tower Height"
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
         Height          =   255
         Index           =   1
         Left            =   120
         TabIndex        =   20
         Top             =   660
         Width           =   1515
      End
      Begin VB.Label lblDesignParametersLabel 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Tower Diameter"
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
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   19
         Top             =   300
         Width           =   1515
      End
   End
   Begin Threed.SSFrame Frame3D1 
      Height          =   1815
      Left            =   60
      TabIndex        =   12
      Top             =   1290
      Width           =   4455
      _Version        =   65536
      _ExtentX        =   7858
      _ExtentY        =   3201
      _StockProps     =   14
      Caption         =   "Tower Parameters:"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Begin VB.TextBox txtTowerParameters 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   0
         Left            =   1740
         TabIndex        =   0
         Top             =   300
         Width           =   1215
      End
      Begin VB.TextBox txtTowerParameters 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   1
         Left            =   1740
         TabIndex        =   1
         Top             =   660
         Width           =   1215
      End
      Begin VB.ComboBox UnitsTowerParam 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   0
         Left            =   3060
         Style           =   2  'Dropdown List
         TabIndex        =   26
         TabStop         =   0   'False
         Top             =   300
         Width           =   1275
      End
      Begin VB.ComboBox UnitsTowerParam 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   1
         Left            =   3060
         Style           =   2  'Dropdown List
         TabIndex        =   25
         TabStop         =   0   'False
         Top             =   660
         Width           =   1275
      End
      Begin VB.ComboBox UnitsTowerParam 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   2
         Left            =   3060
         Style           =   2  'Dropdown List
         TabIndex        =   24
         TabStop         =   0   'False
         Top             =   1020
         Width           =   1275
      End
      Begin VB.ComboBox UnitsTowerParam 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   3
         Left            =   3060
         Style           =   2  'Dropdown List
         TabIndex        =   23
         TabStop         =   0   'False
         Top             =   1380
         Width           =   1275
      End
      Begin VB.Label lblTowerParametersLabel 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Tower Volume"
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
         Height          =   255
         Index           =   3
         Left            =   -1020
         TabIndex        =   32
         Top             =   1380
         Width           =   2655
      End
      Begin VB.Label lblTowerParametersLabel 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Tower Area"
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
         Height          =   255
         Index           =   2
         Left            =   -1020
         TabIndex        =   31
         Top             =   1020
         Width           =   2655
      End
      Begin VB.Label lblTowerParametersLabel 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Specify Diameter"
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
         Height          =   255
         Index           =   0
         Left            =   -1020
         TabIndex        =   30
         Top             =   300
         Width           =   2655
      End
      Begin VB.Label lblTowerParametersLabel 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Specify Height"
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
         Height          =   255
         Index           =   1
         Left            =   -1020
         TabIndex        =   29
         Top             =   660
         Width           =   2655
      End
      Begin VB.Label lblTowerParameters 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
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
         Height          =   255
         Index           =   2
         Left            =   1740
         TabIndex        =   28
         Top             =   1020
         Width           =   1215
      End
      Begin VB.Label lblTowerParameters 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
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
         Height          =   255
         Index           =   3
         Left            =   1740
         TabIndex        =   27
         Top             =   1380
         Width           =   1215
      End
   End
   Begin Threed.SSFrame fraOperatingConditions 
      Height          =   1575
      Left            =   60
      TabIndex        =   13
      Top             =   3180
      Width           =   4455
      _Version        =   65536
      _ExtentX        =   7858
      _ExtentY        =   2778
      _StockProps     =   14
      Caption         =   "Operating Conditions:"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Begin VB.TextBox txtOperatingPressure 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   1740
         TabIndex        =   2
         Top             =   330
         Width           =   1215
      End
      Begin VB.TextBox txtOperatingTemperature 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   1740
         TabIndex        =   3
         Top             =   690
         Width           =   1215
      End
      Begin VB.CommandButton lblDisplayAirWaterProperties 
         Appearance      =   0  'Flat
         Caption         =   "Display Physical Properties of Air and Water"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   180
         TabIndex        =   35
         TabStop         =   0   'False
         Top             =   1050
         Width           =   4155
      End
      Begin VB.ComboBox UnitsOpCond 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   0
         Left            =   3060
         Style           =   2  'Dropdown List
         TabIndex        =   34
         TabStop         =   0   'False
         Top             =   330
         Width           =   1275
      End
      Begin VB.ComboBox UnitsOpCond 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   1
         Left            =   3060
         Style           =   2  'Dropdown List
         TabIndex        =   33
         TabStop         =   0   'False
         Top             =   690
         Width           =   1275
      End
      Begin VB.Label lblOperatingPressure 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Pressure"
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
         Height          =   255
         Left            =   -1020
         TabIndex        =   37
         Top             =   330
         Width           =   2655
      End
      Begin VB.Label lblOperatingTemperature 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Temperature"
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
         Height          =   255
         Left            =   -1020
         TabIndex        =   36
         Top             =   690
         Width           =   2655
      End
   End
   Begin Threed.SSFrame fraPacking 
      Height          =   1095
      Left            =   60
      TabIndex        =   14
      Top             =   4830
      Width           =   4455
      _Version        =   65536
      _ExtentX        =   7858
      _ExtentY        =   1931
      _StockProps     =   14
      Caption         =   "Packing:"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Begin VB.CommandButton cmdSelectPacking 
         Appearance      =   0  'Flat
         Caption         =   "Select Packing Type"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   180
         TabIndex        =   39
         TabStop         =   0   'False
         Top             =   570
         Width           =   4155
      End
      Begin VB.CommandButton cmdPackingType 
         Appearance      =   0  'Flat
         Caption         =   "Type"
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
         Left            =   180
         TabIndex        =   38
         TabStop         =   0   'False
         Top             =   270
         Width           =   855
      End
      Begin VB.Label lblPackingTypeLabel 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Type"
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
         Height          =   255
         Left            =   180
         TabIndex        =   41
         Top             =   270
         Width           =   735
      End
      Begin VB.Label lblPackingType 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
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
         Height          =   255
         Left            =   1140
         TabIndex        =   40
         Top             =   270
         Width           =   3195
      End
   End
   Begin Threed.SSFrame fraFlowsAndLoadings 
      Height          =   2295
      Left            =   4500
      TabIndex        =   15
      Top             =   120
      Width           =   4875
      _Version        =   65536
      _ExtentX        =   8599
      _ExtentY        =   4048
      _StockProps     =   14
      Caption         =   "Flows and Loadings:"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Begin VB.CommandButton cmdPickFlowLoadingParameters 
         Appearance      =   0  'Flat
         Caption         =   "Pick Flow and Loading Parameters to Specify"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   180
         TabIndex        =   46
         TabStop         =   0   'False
         Top             =   240
         Width           =   4335
      End
      Begin VB.TextBox txtFlowsLoadings 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   4
         Left            =   2160
         TabIndex        =   8
         TabStop         =   0   'False
         Top             =   1860
         Width           =   1215
      End
      Begin VB.TextBox txtFlowsLoadings 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   0
         Left            =   2160
         TabIndex        =   4
         Top             =   660
         Width           =   1215
      End
      Begin VB.TextBox txtFlowsLoadings 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   1
         Left            =   2160
         TabIndex        =   5
         Top             =   960
         Width           =   1215
      End
      Begin VB.TextBox txtFlowsLoadings 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   2
         Left            =   2160
         TabIndex        =   6
         TabStop         =   0   'False
         Top             =   1260
         Width           =   1215
      End
      Begin VB.TextBox txtFlowsLoadings 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   3
         Left            =   2160
         TabIndex        =   7
         TabStop         =   0   'False
         Top             =   1560
         Width           =   1215
      End
      Begin VB.ComboBox UnitsFlows 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   0
         Left            =   3480
         Style           =   2  'Dropdown List
         TabIndex        =   45
         TabStop         =   0   'False
         Top             =   660
         Width           =   1275
      End
      Begin VB.ComboBox UnitsFlows 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   1
         Left            =   3480
         Style           =   2  'Dropdown List
         TabIndex        =   44
         TabStop         =   0   'False
         Top             =   960
         Width           =   1275
      End
      Begin VB.ComboBox UnitsFlows 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   3
         Left            =   3480
         Style           =   2  'Dropdown List
         TabIndex        =   43
         TabStop         =   0   'False
         Top             =   1560
         Width           =   1275
      End
      Begin VB.ComboBox UnitsFlows 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   4
         Left            =   3480
         Style           =   2  'Dropdown List
         TabIndex        =   42
         TabStop         =   0   'False
         Top             =   1860
         Width           =   1275
      End
      Begin VB.Label lblFlowsLoadingsLabel 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Water Flow Rate"
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
         Height          =   255
         Index           =   0
         Left            =   -600
         TabIndex        =   52
         Top             =   660
         Width           =   2655
      End
      Begin VB.Label lblFlowsLoadingsLabel 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Air to Water Ratio"
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
         Height          =   255
         Index           =   2
         Left            =   -600
         TabIndex        =   51
         Top             =   1260
         Width           =   2655
      End
      Begin VB.Label lblFlowsLoadingsLabel 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Air Flow Rate"
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
         Height          =   255
         Index           =   1
         Left            =   -600
         TabIndex        =   50
         Top             =   960
         Width           =   2655
      End
      Begin VB.Label lblFlowsLoadingsLabel 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Air Loading Rate"
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
         Height          =   255
         Index           =   4
         Left            =   -600
         TabIndex        =   49
         Top             =   1860
         Width           =   2655
      End
      Begin VB.Label lblFlowsLoadingsLabel 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Water Loading Rate"
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
         Height          =   255
         Index           =   3
         Left            =   -600
         TabIndex        =   48
         Top             =   1560
         Width           =   2655
      End
      Begin VB.Label lblMisc 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "( - )"
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
         Height          =   255
         Index           =   1
         Left            =   3480
         TabIndex        =   47
         Top             =   1260
         Width           =   1275
      End
   End
   Begin Threed.SSFrame Frame3D2 
      Height          =   3705
      Left            =   4500
      TabIndex        =   16
      Top             =   2580
      Width           =   4875
      _Version        =   65536
      _ExtentX        =   8599
      _ExtentY        =   6535
      _StockProps     =   14
      Caption         =   "Contaminant of Interest:"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Begin VB.CommandButton cmdEditComponent 
         Appearance      =   0  'Flat
         Caption         =   "&Edit Properties"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         HelpContextID   =   20
         Left            =   2760
         TabIndex        =   63
         TabStop         =   0   'False
         Top             =   690
         Width           =   1875
      End
      Begin VB.CommandButton cmdDeleteComponent 
         Appearance      =   0  'Flat
         Caption         =   "De&lete"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         HelpContextID   =   20
         Left            =   1560
         TabIndex        =   62
         TabStop         =   0   'False
         Top             =   690
         Width           =   1095
      End
      Begin VB.CommandButton cmdAddComponent 
         Appearance      =   0  'Flat
         Caption         =   "&Add"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         HelpContextID   =   20
         Left            =   300
         TabIndex        =   61
         TabStop         =   0   'False
         Top             =   690
         Width           =   1155
      End
      Begin VB.ComboBox cboSelectCompo 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   300
         Style           =   2  'Dropdown List
         TabIndex        =   60
         TabStop         =   0   'False
         Top             =   330
         Width           =   4335
      End
      Begin VB.TextBox txtDesignConcentrationValue 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   1
         Left            =   2280
         TabIndex        =   9
         Top             =   1470
         Width           =   1095
      End
      Begin VB.TextBox txtDesignConcentrationValue 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   2
         Left            =   2280
         TabIndex        =   10
         Top             =   1770
         Width           =   1095
      End
      Begin VB.ComboBox UnitsInterest 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   0
         Left            =   3480
         Style           =   2  'Dropdown List
         TabIndex        =   59
         TabStop         =   0   'False
         Top             =   1170
         Width           =   1275
      End
      Begin VB.ComboBox UnitsInterest 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   2
         Left            =   3480
         Style           =   2  'Dropdown List
         TabIndex        =   58
         TabStop         =   0   'False
         Top             =   1770
         Width           =   1275
      End
      Begin VB.ComboBox UnitsInterest 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   3
         Left            =   3480
         Style           =   2  'Dropdown List
         TabIndex        =   57
         TabStop         =   0   'False
         Top             =   2070
         Width           =   1275
      End
      Begin VB.ComboBox UnitsInterest 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   4
         Left            =   3480
         Style           =   2  'Dropdown List
         TabIndex        =   56
         TabStop         =   0   'False
         Top             =   2370
         Width           =   1275
      End
      Begin VB.ComboBox UnitsInterest 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   5
         Left            =   3480
         Style           =   2  'Dropdown List
         TabIndex        =   55
         TabStop         =   0   'False
         Top             =   2670
         Width           =   1275
      End
      Begin VB.ComboBox UnitsInterest 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   7
         Left            =   3480
         Style           =   2  'Dropdown List
         TabIndex        =   54
         TabStop         =   0   'False
         Top             =   3270
         Width           =   1275
      End
      Begin VB.CommandButton cmdOndakla 
         Appearance      =   0  'Flat
         Caption         =   "Onda KLa"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   265
         Left            =   180
         TabIndex        =   53
         TabStop         =   0   'False
         Top             =   1155
         Width           =   1935
      End
      Begin VB.Shape Shape1 
         Height          =   795
         Left            =   180
         Top             =   270
         Width           =   4575
      End
      Begin VB.Label lblDesignConcentrationValue 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
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
         Height          =   255
         Index           =   5
         Left            =   2280
         TabIndex        =   79
         Top             =   2670
         Width           =   1095
      End
      Begin VB.Label lblDesignConcentrationValue 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
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
         Height          =   255
         Index           =   4
         Left            =   2280
         TabIndex        =   78
         Top             =   2370
         Width           =   1095
      End
      Begin VB.Label lblDesignConcentrationValue 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
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
         Height          =   255
         Index           =   3
         Left            =   2280
         TabIndex        =   77
         Top             =   2070
         Width           =   1095
      End
      Begin VB.Label lblDesignConcentration 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Percent Removal"
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
         Height          =   255
         Index           =   6
         Left            =   -480
         TabIndex        =   76
         Top             =   2970
         Width           =   2535
      End
      Begin VB.Label lblDesignConcentration 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Treatment Objective"
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
         Height          =   255
         Index           =   4
         Left            =   -480
         TabIndex        =   75
         Top             =   2370
         Width           =   2535
      End
      Begin VB.Label lblDesignConcentration 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Influent Concentration"
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
         Height          =   255
         Index           =   3
         Left            =   -480
         TabIndex        =   74
         Top             =   2070
         Width           =   2535
      End
      Begin VB.Label lblDesignConcentration 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Effluent Concentration"
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
         Height          =   255
         Index           =   5
         Left            =   -480
         TabIndex        =   73
         Top             =   2670
         Width           =   2535
      End
      Begin VB.Label lblDesignConcentration 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Air Pressure Drop"
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
         Height          =   255
         Index           =   7
         Left            =   -480
         TabIndex        =   72
         Top             =   3270
         Width           =   2535
      End
      Begin VB.Label lblDesignConcentration 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Onda KLa"
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
         Height          =   255
         Index           =   0
         Left            =   1020
         TabIndex        =   71
         Top             =   1170
         Width           =   1005
      End
      Begin VB.Label lblDesignConcentrationValue 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
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
         Height          =   255
         Index           =   6
         Left            =   2280
         TabIndex        =   70
         Top             =   2970
         Width           =   1095
      End
      Begin VB.Label lblDesignConcentration 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "KLa Safety Factor"
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
         Height          =   255
         Index           =   1
         Left            =   -480
         TabIndex        =   69
         Top             =   1470
         Width           =   2535
      End
      Begin VB.Label lblDesignConcentration 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Design KLa"
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
         Height          =   255
         Index           =   2
         Left            =   -480
         TabIndex        =   68
         Top             =   1770
         Width           =   2535
      End
      Begin VB.Label lblDesignConcentrationValue 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
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
         Height          =   255
         Index           =   0
         Left            =   2280
         TabIndex        =   67
         Top             =   1170
         Width           =   1095
      End
      Begin VB.Label lblDesignConcentrationValue 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
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
         Height          =   255
         Index           =   7
         Left            =   2280
         TabIndex        =   66
         Top             =   3270
         Width           =   1095
      End
      Begin VB.Label lblMisc 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "( - )"
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
         Height          =   255
         Index           =   0
         Left            =   3480
         TabIndex        =   65
         Top             =   1470
         Width           =   1275
      End
      Begin VB.Label lblMisc 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "( % )"
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
         Height          =   255
         Index           =   2
         Left            =   3480
         TabIndex        =   64
         Top             =   2970
         Width           =   1155
      End
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   330
      Top             =   6120
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Menu mnuFileMenu 
      Caption         =   "&File"
      Begin VB.Menu mnuFile 
         Caption         =   "Switch to &Design Mode"
         Index           =   0
         Shortcut        =   ^D
      End
      Begin VB.Menu mnuFile 
         Caption         =   "-"
         Index           =   1
      End
      Begin VB.Menu mnuFile 
         Caption         =   "&Open"
         Index           =   3
      End
      Begin VB.Menu mnuFile 
         Caption         =   "&Save"
         Index           =   4
      End
      Begin VB.Menu mnuFile 
         Caption         =   "Save &As"
         Index           =   5
      End
      Begin VB.Menu mnuFile 
         Caption         =   "-"
         Index           =   6
      End
      Begin VB.Menu mnuFile 
         Caption         =   "&Print"
         Index           =   7
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
         Caption         =   "Select Printer"
         Index           =   8
      End
      Begin VB.Menu mnuFile 
         Caption         =   "-"
         Index           =   9
      End
      Begin VB.Menu mnuFile 
         Caption         =   "Return to &Main Menu"
         Index           =   10
         Shortcut        =   ^M
      End
      Begin VB.Menu mnuFile 
         Caption         =   "-"
         Index           =   11
      End
      Begin VB.Menu mnuFile 
         Caption         =   "&1 Old File #1"
         Index           =   191
         Visible         =   0   'False
      End
      Begin VB.Menu mnuFile 
         Caption         =   "&2 Old File #2"
         Index           =   192
         Visible         =   0   'False
      End
      Begin VB.Menu mnuFile 
         Caption         =   "&3 Old File #3"
         Index           =   193
         Visible         =   0   'False
      End
      Begin VB.Menu mnuFile 
         Caption         =   "&4 Old File #4"
         Index           =   194
         Visible         =   0   'False
      End
      Begin VB.Menu mnuFile 
         Caption         =   "-"
         Index           =   199
         Visible         =   0   'False
      End
      Begin VB.Menu mnuFile 
         Caption         =   "E&xit"
         Index           =   200
      End
   End
   Begin VB.Menu mnuUnitsMenu 
      Caption         =   "&Units"
      Begin VB.Menu mnuUnits 
         Caption         =   "Standard International (SI)"
         Index           =   0
      End
      Begin VB.Menu mnuUnits 
         Caption         =   "English"
         Index           =   1
      End
   End
   Begin VB.Menu mnuPowerMenu 
      Caption         =   "&Power"
      Begin VB.Menu mnuPower 
         Caption         =   "&Perform Power Calculations"
         Index           =   0
      End
   End
   Begin VB.Menu mnuOptionsMenu 
      Caption         =   "&Results"
      Begin VB.Menu mnuoptions 
         Caption         =   "&View All Concentration Results"
         Index           =   0
      End
      Begin VB.Menu mnuoptions 
         Caption         =   "View &Mass Transfer Parameters"
         Index           =   5
      End
   End
   Begin VB.Menu mnuHelp 
      Caption         =   "&Help"
      Begin VB.Menu mnuHelpItem 
         Caption         =   "Online Help ..."
         Index           =   5
         Visible         =   0   'False
      End
      Begin VB.Menu mnuHelpItem 
         Caption         =   "Online Manual ..."
         Index           =   6
      End
      Begin VB.Menu mnuHelpItem 
         Caption         =   "Manual Printing Instructions ..."
         Index           =   7
      End
      Begin VB.Menu mnuHelpItem 
         Caption         =   "-"
         Index           =   8
      End
      Begin VB.Menu mnuHelpItem 
         Caption         =   "View Version History ..."
         Index           =   10
      End
      Begin VB.Menu mnuHelpItem 
         Caption         =   "View Disclaimer ..."
         Index           =   20
         Visible         =   0   'False
      End
      Begin VB.Menu mnuHelpItem 
         Caption         =   "Technical Assistance Provided By ..."
         Index           =   30
         Visible         =   0   'False
      End
      Begin VB.Menu mnuHelpItem 
         Caption         =   "-"
         Index           =   190
      End
      Begin VB.Menu mnuHelpItem 
         Caption         =   "&About ASAP ..."
         Index           =   200
      End
   End
End
Attribute VB_Name = "frmPTADScreen2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim frmPTADScreen2_Okay_To_Unload As Boolean



Const frmPTADScreen1_declarations_end = True


Sub LOCAL___Reset_DemoVersionDisablings()
  If (IsThisADemo() = True) Then
    mnuFile(3).Enabled = False
    mnuFile(4).Enabled = False
    mnuFile(5).Enabled = False
    mnuFile(191).Enabled = False
    mnuFile(192).Enabled = False
    mnuFile(193).Enabled = False
    mnuFile(194).Enabled = False
    cmdAddComponent.Enabled = False
    cmdDeleteComponent.Enabled = False
    fraFlowsAndLoadings.Caption = "* DEMONSTRATION VERSION *"
    fraFlowsAndLoadings.ForeColor = QBColor(12)
  End If
End Sub


Private Sub AddPrompt(menuID As Integer, prompt As String)
    menuPrompts(iMenuPrompts).menuID = menuID
    menuPrompts(iMenuPrompts).prompt = prompt
    iMenuPrompts = iMenuPrompts + 1
End Sub

Private Sub cboSelectCompo_Click()
    Dim ContaminantIndex As Integer, i As Integer

    ContaminantIndex = cboSelectCompo.ListIndex + 1
    i = ContaminantIndex
    If i = 0 Then Exit Sub
    Scr2.DesignContaminant = Scr2.Contaminant(i)
    Call GetContaminantConcentrationsScreen2


    'Update Variables on Screen
'    Call GetVQmultVQAndAirFlowRate
'    Call GetLoadings
'    Call GetTowerAreaAndDiameter
'    Call GetOndaMassTransferCoefficient
'    Call GetDesignKLaOrKLaSafetyFactor
'    Call GetTowerHeightAndVolume

End Sub

Private Sub cmdAddComponent_Click()
Dim x As rec_frmContaminantPropertyEdit
Dim i As Integer

  If (Scr2.NumChemical + 1 > MAXCHEMICAL) Then
     MsgBox "The maximum number of contaminants has been reached.  It is not possible to input more than " & Format$(MAXCHEMICAL, "0") & " contaminants for design.  To add an additional contaminant, another must first be removed.", MB_ICONSTOP, "Packed Tower Aeration"
     cmdAddComponent.Enabled = False
     Exit Sub
  End If

  x.ModelName = "Packed Tower Aeration"
  x.ModelType = MODELTYPE_PACKEDTOWER
  x.DoEditNumber = -1       'Will be set by frmContaminantPropertyEdit.
  x.DoAdd = True
  x.OldNumCompo = cboSelectCompo.ListCount
  For i = 1 To x.OldNumCompo
    x.Contaminants(i) = Scr2.Contaminant(i)
  Next i

  StEPPImportSuccess = False
  Data_frmContaminantPropertyEdit = x
  frmContaminantPropertyEdit.Show 1
  x = Data_frmContaminantPropertyEdit

  If (StEPPImportSuccess) Or (Not x.CancelledAdd) Then
    For i = x.OldNumCompo + 1 To x.NewNumCompo
      'Incorporate new contaminant.
      If i > 10 Then
       MsgBox "Unable to continue importing file as maximum amount of chemicals in memory reached."
       Me.Show
       Exit Sub
      End If
        
        Scr2.Contaminant(i) = x.Contaminants(i)
        Scr2.NumChemical = Scr2.NumChemical + 1
        'Incorporate new name into ComboBox.
        cboSelectCompo.AddItem Scr2.Contaminant(i).Name
    Next i
  End If
  
  Call SetDesignContaminantEnabled(CInt(cboSelectCompo.ListCount))
  
  If (Scr2.NumChemical > 0) Then
    cmdDeleteComponent.Enabled = True
    cmdEditComponent.Enabled = True
    If (cboSelectCompo.ListIndex = -1) Then
      cboSelectCompo.ListIndex = 0
    End If
  End If
 Me.Show
End Sub

Private Sub cmdDeleteComponent_Click()
  Dim i As Integer

  Scr2.Chemical = cboSelectCompo.ListIndex + 1
  If (Scr2.Chemical = 0) Then Exit Sub

  If MsgBox("Remove" & NL & cboSelectCompo.List(cboSelectCompo.ListIndex), 36, "") = IDYES Then
    cboSelectCompo.RemoveItem cboSelectCompo.ListIndex
    For i = Scr2.Chemical To Scr2.NumChemical - 1
      Scr2.Contaminant(i) = Scr2.Contaminant(i + 1)
    Next i
    Scr2.NumChemical = Scr2.NumChemical - 1
    If (Scr2.NumChemical > 0) Then
      cboSelectCompo.ListIndex = 0
    Else
      cmdDeleteComponent.Enabled = False
      cmdEditComponent.Enabled = False
    End If
    Call SetDesignContaminantEnabled(CInt(cboSelectCompo.ListCount))
  End If

  If Scr2.NumChemical < 10 Then cmdAddComponent.Enabled = True
  Call LOCAL___Reset_DemoVersionDisablings
End Sub

Private Sub cmdEditComponent_Click()
Dim x As rec_frmContaminantPropertyEdit
Dim i As Integer
Dim AListIndex As Integer

  Scr2.Chemical = cboSelectCompo.ListIndex + 1
  If (Scr2.Chemical = 0) Then Exit Sub

  x.ModelName = "Packed Tower Aeration"
  x.ModelType = MODELTYPE_PACKEDTOWER
  x.DoEditNumber = cboSelectCompo.ListIndex + 1
  x.DoAdd = False
  x.OldNumCompo = cboSelectCompo.ListCount
  For i = 1 To x.OldNumCompo
    x.Contaminants(i) = Scr2.Contaminant(i)
  Next i

  Data_frmContaminantPropertyEdit = x
  frmContaminantPropertyEdit.Show 1
  x = Data_frmContaminantPropertyEdit

  If (Not x.CancelledEdit) Then
    For i = 1 To x.NewNumCompo
      Scr2.Contaminant(i) = x.Contaminants(i)
    Next i
    If (x.OldNumCompo < x.NewNumCompo) Then
      'Incorporate new names into ComboBox.
      For i = x.OldNumCompo + 1 To x.NewNumCompo
        cboSelectCompo.AddItem Scr2.Contaminant(i).Name
      Next i
    End If
    'Update ComboBox for any changed names:
    For i = 1 To x.OldNumCompo
      If (Trim$(cboSelectCompo.List(i - 1)) <> Trim$(Scr2.Contaminant(i).Name)) Then
        cboSelectCompo.List(i - 1) = Trim$(Scr2.Contaminant(i).Name)
      End If
    Next i
  Else
    Exit Sub
  End If

  'Generate click event on cboSelectCompo
  AListIndex = cboSelectCompo.ListIndex
  cboSelectCompo.ListIndex = -1
  cboSelectCompo.ListIndex = AListIndex

End Sub

Private Sub cmdOndakla_Click()
       If frmPTADScreen2.lblDesignConcentrationValue(0).Caption = "0.0" Then Exit Sub
       frmShowOndaKLaProperties.Show 1


End Sub

Private Sub cmdPackingType_Click()

  ScreenNumber = 2
  CurrentScreen = Scr2
  Call ShowPackingProperties

End Sub

Private Sub cmdPickFlowLoadingParameters_Click()

    frmFlowsLoadingsScreen2.Show 1
End Sub

Private Sub cmdSelectContaminants_Click()
    
  'frmListcontaminantScreen2.Show 1

End Sub

Private Sub cmdSelectPacking_Click()
    Dim i As Integer, CurrPackingIndex As Integer

    ReadMainPackingDB
    ReadUserPackingDB
  
    PackingDatabaseSource = Scr2.Packing.SourceDatabase

    If Scr2.Packing.Name = "" Then
       Scr2.Packing.Name = frmSelectPacking.cboSelectPacking.List(0)
       PackingChanged = True
    End If

    If PackingDatabaseSource = ORIGINALPACKINGDATABASE Then
       For i = 1 To NumPackingsInDatabase
           If DatabasePacking(i).Name = Scr2.Packing.Name Then
              CurrPackingIndex = i
              frmSelectPacking.cboSelectPacking.ListIndex = CurrPackingIndex - 1
              Exit For
           End If
       Next i
    ElseIf PackingDatabaseSource = USERMODIFIEDPACKINGDATABASE Then
       For i = 1 To NumUserPackings
           If UserPacking(i).Name = Scr2.Packing.Name Then
              CurrPackingIndex = i
              frmSelectPacking.cboSelectPacking.ListIndex = CurrPackingIndex - 1
              Exit For
           End If
       Next i

    End If

    ScreenNumber = 2
    CurrentScreen = Scr2
    frmSelectPacking.Show 1
    Scr2 = CurrentScreen

    If Scr2.NumChemical > 0 Then
       'Update values on screen
       Call GetContaminantConcentrationsScreen2
    End If

End Sub

Private Sub Command1_Click()
Call screen2_results
End Sub

Private Sub Form_Activate()
'    Dim hMenu       As Integer
'    Dim hSubMenu    As Integer
'
''Initialize MsgHook and Load Menu Prompts to Display on Status Bar
'    imenuprompts = 0
'
'    MsgHook1.HwndHook = Me.hWnd
'    MsgHook1.Message(WM_MENUSELECT) = True
'    hMenu = GetMenu(Me.hWnd)
'    '
'    ' Load File menu prompts
'    '
'    hSubMenu = GetSubMenu(hMenu, 0)
'    AddPrompt hSubMenu, "File operations"
'    AddPrompt GetMenuItemID(hSubMenu, 0), "Switch to Design Mode for packed tower aeration"
'    AddPrompt GetMenuItemID(hSubMenu, 2), "Load a rating case from a file"
'    AddPrompt GetMenuItemID(hSubMenu, 3), "Save this rating case to a file"
'    AddPrompt GetMenuItemID(hSubMenu, 4), "Save this rating case to a file"
'    AddPrompt GetMenuItemID(hSubMenu, 6), "Print this rating case"
'    AddPrompt GetMenuItemID(hSubMenu, 7), "Select printer for printing results"
'    AddPrompt GetMenuItemID(hSubMenu, 9), "Leave packed tower aeration and return to main ASAP menu"
'    AddPrompt GetMenuItemID(hSubMenu, 11), "Exit program"
'
'    '
'    ' Load Units menu prompts
'    '
'    hSubMenu = GetSubMenu(hMenu, 1)
'    AddPrompt hSubMenu, "Units operations"
'    AddPrompt GetMenuItemID(hSubMenu, 0), "Display results in Standard International (SI) units"
'    AddPrompt GetMenuItemID(hSubMenu, 1), "Display results in English units"
'    '
'    ' Load Power menu prompt
'    '
'    hSubMenu = GetSubMenu(hMenu, 2)
'    AddPrompt hSubMenu, "Power operations"
'    AddPrompt GetMenuItemID(hSubMenu, 0), "Perform power calculations"
'
'    ' Load Options menu prompt
'    '
'    hSubMenu = GetSubMenu(hMenu, 3)
'    AddPrompt hSubMenu, "Options"
'    AddPrompt GetMenuItemID(hSubMenu, 0), "View Effluent Concentration Results for All Contaminants"

    'Initialize last-few-files list.
    Call LastFewFiles_InitializeList(LASTFEW_WHICHAPP_ASAP, LASTFEW_ASAP_frmPTADScreen2)

End Sub

Private Sub Form_Load()
    
  frmPTADScreen2_Okay_To_Unload = False
  
  frmPTADScreen2.WindowState = 0
  frmPTADScreen2.Width = SCREEN_WIDTH_STANDARD
  frmPTADScreen2.Height = SCREEN_HEIGHT_STANDARD

  'Center the form on the screen
  If (WindowState = 0) Then
    'don't attempt if screen Minimized or Maximized
    Move (Screen.Width - frmPTADScreen2.Width) / 2, (Screen.Height - frmPTADScreen2.Height) / 2
  End If

  Call LabelsPTADScreen2(UNITSTYPE_SI)

  'StatusMessagePanel.BackColor = &HC0C0C0
  'StatusMessagePanel.ForeColor = &H400040
  'StatusBarPanel.BackColor = &HC0C0C0
  'StatusBarPanel.ForeColor = &H400040

  '
  ' DEMO SETTINGS.
  '
  Call LOCAL___Reset_DemoVersionDisablings
End Sub
Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
  If (frmPTADScreen2_Okay_To_Unload) Then
    Cancel = False
  Else
    Cancel = True
  End If
End Sub


Private Sub lblDesignConcentration_Click(Index As Integer)
    If Index = 0 Then
       If frmPTADScreen2.lblDesignConcentrationValue(0).Caption = "0.0" Then Exit Sub
       frmShowOndaKLaProperties.Show 1
    End If

End Sub

Private Sub lblDesignConcentrationValue_Change(Index As Integer)

  Call UnitsInterest_Click(Index)

End Sub

Private Sub lblDisplayAirWaterProperties_Click()

    If HaveValue(Scr2.OperatingPressure.value) And HaveValue(Scr2.operatingtemperature.value) Then
       CurrentScreen = Scr2
       CurrentMode = 2
       frmAirWaterProperties.Show 1
    Else
       MsgBox "You must specify pressure and temperature before physical properties can be displayed.", MB_ICONSTOP, "Error"
    End If

End Sub

Private Sub lblPackingType_Click()
    
  Call cmdPackingType_Click

End Sub

Private Sub mnuFile_Click(Index As Integer)
  Dim i As Integer, Response As Integer
  Dim msg As String
    Screen.MousePointer = 11   'Hourglass

    Select Case Index
       Case 0   'Switch to Design Mode

          If frmPTADScreen2!lblDesignConcentrationValue(6).Caption = "0.0" Then
             Filename$ = "TheDefaultCaseScreen1"
             If (loadscreen1("") = False) Then
               Exit Sub
             End If
             frmPTADScreen2.Hide
             frmPTADScreen1.Show
             Screen.MousePointer = 0
             Exit Sub
          End If
          

          If HaveValue(Scr2.AirPressureDrop.value) Then
             'Give user option to save screen 2 before
             'switching to screen 1
             msg = "Would you like to save the parameters "
             msg = msg + "for this optimization case to a file "
             msg = msg + "before switching to Design "
             msg = msg + "Mode?  " & Chr$(13) & Chr$(13)
             msg = msg + "(Note:  Non-saved information will be lost permanently)."
             Response = MsgBox(msg, MB_ICONquestion + MB_YESNO, "Save Current Design")
             If Response = IDYES Then
                Call SaveScreen2
             End If
          End If

          If ShownScreen1Previously Then   'Show Design Parameters calculated earlier from that case
             frmPTADScreen2.Hide
             scr1.OperatingPressure.ValChanged = True
             Call CalculateAirWaterProperties
             frmPTADScreen1.Show
          Else   'Transfer Parameters from Rating Mode to Design Mode
             'Initialize Values on Screen 1:

             'Pressure and Temperature
             frmPTADScreen1!txtOperatingPressure.Text = frmPTADScreen2!txtOperatingPressure.Text
             scr1.OperatingPressure = Scr2.OperatingPressure
             scr1.OperatingPressure.ValChanged = True
             frmPTADScreen1!txtOperatingTemperature.Text = frmPTADScreen2!txtOperatingTemperature.Text
             scr1.operatingtemperature = Scr2.operatingtemperature
             scr1.operatingtemperature.ValChanged = True

             'Physical Properties
             Call CalculateAirWaterProperties

             'Packing
             frmPTADScreen1!lblPackingType.Caption = frmPTADScreen2!lblPackingType.Caption
             scr1.Packing = Scr2.Packing

             'Contaminant List
             scr1.NumChemical = Scr2.NumChemical
             frmPTADScreen1!cboSelectCompo.Clear
             'frmListContaminant!ListContaminants.Clear
             If scr1.NumChemical > 0 Then
                For i = 1 To scr1.NumChemical
                    scr1.Contaminant(i) = Scr2.Contaminant(i)
                    frmPTADScreen1!cboSelectCompo.AddItem scr1.Contaminant(i).Name
                    'frmListContaminant!ListContaminants.AddItem Scr1.Contaminant(i).Name
                Next i


                'frmListContaminant!mnuOptionsManipulateContaminant(1).Enabled = True
                'frmListContaminant!mnuOptionsManipulateContaminant(3).Enabled = True
                'frmListContaminant!mnuOptionsManipulateContaminant(4).Enabled = True
                'frmListContaminant!mnuOptionsSave.Enabled = True
                'frmListContaminant!mnuOptionsView.Enabled = True

                Call SetDesignContaminantEnabled(CInt(scr1.NumChemical))
             End If

             'Flow and Loading Parameters
             frmPTADScreen1!txtFlowsLoadings(0).Text = frmPTADScreen2!txtFlowsLoadings(0).Text
             scr1.WaterFlowRate = Scr2.WaterFlowRate
             scr1.WaterFlowRate.UserInput = True
             If scr1.AirFlowRate.UserInput = True Then
                frmPTADScreen1!txtFlowsLoadings(4).Text = frmPTADScreen1!txtFlowsLoadings(1).Text
                scr1.AirFlowRate = Scr2.AirFlowRate
                scr1.AirFlowRate.UserInput = True
                scr1.AirToWaterRatio.UserInput = False
                scr1.MultipleOfMinimumAirToWaterRatio.UserInput = False
             Else
                'Consider Air To Water Ratio User Input
                frmPTADScreen1!txtFlowsLoadings(3).Text = frmPTADScreen2!txtFlowsLoadings(2).Text
                scr1.AirToWaterRatio = Scr2.AirToWaterRatio
                scr1.AirToWaterRatio.UserInput = True
                scr1.AirFlowRate.UserInput = False
                scr1.MultipleOfMinimumAirToWaterRatio.UserInput = False
             End If
       
             'KLaSafetyFactor
             scr1.KLaSafetyFactor.value = Scr2.KLaSafetyFactor.value
             scr1.KLaSafetyFactor.UserInput = True
             frmPTADScreen1!txtMassTransfer(1).Text = Format$(scr1.KLaSafetyFactor.value, GetTheFormat(scr1.KLaSafetyFactor.value))

             'Calculate properties of the contaminant in the
             'combo box now.

             scr1.DesignContaminant = Scr2.DesignContaminant
             frmPTADScreen1!cboSelectCompo.ListIndex = frmPTADScreen2!cboSelectCompo.ListIndex

             frmPTADScreen1.Caption = "Packed Tower Aeration - Design Mode (untitled.des)"

             frmPTADScreen2.Hide
             frmPTADScreen1.Show

          End If

       Case 2   'New

       Case 3   'Open
          If HaveValue(Scr2.TowerVolume.value) Then
            If (screen2_savechanges()) Then Exit Sub
          End If
          
          ''''ChDrive SaveAndLoadPath
          ''''ChDir SaveAndLoadPath
          Call ChangeDir_Main
          Call loadscreen2("")

          SaveAndLoadPath = CurDir$
          ''''ChDrive App.Path
          ''''ChDir App.Path
          Call ChangeDir_Main

       Case 4   'Save
          ''''ChDrive SaveAndLoadPath
          ''''ChDir SaveAndLoadPath
          Call ChangeDir_Main
          Call SaveScreen2
          
          'Add this file to the last-few-files list if necessary.
          Call LastFewFiles_MoveFilenameToTop(Filename)
          
          SaveAndLoadPath = CurDir$
          ''''ChDrive App.Path
          ''''ChDir App.Path
          Call ChangeDir_Main

       Case 5   'Save As
          ''''ChDrive SaveAndLoadPath
          ''''ChDir SaveAndLoadPath
          Call ChangeDir_Main
          If Right$(frmPTADScreen2.Caption, 14) <> "(untitled.rat)" Then Call savefilescreen2(Filename)
          Call SaveScreen2
          
          'Add this file to the last-few-files list if necessary.
          Call LastFewFiles_MoveFilenameToTop(Filename)
          
          SaveAndLoadPath = CurDir$
          ''''ChDrive App.Path
          ''''ChDir App.Path
          Call ChangeDir_Main

       Case 7   'Print
          

       Case 8   'Select Printer
            On Error GoTo PrinterError
            'CMDialog1.flags = PD_PRINTSETUP
            'CMDialog1.Action = 5
            CommonDialog1.ShowPrinter
            GoTo ExitSelectPrint
PrinterError:
            Resume ExitSelectPrint:

ExitSelectPrint:

        Case 10   'Return to Main Menu
            If HaveValue(Scr2.TowerVolume.value) Then
               If (screen2_savechanges()) Then Exit Sub
            End If

            'unload Forms for Packed Tower Aeration
            Unload frmAirWaterProperties
            'Unload frmListContaminant
            'Unload frmPropContaminant
            ''''Unload HelpTipForm
            Unload frmShowOndaKLaProperties
            Unload frmSelectPacking
            Unload frmShowPackingProperties
            Unload frmPower
            Unload frmPowerScreen2
            Unload frmFlowsLoadingsScreen2
            'Unload frmListcontaminantScreen2
            'Unload frmPropContaminantScreen2
            Unload frmOptimizeContaminant
            Unload frmPTADScreen1
            frmPTADScreen2_Okay_To_Unload = True
            Unload frmPTADScreen2
            
            frmMainMenu.Show


        Case 200   'Exit
            'Give user option to save design mode before Exiting
            If HaveValue(Scr2.TowerVolume.value) Then
               If (screen2_savechanges()) Then Exit Sub
            End If
            
            'unload Forms for Packed Tower Aeration
            Unload frmAirWaterProperties
            'Unload frmListContaminant
            'Unload frmPropContaminant
            ''''Unload HelpTipForm
            Unload frmShowOndaKLaProperties
            Unload frmSelectPacking
            Unload frmShowPackingProperties
            Unload frmPower
            Unload frmPowerScreen2
            Unload frmFlowsLoadingsScreen2
            'Unload frmListcontaminantScreen2
            'Unload frmPropContaminantScreen2
            Unload frmOptimizeContaminant
            Unload frmPTADScreen1
            frmPTADScreen2_Okay_To_Unload = True
            Unload frmPTADScreen2
            Unload frmMainMenu
            End
    
    End Select
    
    If ((Index >= 191) And (Index <= 194)) Then
      'Handle File|Open of a file here.
      ''''ChDrive SaveAndLoadPath
      ''''ChDir SaveAndLoadPath
      Call ChangeDir_Main
      If (Dir(Current_LastFewFilesRec.FileNames(Index - 190)) = "") Then
        Beep
        MsgBox "That file has been moved or deleted.", MB_ICONEXCLAMATION, Application_Name
      Else
        Call loadscreen2(Current_LastFewFilesRec.FileNames(Index - 190))
        'Add this file to the last-few-files list if necessary.
        Call LastFewFiles_MoveFilenameToTop(Filename)
        SaveAndLoadPath = CurDir$
      End If
      ''''ChDir App.Path
      ''''ChDrive App.Path
      Call ChangeDir_Main
    End If

    Screen.MousePointer = 0   'Arrow

End Sub

Private Sub mnuFilePrint_Click(Index As Integer)

    Select Case Index
       Case 0   'Print to printer
          Call PrintPTADScreen2
       Case 1   'Print to file
          Call PrintPTADScreen2ToFile
    End Select

End Sub

Private Sub mnuHelpItem_Click(Index As Integer)
  Call Launch_ASAP_mnuHelp_Item(Index)
'  Select Case Index
'    'Case 10:
'    '  frmabout2.Show 1
'    'Case 99:
'    '  frmAbout.Show 1
'    Case 300:
'      Call Launch_ASAP_HLP_File
'  End Select
End Sub

Private Sub mnuOndaKla_Click(Index As Integer)
End Sub

Private Sub mnuOptions_Click(Index As Integer)

  Select Case Index
    Case 0: Call screen2_results
    Case 5:
      If frmPTADScreen2.lblDesignConcentrationValue(0).Caption = "0.0" Then Exit Sub
      frmShowOndaKLaProperties.Show 1
  
  End Select

End Sub

Private Sub mnuotheritem_Click()
'frmabout2.Show 1
End Sub

Private Sub mnuPopContaminant_Click(Index As Integer)
    
    'If Index = 1 Then
    '   If frmListcontaminantScreen2!ListContaminants.ListCount > 0 Then
    '      frmListcontaminantScreen2!ListContaminants.ListIndex = 0
    '   End If
    '
    '   frmListcontaminantScreen2.Show 1
    'End If

End Sub

Private Sub mnuPower_Click(Index As Integer)
    Dim CalculatedPower As Integer

    Select Case Index   'Power Calculation

       Case 0
             Call SetPowerPTADScreen2(CalculatedPower)
             If CalculatedPower Then
                frmPowerScreen2.Left = Screen.Width / 2 - frmPowerScreen2.Width / 2
                frmPowerScreen2.Top = Screen.Height / 2 - frmPowerScreen2.Height / 2
                frmPowerScreen2.Show 1
             End If
    End Select

End Sub

Private Sub mnuUnits_Click(Index As Integer)

  Select Case Index
    Case 0        'SI
      Call LabelsPTADScreen2(UNITSTYPE_SI)
    Case 1        'English
      Call LabelsPTADScreen2(UNITSTYPE_ENGLISH)
  End Select

End Sub

Private Sub MsgHook1_Message(msg As Integer, wParam As Integer, lParam As Long, Action As Integer, result As Long)
'    Dim i       As Integer
'    Dim found   As Integer
'
'    '
'    ' Got a menu select message ... see if it's for one of our menus
'    '
'    For i = 0 To iMenuPrompts - 1
'        If (menuPrompts(i).menuID = wParam) Then
'            '
'            ' One of our menus ... display prompt message
'            '
'            StatusMessagePanel.Caption = menuPrompts(i).prompt
'            found = True
'            Exit For
'        End If
'    Next
'    '
'    ' Blank prompt message when no menu selected
'    '
'    If (found <> True) Then
'        StatusMessagePanel.Caption = ""
'    End If
'
End Sub

Private Sub Old_HelpTipTimer_Timer()
''temp kill
'Exit Sub
'''''''''''''''''''''
'
' Dim PointStruct As PointType
' Static PrevioushWnd%
' Dim CurrenthWnd As Integer, TipText As String
'
'If GetActiveWindow() = Me.hWnd Then
'  Call GetCursorPos(PointStruct)
'  CurrenthWnd% = WindowFromPoint(PointStruct.Y, PointStruct.x)
'  If CurrenthWnd <> PrevioushWnd% Then
'    PrevioushWnd% = CurrenthWnd
'    'HelpTipTimer.Interval = 1
'    Select Case CurrenthWnd%
'      Case txtOperatingPressure.hWnd    '<---- Here for the text box txtOperatingPressure
'       StatusMessagePanel.Caption = " Input Operating " & lblOperatingPressure.Caption
'      Case txtOperatingTemperature.hWnd   '<---- Here for the text box txtOperatingTemperature
'       StatusMessagePanel.Caption = " Input Operating " & lblOperatingTemperature.Caption
'      Case lblDisplayAirWaterProperties.hWnd
'       StatusMessagePanel.Caption = " Specify water density, viscosity, and surface tension; and air density and viscosity"
'      Case txtFlowsLoadings(0).hWnd
'       StatusMessagePanel.Caption = "Input " & lblFlowsLoadingsLabel(0).Caption
'      Case txtFlowsLoadings(1).hWnd
'       StatusMessagePanel.Caption = "Input " & lblFlowsLoadingsLabel(1).Caption
'      Case txtFlowsLoadings(2).hWnd
'       StatusMessagePanel.Caption = "Input " & lblFlowsLoadingsLabel(2).Caption
'      Case txtFlowsLoadings(3).hWnd
'       StatusMessagePanel.Caption = "Input " & lblFlowsLoadingsLabel(3).Caption
'      Case txtFlowsLoadings(4).hWnd
'       StatusMessagePanel.Caption = "Input " & lblFlowsLoadingsLabel(4).Caption
'    End Select
'    ShowHelpTip TipText$
'    If Len(TipText$) = 0 Then
'      'HelpTipTimer.Interval = 500 'Milliseconds
'    End If
'  End If
'End If
'
End Sub

Private Sub txtDesignConcentrationValue_GotFocus(Index As Integer)
    
  Call GotFocus_Handle(Me, txtDesignConcentrationValue(Index), Temp_Text)
End Sub

Private Sub txtDesignConcentrationValue_KeyPress(Index As Integer, KeyAscii As Integer)
  Call TextBoxNumber_KeyPress(KeyAscii)
End Sub

Private Sub txtDesignConcentrationValue_LostFocus(Index As Integer)
Dim NewVal As Double
Dim IsNew As Integer
    Dim flag_ok As Integer

   If (LostFocus_IsEvil(Me, txtDesignConcentrationValue(Index))) Then
     Exit Sub
   End If
   
   flag_ok = True

  IsNew = False
  
  Select Case Index
    Case 1        'KLa Safety Factor
      If (NoUnits_LostFocus(txtDesignConcentrationValue(1), NewVal, Temp_Text)) Then
        IsNew = True
        Scr2.KLaSafetyFactor.value = NewVal
        Scr2.KLaSafetyFactor.ValChanged = True
        Scr2.KLaSafetyFactor.UserInput = True
        Call SpecifiedKLaSafetyFactorScreen2
      End If
    
    Case 2        'Design KLa
      If (Unitted_LostFocus(UNITS_INVERSETIME, txtDesignConcentrationValue(2), UnitsInterest(2), NewVal, Temp_Text)) Then
        IsNew = True
        Scr2.DesignMassTransferCoefficient.value = NewVal
        Scr2.DesignMassTransferCoefficient.ValChanged = True
        Scr2.DesignMassTransferCoefficient.UserInput = True
        Call SpecifiedDesignKLaScreen2
      End If

  End Select

  If (IsNew) Then
    Call GetContaminantConcentrationsScreen2
  End If

  Call LostFocus_Handle(Me, txtDesignConcentrationValue(Index), flag_ok)


End Sub

Private Sub txtFlowsLoadings_GotFocus(Index As Integer)
    
  Call GotFocus_Handle(Me, txtFlowsLoadings(Index), Temp_Text)
End Sub

Private Sub txtFlowsLoadings_KeyPress(Index As Integer, KeyAscii As Integer)
  Call TextBoxNumber_KeyPress(KeyAscii)
End Sub

Private Sub txtFlowsLoadings_LostFocus(Index As Integer)
Dim i As Integer
Dim NewVal As Double
Dim IsNew As Integer
    Dim flag_ok As Integer

   If (LostFocus_IsEvil(Me, txtFlowsLoadings(Index))) Then
     Exit Sub
   End If
   
   flag_ok = True

  'Exit Subroutine if not in a text box user can
  'currently specify
  If UsersFlowsLoadingsOption = 0 Then
    Select Case Index
      Case 2 To 4
        Exit Sub
    End Select
  ElseIf UsersFlowsLoadingsOption = 1 Then
    Select Case Index
      Case 1, 3, 4
        Exit Sub
    End Select
  ElseIf UsersFlowsLoadingsOption = 2 Then
    Select Case Index
      Case 0 To 2
        Exit Sub
    End Select
  End If
  
  IsNew = False
  
  Select Case Index

    Case 0        'Water Flow Rate.
      If (Unitted_LostFocus(UNITS_FLOW, txtFlowsLoadings(0), UnitsFlows(0), NewVal, Temp_Text)) Then
        IsNew = True
        Scr2.WaterFlowRate.value = NewVal
        Scr2.WaterFlowRate.ValChanged = True
      End If

    Case 1        'Air Flow Rate
      If (Unitted_LostFocus(UNITS_FLOW, txtFlowsLoadings(1), UnitsFlows(1), NewVal, Temp_Text)) Then
        IsNew = True
        Scr2.AirFlowRate.value = NewVal
        Scr2.AirFlowRate.ValChanged = True
      End If
      
    Case 2        'Air To Water Ratio
      If (NoUnits_LostFocus(txtFlowsLoadings(2), NewVal, Temp_Text)) Then
        IsNew = True
        Scr2.AirToWaterRatio.value = NewVal
        Scr2.AirToWaterRatio.ValChanged = True
      End If
        
    Case 3        'Water Loading Rate
      If (Unitted_LostFocus(UNITS_MASSLOADINGRATE, txtFlowsLoadings(3), UnitsFlows(3), NewVal, Temp_Text)) Then
        IsNew = True
        Scr2.WaterLoadingRate.value = NewVal
        Scr2.WaterLoadingRate.ValChanged = True
      End If

    Case 4        'Air Loading Rate
      If (Unitted_LostFocus(UNITS_MASSLOADINGRATE, txtFlowsLoadings(4), UnitsFlows(4), NewVal, Temp_Text)) Then
        IsNew = True
        scr1.AirLoadingRate.value = NewVal
        scr1.AirLoadingRate.ValChanged = True
      End If
  
  End Select

  If (IsNew) Then
    Call GetFlowsAndLoadingsScreen2

    If Scr2.NumChemical > 0 Then
      'Update Variables on Screen
      Call GetContaminantConcentrationsScreen2
'      Call GetVQmultVQAndAirFlowRate
'      Call GetLoadings
'      Call GetTowerAreaAndDiameter
'      Call GetOndaMassTransferCoefficient
'      Call GetDesignKLaOrKLaSafetyFactor
'      Call GetTowerHeightAndVolume
    End If
  End If

  Call LostFocus_Handle(Me, txtFlowsLoadings(Index), flag_ok)


End Sub

Private Sub txtOperatingPressure_GotFocus()
    
  Call GotFocus_Handle(Me, txtOperatingPressure, Temp_Text)
End Sub

Private Sub txtOperatingPressure_KeyPress(KeyAscii As Integer)
  Call TextBoxNumber_KeyPress(KeyAscii)
End Sub

Private Sub txtOperatingPressure_LostFocus()
Dim NewVal As Double
    Dim flag_ok As Integer

   If (LostFocus_IsEvil(Me, txtOperatingPressure)) Then
     Exit Sub
   End If
   
   flag_ok = True

  If (Unitted_LostFocus(UNITS_PRESSURE, txtOperatingPressure, UnitsOpCond(0), NewVal, Temp_Text)) Then
    Scr2.OperatingPressure.ValChanged = True
    Scr2.OperatingPressure.UserInput = True
    'Note: standard P units are Pa, but
    'OperatingPressure is stored as kPa.
    Scr2.OperatingPressure.value = NewVal * 1# / 101325#

    If (HaveValue(Scr2.OperatingPressure.value) And HaveValue(Scr2.operatingtemperature.value)) Then
      Call CalculateAirWaterPropertiesScreen2
      Call GetFlowsAndLoadingsScreen2

      If (Scr2.NumChemical > 0) Then
        'Update Variables on Screen
        Call GetContaminantConcentrationsScreen2
'        Call GetVQmultVQAndAirFlowRate
'        Call GetLoadings
'        Call GetTowerAreaAndDiameter
'        Call GetOndaMassTransferCoefficient
'        Call GetDesignKLaOrKLaSafetyFactor
'        Call GetTowerHeightAndVolume
      End If
    End If
  End If
  Call LostFocus_Handle(Me, txtOperatingPressure, flag_ok)

  
End Sub

Private Sub txtOperatingTemperature_GotFocus()
    
  Call GotFocus_Handle(Me, txtOperatingTemperature, Temp_Text)
End Sub

Private Sub txtOperatingTemperature_KeyPress(KeyAscii As Integer)
  Call TextBoxNumber_KeyPress(KeyAscii)
End Sub

Private Sub txtOperatingTemperature_LostFocus()
Dim NewVal As Double
    Dim flag_ok As Integer

   If (LostFocus_IsEvil(Me, txtOperatingTemperature)) Then
     Exit Sub
   End If
   
   flag_ok = True

  If (Unitted_LostFocus(UNITS_TEMPERATURE, txtOperatingTemperature, UnitsOpCond(1), NewVal, Temp_Text)) Then
    Scr2.operatingtemperature.ValChanged = True
    Scr2.operatingtemperature.UserInput = True
    Scr2.operatingtemperature.value = NewVal

    If (HaveValue(Scr2.OperatingPressure.value) And HaveValue(Scr2.operatingtemperature.value)) Then
      Call CalculateAirWaterPropertiesScreen2
      Call GetFlowsAndLoadingsScreen2
      
      If (Scr2.NumChemical > 0) Then
        'Update Variables on Screen
        Call GetContaminantConcentrationsScreen2
             
'        Call GetVQmultVQAndAirFlowRate
'        Call GetLoadings
'        Call GetTowerAreaAndDiameter
'        Call GetOndaMassTransferCoefficient
'        Call GetDesignKLaOrKLaSafetyFactor
'        Call GetTowerHeightAndVolume
      End If
    End If
  End If
  Call LostFocus_Handle(Me, txtOperatingTemperature, flag_ok)

    
End Sub

Private Sub txtTowerParameters_GotFocus(Index As Integer)
    
  Call GotFocus_Handle(Me, txtTowerParameters(Index), Temp_Text)
End Sub

Private Sub txtTowerParameters_KeyPress(Index As Integer, KeyAscii As Integer)
  Call TextBoxNumber_KeyPress(KeyAscii)
End Sub

Private Sub txtTowerParameters_LostFocus(Index As Integer)
Dim NewVal As Double
Dim IsNew As Integer
    Dim flag_ok As Integer

   If (LostFocus_IsEvil(Me, txtTowerParameters(Index))) Then
     Exit Sub
   End If
   
   flag_ok = True

  IsNew = False
  If txtTowerParameters(Index) > 2000 Then
    MsgBox "Invalid input:  the input range is (0 to 2000 m)"
    Exit Sub
  End If

  Select Case Index
    Case 0        'Specify Diameter.
      If (Unitted_LostFocus(UNITS_LENGTH, txtTowerParameters(0), UnitsTowerParam(0), NewVal, Temp_Text)) Then
        IsNew = True
        Scr2.SpecifiedTowerDiameter.value = NewVal
        Scr2.SpecifiedTowerDiameter.ValChanged = True
        Scr2.SpecifiedTowerDiameter.UserInput = True
      End If

    Case 1        'Specify Height.
      If (Unitted_LostFocus(UNITS_LENGTH, txtTowerParameters(1), UnitsTowerParam(1), NewVal, Temp_Text)) Then
        IsNew = True
        Scr2.SpecifiedTowerHeight.value = NewVal
        Scr2.SpecifiedTowerHeight.ValChanged = True
        Scr2.SpecifiedTowerHeight.UserInput = True
      End If
    
  End Select

  If (IsNew) Then
    Call GetTowerAreaAndVolume
    Call GetFlowsAndLoadingsScreen2
       
    If (Scr2.NumChemical > 0) Then
      'Update Variables on Screen
      Call GetContaminantConcentrationsScreen2
'
'      Call GetVQmultVQAndAirFlowRate
'      Call GetLoadings
'      Call GetTowerAreaAndDiameter
'      Call GetOndaMassTransferCoefficient
'      Call GetDesignKLaOrKLaSafetyFactor
'      Call GetTowerHeightAndVolume
    End If
    
  End If
  Call LostFocus_Handle(Me, txtTowerParameters(Index), flag_ok)


End Sub

Private Sub UnitsDesignBasis_Click(Index As Integer)

  Select Case Index
    Case 0        'Tower Diameter
      Call Unitted_UnitChange(UNITS_LENGTH, Scr2.TowerDiameter.value, UnitsDesignBasis(0), lblDesignParameters(0))
    
    Case 1        'Tower Height
      Call Unitted_UnitChange(UNITS_LENGTH, Scr2.TowerHeight.value, UnitsDesignBasis(1), lblDesignParameters(1))
  
  End Select

End Sub

Private Sub UnitsDesignBasis_KeyPress(Index As Integer, KeyAscii As Integer)
  Call Combobox_KeyPress(KeyAscii)
End Sub

Private Sub UnitsFlows_Click(Index As Integer)

  Select Case Index
    Case 0        'Water Flow Rate
      Call Unitted_UnitChange(UNITS_FLOW, Scr2.WaterFlowRate.value, UnitsFlows(0), txtFlowsLoadings(0))
    
    Case 1        'Air Flow Rate
      Call Unitted_UnitChange(UNITS_FLOW, Scr2.AirFlowRate.value, UnitsFlows(1), txtFlowsLoadings(1))
  
    Case 3        'Water Loading Rate
      Call Unitted_UnitChange(UNITS_MASSLOADINGRATE, Scr2.WaterLoadingRate.value, UnitsFlows(3), txtFlowsLoadings(3))
  
    Case 4        'Air Loading Rate
      Call Unitted_UnitChange(UNITS_MASSLOADINGRATE, Scr2.AirLoadingRate.value, UnitsFlows(4), txtFlowsLoadings(4))
  
  End Select

End Sub

Private Sub UnitsFlows_KeyPress(Index As Integer, KeyAscii As Integer)
  Call Combobox_KeyPress(KeyAscii)
End Sub

Private Sub UnitsInterest_Click(Index As Integer)
Dim Dummy As Double
Dim ContaminantIndex As Integer
Dim i As Integer

  ContaminantIndex = cboSelectCompo.ListIndex + 1
  i = ContaminantIndex

  On Error GoTo err_UnitsInterest_Click

  Select Case Index
    Case 0            'Onda KLa
      Dummy = Scr2.Onda.OverallMassTransferCoefficient
      Call Unitted_UnitChange(UNITS_INVERSETIME, Dummy, UnitsInterest(0), lblDesignConcentrationValue(0))

    Case 2            'Design KLa
      Dummy = Scr2.DesignMassTransferCoefficient.value
      Call Unitted_UnitChange(UNITS_INVERSETIME, Dummy, UnitsInterest(2), txtDesignConcentrationValue(2))

    Case 3            'Influent Concentration
      Dummy = Scr2.DesignContaminant.Influent.value
      Call Unitted_UnitChange(UNITS_CONCENTRATION, Dummy, UnitsInterest(3), lblDesignConcentrationValue(3))

    Case 4            'Treatment Objective
      Dummy = Scr2.DesignContaminant.TreatmentObjective.value
      Call Unitted_UnitChange(UNITS_CONCENTRATION, Dummy, UnitsInterest(4), lblDesignConcentrationValue(4))

    Case 5            'Effluent Concentration
      Dummy = Scr2.DesignContaminant.Effluent.value
      Call Unitted_UnitChange(UNITS_CONCENTRATION, Dummy, UnitsInterest(5), lblDesignConcentrationValue(5))

    Case 7            'Air Pressure Drop
      Dummy = Scr2.AirPressureDrop.value
      Call Unitted_UnitChange(UNITS_PRESSUREPERLENGTH, Dummy, UnitsInterest(7), lblDesignConcentrationValue(7))
      If (Dummy < -1) Then
        lblDesignConcentrationValue(7) = "N/A"
      End If

  End Select

exit_UnitsInterest_Click:
  Exit Sub

err_UnitsInterest_Click:
  Resume exit_UnitsInterest_Click

End Sub

Private Sub UnitsInterest_KeyPress(Index As Integer, KeyAscii As Integer)
  Call Combobox_KeyPress(KeyAscii)
End Sub

Private Sub UnitsOpCond_Click(Index As Integer)
Dim Dummy As Double

  Select Case Index
    Case 0
      'Note: Standard P units are Pa, but OperatingPressure
      'is stored internally in kPa units.
      Dummy = Scr2.OperatingPressure.value * 101325#
      Call Unitted_UnitChange(UNITS_PRESSURE, Dummy, UnitsOpCond(0), txtOperatingPressure)

    Case 1
      Call Unitted_UnitChange(UNITS_TEMPERATURE, Scr2.operatingtemperature.value, UnitsOpCond(1), txtOperatingTemperature)
    
  End Select

End Sub

Private Sub UnitsOpCond_KeyPress(Index As Integer, KeyAscii As Integer)
  Call Combobox_KeyPress(KeyAscii)
End Sub

Private Sub UnitsTowerParam_Click(Index As Integer)

  Select Case Index
    Case 0        'Specify Diameter
      Call Unitted_UnitChange(UNITS_LENGTH, Scr2.SpecifiedTowerDiameter.value, UnitsTowerParam(0), txtTowerParameters(0))
    
    Case 1        'Specify Height
      Call Unitted_UnitChange(UNITS_LENGTH, Scr2.SpecifiedTowerHeight.value, UnitsTowerParam(1), txtTowerParameters(1))
  
    Case 2        'Tower Area
      Call Unitted_UnitChange(UNITS_AREA, Scr2.TowerArea.value, UnitsTowerParam(2), lblTowerParameters(2))
  
    Case 3        'Tower Volume
      Call Unitted_UnitChange(UNITS_VOLUME, Scr2.TowerVolume.value, UnitsTowerParam(3), lblTowerParameters(3))
  
  End Select

End Sub

Private Sub UnitsTowerParam_KeyPress(Index As Integer, KeyAscii As Integer)
  Call Combobox_KeyPress(KeyAscii)
End Sub


