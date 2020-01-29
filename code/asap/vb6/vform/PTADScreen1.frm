VERSION 5.00
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "THREED32.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmPTADScreen1 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Packed Tower Aeration - Design Mode"
   ClientHeight    =   6795
   ClientLeft      =   990
   ClientTop       =   1740
   ClientWidth     =   9480
   Icon            =   "PTADScreen1.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   ScaleHeight     =   6795
   ScaleWidth      =   9480
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   870
      Top             =   5430
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
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
      Left            =   930
      TabIndex        =   15
      TabStop         =   0   'False
      Top             =   5910
      Width           =   2655
   End
   Begin Threed.SSFrame fraOperatingConditions 
      Height          =   1455
      Left            =   90
      TabIndex        =   9
      Top             =   120
      Width           =   4335
      _Version        =   65536
      _ExtentX        =   7646
      _ExtentY        =   2561
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
         Left            =   120
         TabIndex        =   18
         TabStop         =   0   'False
         Top             =   990
         Width           =   4095
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
         Left            =   1620
         TabIndex        =   1
         Top             =   630
         Width           =   1215
      End
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
         Left            =   1620
         TabIndex        =   0
         Top             =   270
         Width           =   1215
      End
      Begin VB.ComboBox txtPUnits 
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
         Left            =   2940
         Style           =   2  'Dropdown List
         TabIndex        =   17
         TabStop         =   0   'False
         Top             =   270
         Width           =   1275
      End
      Begin VB.ComboBox txtTUnits 
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
         Left            =   2940
         Style           =   2  'Dropdown List
         TabIndex        =   16
         TabStop         =   0   'False
         Top             =   630
         Width           =   1275
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
         Left            =   300
         TabIndex        =   20
         Top             =   690
         Width           =   1215
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
         Left            =   300
         TabIndex        =   19
         Top             =   330
         Width           =   1215
      End
   End
   Begin Threed.SSFrame fraPackingInformation 
      Height          =   1215
      Left            =   90
      TabIndex        =   10
      Top             =   1830
      Width           =   4335
      _Version        =   65536
      _ExtentX        =   7646
      _ExtentY        =   2138
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
         Left            =   120
         TabIndex        =   22
         TabStop         =   0   'False
         Top             =   720
         Width           =   4095
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
         Left            =   120
         TabIndex        =   21
         TabStop         =   0   'False
         Top             =   360
         Width           =   975
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
         Left            =   1200
         TabIndex        =   24
         Top             =   360
         Width           =   3015
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
         Left            =   120
         TabIndex        =   23
         Top             =   360
         Width           =   855
      End
   End
   Begin Threed.SSFrame fraContaminantInformation 
      Height          =   1815
      Left            =   90
      TabIndex        =   11
      Top             =   3330
      Width           =   4335
      _Version        =   65536
      _ExtentX        =   7646
      _ExtentY        =   3201
      _StockProps     =   14
      Caption         =   "Design Contaminant:"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
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
         Left            =   240
         Style           =   2  'Dropdown List
         TabIndex        =   29
         TabStop         =   0   'False
         Top             =   480
         Width           =   3855
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
         Left            =   240
         TabIndex        =   28
         TabStop         =   0   'False
         Top             =   840
         Width           =   1035
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
         Left            =   1380
         TabIndex        =   27
         TabStop         =   0   'False
         Top             =   840
         Width           =   975
      End
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
         Left            =   2460
         TabIndex        =   26
         TabStop         =   0   'False
         Top             =   840
         Width           =   1635
      End
      Begin VB.CommandButton cmdDesignContaminant 
         Appearance      =   0  'Flat
         Caption         =   "Optimize with All Contaminants"
         Enabled         =   0   'False
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
         Left            =   120
         TabIndex        =   25
         TabStop         =   0   'False
         Top             =   1320
         Width           =   4095
      End
      Begin VB.Shape Shape1 
         Height          =   795
         Left            =   120
         Top             =   420
         Width           =   4095
      End
   End
   Begin Threed.SSFrame fraFlowsLoadings 
      Height          =   3135
      Left            =   4410
      TabIndex        =   12
      Top             =   120
      Width           =   4935
      _Version        =   65536
      _ExtentX        =   8705
      _ExtentY        =   5530
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
         Index           =   5
         Left            =   2220
         TabIndex        =   6
         Top             =   2040
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
         Index           =   4
         Left            =   2220
         TabIndex        =   5
         Top             =   1680
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
         Left            =   2220
         TabIndex        =   4
         Top             =   1320
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
         Left            =   2220
         TabIndex        =   3
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
         Index           =   0
         Left            =   2220
         TabIndex        =   2
         Top             =   240
         Width           =   1215
      End
      Begin VB.ComboBox txtFlowsUnits 
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
         Left            =   3540
         Style           =   2  'Dropdown List
         TabIndex        =   34
         TabStop         =   0   'False
         Top             =   240
         Width           =   1275
      End
      Begin VB.ComboBox txtFlowsUnits 
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
         Left            =   3540
         Style           =   2  'Dropdown List
         TabIndex        =   33
         TabStop         =   0   'False
         Top             =   1680
         Width           =   1275
      End
      Begin VB.ComboBox txtFlowsUnits 
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
         Left            =   3540
         Style           =   2  'Dropdown List
         TabIndex        =   32
         TabStop         =   0   'False
         Top             =   2040
         Width           =   1275
      End
      Begin VB.ComboBox lblFlowsUnits 
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
         Index           =   6
         Left            =   3540
         Style           =   2  'Dropdown List
         TabIndex        =   31
         TabStop         =   0   'False
         Top             =   2400
         Width           =   1275
      End
      Begin VB.ComboBox lblFlowsUnits 
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
         Left            =   3540
         Style           =   2  'Dropdown List
         TabIndex        =   30
         TabStop         =   0   'False
         Top             =   2760
         Width           =   1275
      End
      Begin VB.Label lblFlowsLoadings 
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
         ForeColor       =   &H00808080&
         Height          =   285
         Index           =   7
         Left            =   2220
         TabIndex        =   48
         Top             =   2760
         Width           =   1215
      End
      Begin VB.Label lblFlowsLoadings 
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
         ForeColor       =   &H00808080&
         Height          =   285
         Index           =   6
         Left            =   2220
         TabIndex        =   47
         Top             =   2400
         Width           =   1215
      End
      Begin VB.Label lblFlowsLoadings 
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
         ForeColor       =   &H00808080&
         Height          =   285
         Index           =   1
         Left            =   2220
         TabIndex        =   46
         Top             =   600
         Width           =   1215
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
         Index           =   7
         Left            =   120
         TabIndex        =   45
         Top             =   2820
         Width           =   1995
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
         Index           =   6
         Left            =   120
         TabIndex        =   44
         Top             =   2460
         Width           =   1995
      End
      Begin VB.Label lblFlowsLoadingsLabel 
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
         Index           =   5
         Left            =   120
         TabIndex        =   43
         Top             =   2100
         Width           =   1995
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
         Index           =   4
         Left            =   120
         TabIndex        =   42
         Top             =   1740
         Width           =   1995
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
         Index           =   3
         Left            =   120
         TabIndex        =   41
         Top             =   1380
         Width           =   1995
      End
      Begin VB.Label lblFlowsLoadingsLabel 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Multiple of (V/Q)min"
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
         Left            =   120
         TabIndex        =   40
         Top             =   1020
         Width           =   1995
      End
      Begin VB.Label lblFlowsLoadingsLabel 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Min Air to Water Ratio"
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
         TabIndex        =   39
         Top             =   660
         Width           =   1995
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
         Left            =   120
         TabIndex        =   38
         Top             =   300
         Width           =   1995
      End
      Begin VB.Label lblMisc 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "( vol./vol. )"
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
         Left            =   3540
         TabIndex        =   37
         Top             =   660
         Width           =   1275
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
         Left            =   3540
         TabIndex        =   36
         Top             =   1020
         Width           =   1275
      End
      Begin VB.Label lblMisc 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "( vol./vol. )"
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
         Left            =   3540
         TabIndex        =   35
         Top             =   1380
         Width           =   1275
      End
   End
   Begin Threed.SSFrame fraMassTransferParamters 
      Height          =   1335
      Left            =   4410
      TabIndex        =   13
      Top             =   3330
      Width           =   4935
      _Version        =   65536
      _ExtentX        =   8705
      _ExtentY        =   2355
      _StockProps     =   14
      Caption         =   "Mass Transfer Parameters:"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Begin VB.TextBox txtMassTransfer 
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
         Left            =   2220
         TabIndex        =   8
         Top             =   960
         Width           =   1215
      End
      Begin VB.TextBox txtMassTransfer 
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
         Left            =   2220
         TabIndex        =   7
         Top             =   600
         Width           =   1215
      End
      Begin VB.ComboBox UnitsMassTransfer 
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
         Left            =   3540
         Style           =   2  'Dropdown List
         TabIndex        =   51
         TabStop         =   0   'False
         Top             =   240
         Width           =   1275
      End
      Begin VB.ComboBox UnitsMassTransfer 
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
         Left            =   3540
         Style           =   2  'Dropdown List
         TabIndex        =   50
         TabStop         =   0   'False
         Top             =   960
         Width           =   1275
      End
      Begin VB.CommandButton cmdMassTransferLabel 
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
         Height          =   255
         Left            =   120
         TabIndex        =   49
         TabStop         =   0   'False
         Top             =   255
         Width           =   1995
      End
      Begin VB.Label lblMassTransfer 
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
         ForeColor       =   &H00808080&
         Height          =   285
         Index           =   0
         Left            =   2220
         TabIndex        =   55
         Top             =   240
         Width           =   1215
      End
      Begin VB.Label lblMassTransferLabel 
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
         Left            =   360
         TabIndex        =   54
         Top             =   1020
         Width           =   1755
      End
      Begin VB.Label lblMassTransferLabel 
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
         Left            =   360
         TabIndex        =   53
         Top             =   660
         Width           =   1755
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
         Index           =   3
         Left            =   3540
         TabIndex        =   52
         Top             =   660
         Width           =   1275
      End
   End
   Begin Threed.SSFrame fraTowerParamters 
      Height          =   1545
      Left            =   4410
      TabIndex        =   14
      Top             =   4740
      Width           =   4935
      _Version        =   65536
      _ExtentX        =   8705
      _ExtentY        =   2725
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
      Begin VB.ComboBox lblTowerUnits 
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
         Left            =   3540
         Style           =   2  'Dropdown List
         TabIndex        =   59
         TabStop         =   0   'False
         Top             =   270
         Width           =   1275
      End
      Begin VB.ComboBox lblTowerUnits 
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
         Left            =   3540
         Style           =   2  'Dropdown List
         TabIndex        =   58
         TabStop         =   0   'False
         Top             =   570
         Width           =   1275
      End
      Begin VB.ComboBox lblTowerUnits 
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
         Left            =   3540
         Style           =   2  'Dropdown List
         TabIndex        =   57
         TabStop         =   0   'False
         Top             =   870
         Width           =   1275
      End
      Begin VB.ComboBox lblTowerUnits 
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
         Left            =   3540
         Style           =   2  'Dropdown List
         TabIndex        =   56
         TabStop         =   0   'False
         Top             =   1170
         Width           =   1275
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
         ForeColor       =   &H00808080&
         Height          =   255
         Index           =   3
         Left            =   2220
         TabIndex        =   67
         Top             =   1170
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
         ForeColor       =   &H00808080&
         Height          =   255
         Index           =   2
         Left            =   2220
         TabIndex        =   66
         Top             =   870
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
         ForeColor       =   &H00808080&
         Height          =   255
         Index           =   1
         Left            =   2220
         TabIndex        =   65
         Top             =   570
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
         ForeColor       =   &H00808080&
         Height          =   255
         Index           =   0
         Left            =   2220
         TabIndex        =   64
         Top             =   270
         Width           =   1215
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
         Left            =   480
         TabIndex        =   63
         Top             =   1170
         Width           =   1635
      End
      Begin VB.Label lblTowerParametersLabel 
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
         Index           =   2
         Left            =   480
         TabIndex        =   62
         Top             =   870
         Width           =   1635
      End
      Begin VB.Label lblTowerParametersLabel 
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
         Index           =   1
         Left            =   480
         TabIndex        =   61
         Top             =   570
         Width           =   1635
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
         Index           =   0
         Left            =   480
         TabIndex        =   60
         Top             =   270
         Width           =   1635
      End
   End
   Begin Threed.SSFrame xFrame3D1 
      Height          =   1875
      Left            =   9000
      TabIndex        =   68
      Top             =   6270
      Visible         =   0   'False
      Width           =   4335
      _Version        =   65536
      _ExtentX        =   7646
      _ExtentY        =   3307
      _StockProps     =   14
      Caption         =   "Still Needed until labels are de-linked!"
      ForeColor       =   255
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Begin VB.Label xlblDesignConcentration 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Influent Concentration"
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   0
         Left            =   180
         TabIndex        =   75
         Top             =   270
         Width           =   2535
      End
      Begin VB.Label xlblDesignConcentration 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Treatment Objective"
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   1
         Left            =   180
         TabIndex        =   74
         Top             =   570
         Width           =   2535
      End
      Begin VB.Label xlblDesignConcentration 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Percent Removal"
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   2
         Left            =   180
         TabIndex        =   73
         Top             =   870
         Width           =   2535
      End
      Begin VB.Label lblDesignConcentrationValue 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   0
         Left            =   2940
         TabIndex        =   72
         Top             =   270
         Width           =   1095
      End
      Begin VB.Label lblDesignConcentrationValue 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   1
         Left            =   2940
         TabIndex        =   71
         Top             =   570
         Width           =   1095
      End
      Begin VB.Label lblDesignConcentrationValue 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   2
         Left            =   2940
         TabIndex        =   70
         Top             =   870
         Width           =   1095
      End
      Begin VB.Label lblMassTransferLabel 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Onda KLa"
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   0
         Left            =   660
         TabIndex        =   69
         Top             =   1290
         Visible         =   0   'False
         Width           =   1755
      End
   End
   Begin VB.Menu mnuFileMenu 
      Caption         =   "&File"
      Begin VB.Menu mnuFile 
         Caption         =   "Switch to &Rating Mode"
         Index           =   0
         Shortcut        =   ^R
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
      Begin VB.Menu mnuOptions 
         Caption         =   "&View All Concentration Results"
         Index           =   0
      End
      Begin VB.Menu mnuOptions 
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
Attribute VB_Name = "frmPTADScreen1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim Temp_Text As String
Dim frmPTADScreen1_Okay_To_Unload As Boolean




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
    fraFlowsLoadings.Caption = "* DEMONSTRATION VERSION *"
    fraFlowsLoadings.ForeColor = QBColor(12)
  End If
End Sub


Private Sub AddPrompt(menuID As Integer, prompt As String)
    menuPrompts(iMenuPrompts).menuID = menuID
    menuPrompts(iMenuPrompts).prompt = prompt
    iMenuPrompts = iMenuPrompts + 1
End Sub

Private Sub cboDesignContaminant_Click()
    
End Sub

Private Sub cboDesignContaminant_GotFocus()
'    Temp_Text = cboDesignContaminant.Text
End Sub

Private Sub cboDesignContaminant_KeyPress(KeyAscii As Integer)
    'Dim i As Integer
    'Dim ComboText As String
    '
    'If KeyAscii <> 13 Then Exit Sub
    'KeyAscii = 0

End Sub

Private Sub cboDesignContaminant_LostFocus()
    'Dim KeyAscii As Integer
    '
    'KeyAscii = 13
    'cboDesignContaminant_KeyPress (KeyAscii)
End Sub

Private Sub cboSelectCompo_Click()
    Dim ContaminantIndex As Integer, i As Integer
    Dim PercentRemoval As Double
    
    ContaminantIndex = cboSelectCompo.ListIndex + 1
    i = ContaminantIndex
    If i = 0 Then Exit Sub
    lblDesignConcentrationValue(0).Caption = Format$(scr1.Contaminant(i).Influent.value, GetTheFormat(scr1.Contaminant(i).Influent.value))
    lblDesignConcentrationValue(1).Caption = Format$(scr1.Contaminant(i).TreatmentObjective.value, GetTheFormat(scr1.Contaminant(i).TreatmentObjective.value))
    Call REMOVPT(PercentRemoval, scr1.Contaminant(i).Influent.value, scr1.Contaminant(i).TreatmentObjective.value)
    lblDesignConcentrationValue(2).Caption = Format$(PercentRemoval, "0.0")
    Call PT1VQMIN(scr1.MinimumAirToWaterRatio.value, scr1.Contaminant(i).Influent.value, scr1.Contaminant(i).TreatmentObjective.value, scr1.Contaminant(i).HenrysConstant.value)
    lblFlowsLoadings(1).Caption = Format$(scr1.MinimumAirToWaterRatio.value, GetTheFormat(scr1.MinimumAirToWaterRatio.value))
    
    scr1.DesignContaminant = scr1.Contaminant(i)
    
    'Update Variables on Screen
    Call GetVQmultVQAndAirFlowRate
    Call GetLoadings
    Call GetTowerAreaAndDiameter
    Call GetOndaMassTransferCoefficient
    Call GetDesignKLaOrKLaSafetyFactor
    Call GetTowerHeightAndVolume

End Sub

Private Sub cmdAddComponent_Click()
Dim x As rec_frmContaminantPropertyEdit
Dim i As Integer

  If (scr1.NumChemical + 1 > MAXCHEMICAL) Then
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
    x.Contaminants(i) = scr1.Contaminant(i)
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

      scr1.Contaminant(i) = x.Contaminants(i)
      scr1.NumChemical = scr1.NumChemical + 1
      'Incorporate new name into ComboBox.
      cboSelectCompo.AddItem scr1.Contaminant(i).Name
    Next i
  End If
  
  Call SetDesignContaminantEnabled(CInt(cboSelectCompo.ListCount))
  
  If (scr1.NumChemical > 0) Then
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

  scr1.Chemical = cboSelectCompo.ListIndex + 1
  If (scr1.Chemical = 0) Then Exit Sub

  If MsgBox("Remove" & NL & cboSelectCompo.List(cboSelectCompo.ListIndex), 36, "") = IDYES Then
    cboSelectCompo.RemoveItem cboSelectCompo.ListIndex
    For i = scr1.Chemical To scr1.NumChemical - 1
      scr1.Contaminant(i) = scr1.Contaminant(i + 1)
    Next i
    scr1.NumChemical = scr1.NumChemical - 1
    If (scr1.NumChemical > 0) Then
      cboSelectCompo.ListIndex = 0
    Else
      cmdDeleteComponent.Enabled = False
      cmdEditComponent.Enabled = False
    End If
    Call SetDesignContaminantEnabled(CInt(cboSelectCompo.ListCount))
  End If

    If scr1.NumChemical < 10 Then cmdAddComponent.Enabled = True
  Call LOCAL___Reset_DemoVersionDisablings
End Sub

Private Sub cmdDesignContaminant_Click()
    Call OptimizeDesignContaminant
End Sub

Private Sub cmdEditComponent_Click()
Dim x As rec_frmContaminantPropertyEdit
Dim i As Integer
Dim AListIndex As Integer

  scr1.Chemical = cboSelectCompo.ListIndex + 1
  If (scr1.Chemical = 0) Then Exit Sub

  x.ModelName = "Packed Tower Aeration"
  x.ModelType = MODELTYPE_PACKEDTOWER
  x.DoEditNumber = cboSelectCompo.ListIndex + 1
  x.DoAdd = False
  x.OldNumCompo = cboSelectCompo.ListCount
  For i = 1 To x.OldNumCompo
    x.Contaminants(i) = scr1.Contaminant(i)
  Next i

  Data_frmContaminantPropertyEdit = x
  frmContaminantPropertyEdit.Show 1
  x = Data_frmContaminantPropertyEdit

  If (Not x.CancelledEdit) Then
    For i = 1 To x.NewNumCompo
      scr1.Contaminant(i) = x.Contaminants(i)
    Next i
    If (x.OldNumCompo < x.NewNumCompo) Then
      'Incorporate new names into ComboBox.
      For i = x.OldNumCompo + 1 To x.NewNumCompo
        cboSelectCompo.AddItem scr1.Contaminant(i).Name
      Next i
    End If
    'Update ComboBox for any changed names:
    For i = 1 To x.OldNumCompo
      If (Trim$(cboSelectCompo.List(i - 1)) <> Trim$(scr1.Contaminant(i).Name)) Then
        cboSelectCompo.List(i - 1) = Trim$(scr1.Contaminant(i).Name)
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

Private Sub cmdMassTransferLabel_Click()

    If frmPTADScreen1.lblMassTransfer(0).Caption = "0.0" Then Exit Sub
    frmShowOndaKLaProperties.Show 1

End Sub

Private Sub cmdPackingType_Click()

  ScreenNumber = 1
  CurrentScreen = scr1
  Call ShowPackingProperties

End Sub

Private Sub cmdSelectContaminants_Click()

    'frmListContaminant.Show 1

End Sub

Private Sub cmdSelectPacking_Click()
    Dim i As Integer, CurrPackingIndex As Integer
    
    ReadMainPackingDB
    ReadUserPackingDB

    PackingDatabaseSource = scr1.Packing.SourceDatabase

    If scr1.Packing.Name = "" Then
       scr1.Packing.Name = frmSelectPacking.cboSelectPacking.List(0)
       PackingChanged = True
    End If

    If PackingDatabaseSource = ORIGINALPACKINGDATABASE Then
       For i = 1 To NumPackingsInDatabase
           If DatabasePacking(i).Name = scr1.Packing.Name Then
              CurrPackingIndex = i
              frmSelectPacking!cboSelectPacking.ListIndex = CurrPackingIndex - 1
              Exit For
           End If
       Next i
    ElseIf PackingDatabaseSource = USERMODIFIEDPACKINGDATABASE Then
       For i = 1 To NumUserPackings
           If UserPacking(i).Name = scr1.Packing.Name Then
              CurrPackingIndex = i
              frmSelectPacking.cboSelectPacking.ListIndex = CurrPackingIndex - 1
              Exit For
           End If
       Next i

    End If
    ScreenNumber = 1
    CurrentScreen = scr1
    frmSelectPacking.Show 1
    
    'Reflect changes in packing into Scr1
    scr1 = CurrentScreen

    If scr1.NumChemical > 0 Then
       'Update Variables on Screen
       Call GetVQmultVQAndAirFlowRate
       Call GetLoadings
       Call GetTowerAreaAndDiameter
       Call GetOndaMassTransferCoefficient
       Call GetDesignKLaOrKLaSafetyFactor
       Call GetTowerHeightAndVolume
    End If

End Sub

Private Sub Command1_Click()

Call screen1_results
 

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
'    AddPrompt GetMenuItemID(hSubMenu, 0), "Switch to Rating Mode for packed tower aeration"
'    AddPrompt GetMenuItemID(hSubMenu, 2), "Load a design case from a file"
'    AddPrompt GetMenuItemID(hSubMenu, 3), "Save this design case to a file"
'    AddPrompt GetMenuItemID(hSubMenu, 4), "Save this design case to a file"
'    AddPrompt GetMenuItemID(hSubMenu, 6), "Print this design case"
'    AddPrompt GetMenuItemID(hSubMenu, 7), "Select a printer for printing results"
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
'    '
'    ' Load Options menu prompt
'    '
'    hSubMenu = GetSubMenu(hMenu, 3)
'    AddPrompt hSubMenu, "Options"
'    AddPrompt GetMenuItemID(hSubMenu, 0), "View Effluent Concentration Results for All Contaminants"

    'Initialize last-few-files list.
    Call LastFewFiles_InitializeList(LASTFEW_WHICHAPP_ASAP, LASTFEW_ASAP_frmPTADScreen1)

End Sub

Private Sub Form_Load()
    Dim i As Integer
    
  frmPTADScreen1_Okay_To_Unload = False

scr1.NumChemical = 0
frmPTADScreen1.Width = SCREEN_WIDTH_STANDARD
frmPTADScreen1.Height = SCREEN_HEIGHT_STANDARD
    
    'Center the form on the screen
    If WindowState = 0 Then
       'don't attempt if screen Minimized or Maximized
       Move (Screen.Width - frmPTADScreen1.Width) / 2, (Screen.Height - frmPTADScreen1.Height) / 2
    End If
    
    'Initialize Labels on frmPTADScreen1
    Call LabelsPTADScreen1(UNITSTYPE_SI)

    'Initialize Values for Pressure and Temperature
    Call InitializePressureTemperature

    'Initialize values for water density, water viscosity,
    'water surface tension, air density, and air viscosity
    'based on default temperature and pressure
    Call CalculateAirWaterProperties

    'Initialize design packing to a default
    Call InitializePacking
    
    'Initialize value for Water Flow Rate
    Call InitializeWaterFlowRate

    'Initialize value for Multiple of Minimum Air to Water Ratio
    Call InitializeVQminMultiple

    'Initialize Value for Air Pressure Drop
    Call InitializeAirPressureDrop

    'Initialize Value for KLaSafetyFactor
    Call InitializeKLaSafetyFactor

    'Initialize calculated properties text boxes to 0 and disabled
    Call InitializeCalculatedProperties

    Load frmAirWaterProperties
    Load frmShowPackingProperties
    'Load frmListContaminant

    Call LastFewFiles_InitializeList(LASTFEW_WHICHAPP_ASAP, LASTFEW_ASAP_frmPTADScreen1)
    '
    ' DEMO SETTINGS.
    '
    Call LOCAL___Reset_DemoVersionDisablings
End Sub
Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
  If (frmPTADScreen1_Okay_To_Unload) Then
    Cancel = False
  Else
    Cancel = True
  End If
End Sub


Private Sub lblDisplayAirWaterProperties_Click()
    If HaveValue(scr1.OperatingPressure.value) And HaveValue(scr1.operatingtemperature.value) Then
       CurrentScreen = scr1
       CurrentMode = 1
       frmAirWaterProperties.Show 1
    Else
       MsgBox "You must specify pressure and temperature before physical properties can be displayed.", MB_ICONSTOP, "Error"
    End If
End Sub

Private Sub lblFlowsLoadings_Change(Index As Integer)

  If (Index = 6) Or (Index = 7) Then
    Call lblFlowsUnits_Click(Index)
  End If

End Sub

Private Sub lblFlowsUnits_Click(Index As Integer)

  Select Case Index
    Case 6        'Air Loading Rate
      Call Unitted_UnitChange(UNITS_MASSLOADINGRATE, scr1.AirLoadingRate.value, lblFlowsUnits(6), lblFlowsLoadings(6))

    Case 7        'Water Loading Rate
      Call Unitted_UnitChange(UNITS_MASSLOADINGRATE, scr1.WaterLoadingRate.value, lblFlowsUnits(7), lblFlowsLoadings(7))
  End Select

End Sub

Private Sub lblFlowsUnits_KeyPress(Index As Integer, KeyAscii As Integer)
  Call Combobox_KeyPress(KeyAscii)
End Sub

Private Sub lblMassTransfer_Change(Index As Integer)

  Call UnitsMassTransfer_Click(0)

End Sub

Private Sub lblPackingType_Click()

  Call cmdPackingType_Click
  
  End Sub

Private Sub lblTowerParameters_Change(Index As Integer)

  Call lblTowerUnits_Click(Index)

End Sub

Private Sub lblTowerUnits_Click(Index As Integer)

  Select Case Index
    Case 0        'Tower Area
      Call Unitted_UnitChange(UNITS_AREA, scr1.TowerArea.value, lblTowerUnits(0), lblTowerParameters(0))
    
    Case 1        'Tower Diameter
      Call Unitted_UnitChange(UNITS_LENGTH, scr1.TowerDiameter.value, lblTowerUnits(1), lblTowerParameters(1))
    
    Case 2        'Tower Height
      Call Unitted_UnitChange(UNITS_LENGTH, scr1.TowerHeight.value, lblTowerUnits(2), lblTowerParameters(2))
    
    Case 3        'Tower Volume
      Call Unitted_UnitChange(UNITS_VOLUME, scr1.TowerVolume.value, lblTowerUnits(3), lblTowerParameters(3))
    
  End Select

End Sub

Private Sub lblTowerUnits_KeyPress(Index As Integer, KeyAscii As Integer)
  Call Combobox_KeyPress(KeyAscii)
End Sub

Private Sub mnuFile_Click(Index As Integer)
    Dim i As Integer
    Dim msg As String, Response As Integer

    Screen.MousePointer = 11   'Hourglass

    Select Case Index
       Case 0   'Switch to Rating Mode

          If frmPTADScreen1!lblTowerParameters(3).Caption = "0.0" Then
             Filename$ = "TheDefaultCaseScreen2"
             Call loadscreen2("")
             frmPTADScreen1.Hide
             frmPTADScreen2.Show
             Screen.MousePointer = 0
             Exit Sub
          End If

          If HaveValue(scr1.TowerVolume.value) Then
             'Give user option to save screen 1 before
             'switching to screen 2
             msg = "Would you like to save the parameters "
             msg = msg + "for this design case to a file "
             msg = msg + "before switching to Rating "
             msg = msg + "Mode?"
             Response = MsgBox(msg, MB_ICONquestion + MB_YESNO, "Save Current Design")
             If Response = IDYES Then
                Call SaveScreen1
             End If

             'Initialize Values on Screen 2:

             'First, switch back to SI units so we can do a direct copy of the textboxes from this screen to that screen.
             Call LabelsPTADScreen1(UNITSTYPE_SI)
             
             'Design Tower Diameter
             frmPTADScreen2!lblDesignParameters(0).Caption = frmPTADScreen1!lblTowerParameters(1).Caption
             Scr2.TowerDiameter = scr1.TowerDiameter

             'Design Tower Height
             frmPTADScreen2!lblDesignParameters(1).Caption = frmPTADScreen1!lblTowerParameters(2).Caption
             Scr2.TowerHeight = scr1.TowerHeight
       
             'UserSpecified Tower Diameter and Tower Height
             Scr2.SpecifiedTowerDiameter.value = scr1.TowerDiameter.value
             Scr2.SpecifiedTowerHeight.value = scr1.TowerHeight.value
             frmPTADScreen2!txtTowerParameters(0).Text = frmPTADScreen2!lblDesignParameters(0).Caption
             frmPTADScreen2!txtTowerParameters(1).Text = frmPTADScreen2!lblDesignParameters(1).Caption

             'Calculate Tower Area and Volume with specified tower
             'diameter and tower height
             Call GetTowerAreaAndVolume
             frmPTADScreen2!lblTowerParameters(2).Caption = Format$(Scr2.TowerArea.value, GetTheFormat(Scr2.TowerArea.value))
             frmPTADScreen2!lblTowerParameters(3).Caption = Format$(Scr2.TowerVolume.value, GetTheFormat(Scr2.TowerVolume.value))

             'Flow and Loading Parameters
             frmPTADScreen2!txtFlowsLoadings(0).Text = frmPTADScreen1!txtFlowsLoadings(0).Text
             frmPTADScreen2!txtFlowsLoadings(0).Enabled = True
             Scr2.WaterFlowRate = scr1.WaterFlowRate
             Scr2.WaterFlowRate.UserInput = True
             If scr1.AirFlowRate.UserInput = True Then
                'Water Flow Rate and air flow rate are initial
                'two flow and loading values that are user input
                'on screen 2
                frmPTADScreen2!txtFlowsLoadings(1).Text = frmPTADScreen1!txtFlowsLoadings(4).Text
                frmPTADScreen2!txtFlowsLoadings(1).Enabled = True
                Scr2.AirFlowRate = scr1.AirFlowRate
                Scr2.AirFlowRate.UserInput = True
                For i = 2 To 4
                    frmPTADScreen2!txtFlowsLoadings(i).Enabled = False
                    frmPTADScreen2!txtFlowsLoadings(i).Text = ""
                Next i
                Scr2.AirToWaterRatio.UserInput = False
                Scr2.AirLoadingRate.UserInput = False
                Scr2.WaterLoadingRate.UserInput = False
                frmFlowsLoadingsScreen2!optFlowsLoadings(0).value = True
                UsersFlowsLoadingsOption = 0
             Else
                'Water flow rate and air to water ratio are
                'initial two flow and loading values that are
                'user-input on screen 2
                frmPTADScreen2!txtFlowsLoadings(2).Text = frmPTADScreen1!txtFlowsLoadings(3).Text
                frmPTADScreen2!txtFlowsLoadings(2).Enabled = True
                Scr2.AirToWaterRatio = scr1.AirToWaterRatio
                Scr2.AirToWaterRatio.UserInput = True
                For i = 1 To 4
                    If i <> 2 Then
                       frmPTADScreen2!txtFlowsLoadings(i).Enabled = False
                       frmPTADScreen2!txtFlowsLoadings(i).Text = ""
                    End If
                Next i
                Scr2.AirFlowRate.UserInput = False
                Scr2.AirLoadingRate.UserInput = False
                Scr2.WaterLoadingRate.UserInput = False
                frmFlowsLoadingsScreen2!optFlowsLoadings(1).value = True
                UsersFlowsLoadingsOption = 1
             End If
       

             'Pressure and Temperature
             frmPTADScreen2!txtOperatingPressure.Text = frmPTADScreen1!txtOperatingPressure.Text
             Scr2.OperatingPressure = scr1.OperatingPressure
             Scr2.OperatingPressure.ValChanged = True
             frmPTADScreen2!txtOperatingTemperature.Text = frmPTADScreen1!txtOperatingTemperature.Text
             Scr2.operatingtemperature = scr1.operatingtemperature
             Scr2.operatingtemperature.ValChanged = True

             'Physical Properties
             Call CalculateAirWaterPropertiesScreen2

             'Packing
             frmPTADScreen2!lblPackingType.Caption = frmPTADScreen1!lblPackingType.Caption
             Scr2.Packing = scr1.Packing

             'Contaminant List
             Scr2.NumChemical = scr1.NumChemical
             frmPTADScreen2!cboSelectCompo.Clear
             'frmListcontaminantScreen2!ListContaminants.Clear
             If Scr2.NumChemical > 0 Then
                For i = 1 To Scr2.NumChemical
                    Scr2.Contaminant(i) = scr1.Contaminant(i)
                    frmPTADScreen2!cboSelectCompo.AddItem Scr2.Contaminant(i).Name
                    'frmListcontaminantScreen2!ListContaminants.AddItem Scr2.Contaminant(i).Name
                Next i


                'frmListcontaminantScreen2!mnuOptionsManipulateContaminant(1).Enabled = True
                'frmListcontaminantScreen2!mnuOptionsManipulateContaminant(3).Enabled = True
                'frmListcontaminantScreen2!mnuOptionsManipulateContaminant(4).Enabled = True
                'frmListcontaminantScreen2!mnuOptionsSave.Enabled = True
                'frmListcontaminantScreen2!mnuOptionsView.Enabled = True

                Call SetDesignContaminantEnabledScreen2(Scr2.NumChemical)
             End If

             'Calculated Flows and Loadings
             Call GetFlowsAndLoadingsScreen2

             'KLaSafetyFactor
             Scr2.KLaSafetyFactor.value = scr1.KLaSafetyFactor.value
             Scr2.KLaSafetyFactor.UserInput = True
             frmPTADScreen2!txtDesignConcentrationValue(1).Text = Format$(Scr2.KLaSafetyFactor.value, GetTheFormat(Scr2.KLaSafetyFactor.value))

             'Calculate properties of the contaminant in the
             'combo box now.

             Scr2.DesignContaminant = scr1.DesignContaminant
             frmPTADScreen2!cboSelectCompo.ListIndex = frmPTADScreen1!cboSelectCompo.ListIndex
             frmPTADScreen2.Caption = "Packed Tower Aeration - Rating Mode (untitled.rat)"

             frmPTADScreen1.Hide
             ShownScreen1Previously = True
             frmPTADScreen2.Show
       
          Else
             frmPTADScreen1.Hide
             frmPTADScreen2.Show
             Filename = "TheDefaultCaseScreen2"
             Call StartScreen2DefaultCase
          End If

       Case 2   'New

       Case 3   'Open
          If HaveValue(scr1.TowerVolume.value) Then
             If (screen1_savechanges()) Then Exit Sub
          End If

          ''''ChDrive SaveAndLoadPath
          ''''ChDir SaveAndLoadPath
          Call ChangeDir_Main
          Call loadscreen1("")
          
          '''''Add this file to the last-few-files list if necessary.
          ''''Call LastFewFiles_MoveFilenameToTop(Filename)
          
          SaveAndLoadPath = CurDir$
          ''''ChDir App.Path
          ''''ChDrive App.Path
          Call ChangeDir_Main

       Case 4   'Save
          ''''ChDrive SaveAndLoadPath
          ''''ChDir SaveAndLoadPath
          Call ChangeDir_Main
          Call SaveScreen1
          
          'Add this file to the last-few-files list if necessary.
          Call LastFewFiles_MoveFilenameToTop(Filename)
          
          SaveAndLoadPath = CurDir$
          ''''ChDir App.Path
          ''''ChDrive App.Path
          Call ChangeDir_Main

       Case 5   'Save As
          ''''ChDrive SaveAndLoadPath
          ''''ChDir SaveAndLoadPath
          Call ChangeDir_Main
          If Right$(frmPTADScreen1.Caption, 14) <> "(untitled.des)" Then
            Call savefilescreen1(Filename)
          End If
          Call SaveScreen1
          
          'Add this file to the last-few-files list if necessary.
          Call LastFewFiles_MoveFilenameToTop(Filename)
          
          SaveAndLoadPath = CurDir$
          ''''ChDir App.Path
          ''''ChDrive App.Path
          Call ChangeDir_Main

       Case 7   'Print
          
          
       Case 8   'Select Printer
            On Error GoTo PrinterError
            CommonDialog1.ShowPrinter
            'CMDialog1.flags = PD_PRINTSETUP
            'CMDialog1.Action = 5
            GoTo ExitSelectPrint
PrinterError:
            Resume ExitSelectPrint:

ExitSelectPrint:

        Case 10   'Return To Main Menu
            If HaveValue(scr1.TowerVolume.value) Then
               If (screen1_savechanges()) Then Exit Sub
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
            Unload frmPTADScreen2
            
            frmMainMenu.Show
        

        Case 200   'Exit
            'Give user option to save design mode before Exiting
            If HaveValue(scr1.TowerVolume.value) Then
               If (screen1_savechanges()) Then Exit Sub
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
            Unload frmPTADScreen2
            frmPTADScreen1_Okay_To_Unload = True
            Unload frmMainMenu
            End
            
    End Select
    
    If ((Index >= 191) And (Index <= 194)) Then
      'Handle File|Open of a file here.
      On Error GoTo notfound
      ''''ChDrive SaveAndLoadPath
      ''''ChDir SaveAndLoadPath
      Call ChangeDir_Main
      If (Dir(Current_LastFewFilesRec.FileNames(Index - 190)) = "") Then
        Beep
        MsgBox "That file has been moved or deleted.", MB_ICONEXCLAMATION, Application_Name
      Else
        Call loadscreen1(Current_LastFewFilesRec.FileNames(Index - 190))
        'Add this file to the last-few-files list if necessary.
        Call LastFewFiles_MoveFilenameToTop(Filename)
        SaveAndLoadPath = CurDir$
      End If
      ''''ChDir App.Path
      ''''ChDrive App.Path
      Call ChangeDir_Main
    End If

    
    Screen.MousePointer = 0   'Arrow
Exit Sub

notfound:
MsgBox "The file was not found"
Resume Next

End Sub

Private Sub mnuFilePrint_Click(Index As Integer)
    
    Select Case Index
       Case 0   'Print to printer
          Call PrintPTADScreen1
       Case 1   'Print to file
          Call PrintPTADScreen1ToFile
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


Private Sub mnuOptions_Click(Index As Integer)

          
Select Case Index
  Case 0: Call screen1_results

  Case 5: '----- View Mass Transfer Parameters
    Call cmdMassTransferLabel_Click

End Select


End Sub

Private Sub mnuotheritem_Click()
'frmabout2.Show 1
End Sub

Private Sub mnuPopContaminant_Click(Index As Integer)
    'If Index = 1 Then
    '   frmListContaminant.Show 1
    'End If
End Sub

Private Sub mnuPower_Click(Index As Integer)
    Dim CalculatedPower As Integer

    Select Case Index   'Power Calculation

       Case 0
          Call SetPowerPTADScreen1(CalculatedPower)
          If CalculatedPower Then
             frmPower.Left = Screen.Width / 2 - frmPower.Width / 2
             frmPower.Top = Screen.Height / 2 - frmPower.Height / 2
             frmPower.Show 1
          End If
    End Select
End Sub

Private Sub mnuUnits_Click(Index As Integer)

  Select Case Index
    Case 0        'SI
      Call LabelsPTADScreen1(UNITSTYPE_SI)
    Case 1        'English
      Call LabelsPTADScreen1(UNITSTYPE_ENGLISH)
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
'temp kill
Exit Sub
'''''''''''''
Dim PointStruct As PointType
 Static PrevioushWnd%
 Dim CurrenthWnd As Integer, TipText As String

If GetActiveWindow() = Me.hwnd Then
  Call GetCursorPos(PointStruct)
  CurrenthWnd% = WindowFromPoint(PointStruct.Y, PointStruct.x)
  If CurrenthWnd <> PrevioushWnd% Then
    PrevioushWnd% = CurrenthWnd
    'HelpTipTimer.Interval = 1
    'Select Case CurrenthWnd%
    '  Case txtOperatingPressure.hWnd    '<---- Here for the text box txtOperatingPressure
    '   StatusMessagePanel.Caption = " Input Operating " & lblOperatingPressure.Caption
    '  Case txtOperatingTemperature.hWnd   '<---- Here for the text box txtOperatingTemperature
    '   StatusMessagePanel.Caption = " Input Operating " & lblOperatingTemperature.Caption
    '  Case lblDisplayAirWaterProperties.hWnd
    '   StatusMessagePanel.Caption = " Specify water density, viscosity, and surface tension; and air density and viscosity"
    'End Select
    ShowHelpTip TipText$
    If Len(TipText$) = 0 Then
      'HelpTipTimer.Interval = 500 'Milliseconds
    End If
  End If
End If

End Sub

Private Sub txtFlowsLoadings_Change(Index As Integer)

'  If (Index = 0) Or (Index = 4) Or (Index = 5) Then
'    Call txtFlowsUnits_Click(Index)
'  End If

End Sub

Private Sub txtFlowsLoadings_GotFocus(Index As Integer)
    
  Call GotFocus_Handle(Me, txtFlowsLoadings(Index), Temp_Text)

End Sub

Private Sub txtFlowsLoadings_KeyPress(Index As Integer, KeyAscii As Integer)
  Call TextBoxNumber_KeyPress(KeyAscii)
End Sub

Private Sub txtFlowsLoadings_LostFocus(Index As Integer)
Dim NewVal As Double
Dim IsNew As Integer
    Dim flag_ok As Integer

   If (LostFocus_IsEvil(Me, txtFlowsLoadings(Index))) Then
     Exit Sub
   End If
   
   flag_ok = True

  IsNew = False
  
  Select Case Index
    Case 0        'Water Flow Rate.
      If (Unitted_LostFocus(UNITS_FLOW, txtFlowsLoadings(0), txtFlowsUnits(0), NewVal, Temp_Text)) Then
        IsNew = True
        scr1.WaterFlowRate.ValChanged = True
        scr1.WaterFlowRate.UserInput = True
        scr1.WaterFlowRate.value = NewVal
      End If

    Case 2        'Multiple of Minimum Air To Water Ratio
      If (NoUnits_LostFocus(txtFlowsLoadings(2), NewVal, Temp_Text)) Then
        IsNew = True
        scr1.MultipleOfMinimumAirToWaterRatio.value = NewVal
        scr1.MultipleOfMinimumAirToWaterRatio.ValChanged = True
        scr1.MultipleOfMinimumAirToWaterRatio.UserInput = True
        Call SpecifiedVQminMultiple
      End If
      
    Case 3        'Air To Water Ratio
      If (NoUnits_LostFocus(txtFlowsLoadings(3), NewVal, Temp_Text)) Then
        IsNew = True
        scr1.AirToWaterRatio.value = NewVal
        scr1.AirToWaterRatio.ValChanged = True
        scr1.AirToWaterRatio.UserInput = True
        Call SpecifiedAirToWaterRatio
      End If
        
    Case 4        'Air Flow Rate
      If (Unitted_LostFocus(UNITS_FLOW, txtFlowsLoadings(4), txtFlowsUnits(4), NewVal, Temp_Text)) Then
        IsNew = True
        scr1.AirFlowRate.value = NewVal
        scr1.AirFlowRate.ValChanged = True
        scr1.AirFlowRate.UserInput = True
        Call SpecifiedAirFlowRate
      End If

    Case 5        'Air Pressure Drop
      If (Unitted_LostFocus(UNITS_PRESSUREPERLENGTH, txtFlowsLoadings(5), txtFlowsUnits(5), NewVal, Temp_Text)) Then
        IsNew = True
        scr1.AirPressureDrop.value = NewVal
        scr1.AirPressureDrop.ValChanged = True
        scr1.AirPressureDrop.UserInput = True
      End If
  
  End Select

  If (IsNew) Then
    If (Index = 0) Then Call GetVQmultVQAndAirFlowRate
    Call GetLoadings
    Call GetTowerAreaAndDiameter
    Call GetOndaMassTransferCoefficient
    Call GetDesignKLaOrKLaSafetyFactor
    Call GetTowerHeightAndVolume
  End If
  Call LostFocus_Handle(Me, txtFlowsLoadings(Index), flag_ok)


End Sub

Private Sub txtFlowsUnits_Click(Index As Integer)

  Select Case Index
    Case 0        'Water Flow Rate
      Call Unitted_UnitChange(UNITS_FLOW, scr1.WaterFlowRate.value, txtFlowsUnits(0), txtFlowsLoadings(0))

    Case 4        'Air Flow Rate
      Call Unitted_UnitChange(UNITS_FLOW, scr1.AirFlowRate.value, txtFlowsUnits(4), txtFlowsLoadings(4))
    
    Case 5        'Air Pressure Drop
      Call Unitted_UnitChange(UNITS_PRESSUREPERLENGTH, scr1.AirPressureDrop.value, txtFlowsUnits(5), txtFlowsLoadings(5))
  End Select

End Sub

Private Sub txtFlowsUnits_GotFocus(Index As Integer)
    
  Call GotFocus_Handle(Me, txtFlowsUnits(Index), Temp_Text)

End Sub

Private Sub txtFlowsUnits_KeyPress(Index As Integer, KeyAscii As Integer)
  Call Combobox_KeyPress(KeyAscii)
End Sub

Private Sub txtFlowsUnits_LostFocus(Index As Integer)
    Dim flag_ok As Integer

   If (LostFocus_IsEvil(Me, txtFlowsUnits(Index))) Then
     Exit Sub
   End If
   
   flag_ok = True
  Call LostFocus_Handle(Me, txtFlowsUnits(Index), flag_ok)

 
End Sub

Private Sub txtMassTransfer_Change(Index As Integer)

'  Call UnitsMassTransfer_Click(2)

End Sub

Private Sub txtMassTransfer_GotFocus(Index As Integer)
    
  Call GotFocus_Handle(Me, txtMassTransfer(Index), Temp_Text)
End Sub

Private Sub txtMassTransfer_KeyPress(Index As Integer, KeyAscii As Integer)
  Call TextBoxNumber_KeyPress(KeyAscii)
End Sub

Private Sub txtMassTransfer_LostFocus(Index As Integer)
Dim NewVal As Double
Dim IsNew As Integer
    Dim flag_ok As Integer

   If (LostFocus_IsEvil(Me, txtMassTransfer(Index))) Then
     Exit Sub
   End If
   
   flag_ok = True

  IsNew = False
  
  Select Case Index
    Case 1        'KLa Safety Factor
      If (NoUnits_LostFocus(txtMassTransfer(1), NewVal, Temp_Text)) Then
        IsNew = True
        scr1.KLaSafetyFactor.value = NewVal
        scr1.KLaSafetyFactor.ValChanged = True
        scr1.KLaSafetyFactor.UserInput = True
        Call SpecifiedKLaSafetyFactor
      End If
    
    Case 2        'Design KLa
      If (Unitted_LostFocus(UNITS_INVERSETIME, txtMassTransfer(2), UnitsMassTransfer(2), NewVal, Temp_Text)) Then
        IsNew = True
        scr1.DesignMassTransferCoefficient.value = NewVal
        scr1.DesignMassTransferCoefficient.ValChanged = True
        scr1.DesignMassTransferCoefficient.UserInput = True
        Call SpecifiedDesignMassTransferCoefficient
      End If
  
  End Select

  If (IsNew) Then
    Call GetTowerHeightAndVolume
  End If
  Call LostFocus_Handle(Me, txtMassTransfer(Index), flag_ok)


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

  If (Unitted_LostFocus(UNITS_PRESSURE, txtOperatingPressure, txtPUnits, NewVal, Temp_Text)) Then
    scr1.OperatingPressure.ValChanged = True
    scr1.OperatingPressure.UserInput = True
    'Note: standard P units are Pa, but
    'OperatingPressure is stored as kPa.
    scr1.OperatingPressure.value = NewVal * 1# / 101325#

    If (HaveValue(scr1.OperatingPressure.value) And HaveValue(scr1.operatingtemperature.value)) Then
      Call CalculateAirWaterProperties
      If (scr1.NumChemical > 0) Then
        'Update Variables on Screen
        Call GetVQmultVQAndAirFlowRate
        Call GetLoadings
        Call GetTowerAreaAndDiameter
        Call GetOndaMassTransferCoefficient
        Call GetDesignKLaOrKLaSafetyFactor
        Call GetTowerHeightAndVolume
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

  If (Unitted_LostFocus(UNITS_TEMPERATURE, txtOperatingTemperature, txtTUnits, NewVal, Temp_Text)) Then
    scr1.operatingtemperature.ValChanged = True
    scr1.operatingtemperature.UserInput = True
    scr1.operatingtemperature.value = NewVal

    If (HaveValue(scr1.OperatingPressure.value) And HaveValue(scr1.operatingtemperature.value)) Then
      Call CalculateAirWaterProperties

      If scr1.NumChemical > 0 Then
        'Update Variables on Screen
        Call GetVQmultVQAndAirFlowRate
        Call GetLoadings
        Call GetTowerAreaAndDiameter
        Call GetOndaMassTransferCoefficient
        Call GetDesignKLaOrKLaSafetyFactor
        Call GetTowerHeightAndVolume
      End If
    End If
  End If
  Call LostFocus_Handle(Me, txtOperatingTemperature, flag_ok)

    
End Sub

Private Sub txtPUnits_Click()
Dim Dummy As Double

  'Note: Standard P units are Pa, but OperatingPressure
  'is stored internally in kPa units.
  Dummy = scr1.OperatingPressure.value * 101325#
  Call Unitted_UnitChange(UNITS_PRESSURE, Dummy, txtPUnits, txtOperatingPressure)

End Sub

Private Sub txtPUnits_KeyPress(KeyAscii As Integer)
  Call Combobox_KeyPress(KeyAscii)
End Sub

Private Sub txtTUnits_Click()
  
  Call Unitted_UnitChange(UNITS_TEMPERATURE, scr1.operatingtemperature.value, txtTUnits, txtOperatingTemperature)

End Sub

Private Sub txtTUnits_KeyPress(KeyAscii As Integer)
  Call Combobox_KeyPress(KeyAscii)
End Sub

Private Sub UnitsMassTransfer_Click(Index As Integer)

  Select Case Index
    Case 0        'Onda KLa
      Call Unitted_UnitChange(UNITS_INVERSETIME, scr1.Onda.OverallMassTransferCoefficient, UnitsMassTransfer(0), lblMassTransfer(0))
  
    Case 2        'Design KLa
      Call Unitted_UnitChange(UNITS_INVERSETIME, scr1.DesignMassTransferCoefficient.value, UnitsMassTransfer(2), txtMassTransfer(2))

  End Select
      
End Sub

Private Sub UnitsMassTransfer_KeyPress(Index As Integer, KeyAscii As Integer)
  Call Combobox_KeyPress(KeyAscii)
End Sub


