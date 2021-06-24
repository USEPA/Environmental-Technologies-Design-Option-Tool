VERSION 5.00
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "THREED32.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form contam_prop_form 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "StEPP - Software to Estimate Physical Properties"
   ClientHeight    =   6795
   ClientLeft      =   585
   ClientTop       =   2625
   ClientWidth     =   9480
   Icon            =   "contam_prop.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   ScaleHeight     =   6795
   ScaleWidth      =   9480
   Begin Threed.SSFrame SSFrame1 
      Height          =   615
      Left            =   840
      TabIndex        =   53
      Top             =   6060
      Visible         =   0   'False
      Width           =   4545
      _Version        =   65536
      _ExtentX        =   8017
      _ExtentY        =   1085
      _StockProps     =   14
      Caption         =   "Invisible"
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
      Begin MSComDlg.CommonDialog CommonDialog1 
         Left            =   3060
         Top             =   120
         _ExtentX        =   847
         _ExtentY        =   847
         _Version        =   393216
      End
      Begin VB.Data Data1_OLD 
         Caption         =   "Data1_OLD"
         Connect         =   "Access"
         DatabaseName    =   ""
         DefaultCursorType=   0  'DefaultCursor
         DefaultType     =   2  'UseODBC
         Exclusive       =   0   'False
         Height          =   360
         Left            =   900
         Options         =   0
         ReadOnly        =   0   'False
         RecordsetType   =   1  'Dynaset
         RecordSource    =   ""
         Top             =   180
         Visible         =   0   'False
         Width           =   1995
      End
      Begin MSComDlg.CommonDialog CommonDialog2 
         Left            =   3540
         Top             =   120
         _ExtentX        =   847
         _ExtentY        =   847
         _Version        =   393216
      End
   End
   Begin Threed.SSFrame fraOperatingConditions 
      Height          =   975
      Left            =   60
      TabIndex        =   0
      Top             =   60
      Width           =   4575
      _Version        =   65536
      _ExtentX        =   8064
      _ExtentY        =   1714
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
         Left            =   3240
         TabIndex        =   6
         Top             =   600
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
         Left            =   3240
         TabIndex        =   5
         Top             =   240
         Width           =   1215
      End
      Begin VB.Label lblOperatingConditions 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
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
         Index           =   1
         Left            =   150
         TabIndex        =   8
         Top             =   600
         Width           =   3015
      End
      Begin VB.Label lblOperatingConditions 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
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
         Index           =   0
         Left            =   120
         TabIndex        =   7
         Top             =   240
         Width           =   3015
      End
   End
   Begin Threed.SSFrame fraAvailableContaminants 
      Height          =   2355
      Left            =   60
      TabIndex        =   1
      Top             =   1110
      Width           =   4575
      _Version        =   65536
      _ExtentX        =   8064
      _ExtentY        =   4149
      _StockProps     =   14
      Caption         =   "Available Contaminants:"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Begin VB.ListBox contam_combo 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1230
         Left            =   120
         TabIndex        =   13
         Top             =   270
         Width           =   4335
      End
      Begin VB.CommandButton cmdSearch 
         Appearance      =   0  'Flat
         Caption         =   "Find Next Occurrence"
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
         Index           =   1
         Left            =   2280
         TabIndex        =   12
         Top             =   1590
         Width           =   2175
      End
      Begin VB.CommandButton cmdSearch 
         Appearance      =   0  'Flat
         Caption         =   "Find"
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
         Index           =   0
         Left            =   120
         TabIndex        =   11
         Top             =   1590
         Width           =   2175
      End
      Begin VB.CommandButton cmdSelectContaminant 
         Appearance      =   0  'Flat
         Caption         =   "Select Current Contaminant"
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
         Left            =   1320
         TabIndex        =   10
         Top             =   1890
         Width           =   3135
      End
      Begin VB.CommandButton cmdSynonyms 
         Appearance      =   0  'Flat
         Caption         =   "Synonyms"
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
         TabIndex        =   9
         Top             =   1890
         Width           =   1215
      End
   End
   Begin Threed.SSFrame fraSelectedContaminants 
      Height          =   2475
      Left            =   60
      TabIndex        =   2
      Top             =   3540
      Width           =   4575
      _Version        =   65536
      _ExtentX        =   8070
      _ExtentY        =   4366
      _StockProps     =   14
      Caption         =   "Selected Contaminants:"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Begin VB.CommandButton cmdUnselectContaminant 
         Appearance      =   0  'Flat
         Caption         =   "Unselect Current Contaminant"
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
         TabIndex        =   15
         Top             =   1980
         Width           =   4335
      End
      Begin VB.ComboBox cboSelectContaminant 
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
         Height          =   1350
         Left            =   120
         Style           =   1  'Simple Combo
         TabIndex        =   14
         Top             =   360
         Width           =   4335
      End
   End
   Begin Threed.SSFrame fraContaminantProperties 
      Height          =   4092
      Left            =   4770
      TabIndex        =   3
      Top             =   60
      Width           =   4572
      _Version        =   65536
      _ExtentX        =   8064
      _ExtentY        =   7218
      _StockProps     =   14
      Caption         =   "Properties of the Contaminant:"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Begin VB.Label lblSelectedContaminant 
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
         Left            =   240
         TabIndex        =   42
         Top             =   420
         Width           =   4095
      End
      Begin VB.Label lblContaminantPropertiesLabel 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         Caption         =   "Gas Diffusivity"
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
         Index           =   12
         Left            =   120
         TabIndex        =   41
         Top             =   3720
         Width           =   3000
      End
      Begin VB.Label lblContaminantPropertiesLabel 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         Caption         =   "Liquid Diffusivity"
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
         Index           =   11
         Left            =   120
         TabIndex        =   40
         Top             =   3480
         Width           =   3000
      End
      Begin VB.Label lblContaminantPropertiesLabel 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         Caption         =   "log Octanol Water Part. Coeff."
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
         Index           =   10
         Left            =   120
         TabIndex        =   39
         Top             =   3240
         Width           =   3000
      End
      Begin VB.Label lblContaminantPropertiesLabel 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         Caption         =   "Aqueous Solubility"
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
         Index           =   9
         Left            =   120
         TabIndex        =   38
         Top             =   3000
         Width           =   3000
      End
      Begin VB.Label lblContaminantPropertiesLabel 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         Caption         =   "Refractive Index"
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
         Index           =   8
         Left            =   120
         TabIndex        =   37
         Top             =   2760
         Width           =   3000
      End
      Begin VB.Label lblContaminantPropertiesLabel 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         Caption         =   "Molar Volume @ NBP"
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
         TabIndex        =   36
         Top             =   2520
         Width           =   3000
      End
      Begin VB.Label lblContaminantPropertiesLabel 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         Caption         =   "Molar Volume @ Op.T"
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
         TabIndex        =   35
         Top             =   2280
         Width           =   3000
      End
      Begin VB.Label lblContaminantPropertiesLabel 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         Caption         =   "Liquid Density"
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
         TabIndex        =   34
         Top             =   2040
         Width           =   3000
      End
      Begin VB.Label lblContaminantPropertiesLabel 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         Caption         =   "Normal Boiling Point"
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
         TabIndex        =   33
         Top             =   1800
         Width           =   3000
      End
      Begin VB.Label lblContaminantPropertiesLabel 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         Caption         =   "Molecular Weight"
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
         TabIndex        =   32
         Top             =   1560
         Width           =   3000
      End
      Begin VB.Label lblContaminantPropertiesLabel 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         Caption         =   "Henry's Constant"
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
         TabIndex        =   31
         Top             =   1320
         Width           =   3000
      End
      Begin VB.Label lblContaminantPropertiesLabel 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         Caption         =   "Infinite Dilution Activity Coeff."
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
         TabIndex        =   30
         Top             =   1080
         Width           =   3000
      End
      Begin VB.Label lblContaminantPropertiesLabel 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         Caption         =   "Vapor Pressure"
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
         TabIndex        =   29
         Top             =   840
         Width           =   3000
      End
      Begin VB.Label lblContaminantProperties 
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
         Index           =   12
         Left            =   3240
         TabIndex        =   28
         Top             =   3720
         Width           =   1215
      End
      Begin VB.Label lblContaminantProperties 
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
         Index           =   11
         Left            =   3240
         TabIndex        =   27
         Top             =   3480
         Width           =   1215
      End
      Begin VB.Label lblContaminantProperties 
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
         Index           =   10
         Left            =   3240
         TabIndex        =   26
         Top             =   3240
         Width           =   1215
      End
      Begin VB.Label lblContaminantProperties 
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
         Index           =   9
         Left            =   3240
         TabIndex        =   25
         Top             =   3000
         Width           =   1215
      End
      Begin VB.Label lblContaminantProperties 
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
         Index           =   8
         Left            =   3240
         TabIndex        =   24
         Top             =   2760
         Width           =   1215
      End
      Begin VB.Label lblContaminantProperties 
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
         Left            =   3240
         TabIndex        =   23
         Top             =   2520
         Width           =   1215
      End
      Begin VB.Label lblContaminantProperties 
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
         Left            =   3240
         TabIndex        =   22
         Top             =   2280
         Width           =   1215
      End
      Begin VB.Label lblContaminantProperties 
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
         Left            =   3240
         TabIndex        =   21
         Top             =   2040
         Width           =   1215
      End
      Begin VB.Label lblContaminantProperties 
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
         Left            =   3240
         TabIndex        =   20
         Top             =   1800
         Width           =   1215
      End
      Begin VB.Label lblContaminantProperties 
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
         Left            =   3240
         TabIndex        =   19
         Top             =   1560
         Width           =   1215
      End
      Begin VB.Label lblContaminantProperties 
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
         Left            =   3240
         TabIndex        =   18
         Top             =   1320
         Width           =   1215
      End
      Begin VB.Label lblContaminantProperties 
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
         Left            =   3240
         TabIndex        =   17
         Top             =   1080
         Width           =   1215
      End
      Begin VB.Label lblContaminantProperties 
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
         Left            =   3240
         TabIndex        =   16
         Top             =   840
         Width           =   1215
      End
   End
   Begin Threed.SSFrame fraAirWaterProperties 
      Height          =   1755
      Left            =   4770
      TabIndex        =   4
      Top             =   4260
      Width           =   4575
      _Version        =   65536
      _ExtentX        =   8064
      _ExtentY        =   3090
      _StockProps     =   14
      Caption         =   "Properties of Air and Water:"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Begin VB.Label lblAirWaterProperties 
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
         Left            =   3240
         TabIndex        =   52
         Top             =   1380
         Width           =   1215
      End
      Begin VB.Label lblAirWaterProperties 
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
         Left            =   3240
         TabIndex        =   51
         Top             =   1140
         Width           =   1215
      End
      Begin VB.Label lblAirWaterProperties 
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
         Left            =   3240
         TabIndex        =   50
         Top             =   900
         Width           =   1215
      End
      Begin VB.Label lblAirWaterProperties 
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
         Left            =   3240
         TabIndex        =   49
         Top             =   660
         Width           =   1215
      End
      Begin VB.Label lblAirWaterPropertiesLabel 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         Caption         =   "Air Viscosity"
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
         TabIndex        =   48
         Top             =   1380
         Width           =   3000
      End
      Begin VB.Label lblAirWaterPropertiesLabel 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         Caption         =   "Air Density"
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
         TabIndex        =   47
         Top             =   1140
         Width           =   3000
      End
      Begin VB.Label lblAirWaterPropertiesLabel 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         Caption         =   "Water Surface Tension"
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
         TabIndex        =   46
         Top             =   900
         Width           =   3000
      End
      Begin VB.Label lblAirWaterPropertiesLabel 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         Caption         =   "Water Viscosity"
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
         TabIndex        =   45
         Top             =   660
         Width           =   3000
      End
      Begin VB.Label lblAirWaterProperties 
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
         Left            =   3240
         TabIndex        =   44
         Top             =   420
         Width           =   1215
      End
      Begin VB.Label lblAirWaterPropertiesLabel 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         Caption         =   "Water Density"
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
         TabIndex        =   43
         Top             =   420
         Width           =   3000
      End
   End
   Begin VB.Menu mnuFileMenu 
      Caption         =   "&File"
      Begin VB.Menu mnuFile 
         Caption         =   "&New"
         Index           =   0
      End
      Begin VB.Menu mnuFile 
         Caption         =   "&Open ..."
         Index           =   1
      End
      Begin VB.Menu mnuFile 
         Caption         =   "- (INVISIBLE)"
         Index           =   3
         Visible         =   0   'False
      End
      Begin VB.Menu mnuFile 
         Caption         =   "&Save"
         Index           =   4
      End
      Begin VB.Menu mnuFile 
         Caption         =   "Save &As ..."
         Index           =   5
      End
      Begin VB.Menu mnuFile 
         Caption         =   "-"
         Index           =   6
      End
      Begin VB.Menu mnuFile 
         Caption         =   "&Print ..."
         Index           =   7
      End
      Begin VB.Menu mnuFile 
         Caption         =   "Select Printer ..."
         Index           =   8
      End
      Begin VB.Menu mnuFile 
         Caption         =   "-"
         Index           =   9
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
         Caption         =   "&Standard International (SI)"
         Index           =   0
      End
      Begin VB.Menu mnuUnits 
         Caption         =   "&English"
         Index           =   1
      End
   End
   Begin VB.Menu mnuOptions 
      Caption         =   "&Options"
      Index           =   10
      Begin VB.Menu mnuOptionsItem 
         Caption         =   "Create Export File for AdDesignS/ASAP"
         Index           =   10
      End
      Begin VB.Menu mnuOptionsItem 
         Caption         =   "Copy to Clipboard for AdDesignS/ASAP"
         Index           =   20
      End
   End
   Begin VB.Menu mnuOptionsEtcMenu 
      Caption         =   "&OptionsOld"
      Visible         =   0   'False
      Begin VB.Menu export 
         Caption         =   "Create Export File for AdDesignS/ASAP"
      End
      Begin VB.Menu mnuOptionsEtc 
         Caption         =   "Modify &Hierarchy"
         Enabled         =   0   'False
         Index           =   0
         Visible         =   0   'False
      End
      Begin VB.Menu mnuAboutMenu 
         Caption         =   "&About"
         Visible         =   0   'False
      End
      Begin VB.Menu mnuAbout 
         Caption         =   "&StEPP Version 1.00"
         Index           =   0
         Visible         =   0   'False
      End
      Begin VB.Menu mnuAbout 
         Caption         =   "The &Authors"
         Index           =   1
         Visible         =   0   'False
      End
      Begin VB.Menu mnuAbout 
         Caption         =   "&Programming Support"
         Index           =   2
         Visible         =   0   'False
      End
      Begin VB.Menu mnuAbout 
         Caption         =   "&Obtaining Additional Information"
         Index           =   3
         Visible         =   0   'False
      End
   End
   Begin VB.Menu mnuHelp 
      Caption         =   "&Help"
      Index           =   0
      Begin VB.Menu frmHelpIndex 
         Caption         =   "&Online Help ..."
         Index           =   10
         Visible         =   0   'False
      End
      Begin VB.Menu frmHelpIndex 
         Caption         =   "&Online Manual ..."
         Index           =   20
      End
      Begin VB.Menu frmHelpIndex 
         Caption         =   "Manual Printing Instructions ..."
         Index           =   22
      End
      Begin VB.Menu frmHelpIndex 
         Caption         =   "-"
         Index           =   25
      End
      Begin VB.Menu frmHelpIndex 
         Caption         =   "View Version History ..."
         Index           =   30
      End
      Begin VB.Menu frmHelpIndex 
         Caption         =   "View Disclaimer ..."
         Index           =   40
         Visible         =   0   'False
      End
      Begin VB.Menu frmHelpIndex 
         Caption         =   "Technical Assistance Provided By ..."
         Index           =   50
         Visible         =   0   'False
      End
      Begin VB.Menu frmHelpIndex 
         Caption         =   "-"
         Index           =   98
      End
      Begin VB.Menu frmHelpIndex 
         Caption         =   "&About StEPP ..."
         Index           =   99
      End
   End
End
Attribute VB_Name = "contam_prop_form"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
''''Option Explicit

Dim contam_prop_form_ActivatedYet As Integer





Const contam_prop_form_declarations_end = True


Sub frmMain_Reset_DemoVersionDisablings()
  If (IsThisADemo() = True) Then
    mnuFile(1).Enabled = False
    mnuFile(4).Enabled = False
    mnuFile(5).Enabled = False
    mnuFile(191).Enabled = False
    mnuFile(192).Enabled = False
    mnuFile(193).Enabled = False
    mnuFile(194).Enabled = False
    fraContaminantProperties.ForeColor = QBColor(12)
    fraContaminantProperties.Caption = "* DEMONSTRATION VERSION *"
  End If
End Sub


Private Sub BlankTextBoxesPressure()
    'Blank text boxes that are recalculated when Pressure changes
  
    lblContaminantProperties(12).Caption = ""

End Sub

Private Sub BlankTextBoxesTemp()
    'Blank text boxes that are recalculated when T changes
    Dim i As Integer
  
    lblContaminantProperties(0).Caption = ""
    If phprop.ActivityCoefficient.BinaryInteractionParameterDatabase > 0 Then lblContaminantProperties(1).Caption = ""
    lblContaminantProperties(2).Caption = ""
    lblContaminantProperties(5).Caption = ""
    lblContaminantProperties(6).Caption = ""
    If phprop.AqueousSolubility.BinaryInteractionParameterDatabase > 0 Then lblContaminantProperties(9).Caption = ""
    If phprop.OctWaterPartCoeff.BinaryInteractionParameterDatabase > 0 Then lblContaminantProperties(10).Caption = ""
    lblContaminantProperties(11).Caption = ""
    lblContaminantProperties(12).Caption = ""

    For i = 0 To 4
        lblAirWaterProperties(i).Caption = ""
    Next i

End Sub

Private Sub cboSelectContaminant_Click()
    Dim ContaminantName As String, Ch As String, LastName As String
    Dim hc_database_value As String * 40
    Dim hc_database_temp As String
    Dim hc_string As String
    Dim hc_unifac_value As String * 40
    Dim hc_unifac_temp As String
    Dim SIValue As Double, EnglishValue As Double

    If cboSelectContaminant.ListIndex = -1 Then Exit Sub

    ContaminantName = cboSelectContaminant.Text
    ContaminantName = Right$(ContaminantName, Len(ContaminantName) - 1)
    Ch = Left$(ContaminantName, 1)
    

    While Ch <> " "
      
       ContaminantName = Right$(ContaminantName, Len(ContaminantName) - 1)
       Ch = Left$(ContaminantName, 1)
    Wend
    ContaminantName = Right$(ContaminantName, Len(ContaminantName) - 1)

    lblSelectedContaminant.Caption = ContaminantName
'    If (cboSelectContaminant.ListCount < 2) Then Exit Sub

    If JustLoadedFile Then
       For i = 1 To NUMBER_OF_PROPERTIES
           HaveProperty(i) = phprop.HaveProperty(i)
       Next i
       For i = 1 To NUMBER_OF_PROPERTIES_AVAILABLE
           PROPAVAILABLE(i) = phprop.PROPAVAILABLE(i)
       Next i
       Call InitializeHilights

       If CurrentUnits = SIUnits Then
          txtOperatingPressure.Text = Str$(phprop.OperatingPressure)
          txtOperatingTemperature.Text = Str$(phprop.OperatingTemperature)
       Else
          SIValue = phprop.OperatingPressure
          Call PRESSCNV(EnglishValue, SIValue)
          txtOperatingPressure.Text = Str$(EnglishValue)

          SIValue = phprop.OperatingTemperature
          Call TEMPCNV(EnglishValue, SIValue)
          txtOperatingTemperature.Text = Str$(EnglishValue)
       End If

       Call DisplayAllProperties

       PreviouslySelectedIndex = cboSelectContaminant.ListIndex + 1
       JustLoadedFile = False
       Exit Sub
    End If

    If PropContaminant(cboSelectContaminant.ListIndex + 1).CasNumber <> phprop.CasNumber Then
       For i = 1 To NUMBER_OF_PROPERTIES
           phprop.HaveProperty(i) = HaveProperty(i)
       Next i
       For i = 1 To NUMBER_OF_PROPERTIES_AVAILABLE
           phprop.PROPAVAILABLE(i) = PROPAVAILABLE(i)
       Next i
       If PreviouslySelectedIndex > 0 Then
          PropContaminant(PreviouslySelectedIndex) = phprop
       End If

       phprop = PropContaminant(cboSelectContaminant.ListIndex + 1)
       For i = 1 To NUMBER_OF_PROPERTIES
           HaveProperty(i) = phprop.HaveProperty(i)
       Next i
       For i = 1 To NUMBER_OF_PROPERTIES_AVAILABLE
           PROPAVAILABLE(i) = phprop.PROPAVAILABLE(i)
       Next i
       Call InitializeHilights

       If CurrentUnits = SIUnits Then
          txtOperatingPressure.Text = Str$(phprop.OperatingPressure)
          txtOperatingTemperature.Text = Str$(phprop.OperatingTemperature)
       Else
          SIValue = phprop.OperatingPressure
          Call PRESSCNV(EnglishValue, SIValue)
          txtOperatingPressure.Text = Str$(EnglishValue)

          SIValue = phprop.OperatingTemperature
          Call TEMPCNV(EnglishValue, SIValue)
          txtOperatingTemperature.Text = Str$(EnglishValue)
       End If

       Call DisplayAllProperties

       PreviouslySelectedIndex = cboSelectContaminant.ListIndex + 1
    End If

End Sub

Private Sub cboSelectContaminant_KeyPress(KeyAscii As Integer)
    KeyAscii = 0
End Sub

Private Sub cmdCancel_Click()
    End
End Sub

Private Sub cmdSearch_Click(Index As Integer)
    Dim contam_combo_index As Integer

    Select Case Index
       Case 0
          frmStringS.Show 1
          If Find_String <> "" Then
             Screen.MousePointer = 11  'HourGlass
             Call Search_String(-1)
             Screen.MousePointer = 0   'Arrow
          End If
       Case 1
          If Find_String <> "" Then
             contam_combo_index = contam_combo.ListIndex
             Screen.MousePointer = 11  'HourGlass
             Call Search_String(contam_combo_index)
             Screen.MousePointer = 0   'Arrow
          End If

    End Select

End Sub

Private Sub cmdSelectContaminant_Click()
    Dim i As Integer, J As Integer
    Dim msg$, Response As Integer
Dim strThisChem As String
  If (contam_combo.ListIndex < 0) Then Exit Sub
  If (IsThisADemo() = True) Then
    strThisChem = contam_combo.List(contam_combo.ListIndex)
    If (InStr(strThisChem, " 56235 ") = 0) Then
      Call Demo_ShowError("In this demonstration version, " & _
          "only the chemical CARBON TETRACHLORIDE may be selected.")
      Exit Sub
    End If
  End If
  
''''' RETURNS FALSE IF THE CHEMICAL CAN CONTINUE ON
''''' ALWAYS RETURNS FALSE IF NOT IN DEMOMODE CHECK DEMOMODE.BAS
''''    If (demo_check_chemicals(contam_combo)) Then Exit Sub

    If NumSelectedChemicals = MAXSELECTEDCHEMICALS Then
       msg$ = "The maximum number of contaminants that can be selected at a time in the StEPP program is " & Str$(MAXSELECTEDCHEMICALS) & ".  Therefore, you may not select this chemical unless you Unselect a contaminant you selected previously or begin the program again."
       MsgBox msg$, MB_ICONSTOP, "Too Many Contaminants Selected"
       Exit Sub
    End If

    Screen.MousePointer = 11   'Hourglass

    Update_Fields (contam_combo.ListIndex)

    If NumSelectedChemicals = 0 Then
       contam_prop_form!mnuFile(4).Enabled = True
       contam_prop_form!mnuFile(5).Enabled = True
       contam_prop_form!mnuFile(7).Enabled = True
       contam_prop_form!cmdUnselectContaminant.Enabled = True
       Call frmMain_Reset_DemoVersionDisablings
    End If

    For i = 0 To cboSelectContaminant.ListCount - 1
        If Trim$(cboSelectContaminant.List(i)) = Trim$(contam_combo.List(contam_combo.ListIndex)) Then
           msg$ = "There is already a contaminant named "
           msg$ = msg$ + contam_combo.List(contam_combo.ListIndex) + " selected. "
           msg$ = msg$ + Chr$(13) + Chr$(13)
           msg$ = msg$ + "Do you wish to reinitialize it to default properties by selecting it now?"
           Response = MsgBox(msg$, MB_ICONQUESTION + MB_YESNO, "Contaminant Already Selected")
           If Response = IDYES Then
              If Trim$(contam_combo.List(contam_combo.ListIndex)) <> Trim$(cboSelectContaminant.Text) Then  'If contaminant currently selected is not the one being replaced then update its values before performing calculations
                 For J = 1 To NUMBER_OF_PROPERTIES
                     phprop.HaveProperty(J) = HaveProperty(J)
                 Next J
                 For J = 1 To NUMBER_OF_PROPERTIES_AVAILABLE
                     phprop.PROPAVAILABLE(J) = PROPAVAILABLE(J)
                 Next J
                 PropContaminant(PreviouslySelectedIndex) = phprop
              End If

              cboSelectContaminant.RemoveItem i
              For J = i + 2 To NumSelectedChemicals
                  PropContaminant(J - 1) = PropContaminant(J)
              Next J
              NumSelectedChemicals = NumSelectedChemicals - 1
              Exit For
           Else
              Screen.MousePointer = 0   'Arrow
              Exit Sub
           End If
        End If
    Next i

    cboSelectContaminant.AddItem contam_combo.List(contam_combo.ListIndex)
    lblSelectedContaminant.Caption = Trim$(dbinput.Name)

    'Update the contaminant selected prior to the new one if necessary
    If PreviouslySelectedIndex >= 0 Then
       For J = 1 To NUMBER_OF_PROPERTIES
           phprop.HaveProperty(J) = HaveProperty(J)
       Next J
       For J = 1 To NUMBER_OF_PROPERTIES_AVAILABLE
           phprop.PROPAVAILABLE(J) = PROPAVAILABLE(J)
       Next J
       PropContaminant(PreviouslySelectedIndex) = phprop
    End If

'* initialize binary interaction parameter database choices
    phprop.ActivityCoefficient.BinaryInteractionParameterDatabase = BIP_dbHierarchy.ActivityCoefficient(1)
    phprop.AqueousSolubility.BinaryInteractionParameterDatabase = BIP_dbHierarchy.AqueousSolubility(1)
    phprop.OctWaterPartCoeff.BinaryInteractionParameterDatabase = BIP_dbHierarchy.OctWaterPartCoeff(1)
    For i = 1 To 3
        phprop.ActivityCoefficient.BinaryInteractionParameterDBAvailable(i) = True
        phprop.AqueousSolubility.BinaryInteractionParameterDBAvailable(i) = True
        If i <> 3 Then phprop.OctWaterPartCoeff.BinaryInteractionParameterDBAvailable(i) = True
    Next i
    UserSelectedTheUnifacBIPDBActCoeff = False
    UserSelectedTheUnifacBIPDBAqSol = False
    UserSelectedTheUnifacBIPDBKow = False

'* Set Current Selections to None
    Call InitializeCurrentSelections

    NumSelectedChemicals = NumSelectedChemicals + 1
    Call InitializeHilights
    Call InitializePROPandHAVEAVAILABLEArrays
    Call InitializeUserInputs

    Call BlankAllTextBoxes
    frmWaitForCalculations.Show
    frmWaitForCalculations.Refresh

' THIS IS HERE TO MAKE SURE THAT THE CURRENT DIRECTORY IS WITH THE FORTRAN DLL
' FILES.   THE DIFFERENT *.DAT FILES THAT ARE USED BY THE FORTRAN DLLS MUST BE THERE
'    msg$ = CurDir$
'    ChDrive app.path
'    ChDir app.path + "\dlls"

    Call DoCalculationForThisContaminant

' RETURNING TO WHERE WE WERE BEFORE
'    ChDrive msg$
'    ChDir msg$

    phprop.AqueousSolubility.PreviousBinaryInteractionParameterDB = phprop.AqueousSolubility.BinaryInteractionParameterDatabase
    
    If NumSelectedChemicals > 0 Then cboSelectContaminant.Enabled = True
    frmWaitForCalculations.Hide
    cboSelectContaminant.ListIndex = cboSelectContaminant.ListCount - 1
    cboSelectContaminant.SetFocus
  
    Screen.MousePointer = 0   'Arrow

End Sub

'Private Sub cmdSteppLink_Click(Index As Integer)
'Dim f As Integer
'Dim i As Integer
'Dim this_name As String
'Dim this_cas As String
'Dim temp1 As String
'
'  If (Index = 0) Then
'    If cboSelectContaminant.ListCount = 0 Then
'      MsgBox "No chemicals were selected for transfer!  Choose 'Cancel' if you want to abort.", MB_ICONEXCLAMATION, Application_Name
'      Exit Sub
'    End If
'  End If
'
'  Select Case Index
'    Case 0
'      '---- Output property file
'      Call GetTempFilename(CStr(App.Path), SteppLink_fn_properties)
'      SteppLink_fn_properties = App.Path & "\" & SteppLink_fn_properties
''MsgBox "Outputting to `" & SteppLink_fn_properties & "` ..."
'      f = FreeFile
'      Open SteppLink_fn_properties For Output As #f
'      For i = 0 To cboSelectContaminant.ListCount - 1
'        cboSelectContaminant.ListIndex = i
'        temp1 = LTrim$(cboSelectContaminant.List(cboSelectContaminant.ListIndex))
'        Call parsedargs_getarg(" ", temp1, 1, this_cas)
'        this_name = Trim$(lblSelectedContaminant)
'        Write #f, "Chemical", this_name, this_cas
'        Call SteppLink_OutputProperty(f, 0, "VaporPressure", "Pa")
'        Call SteppLink_OutputProperty(f, 1, "ActivityCoefficient", "-")
'        Call SteppLink_OutputProperty(f, 2, "HenrysConstant", "-")
'        Call SteppLink_OutputProperty(f, 3, "MolecularWeight", "kg/kmol")
'        Call SteppLink_OutputProperty(f, 4, "NormalBoilingPoint", "C")
'        Call SteppLink_OutputProperty(f, 5, "LiquidDensity", "kg/m3")
'        Call SteppLink_OutputProperty(f, 6, "MolarVolumeAtOpT", "m3/kmol")
'        Call SteppLink_OutputProperty(f, 7, "MolarVolumeAtNBP", "m3/kmol")
'        Call SteppLink_OutputProperty(f, 8, "RefractiveIndex", "-")
'        Call SteppLink_OutputProperty(f, 9, "AqueousSolubility", "PPMw")
'        Call SteppLink_OutputProperty(f, 10, "LogKOW", "-")
'        Call SteppLink_OutputProperty(f, 11, "LiquidDiffusivity", "m2/s")
'        Call SteppLink_OutputProperty(f, 12, "GasDiffusivity", "m2/s")
'      Next i
'      Write #f, "END_OF_FILE", "", ""
'      Close #f
'
'      '---- Re-create wait-file to signal Client of successful link
'      f = FreeFile
''MsgBox "Outputting to `" & SteppLink_fn_done_waitfile & "` ..."
'      Open SteppLink_fn_done_waitfile For Output As #f
'      Print #f, "OK"
'      Print #f, SteppLink_fn_properties
'      Close #f
'    Case 1
'      f = FreeFile
''MsgBox "Outputting to `" & SteppLink_fn_done_waitfile & "` ..."
'      Open SteppLink_fn_done_waitfile For Output As #f
'      Print #f, "CANCEL"
'      Close #f
'  End Select
'
'  Select Case Index
'    Case 0        'Use contaminants in {Client}
'      dde_stepplink_status.Text = "complete"
'      SteppLink_Status = STEPPLINK_STATUS_INACTIVE
'    Case 1        'Cancel StEPP-{Client} link
'      dde_stepplink_status.Text = "cancel"
'      SteppLink_Status = STEPPLINK_STATUS_INACTIVE
'  End Select
'
'  '---- Various screen-related B.S. for closing the link
'  cmdSteppLink(0).Visible = False
'  cmdSteppLink(1).Visible = False
'  txtOperatingPressure.Enabled = True
'  txtOperatingTemperature.Enabled = True
'  SteppLink_SpecifiedPressure = ""
'  SteppLink_SpecifiedTemperature = ""
'  If (frmstepinfo.Visible) Then
'    '-- Ensure they don't click "never show again" without clicking OK
'    frmstepinfo.chkdisplay = False
'    Unload frmstepinfo
'  End If
'
'  '---- Minimize StEPP
'  contam_prop_form.WindowState = 1
'
'End Sub

Private Sub cmdSynonyms_Click()
frmAlias.Show 1

End Sub

Private Sub cmdUnselectContaminant_Click()
    
  Call Do_UnselectCurrentContaminant(False)

  If NumSelectedChemicals = 0 Then
     cboSelectContaminant.Enabled = False
     mnuFile(4).Enabled = False
     mnuFile(5).Enabled = False
     mnuFile(7).Enabled = False
     Call frmMain_Reset_DemoVersionDisablings
  End If
  
End Sub

Private Sub Command1_Click()

  SteppLink_SpecifiedPressure = InputBox("Specify forcing pressure.")
  SteppLink_SpecifiedTemperature = InputBox("Specify forcing temperature.")


Call Update_P_and_T_StEPPLink

End Sub

Private Sub Command3D1_Click()
    End
End Sub

Private Sub contam_combo_Click()
'    TempIndex = contam_combo.ListIndex
    'call eliminated during restructuring of database
    'by F. Gobin
'    update_fields (contam_combo.ListIndex)
    
End Sub

Private Sub contam_combo_DblClick()

cmdSelectContaminant_Click

End Sub

Private Sub contam_combo_KeyPress(KeyAscii As Integer)

    If KeyAscii = 13 Then
'       KeyAscii = 0
'       cmdSelectContaminant.SetFocus
    Else
       KeyAscii = 0
    End If

End Sub

'Private Sub dde_msg_Change()
'Dim X As Integer
'Dim response As String
'Dim params As String
'Dim NumParams As Integer
'
'  On Error GoTo err_dde_msg
'
'  If (StrComp(dde_msg.Text, "querynumcomp") = 0) Then
'
'    'NOTE: This is an OLD StEPP Link command, preserved only for backward compatibility
'    'Set dde_numcomp to the number of components currently selected.
'
'    dde_numcomp.Text = Trim$(Str$(cboSelectContaminant.ListCount))
'
'  ElseIf (StrComp(Left$(dde_msg.Text, 14), "querycompname ") = 0) Then
'
'    'NOTE: This is an OLD StEPP Link command, preserved only for backward compatibility
'    'Set dde_thiscompname to the component name requested.
'
'    X = CInt(Trim$(Right$(dde_msg.Text, Len(dde_msg.Text) - 14)))
'    If (X < 0) Or (X > cboSelectContaminant.ListCount - 1) Then
'      GoTo err_dde_msg_OutOfBounds
'    End If
'    dde_thiscompname.Text = cboSelectContaminant.List(X)
'
'  ElseIf (StrComp(Left$(dde_msg.Text, 12), "displaycomp ") = 0) Then
'
'    'NOTE: This is an OLD StEPP Link command, preserved only for backward compatibility
'    'Display the properties of the component requested.
'
'    X = CInt(Trim$(Right$(dde_msg.Text, Len(dde_msg.Text) - 12)))
'    If (X < 0) Or (X > cboSelectContaminant.ListCount - 1) Then
'      GoTo err_dde_msg_OutOfBounds
'    End If
'    cboSelectContaminant.ListIndex = X
'
'  ElseIf (StrComp(Left$(dde_msg.Text, 16), "stepplink_begin ") = 0) Then
'
'    'Begin a StEPP-Client link if one is not currently in progress.
'    'PARAMETERS: stepplink_begin {Stepp_ClientProgram} {fn_done_waitfile} {PRESSURE_in_Pa} {TEMPERATURE_in_C}
'
'    If (SteppLink_Status = STEPPLINK_STATUS_ACTIVE) Then
'      'Do nothing.
'    ElseIf (SteppLink_Status = STEPPLINK_STATUS_INACTIVE) Then
'
'      'Set statuses to ACTIVE.
'      SteppLink_Status = STEPPLINK_STATUS_ACTIVE
'      dde_stepplink_status.Text = "active"
'
'      ''Force removal of all selected contaminants (if any).
'      'Do
'      '  If (cboSelectContaminant.ListCount = 0) Then Exit Do
'      '  cboSelectContaminant.ListIndex = 0
'      '  Call Do_UnselectCurrentContaminant(True)
'      'Loop Until (1 <> 1)
'
'
'      'Parse parameters.
'      params = Trim$(Mid$(dde_msg.Text, 17, Len(dde_msg.Text) - 16))
'      NumParams = ParsedArgs_GetNum(" ", params)
'      If (NumParams >= 1) Then
'        'Find out which client program it is (ASAP or ADSIM).
'        Call parsedargs_getarg(" ", params, 1, SteppLink_ClientProgram)
'        Stepp_ClientProgram = SteppLink_ClientProgram
'      End If
'      If (NumParams >= 2) Then
'        'Set filename of waitfile {fn_done_waitfile}.
'        Call parsedargs_getarg(" ", params, 2, SteppLink_fn_done_waitfile)
'      Else
'        SteppLink_fn_done_waitfile = ""
'      End If
'      If (NumParams >= 3) Then
'        'Set pressure to that specified by the client.
'        Call parsedargs_getarg(" ", params, 3, SteppLink_SpecifiedPressure)
'        txtOperatingPressure = SteppLink_SpecifiedPressure
'        phprop.OperatingPressure = CDbl(SteppLink_SpecifiedPressure)
'        txtOperatingPressure.Enabled = False
'      Else
'        SteppLink_SpecifiedPressure = ""
'      End If
'      If (NumParams >= 4) Then
'        'Set temperature to that specified by the client.
'        Call parsedargs_getarg(" ", params, 4, SteppLink_SpecifiedTemperature)
'        txtOperatingTemperature = SteppLink_SpecifiedTemperature
'        dbinput.OperatingTemperature = CDbl(SteppLink_SpecifiedTemperature)
'         phprop.OperatingTemperature = CDbl(SteppLink_SpecifiedTemperature)
'        txtOperatingTemperature.Enabled = False
'      Else
'        SteppLink_SpecifiedTemperature = ""
'      End If
'
'      '''MsgBox "Specified P & T = " & SteppLink_SpecifiedPressure & ", " & SteppLink_SpecifiedTemperature, MB_ICONEXCLAMATION, "StEPP"
'
'      '----- Tell Client program that the message is confirmed.
'      dde_msg.Text = ""
'
'      '----- Update P & T of all current contaminants if necessary.
'      Call Update_P_and_T_StEPPLink
'
'
''MsgBox "test point A"
'      'Set up OK and Cancel buttons for StEPP Link.
'      cmdSteppLink(0).Caption = "Use these contaminants in " & SteppLink_ClientProgram
'      cmdSteppLink(1).Caption = "Cancel StEPP-" & SteppLink_ClientProgram & " link"
'      cmdSteppLink(0).Visible = True
'      cmdSteppLink(1).Visible = True
'
''MsgBox "test point b"
'      'Convert to SI units if necessary.
'      Call mnuunits_click(0)
'
''MsgBox "test point c"
'      'Restore StEPP to visibility if minimized (or maximized).
'      If (contam_prop_form.WindowState <> 0) Then
'        contam_prop_form.WindowState = 0
'      End If
'
''MsgBox "test point d"
'      'Delete waitfile to signal client that I have recognized their request
'      'MsgBox "About to delete `" & SteppLink_fn_done_waitfile & "` ..."
'      If (SteppLink_fn_done_waitfile <> "") Then
'        If (Dir(SteppLink_fn_done_waitfile) <> "") Then
'          Kill SteppLink_fn_done_waitfile
'        End If
'      End If
'
''MsgBox "test point e"
'      'Display the StEPP Link Instructions if necessary.
'      'response = ini_getsetting(INI_FileName, INI_ProgramType, "has_seen_steppinfo")
'      response = INI_Getsetting("has_seen_steppinfo")
'      If (response <> "1") Then
'        frmstepinfo.Show
'        frmstepinfo.Command1.SetFocus
'      End If
'
''MsgBox "test point f"
'    Else
'      'Do nothing.
'    End If
'
'  ElseIf (StrComp(Left$(dde_msg.Text, 16), "stepplink_cancel") = 0) Then
'
'    'Cancel a StEPP-Client link if one is in progress.
'
'    If (SteppLink_Status = STEPPLINK_STATUS_ACTIVE) Then
'      'Set statuses to INACTIVE.
'      SteppLink_Status = STEPPLINK_STATUS_INACTIVE
'      dde_stepplink_status.Text = "inactive"
'
'      'Turn off StEPP-Client link buttons.
'      cmdSteppLink(0).Visible = False
'      cmdSteppLink(1).Visible = False
'
'    ElseIf (SteppLink_Status = STEPPLINK_STATUS_INACTIVE) Then
'      'Do nothing.
'    Else
'      'Do nothing.
'    End If
'
'  End If
'
'  dde_msg.Text = ""
'
'
'
'exit_dde_msg:
'  On Error GoTo 0
'  Exit Sub
'
'err_dde_msg_OutOfBounds:
'  MsgBox "Error in DDE communication--index out of bounds.", MB_ICONSTOP, "StEPP"
'  GoTo exit_dde_msg
'
'err_dde_msg:
'  MsgBox "Error in DDE communication.", MB_ICONSTOP, "StEPP"
'  Resume exit_dde_msg
'
'End Sub

Private Sub Do_UnselectCurrentContaminant(ForceIt As Integer)
Dim NumRemovedChemical As Integer
Dim i As Integer

  'NOTE: The variable ForceIt is not currently used.
  'If any code is ever added to prompt the user if
  'they are "sure" they want to unselect the current
  'contaminant, you MUST NOT put up a prompt if
  'ForceIt is set to True.  -ejo, 6/15/96
  
  Screen.MousePointer = 11   'Hourglass

  NumRemovedChemical = cboSelectContaminant.ListIndex + 1
  For i = (NumRemovedChemical + 1) To NumSelectedChemicals
    PropContaminant(i - 1) = PropContaminant(i)
  Next i
  NumSelectedChemicals = NumSelectedChemicals - 1

  cboSelectContaminant.RemoveItem cboSelectContaminant.ListIndex

  If NumSelectedChemicals > 0 Then
    PreviouslySelectedIndex = -1
    cboSelectContaminant.ListIndex = 0
  Else
    cmdUnselectContaminant.Enabled = False
    Call BlankAllTextBoxes
  End If

  Screen.MousePointer = 0   'arrow

End Sub

Private Sub ExportFileGeneration_GetText(cliptext As String)
Dim f As Integer
Dim i As Integer
Dim index0 As Integer
Dim this_name As String
Dim this_cas As String
Dim temp1 As String
Dim ExportFile As String
'Dim cliptext As String
Dim vb3CrLf As String
      
  'get chemical selected from "selected contaminants and see if valid
  index0 = cboSelectContaminant.ListCount
  If (index0 = 0) Then
    MsgBox "No chemicals were selected for export!", MB_ICONEXCLAMATION, "StEPP"
    Exit Sub
  End If
  
  ''******** add box to choose file type
  'On Error Resume Next
  'contam_prop_form!CMDialog1.Filename = ""
  'contam_prop_form!CMDialog1.DefaultExt = "exp"
  'contam_prop_form!CMDialog1.Filter = "StEPP export Files (*.exp)|*.exp"
  'contam_prop_form!CMDialog1.DialogTitle = "Save StEPP export File"
  'contam_prop_form!CMDialog1.Flags = OFN_OVERWRITEPROMPT Or OFN_PATHMUSTEXIST
  'contam_prop_form!CMDialog1.CancelError = True
  'contam_prop_form!CMDialog1.Action = 2
  'ExportFile = contam_prop_form!CMDialog1.Filename
  'contam_prop_form!CMDialog1.Filename = ""
  'If Err = 32755 Then   'Cancel selected by user
  '  Exit Sub
  'End If

  'CONVERT TO SI UNITS IF REQUIRED.
  If (Not mnuUnits(0).CHECKED) Then
    'mnuUnits(1).Checked = False
    'mnuUnits(0).Checked = True
    Call mnuunits_click(0)
    DoEvents
  End If

  'SET UP TEXT TO BE COPIED TO CLIPBOARD.
  vb3CrLf = Chr$(13) & Chr$(10)
  cliptext = "1234567890:START_OF_STEPP_CLIPBOARD_EXPORT"
  cliptext = cliptext & vb3CrLf
  Call SteppLink_AddItemToClipboard("Operating Pressure, Pa", Trim$(Str$(phprop.OperatingPressure)), cliptext)
  Call SteppLink_AddItemToClipboard("Operating Pressure, degC", Trim$(Str$(phprop.OperatingTemperature)), cliptext)
  Call SteppLink_AddItemToClipboard("Number Components On Clipboard", Trim$(Str$(cboSelectContaminant.ListCount)), cliptext)
  For i = 0 To cboSelectContaminant.ListCount - 1
    cboSelectContaminant.ListIndex = i
    temp1 = LTrim$(cboSelectContaminant.List(cboSelectContaminant.ListIndex))
    Call parsedargs_getarg(" ", temp1, 1, this_cas)
    this_name = Trim$(lblSelectedContaminant)
    Call SteppLink_AddItemToClipboard("Chemical Name, -", this_name, cliptext)
    Call SteppLink_AddItemToClipboard("Chemical CAS, -", this_cas, cliptext)
    Call SteppLink_AddItemToClipboard("VaporPressure, Pa", SteppLink_GetPropertyForOutput(0), cliptext)
    Call SteppLink_AddItemToClipboard("ActivityCoefficient, -", SteppLink_GetPropertyForOutput(1), cliptext)
    Call SteppLink_AddItemToClipboard("HenrysConstant, -", SteppLink_GetPropertyForOutput(2), cliptext)
    Call SteppLink_AddItemToClipboard("MolecularWeight, kg/kmol", SteppLink_GetPropertyForOutput(3), cliptext)
    Call SteppLink_AddItemToClipboard("NormalBoilingPoint, C", SteppLink_GetPropertyForOutput(4), cliptext)
    Call SteppLink_AddItemToClipboard("LiquidDensity, kg/m3", SteppLink_GetPropertyForOutput(5), cliptext)
    Call SteppLink_AddItemToClipboard("MolarVolumeAtOpT, m3/kmol", SteppLink_GetPropertyForOutput(6), cliptext)
    Call SteppLink_AddItemToClipboard("MolarVolumeAtNBP, m3/kmol", SteppLink_GetPropertyForOutput(7), cliptext)
    Call SteppLink_AddItemToClipboard("RefractiveIndex, -", SteppLink_GetPropertyForOutput(8), cliptext)
    Call SteppLink_AddItemToClipboard("AqueousSolubility, PPMw", SteppLink_GetPropertyForOutput(9), cliptext)
    Call SteppLink_AddItemToClipboard("LogKOW, -", SteppLink_GetPropertyForOutput(10), cliptext)
    Call SteppLink_AddItemToClipboard("LiquidDiffusivity, m2/s", SteppLink_GetPropertyForOutput(11), cliptext)
    Call SteppLink_AddItemToClipboard("GasDiffusivity, m2/s", SteppLink_GetPropertyForOutput(12), cliptext)
  Next i
  cliptext = cliptext & "1234567890:END_OF_STEPP_CLIPBOARD_EXPORT"
  cliptext = cliptext & vb3CrLf

End Sub

Private Sub Form_Activate()
Dim Response As String

  If (Not contam_prop_form_ActivatedYet) Then
    Call LastFewFiles_InitializeList(LASTFEW_WHICHAPP_STEPP, LASTFEW_STEPP_contam_prop_form)
    contam_prop_form_ActivatedYet = True
  End If

End Sub

Private Sub Form_Load()

  Me.Top = (Screen.Height - Me.Height) / 2
  Me.Left = (Screen.Width - Me.Width) / 2
  
  'The following line MUST BE HERE, and NOT in setup_form,
  'otherwise it will be reset every time there is a
  'file operation, which is stupid.
  SteppLink_Status = STEPPLINK_STATUS_INACTIVE

  'Set up the main form.
  Call setup_form
  '
  ' DEMO SETTINGS.
  '
  Call frmMain_Reset_DemoVersionDisablings
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
Dim Response As Integer
Dim msg As String


If (SteppLink_Status = STEPPLINK_STATUS_ACTIVE) Then
    msg = "In order to leave StEPP please click either the 'Cancel StEPP link' or"
    msg = msg + " the 'Use these contaminants' button below."
    MsgBox msg, MB_ICONEXCLAMATION, "StEPP"
   Cancel = True
    Screen.MousePointer = 0
    Exit Sub
End If

            Response = MsgBox("Save current data?", MB_ICONQUESTION + MB_YESNOCANCEL, "StEPP")
            If Response = IDCANCEL Then
              Screen.MousePointer = 0
              Cancel = True
              Exit Sub
            End If
            If Response = IDYES Then
              'ChDrive SaveAndLoadPath
              'ChDir SaveAndLoadPath
              Call ChangeDir_Main
              Call SaveStEPPDesign
              'Add this file to the last-few-files list if necessary.
              Call LastFewFiles_MoveFilenameToTop(FileName$)
              SaveAndLoadPath = CurDir$
              'ChDrive steppPath
              'ChDir steppPath
              Call ChangeDir_Main
            End If

End Sub

Private Sub Form_Unload(Cancel As Integer)

End
End Sub

Private Sub frmHelpIndex_Click(Index As Integer)
Dim msg As String
Dim fn_This As String
  Select Case Index
'    Case 10:      'CONTENTS.
'      ' Msg = "In the near future, an online help system similar "
'      ' Msg = Msg + "in format to online help systems for other Windows "
'      ' Msg = Msg + "applications will be implemented here." & Chr$(13) & Chr$(13)
'      ' Msg = Msg + "It will explain the correlations and parameter "
'      ' Msg = Msg + "estimation techniques used to obtain the properties "
'      ' Msg = Msg + "available in StEPP.  It will also detail general "
'      ' Msg = Msg + "explanations of how to manipulate this Windows-based "
'      ' Msg = Msg + "program."
'      ' MsgBox Msg, MB_ICONINFORMATION, "Online Help System"
'      SendKeys "{F1}", True
    Case 10:      'ONLINE HELP.
      fn_This = MAIN_APP_PATH & "\help\stepp.hlp"
      If (fileexists(fn_This) = False) Then
        Call Show_Message("The file `" & fn_This & "` is missing.")
        Exit Sub
      End If
      Call ShellExecute_LocalFile(fn_This)
      ''''Call LaunchFile_General("", fn_This)
      '''''Call LaunchFile_General("", MAIN_APP_PATH & "\help\stepp.hlp")
    Case 20:      'ONLINE MANUAL.
      ''''fn_This = MAIN_APP_PATH & "\help\stepp.pdf"
      fn_This = MAIN_APP_PATH & "\help\stepp.doc"
      If (fileexists(fn_This) = False) Then
        Call Show_Message("The file `" & fn_This & "` is missing.")
        Exit Sub
      End If
      Call ShellExecute_LocalFile(fn_This)
      ''''Call LaunchFile_General("", fn_This)
      '''''Call LaunchFile_General("", MAIN_APP_PATH & "\help\stepp.pdf")
    Case 22:      'MANUAL PRINTING INSTRUCTIONS.
      fn_This = Global_fpath_dir_CPAS & "\dbase\printing.txt"
      If (fileexists(fn_This) = False) Then
        Call Show_Message("The file `" & fn_This & "` is missing.")
        Exit Sub
      End If
      Call Launch_Notepad(fn_This)
    Case 30:      'VIEW VERSION HISTORY.
      fn_This = App.Path & "\dbase\readme.txt"
      If (fileexists(fn_This) = False) Then
        Call Show_Message("The file `" & fn_This & "` is missing.")
        Exit Sub
      End If
      Call Launch_Notepad(fn_This)
    Case 40:      'VIEW DISCLAIMER.
      'SHOW THE DISCLAIMER WINDOW.
      splash_mode = 101
      splash_button_pressed = 0
      frmSplash.Show 1
    Case 50:      'TECHNICAL ASSISTANCE PROVIDED BY.
      'frmAbout2.Show 1
      frmTechnicalAssistance.Show 1
    Case 99:      'ABOUT.
      frmAbout.Show 1
  End Select
End Sub

Private Sub InitializeCurrentSelectionsTemp()

End Sub

Private Sub lblAirWaterProperties_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)

    If lblAirWaterProperties(Index).Caption = "" Then
       Exit Sub
    End If

    If Button = 1 Then 'left mouse button click so show appropriate form
       Select Case Index
          Case 0   'Water Density
               frmWaterDensity.Show 1
          Case 1   'Water Viscosity
               frmWaterViscosity.Show 1
          Case 2   'Water Surface Tension
               frmWaterSurfaceTension.Show 1
          Case 3   'Air Density
               frmAirDensity.Show 1
          Case 4   'Air Viscosity
               frmAirViscosity.Show 1
       End Select
    End If

    If Button = 2 Then 'right mouse button click so display appropriate message
       Select Case Index
          Case 0   'Water Density
               If lblAirWaterProperties(0).Caption = "Not Available" Then
                  msg = "Water Density is not available from StEPP."
                  MsgBox msg, MB_ICONINFORMATION, "Water - Data Unavailable"
               Else
                  Select Case phprop.WaterDensity.CurrentSelection.choice
                     Case WATER_DENSITY_CORRELATION
                        msg = "The currently selected Water Density is from a Polynomial Fit of Data Given in McCabe and Smith (1986).  For more detailed information about Water Density, click the left mouse button on the Water Density label or value on this screen."
                     Case WATER_DENSITY_INPUT
                        msg = "The currently selected Water Density is from User Input.  For more detailed information about Water Density, click the left mouse button on the Water Density label or value on this screen."
                  End Select
                  MsgBox msg, MB_ICONINFORMATION, "Water - Water Density Info."
               End If
               
          Case 1   'Water Viscosity
               If lblAirWaterProperties(1).Caption = "Not Available" Then
                  msg = "Water Viscosity is not available from StEPP."
                  MsgBox msg, MB_ICONINFORMATION, "Water - Data Unavailable"
               Else
                  Select Case phprop.WaterViscosity.CurrentSelection.choice
                     Case WATER_VISCOSITY_CORRELATION
                        msg = "The currently selected Water Viscosity is from a Correlation Presented in Reid, Prausnitz, and Poling (1987).  For more detailed information about Water Viscosity, click the left mouse button on the Water Viscosity label or value on this screen."
                     Case WATER_VISCOSITY_INPUT
                        msg = "The currently selected Water Viscosity is from User Input.  For more detailed information about Water Viscosity, click the left mouse button on the Water Viscosity label or value on this screen."
                  End Select
                  MsgBox msg, MB_ICONINFORMATION, "Water - Water Viscosity Info."
               End If
               
          Case 2   'Water Surface Tension
               If lblAirWaterProperties(2).Caption = "Not Available" Then
                  msg = "Water Surface Tension is not available from StEPP."
                  MsgBox msg, MB_ICONINFORMATION, "Water - Data Unavailable"
               Else
                  Select Case phprop.WaterSurfaceTension.CurrentSelection.choice
                     Case WATER_SURF_TENSION_CORRELATION
                        msg = "The currently selected Water Surface Tension is from a Correlation Presented in Cummins and Westrick (1983).  For more detailed information about Water Surface Tension, click the left mouse button on the Water Surface Tension label or value on this screen."
                     Case WATER_SURF_TENSION_INPUT
                        msg = "The currently selected Water Surface Tension is from User Input.  For more detailed information about Water Surface Tension, click the left mouse button on the Water Surface Tension label or value on this screen."
                  End Select
                  MsgBox msg, MB_ICONINFORMATION, "Water - Water Surface Tension Info."
               End If
               
          Case 3   'Air Density
               If lblAirWaterProperties(3).Caption = "Not Available" Then
                  msg = "Air Density is not available from StEPP."
                  MsgBox msg, MB_ICONINFORMATION, "Air - Data Unavailable"
               Else
                  Select Case phprop.AirDensity.CurrentSelection.choice
                     Case AIR_DENSITY_CORRELATION
                        msg = "The currently selected Air Density is from the Ideal Gas Law.  For more detailed information about Air Density, click the left mouse button on the Air Density label or value on this screen."
                     Case AIR_DENSITY_INPUT
                        msg = "The currently selected Air Density is from User Input.  For more detailed information about Air Density, click the left mouse button on the Air Density label or value on this screen."
                  End Select
                  MsgBox msg, MB_ICONINFORMATION, "Air - Air Density Info."
               End If
              
          Case 4   'Air Viscosity
               If lblAirWaterProperties(4).Caption = "Not Available" Then
                  msg = "Air Viscosity is not available from StEPP."
                  MsgBox msg, MB_ICONINFORMATION, "Air - Data Unavailable"
               Else
                  Select Case phprop.AirViscosity.CurrentSelection.choice
                     Case AIR_VISCOSITY_CORRELATION
                        msg = "The currently selected Air Viscosity is from a Correlation Presented in Cummins and Westrick (1983).  For more detailed information about Air Viscosity, click the left mouse button on the Air Viscosity label or value on this screen."
                     Case AIR_VISCOSITY_INPUT
                        msg = "The currently selected Air Viscosity is from User Input.  For more detailed information about Air Viscosity, click the left mouse button on the Air Viscosity label or value on this screen."
                  End Select
                  MsgBox msg, MB_ICONINFORMATION, "Air - Air Viscosity Info."
               End If
               
       End Select
    End If


End Sub

Private Sub lblAirWaterPropertiesLabel_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)

    Call lblAirWaterProperties_MouseDown(Index, Button, Shift, X, Y)

End Sub

Private Sub lblContaminantProperties_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim msg As String

    If lblContaminantProperties(Index).Caption = "" Then
       Exit Sub
    End If

    If Button = 1 Then 'left mouse button click so show appropriate form
       Select Case Index
          Case 0   'Vapor Pressure
               vp_form.Show 1
          Case 1   'Infinite Dilution Activity Coefficient
               Infinite_dilution_form.Show 1
          Case 2   'Henry's Constant
               hc_form.Show 1
          Case 3   'Molecular Weight
               mwt_form.Show 1
          Case 4   'Normal Boiling Point
               nbp_form.Show 1
          Case 5   'Liquid Density
               ldens_form.Show 1
          Case 6   'Molar Volume at Temperature of Interest
               molar_vol_form.Show 1
          Case 7   'Molar Volume at Normal Boiling Point
               mv_nbp_form.Show 1
          Case 8   'Refractive Index
               rindex_form.Show 1
          Case 9   'Aqueous Solubility
               aqsol_form.Show 1
          Case 10  'Octanol Water Partition Coefficient
               octanol_form.Show 1
          Case 11  'Liquid Diffusivity
               liquid_diff_form.Show 1
          Case 12  'Gas Diffusivity
               gas_diff_form.Show 1
       End Select
    End If

    If Button = 2 Then 'right mouse button click so display appropriate message
       Select Case Index
          Case 0   'Vapor Pressure
               If lblContaminantProperties(0).Caption = "Not Available" Then
                  msg = "Vapor Pressure is not available from StEPP for this chemical."
                  MsgBox msg, MB_ICONINFORMATION, Trim$(phprop.Name) & " - Data Unavailable"
               Else
                  Select Case phprop.VaporPressure.CurrentSelection.choice
                     Case VAPOR_PRESSURE_DATABASE
                        msg = "The currently selected Vapor Pressure is from the StEPP database.  For more detailed information about Vapor Pressure, click the left mouse button on the Vapor Pressure label or value on this screen."
                     Case VAPOR_PRESSURE_INPUT
                        msg = "The currently selected Vapor Pressure is from user input.  For more detailed information about Vapor Pressure, click the left mouse button on the Vapor Pressure label or value on this screen."
                  End Select
                  MsgBox msg, MB_ICONINFORMATION, Trim$(phprop.Name) & " - Vapor Pressure Info."
               End If
          Case 1   'Infinite Dilution Activity Coefficient
               If lblContaminantProperties(1).Caption = "Not Available" Then
                  msg = "Infinite Dilution Activity Coefficient is not available from StEPP for this chemical."
                  MsgBox msg, MB_ICONINFORMATION, Trim$(phprop.Name) & " - Data Unavailable"
               Else
                  Select Case phprop.ActivityCoefficient.CurrentSelection.choice
                     Case ACTIVITY_COEFFICIENT_UNIFAC
                        msg = "The currently selected Infinite Dilution Activity Coefficient is from UNIFAC.  For more detailed information about Activity Coefficient, click the left mouse button on the Activity Coefficient label or value on this screen."
                  End Select
                  MsgBox msg, MB_ICONINFORMATION, Trim$(phprop.Name) & " - Activity Coefficient Info."
               End If
          Case 2   'Henry's Constant
               If lblContaminantProperties(2).Caption = "Not Available" Then
                  msg = "Henry's Constant is not available from StEPP for this chemical."
                  MsgBox msg, MB_ICONINFORMATION, Trim$(phprop.Name) & " - Data Unavailable"
               Else
                  Select Case phprop.HenrysConstant.CurrentSelection.choice
                     Case HENRYS_CONSTANT_REGRESS
                        msg = "The currently selected Henry's Constant is from a Regression of Data Points in the StEPP database.  For more detailed information about Henry's Constant, click the left mouse button on the Henry's Constant label or value on this screen."
                     Case HENRYS_CONSTANT_FIT
                        msg = "The currently selected Henry's Constant is from a UNIFAC Fit with a Data Point.  For more detailed information about Henry's Constant, click the left mouse button on the Henry's Constant label or value on this screen."
                     Case HENRYS_CONSTANT_OPT_UNIFAC
                        msg = "The currently selected Henry's Constant is from UNIFAC at the Operating Temperature.  For more detailed information about Henry's Constant, click the left mouse button on the Henry's Constant label or value on this screen."
                     Case HENRYS_CONSTANT_DATABASE
                        msg = "The currently selected Henry's Constant is from the StEPP Database.  For more detailed information about Henry's Constant, click the left mouse button on the Henry's Constant label or value on this screen."
                     Case HENRYS_CONSTANT_UNIFAC
                        msg = "The currently selected Henry's Constant is from UNIFAC at a Temperature Corresponding to a Database Value.  For more detailed information about Henry's Constant, click the left mouse button on the Henry's Constant label or value on this screen."
                     Case HENRYS_CONSTANT_INPUT
                        msg = "The currently selected Henry's Constant is from User Input.  For more detailed information about Henry's Constant, click the left mouse button on the Henry's Constant label or value on this screen."
                  End Select
                  MsgBox msg, MB_ICONINFORMATION, Trim$(phprop.Name) & " - Henry's Constant Info."
               End If
          Case 3   'Molecular Weight
               If lblContaminantProperties(3).Caption = "Not Available" Then
                  msg = "Molecular Weight is not available from StEPP for this chemical."
                  MsgBox msg, MB_ICONINFORMATION, Trim$(phprop.Name) & " - Data Unavailable"
               Else
                  Select Case phprop.MolecularWeight.CurrentSelection.choice
                     Case MOLECULAR_WEIGHT_DATABASE
                        msg = "The currently selected Molecular Weight is from the StEPP Database.  For more detailed information about Molecular Weight, click the left mouse button on the Molecular Weight label or value on this screen."
                     Case MOLECULAR_WEIGHT_UNIFAC
                        msg = "The currently selected Molecular Weight is from Group Contribution Method.  For more detailed information about Molecular Weight, click the left mouse button on the Molecular Weight label or value on this screen."
                     Case MOLECULAR_WEIGHT_INPUT
                        msg = "The currently selected Molecular Weight is from User Input.  For more detailed information about Molecular Weight, click the left mouse button on the Molecular Weight label or value on this screen."
                  End Select
                  MsgBox msg, MB_ICONINFORMATION, Trim$(phprop.Name) & " - Molecular Weight Info."
               End If
              
          Case 4   'Normal Boiling Point
               If lblContaminantProperties(4).Caption = "Not Available" Then
                  msg = "Normal Boiling Point is not available from StEPP for this chemical."
                  MsgBox msg, MB_ICONINFORMATION, Trim$(phprop.Name) & " - Data Unavailable"
               Else
                  Select Case phprop.BoilingPoint.CurrentSelection.choice
                     Case BOILING_POINT_DATABASE
                        msg = "The currently selected Normal Boiling Point is from the StEPP Database.  For more detailed information about Normal Boiling Point, click the left mouse button on the Normal Boiling Point label or value on this screen."
                     Case BOILING_POINT_INPUT
                        msg = "The currently selected Normal Boiling Point is from User Input.  For more detailed information about Normal Boiling Point, click the left mouse button on the Normal Boiling Point label or value on this screen."
                  End Select
                  MsgBox msg, MB_ICONINFORMATION, Trim$(phprop.Name) & " - Boiling Point Info."
               End If
              
          Case 5   'Liquid Density
               If lblContaminantProperties(5).Caption = "Not Available" Then
                  msg = "Liquid Density is not available from StEPP for this chemical."
                  MsgBox msg, MB_ICONINFORMATION, Trim$(phprop.Name) & " - Data Unavailable"
               Else
                  Select Case phprop.LiquidDensity.CurrentSelection.choice
                     Case LIQUID_DENSITY_DATABASE
                        msg = "The currently selected Liquid Density is from the StEPP database.  For more detailed information about Liquid Density, click the left mouse button on the Liquid Density label or value on this screen."
                     Case LIQUID_DENSITY_UNIFAC
                        msg = "The currently selected Liquid Density is from Group Contribution Method.  For more detailed information about Liquid Density, click the left mouse button on the Liquid Density label or value on this screen."
                     Case LIQUID_DENSITY_INPUT
                        msg = "The currently selected Liquid Density is from User Input.  For more detailed information about Liquid Density, click the left mouse button on the Liquid Density label or value on this screen."
                  End Select
                  MsgBox msg, MB_ICONINFORMATION, Trim$(phprop.Name) & " - Liquid Density Info."
               End If
              
          Case 6   'Molar Volume at Temperature of Interest
               If lblContaminantProperties(6).Caption = "Not Available" Then
                  msg = "Molar Volume at the Operating Temperature is not available from StEPP for this chemical."
                  MsgBox msg, MB_ICONINFORMATION, Trim$(phprop.Name) & " - Data Unavailable"
               Else
                  Select Case phprop.MolarVolume.operatingT.CurrentSelection.choice
                     Case MOLAR_VOLUME_OPT_DATABASE
                        msg = "The currently selected Molar Volume at the Operating Temperature is from the StEPP database.  For more detailed information about Molar Volume at the Operating Temperature, click the left mouse button on the Molar Volume label or value on this screen."
                     Case MOLAR_VOLUME_OPT_UNIFAC
                        msg = "The currently selected Molar Volume at the Operating Temperature is from Group Contribution Method.  For more detailed information about Molar Volume at the Operating Temperature, click the left mouse button on the Molar Volume label or value on this screen."
                     Case MOLAR_VOLUME_OPT_INPUT
                        msg = "The currently selected Molar Volume at the Operating Temperature is from User Input.  For more detailed information about Molar Volume at the Operating Temperature, click the left mouse button on the Molar Volume label or value on this screen."
                  End Select
                  MsgBox msg, MB_ICONINFORMATION, Trim$(phprop.Name) & " - Molar Volume Info."
               End If
             
          Case 7   'Molar Volume at Normal Boiling Point
               If lblContaminantProperties(7).Caption = "Not Available" Then
                  msg = "Molar Volume at the Normal Boiling Point is not available from StEPP for this chemical."
                  MsgBox msg, MB_ICONINFORMATION, Trim$(phprop.Name) & " - Data Unavailable"
               Else
                  Select Case phprop.MolarVolume.BoilingPoint.CurrentSelection.choice
                     Case MOLAR_VOLUME_NBP_UNIFAC
                        msg = "The currently selected Molar Volume at the Normal Boiling Point is from Schroeder's Method.  For more detailed information about Molar Volume at the Normal Boiling Point, click the left mouse button on the Molar Volume label or value on this screen."
                     Case MOLAR_VOLUME_NBP_INPUT
                        msg = "The currently selected Molar Volume at the Normal Boiling Point is from User Input.  For more detailed information about Molar Volume at the Normal Boiling Point, click the left mouse button on the Molar Volume label or value on this screen."
                  End Select
                  MsgBox msg, MB_ICONINFORMATION, Trim$(phprop.Name) & " - Molar Volume Info."
               End If
             
          Case 8   'Refractive Index
               If lblContaminantProperties(8).Caption = "Not Available" Then
                  msg = "Refractive Index is not available from StEPP for this chemical."
                  MsgBox msg, MB_ICONINFORMATION, Trim$(phprop.Name) & " - Data Unavailable"
               Else
                  Select Case phprop.RefractiveIndex.CurrentSelection.choice
                     Case REFRACTIVE_INDEX_DATABASE
                        msg = "The currently selected Refractive Index is from the StEPP database.  For more detailed information about Refractive Index, click the left mouse button on the Refractive Index label or value on this screen."
                     Case REFRACTIVE_INDEX_INPUT
                        msg = "The currently selected Refractive Index is from User Input.  For more detailed information about Refractive Index, click the left mouse button on the Refractive Index label or value on this screen."
                  End Select
                  MsgBox msg, MB_ICONINFORMATION, Trim$(phprop.Name) & " - Refractive Index Info."
               End If
            
          Case 9   'Aqueous Solubility
               If lblContaminantProperties(9).Caption = "Not Available" Then
                  msg = "Aqueous Solubility is not available from StEPP for this chemical."
                  MsgBox msg, MB_ICONINFORMATION, Trim$(phprop.Name) & " - Data Unavailable"
               Else
                  Select Case phprop.AqueousSolubility.CurrentSelection.choice
                     Case AQUEOUS_SOLUBILITY_FIT
                        msg = "The currently selected Aqueous Solubility is from a Fit of the UNIFAC Curve with a Data Point.  For more detailed information about Aqueous Solubility, click the left mouse button on the Aqueous Solubility label or value on this screen."
                     Case AQUEOUS_SOLUBILITY_OPT_UNIFAC
                        msg = "The currently selected Aqueous Solubility is from a UNIFAC at the Operating Temperature.  For more detailed information about Aqueous Solubility, click the left mouse button on the Aqueous Solubility label or value on this screen."
                     Case AQUEOUS_SOLUBILITY_DATABASE
                        msg = "The currently selected Aqueous Solubility is from the StEPP database.  For more detailed information about Aqueous Solubility, click the left mouse button on the Aqueous Solubility label or value on this screen."
                     Case AQUEOUS_SOLUBILITY_DBT_UNIFAC
                        msg = "The currently selected Aqueous Solubility is from UNIFAC at a temperature corresponding to a point in the StEPP database.  For more detailed information about Aqueous Solubility, click the left mouse button on the Aqueous Solubility label or value on this screen."
                     Case AQUEOUS_SOLUBILITY_INPUT
                        msg = "The currently selected Aqueous Solubility is from User Input.  For more detailed information about Aqueous Solubility, click the left mouse button on the Aqueous Solubility label or value on this screen."
                  End Select
                  MsgBox msg, MB_ICONINFORMATION, Trim$(phprop.Name) & " - Aqueous Solubility Info."
               End If
              
          Case 10  'Octanol Water Partition Coefficient
               If lblContaminantProperties(10).Caption = "Not Available" Then
                  msg = "Octanol Water Partition Coefficient is not available from StEPP for this chemical."
                  MsgBox msg, MB_ICONINFORMATION, Trim$(phprop.Name) & " - Data Unavailable"
               Else
                  Select Case phprop.OctWaterPartCoeff.CurrentSelection.choice
                     Case OCT_WATER_PART_COEFF_DB
                        msg = "The currently selected Octanol Water Partition Coefficient is from the StEPP database.  For more detailed information about Octanol Water Partition Coefficient, click the left mouse button on the Octanol Water Partition Coefficient label or value on this screen."
                     Case OCT_WATER_PART_COEFF_DBT_UNIFAC
                        msg = "The currently selected Octanol Water Partition Coefficient is from UNIFAC at a temperature corresponding to a point in the StEPP database.  For more detailed information about Octanol Water Partition Coefficient, click the left mouse button on the Octanol Water Partition Coefficient label or value on this screen."
                     Case OCT_WATER_PART_COEFF_OPT_UNIFAC
                        msg = "The currently selected Octanol Water Partition Coefficient is from UNIFAC at the Operating Temperature.  For more detailed information about Octanol Water Partition Coefficient, click the left mouse button on the Octanol Water Partition Coefficient label or value on this screen."
                     Case OCT_WATER_PART_COEFF_INPUT
                        msg = "The currently selected Octanol Water Partition Coefficient is from User Input.  For more detailed information about Octanol Water Partition Coefficient, click the left mouse button on the Octanol Water Partition Coefficient label or value on this screen."
                  End Select
                  MsgBox msg, MB_ICONINFORMATION, Trim$(phprop.Name) & " - Oct. Water Part. Coeff. Info."
               End If
           
          Case 11  'Liquid Diffusivity
               If lblContaminantProperties(11).Caption = "Not Available" Then
                  msg = "Liquid Diffusivity is not available from StEPP for this chemical."
                  MsgBox msg, MB_ICONINFORMATION, Trim$(phprop.Name) & " - Data Unavailable"
               Else
                  Select Case phprop.LiquidDiffusivity.CurrentSelection.choice
                     Case LIQUID_DIFFUSIVITY_POLSON
                        msg = "The currently selected Liquid Diffusivity is from the Polson Correlation.  For more detailed information about Liquid Diffusivity, click the left mouse button on the Liquid Diffusivity label or value on this screen."
                     Case LIQUID_DIFFUSIVITY_HAYDUKLAUDIE
                        msg = "The currently selected Liquid Diffusivity is from the Hayduk and Laudie Correlation.  For more detailed information about Liquid Diffusivity, click the left mouse button on the Liquid Diffusivity label or value on this screen."
                     Case LIQUID_DIFFUSIVITY_WILKECHANG
                        msg = "The currently selected Liquid Diffusivity is from the Wilke-Chang Correlation.  For more detailed information about Liquid Diffusivity, click the left mouse button on the Liquid Diffusivity label or value on this screen."
                     Case LIQUID_DIFFUSIVITY_INPUT
                        msg = "The currently selected Liquid Diffusivity is from User Input.  For more detailed information about Liquid Diffusivity, click the left mouse button on the Liquid Diffusivity label or value on this screen."
                  End Select
                  MsgBox msg, MB_ICONINFORMATION, Trim$(phprop.Name) & " - Liquid Diffusivity Info."
               End If
            
          Case 12  'Gas Diffusivity
               If lblContaminantProperties(12).Caption = "Not Available" Then
                  msg = "Gas Diffusivity is not available from StEPP for this chemical."
                  MsgBox msg, MB_ICONINFORMATION, Trim$(phprop.Name) & " - Data Unavailable"
               Else
                  Select Case phprop.GasDiffusivity.CurrentSelection.choice
                     Case GAS_DIFFUSIVITY_WILKELEE
                        msg = "The currently selected Gas Diffusivity is from the Wilke-Lee Modification of the Hirschfelder-Bird-Spotz Method.  For more detailed information about Gas Diffusivity, click the left mouse button on the Gas Diffusivity label or value on this screen."
                     Case GAS_DIFFUSIVITY_INPUT
                        msg = "The currently selected Gas Diffusivity is from User Input.  For more detailed information about Gas Diffusivity, click the left mouse button on the Gas Diffusivity label or value on this screen."
                  End Select
                  MsgBox msg, MB_ICONINFORMATION, Trim$(phprop.Name) & " - Gas Diffusivity Info."
               End If
              
       End Select
         
    End If

End Sub

Private Sub lblContaminantPropertiesLabel_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    
    Call lblContaminantProperties_MouseDown(Index, Button, Shift, X, Y)
End Sub

Private Sub mnuAbout_Click(Index As Integer)
    Dim msg As String

    Select Case Index
       Case 0
          msg = "StEPP - Software to Estimate Physical Properties" & Chr$(13) & Chr$(13)
          msg = msg + "Version 1.00" & Chr$(13)
          msg = msg + "Copyright 1993, 1996" + Chr$(13)
          msg = msg + "Center for Clean Industrial and Treatment Technologies" & Chr$(13)
          msg = msg + "Michigan Technological University" & Chr$(13)
          msg = msg + "June 1, 1996" & Chr$(13) & Chr$(13)
'          msg = msg + "FOR PROPOSAL REVIEW ONLY" & Chr$(13)
'          msg = msg + "NOT FOR DISTRIBUTION" & Chr$(13) & Chr$(13)
'          msg = msg + "Status of This Program:  In Development" & Chr$(13) & Chr$(13) & Chr$(13)
          MsgBox msg, MB_ICONINFORMATION, "About StEPP"

'    msg = ""
'    msg = msg + "For further information, contact:" & Chr$(13) & Chr$(13)
'    msg = msg + "Tony Rogers" & Chr$(13)
'    msg = msg + "Dept. of Chemical Engineering" & Chr$(13)
'    msg = msg + "Michigan Technological University" & Chr$(13)
'    msg = msg + "1400 Townsend Drive" & Chr$(13)
'    msg = msg + "Houghton, MI 49931" & Chr$(13)
'    msg = msg + "Phone:  (906) 487-2210" & Chr$(13)
'    msg = msg + "Fax:  (906) 487-3213" & Chr$(13)
'    msg = msg + "E-mail:  tnrogers@mtu.edu" & Chr$(13)
    
       Case 1   'About the Authors
          msg = "David R. Hokanson:  " & Chr$(13)
          msg = msg + "     M.S. Candidate in Civil and Env. Eng." & Chr$(13) & Chr$(13)
          msg = msg + "Michael D. Miller:" & Chr$(13)
          msg = msg + "     M.S. Candidate in Chem. Eng." & Chr$(13) & Chr$(13)
          msg = msg + "Tony N. Rogers:" & Chr$(13)
          msg = msg + "     Asst. Professor of Chem. Eng." & Chr$(13) & Chr$(13)
          msg = msg + "David W. Hand:" & Chr$(13)
          msg = msg + "     Asst. Professor of Civil and Env. Eng." & Chr$(13) & Chr$(13)
          msg = msg + "Fr" & Chr$(233) & "d" & Chr$(233) & "ric Gobin:" & Chr$(13)
          msg = msg + "     Chemical Engineer:  Elf Aquitaine, Inc." & Chr$(13) & Chr$(13)
          msg = msg + "Matthew Buchkowski:" & Chr$(13)
          msg = msg + "     B.S. Candidate in Chem. Eng." & Chr$(13) & Chr$(13)
          msg = msg + "John C. Crittenden:" & Chr$(13)
          msg = msg + "     Presidential Professor of Civil and Env. Eng" & Chr$(13) & Chr$(13)

          MsgBox msg, MB_ICONINFORMATION, "About the Authors"

        Case 2   'About Programming Support for this Product
           'msg = "Programming support for the database used in "
           'msg = msg + "this product was provided by Thomas F. Budd, "
           'msg = msg + "a B.S. Candidate in Computer Science."
            msg = "Richard J. Hossli:  " & Chr$(13)
            msg = msg + "     a Software developer" & Chr$(13) & Chr$(13)
            msg = msg + "Jason E. Mclean:  " & Chr$(13)
            msg = msg + "     a B.S. Candidate in Computer Science" & Chr$(13) & Chr$(13)
            msg = msg + "Eric J. Oman:  " & Chr$(13)
            msg = msg + "     a B.S. Candidate in Chem. Eng" & Chr$(13) & Chr$(13)
            msg = msg + "Thomas F. Budd:  " & Chr$(13)
            msg = msg + "     a B.S. Candidate in Computer Science" & Chr$(13) & Chr$(13)

           MsgBox msg, MB_ICONINFORMATION, "About Programming Support"

        Case 3   'Obtaining Additional Information
           msg = "For product release information when it becomes "
           msg = msg + "available, contact:" & Chr$(13) & Chr$(13)
           msg = msg + "    Dr. David Hand" & Chr$(13)
           msg = msg + "    Dept. of Civil and Env. Eng." & Chr$(13)
           msg = msg + "    Michigan Tech. University" & Chr$(13)
           msg = msg + "    1400 Townsend Drive" & Chr$(13)
           msg = msg + "    Houghton, MI 49931" & Chr$(13)
           msg = msg + "    Phone:   (906) 487-2777" & Chr$(13)
           msg = msg + "    E-mail:  dwhand@mtu.edu" & Chr$(13) & Chr$(13) & Chr$(13)
           msg = msg + "Also feel free to contact Dr. Hand with "
           msg = msg + "any comments, problems, or suggestions related "
           msg = msg + "to the StEPP program."

           MsgBox msg, MB_ICONINFORMATION, "About Obtaining Additional Information"

    End Select

End Sub

Private Sub mnuFile_Click(Index As Integer)
    Dim J As Integer, Response As Integer
    Dim msg As String

    Screen.MousePointer = 11   'Hourglass

    Select Case Index
       Case 0   'New
            Response = MsgBox("Save current data?", MB_ICONQUESTION + MB_YESNOCANCEL, "StEPP")
            If Response = IDCANCEL Then
              Screen.MousePointer = 0
              Exit Sub
            End If
            If Response = IDYES Then
              'ChDrive SaveAndLoadPath
              'ChDir SaveAndLoadPath
              Call ChangeDir_Main
              Call SaveStEPPDesign
              'Add this file to the last-few-files list if necessary.
              Call LastFewFiles_MoveFilenameToTop(FileName$)
              SaveAndLoadPath = CurDir$
              'ChDrive steppPath
              'ChDir steppPath
              Call ChangeDir_Main
            End If

            cboSelectContaminant.Clear
            'For j = 0 To cboSelectContaminant.ListCount - 1
            '  cboSelectContaminant.RemoveItem j
            'Next

            For J = 0 To 12
              lblContaminantProperties(J).Caption = ""
            Next
            
            For J = 0 To 4
              lblAirWaterProperties(J).Caption = ""
            Next

            lblSelectedContaminant.Caption = ""
            NumSelectedChemicals = 0
            Call setup_form

       Case 1   'Open
            Response = MsgBox("Save current data?", MB_ICONQUESTION + MB_YESNOCANCEL, "StEPP")
            If Response = IDCANCEL Then
              Screen.MousePointer = 0
              Exit Sub
            End If
            If Response = IDYES Then
              'ChDrive SaveAndLoadPath
              'ChDir SaveAndLoadPath
              Call ChangeDir_Main
              Call SaveStEPPDesign
              '''''Add this file to the last-few-files list if necessary.
              ''''Call LastFewFiles_MoveFilenameToTop(FileName$)
              SaveAndLoadPath = CurDir$
              'ChDrive steppPath
              'ChDir steppPath
              Call ChangeDir_Main
            End If
            
            'ChDrive SaveAndLoadPath
            'ChDir SaveAndLoadPath
            Call ChangeDir_Main
            
            Call LoadStEPPDesign("")
       Call Update_P_and_T_StEPPLink
       'If (SteppLink_SpecifiedPressure <> "") Then
       '  Call UpdatePressAllCompounds(CDbl(SteppLink_SpecifiedPressure))
       'End If
       'If (SteppLink_SpecifiedTemperature <> "") Then
       '  Call UpdateTempAllCompounds(CDbl(SteppLink_SpecifiedTemperature))
       'End If
       
            '''''Add this file to the last-few-files list if necessary.
            ''''Call LastFewFiles_MoveFilenameToTop(FileName$)
            SaveAndLoadPath = CurDir$
            txtOperatingTemperature.SetFocus
            'ChDrive steppPath
            'ChDir steppPath
            Call ChangeDir_Main
       
       Case 4   'Save
                'ChDrive SaveAndLoadPath
                'ChDir SaveAndLoadPath
                Call ChangeDir_Main
                Call SaveStEPPDesign
                'Add this file to the last-few-files list if necessary.
                Call LastFewFiles_MoveFilenameToTop(FileName$)
                SaveAndLoadPath = CurDir$
                'ChDrive steppPath
                'ChDir steppPath
                Call ChangeDir_Main

       Case 5   'Save As
            'ChDrive SaveAndLoadPath
            'ChDir SaveAndLoadPath
            Call ChangeDir_Main
            OldFileName$ = FileName$
            FileName$ = ""
            Call SaveStEPPDesign
            'Add this file to the last-few-files list if necessary.
            Call LastFewFiles_MoveFilenameToTop(FileName$)
            SaveAndLoadPath = CurDir$
            'ChDrive steppPath
            'ChDir steppPath
            Call ChangeDir_Main
       
       Case 7   'Print
            'Place current properties into PropContaminant Structure
           PropContaminant(contam_prop_form!cboSelectContaminant.ListIndex + 1) = phprop

           For J = 1 To NUMBER_OF_PROPERTIES
               PropContaminant(contam_prop_form!cboSelectContaminant.ListIndex + 1).HaveProperty(J) = HaveProperty(J)
           Next J
           For J = 1 To NUMBER_OF_PROPERTIES_AVAILABLE
               PropContaminant(contam_prop_form!cboSelectContaminant.ListIndex + 1).PROPAVAILABLE(J) = PROPAVAILABLE(J)
           Next J

            frmPrint!lblCurrentContaminant.Caption = contam_prop_form!lblSelectedContaminant.Caption
            frmPrint.Show
       
       Case 8   'Select Printer
          On Error GoTo No_Default_Printer
          'CMDialog1.flags = PD_PRINTSETUP
          'CMDialog1.Action = 5
          CommonDialog1.ShowPrinter
       
       Case 200  'Exit
            'NOTE: It is safe to have an "unload me" command here
            'due to the fact that Form_QueryUnload() takes care of checking
            'on (a) asking the user if they want to save their data and
            '(b) telling the user they can't leave before they either
            'cancel or use components in the current StEPP link.
            Unload Me
            Exit Sub
    
    End Select


    If ((Index >= 191) And (Index <= 194)) Then
      'Handle File|Open of a file here.
      'ChDrive SaveAndLoadPath
      'ChDir SaveAndLoadPath
      Call ChangeDir_Main
      If (Dir(Current_LastFewFilesRec.FileNames(Index - 190)) = "") Then
        Beep
        MsgBox "That file has been moved or deleted.", MB_ICONEXCLAMATION, "StEPP"
      Else
            Response = MsgBox("Save current data?", MB_ICONQUESTION + MB_YESNOCANCEL, "StEPP")
            If Response = IDCANCEL Then
              Screen.MousePointer = 0
              Exit Sub
            End If
            If Response = IDYES Then
              'ChDrive SaveAndLoadPath
              'ChDir SaveAndLoadPath
              Call ChangeDir_Main
              Call SaveStEPPDesign
              'Add this file to the last-few-files list if necessary.
              Call LastFewFiles_MoveFilenameToTop(FileName$)
              SaveAndLoadPath = CurDir$
              'ChDrive steppPath
              'ChDir steppPath
              Call ChangeDir_Main
            End If
        
        Call LoadStEPPDesign(Current_LastFewFilesRec.FileNames(Index - 190))
       Call Update_P_and_T_StEPPLink
       'If (SteppLink_SpecifiedPressure <> "") Then
       '  Call UpdatePressAllCompounds(CDbl(SteppLink_SpecifiedPressure))
       'End If
       'If (SteppLink_SpecifiedTemperature <> "") Then
       '  Call UpdateTempAllCompounds(CDbl(SteppLink_SpecifiedTemperature))
       'End If
        
        'Add this file to the last-few-files list if necessary.
        Call LastFewFiles_MoveFilenameToTop(FileName$)
        SaveAndLoadPath = CurDir$
      End If
      'ChDir App.Path
      'ChDrive App.Path
      Call ChangeDir_Main
    End If

    Screen.MousePointer = 0    'Arrow
Exit Sub

No_Default_Printer:
    Screen.MousePointer = 0    'Arrow     'err.description
    If Err = 28663 Then
        'do nothing msg already comes up
    Else
    MsgBox "Error " & Err & " Occured"
       
    End If

Resume Next

End Sub

Private Sub mnuOptionsEtc_Click(Index As Integer)

    Select Case Index
       Case 0   'Modify Hierarchy
'          frmHierarchy.Show 1
    End Select

End Sub

Private Sub mnuOptionsItem_Click(Index As Integer)
Dim f As Integer
Dim i As Integer
Dim index0 As Integer
Dim this_name As String
Dim this_cas As String
Dim temp1 As String
Dim ExportFile As String
Dim cliptext As String
Dim vb3CrLf As String
Dim Ctl As Control
Set Ctl = contam_prop_form.CommonDialog2

  index0 = cboSelectContaminant.ListCount
  If (index0 = 0) Then
    MsgBox "No chemicals were selected for export!", MB_ICONEXCLAMATION, "StEPP"
    Exit Sub
  End If
  
  Select Case Index
    Case 10:
      'get chemical selected from "selected contaminants and see if valid
      On Error Resume Next
'      contam_prop_form!CMDialog1.FileName = ""
'      contam_prop_form!CMDialog1.DefaultExt = "exp"
'      contam_prop_form!CMDialog1.Filter = "StEPP Export Files (*.exp)|*.exp"
'      contam_prop_form!CMDialog1.DialogTitle = "Save StEPP Export File"
'      contam_prop_form!CMDialog1.flags = OFN_OVERWRITEPROMPT Or OFN_PATHMUSTEXIST
'      contam_prop_form!CMDialog1.CancelError = True
'      contam_prop_form!CMDialog1.Action = 2
'      ExportFile = contam_prop_form!CMDialog1.FileName
'      contam_prop_form!CMDialog1.FileName = ""
      ''''Ctl.FileName = ""
      Ctl.DefaultExt = "exp"
      Ctl.Filter = "StEPP Export Files (*.exp)|*.exp"
      Ctl.DialogTitle = "Save StEPP Export File"
      Ctl.flags = OFN_OVERWRITEPROMPT Or OFN_PATHMUSTEXIST
      Ctl.CancelError = True
      Ctl.Action = 2
      ExportFile = Ctl.FileName
      ''''Ctl.FileName = ""
      If Err = 32755 Then   'Cancel selected by user
        Exit Sub
      End If
      'GENERATE TEXT TO BE COPIED.
      Call ExportFileGeneration_GetText(cliptext)
      'PERFORM ACTUAL OUTPUT OF FILE.
      f = FreeFile
      Open ExportFile For Output As #f
      Print #f, cliptext
      Close #f
      '
      ' DISPLAY EXPLANATION MESSAGE.
      '
      Call Show_Message("The StEPP export file named `" & ExportFile & _
          "` was successfully generated.")
          

      ''get chemical selected from "selected contaminants and see if valid
      'index0 = cboSelectContaminant.ListCount
      '
      'If (index0 = 0) Then
      '    MsgBox "No chemicals were selected for export!", MB_ICONEXCLAMATION, "StEPP"
      '    Exit Sub
      'End If
      '
      ''******** add box to choose file type
      'On Error Resume Next
      'contam_prop_form!CMDialog1.Filename = ""
      'contam_prop_form!CMDialog1.DefaultExt = "exp"
      'contam_prop_form!CMDialog1.Filter = "StEPP export Files (*.exp)|*.exp"
      'contam_prop_form!CMDialog1.DialogTitle = "Save StEPP export File"
      'contam_prop_form!CMDialog1.Flags = OFN_OVERWRITEPROMPT Or OFN_PATHMUSTEXIST
      'contam_prop_form!CMDialog1.CancelError = True
      'contam_prop_form!CMDialog1.Action = 2
      'ExportFile = contam_prop_form!CMDialog1.Filename
      'contam_prop_form!CMDialog1.Filename = ""
      '
      'If Err = 32755 Then   'Cancel selected by user
      '  Exit Sub
      'End If
      '
      ''open file and spew all data on chemical
      'f = FreeFile
      'Open ExportFile For Output As #f
      '
      ''the extra data that will be the check for right file
      'Write #f, "1234567890"
      'Write #f, txtoperatingpressure
      'Write #f, txtOperatingTemperature
      '
      ''spew out the data
      'For i = 0 To cboSelectContaminant.ListCount - 1
      '  cboSelectContaminant.ListIndex = i
      '  temp1 = LTrim$(cboSelectContaminant.List(cboSelectContaminant.ListIndex))
      '  Call parsedargs_getarg(" ", temp1, 1, this_cas)
      '  this_name = Trim$(lblSelectedContaminant)
      '  Write #f, "Chemical", this_name, this_cas
      '  Call SteppLink_OutputProperty(f, 0, "VaporPressure", "Pa")
      '  Call SteppLink_OutputProperty(f, 1, "ActivityCoefficient", "-")
      '  Call SteppLink_OutputProperty(f, 2, "HenrysConstant", "-")
      '  Call SteppLink_OutputProperty(f, 3, "MolecularWeight", "kg/kmol")
      '  Call SteppLink_OutputProperty(f, 4, "NormalBoilingPoint", "C")
      '  Call SteppLink_OutputProperty(f, 5, "LiquidDensity", "kg/m3")
      '  Call SteppLink_OutputProperty(f, 6, "MolarVolumeAtOpT", "m3/kmol")
      '  Call SteppLink_OutputProperty(f, 7, "MolarVolumeAtNBP", "m3/kmol")
      '  Call SteppLink_OutputProperty(f, 8, "RefractiveIndex", "-")
      '  Call SteppLink_OutputProperty(f, 9, "AqueousSolubility", "PPMw")
      '  Call SteppLink_OutputProperty(f, 10, "LogKOW", "-")
      '  Call SteppLink_OutputProperty(f, 11, "LiquidDiffusivity", "m2/s")
      '  Call SteppLink_OutputProperty(f, 12, "GasDiffusivity", "m2/s")
      'Next i
      'Write #f, "END_OF_FILE", "", ""
      'Close #f
      '
      ''go back to correct directory
      '
      'ChDrive steppPath
      'ChDir steppPath
  
    Case 20:
      'GENERATE TEXT TO BE COPIED.
      Call ExportFileGeneration_GetText(cliptext)
      'PERFORM ACTUAL COPY TO CLIPBOARD.
      Clipboard.Clear
      Clipboard.SetText cliptext
      '
      ' DISPLAY EXPLANATION MESSAGE.
      '
      Call Show_Message("The StEPP properties were successfully copied " & _
          "to the clipboard.  They are now available for import into " & _
          "the AdDesignS or ASAP programs.  Warning: If you copy anything " & _
          "else to the clipboard, the property data will be erased.")
  End Select

End Sub

Private Sub mnuunits_click(Index As Integer)
    Dim SIValue As Double
    Dim EnglishValue As Double
    Dim msg As String

    If mnuUnits(Index).CHECKED = True Then Exit Sub

    If (SteppLink_Status = STEPPLINK_STATUS_ACTIVE) Then
      msg = "Cannot perform unit change--the StEPP link requires that properties be kept in SI units."
      MsgBox msg, MB_ICONEXCLAMATION, "StEPP"
      Exit Sub
    End If

    cmdSelectContaminant.SetFocus

    mnuUnits(Index).CHECKED = True
    Select Case Index
       Case 0   'SI Units
          CurrentUnits = SIUnits
          mnuUnits(1).CHECKED = False
       Case 1   'English Units
          CurrentUnits = EnglishUnits
          mnuUnits(0).CHECKED = False
    End Select

    Call GetUnits

    'Convert operating temperature and operating pressure
    Select Case Index
       Case 0   'SI Units
            txtOperatingPressure.Text = Str$(phprop.OperatingPressure)
            txtOperatingTemperature.Text = Str$(phprop.OperatingTemperature)
       Case 1   'English Units
            SIValue = phprop.OperatingPressure
            Call PRESSCNV(EnglishValue, SIValue)
            txtOperatingPressure.Text = Str$(EnglishValue)

            SIValue = phprop.OperatingTemperature
            Call TEMPCNV(EnglishValue, SIValue)
            txtOperatingTemperature.Text = Str$(EnglishValue)
    End Select

    If NumSelectedChemicals = 0 Then Exit Sub

    Call DisplayAllProperties

End Sub

Private Sub Search_String(J As Integer)
    Dim i As Integer, Res As Integer
    For i = J + 1 To contam_combo.ListCount
      Res = InStr(1, contam_combo.List(i), Find_String, 1)
      If Res > 0 Then
        contam_combo.ListIndex = i
        contam_combo.TopIndex = contam_combo.ListIndex
        'contam_combo.Selected(0) = True
        contam_combo.SetFocus
        Exit Sub
      End If
    Next i
    For i = 0 To J
      Res = InStr(1, contam_combo.List(i), Find_String, 1)
      If Res > 0 Then
        contam_combo.ListIndex = i
        contam_combo.TopIndex = contam_combo.ListIndex
        'contam_combo.Selected(0) = True
        contam_combo.SetFocus
        Exit Sub
      End If
    Next i

MsgBox "String not Found", 64, "Warning"
End Sub

Private Sub set_popup_defaults()
    vp_form.Option1(1) = True
    Infinite_dilution_form.Option1(1) = True
    hc_form.Option1(3) = True
    mwt_form.Option1(1) = True
    nbp_form.Option1(1) = True
    ldens_form.Option1(1) = True
    molar_vol_form.Option1(1) = True
    mv_nbp_form.Option1(1) = True
    rindex_form.Option1(1) = True
    aqsol_form.Option1(1) = True
    octanol_form.Option1(1) = True
    liquid_diff_form.Option1(1) = True
    gas_diff_form.Option1(1) = True
End Sub

Private Sub setup_form()
    Dim i As Integer
    Dim arraysize As Integer
    Const chunksize = 1500
    Dim lastcas As Double
    Dim Index As Integer    ' current record position in db
    
    set_popup_defaults
'
' We want the properties window up... not the test one
    
    contam_prop_form.Width = 9600
    contam_prop_form.Height = 7200
'    contam_prop_form.Height = 7200
'    contam_prop_form.Left = 0
'    contam_prop_form.Top = 0
    contam_prop_form.WindowState = 0

    'Center the form on the screen
    If WindowState = 0 Then
       'don't attempt if screen Minimized or Maximized
       Move (Screen.Width - contam_prop_form.Width) / 2, (Screen.Height - contam_prop_form.Height) / 2
    End If
    cmdUnselectContaminant.Enabled = False


'Initialize pressure and temperature

     txtOperatingPressure.Text = "101325"
     txtOperatingTemperature.Text = "25.0"
     phprop.OperatingPressure = CDbl(txtOperatingPressure.Text)
     phprop.OperatingTemperature = CDbl(txtOperatingTemperature.Text)
     dbinput.OperatingTemperature = phprop.OperatingTemperature


'Place labels on to all forms

'     CurrentUnits = ENGLISHUNITS
     CurrentUnits = SIUnits
     mnuUnits(0).CHECKED = True
     Call GetUnits

'Initialize File Name
     FileName$ = ""
     OldFileName$ = ""
     contam_prop_form!mnuFile(4).Enabled = False
     contam_prop_form!mnuFile(5).Enabled = False
     contam_prop_form!mnuFile(7).Enabled = False
     Call frmMain_Reset_DemoVersionDisablings
     
'Initialize things on frmPrint
    frmPrint!optDestination(0).Value = True
    frmPrint!optPrintContaminants(0).Value = True
    frmPrint!optPrintProperties(0).Value = True
    For i = 0 To 18
        frmPrint!chkProperties(i).Enabled = False
    Next i

    frmPrint!cboPropertyDescription.AddItem "Print Selected Values Only"
    frmPrint!cboPropertyDescription.AddItem "Print Full Description of Properties"
    frmPrint!cboPropertyDescription.ListIndex = 0

    frmPrint!cboUnits.AddItem "Print Values in SI Units"
    frmPrint!cboUnits.AddItem "Print Values in English Units"
    frmPrint!cboUnits.ListIndex = 0

    PreviouslySelectedIndex = -1

    contam_prop_form_ActivatedYet = False

    cboSelectContaminant.Enabled = False

    '---- Signal client program that StEPP has loaded
    If (SteppLink_fn_loadup_waitfile <> "") Then
      If (Dir(SteppLink_fn_loadup_waitfile) <> "") Then
        Kill SteppLink_fn_loadup_waitfile
      End If
    End If

End Sub

Private Sub txtOperatingPressure_GotFocus()
  Call GotFocus_Handle(Me, txtOperatingPressure, Temp_Text)

End Sub

Private Sub txtOperatingPressure_KeyPress(KeyAscii As Integer)
 
    If KeyAscii = 13 Then
       KeyAscii = 0
       txtOperatingTemperature.SetFocus
       Exit Sub
    End If
    Call NumberCheck(KeyAscii)
    
End Sub

Private Sub txtOperatingPressure_LostFocus()
    Dim i As Integer
    Dim msg As String, Response As Integer
    Dim Answer As Integer
    Dim IsError As Integer
    Dim ValueChanged As Integer
    Dim NumContaminantInList As Integer
    Dim EnglishValue As Double, SIValue As Double
    Dim NewPressure As Double
    Dim flag_ok As Integer
    Dim NewTemperature As Double

   If (LostFocus_IsEvil(Me, txtOperatingPressure)) Then
     Exit Sub
   End If

   flag_ok = True
    Call TextHandleError(IsError, txtOperatingPressure, Temp_Text)
    If Not IsError Then
       If Not HaveNumber(CDbl(txtOperatingPressure.Text)) Then
          txtOperatingPressure.Text = Temp_Text
          txtOperatingPressure.SetFocus
          Call LostFocus_Handle(Me, txtOperatingPressure, flag_ok)
          Exit Sub
       End If

       Call TextNumberChanged(ValueChanged, txtOperatingPressure, Temp_Text)

       If ValueChanged Then
          If CurrentUnits = SIUnits Then
             phprop.OperatingPressure = CDbl(txtOperatingPressure.Text)
          Else
             EnglishValue = CDbl(txtOperatingPressure.Text)
             Call PRESENSI(SIValue, EnglishValue)
             phprop.OperatingPressure = SIValue
          End If
          NewPressure = phprop.OperatingPressure
       Else
         Call LostFocus_Handle(Me, txtOperatingPressure, flag_ok)
         Exit Sub
       End If

       If contam_prop_form!cboSelectContaminant.ListCount = 0 Then
            Call LostFocus_Handle(Me, txtOperatingPressure, flag_ok)
            Exit Sub
       End If

       If HaveNumber(phprop.OperatingPressure) And HaveNumber(phprop.OperatingTemperature) Then
          'If updating pressure is just supposed to update the
          'currently highlighted component then call this routine:
             'Call UpdatePressCurrentCompound

          'If updating pressure is supposed to update ALL
          'currently selected components then call this routine
             
             ''''Call UpdatePressAllCompounds(NewPressure)
             
             'INSTEAD OF CALLING THE PRESSURE UPDATE ROUTINE,
             'THE TEMPERATURE UPDATE ROUTINE IS CALLED.  THE ONLY
             'PRESSURE-DEPENDENT PROPERTIES ARE (I THINK):
             'GAS DIFFUSIVITY AND AIR DENSITY.  UNFORTUNATELY, THE
             'UpdatePressAllCompounds() SUBROUTINE WAS NOT HANDLING
             'THESE UPDATES PROPERLY.  THE UpdateTempAllCompounds()
             'SUBROUTINE SEEMS TO HANDLE THEM PROPERLY.
             '    - ERIC J. OMAN, 6/18/98
             NewTemperature = phprop.OperatingTemperature
             Call UpdateTempAllCompounds(NewTemperature)

       End If
    End If

  Call LostFocus_Handle(Me, txtOperatingPressure, flag_ok)

End Sub

Private Sub txtOperatingTemperature_GotFocus()
  Call GotFocus_Handle(Me, txtOperatingTemperature, Temp_Text)
End Sub

Private Sub txtOperatingTemperature_KeyPress(KeyAscii As Integer)

    If KeyAscii = 13 Then
       KeyAscii = 0
       contam_combo.SetFocus
       Exit Sub
    End If
    
    Call NumberCheck(KeyAscii)

End Sub

Private Sub txtOperatingTemperature_LostFocus()
    Dim NumContaminantInList As Integer
    Dim EnglishValue As Double, SIValue As Double
    Dim NewTemperature As Double
    Dim flag_ok As Integer

   If (LostFocus_IsEvil(Me, txtOperatingTemperature)) Then
     Exit Sub
   End If

   flag_ok = True

    Dim Answer As Integer, Response As Integer
    Dim msg As String, IsError As Integer, ValueChanged As Integer

    Call TextHandleError(IsError, txtOperatingTemperature, Temp_Text)
    If Not IsError Then
       If Not HaveTemp(CDbl(txtOperatingTemperature.Text)) Then
          txtOperatingTemperature.Text = Temp_Text
          txtOperatingTemperature.SetFocus
          Call LostFocus_Handle(Me, txtOperatingTemperature, flag_ok)
          Exit Sub
       End If
       
       Call TextNumberChanged(ValueChanged, txtOperatingTemperature, Temp_Text)

       If ValueChanged Then
          If CurrentUnits = SIUnits Then
             phprop.OperatingTemperature = CDbl(txtOperatingTemperature.Text)
             dbinput.OperatingTemperature = phprop.OperatingTemperature
          Else
             EnglishValue = CDbl(txtOperatingTemperature.Text)
             Call TEMPENSI(SIValue, EnglishValue)
             phprop.OperatingTemperature = SIValue
             dbinput.OperatingTemperature = phprop.OperatingTemperature
          End If
          NewTemperature = phprop.OperatingTemperature
       Else
          Call LostFocus_Handle(Me, txtOperatingTemperature, flag_ok)
          Exit Sub
       End If

       If contam_prop_form!cboSelectContaminant.ListCount = 0 Then
            Call LostFocus_Handle(Me, txtOperatingTemperature, flag_ok)
            Exit Sub
       End If

       If HaveNumber(phprop.OperatingPressure) And HaveTemp(phprop.OperatingTemperature) Then
          
          'If updating temperature is just supposed to update the
          'currently highlighted component then call this routine:
             'Call UpdateTempCurrentCompound

          'If updating temperature is supposed to update ALL
          'currently selected components then call this routine
             Call UpdateTempAllCompounds(NewTemperature)

       End If
    End If
  Call LostFocus_Handle(Me, txtOperatingTemperature, flag_ok)
    
End Sub

'
' NOTE: The Update_Fieldstuff() SUBROUTINE IS NOT CALLED
' FROM ANYWHERE ELSE IN THE ENTIRE PROJECT.
'
'Private Sub Update_Fieldstuff(RecordNo As Integer)
'
'    Dim i As Long
'    Dim j As Long
'    Dim K As Long
'    Dim TempD As Double
'    Dim HC_Count As Long
'    Dim HC_DB_Source As String
'    Dim HC_DB_Value As String * 36
'    Dim HC_DB_Temp As String
'
'    dbinput.CASNumber = db_index(RecordNo + 1)
'
'    '
'    ' OPEN RECORDSET.
'    '
'    Set RS_Main = DB_Main.OpenRecordset( _
'        "SELECT * FROM [Names (Master)] WHERE [Names (Master)].CAS = " & _
'        Format$(dbinput.CASNumber, "0"))
'    RS_Main.MoveFirst
'    RS_Main.MoveLast
'    RS_Main.MoveFirst
'    Set Selection = RS_Main
'    'If (DemoMode) Then
'    '    Data1.DatabaseName = Database_Path + "\demo_db.mdb"
'    'Else
'    '    Data1.DatabaseName = Database_Path + "\stepp_db.mdb"
'    'End If
'    'Data1.RecordSource = "SELECT * FROM [Names (Master)] WHERE [Names (Master)].CAS = " & Format$(dbinput.CASNumber, "0")
'    'Data1.Refresh
'    'Set Selection = Data1.Recordset
'
'    dbinput.Name = Selection(2)
'
'    'Look into the Properties Table ----------------------------------
'
'    '
'    ' OPEN RECORDSET.
'    '
'    Set RS_Main = DB_Main.OpenRecordset( _
'        "SELECT * FROM DIPPR801 WHERE DIPPR801.CAS = " & _
'        Format$(dbinput.CASNumber, "0"))
'    RS_Main.MoveFirst
'    RS_Main.MoveLast
'    RS_Main.MoveFirst
'    Set Selection = RS_Main
'    'Data1.RecordSource = "SELECT * FROM DIPPR801 WHERE DIPPR801.CAS = " & Format$(dbinput.CASNumber, "0")
'    'Data1.Refresh
'    'Set Selection = Data1.Recordset
'
'    If Selection.EOF = False Then
'
'        dbinput.formula = nullcheck(Selection("FORM"))
'        dbinput.MolecularWeight = Selection("MW")
'        dbinput.BoilingPoint = Selection("NBP")
'        dbinput.BoilingPointSource = get_source(nullcheck("DIPPR801"))
'        dbinput.RefractiveIndex = Selection("RI")
'        dbinput.VaporPressureDatabaseEquation = Selection("VPEQN")
'        dbinput.VaporPressureNumberCoefficients = Selection("VPNUM")
'        dbinput.VaporPressureAntoineA = Selection("VPA")
'        dbinput.VaporPressureAntoineB = Selection("VPB")
'        dbinput.VaporPressureAntoineC = Selection("VPC")
'        dbinput.VaporPressureAntoineD = Selection("VPD")
'        dbinput.VaporPressureAntoineE = Selection("VPE")
'        dbinput.VaporPressureMinimumT = Selection("VPTMIN")
'        dbinput.VaporPressureMaximumT = Selection("VPTMAX")
'        dbinput.VaporPressureSource = get_source(nullcheck("DIPPR801"))
'        dbinput.LiquidDensityEquation = Selection("LDNEQN")
'        dbinput.LiquidDensityNumberCoefficients = Selection("LDNNUM")
'        dbinput.LiquidDensityCoefficientA = Selection("LDNA")
'        dbinput.LiquidDensityCoefficientB = Selection("LDNB")
'        dbinput.LiquidDensityCoefficientC = Selection("LDNC")
'        dbinput.LiquidDensityCoefficientD = Selection("LDND")
'        dbinput.LiquidDensityMinimumT = Selection("LDNTMIN")
'        dbinput.LiquidDensityMaximumT = Selection("LDNTMAX")
'        dbinput.LiquidDensitySource = get_source(nullcheck("DIPPR801"))
'
'    Else
'
'        dbinput.MolecularWeight = -1
'        dbinput.BoilingPointSource = -1
'        dbinput.RefractiveIndex = -1
'        dbinput.VaporPressureAntoineA = -1
'        dbinput.VaporPressureDatabaseEquation = -1
'        dbinput.LiquidDensityEquation = -1
'
'    End If
'
'    If dbinput.MolecularWeight = 0 Then
'        dbinput.MolecularWeight = -1
'    End If
'
'    If dbinput.BoilingPoint = 0 Then
'        dbinput.BoilingPointSource = -1
'    End If
'
'    If dbinput.RefractiveIndex = 0 Then
'        dbinput.RefractiveIndex = -1
'    End If
'
'    If dbinput.VaporPressureAntoineA = 0 Then
'        dbinput.VaporPressureAntoineA = -1
'        dbinput.VaporPressureDatabaseEquation = -1
'    End If
'
'    If dbinput.LiquidDensityEquation = 0 Then
'        dbinput.LiquidDensityEquation = -1
'    End If
'
'    If dbinput.VaporPressureAntoineA = -1 Then
'
'        '
'        ' OPEN RECORDSET.
'        '
'        Set RS_Main = DB_Main.OpenRecordset( _
'            "SELECT * FROM [VP Yaws] WHERE [VP Yaws].CAS = " & _
'            Format$(dbinput.CASNumber, "0"))
'        RS_Main.MoveFirst
'        RS_Main.MoveLast
'        RS_Main.MoveFirst
'        Set Selection = RS_Main
'        'Data1.RecordSource = "SELECT * FROM [VP Yaws] WHERE [VP Yaws].CAS = " & Format$(dbinput.CASNumber, "0")
'        'Data1.Refresh
'        'Set Selection = Data1.Recordset
'
'        If Selection.EOF = False Then
'
'            dbinput.VaporPressureNumberCoefficients = 3
'            dbinput.VaporPressureAntoineA = Selection("ANTA")
'            dbinput.VaporPressureAntoineB = Selection("ANTB")
'            dbinput.VaporPressureAntoineC = Selection("ANTC")
'            dbinput.VaporPressureMinimumT = Selection("MINT")
'            dbinput.VaporPressureMaximumT = Selection("MAXT")
'            dbinput.VaporPressureSource = get_source(nullcheck("YAWS"))
'
'        Else
'
'            dbinput.VaporPressureAntoineA = -1
'            dbinput.VaporPressureDatabaseEquation = -1
'
'        End If
'
'    End If
'
'    If dbinput.VaporPressureAntoineA = 0 Then
'        dbinput.VaporPressureAntoineA = -1
'        dbinput.VaporPressureDatabaseEquation = -1
'    End If
'
'    If dbinput.VaporPressureAntoineA = -1 Then
'
'        '
'        ' OPEN RECORDSET.
'        '
'        Set RS_Main = DB_Main.OpenRecordset( _
'            "SELECT * FROM [VP@25 Superfund] WHERE [VP@25 Superfund].CAS = " & _
'            Format$(dbinput.CASNumber, "0"))
'        RS_Main.MoveFirst
'        RS_Main.MoveLast
'        RS_Main.MoveFirst
'        Set Selection = RS_Main
'        'Data1.RecordSource = "SELECT * FROM [VP@25 Superfund] WHERE [VP@25 Superfund].CAS = " & Format$(dbinput.CASNumber, "0")
'        'Data1.Refresh
'        'Set Selection = Data1.Recordset
'
'        If Selection.EOF = False Then
'            dbinput.VaporPressureSuperfund = Selection("VP")
'            dbinput.VaporPressureSuperfundTemperature = 25
'            dbinput.VaporPressureSource = get_source(nullcheck("SUPERFUND"))
'        Else
'            dbinput.VaporPressureSuperfund = -1
'        End If
'
'    End If
'
'    If dbinput.VaporPressureSuperfund = 0 Then
'        dbinput.VaporPressureSuperfund = -1
'    End If
'
'    '
'    ' OPEN RECORDSET.
'    '
'    Set RS_Main = DB_Main.OpenRecordset( _
'        "SELECT * FROM [SB@25 Yaws] WHERE [SB@25 Yaws].CAS = " & _
'        Format$(dbinput.CASNumber, "0"))
'    RS_Main.MoveFirst
'    RS_Main.MoveLast
'    RS_Main.MoveFirst
'    Set Selection = RS_Main
'    'Data1.RecordSource = "SELECT * FROM [SB@25 Yaws] WHERE [SB@25 Yaws].CAS = " & Format$(dbinput.CASNumber, "0")
'    'Data1.Refresh
'    'Set Selection = Data1.Recordset
'
'    If Selection.EOF = False Then
'        dbinput.AqueousSolubility = Selection("Sol")
'        dbinput.AqueousSolubilityTemperature = 25
'        dbinput.AqueousSolubilitySource = get_source(nullcheck("YAWS"))
'    Else
'        dbinput.AqueousSolubility = -1
'    End If
'
'    If dbinput.AqueousSolubility = 0 Then
'
'        '
'        ' OPEN RECORDSET.
'        '
'        Set RS_Main = DB_Main.OpenRecordset( _
'            "SELECT * FROM [SB@25 Superfund] WHERE [SB@25 Superfund].CAS = " & _
'            Format$(dbinput.CASNumber, "0"))
'        RS_Main.MoveFirst
'        RS_Main.MoveLast
'        RS_Main.MoveFirst
'        Set Selection = RS_Main
'        'Data1.RecordSource = "SELECT * FROM [SB@25 Superfund] WHERE [SB@25 Superfund].CAS = " & Format$(dbinput.CASNumber, "0")
'        'Data1.Refresh
'        'Set Selection = Data1.Recordset
'
'        If Selection.EOF = False Then
'            dbinput.AqueousSolubility = Selection("Sol")
'            dbinput.AqueousSolubilityTemperature = 25
'            dbinput.AqueousSolubilitySource = get_source(nullcheck("SUPERFUND"))
'
'        Else
'
'            dbinput.AqueousSolubility = -1
'
'        End If
'
'    End If
'
'    If dbinput.AqueousSolubility = 0 Then
'        dbinput.AqueousSolubility = -1
'    End If
'
'    '
'    ' OPEN RECORDSET.
'    '
'    Set RS_Main = DB_Main.OpenRecordset( _
'        "SELECT * FROM [Kow@25 Superfund] WHERE [Kow@25 Superfund].CAS = " & _
'        Format$(dbinput.CASNumber, "0"))
'    RS_Main.MoveFirst
'    RS_Main.MoveLast
'    RS_Main.MoveFirst
'    Set Selection = RS_Main
'    'Data1.RecordSource = "SELECT * FROM [Kow@25 Superfund] WHERE [Kow@25 Superfund].CAS = " & Format$(dbinput.CASNumber, "0")
'    'Data1.Refresh
'    'Set Selection = Data1.Recordset
'
'    If Selection.EOF = False Then
'        dbinput.OctWaterPartCoeff = Selection("log Kow")
'        dbinput.OctWaterPartCoeffTemperature = 25
'        dbinput.OctWaterPartCoeffSource = get_source(nullcheck("SUPERFUND"))
'    Else
'        dbinput.OctWaterPartCoeff = -1
'    End If
'
'    If dbinput.OctWaterPartCoeff = 0 Then
'        dbinput.OctWaterPartCoeff = -1
'    End If
'
'    '
'    ' OPEN RECORDSET.
'    '
'    Set RS_Main = DB_Main.OpenRecordset( _
'        "SELECT * FROM [Rogers/Miller] WHERE [Rogers/Miller].CAS = " & _
'        Format$(dbinput.CASNumber, "0"))
'    RS_Main.MoveFirst
'    RS_Main.MoveLast
'    RS_Main.MoveFirst
'    Set Selection = RS_Main
'    'Data1.RecordSource = "SELECT * FROM [Rogers/Miller] WHERE [Rogers/Miller].CAS = " & Format$(dbinput.CASNumber, "0")
'    'Data1.Refresh
'    'Set Selection = Data1.Recordset
'
'    If Selection.EOF = False Then
'
'        If Selection("MX") <= 0 Then dbinput.MaximumUnifacGroups = 0
'
'        For i = 1 To NC
'            For j = 1 To 10
'                For K = 1 To 2
'                    dbinput.MS(i, j, K) = 0
'                Next K
'            Next j
'        Next i
'
'        dbinput.NumberofRingsinCompound = Selection("RG")
'        dbinput.MaximumUnifacGroups = Selection("MX")
'
'        For i = 1 To dbinput.MaximumUnifacGroups
'            dbinput.MS(NC, i, 1) = Selection("G" + Trim$(Str$(i)))
'            dbinput.MS(NC, i, 2) = Selection("N" + Trim$(Str$(i)))
'        Next i
'
'    Else
'
'       dbinput.NumberofRingsinCompound = -1
'       dbinput.MaximumUnifacGroups = -1
'
'    End If
'
'    If dbinput.formula = "" Then
'        If Selection.EOF = False Then
'            dbinput.formula = Selection("Formula")
'        End If
'    End If
'
'    HC_Count = 0
'    hc_string = ""
'
'    '
'    ' OPEN RECORDSET.
'    '
'    Set RS_Main = DB_Main.OpenRecordset( _
'        "SELECT * FROM [HC RTI] WHERE [HC RTI].CAS = " & _
'        Format$(dbinput.CASNumber, "0"))
'    RS_Main.MoveFirst
'    RS_Main.MoveLast
'    RS_Main.MoveFirst
'    Set Selection = RS_Main
'    'Data1.RecordSource = "SELECT * FROM [HC RTI] WHERE [HC RTI].CAS = " & Format$(dbinput.CASNumber, "0")
'    'Data1.Refresh
'    'Set Selection = Data1.Recordset
'
'    Do While Not Selection.EOF
'
'        If number(Selection(1)) <> "" Then
'
'            last_hc_string = hc_string
'            HC_DB_Source = "RTI"
'            hc_form!lblDatabase = HC_DB_Source
'            LSet HC_DB_Value = Format$(number(Selection(1)), GetTheFormat(CDbl(number(Selection(1)))))
'            HC_DB_Temp = Format$(number(Selection(2)), GetTheFormat(CDbl(number(Selection(2)))))
'            hc_string = HC_DB_Value + HC_DB_Temp
'
'            If hc_string <> last_hc_string Then
'                HC_Count = HC_Count + 1
'                dbinput.HenrysConstantSource = get_source(nullcheck("RTI"))
'                dbinput.HenrysConstant(HC_Count) = Selection(1)
'                dbinput.HenrysConstantTemperature(HC_Count) = Selection(2)
'            End If
'
'        End If
'
'        Selection.MoveNext
'
'    Loop
'
'    dbinput.NumberOfDatabaseHenrysConstants = HC_Count
'
'    If dbinput.NumberOfDatabaseHenrysConstants = 0 Then
'
'        '
'        ' OPEN RECORDSET.
'        '
'        Set RS_Main = DB_Main.OpenRecordset( _
'            "SELECT * FROM [HC Superfund] WHERE [HC Superfund].CAS = " & _
'            Format$(dbinput.CASNumber, "0"))
'        RS_Main.MoveFirst
'        RS_Main.MoveLast
'        RS_Main.MoveFirst
'        Set Selection = RS_Main
'        'Data1.RecordSource = "SELECT * FROM [HC Superfund] WHERE [HC Superfund].CAS = " & Format$(dbinput.CASNumber, "0")
'        'Data1.Refresh
'        'Set Selection = Data1.Recordset
'
'        Do While Not Selection.EOF
'
'            If number(Selection(1)) <> "" Then
'
'                last_hc_string = hc_string
'                HC_DB_Source = "SUPERFUND"
'                hc_form!lblDatabase = HC_DB_Source
'                LSet HC_DB_Value = Format$(number(Selection(1)), GetTheFormat(CDbl(number(Selection(1)))))
'                HC_DB_Temp = Format$(number(Selection(2)), GetTheFormat(CDbl(number(Selection(2)))))
'                hc_string = HC_DB_Value + HC_DB_Temp
'
'                If hc_string <> last_hc_string Then
'                    HC_Count = HC_Count + 1
'                    dbinput.HenrysConstantSource = get_source(nullcheck("SUPERFUND"))
'                    dbinput.HenrysConstant(HC_Count) = Selection(1)
'                    dbinput.HenrysConstantTemperature(HC_Count) = Selection(2)
'                End If
'
'            End If
'
'            Selection.MoveNext
'
'        Loop
'
'        dbinput.NumberOfDatabaseHenrysConstants = HC_Count
'
'    End If
'
'    If dbinput.NumberOfDatabaseHenrysConstants = 0 Then
'
'        '
'        ' OPEN RECORDSET.
'        '
'        Set RS_Main = DB_Main.OpenRecordset( _
'            "SELECT * FROM [HC Yaws] WHERE [HC Yaws].CAS = " & _
'            Format$(dbinput.CASNumber, "0"))
'        RS_Main.MoveFirst
'        RS_Main.MoveLast
'        RS_Main.MoveFirst
'        Set Selection = RS_Main
'        'Data1.RecordSource = "SELECT * FROM [HC Yaws] WHERE [HC Yaws].CAS = " & Format$(dbinput.CASNumber, "0")
'        'Data1.Refresh
'        'Set Selection = Data1.Recordset
'
'        Do While Not Selection.EOF
'
'            If number(Selection(1)) <> "" Then
'                last_hc_string = hc_string
'                HC_DB_Source = "YAWS"
'                hc_form!lblDatabase = HC_DB_Source
'                LSet HC_DB_Value = Format$(number(Selection(1)), GetTheFormat(CDbl(number(Selection(1)))))
'                HC_DB_Temp = Format$(number(Selection(2)), GetTheFormat(CDbl(number(Selection(2)))))
'                hc_string = HC_DB_Value + HC_DB_Temp
'                If hc_string <> last_hc_string Then
'                    HC_Count = HC_Count + 1
'                    dbinput.HenrysConstantSource = get_source(nullcheck("YAWS"))
'                    dbinput.HenrysConstant(HC_Count) = Selection(1)
'                    dbinput.HenrysConstantTemperature(HC_Count) = Selection(2)
'                End If
'
'            End If
'
'        Selection.MoveNext
'
'        Loop
'
'        dbinput.NumberOfDatabaseHenrysConstants = HC_Count
'
'     End If
'
'     'Convert database Henry's constants to dimensionless units
'
'     If dbinput.NumberOfDatabaseHenrysConstants > 0 Then
'         Call HCDBCONV(dbinput.HenrysConstant(1), dbinput.HenrysConstantTemperature(1), dbinput.NumberOfDatabaseHenrysConstants, dbinput.HenrysConstantSource)
'     End If
'
'End Sub


Private Sub Update_P_and_T_StEPPLink()
Dim i As Integer

  '----- Critical that these lines be here!
  'ChDir App.Path
  'ChDrive App.Path
  Call ChangeDir_Main

  If ((SteppLink_SpecifiedPressure <> "") Or (SteppLink_SpecifiedTemperature <> "")) Then
    If (SteppLink_SpecifiedPressure <> "") Then
      Call UpdatePressAllCompounds(CDbl(SteppLink_SpecifiedPressure))
    End If
    If (SteppLink_SpecifiedTemperature <> "") Then
      Call UpdateTempAllCompounds(CDbl(SteppLink_SpecifiedTemperature))
    End If
  End If

    
    'For i = 1 To cboSelectContaminant.ListCount
    '  cboSelectContaminant.ListIndex = i - 1
    '  If (SteppLink_SpecifiedPressure <> "") Then
    '    If (CDbl(Trim$(txtOperatingPressure)) <> CDbl(Trim$(SteppLink_SpecifiedPressure))) Then
    '      phprop.OperatingPressure = CDbl(SteppLink_SpecifiedPressure)
    '      txtOperatingPressure = SteppLink_SpecifiedPressure
    '      Call UpdatePressCurrentCompound
    '    End If
    '  End If
    '  If (SteppLink_SpecifiedTemperature <> "") Then
    '    If (CDbl(Trim$(txtOperatingTemperature)) <> CDbl(Trim$(SteppLink_SpecifiedTemperature))) Then
    '      phprop.OperatingTemperature = CDbl(SteppLink_SpecifiedTemperature)
    '      txtOperatingTemperature = SteppLink_SpecifiedTemperature
    '      Call UpdateTempCurrentCompound
    '    End If
    '  End If
    'Next i
  
End Sub

Private Sub update_window()


End Sub

Private Sub UpdatePressAllCompounds(NewPressure As Double)
    Dim ii As Integer, origIndex As Integer, i As Integer

      frmWaitForCalculations!Panel3D1.FontSize = 13.8
      frmWaitForCalculations!Panel3D1.Caption = "Performing Calculations" & Chr$(13) & Chr$(13) & Chr$(13) & Chr$(13) & "Please Wait"
      frmWaitForCalculations!Panel3D2.FontSize = 10#
      frmWaitForCalculations!Panel3D2.Visible = True
      frmWaitForCalculations!Panel3D2.Caption = "Updating" & Chr$(13) & "Each Selected Compound"
      frmWaitForCalculations.Show
      frmWaitForCalculations.Refresh

      Screen.MousePointer = 11   'Hourglass

    origIndex = cboSelectContaminant.ListIndex

    For ii = 0 To cboSelectContaminant.ListCount - 1
    cboSelectContaminant.ListIndex = -1
    cboSelectContaminant.ListIndex = ii
    phprop.OperatingPressure = NewPressure
    txtOperatingPressure = Format$(NewPressure, "0.00")


          Call BlankTextBoxesPressure

          Call CalculateGasDiffusivity
          contam_prop_form.Refresh


'*** Place PROPAVAILABLE and HAVEPROPERTY arrays into phprop structure
       For i = 1 To NUMBER_OF_PROPERTIES_AVAILABLE
           phprop.PROPAVAILABLE(i) = PROPAVAILABLE(i)
       Next i
       For i = 1 To NUMBER_OF_PROPERTIES
           phprop.HaveProperty(i) = HaveProperty(i)
       Next i
     
      NumContaminantInList = cboSelectContaminant.ListIndex + 1
      PropContaminant(NumContaminantInList) = phprop
    
    Next ii

    cboSelectContaminant.ListIndex = origIndex

      frmWaitForCalculations.Hide

      frmWaitForCalculations!Panel3D1.FontSize = 13.8
      frmWaitForCalculations!Panel3D1.Caption = "Performing Calculations" & Chr$(13) & Chr$(13) & "Please Wait"
      frmWaitForCalculations!Panel3D2.Visible = False

      Screen.MousePointer = 0    'Arrow

End Sub

Private Sub UpdatePressCurrentCompound()
    Dim i As Integer

'NOTE: THIS SUBROUTINE IS NO LONGER USED ANYWHERE ELSE
'IN THE PROGRAM !!!
Exit Sub
          



          frmWaitForCalculations.Show
          frmWaitForCalculations.Refresh

          Call BlankTextBoxesPressure

          Screen.MousePointer = 11   'Hourglass
          Call CalculateGasDiffusivity
          Screen.MousePointer = 0    'Arrow
          contam_prop_form.Refresh

          frmWaitForCalculations.Hide

'*** Place PROPAVAILABLE and HAVEPROPERTY arrays into phprop structure
     For i = 1 To NUMBER_OF_PROPERTIES_AVAILABLE
         phprop.PROPAVAILABLE(i) = PROPAVAILABLE(i)
     Next i
     For i = 1 To NUMBER_OF_PROPERTIES
         phprop.HaveProperty(i) = HaveProperty(i)
     Next i

          NumContaminantInList = cboSelectContaminant.ListIndex + 1
          PropContaminant(NumContaminantInList) = phprop

End Sub

Private Sub UpdateTempAllCompounds(NewTemperature As Double)
Dim ii As Integer, origIndex As Integer, i As Integer
Dim SIValue As Double
Dim EnglishValue As Double

      frmWaitForCalculations!Panel3D1.FontSize = 13.8
      frmWaitForCalculations!Panel3D1.Caption = "Performing Calculations" & Chr$(13) & Chr$(13) & Chr$(13) & Chr$(13) & "Please Wait"
      frmWaitForCalculations!Panel3D2.FontSize = 10#
      frmWaitForCalculations!Panel3D2.Visible = True
      frmWaitForCalculations!Panel3D2.Caption = "Updating" & Chr$(13) & "Each Selected Compound"
      frmWaitForCalculations.Show
      frmWaitForCalculations.Refresh

      Screen.MousePointer = 11   'Hourglass

    origIndex = cboSelectContaminant.ListIndex

    For ii = 0 To cboSelectContaminant.ListCount - 1
    cboSelectContaminant.ListIndex = -1
    cboSelectContaminant.ListIndex = ii
    phprop.OperatingTemperature = NewTemperature
    
    If CurrentUnits = SIUnits Then
      'DISPLAY TEMPERATURE IN SI UNITS.
      SIValue = NewTemperature
      txtOperatingTemperature.Text = Format$(SIValue, "0.00")
    Else
      'DISPLAY TEMPERATURE IN ENGLISH UNITS.
      SIValue = NewTemperature
      Call TEMPCNV(EnglishValue, SIValue)
      txtOperatingTemperature.Text = Format$(EnglishValue, "0.00")
    End If

      Call BlankTextBoxesTemp

      Call CalculateVaporPressure
      contam_prop_form.Refresh

      If phprop.ActivityCoefficient.BinaryInteractionParameterDatabase > 0 Then
         Call CalculateActivityCoefficient
         contam_prop_form.Refresh
      End If
      
      Call CalculateHenrysConstant
      contam_prop_form.Refresh

      Call CalculateLiquidDensity
      contam_prop_form.Refresh

      Call CalculateMolarVolumeOpT
      contam_prop_form.Refresh

      If phprop.AqueousSolubility.BinaryInteractionParameterDatabase > 0 Then
         Call CalculateAqueousSolubility
         contam_prop_form.Refresh
      End If

      If phprop.OctWaterPartCoeff.BinaryInteractionParameterDatabase > 0 Then
         Call CalculateOctWaterPartCoeff
         contam_prop_form.Refresh
      End If

      Call CalculateLiquidDiffusivity
      contam_prop_form.Refresh

      Call CalculateGasDiffusivity
      contam_prop_form.Refresh

      Call CalculateWaterDensity
      contam_prop_form.Refresh

      Call CalculateWaterViscosity
      contam_prop_form.Refresh

      Call CalculateWaterSurfaceTension
      contam_prop_form.Refresh

      Call CalculateAirDensity
      contam_prop_form.Refresh

      Call CalculateAirViscosity
      contam_prop_form.Refresh


'*** Place PROPAVAILABLE and HAVEPROPERTY arrays into phprop structure
       For i = 1 To NUMBER_OF_PROPERTIES_AVAILABLE
           phprop.PROPAVAILABLE(i) = PROPAVAILABLE(i)
       Next i
       For i = 1 To NUMBER_OF_PROPERTIES
           phprop.HaveProperty(i) = HaveProperty(i)
       Next i
     
      NumContaminantInList = cboSelectContaminant.ListIndex + 1
      PropContaminant(NumContaminantInList) = phprop
    
    Next ii

    cboSelectContaminant.ListIndex = origIndex

      frmWaitForCalculations.Hide

      frmWaitForCalculations!Panel3D1.FontSize = 13.8
      frmWaitForCalculations!Panel3D1.Caption = "Performing Calculations" & Chr$(13) & Chr$(13) & "Please Wait"
      frmWaitForCalculations!Panel3D2.Visible = False

      Screen.MousePointer = 0    'Arrow

End Sub

Private Sub UpdateTempCurrentCompound()
    Dim i As Integer


'NOTE: THIS SUBROUTINE IS NO LONGER USED ANYWHERE ELSE
'IN THE PROGRAM !!!
Exit Sub

          
          
          frmWaitForCalculations.Show
          frmWaitForCalculations.Refresh

          Screen.MousePointer = 11   'Hourglass

          Call BlankTextBoxesTemp

          Call CalculateVaporPressure
          contam_prop_form.Refresh

          If phprop.ActivityCoefficient.BinaryInteractionParameterDatabase > 0 Then
             Call CalculateActivityCoefficient
             contam_prop_form.Refresh
          End If
          
          Call CalculateHenrysConstant
          contam_prop_form.Refresh

          Call CalculateLiquidDensity
          contam_prop_form.Refresh

          Call CalculateMolarVolumeOpT
          contam_prop_form.Refresh

          If phprop.AqueousSolubility.BinaryInteractionParameterDatabase > 0 Then
             Call CalculateAqueousSolubility
             contam_prop_form.Refresh
          End If

          If phprop.OctWaterPartCoeff.BinaryInteractionParameterDatabase > 0 Then
             Call CalculateOctWaterPartCoeff
             contam_prop_form.Refresh
          End If

          Call CalculateLiquidDiffusivity
          contam_prop_form.Refresh

          Call CalculateGasDiffusivity
          contam_prop_form.Refresh

          Call CalculateWaterDensity
          contam_prop_form.Refresh

          Call CalculateWaterViscosity
          contam_prop_form.Refresh

          Call CalculateWaterSurfaceTension
          contam_prop_form.Refresh

          Call CalculateAirDensity
          contam_prop_form.Refresh

          Call CalculateAirViscosity
          contam_prop_form.Refresh

          frmWaitForCalculations.Hide

          Screen.MousePointer = 0    'Arrow

'*** Place PROPAVAILABLE and HAVEPROPERTY arrays into phprop structure
           For i = 1 To NUMBER_OF_PROPERTIES_AVAILABLE
               phprop.PROPAVAILABLE(i) = PROPAVAILABLE(i)
           Next i
           For i = 1 To NUMBER_OF_PROPERTIES
               phprop.HaveProperty(i) = HaveProperty(i)
           Next i
     
          NumContaminantInList = cboSelectContaminant.ListIndex + 1
          PropContaminant(NumContaminantInList) = phprop

End Sub


