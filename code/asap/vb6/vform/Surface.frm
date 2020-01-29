VERSION 5.00
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "THREED32.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmSurface 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Surface Aeration"
   ClientHeight    =   6840
   ClientLeft      =   2325
   ClientTop       =   1575
   ClientWidth     =   9480
   Icon            =   "Surface.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   ScaleHeight     =   6840
   ScaleWidth      =   9480
   Begin VB.TextBox txtFlowParameters 
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
      Left            =   6870
      TabIndex        =   6
      Top             =   90
      Width           =   1215
   End
   Begin VB.ComboBox UnitsFlowParam 
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
      Left            =   8130
      Style           =   2  'Dropdown List
      TabIndex        =   57
      TabStop         =   0   'False
      Top             =   90
      Width           =   1155
   End
   Begin VB.TextBox txtPowerInput 
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
      Left            =   1830
      TabIndex        =   2
      Top             =   1470
      Width           =   1215
   End
   Begin VB.ComboBox UnitsPowerInput 
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
      Left            =   3090
      Style           =   2  'Dropdown List
      TabIndex        =   25
      TabStop         =   0   'False
      Top             =   1470
      Width           =   1275
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
      Left            =   1080
      TabIndex        =   19
      TabStop         =   0   'False
      Top             =   6120
      Width           =   2655
   End
   Begin Threed.SSFrame fraOperatingConditions 
      Height          =   1275
      Left            =   60
      TabIndex        =   13
      Top             =   120
      Width           =   4395
      _Version        =   65536
      _ExtentX        =   7752
      _ExtentY        =   2249
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
         Caption         =   "Display Physical Properties of Water"
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
         Left            =   150
         TabIndex        =   22
         TabStop         =   0   'False
         Top             =   900
         Width           =   4155
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
         Left            =   1770
         TabIndex        =   1
         Top             =   540
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
         Left            =   1770
         TabIndex        =   0
         Top             =   240
         Width           =   1215
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
         Left            =   3030
         Style           =   2  'Dropdown List
         TabIndex        =   21
         TabStop         =   0   'False
         Top             =   240
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
         Left            =   3030
         Style           =   2  'Dropdown List
         TabIndex        =   20
         TabStop         =   0   'False
         Top             =   540
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
         Left            =   -990
         TabIndex        =   24
         Top             =   540
         Width           =   2655
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
         Left            =   -990
         TabIndex        =   23
         Top             =   240
         Width           =   2655
      End
   End
   Begin Threed.SSFrame fraOxygen 
      Height          =   1395
      Left            =   60
      TabIndex        =   14
      Top             =   1890
      Width           =   4395
      _Version        =   65536
      _ExtentX        =   7752
      _ExtentY        =   2461
      _StockProps     =   14
      Caption         =   "Oxygen (reference compound):"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Begin VB.ComboBox cboOxygen 
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
         Left            =   1350
         Style           =   2  'Dropdown List
         TabIndex        =   29
         TabStop         =   0   'False
         Top             =   300
         Width           =   2955
      End
      Begin VB.TextBox txtOxygen 
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
         Left            =   1770
         TabIndex        =   4
         Top             =   1020
         Width           =   1215
      End
      Begin VB.TextBox txtOxygen 
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
         Left            =   1770
         TabIndex        =   3
         Top             =   660
         Width           =   1215
      End
      Begin VB.ComboBox UnitsOxygenRef 
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
         Left            =   3030
         Style           =   2  'Dropdown List
         TabIndex        =   28
         TabStop         =   0   'False
         Top             =   660
         Width           =   1275
      End
      Begin VB.ComboBox UnitsOxygenRef 
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
         Left            =   3030
         Style           =   2  'Dropdown List
         TabIndex        =   27
         TabStop         =   0   'False
         Top             =   1020
         Width           =   1275
      End
      Begin VB.Label lblOxygenLabel 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "KLa Method:"
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
         Left            =   30
         TabIndex        =   32
         Top             =   360
         Width           =   1215
      End
      Begin VB.Label lblOxygenLabel 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "KLa"
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
         Left            =   -990
         TabIndex        =   31
         Top             =   1020
         Width           =   2655
      End
      Begin VB.Label lblOxygenLabel 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
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
         Index           =   1
         Left            =   -990
         TabIndex        =   30
         Top             =   660
         Width           =   2655
      End
   End
   Begin Threed.SSFrame fraContaminantInformation 
      Height          =   2655
      Left            =   60
      TabIndex        =   15
      Top             =   3360
      Width           =   4395
      _Version        =   65536
      _ExtentX        =   7752
      _ExtentY        =   4683
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
         Index           =   3
         Left            =   1890
         TabIndex        =   5
         Top             =   2280
         Width           =   1095
      End
      Begin VB.ComboBox cboDesignContaminant 
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
         Left            =   210
         Style           =   2  'Dropdown List
         TabIndex        =   39
         TabStop         =   0   'False
         Top             =   480
         Width           =   3975
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
         Left            =   210
         TabIndex        =   38
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
         Left            =   1350
         TabIndex        =   37
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
         Left            =   2430
         TabIndex        =   36
         TabStop         =   0   'False
         Top             =   840
         Width           =   1755
      End
      Begin VB.ComboBox UnitsDesignContam 
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
         Left            =   3030
         Style           =   2  'Dropdown List
         TabIndex        =   35
         TabStop         =   0   'False
         Top             =   1380
         Width           =   1275
      End
      Begin VB.ComboBox UnitsDesignContam 
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
         Left            =   3030
         Style           =   2  'Dropdown List
         TabIndex        =   34
         TabStop         =   0   'False
         Top             =   1680
         Width           =   1275
      End
      Begin VB.ComboBox UnitsDesignContam 
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
         Left            =   3030
         Style           =   2  'Dropdown List
         TabIndex        =   33
         TabStop         =   0   'False
         Top             =   2280
         Width           =   1275
      End
      Begin VB.Label lblDesignConcentration 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "KLa"
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
         Left            =   -750
         TabIndex        =   47
         Top             =   2280
         Width           =   2535
      End
      Begin VB.Label lblDesignConcentration 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Influent Conc."
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
         Left            =   -750
         TabIndex        =   46
         Top             =   1380
         Width           =   2535
      End
      Begin VB.Label lblDesignConcentration 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Treatment Obj."
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
         Left            =   -750
         TabIndex        =   45
         Top             =   1680
         Width           =   2535
      End
      Begin VB.Label lblDesignConcentration 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Desired % Removal"
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
         Left            =   -750
         TabIndex        =   44
         Top             =   1980
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
         Left            =   1890
         TabIndex        =   43
         Top             =   1380
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
         Index           =   1
         Left            =   1890
         TabIndex        =   42
         Top             =   1680
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
         Index           =   2
         Left            =   1890
         TabIndex        =   41
         Top             =   1980
         Width           =   1095
      End
      Begin VB.Shape Shape1 
         Height          =   795
         Left            =   90
         Top             =   420
         Width           =   4215
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
         Index           =   0
         Left            =   3030
         TabIndex        =   40
         Top             =   1980
         Width           =   1155
      End
   End
   Begin Threed.SSFrame fraTankParameters 
      Height          =   1815
      Left            =   4440
      TabIndex        =   16
      Top             =   510
      Width           =   4935
      _Version        =   65536
      _ExtentX        =   8705
      _ExtentY        =   3201
      _StockProps     =   14
      Caption         =   "Tank Parameters"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Begin VB.TextBox txtTankParameters 
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
         Left            =   2430
         TabIndex        =   7
         Top             =   240
         Width           =   1215
      End
      Begin VB.TextBox txtTankParameters 
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
         Left            =   2430
         TabIndex        =   8
         Top             =   540
         Width           =   1215
      End
      Begin VB.TextBox txtTankParameters 
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
         Left            =   2430
         TabIndex        =   9
         Top             =   840
         Width           =   1215
      End
      Begin VB.TextBox txtTankParameters 
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
         Left            =   2430
         TabIndex        =   10
         Top             =   1140
         Width           =   1215
      End
      Begin VB.TextBox txtTankParameters 
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
         Left            =   2430
         TabIndex        =   11
         Top             =   1440
         Width           =   1215
      End
      Begin VB.ComboBox UnitsTankParam 
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
         Left            =   3690
         Style           =   2  'Dropdown List
         TabIndex        =   51
         TabStop         =   0   'False
         Top             =   540
         Width           =   1155
      End
      Begin VB.ComboBox UnitsTankParam 
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
         Left            =   3690
         Style           =   2  'Dropdown List
         TabIndex        =   50
         TabStop         =   0   'False
         Top             =   840
         Width           =   1155
      End
      Begin VB.ComboBox UnitsTankParam 
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
         Left            =   3690
         Style           =   2  'Dropdown List
         TabIndex        =   49
         TabStop         =   0   'False
         Top             =   1140
         Width           =   1155
      End
      Begin VB.ComboBox UnitsTankParam 
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
         Left            =   3690
         Style           =   2  'Dropdown List
         TabIndex        =   48
         TabStop         =   0   'False
         Top             =   1440
         Width           =   1155
      End
      Begin VB.Label lblTankParametersLabel 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "No. of Tanks (series)"
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
         Height          =   195
         Index           =   0
         Left            =   -390
         TabIndex        =   56
         Top             =   300
         Width           =   2715
      End
      Begin VB.Label lblTankParametersLabel 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Retention Time (1 Tank)"
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
         Height          =   195
         Index           =   1
         Left            =   -390
         TabIndex        =   55
         Top             =   600
         Width           =   2715
      End
      Begin VB.Label lblTankParametersLabel 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Retention Time (All Tanks)"
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
         Height          =   195
         Index           =   2
         Left            =   -390
         TabIndex        =   54
         Top             =   900
         Width           =   2715
      End
      Begin VB.Label lblTankParametersLabel 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Volume (1 Tank)"
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
         Height          =   195
         Index           =   3
         Left            =   -390
         TabIndex        =   53
         Top             =   1200
         Width           =   2715
      End
      Begin VB.Label lblTankParametersLabel 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Volume (All Tanks)"
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
         Height          =   195
         Index           =   4
         Left            =   -390
         TabIndex        =   52
         Top             =   1500
         Width           =   2715
      End
   End
   Begin Threed.SSFrame fraConcentrationResults 
      Height          =   2115
      Left            =   4440
      TabIndex        =   17
      Top             =   2400
      Width           =   4935
      _Version        =   65536
      _ExtentX        =   8705
      _ExtentY        =   3731
      _StockProps     =   14
      Caption         =   "Concentration Results:"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Begin VB.CommandButton cmdConcentrationResults 
         Appearance      =   0  'Flat
         Caption         =   "View Effluent Concentrations from All Tanks"
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
         Left            =   90
         TabIndex        =   61
         TabStop         =   0   'False
         Top             =   990
         Width           =   4755
      End
      Begin VB.ComboBox UnitsConcResults 
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
         Left            =   3690
         Style           =   2  'Dropdown List
         TabIndex        =   60
         TabStop         =   0   'False
         Top             =   630
         Width           =   1155
      End
      Begin VB.ComboBox UnitsConcResults 
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
         Left            =   3690
         Style           =   2  'Dropdown List
         TabIndex        =   59
         TabStop         =   0   'False
         Top             =   1410
         Width           =   1155
      End
      Begin VB.Label lblConcentrationResultsLabel 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Compound"
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
         Height          =   195
         Index           =   0
         Left            =   -150
         TabIndex        =   70
         Top             =   330
         Width           =   1155
      End
      Begin VB.Label lblConcentrationResultsLabel 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Ci to Tank 1"
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
         Height          =   195
         Index           =   1
         Left            =   -270
         TabIndex        =   69
         Top             =   630
         Width           =   2595
      End
      Begin VB.Label lblConcentrationResultsLabel 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Ce from Last Tank"
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
         Height          =   195
         Index           =   3
         Left            =   -270
         TabIndex        =   68
         Top             =   1470
         Width           =   2595
      End
      Begin VB.Label lblConcentrationResultsLabel 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Achieved % Removal"
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
         Height          =   195
         Index           =   4
         Left            =   -270
         TabIndex        =   67
         Top             =   1830
         Width           =   2595
      End
      Begin VB.Label lblConcentrationResults 
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
         Left            =   1170
         TabIndex        =   66
         Top             =   330
         Width           =   2715
      End
      Begin VB.Label lblConcentrationResults 
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
         Left            =   2430
         TabIndex        =   65
         Top             =   630
         Width           =   1215
      End
      Begin VB.Label lblConcentrationResults 
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
         Left            =   2430
         TabIndex        =   64
         Top             =   1410
         Width           =   1215
      End
      Begin VB.Label lblConcentrationResults 
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
         Left            =   2430
         TabIndex        =   63
         Top             =   1770
         Width           =   1215
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
         Left            =   3690
         TabIndex        =   62
         Top             =   1770
         Width           =   1155
      End
   End
   Begin Threed.SSFrame fraPower 
      Height          =   1275
      Left            =   4440
      TabIndex        =   18
      Top             =   4590
      Width           =   4935
      _Version        =   65536
      _ExtentX        =   8705
      _ExtentY        =   2249
      _StockProps     =   14
      Caption         =   "Power Calculation:"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Begin VB.TextBox txtPowerCalculation 
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
         Left            =   2430
         TabIndex        =   12
         Top             =   240
         Width           =   1215
      End
      Begin VB.ComboBox UnitsPowerCalc 
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
         Left            =   3690
         Style           =   2  'Dropdown List
         TabIndex        =   72
         TabStop         =   0   'False
         Top             =   600
         Width           =   1155
      End
      Begin VB.ComboBox UnitsPowerCalc 
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
         Left            =   3690
         Style           =   2  'Dropdown List
         TabIndex        =   71
         TabStop         =   0   'False
         Top             =   900
         Width           =   1155
      End
      Begin VB.Label lblPowerCalculationLabel 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Aerator Motor Efficiency"
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
         Height          =   195
         Index           =   0
         Left            =   -390
         TabIndex        =   78
         Top             =   300
         Width           =   2715
      End
      Begin VB.Label lblPowerCalculationLabel 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Power Required per Tank"
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
         Height          =   195
         Index           =   1
         Left            =   -390
         TabIndex        =   77
         Top             =   660
         Width           =   2715
      End
      Begin VB.Label lblPowerCalculationLabel 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Total Power Required"
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
         Height          =   195
         Index           =   2
         Left            =   -390
         TabIndex        =   76
         Top             =   960
         Width           =   2715
      End
      Begin VB.Label lblPowerCalculation 
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
         Left            =   2430
         TabIndex        =   75
         Top             =   600
         Width           =   1215
      End
      Begin VB.Label lblPowerCalculation 
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
         Left            =   2430
         TabIndex        =   74
         Top             =   900
         Width           =   1215
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
         Index           =   1
         Left            =   3690
         TabIndex        =   73
         Top             =   240
         Width           =   1155
      End
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   360
      Top             =   6150
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Label lblFlowParametersLabel 
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
      Left            =   4050
      TabIndex        =   58
      Top             =   150
      Width           =   2715
   End
   Begin VB.Label lblPowerInputLabel 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Power Input, P/V"
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
      Height          =   195
      Left            =   -990
      TabIndex        =   26
      Top             =   1530
      Width           =   2715
   End
   Begin VB.Menu mnuFileMenu 
      Caption         =   "&File"
      Begin VB.Menu mnuFile 
         Caption         =   "Switch Modes"
         Index           =   0
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
   Begin VB.Menu mnuOptionsMenu 
      Caption         =   "&Results"
      Begin VB.Menu mnuOptions 
         Caption         =   "&View All Concentration Results"
         Index           =   0
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
Attribute VB_Name = "frmSurface"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim Temp_Text As String
Public frmSurface_Okay_To_Unload As Boolean



Const frmSurface_declarations_end = True


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
    fraTankParameters.Caption = "* DEMONSTRATION VERSION *"
    fraTankParameters.ForeColor = QBColor(12)
  End If
End Sub


Private Sub AddPrompt(menuID As Integer, prompt As String)
    menuPrompts(iMenuPrompts).menuID = menuID
    menuPrompts(iMenuPrompts).prompt = prompt
    iMenuPrompts = iMenuPrompts + 1
End Sub

Private Sub cboDesignContaminant_Click()
Dim ContaminantIndex As Integer, i As Integer
Dim PercentRemoval As Double
Dim Dummy As Double

    ContaminantIndex = cboDesignContaminant.ListIndex + 1
    i = ContaminantIndex
    If i = 0 Then Exit Sub
    sur.DesignContaminant = sur.Contaminant(i)

    'Update Design Contaminant | Influent Conc.
    'lblDesignConcentrationValue(0).Caption = Format$(sur.Contaminant(i).Influent.Value, "0.0")
    Call UnitsDesignContam_Click(0)
    ''''Dummy = sur.Contaminant(i).Influent.Value
    ''''Call Unitted_UnitChange(UNITS_CONCENTRATION, Dummy, UnitsDesignContam(0), lblDesignConcentrationValue(0))
    
    'Update Design Contaminant | Treatment Obj.
    'lblDesignConcentrationValue(1).Caption = Format$(sur.Contaminant(i).TreatmentObjective.Value, "0.0")
    Call UnitsDesignContam_Click(1)
    ''''Dummy = sur.Contaminant(i).TreatmentObjective.Value
    ''''Call Unitted_UnitChange(UNITS_CONCENTRATION, Dummy, UnitsDesignContam(1), lblDesignConcentrationValue(1))
    
    'Call the FORTRAN routine.
    Call REMOVBUB(sur.DesiredPercentRemoval, sur.Contaminant(i).Influent.value, sur.Contaminant(i).TreatmentObjective.value)
    
    lblDesignConcentrationValue(2).Caption = Format$(sur.DesiredPercentRemoval, "0.0")
    
    lblConcentrationResults(0).Caption = cboDesignContaminant.Text
    
    'Update Concentration Results | Ci to Tank 1.
    'lblConcentrationResults(1).Caption = lblDesignConcentrationValue(0).Caption
    Call UnitsConcResults_Click(1)
    
    'Update Concentration Results | Ce from Last Tank.
    Call UnitsConcResults_Click(3)
    'UPDATED_UNITS

    Call CalculateContaminantMTCoeffSurface
    
    'Update Concentration Results | KLa.
    Call UnitsDesignContam_Click(3)
    
    If SurfaceAerationMode = DESIGN_MODE Then
       Call CalculateRetentionTimeSurface
    End If
    Call CalculateTausAndTankVolumesSurface
    Call CalculateEffluentConcentrationsSurface
    Call CalculatePowerSurface

    'Update Tank Parameters:
    For i = 1 To 4
      Call UnitsTankParam_Click(i)
    Next i

End Sub

Private Sub cboDesignContaminant_GotFocus()
    Temp_Text = cboDesignContaminant.Text
End Sub

Private Sub cboDesignContaminant_KeyPress(KeyAscii As Integer)
  Call TextBoxNumber_KeyPress(KeyAscii)
End Sub

Private Sub cboDesignContaminant_LostFocus()
  'Call TextBoxNumber_KeyPress(KeyAscii)
End Sub

Private Sub cboOxygen_Click()
   
   If cboOxygen.ListIndex = 0 Then   'Clean Water Oxygen Transfer Test Data
      If sur.Oxygen.KLaMethod = 2 Then
         sur.Oxygen.KLaMethod = 1
         Call CalculateOxygenMTCoeffSurface
      End If
   Else
      sur.Oxygen.KLaMethod = 2
      sur.Oxygen.MassTransferCoefficient.UserInput = True
   End If

End Sub

Private Sub cmdAddComponent_Click()
Dim x As rec_frmContaminantPropertyEdit
Dim i As Integer

  If (sur.NumChemical + 1 > MAXCHEMICAL) Then
     MsgBox "The maximum number of contaminants has been reached.  It is not possible to input more than " & Format$(MAXCHEMICAL, "0") & " contaminants for design.  To add an additional contaminant, another must first be removed.", MB_ICONSTOP, "Surface Aeration"
     cmdAddComponent.Enabled = False
     Exit Sub
  End If

  x.ModelName = "Surface Aeration"
  x.ModelType = MODELTYPE_SURFACE
  x.DoEditNumber = -1       'Will be set by frmContaminantPropertyEdit.
  x.DoAdd = True
  x.OldNumCompo = cboDesignContaminant.ListCount
  For i = 1 To x.OldNumCompo
    x.Contaminants(i).Name = sur.Contaminant(i).Name
    x.Contaminants(i).MolecularWeight.value = sur.Contaminant(i).MolecularWeight.value
    x.Contaminants(i).HenrysConstant.value = sur.Contaminant(i).HenrysConstant.value
    x.Contaminants(i).MolarVolume.value = sur.Contaminant(i).MolarVolume.value
    x.Contaminants(i).LiquidDiffusivity.value = sur.Contaminant(i).LiquidDiffusivity.value
    x.Contaminants(i).Influent.value = sur.Contaminant(i).Influent.value
    x.Contaminants(i).TreatmentObjective.value = sur.Contaminant(i).TreatmentObjective.value
  Next i

  StEPPImportSuccess = False
  Data_frmContaminantPropertyEdit = x
  frmContaminantPropertyEdit.Show 1
  x = Data_frmContaminantPropertyEdit

  If (StEPPImportSuccess) Or (Not x.CancelledAdd) Then
    For i = x.OldNumCompo + 1 To x.NewNumCompo
      
      If i > 10 Then
       MsgBox "Unable to continue importing file as maximum amount of chemicals in memory reached."
       Me.Show
       Exit Sub
      End If
      
      'Incorporate new contaminant.
      sur.Contaminant(i).Name = x.Contaminants(i).Name
      sur.Contaminant(i).MolecularWeight.value = x.Contaminants(i).MolecularWeight.value
      sur.Contaminant(i).HenrysConstant.value = x.Contaminants(i).HenrysConstant.value
      sur.Contaminant(i).MolarVolume.value = x.Contaminants(i).MolarVolume.value
      sur.Contaminant(i).LiquidDiffusivity.value = x.Contaminants(i).LiquidDiffusivity.value
      sur.Contaminant(i).Influent.value = x.Contaminants(i).Influent.value
      sur.Contaminant(i).TreatmentObjective.value = x.Contaminants(i).TreatmentObjective.value
      sur.NumChemical = sur.NumChemical + 1
      'Incorporate new name into ComboBox.
      cboDesignContaminant.AddItem sur.Contaminant(i).Name


    Next i
  End If
  
  Call SetDesignContaminantEnabledSurface(CInt(cboDesignContaminant.ListCount))
  
  If (sur.NumChemical > 0) Then
    cmdDeleteComponent.Enabled = True
    cmdEditComponent.Enabled = True
    If (cboDesignContaminant.ListIndex = -1) Then
      cboDesignContaminant.ListIndex = 0
    End If
  End If
  Me.Show
End Sub

Private Sub cmdConcentrationResults_Click()
    frmSurfaceEffluentConcentrations.Show 1

End Sub

Private Sub cmdDeleteComponent_Click()
  Dim i As Integer

  sur.Chemical = cboDesignContaminant.ListIndex + 1
  If (sur.Chemical = 0) Then Exit Sub

  If MsgBox("Remove" & NL & cboDesignContaminant.List(cboDesignContaminant.ListIndex), 36, "") = IDYES Then
    cboDesignContaminant.RemoveItem cboDesignContaminant.ListIndex
    For i = sur.Chemical To sur.NumChemical - 1
      sur.Contaminant(i) = sur.Contaminant(i + 1)
    Next i
    sur.NumChemical = sur.NumChemical - 1
    If (sur.NumChemical > 0) Then
      cboDesignContaminant.ListIndex = 0
    Else
      cmdDeleteComponent.Enabled = False
      cmdEditComponent.Enabled = False
    End If
    Call SetDesignContaminantEnabledSurface(CInt(cboDesignContaminant.ListCount))
  End If

  If sur.NumChemical < 10 Then cmdAddComponent.Enabled = True
  Call LOCAL___Reset_DemoVersionDisablings
End Sub

Private Sub cmdEditComponent_Click()
Dim x As rec_frmContaminantPropertyEdit
Dim i As Integer
Dim AListIndex As Integer

  sur.Chemical = cboDesignContaminant.ListIndex + 1
  If (sur.Chemical = 0) Then Exit Sub

  x.ModelName = "Surface Aeration"
  x.ModelType = MODELTYPE_SURFACE
  x.DoEditNumber = cboDesignContaminant.ListIndex + 1
  x.DoAdd = False
  x.OldNumCompo = cboDesignContaminant.ListCount
  For i = 1 To x.OldNumCompo
    x.Contaminants(i).Name = sur.Contaminant(i).Name
    x.Contaminants(i).MolecularWeight.value = sur.Contaminant(i).MolecularWeight.value
    x.Contaminants(i).HenrysConstant.value = sur.Contaminant(i).HenrysConstant.value
    x.Contaminants(i).MolarVolume.value = sur.Contaminant(i).MolarVolume.value
    x.Contaminants(i).LiquidDiffusivity.value = sur.Contaminant(i).LiquidDiffusivity.value
    x.Contaminants(i).Influent.value = sur.Contaminant(i).Influent.value
    x.Contaminants(i).TreatmentObjective.value = sur.Contaminant(i).TreatmentObjective.value
  Next i

  Data_frmContaminantPropertyEdit = x
  frmContaminantPropertyEdit.Show 1
  x = Data_frmContaminantPropertyEdit

  If (Not x.CancelledEdit) Then
    For i = 1 To x.NewNumCompo
      sur.Contaminant(i).Name = x.Contaminants(i).Name
      sur.Contaminant(i).MolecularWeight.value = x.Contaminants(i).MolecularWeight.value
      sur.Contaminant(i).HenrysConstant.value = x.Contaminants(i).HenrysConstant.value
      sur.Contaminant(i).MolarVolume.value = x.Contaminants(i).MolarVolume.value
      sur.Contaminant(i).LiquidDiffusivity.value = x.Contaminants(i).LiquidDiffusivity.value
      sur.Contaminant(i).Influent.value = x.Contaminants(i).Influent.value
      sur.Contaminant(i).TreatmentObjective.value = x.Contaminants(i).TreatmentObjective.value
    Next i
    If (x.OldNumCompo < x.NewNumCompo) Then
      'Incorporate new names into ComboBox.
      For i = x.OldNumCompo + 1 To x.NewNumCompo
        cboDesignContaminant.AddItem sur.Contaminant(i).Name
      Next i
    End If
    'Update ComboBox for any changed names:
    For i = 1 To x.OldNumCompo
      If (Trim$(cboDesignContaminant.List(i - 1)) <> Trim$(sur.Contaminant(i).Name)) Then
        cboDesignContaminant.List(i - 1) = Trim$(sur.Contaminant(i).Name)
      End If
    Next i
  Else
     Exit Sub
  End If

  'Generate click event on cboDesignContaminant
  AListIndex = cboDesignContaminant.ListIndex
  cboDesignContaminant.ListIndex = -1
  cboDesignContaminant.ListIndex = AListIndex

End Sub

Private Sub cmdSelectContaminants_Click()
    'frmListContaminantSurface.Show 1
End Sub

Private Sub Command1_Click()
Call surface_results
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
'    AddPrompt GetMenuItemID(hSubMenu, 0), "Switch modes for surface aeration"
'    AddPrompt GetMenuItemID(hSubMenu, 2), "Load a design case from a file"
'    AddPrompt GetMenuItemID(hSubMenu, 3), "Save this design case to a file"
'    AddPrompt GetMenuItemID(hSubMenu, 4), "Save this design case to a file"
'    AddPrompt GetMenuItemID(hSubMenu, 6), "Print this design case"
'    AddPrompt GetMenuItemID(hSubMenu, 7), "Select printer for printing results"
'    AddPrompt GetMenuItemID(hSubMenu, 9), "Leave surface aeration and return to main ASAP menu"
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
'    ' Load Options menu prompt
'    '
'    hSubMenu = GetSubMenu(hMenu, 2)
'    AddPrompt hSubMenu, "Options"
'    AddPrompt GetMenuItemID(hSubMenu, 0), "View Effluent Concentration Results for All Contaminants"

    'Initialize last-few-files list.
    Call frmSurface_InitializeLastFewFilesList

End Sub

Private Sub Form_Load()
    Dim i As Integer

    frmSurface_Okay_To_Unload = False
    sur.NumChemical = 0
    frmSurface.WindowState = 0
    frmSurface.Width = SCREEN_WIDTH_STANDARD
    frmSurface.Height = SCREEN_HEIGHT_STANDARD
    'StatusMessagePanel.BackColor = &HC0C0C0
    'StatusMessagePanel.ForeColor = &H400040
    'StatusBarPanel.BackColor = &HC0C0C0
    'StatusBarPanel.ForeColor = &H400040
    


    'Center the form on the screen
    If WindowState = 0 Then
       'don't attempt if screen Minimized or Maximized
       Move (Screen.Width - frmSurface.Width) / 2, (Screen.Height - frmSurface.Height) / 2
    End If


    'Initialize Labels on frmSurface
    Call LabelsSurface(UNITSTYPE_SI)


    'Load KLa Method Combo box
    cboOxygen.AddItem "Roberts & Dandliker Corr."
    cboOxygen.AddItem "User Input"


    Load frmWaterPropertiesSurface
    


    '
    ' DEMO SETTINGS.
    '
    Call LOCAL___Reset_DemoVersionDisablings
End Sub

Private Sub frmSurface_InitializeLastFewFilesList()

  'Initialize last-few-files list.
  Select Case SurfaceAerationMode
    Case DESIGN_MODE:
      Call LastFewFiles_InitializeList(LASTFEW_WHICHAPP_ASAP, LASTFEW_ASAP_frmSurface_DESIGN)
    Case RATING_MODE:
      Call LastFewFiles_InitializeList(LASTFEW_WHICHAPP_ASAP, LASTFEW_ASAP_frmSurface_RATING)
  End Select

End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
  If (frmSurface_Okay_To_Unload) Then
    Cancel = False
  Else
    Cancel = True
  End If
End Sub

Private Sub lblDisplayAirWaterProperties_Click()
    If HaveValue(sur.OperatingPressure.value) And HaveValue(sur.operatingtemperature.value) Then
       frmWaterPropertiesSurface.Show 1
    Else
       MsgBox "You must specify pressure and temperature before physical properties can be displayed.", MB_ICONSTOP, "Error"
    End If
End Sub

Private Sub mnuFile_Click(Index As Integer)
    Dim i As Integer
    Dim msg As String, Response As Integer

    Screen.MousePointer = 11   'Hourglass
    Select Case Index
       Case 0   'Switch Surface Aeration Modes
          Select Case SurfaceAerationMode
             
             Case DESIGN_MODE   'Switch to Rating Mode
                If lblConcentrationResults(4).Caption <> "" Then
                   'Give user option to save design mode before
                   'switching to rating mode
                   msg = "Would you like to save the parameters "
                   msg = msg + "for this design case to a file "
                   msg = msg + "before switching to Rating "
                   msg = msg + "Mode?"
                   Response = MsgBox(msg, MB_ICONquestion + MB_YESNO, "Save Current Design")
                   If Response = IDYES Then
                      Call SaveSurface
                   End If

                   frmSurface.Caption = "Surface Aeration - Rating Mode (untitled.sur)"
                   frmSurface!mnuFile(0).Caption = "Switch to &Design Mode"
                   SurfaceAerationMode = RATING_MODE
                
                   'Initialize last-few-files list.
                   Call frmSurface_InitializeLastFewFilesList
                   
                   sur.TankVolume.UserInput = True
                   For i = 1 To 4
                       frmSurface!txtTankParameters(i).Enabled = True
                   Next i
                   Call CalculateMinAirToWaterRatio
                Else
                   SurfaceAerationMode = RATING_MODE
                   Filename$ = "TheDefaultCaseSurface"
                   If (loadsurface("") = False) Then
                     Exit Sub
                   End If
                End If

             Case RATING_MODE   'Switch to Design Mode
                If lblConcentrationResults(4).Caption <> "" Then
                   'Give user option to save rating mode before
                   'switching to design mode
                   msg = "Would you like to save the parameters "
                   msg = msg + "for this rating case to a file "
                   msg = msg + "before switching to Design "
                   msg = msg + "Mode?"
                   Response = MsgBox(msg, MB_ICONquestion + MB_YESNO, "Save Current Design")
                   If Response = IDYES Then
                      Call SaveSurface
                   End If
    
                   frmSurface.Caption = "Surface Aeration - Design Mode (untitled.sur)"
                   frmSurface!mnuFile(0).Caption = "Switch to &Rating Mode"
                   SurfaceAerationMode = DESIGN_MODE
                   
                   'Initialize last-few-files list.
                   Call frmSurface_InitializeLastFewFilesList
                   
                   sur.CodeForTausAndTankVolumes = 1
                   Call CalculateRetentionTimeSurface
                   Call CalculateTausAndTankVolumesSurface
                   For i = 1 To 4
                       frmSurface!txtTankParameters(i).Enabled = False
                   Next i
                   Call CalculateEffluentConcentrationsSurface
                Else
                   SurfaceAerationMode = DESIGN_MODE
                   Filename$ = "TheDefaultCaseSurface"
                   If (loadsurface("") = False) Then
                     Exit Sub
                   End If
                End If
          End Select

       Case 3   'Open
          If frmSurface!lblConcentrationResults(4).Caption <> "" Then
            If (surface_savechanges()) Then Exit Sub
          End If
          
          ''''ChDrive SaveAndLoadPath
          ''''ChDir SaveAndLoadPath
          Call ChangeDir_Main
          Call loadsurface("")
            
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
          Call SaveSurface
          
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
          If Right$(frmSurface.Caption, 14) <> "(untitled.sur)" Then Call savefilesurface(Filename)
          Call SaveSurface
          
          'Add this file to the last-few-files list if necessary.
          Call LastFewFiles_MoveFilenameToTop(Filename)

          SaveAndLoadPath = CurDir$
          ''''ChDir App.Path
          ''''ChDrive App.Path
          Call ChangeDir_Main

       Case 7   'Print
          

       Case 8   'Select Printer
            On Error GoTo PrinterError
            ''''CMDialog1.flags = PD_PRINTSETUP
            ''''CMDialog1.Action = 5
            CommonDialog1.ShowPrinter
            GoTo ExitSelectPrint:
PrinterError:
            Resume ExitSelectPrint:

ExitSelectPrint:

        Case 10   'Return To Main Menu
            
            If frmSurface!lblConcentrationResults(4).Caption <> "" Then
              If (surface_savechanges()) Then Exit Sub
            End If
            
            'Unload Forms for Surface Aeration
            frmSurface_Okay_To_Unload = True
            Unload frmSurface
            Unload frmSurfaceEffluentConcentrations
            'Unload frmListContaminantSurface
            'Unload frmPropContaminantSurface
            Unload frmWaterPropertiesSurface
            
            frmMainMenu.Show
            
        
        Case 200   'Exit
            'Give user option to save design mode before Exiting

            If frmSurface!lblConcentrationResults(4).Caption <> "" Then
              If (surface_savechanges()) Then Exit Sub
            End If

            'Unload Forms for Surface Aeration
            frmSurface_Okay_To_Unload = True
            Unload frmSurface
            Unload frmSurfaceEffluentConcentrations
            'Unload frmListContaminantSurface
            'Unload frmPropContaminantSurface
            Unload frmWaterPropertiesSurface
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
        Call loadsurface(Current_LastFewFilesRec.FileNames(Index - 190))
        'Add this file to the last-few-files list if necessary.
        Call LastFewFiles_MoveFilenameToTop(Filename)
        SaveAndLoadPath = CurDir$
      End If
      ''''ChDrive App.Path
      ''''ChDir App.Path
      Call ChangeDir_Main
    End If
    
    Screen.MousePointer = 0   'Arrow

End Sub

Private Sub mnuFilePrint_Click(Index As Integer)

    Select Case Index
       Case 0   'Print to printer
          Call PrintSurface
       Case 1   'Print to file
          Call PrintSurfaceToFile
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
Call surface_results

End Sub

Private Sub mnuotheritem_Click()
'frmabout2.Show 1
End Sub

Private Sub mnuUnits_Click(Index As Integer)
  
  Select Case Index
    Case 0        'SI
      Call LabelsSurface(UNITSTYPE_SI)
    Case 1        'English
      Call LabelsSurface(UNITSTYPE_ENGLISH)
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
''''''''''''''''''''''''
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
    Dim ValueChanged As Integer, Dummy As Double
    Dim msg As String
    Dim flag_ok As Integer

   If (LostFocus_IsEvil(Me, txtDesignConcentrationValue(Index))) Then
     Exit Sub
   End If
   
   flag_ok = True


    Call TextHandleError(IsError, txtDesignConcentrationValue(Index), Temp_Text)

    If Not IsError Then
       Dummy = CDbl(txtDesignConcentrationValue(Index).Text)
       Select Case Index
          Case 3    'Contaminant Mass Transfer Coefficient
             Call TextNumberChanged(ValueChanged, txtDesignConcentrationValue(3), Temp_Text)
             If ValueChanged Then
                If HaveValue(Dummy) Then
                   sur.ContaminantMassTransferCoefficient.value = Dummy
                   sur.ContaminantMassTransferCoefficient.ValChanged = True
                   sur.ContaminantMassTransferCoefficient.UserInput = True
                Else
                   txtDesignConcentrationValue(3).Text = Temp_Text
                   txtDesignConcentrationValue(3).SetFocus
                   Exit Sub
                End If
             End If

       End Select

       If ValueChanged Then
          If SurfaceAerationMode = DESIGN_MODE Then
             Call CalculateRetentionTimeSurface
             Call CalculateTausAndTankVolumesSurface
          End If

          Call CalculateEffluentConcentrationsSurface
          Call CalculatePowerSurface
       End If

    End If
  Call LostFocus_Handle(Me, txtDesignConcentrationValue(Index), flag_ok)


End Sub

Private Sub txtFlowParameters_GotFocus(Index As Integer)
    
  Call GotFocus_Handle(Me, txtFlowParameters(Index), Temp_Text)
End Sub

Private Sub txtFlowParameters_KeyPress(Index As Integer, KeyAscii As Integer)
  Call TextBoxNumber_KeyPress(KeyAscii)
End Sub

Private Sub txtFlowParameters_LostFocus(Index As Integer)
Dim NewVal As Double
Dim IsNew As Integer
    Dim flag_ok As Integer

   If (LostFocus_IsEvil(Me, txtFlowParameters(Index))) Then
     Exit Sub
   End If
   
   flag_ok = True


  If (Unitted_LostFocus(UNITS_FLOW, txtFlowParameters(0), UnitsFlowParam(0), NewVal, Temp_Text)) Then
    sur.WaterFlowRate.ValChanged = True
    sur.WaterFlowRate.UserInput = True
    sur.WaterFlowRate.value = NewVal
    Call CalculateRetentionTimesAndTankVolumes

    If (SurfaceAerationMode = DESIGN_MODE) Then
      Call CalculateRetentionTimeSurface
      Call CalculateTausAndTankVolumesSurface
    Else
      Call CalculateTausAndTankVolumesSurface
    End If
    Call CalculateEffluentConcentrationsSurface
  
    '
    ' REFRESH THE DISPLAY.
    '
Dim SaveIndex As Integer
    SaveIndex = cboDesignContaminant.ListIndex
    cboDesignContaminant.ListIndex = -1
    cboDesignContaminant.ListIndex = SaveIndex
  End If
  Call LostFocus_Handle(Me, txtFlowParameters(Index), flag_ok)


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
    sur.OperatingPressure.ValChanged = True
    sur.OperatingPressure.UserInput = True
    'Note: standard P units are Pa, but
    'OperatingPressure is stored as kPa.
    sur.OperatingPressure.value = NewVal * 1# / 101325#

    If (HaveValue(sur.OperatingPressure.value) And HaveValue(sur.operatingtemperature.value)) Then
      Call CalculateWaterPropertiesSurface
      If sur.NumChemical > 0 Then
        'Update Variables on Screen
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
    sur.operatingtemperature.ValChanged = True
    sur.operatingtemperature.UserInput = True
    sur.operatingtemperature.value = NewVal

    If (HaveValue(sur.OperatingPressure.value) And HaveValue(sur.operatingtemperature.value)) Then
      Call CalculateWaterPropertiesSurface
      Call CalculateOxygenLiquidDiffSurface
      If cboOxygen.ListIndex = 0 Then
        Call CalculateOxygenMTCoeffSurface
      End If
         
      If (sur.NumChemical > 0) Then
        Call CalculateContaminantMTCoeffSurface
        If SurfaceAerationMode = DESIGN_MODE Then
          Call CalculateRetentionTimeSurface
          Call CalculateTausAndTankVolumesSurface
        End If
        Call CalculateEffluentConcentrationsSurface
        Call CalculatePowerSurface
      End If
    End If
  End If
  Call LostFocus_Handle(Me, txtOperatingTemperature, flag_ok)
                               
End Sub

Private Sub txtOxygen_GotFocus(Index As Integer)
    
  Call GotFocus_Handle(Me, txtOxygen(Index), Temp_Text)
End Sub

Private Sub txtOxygen_KeyPress(Index As Integer, KeyAscii As Integer)
  Call TextBoxNumber_KeyPress(KeyAscii)
End Sub

Private Sub txtOxygen_LostFocus(Index As Integer)
Dim NewVal As Double
Dim IsNew As Integer
    Dim flag_ok As Integer

   If (LostFocus_IsEvil(Me, txtOxygen(Index))) Then
     Exit Sub
   End If
   
   flag_ok = True


  IsNew = False
  
  Select Case Index
    Case 1        'Liquid Diffusivity.
      If (Unitted_LostFocus(UNITS_DIFFUSIVITY, txtOxygen(1), UnitsOxygenRef(1), NewVal, Temp_Text)) Then
        IsNew = True
        sur.Oxygen.LiquidDiffusivity.value = NewVal
        sur.Oxygen.LiquidDiffusivity.ValChanged = True
        sur.Oxygen.LiquidDiffusivity.UserInput = True
      End If

    Case 2        'KLa.
      If (Unitted_LostFocus(UNITS_INVERSETIME, txtOxygen(2), UnitsOxygenRef(2), NewVal, Temp_Text)) Then
        IsNew = True
        sur.Oxygen.MassTransferCoefficient.value = NewVal
        sur.Oxygen.MassTransferCoefficient.ValChanged = True
        sur.Oxygen.MassTransferCoefficient.UserInput = True
      End If
    
  End Select

  If (IsNew) Then
    Select Case Index
      Case 1
        If HaveValue(sur.OperatingPressure.value) And HaveValue(sur.operatingtemperature.value) And HaveValue(sur.Oxygen.LiquidDiffusivity.value) Then
          If sur.NumChemical > 0 Then
            Call CalculateContaminantMTCoeffSurface
            If SurfaceAerationMode = DESIGN_MODE Then
              Call CalculateRetentionTimeSurface
              Call CalculateTausAndTankVolumesSurface
            End If
            Call CalculateEffluentConcentrationsSurface
            Call CalculatePowerSurface
          End If
        End If
      
      Case 2
        If HaveValue(sur.OperatingPressure.value) And HaveValue(sur.operatingtemperature.value) And HaveValue(sur.Oxygen.LiquidDiffusivity.value) And HaveValue(sur.Oxygen.MassTransferCoefficient.value) Then
          If sur.NumChemical > 0 Then
            Call CalculateContaminantMTCoeffSurface
            If SurfaceAerationMode = DESIGN_MODE Then
              Call CalculateRetentionTimeSurface
              Call CalculateTausAndTankVolumesSurface
            End If
            Call CalculateEffluentConcentrationsSurface
            Call CalculatePowerSurface
          End If
        End If
        If cboOxygen.ListIndex = 0 Then cboOxygen.ListIndex = 1
    
    End Select
  End If
  Call LostFocus_Handle(Me, txtOxygen(Index), flag_ok)


End Sub

Private Sub txtPowerCalculation_GotFocus(Index As Integer)
    
  Call GotFocus_Handle(Me, txtPowerCalculation(Index), Temp_Text)
End Sub

Private Sub txtPowerCalculation_KeyPress(Index As Integer, KeyAscii As Integer)
  Call TextBoxNumber_KeyPress(KeyAscii)
End Sub

Private Sub txtPowerCalculation_LostFocus(Index As Integer)
    Dim Answer As Integer, Response As Integer
    Dim msg As String
    Dim ValueChanged As Integer
    Dim flag_ok As Integer

   If (LostFocus_IsEvil(Me, txtPowerCalculation(Index))) Then
     Exit Sub
   End If
   
   flag_ok = True


    Call TextHandleError(IsError, txtPowerCalculation(0), Temp_Text)
    If Not IsError Then
       If Not HaveValue(CDbl(txtPowerCalculation(0).Text)) Then
          txtPowerCalculation(0).Text = Temp_Text
          txtPowerCalculation(0).SetFocus
          Exit Sub
       End If
       
       Call TextNumberChanged(ValueChanged, txtPowerCalculation(0), Temp_Text)

       If ValueChanged Then
          sur.Power.AeratorMotorEfficiency = CDbl(txtPowerCalculation(0).Text)
       Else
         Call LostFocus_Handle(Me, txtPowerCalculation(Index), flag_ok)
          Exit Sub
       End If

       Call CalculatePowerSurface

    End If
  Call LostFocus_Handle(Me, txtPowerCalculation(Index), flag_ok)


End Sub

Private Sub txtPowerInput_GotFocus()
    
  Call GotFocus_Handle(Me, txtPowerInput, Temp_Text)
End Sub

Private Sub txtPowerInput_KeyPress(KeyAscii As Integer)
  Call TextBoxNumber_KeyPress(KeyAscii)
End Sub

Private Sub txtPowerInput_LostFocus()
Dim NewVal As Double
    Dim flag_ok As Integer

   If (LostFocus_IsEvil(Me, txtPowerInput)) Then
     Exit Sub
   End If
   
   flag_ok = True


  If (Unitted_LostFocus(UNITS_PRESSURE, txtPowerInput, UnitsPowerInput, NewVal, Temp_Text)) Then
    sur.PowerInput_PoverV.ValChanged = True
    sur.PowerInput_PoverV.UserInput = True
    sur.PowerInput_PoverV.value = NewVal

    Call CalculateOxygenLiquidDiffSurface
    If cboOxygen.ListIndex = 0 Then
      Call CalculateOxygenMTCoeffSurface
    End If
         
    If sur.NumChemical > 0 Then
      Call CalculateContaminantMTCoeffSurface
      If SurfaceAerationMode = DESIGN_MODE Then
        Call CalculateRetentionTimeSurface
        Call CalculateTausAndTankVolumesSurface
      End If
      Call CalculateEffluentConcentrationsSurface
      Call CalculatePowerSurface
    End If

  End If
  Call LostFocus_Handle(Me, txtPowerInput, flag_ok)

    
End Sub

Private Sub txtTankParameters_Change(Index As Integer)

'  Call UnitsTankParam_Click(Index)

End Sub

Private Sub txtTankParameters_GotFocus(Index As Integer)
    
  Call GotFocus_Handle(Me, txtTankParameters(Index), Temp_Text)
End Sub

Private Sub txtTankParameters_KeyPress(Index As Integer, KeyAscii As Integer)
  Call TextBoxNumber_KeyPress(KeyAscii)
End Sub

Private Sub txtTankParameters_LostFocus(Index As Integer)
Dim NewVal As Double
Dim IsNew As Integer
    Dim flag_ok As Integer

   If (LostFocus_IsEvil(Me, txtTankParameters(Index))) Then
     Exit Sub
   End If
   
   flag_ok = True


  IsNew = False
  
  Select Case Index
    Case 0        'Number of Tanks in Series.
      If (NoUnits_LostFocus(txtTankParameters(0), NewVal, Temp_Text)) Then
        IsNew = True
        If (NewVal > MAXIMUM_TANKS) Then
          MsgBox "The number of tanks specified exceeds the maximum number of tanks allowed in the program.  The number of tanks must be less than or equal to " & Format$(MAXIMUM_TANKS, "0") & ".", MB_ICONSTOP, "Bubble Aeration"
          txtTankParameters(0).Text = Temp_Text
          txtTankParameters(0).SetFocus
          Exit Sub
        End If

        sur.NumberOfTanks.ValChanged = True
        sur.NumberOfTanks.UserInput = True
        sur.NumberOfTanks.value = NewVal
       
        If SurfaceAerationMode = DESIGN_MODE Then
          Call CalculateRetentionTimeSurface
          Call CalculateTausAndTankVolumesSurface
        Else
          Call CalculateTausAndTankVolumesSurface
        End If
      End If

    Case 1        'Retention Time for 1 Tank.
      If (Unitted_LostFocus(UNITS_TIME, txtTankParameters(1), UnitsTankParam(1), NewVal, Temp_Text)) Then
        IsNew = True
        'Standard time units are seconds, but TankHydraulicRetentionTime
        'is stored internally as hours.
        sur.TankHydraulicRetentionTime.value = NewVal / 60# / 60#
        sur.TankHydraulicRetentionTime.ValChanged = True
        sur.TankHydraulicRetentionTime.UserInput = True
        sur.CodeForTausAndTankVolumes = 1
        Call CalculateTausAndTankVolumesSurface
      End If
        
    Case 2        'Retention Time for All Tanks.
      If (Unitted_LostFocus(UNITS_TIME, txtTankParameters(2), UnitsTankParam(2), NewVal, Temp_Text)) Then
        IsNew = True
        'Standard time units are seconds, but TotalHydraulicRetentionTime
        'is stored internally as hours.
        sur.TotalHydraulicRetentionTime.value = NewVal / 60# / 60#
        sur.TotalHydraulicRetentionTime.ValChanged = True
        sur.TotalHydraulicRetentionTime.UserInput = True
        sur.CodeForTausAndTankVolumes = 2
        Call CalculateTausAndTankVolumesSurface
      End If

    Case 3        'Volume of Each Tank.
      If (Unitted_LostFocus(UNITS_VOLUME, txtTankParameters(3), UnitsTankParam(3), NewVal, Temp_Text)) Then
        IsNew = True
        sur.TankVolume.value = NewVal
        sur.TankVolume.ValChanged = True
        sur.TankVolume.UserInput = True
        sur.CodeForTausAndTankVolumes = 3
        Call CalculateTausAndTankVolumesSurface
      End If

    Case 4        'Volume of All Tanks.
      If (Unitted_LostFocus(UNITS_VOLUME, txtTankParameters(4), UnitsTankParam(4), NewVal, Temp_Text)) Then
        IsNew = True
        sur.TotalTankVolume.value = NewVal
        sur.TotalTankVolume.ValChanged = True
        sur.TotalTankVolume.UserInput = True
        sur.CodeForTausAndTankVolumes = 4
        Call CalculateTausAndTankVolumesSurface
      End If

  End Select

  If (IsNew) Then
    Call CalculateEffluentConcentrationsSurface
    Call CalculatePowerSurface
    Call cboDesignContaminant_Click
  End If
  Call LostFocus_Handle(Me, txtTankParameters(Index), flag_ok)


End Sub

Private Sub UnitsConcResults_Click(Index As Integer)
Dim Dummy As Double
Dim ContaminantIndex As Integer
Dim i As Integer

  ContaminantIndex = cboDesignContaminant.ListIndex + 1
  i = ContaminantIndex

  On Error GoTo err_UnitsConcResults_Click

  Select Case Index
    Case 1            'Ci to Tank 1.
      Dummy = sur.Contaminant(i).Influent.value
      Call Unitted_UnitChange(UNITS_CONCENTRATION, Dummy, UnitsConcResults(1), lblConcentrationResults(1))

    Case 3            'Ce from Last Tank.
      Dummy = sur.DesignContaminant.Effluent(sur.NumberOfTanks.value)
      Call Unitted_UnitChange(UNITS_CONCENTRATION, Dummy, UnitsConcResults(3), lblConcentrationResults(3))

  End Select

exit_UnitsConcResults_Click:
  Exit Sub

err_UnitsConcResults_Click:
  Resume exit_UnitsConcResults_Click

End Sub

Private Sub UnitsConcResults_KeyPress(Index As Integer, KeyAscii As Integer)
  Call Combobox_KeyPress(KeyAscii)
End Sub

Private Sub UnitsDesignContam_Click(Index As Integer)
Dim Dummy As Double
Dim ContaminantIndex As Integer
Dim i As Integer

  ContaminantIndex = cboDesignContaminant.ListIndex + 1
  i = ContaminantIndex

  On Error GoTo err_UnitsDesignContam_Click

  Select Case Index
    Case 0            'Influent Conc.
      Dummy = sur.Contaminant(i).Influent.value
      Call Unitted_UnitChange(UNITS_CONCENTRATION, Dummy, UnitsDesignContam(0), lblDesignConcentrationValue(0))

    Case 1            'Treatment Obj.
      Dummy = sur.Contaminant(i).TreatmentObjective.value
      Call Unitted_UnitChange(UNITS_CONCENTRATION, Dummy, UnitsDesignContam(1), lblDesignConcentrationValue(1))

    Case 3            'KLa
      Dummy = sur.ContaminantMassTransferCoefficient.value
      Call Unitted_UnitChange(UNITS_INVERSETIME, Dummy, UnitsDesignContam(3), txtDesignConcentrationValue(3))

  End Select

exit_UnitsDesignContam_Click:
  Exit Sub

err_UnitsDesignContam_Click:
  Resume exit_UnitsDesignContam_Click

End Sub

Private Sub UnitsDesignContam_KeyPress(Index As Integer, KeyAscii As Integer)
  Call Combobox_KeyPress(KeyAscii)
End Sub

Private Sub UnitsFlowParam_Click(Index As Integer)

  Select Case Index
    Case 0            'Water Flow Rate.
      Call Unitted_UnitChange(UNITS_FLOW, sur.WaterFlowRate.value, UnitsFlowParam(0), txtFlowParameters(0))
  End Select

End Sub

Private Sub UnitsFlowParam_KeyPress(Index As Integer, KeyAscii As Integer)
  Call Combobox_KeyPress(KeyAscii)
End Sub

Private Sub UnitsOpCond_Click(Index As Integer)
Dim Dummy As Double

  Select Case Index
    Case 0
      'Note: Standard P units are Pa, but OperatingPressure
      'is stored internally in kPa units.
      Dummy = sur.OperatingPressure.value * 101325#
      Call Unitted_UnitChange(UNITS_PRESSURE, Dummy, UnitsOpCond(0), txtOperatingPressure)

    Case 1
      Call Unitted_UnitChange(UNITS_TEMPERATURE, sur.operatingtemperature.value, UnitsOpCond(1), txtOperatingTemperature)
    
  End Select

End Sub

Private Sub UnitsOpCond_KeyPress(Index As Integer, KeyAscii As Integer)
  Call Combobox_KeyPress(KeyAscii)
End Sub

Private Sub UnitsOxygenRef_Click(Index As Integer)

  Select Case Index
    Case 1            'Liquid Diffusivity
      Call Unitted_UnitChange(UNITS_DIFFUSIVITY, sur.Oxygen.LiquidDiffusivity.value, UnitsOxygenRef(1), txtOxygen(1))

    Case 2            'KLa
      Call Unitted_UnitChange(UNITS_INVERSETIME, sur.Oxygen.MassTransferCoefficient.value, UnitsOxygenRef(2), txtOxygen(2))
    
  End Select

End Sub

Private Sub UnitsOxygenRef_KeyPress(Index As Integer, KeyAscii As Integer)
  Call Combobox_KeyPress(KeyAscii)
End Sub

Private Sub UnitsPowerCalc_Click(Index As Integer)
Dim Dummy As Double
Dim ContaminantIndex As Integer
Dim i As Integer

  ContaminantIndex = cboDesignContaminant.ListIndex + 1
  i = ContaminantIndex

  On Error GoTo err_UnitsPowerCalc_Change

  Select Case Index
    Case 1            'Power Required per Tank.
      Dummy = sur.Power.PowerForEachTank
      Call Unitted_UnitChange(UNITS_POWER, Dummy, UnitsPowerCalc(1), lblPowerCalculation(1))

    Case 2            'Total Power Required.
      Dummy = sur.Power.TotalPowerForAllTanks
      Call Unitted_UnitChange(UNITS_POWER, Dummy, UnitsPowerCalc(2), lblPowerCalculation(2))

  End Select

exit_UnitsPowerCalc_Change:
  Exit Sub

err_UnitsPowerCalc_Change:
  Resume exit_UnitsPowerCalc_Change

End Sub

Private Sub UnitsPowerCalc_KeyPress(Index As Integer, KeyAscii As Integer)
  Call Combobox_KeyPress(KeyAscii)
End Sub

Private Sub UnitsPowerInput_Click()
Dim Dummy As Double

  Dummy = sur.PowerInput_PoverV.value
  Call Unitted_UnitChange(UNITS_POWERPERVOLUME, Dummy, UnitsPowerInput, txtPowerInput)

End Sub

Private Sub UnitsTankParam_Click(Index As Integer)
Dim Dummy As Double

  Select Case Index
    Case 1            'Retention Time (1 Tank).
      'Standard time units are seconds, but TankHydraulicRetentionTime
      'is stored internally as hours.
      Dummy = sur.TankHydraulicRetentionTime.value * 60 * 60#
      Call Unitted_UnitChange(UNITS_TIME, Dummy, UnitsTankParam(1), txtTankParameters(1))

    Case 2            'Retention Time (All Tanks).
      'Standard time units are seconds, but TotalHydraulicRetentionTime
      'is stored internally as hours.
      Dummy = sur.TotalHydraulicRetentionTime.value * 60 * 60#
      Call Unitted_UnitChange(UNITS_TIME, Dummy, UnitsTankParam(2), txtTankParameters(2))

    Case 3            'Volume (1 Tank).
      Call Unitted_UnitChange(UNITS_VOLUME, sur.TankVolume.value, UnitsTankParam(3), txtTankParameters(3))

    Case 4            'Volume (All Tanks).
      Call Unitted_UnitChange(UNITS_VOLUME, sur.TotalTankVolume.value, UnitsTankParam(4), txtTankParameters(4))
    
  End Select

End Sub

Private Sub UnitsTankParam_KeyPress(Index As Integer, KeyAscii As Integer)
  Call Combobox_KeyPress(KeyAscii)
End Sub


