VERSION 5.00
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "THREED32.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmBubble 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Bubble Aeration"
   ClientHeight    =   6795
   ClientLeft      =   1860
   ClientTop       =   1365
   ClientWidth     =   9480
   Icon            =   "Bubble.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   ScaleHeight     =   6795
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
      TabIndex        =   19
      TabStop         =   0   'False
      Top             =   5910
      Width           =   2655
   End
   Begin Threed.SSFrame fraOperatingConditions 
      Height          =   1335
      Left            =   30
      TabIndex        =   13
      Top             =   90
      Width           =   4455
      _Version        =   65536
      _ExtentX        =   7858
      _ExtentY        =   2355
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
         Left            =   120
         TabIndex        =   22
         TabStop         =   0   'False
         Top             =   930
         Width           =   4215
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
         Left            =   1800
         TabIndex        =   1
         Top             =   570
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
         Left            =   1800
         TabIndex        =   0
         Top             =   270
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
         Left            =   3060
         Style           =   2  'Dropdown List
         TabIndex        =   21
         TabStop         =   0   'False
         Top             =   270
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
         TabIndex        =   20
         TabStop         =   0   'False
         Top             =   570
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
         Left            =   -960
         TabIndex        =   24
         Top             =   570
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
         Left            =   -960
         TabIndex        =   23
         Top             =   270
         Width           =   2655
      End
   End
   Begin Threed.SSFrame fraOxygen 
      Height          =   1455
      Left            =   30
      TabIndex        =   14
      Top             =   1500
      Width           =   4455
      _Version        =   65536
      _ExtentX        =   7858
      _ExtentY        =   2566
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
         Left            =   1320
         Style           =   2  'Dropdown List
         TabIndex        =   27
         TabStop         =   0   'False
         Top             =   270
         Width           =   3015
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
         Left            =   1800
         TabIndex        =   3
         Top             =   1050
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
         Left            =   1800
         TabIndex        =   2
         Top             =   690
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
         Left            =   3060
         Style           =   2  'Dropdown List
         TabIndex        =   26
         TabStop         =   0   'False
         Top             =   690
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
         Left            =   3060
         Style           =   2  'Dropdown List
         TabIndex        =   25
         TabStop         =   0   'False
         Top             =   1050
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
         Left            =   0
         TabIndex        =   30
         Top             =   330
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
         Left            =   -960
         TabIndex        =   29
         Top             =   1050
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
         Left            =   -960
         TabIndex        =   28
         Top             =   690
         Width           =   2655
      End
   End
   Begin Threed.SSFrame fraContaminantInformation 
      Height          =   2655
      Left            =   30
      TabIndex        =   15
      Top             =   3060
      Width           =   4455
      _Version        =   65536
      _ExtentX        =   7858
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
         TabIndex        =   37
         TabStop         =   0   'False
         Top             =   810
         Width           =   1755
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
         TabIndex        =   36
         TabStop         =   0   'False
         Top             =   810
         Width           =   975
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
         TabIndex        =   35
         TabStop         =   0   'False
         Top             =   810
         Width           =   1035
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
         Left            =   240
         Style           =   2  'Dropdown List
         TabIndex        =   34
         TabStop         =   0   'False
         Top             =   450
         Width           =   3975
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
         Index           =   3
         Left            =   1920
         TabIndex        =   4
         Top             =   2250
         Width           =   1095
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
         Left            =   3060
         Style           =   2  'Dropdown List
         TabIndex        =   33
         TabStop         =   0   'False
         Top             =   1350
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
         Left            =   3060
         Style           =   2  'Dropdown List
         TabIndex        =   32
         TabStop         =   0   'False
         Top             =   1650
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
         Left            =   3060
         Style           =   2  'Dropdown List
         TabIndex        =   31
         TabStop         =   0   'False
         Top             =   2250
         Width           =   1275
      End
      Begin VB.Shape Shape1 
         Height          =   795
         Left            =   120
         Top             =   390
         Width           =   4215
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
         Left            =   -720
         TabIndex        =   45
         Top             =   2250
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
         Left            =   -720
         TabIndex        =   44
         Top             =   1350
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
         Left            =   -720
         TabIndex        =   43
         Top             =   1650
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
         Left            =   -720
         TabIndex        =   42
         Top             =   1950
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
         Left            =   1920
         TabIndex        =   41
         Top             =   1350
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
         Left            =   1920
         TabIndex        =   40
         Top             =   1650
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
         Left            =   1920
         TabIndex        =   39
         Top             =   1950
         Width           =   1095
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
         Left            =   3060
         TabIndex        =   38
         Top             =   1980
         Width           =   1155
      End
   End
   Begin Threed.SSFrame fraFlowParameters 
      Height          =   1515
      Left            =   4470
      TabIndex        =   16
      Top             =   90
      Width           =   4935
      _Version        =   65536
      _ExtentX        =   8705
      _ExtentY        =   2672
      _StockProps     =   14
      Caption         =   "Flow Parameters"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
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
         Index           =   3
         Left            =   2400
         TabIndex        =   7
         Top             =   1110
         Width           =   1215
      End
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
         Index           =   2
         Left            =   2400
         TabIndex        =   6
         Top             =   810
         Width           =   1215
      End
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
         Left            =   2400
         TabIndex        =   5
         Top             =   150
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
         Left            =   3660
         Style           =   2  'Dropdown List
         TabIndex        =   47
         TabStop         =   0   'False
         Top             =   150
         Width           =   1155
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
         Index           =   3
         Left            =   3660
         Style           =   2  'Dropdown List
         TabIndex        =   46
         TabStop         =   0   'False
         Top             =   1110
         Width           =   1155
      End
      Begin VB.Label lblFlowParametersLabel 
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
         Index           =   3
         Left            =   -420
         TabIndex        =   54
         Top             =   1110
         Width           =   2715
      End
      Begin VB.Label lblFlowParametersLabel 
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
         Left            =   -360
         TabIndex        =   53
         Top             =   810
         Width           =   2655
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
         Left            =   -480
         TabIndex        =   52
         Top             =   210
         Width           =   2715
      End
      Begin VB.Label lblFlowParameters 
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
         Left            =   2400
         TabIndex        =   51
         Top             =   510
         Width           =   1215
      End
      Begin VB.Label lblFlowParametersLabel 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Min. Air To Water Ratio"
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
         Left            =   -420
         TabIndex        =   50
         Top             =   510
         Width           =   2715
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
         Left            =   3660
         TabIndex        =   49
         Top             =   510
         Width           =   1155
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
         Index           =   1
         Left            =   3660
         TabIndex        =   48
         Top             =   810
         Width           =   1155
      End
   End
   Begin Threed.SSFrame fraTankParameters 
      Height          =   1815
      Left            =   4470
      TabIndex        =   17
      Top             =   1680
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
         Left            =   2400
         TabIndex        =   8
         Top             =   210
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
         Left            =   2400
         TabIndex        =   9
         Top             =   510
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
         Left            =   2400
         TabIndex        =   10
         Top             =   810
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
         Left            =   2400
         TabIndex        =   11
         Top             =   1110
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
         Left            =   2400
         TabIndex        =   12
         Top             =   1410
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
         Left            =   3660
         Style           =   2  'Dropdown List
         TabIndex        =   58
         TabStop         =   0   'False
         Top             =   510
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
         Left            =   3660
         Style           =   2  'Dropdown List
         TabIndex        =   57
         TabStop         =   0   'False
         Top             =   810
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
         Left            =   3660
         Style           =   2  'Dropdown List
         TabIndex        =   56
         TabStop         =   0   'False
         Top             =   1110
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
         Left            =   3660
         Style           =   2  'Dropdown List
         TabIndex        =   55
         TabStop         =   0   'False
         Top             =   1410
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
         Left            =   -420
         TabIndex        =   63
         Top             =   270
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
         Left            =   -420
         TabIndex        =   62
         Top             =   570
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
         Left            =   -420
         TabIndex        =   61
         Top             =   870
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
         Left            =   -420
         TabIndex        =   60
         Top             =   1170
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
         Left            =   -420
         TabIndex        =   59
         Top             =   1470
         Width           =   2715
      End
   End
   Begin Threed.SSFrame fraConcentrationResults 
      Height          =   2295
      Left            =   4470
      TabIndex        =   18
      Top             =   3900
      Width           =   4935
      _Version        =   65536
      _ExtentX        =   8705
      _ExtentY        =   4048
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
         Left            =   60
         TabIndex        =   67
         TabStop         =   0   'False
         Top             =   1200
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
         Left            =   3660
         Style           =   2  'Dropdown List
         TabIndex        =   66
         TabStop         =   0   'False
         Top             =   540
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
         Index           =   2
         Left            =   3660
         Style           =   2  'Dropdown List
         TabIndex        =   65
         TabStop         =   0   'False
         Top             =   840
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
         Left            =   3660
         Style           =   2  'Dropdown List
         TabIndex        =   64
         TabStop         =   0   'False
         Top             =   1620
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
         TabIndex        =   78
         Top             =   240
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
         Left            =   -300
         TabIndex        =   77
         Top             =   540
         Width           =   2595
      End
      Begin VB.Label lblConcentrationResultsLabel 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Yi to All Tanks"
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
         Left            =   -300
         TabIndex        =   76
         Top             =   840
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
         Left            =   -300
         TabIndex        =   75
         Top             =   1620
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
         Left            =   -300
         TabIndex        =   74
         Top             =   1920
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
         Left            =   1140
         TabIndex        =   73
         Top             =   240
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
         Left            =   2400
         TabIndex        =   72
         Top             =   540
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
         Index           =   2
         Left            =   2400
         TabIndex        =   71
         Top             =   840
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
         Left            =   2400
         TabIndex        =   70
         Top             =   1620
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
         Left            =   2400
         TabIndex        =   69
         Top             =   1920
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
         Index           =   3
         Left            =   3660
         TabIndex        =   68
         Top             =   1920
         Width           =   1155
      End
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   210
      Top             =   6180
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Label lblStantonLabel 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Stanton Number"
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   4050
      TabIndex        =   80
      Top             =   3570
      Width           =   2715
   End
   Begin VB.Label lblStanton 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   6870
      TabIndex        =   79
      Top             =   3570
      Width           =   1215
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
   Begin VB.Menu mnuPopUpContaminant 
      Caption         =   "PopupContaminant"
      Visible         =   0   'False
      Begin VB.Menu mnuPopContaminant 
         Caption         =   "Call &StEPP"
         Enabled         =   0   'False
         Index           =   0
      End
      Begin VB.Menu mnuPopContaminant 
         Caption         =   "User &Input"
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
Attribute VB_Name = "frmBubble"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim Temp_Text As String
Public frmBubble_Okay_To_Unload As Boolean



Const frmBubble_declarations_end = True


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
    fraFlowParameters.Caption = "* DEMONSTRATION VERSION *"
    fraFlowParameters.ForeColor = QBColor(12)
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

    ContaminantIndex = cboDesignContaminant.ListIndex + 1
    i = ContaminantIndex
    If i = 0 Then Exit Sub
    bub.DesignContaminant = bub.Contaminant(i)

    'Update Design Contaminant | Influent Conc.
    'lblDesignConcentrationValue(0).Caption = Format$(bub.Contaminant(i).Influent.Value, "0.0")
    Call UnitsDesignContam_Click(0)
    
    'Update Design Contaminant | Treatment Obj.
    'lblDesignConcentrationValue(1).Caption = Format$(bub.Contaminant(i).TreatmentObjective.Value, "0.0")
    Call UnitsDesignContam_Click(1)

    'Call system_log("REMOVBUB " & Str$(bub.DesiredPercentRemoval) & ", " & Str$(bub.Contaminant(i).Influent.value) & ", " & Str$(bub.Contaminant(i).TreatmentObjective.value))
    Call REMOVBUB(bub.DesiredPercentRemoval, bub.Contaminant(i).Influent.value, bub.Contaminant(i).TreatmentObjective.value)
    Call CalculateMinAirToWaterRatio
    
    lblDesignConcentrationValue(2).Caption = Format$(bub.DesiredPercentRemoval, "0.0")
    
    lblConcentrationResults(0).Caption = cboDesignContaminant.Text

    'Update Concentration Results | Ci to Tank 1.
    'lblConcentrationResults(1).Caption = lblDesignConcentrationValue(0).Caption
    Call UnitsConcResults_Click(1)

    'Update Concentration Results | Yi to All Tanks.
    Call UnitsConcResults_Click(2)    'Note: Always 0.
    
    'Update Concentration Results | Ce from Last Tank.
    Call UnitsConcResults_Click(3)
    'UPDATED_UNITS

    Call CalculateContaminantMTCoeff

    'Update Design Contaminant | KLa
    Call UnitsDesignContam_Click(3)

    If BubbleAerationMode = DESIGN_MODE Then
       If bub.AirToWaterRatio.value < bub.MinimumAirToWaterRatio.value Then
          frmBubbleAchievingRemovalEfficiency!lblAchieving(0).Caption = Format$(bub.MinimumAirToWaterRatio.value, GetTheFormat(bub.MinimumAirToWaterRatio.value))
          frmBubbleAchievingRemovalEfficiency!txtAchieving(1).Text = Format$(bub.AirToWaterRatio.value, GetTheFormat(bub.AirToWaterRatio.value))
          frmBubbleAchievingRemovalEfficiency!txtAchieving(2).Text = Format$(bub.NumberOfTanks.value, "0")
          frmBubbleAchievingRemovalEfficiency.Show 1
       End If
       Call CalculateTankVolumeBubble
       Call CalculateRetentionTimesAndTankVolumes
    End If
    Call CalculateStantonNo
    Call CalculateEffluentConcentrationsBubble

    'Update Tank Parameters:
    For i = 1 To 4
      Call UnitsTankParam_Click(i)
    Next i

    'Update Flow Parameters:
    Call UnitsFlowParam_Click(0)
    Call UnitsFlowParam_Click(3)
    
End Sub

Private Sub cboDesignContaminant_GotFocus()
    Temp_Text = cboDesignContaminant.Text
End Sub

Private Sub cboDesignContaminant_KeyPress(KeyAscii As Integer)
  Call TextBoxNumber_KeyPress(KeyAscii)
End Sub

Private Sub cboOxygen_Click()
   
   If cboOxygen.ListIndex = 0 Then   'Clean Water Oxygen Transfer Test Data
      frmOxygenMassTransferCoeff.Show
   Else
      bub.Oxygen.KLaMethod = 2
      bub.Oxygen.MassTransferCoefficient.UserInput = True
   End If

End Sub

Private Sub cmdAddComponent_Click()
Dim x As rec_frmContaminantPropertyEdit
Dim i As Integer

  If (bub.NumChemical + 1 > MAXCHEMICAL) Then
     MsgBox "The maximum number of contaminants has been reached.  It is not possible to input more than " & Format$(MAXCHEMICAL, "0") & " contaminants for design.  To add an additional contaminant, another must first be removed.", MB_ICONSTOP, "Bubble Aeration"
     cmdAddComponent.Enabled = False
     Exit Sub
  End If

  x.ModelName = "Bubble Aeration"
  x.ModelType = MODELTYPE_BUBBLE
  x.DoEditNumber = -1       'Will be set by frmContaminantPropertyEdit.
  x.DoAdd = True
  x.OldNumCompo = cboDesignContaminant.ListCount
  For i = 1 To x.OldNumCompo
    x.Contaminants(i).Name = bub.Contaminant(i).Name
    x.Contaminants(i).MolecularWeight.value = bub.Contaminant(i).MolecularWeight.value
    x.Contaminants(i).HenrysConstant.value = bub.Contaminant(i).HenrysConstant.value
    x.Contaminants(i).MolarVolume.value = bub.Contaminant(i).MolarVolume.value
    x.Contaminants(i).LiquidDiffusivity.value = bub.Contaminant(i).LiquidDiffusivity.value
    x.Contaminants(i).Influent.value = bub.Contaminant(i).Influent.value
    x.Contaminants(i).TreatmentObjective.value = bub.Contaminant(i).TreatmentObjective.value
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

      bub.Contaminant(i).Name = x.Contaminants(i).Name
      bub.Contaminant(i).MolecularWeight.value = x.Contaminants(i).MolecularWeight.value
      bub.Contaminant(i).HenrysConstant.value = x.Contaminants(i).HenrysConstant.value
      bub.Contaminant(i).MolarVolume.value = x.Contaminants(i).MolarVolume.value
      bub.Contaminant(i).LiquidDiffusivity.value = x.Contaminants(i).LiquidDiffusivity.value
      bub.Contaminant(i).Influent.value = x.Contaminants(i).Influent.value
      bub.Contaminant(i).TreatmentObjective.value = x.Contaminants(i).TreatmentObjective.value
      bub.NumChemical = bub.NumChemical + 1
      'Incorporate new name into ComboBox.
      cboDesignContaminant.AddItem bub.Contaminant(i).Name
    Next i
  End If
  
  Call SetDesignContaminantEnabledBubble(CInt(cboDesignContaminant.ListCount))
  
  If (bub.NumChemical > 0) Then
    cmdDeleteComponent.Enabled = True
    cmdEditComponent.Enabled = True
    If (cboDesignContaminant.ListIndex = -1) Then
      cboDesignContaminant.ListIndex = 0
    End If
  End If
 Me.Show
End Sub

Private Sub cmdConcentrationResults_Click()
    frmBubbleEffluentConcentrations.Show 1

End Sub

Private Sub cmdDeleteComponent_Click()
  Dim i As Integer

  bub.Chemical = cboDesignContaminant.ListIndex + 1
  If (bub.Chemical = 0) Then Exit Sub

  If MsgBox("Remove" & NL & cboDesignContaminant.List(cboDesignContaminant.ListIndex), 36, "") = IDYES Then
    cboDesignContaminant.RemoveItem cboDesignContaminant.ListIndex
    For i = bub.Chemical To bub.NumChemical - 1
      bub.Contaminant(i) = bub.Contaminant(i + 1)
    Next i
    bub.NumChemical = bub.NumChemical - 1
    If (bub.NumChemical > 0) Then
      cboDesignContaminant.ListIndex = 0
    Else
      cmdDeleteComponent.Enabled = False
      cmdEditComponent.Enabled = False
    End If
    Call SetDesignContaminantEnabledBubble(CInt(cboDesignContaminant.ListCount))
  End If

    If bub.NumChemical < 10 Then cmdAddComponent.Enabled = True
    Call LOCAL___Reset_DemoVersionDisablings
End Sub

Private Sub cmdEditComponent_Click()
Dim x As rec_frmContaminantPropertyEdit
Dim i As Integer
Dim AListIndex As Integer

  bub.Chemical = cboDesignContaminant.ListIndex + 1
  If (bub.Chemical = 0) Then Exit Sub

  x.ModelName = "Bubble Aeration"
  x.ModelType = MODELTYPE_BUBBLE
  x.DoEditNumber = cboDesignContaminant.ListIndex + 1
  x.DoAdd = False
  x.OldNumCompo = cboDesignContaminant.ListCount
  For i = 1 To x.OldNumCompo
    x.Contaminants(i).Name = bub.Contaminant(i).Name
    x.Contaminants(i).MolecularWeight.value = bub.Contaminant(i).MolecularWeight.value
    x.Contaminants(i).HenrysConstant.value = bub.Contaminant(i).HenrysConstant.value
    x.Contaminants(i).MolarVolume.value = bub.Contaminant(i).MolarVolume.value
    x.Contaminants(i).LiquidDiffusivity.value = bub.Contaminant(i).LiquidDiffusivity.value
    x.Contaminants(i).Influent.value = bub.Contaminant(i).Influent.value
    x.Contaminants(i).TreatmentObjective.value = bub.Contaminant(i).TreatmentObjective.value
  Next i

  Data_frmContaminantPropertyEdit = x
  frmContaminantPropertyEdit.Show 1
  x = Data_frmContaminantPropertyEdit

  If (Not x.CancelledEdit) Then
    For i = 1 To x.NewNumCompo
      bub.Contaminant(i).Name = x.Contaminants(i).Name
      bub.Contaminant(i).MolecularWeight.value = x.Contaminants(i).MolecularWeight.value
      bub.Contaminant(i).HenrysConstant.value = x.Contaminants(i).HenrysConstant.value
      bub.Contaminant(i).MolarVolume.value = x.Contaminants(i).MolarVolume.value
      bub.Contaminant(i).LiquidDiffusivity.value = x.Contaminants(i).LiquidDiffusivity.value
      bub.Contaminant(i).Influent.value = x.Contaminants(i).Influent.value
      bub.Contaminant(i).TreatmentObjective.value = x.Contaminants(i).TreatmentObjective.value
    Next i
    If (x.OldNumCompo < x.NewNumCompo) Then
      'Incorporate new names into ComboBox.
      For i = x.OldNumCompo + 1 To x.NewNumCompo
        cboDesignContaminant.AddItem bub.Contaminant(i).Name
      Next i
    End If
    'Update ComboBox for any changed names:
    For i = 1 To x.OldNumCompo
      If (Trim$(cboDesignContaminant.List(i - 1)) <> Trim$(bub.Contaminant(i).Name)) Then
        cboDesignContaminant.List(i - 1) = Trim$(bub.Contaminant(i).Name)
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
    'frmListContaminantBubble.Show 1
End Sub

Private Sub Command1_Click()
Call bubble_results
End Sub

Private Sub Form_Activate()
  frmBubble.WindowState = 0
  frmBubble.Width = SCREEN_WIDTH_STANDARD
  frmBubble.Height = SCREEN_HEIGHT_STANDARD

  'Center the form on the screen
  If (WindowState = 0) Then
    'don't attempt if screen Minimized or Maximized
    Move (Screen.Width - frmBubble.Width) / 2, (Screen.Height - frmBubble.Height) / 2
  End If

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
'    AddPrompt GetMenuItemID(hSubMenu, 0), "Switch modes for bubble aeration"
'    AddPrompt GetMenuItemID(hSubMenu, 2), "Load a design case from a file"
'    AddPrompt GetMenuItemID(hSubMenu, 3), "Save this design case to a file"
'    AddPrompt GetMenuItemID(hSubMenu, 4), "Save this design case to a file"
'    AddPrompt GetMenuItemID(hSubMenu, 6), "Print this design case"
'    AddPrompt GetMenuItemID(hSubMenu, 7), "Select printer for printing results"
'    AddPrompt GetMenuItemID(hSubMenu, 9), "Leave bubble aeration and return to main ASAP menu"
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
'    'Load Options Menu Prompt
'    hSubMenu = GetSubMenu(hMenu, 3)
'    AddPrompt hSubMenu, "Options"
'    AddPrompt GetMenuItemID(hSubMenu, 0), "View effluent concentration results for all contaminants"

    'Initialize last-few-files list.
    Call frmBubble_InitializeLastFewFilesList

End Sub

Private Sub Form_Load()
    Dim i As Integer

    frmBubble_Okay_To_Unload = False
    bub.NumChemical = 0
    frmBubble.WindowState = 0
    frmBubble.Width = SCREEN_WIDTH_STANDARD
    frmBubble.Height = SCREEN_HEIGHT_STANDARD
    'StatusMessagePanel.BackColor = &HC0C0C0
    'StatusMessagePanel.ForeColor = &H400040
    'StatusBarPanel.BackColor = &HC0C0C0
    'StatusBarPanel.ForeColor = &H400040
    


    Call CenterThisForm(Me)

    'Initialize Labels on frmBubble
    Call LabelsBubble(UNITSTYPE_SI)

    'Initialize Values for Pressure and Temperature
'    Call InitializePressureTemperatureBubble

    'Initialize values for water density and water viscosity
    'based on default temperature and pressure

 '   Call CalculateWaterPropertiesBubble

    'Load KLa Method Combo box
    cboOxygen.AddItem "CW O2 Transfer Test Data"
    cboOxygen.AddItem "User Input"

    'Initialize Oxygen Mass Transfer Coefficient
    Call InitializeOxygenMTCoeff

'    'Initialize value for Multiple of Minimum Air to Water Ratio
'    Call InitializeVQminMultiple

'    'Initialize Value for Air Pressure Drop
'    Call InitializeAirPressureDrop

'    'Initialize Value for KLaSafetyFactor
'    Call InitializeKLaSafetyFactor

    'Initialize calculated properties text boxes to 0
    'and disabled
'    Call InitializeCalculatedProperties

    Load frmWaterPropertiesBubble
    Load frmOxygenMassTransferCoeff

    lblConcentrationResults(2).Caption = "0"

    '
    ' DEMO SETTINGS.
    '
    Call LOCAL___Reset_DemoVersionDisablings
End Sub

Private Sub frmBubble_InitializeLastFewFilesList()

  'Initialize last-few-files list.
  Select Case BubbleAerationMode
    Case DESIGN_MODE:
      Call LastFewFiles_InitializeList(LASTFEW_WHICHAPP_ASAP, LASTFEW_ASAP_frmBubble_DESIGN)
    Case RATING_MODE:
      Call LastFewFiles_InitializeList(LASTFEW_WHICHAPP_ASAP, LASTFEW_ASAP_frmBubble_RATING)
  End Select

End Sub
Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
  If (frmBubble_Okay_To_Unload) Then
    Cancel = False
  Else
    Cancel = True
  End If
End Sub


Private Sub lblDisplayAirWaterProperties_Click()
    If HaveValue(bub.OperatingPressure.value) And HaveValue(bub.operatingtemperature.value) Then
       frmWaterPropertiesBubble.Show 1
    Else
       MsgBox "You must specify pressure and temperature before physical properties can be displayed.", MB_ICONSTOP, "Error"
    End If
End Sub

Private Sub mnuFile_Click(Index As Integer)
    Dim i As Integer
    Dim msg As String, Response As Integer

    Screen.MousePointer = 11   'Hourglass
    Select Case Index
        Case 0   'Switch Bubble Aeration Modes
            Select Case BubbleAerationMode
            Case DESIGN_MODE   'Switch to Rating Mode
                If lblStanton.Caption <> "" Then
                    'Give user option to save design mode before
                    'switching to rating mode
                    msg = "Would you like to save the parameters "
                    msg = msg + "for this design case to a file "
                    msg = msg + "before switching to Rating "
                    msg = msg + "Mode?"
                    Response = MsgBox(msg, MB_ICONquestion + MB_YESNO, "Save Current Design")
                    If Response = IDYES Then
                        Call savebubble
                    End If
                End If
                
                frmBubble.Caption = "Bubble Aeration - Rating Mode (untitled.bub)"
                frmBubble!mnuFile(0).Caption = "Switch to &Design Mode"
                BubbleAerationMode = RATING_MODE
                
                'Initialize last-few-files list.
                Call frmBubble_InitializeLastFewFilesList
                
                If frmBubble!lblStanton.Caption <> "" Then
                    bub.TankVolume.UserInput = True
                    For i = 1 To 4
                        frmBubble!txtTankParameters(i).Enabled = True
                    Next i
                    Call CalculateMinAirToWaterRatio
                Else
                    Filename$ = "TheDefaultCaseBubble"
                    If (loadbubble("") = False) Then
                      Exit Sub
                    End If
                End If
            
            Case RATING_MODE   'Switch to Design Mode
                If lblStanton.Caption <> "" Then
                    'Give user option to save rating mode before
                    'switching to design mode
                    msg = "Would you like to save the parameters "
                    msg = msg + "for this rating case to a file "
                    msg = msg + "before switching to Design "
                    msg = msg + "Mode?"
                    Response = MsgBox(msg, MB_ICONquestion + MB_YESNO, "Save Current Design")
                    If Response = IDYES Then
                        Call savebubble
                    End If
                End If
                
                frmBubble.Caption = "Bubble Aeration - Design Mode (untitled.bub)"
                frmBubble!mnuFile(0).Caption = "Switch to &Rating Mode"
                BubbleAerationMode = DESIGN_MODE
                
                'Initialize last-few-files list.
                Call frmBubble_InitializeLastFewFilesList
                
                If frmBubble!lblStanton.Caption <> "" Then
                    bub.CodeForTausAndTankVolumes = 3
                    Call CalculateMinAirToWaterRatio
                    If bub.AirToWaterRatio.value < bub.MinimumAirToWaterRatio.value Then
                        frmBubbleAchievingRemovalEfficiency!lblAchieving(0).Caption = Format$(bub.MinimumAirToWaterRatio.value, GetTheFormat(bub.MinimumAirToWaterRatio.value))
                        frmBubbleAchievingRemovalEfficiency!txtAchieving(1).Text = Format$(bub.AirToWaterRatio.value, GetTheFormat(bub.AirToWaterRatio.value))
                        frmBubbleAchievingRemovalEfficiency!txtAchieving(2).Text = Format$(bub.NumberOfTanks.value, "0")
                        frmBubbleAchievingRemovalEfficiency.Show 1
                    End If
                    Call CalculateTankVolumeBubble
                    Call CalculateRetentionTimesAndTankVolumes
                    For i = 1 To 4
                        frmBubble!txtTankParameters(i).Enabled = False
                    Next i
                    Call CalculateStantonNo
                    Call CalculateEffluentConcentrationsBubble
                Else
                    Filename$ = "TheDefaultCaseBubble"
                    If (loadbubble("") = False) Then
                      Exit Sub
                    End If
                End If
            
            End Select

        Case 3   'Open
            If frmBubble!lblStanton.Caption <> "" Then
              If (bubble_savechanges()) Then Exit Sub
            End If
             
            ''''ChDrive SaveAndLoadPath
            ''''ChDir SaveAndLoadPath
            Call ChangeDir_Main
            Call loadbubble("")

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
            Call savebubble

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
            If Right$(frmBubble.Caption, 14) <> "(untitled.bub)" Then Call savefilebubble(Filename)
            Call savebubble
            
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
            GoTo ExitSelectPrint
PrinterError:
            Resume ExitSelectPrint:
    
ExitSelectPrint:

        Case 10   'Return To Main Menu
            '********************
            If frmBubble!lblStanton.Caption <> "" Then
              If (bubble_savechanges()) Then Exit Sub
            End If
            
            'Unload Forms for Bubble Aeration
            frmBubble_Okay_To_Unload = True
            Unload frmBubble
            Unload frmBubbleEffluentConcentrations
            Unload frmOxygenMassTransferCoeff
            'Unload frmListContaminantBubble
            Unload frmBubblePower
            'Unload frmPropContaminantBubble
            Unload frmBubbleAchievingRemovalEfficiency
            Unload frmWaterPropertiesBubble
            
            frmMainMenu.Show
        
        
        Case 200   'Exit
            'Give user option to save design mode before Exiting
            If frmBubble!lblStanton.Caption <> "" Then
              If (bubble_savechanges()) Then Exit Sub
            End If

            'Unload Forms for Bubble Aeration
            frmBubble_Okay_To_Unload = True
            Unload frmBubble
            Unload frmBubbleEffluentConcentrations
            Unload frmOxygenMassTransferCoeff
            'Unload frmListContaminantBubble
            Unload frmBubblePower
            'Unload frmPropContaminantBubble
            Unload frmBubbleAchievingRemovalEfficiency
            Unload frmWaterPropertiesBubble
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
        Call loadbubble(Current_LastFewFilesRec.FileNames(Index - 190))
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


End Sub

Private Sub mnuFilePrint_Click(Index As Integer)

    Select Case Index
       Case 0   'Print to printer
          Call PrintBubble
       Case 1   'Print to file
          Call PrintBubbleToFile
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
Call bubble_results
End Sub

Private Sub mnuotheritem_Click()
'frmabout2.Show 1
End Sub

Private Sub mnuPopContaminant_Click(Index As Integer)
    'If Index = 1 Then
    '   frmListContaminantBubble.Show 1
    'End If
End Sub

Private Sub mnuPower_Click(Index As Integer)
    Dim CalculatedPower As Integer

    Select Case Index   'Power Calculation

       Case 0
             Call SetPowerBubble
             frmBubblePower.Left = Screen.Width / 2 - frmBubblePower.Width / 2
             frmBubblePower.Top = Screen.Height / 2 - frmBubblePower.Height / 2
             frmBubblePower.Show 1
    End Select

End Sub

Private Sub mnuUnits_Click(Index As Integer)

  Select Case Index
    Case 0        'SI
      Call LabelsBubble(UNITSTYPE_SI)
    Case 1        'English
      Call LabelsBubble(UNITSTYPE_ENGLISH)
  End Select

End Sub

Private Sub MsgHook1_Message(msg As Integer, wParam As Integer, lParam As Long, Action As Integer, result As Long)
'    Dim i       As Integer
'    Dim found   As Integer
'
'Exit Sub
'
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
' 'temp kill
'Exit Sub
''''''''''''''
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
                   bub.ContaminantMassTransferCoefficient.value = Dummy
                   bub.ContaminantMassTransferCoefficient.ValChanged = True
                   bub.ContaminantMassTransferCoefficient.UserInput = True
                Else
                   txtDesignConcentrationValue(3).Text = Temp_Text
                   txtDesignConcentrationValue(3).SetFocus
                   Exit Sub
                End If
             End If

       End Select

       If ValueChanged Then
          If BubbleAerationMode = DESIGN_MODE Then
             Call CalculateTankVolumeBubble
             Call CalculateRetentionTimesAndTankVolumes
          End If

          Call CalculateStantonNo
          Call CalculateEffluentConcentrationsBubble
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
  Dim bypass_it As Integer

   If (LostFocus_IsEvil(Me, txtFlowParameters(Index))) Then
     Exit Sub
   End If
   
   flag_ok = True
  

  IsNew = False
  
  Select Case Index
    Case 0        'Water Flow Rate.
      If (Unitted_LostFocus(UNITS_FLOW, txtFlowParameters(0), UnitsFlowParam(0), NewVal, Temp_Text)) Then
        IsNew = True
        bub.WaterFlowRate.ValChanged = True
        bub.WaterFlowRate.UserInput = True
        bub.WaterFlowRate.value = NewVal
        If bub.AirToWaterRatio.UserInput = True Then
           Call CalculateAirFlowRate
        Else
           Call CalculateAirToWaterRatio
        End If
        Call CalculateRetentionTimesAndTankVolumes
      End If

    Case 2        'Air to Water Ratio.
      If (NoUnits_LostFocus(txtFlowParameters(2), NewVal, Temp_Text)) Then
        IsNew = True
        bub.AirToWaterRatio.ValChanged = True
        bub.AirToWaterRatio.UserInput = True
        bub.AirToWaterRatio.value = NewVal
        Call CalculateAirFlowRate
      End If
    
    Case 3        'Air Flow Rate.
      If (Unitted_LostFocus(UNITS_FLOW, txtFlowParameters(3), UnitsFlowParam(3), NewVal, Temp_Text)) Then
        IsNew = True
        bub.AirFlowRate.ValChanged = True
        bub.AirFlowRate.UserInput = True
        bub.AirFlowRate.value = NewVal
        Call CalculateAirToWaterRatio
      End If

  End Select

  If (IsNew) Then
    bypass_it = False
    If BubbleAerationMode = DESIGN_MODE Then
      If bub.AirToWaterRatio.value < bub.MinimumAirToWaterRatio.value Then
        Call LostFocus_Handle(Me, txtFlowParameters(Index), flag_ok)
        frmBubbleAchievingRemovalEfficiency!lblAchieving(0).Caption = Format$(bub.MinimumAirToWaterRatio.value, GetTheFormat(bub.MinimumAirToWaterRatio.value))
        frmBubbleAchievingRemovalEfficiency!txtAchieving(1).Text = Format$(bub.AirToWaterRatio.value, GetTheFormat(bub.AirToWaterRatio.value))
        frmBubbleAchievingRemovalEfficiency!txtAchieving(2).Text = Format$(bub.NumberOfTanks.value, "0")
        frmBubbleAchievingRemovalEfficiency.Show 1
        bypass_it = True
      End If
      Call CalculateTankVolumeBubble
      Call CalculateRetentionTimesAndTankVolumes
    End If
    Call CalculateStantonNo
    Call CalculateEffluentConcentrationsBubble
    Call cboDesignContaminant_Click
    If (bypass_it) Then Exit Sub
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
On Error GoTo err_ThisFunc
   If (LostFocus_IsEvil(Me, txtOperatingPressure)) Then
     Exit Sub
   End If
   
   flag_ok = True
  
  If (Unitted_LostFocus(UNITS_PRESSURE, txtOperatingPressure, UnitsOpCond(0), NewVal, Temp_Text)) Then
    bub.OperatingPressure.ValChanged = True
    bub.OperatingPressure.UserInput = True
    'Note: standard P units are Pa, but
    'OperatingPressure is stored as kPa.
    bub.OperatingPressure.value = NewVal * 1# / 101325#

    If (HaveValue(bub.OperatingPressure.value) And HaveValue(bub.operatingtemperature.value)) Then
      Call CalculateWaterPropertiesBubble
      If (bub.NumChemical > 0) Then
        'Update Variables on Screen
      End If
    End If
    Call cboDesignContaminant_Click
  End If

  Call LostFocus_Handle(Me, txtOperatingPressure, flag_ok)


exit_normally_ThisFunc:
  'txtOperatingPressure_LostFocus = True
  Exit Sub
exit_err_ThisFunc:
  'txtOperatingPressure_LostFocus = False
  Exit Sub
err_ThisFunc:
  ''''Call Show_Trapped_Error("txtOperatingPressure_LostFocus")
  Resume exit_err_ThisFunc
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
    bub.operatingtemperature.ValChanged = True
    bub.operatingtemperature.UserInput = True
    bub.operatingtemperature.value = NewVal

    If (HaveValue(bub.OperatingPressure.value) And HaveValue(bub.operatingtemperature.value)) Then
      Call CalculateWaterPropertiesBubble
      Call CalculateOxygenLiquidDiffusivity
      If (cboOxygen.ListIndex = 0) Then
        Call CalculateTrueKLa
        bub.Oxygen.MassTransferCoefficient.value = bub.Oxygen.CWO2TestData.TrueOxygenMTCoeffOperatingT_KLAO2
        frmBubble!txtOxygen(2).Text = Format$(bub.Oxygen.MassTransferCoefficient.value, GetTheFormat(bub.Oxygen.MassTransferCoefficient.value))
      End If
         
      If (bub.NumChemical > 0) Then
        Call CalculateContaminantMTCoeff
        If (BubbleAerationMode = DESIGN_MODE) Then
          Call CalculateTankVolumeBubble
          Call CalculateRetentionTimesAndTankVolumes
        End If
        Call CalculateStantonNo
        Call CalculateEffluentConcentrationsBubble
      End If
    End If
    Call cboDesignContaminant_Click
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
        bub.Oxygen.LiquidDiffusivity.value = NewVal
        bub.Oxygen.LiquidDiffusivity.ValChanged = True
        bub.Oxygen.LiquidDiffusivity.UserInput = True
      End If

    Case 2        'KLa.
      If (Unitted_LostFocus(UNITS_INVERSETIME, txtOxygen(2), UnitsOxygenRef(2), NewVal, Temp_Text)) Then
        IsNew = True
        bub.Oxygen.MassTransferCoefficient.value = NewVal
        bub.Oxygen.MassTransferCoefficient.ValChanged = True
        bub.Oxygen.MassTransferCoefficient.UserInput = True
      End If
    
  End Select

  If (IsNew) Then
    Select Case Index
      Case 1
        If HaveValue(bub.OperatingPressure.value) And HaveValue(bub.operatingtemperature.value) And HaveValue(bub.Oxygen.LiquidDiffusivity.value) Then
          If bub.NumChemical > 0 Then
            Call CalculateContaminantMTCoeff
            If BubbleAerationMode = DESIGN_MODE Then
              Call CalculateTankVolumeBubble
              Call CalculateRetentionTimesAndTankVolumes
            End If
            Call CalculateStantonNo
            Call CalculateEffluentConcentrationsBubble
          End If
        End If
      
      Case 2
        If HaveValue(bub.OperatingPressure.value) And HaveValue(bub.operatingtemperature.value) And HaveValue(bub.Oxygen.LiquidDiffusivity.value) And HaveValue(bub.Oxygen.MassTransferCoefficient.value) Then
          If bub.NumChemical > 0 Then
            Call CalculateContaminantMTCoeff
            If BubbleAerationMode = DESIGN_MODE Then
              Call CalculateTankVolumeBubble
              Call CalculateRetentionTimesAndTankVolumes
            End If
            Call CalculateStantonNo
            Call CalculateEffluentConcentrationsBubble
          End If
        End If
        If cboOxygen.ListIndex = 0 Then cboOxygen.ListIndex = 1
    
    End Select
  End If

  Call LostFocus_Handle(Me, txtOxygen(Index), flag_ok)

End Sub

Private Sub txtTankParameters_Change(Index As Integer)

  'If (Index >= 1 And Index <= 4) Then
  '  Call UnitsTankParam_Click(Index)
  'End If
  
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
          Call LostFocus_Handle(Me, txtTankParameters(Index), flag_ok)
          Exit Sub
        End If

        bub.NumberOfTanks.ValChanged = True
        bub.NumberOfTanks.UserInput = True
        bub.NumberOfTanks.value = NewVal
        Call CalculateMinAirToWaterRatio

        If BubbleAerationMode = DESIGN_MODE Then
          If bub.AirToWaterRatio.value < bub.MinimumAirToWaterRatio.value Then
            frmBubbleAchievingRemovalEfficiency!lblAchieving(0).Caption = Format$(bub.MinimumAirToWaterRatio.value, GetTheFormat(bub.MinimumAirToWaterRatio.value))
            frmBubbleAchievingRemovalEfficiency!txtAchieving(1).Text = Format$(bub.AirToWaterRatio.value, GetTheFormat(bub.AirToWaterRatio.value))
            frmBubbleAchievingRemovalEfficiency!txtAchieving(2).Text = Format$(bub.NumberOfTanks.value, "0")
            
            Call LostFocus_Handle(Me, txtTankParameters(Index), flag_ok)
            
            frmBubbleAchievingRemovalEfficiency.Show 1
          End If
          Call CalculateTankVolumeBubble
          Call CalculateRetentionTimesAndTankVolumes
        Else
          Call CalculateRetentionTimesAndTankVolumes
        End If
      End If

    Case 1        'Retention Time for 1 Tank.
      If (Unitted_LostFocus(UNITS_TIME, txtTankParameters(1), UnitsTankParam(1), NewVal, Temp_Text)) Then
        IsNew = True
        bub.TankHydraulicRetentionTime.ValChanged = True
        bub.TankHydraulicRetentionTime.UserInput = True
        'Standard time units are seconds, but TankHydraulicRetentionTime
        'is stored internally as hours.
        bub.TankHydraulicRetentionTime.value = NewVal / 60# / 60#
        bub.CodeForTausAndTankVolumes = 1
        Call CalculateRetentionTimesAndTankVolumes
      End If
        
    Case 2        'Retention Time for All Tanks.
      If (Unitted_LostFocus(UNITS_TIME, txtTankParameters(2), UnitsTankParam(2), NewVal, Temp_Text)) Then
        IsNew = True
        bub.TotalHydraulicRetentionTime.ValChanged = True
        bub.TotalHydraulicRetentionTime.UserInput = True
        'Standard time units are seconds, but TotalHydraulicRetentionTime
        'is stored internally as hours.
        bub.TotalHydraulicRetentionTime.value = NewVal / 60# / 60#
        bub.CodeForTausAndTankVolumes = 2
        Call CalculateRetentionTimesAndTankVolumes
      End If

    Case 3        'Volume of Each Tank.
      If (Unitted_LostFocus(UNITS_VOLUME, txtTankParameters(3), UnitsTankParam(3), NewVal, Temp_Text)) Then
        IsNew = True
        bub.TankVolume.ValChanged = True
        bub.TankVolume.UserInput = True
        bub.TankVolume.value = NewVal
        Call CalculateRetentionTimesAndTankVolumes
      End If

    Case 4        'Volume of All Tanks.
      If (Unitted_LostFocus(UNITS_VOLUME, txtTankParameters(4), UnitsTankParam(4), NewVal, Temp_Text)) Then
        IsNew = True
        bub.TotalTankVolume.ValChanged = True
        bub.TotalTankVolume.UserInput = True
        bub.TotalTankVolume.value = NewVal
        bub.TankVolume.ValChanged = True
        bub.TankVolume.UserInput = True
        bub.TankVolume.value = bub.TotalTankVolume.value / bub.NumberOfTanks.value
        
        Call CalculateRetentionTimesAndTankVolumes
      End If

  End Select

  If (IsNew) Then
    Call CalculateStantonNo

    Call CalculateEffluentConcentrationsBubble
    'Update everything:
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
      Dummy = bub.Contaminant(i).Influent.value
      Call Unitted_UnitChange(UNITS_CONCENTRATION, Dummy, UnitsConcResults(1), lblConcentrationResults(1))

    Case 2            'Yi to All Tanks.
      Dummy = 0
      Call Unitted_UnitChange(UNITS_CONCENTRATION, Dummy, UnitsConcResults(2), lblConcentrationResults(2))

    Case 3            'Ce from Last Tank.
      Dummy = bub.DesignContaminant.Effluent(bub.NumberOfTanks.value)
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
      Dummy = bub.Contaminant(i).Influent.value
      Call Unitted_UnitChange(UNITS_CONCENTRATION, Dummy, UnitsDesignContam(0), lblDesignConcentrationValue(0))

    Case 1            'Treatment Obj.
      Dummy = bub.Contaminant(i).TreatmentObjective.value
      Call Unitted_UnitChange(UNITS_CONCENTRATION, Dummy, UnitsDesignContam(1), lblDesignConcentrationValue(1))

    Case 3            'KLa
      Dummy = bub.ContaminantMassTransferCoefficient.value
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
      Call Unitted_UnitChange(UNITS_FLOW, bub.WaterFlowRate.value, UnitsFlowParam(0), txtFlowParameters(0))

    Case 3            'Air Flow Rate.
      Call Unitted_UnitChange(UNITS_FLOW, bub.AirFlowRate.value, UnitsFlowParam(3), txtFlowParameters(3))
    
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
      Dummy = bub.OperatingPressure.value * 101325#
      Call Unitted_UnitChange(UNITS_PRESSURE, Dummy, UnitsOpCond(0), txtOperatingPressure)

    Case 1
      Call Unitted_UnitChange(UNITS_TEMPERATURE, bub.operatingtemperature.value, UnitsOpCond(1), txtOperatingTemperature)
    
  End Select

End Sub

Private Sub UnitsOpCond_KeyPress(Index As Integer, KeyAscii As Integer)
  Call Combobox_KeyPress(KeyAscii)
End Sub

Private Sub UnitsOxygenRef_Click(Index As Integer)

  Select Case Index
    Case 1            'Liquid Diffusivity
      Call Unitted_UnitChange(UNITS_DIFFUSIVITY, bub.Oxygen.LiquidDiffusivity.value, UnitsOxygenRef(1), txtOxygen(1))

    Case 2            'KLa
      Call Unitted_UnitChange(UNITS_INVERSETIME, bub.Oxygen.MassTransferCoefficient.value, UnitsOxygenRef(2), txtOxygen(2))
    
  End Select

End Sub

Private Sub UnitsOxygenRef_KeyPress(Index As Integer, KeyAscii As Integer)
  Call Combobox_KeyPress(KeyAscii)
End Sub

Private Sub UnitsTankParam_Click(Index As Integer)
Dim Dummy As Double

  Select Case Index
    Case 1            'Retention Time (1 Tank).
      'Standard time units are seconds, but TankHydraulicRetentionTime
      'is stored internally as hours.
      Dummy = bub.TankHydraulicRetentionTime.value * 60 * 60#
      Call Unitted_UnitChange(UNITS_TIME, Dummy, UnitsTankParam(1), txtTankParameters(1))

    Case 2            'Retention Time (All Tanks).
      'Standard time units are seconds, but TotalHydraulicRetentionTime
      'is stored internally as hours.
      Dummy = bub.TotalHydraulicRetentionTime.value * 60 * 60#
      Call Unitted_UnitChange(UNITS_TIME, Dummy, UnitsTankParam(2), txtTankParameters(2))

    Case 3            'Volume (1 Tank).
      Call Unitted_UnitChange(UNITS_VOLUME, bub.TankVolume.value, UnitsTankParam(3), txtTankParameters(3))

    Case 4            'Volume (All Tanks).
      Call Unitted_UnitChange(UNITS_VOLUME, bub.TotalTankVolume.value, UnitsTankParam(4), txtTankParameters(4))
    
  End Select

End Sub

Private Sub UnitsTankParam_KeyPress(Index As Integer, KeyAscii As Integer)
  Call Combobox_KeyPress(KeyAscii)
End Sub


