VERSION 5.00
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "THREED32.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{042BADC8-5E58-11CE-B610-524153480001}#1.0#0"; "VCF132.OCX"
Begin VB.Form frmMain 
   Appearance      =   0  'Flat
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "{Generic Application -- Me.Caption set as Name_App_Short}"
   ClientHeight    =   6870
   ClientLeft      =   405
   ClientTop       =   2205
   ClientWidth     =   11130
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
   HelpContextID   =   1000
   Icon            =   "Main.frx":0000
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   6870
   ScaleWidth      =   11130
   Begin VB.ComboBox cboUnitsOutput 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Index           =   3
      Left            =   9000
      Style           =   2  'Dropdown List
      TabIndex        =   67
      TabStop         =   0   'False
      Top             =   5865
      Width           =   1485
   End
   Begin VB.ComboBox cboUnitsOutput 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Index           =   2
      Left            =   9000
      Style           =   2  'Dropdown List
      TabIndex        =   66
      TabStop         =   0   'False
      Top             =   5505
      Width           =   1485
   End
   Begin VB.ComboBox cboUnitsOutput 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Index           =   1
      Left            =   4050
      Style           =   2  'Dropdown List
      TabIndex        =   65
      TabStop         =   0   'False
      Top             =   5865
      Width           =   1485
   End
   Begin Threed.SSPanel SSPanel2 
      Align           =   2  'Align Bottom
      Height          =   405
      Left            =   0
      TabIndex        =   0
      Top             =   6465
      Width           =   11130
      _Version        =   65536
      _ExtentX        =   19632
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
      Begin Threed.SSPanel sspanel_Status 
         Height          =   285
         Left            =   2220
         TabIndex        =   2
         Top             =   60
         Width           =   8805
         _Version        =   65536
         _ExtentX        =   15531
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
      Begin Threed.SSPanel sspanel_Dirty 
         Height          =   285
         Left            =   60
         TabIndex        =   1
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
   End
   Begin Threed.SSFrame SSFrame2 
      Height          =   2295
      Left            =   6180
      TabIndex        =   52
      Top             =   6240
      Visible         =   0   'False
      Width           =   8865
      _Version        =   65536
      _ExtentX        =   15637
      _ExtentY        =   4048
      _StockProps     =   14
      Caption         =   "Unused -- Invisible"
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
      Begin VB.ComboBox cboUnitsOutput 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   0
         Left            =   2280
         Style           =   2  'Dropdown List
         TabIndex        =   68
         TabStop         =   0   'False
         Top             =   180
         Width           =   1485
      End
      Begin VB.TextBox txtOutput 
         Alignment       =   2  'Center
         BackColor       =   &H8000000F&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00008000&
         Height          =   300
         Index           =   7
         Left            =   2040
         Locked          =   -1  'True
         TabIndex        =   57
         Text            =   "txtOutput(0)"
         Top             =   1050
         Width           =   2265
      End
      Begin VB.TextBox txtOutput 
         Alignment       =   2  'Center
         BackColor       =   &H8000000F&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   300
         Index           =   6
         Left            =   2040
         Locked          =   -1  'True
         TabIndex        =   56
         Text            =   "txtOutput(1)"
         Top             =   1410
         Width           =   1395
      End
      Begin VB.TextBox txtOutput 
         Alignment       =   2  'Center
         BackColor       =   &H8000000F&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   300
         Index           =   5
         Left            =   6510
         Locked          =   -1  'True
         TabIndex        =   55
         Text            =   "txtOutput(2)"
         Top             =   1050
         Width           =   1395
      End
      Begin VB.TextBox txtOutput 
         Alignment       =   2  'Center
         BackColor       =   &H8000000F&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   300
         Index           =   4
         Left            =   6510
         Locked          =   -1  'True
         TabIndex        =   54
         Text            =   "txtOutput(3)"
         Top             =   1410
         Width           =   1395
      End
      Begin VB.Label lblOutput 
         Alignment       =   1  'Right Justify
         Caption         =   "Contaminant:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00008000&
         Height          =   285
         Index           =   7
         Left            =   120
         TabIndex        =   64
         Top             =   1095
         Width           =   1845
      End
      Begin VB.Label lblOutput 
         Alignment       =   1  'Right Justify
         Caption         =   "Total Amount Influent:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   285
         Index           =   6
         Left            =   120
         TabIndex        =   63
         Top             =   1455
         Width           =   1845
      End
      Begin VB.Label lblOutputUnits 
         Alignment       =   2  'Center
         Caption         =   "(kg/day)"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   285
         Index           =   6
         Left            =   3480
         TabIndex        =   62
         Top             =   1455
         Width           =   885
      End
      Begin VB.Label lblOutput 
         Alignment       =   1  'Right Justify
         Caption         =   "Total Amount Effluent:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   285
         Index           =   5
         Left            =   4590
         TabIndex        =   61
         Top             =   1095
         Width           =   1845
      End
      Begin VB.Label lblOutputUnits 
         Alignment       =   2  'Center
         Caption         =   "(kg/day)"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   285
         Index           =   5
         Left            =   7950
         TabIndex        =   60
         Top             =   1095
         Width           =   885
      End
      Begin VB.Label lblOutput 
         Alignment       =   1  'Right Justify
         Caption         =   "Total Percent Removed:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   285
         Index           =   4
         Left            =   4590
         TabIndex        =   59
         Top             =   1455
         Width           =   1845
      End
      Begin VB.Label lblOutputUnits 
         Alignment       =   2  'Center
         Caption         =   "(kg/day)"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   285
         Index           =   4
         Left            =   7950
         TabIndex        =   58
         Top             =   1455
         Width           =   885
      End
      Begin VB.Label lblOutputUnits 
         Alignment       =   2  'Center
         Caption         =   "(kg/day)"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00008000&
         Height          =   285
         Index           =   0
         Left            =   480
         TabIndex        =   53
         Top             =   570
         Width           =   885
      End
   End
   Begin VB.TextBox txtOutput 
      Alignment       =   2  'Center
      BackColor       =   &H8000000F&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   300
      Index           =   3
      Left            =   7560
      Locked          =   -1  'True
      TabIndex        =   50
      Text            =   "txtOutput(3)"
      Top             =   5880
      Width           =   1395
   End
   Begin VB.TextBox txtOutput 
      Alignment       =   2  'Center
      BackColor       =   &H8000000F&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   300
      Index           =   2
      Left            =   7560
      Locked          =   -1  'True
      TabIndex        =   48
      Text            =   "txtOutput(2)"
      Top             =   5520
      Width           =   1395
   End
   Begin VB.TextBox txtOutput 
      Alignment       =   2  'Center
      BackColor       =   &H8000000F&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   300
      Index           =   1
      Left            =   2610
      Locked          =   -1  'True
      TabIndex        =   46
      Text            =   "txtOutput(1)"
      Top             =   5880
      Width           =   1395
   End
   Begin VB.TextBox txtOutput 
      Alignment       =   2  'Center
      BackColor       =   &H8000000F&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00008000&
      Height          =   300
      Index           =   0
      Left            =   2610
      Locked          =   -1  'True
      TabIndex        =   44
      Text            =   "txtOutput(0)"
      Top             =   5520
      Width           =   2265
   End
   Begin VB.ComboBox cboUnits 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Index           =   5
      Left            =   9150
      Style           =   2  'Dropdown List
      TabIndex        =   42
      TabStop         =   0   'False
      Top             =   1710
      Width           =   1300
   End
   Begin VB.TextBox txtData 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Index           =   5
      Left            =   7710
      TabIndex        =   41
      Text            =   "txtData(5)"
      Top             =   1725
      Width           =   1395
   End
   Begin VB.ComboBox cboUnits 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Index           =   4
      Left            =   3870
      Style           =   2  'Dropdown List
      TabIndex        =   39
      TabStop         =   0   'False
      Top             =   1710
      Width           =   1300
   End
   Begin VB.TextBox txtData 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Index           =   4
      Left            =   2430
      TabIndex        =   38
      Text            =   "txtData(4)"
      Top             =   1725
      Width           =   1395
   End
   Begin Threed.SSPanel sspanel_GridHolder 
      Height          =   3375
      Left            =   60
      TabIndex        =   36
      Top             =   2100
      Width           =   10995
      _Version        =   65536
      _ExtentX        =   19394
      _ExtentY        =   5953
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
      Begin VCIF1Lib.F1Book F1Book1 
         Height          =   3195
         Left            =   60
         OleObjectBlob   =   "Main.frx":030A
         TabIndex        =   37
         Top             =   30
         Width           =   10815
      End
   End
   Begin Threed.SSPanel sspanel_ButtonHolder 
      Height          =   1695
      Left            =   60
      TabIndex        =   17
      Top             =   0
      Width           =   10965
      _Version        =   65536
      _ExtentX        =   19341
      _ExtentY        =   2990
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
      Begin Threed.SSCommand cmdUnitChange 
         Height          =   405
         Index           =   0
         Left            =   9420
         TabIndex        =   69
         Top             =   90
         Width           =   1425
         _Version        =   65536
         _ExtentX        =   2514
         _ExtentY        =   714
         _StockProps     =   78
         Caption         =   "&English Units"
         Outline         =   0   'False
      End
      Begin Threed.SSCommand cmdMainButton 
         Height          =   810
         Index           =   8
         Left            =   8280
         TabIndex        =   18
         Top             =   90
         Width           =   810
         _Version        =   65536
         _ExtentX        =   1429
         _ExtentY        =   1429
         _StockProps     =   78
         Outline         =   0   'False
         AutoSize        =   2
         Picture         =   "Main.frx":1D93
      End
      Begin Threed.SSCommand cmdMainButton 
         Height          =   810
         Index           =   0
         Left            =   90
         TabIndex        =   19
         Top             =   90
         Width           =   810
         _Version        =   65536
         _ExtentX        =   1429
         _ExtentY        =   1429
         _StockProps     =   78
         Outline         =   0   'False
         AutoSize        =   2
         Picture         =   "Main.frx":2AE5
      End
      Begin Threed.SSCommand cmdMainButton 
         Height          =   810
         Index           =   5
         Left            =   5190
         TabIndex        =   20
         Top             =   90
         Width           =   810
         _Version        =   65536
         _ExtentX        =   1429
         _ExtentY        =   1429
         _StockProps     =   78
         Outline         =   0   'False
         AutoSize        =   2
         Picture         =   "Main.frx":2FF7
      End
      Begin Threed.SSCommand cmdMainButton 
         Height          =   810
         Index           =   6
         Left            =   6150
         TabIndex        =   21
         Top             =   90
         Width           =   810
         _Version        =   65536
         _ExtentX        =   1429
         _ExtentY        =   1429
         _StockProps     =   78
         Outline         =   0   'False
         AutoSize        =   2
         Picture         =   "Main.frx":3509
      End
      Begin Threed.SSCommand cmdMainButton 
         Height          =   810
         Index           =   3
         Left            =   3270
         TabIndex        =   22
         Top             =   90
         Width           =   810
         _Version        =   65536
         _ExtentX        =   1429
         _ExtentY        =   1429
         _StockProps     =   78
         Outline         =   0   'False
         AutoSize        =   2
         Picture         =   "Main.frx":3A1B
      End
      Begin Threed.SSCommand cmdMainButton 
         Height          =   810
         Index           =   7
         Left            =   7110
         TabIndex        =   23
         Top             =   90
         Width           =   810
         _Version        =   65536
         _ExtentX        =   1429
         _ExtentY        =   1429
         _StockProps     =   78
         Outline         =   0   'False
         AutoSize        =   2
         Picture         =   "Main.frx":3F2D
      End
      Begin Threed.SSCommand cmdMainButton 
         Height          =   810
         Index           =   4
         Left            =   4230
         TabIndex        =   24
         Top             =   90
         Width           =   810
         _Version        =   65536
         _ExtentX        =   1429
         _ExtentY        =   1429
         _StockProps     =   78
         Outline         =   0   'False
         AutoSize        =   2
         Picture         =   "Main.frx":443F
      End
      Begin Threed.SSCommand cmdMainButton 
         Height          =   810
         Index           =   2
         Left            =   2310
         TabIndex        =   25
         Top             =   90
         Width           =   810
         _Version        =   65536
         _ExtentX        =   1429
         _ExtentY        =   1429
         _StockProps     =   78
         Outline         =   0   'False
         AutoSize        =   2
         Picture         =   "Main.frx":4951
      End
      Begin Threed.SSCommand cmdMainButton 
         Height          =   810
         Index           =   1
         Left            =   1350
         TabIndex        =   26
         Top             =   90
         Width           =   810
         _Version        =   65536
         _ExtentX        =   1429
         _ExtentY        =   1429
         _StockProps     =   78
         Outline         =   0   'False
         AutoSize        =   2
         Picture         =   "Main.frx":4E63
      End
      Begin Threed.SSCommand cmdUnitChange 
         Height          =   405
         Index           =   1
         Left            =   9420
         TabIndex        =   71
         Top             =   480
         Width           =   1425
         _Version        =   65536
         _ExtentX        =   2514
         _ExtentY        =   714
         _StockProps     =   78
         Caption         =   "&S.I. Units"
         Outline         =   0   'False
      End
      Begin VB.Label lblButton 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         Caption         =   "Change Units"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   195
         Index           =   9
         Left            =   9615
         TabIndex        =   70
         Top             =   990
         Width           =   1035
         WordWrap        =   -1  'True
      End
      Begin VB.Label lblButton 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         Caption         =   "Influent Weir*"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   195
         Index           =   1
         Left            =   1275
         TabIndex        =   35
         Top             =   990
         Width           =   1005
         WordWrap        =   -1  'True
      End
      Begin VB.Label lblButton 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         Caption         =   "Aerated Grit Chamber*"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   390
         Index           =   2
         Left            =   2265
         TabIndex        =   34
         Top             =   990
         Width           =   1005
         WordWrap        =   -1  'True
      End
      Begin VB.Label lblButton 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         Caption         =   "Primary Clarifier"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   390
         Index           =   3
         Left            =   3195
         TabIndex        =   33
         Top             =   990
         Width           =   1005
         WordWrap        =   -1  'True
      End
      Begin VB.Label lblButton 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         Caption         =   "Primary Clarifier  Weir*"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   585
         Index           =   4
         Left            =   4200
         TabIndex        =   32
         Top             =   990
         Width           =   1005
         WordWrap        =   -1  'True
      End
      Begin VB.Label lblButton 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         Caption         =   "Aeration Basin(s)"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   390
         Index           =   5
         Left            =   5130
         TabIndex        =   31
         Top             =   990
         Width           =   1005
         WordWrap        =   -1  'True
      End
      Begin VB.Label lblButton 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         Caption         =   "Secondary Clarifier"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   390
         Index           =   6
         Left            =   6060
         TabIndex        =   30
         Top             =   990
         Width           =   1005
         WordWrap        =   -1  'True
      End
      Begin VB.Label lblButton 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         Caption         =   "Secondary Clarifier Weir*"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   390
         Index           =   7
         Left            =   7050
         TabIndex        =   29
         Top             =   990
         Width           =   1005
         WordWrap        =   -1  'True
      End
      Begin VB.Label lblButton 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         Caption         =   "Physico- Chemical Properties"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   585
         Index           =   0
         Left            =   75
         TabIndex        =   28
         Top             =   990
         Width           =   915
         WordWrap        =   -1  'True
      End
      Begin VB.Label lblButton 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         Caption         =   "Calculate!"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   195
         Index           =   8
         Left            =   8265
         TabIndex        =   27
         Top             =   990
         Width           =   915
         WordWrap        =   -1  'True
      End
   End
   Begin Threed.SSFrame SSFrame1 
      Height          =   2175
      Left            =   270
      TabIndex        =   4
      Top             =   6240
      Visible         =   0   'False
      Width           =   5535
      _Version        =   65536
      _ExtentX        =   9763
      _ExtentY        =   3836
      _StockProps     =   14
      Caption         =   "Test Frame"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Begin VB.ComboBox cboUnits 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Index           =   3
         Left            =   3510
         Style           =   2  'Dropdown List
         TabIndex        =   15
         TabStop         =   0   'False
         Top             =   1650
         Width           =   1545
      End
      Begin VB.TextBox txtData 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Index           =   3
         Left            =   1440
         TabIndex        =   14
         Text            =   "txtData(3)"
         Top             =   1680
         Width           =   1995
      End
      Begin VB.ComboBox cboUnits 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Index           =   2
         Left            =   3510
         Style           =   2  'Dropdown List
         TabIndex        =   12
         TabStop         =   0   'False
         Top             =   1230
         Width           =   1545
      End
      Begin VB.TextBox txtData 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Index           =   2
         Left            =   1440
         TabIndex        =   11
         Text            =   "txtData(2)"
         Top             =   1260
         Width           =   1995
      End
      Begin VB.ComboBox cboUnits 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Index           =   1
         Left            =   3510
         Style           =   2  'Dropdown List
         TabIndex        =   9
         TabStop         =   0   'False
         Top             =   810
         Width           =   1545
      End
      Begin VB.TextBox txtData 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Index           =   1
         Left            =   1440
         TabIndex        =   8
         Text            =   "txtData(1)"
         Top             =   840
         Width           =   1995
      End
      Begin VB.TextBox txtData 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Index           =   0
         Left            =   1440
         TabIndex        =   6
         Text            =   "txtData(0)"
         Top             =   420
         Width           =   1995
      End
      Begin VB.ComboBox cboUnits 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Index           =   0
         Left            =   3510
         Style           =   2  'Dropdown List
         TabIndex        =   5
         TabStop         =   0   'False
         Top             =   390
         Width           =   1545
      End
      Begin VB.Label lblData 
         Alignment       =   1  'Right Justify
         Caption         =   "Flow Rate:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   3
         Left            =   -510
         TabIndex        =   16
         Top             =   1710
         Width           =   1845
      End
      Begin VB.Label lblData 
         Alignment       =   1  'Right Justify
         Caption         =   "Mass:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   2
         Left            =   -510
         TabIndex        =   13
         Top             =   1290
         Width           =   1845
      End
      Begin VB.Label lblData 
         Alignment       =   1  'Right Justify
         Caption         =   "Diameter:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   1
         Left            =   -510
         TabIndex        =   10
         Top             =   870
         Width           =   1845
      End
      Begin VB.Label lblData 
         Alignment       =   1  'Right Justify
         Caption         =   "Length:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   0
         Left            =   -510
         TabIndex        =   7
         Top             =   450
         Width           =   1845
      End
   End
   Begin Threed.SSFrame SSFrame4 
      Height          =   1035
      Left            =   10590
      TabIndex        =   3
      Top             =   5820
      Visible         =   0   'False
      Width           =   3015
      _Version        =   65536
      _ExtentX        =   5318
      _ExtentY        =   1826
      _StockProps     =   14
      Caption         =   "Used -- Invisible"
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
      Begin MSComDlg.CommonDialog CommonDialog1 
         Left            =   240
         Top             =   240
         _ExtentX        =   847
         _ExtentY        =   847
         _Version        =   393216
      End
   End
   Begin VB.Label lblOutput 
      Alignment       =   1  'Right Justify
      Caption         =   "Total Percent Removed:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   285
      Index           =   3
      Left            =   5640
      TabIndex        =   51
      Top             =   5925
      Width           =   1845
   End
   Begin VB.Label lblOutput 
      Alignment       =   1  'Right Justify
      Caption         =   "Total Amount Effluent:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   285
      Index           =   2
      Left            =   5640
      TabIndex        =   49
      Top             =   5565
      Width           =   1845
   End
   Begin VB.Label lblOutput 
      Alignment       =   1  'Right Justify
      Caption         =   "Total Amount Influent:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   285
      Index           =   1
      Left            =   690
      TabIndex        =   47
      Top             =   5925
      Width           =   1845
   End
   Begin VB.Label lblOutput 
      Alignment       =   1  'Right Justify
      Caption         =   "Contaminant:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00008000&
      Height          =   285
      Index           =   0
      Left            =   690
      TabIndex        =   45
      Top             =   5565
      Width           =   1845
   End
   Begin VB.Label lblData 
      Alignment       =   1  'Right Justify
      Caption         =   "Influent Solids Concentration:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   5
      Left            =   5280
      TabIndex        =   43
      Top             =   1770
      Width           =   2385
   End
   Begin VB.Label lblData 
      Alignment       =   1  'Right Justify
      Caption         =   "Influent Plant Flow Rate:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   4
      Left            =   540
      TabIndex        =   40
      Top             =   1770
      Width           =   1845
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
         Caption         =   "&Print ..."
         Index           =   85
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
   Begin VB.Menu mnuPlant 
      Caption         =   "&Plant Configuration"
      Begin VB.Menu mnuPlantItem 
         Caption         =   "Influent Weir Enabled"
         Checked         =   -1  'True
         Index           =   10
      End
      Begin VB.Menu mnuPlantItem 
         Caption         =   "Grit Chamber Enabled"
         Checked         =   -1  'True
         Index           =   20
      End
      Begin VB.Menu mnuPlantItem 
         Caption         =   "Primary Clarifier Weir Enabled"
         Checked         =   -1  'True
         Index           =   30
      End
      Begin VB.Menu mnuPlantItem 
         Caption         =   "Secondary Clarifier Weir Enabled"
         Checked         =   -1  'True
         Index           =   40
      End
   End
   Begin VB.Menu mnuView 
      Caption         =   "&View"
      Begin VB.Menu mnuViewItem 
         Caption         =   "&Physico-Chemical Properties ..."
         Index           =   10
      End
      Begin VB.Menu mnuViewItem 
         Caption         =   "-"
         Index           =   20
      End
      Begin VB.Menu mnuViewItem 
         Caption         =   "&Influent Weir ..."
         Index           =   30
      End
      Begin VB.Menu mnuViewItem 
         Caption         =   "&Aerated Grit Chamber ..."
         Index           =   40
      End
      Begin VB.Menu mnuViewItem 
         Caption         =   "Primary &Clarifier ..."
         Index           =   50
      End
      Begin VB.Menu mnuViewItem 
         Caption         =   "Primary Clarifier &Weir ..."
         Index           =   60
      End
      Begin VB.Menu mnuViewItem 
         Caption         =   "Aeration &Basin(s) ..."
         Index           =   70
      End
      Begin VB.Menu mnuViewItem 
         Caption         =   "Sec&ondary Clarifier ..."
         Index           =   80
      End
      Begin VB.Menu mnuViewItem 
         Caption         =   "Secondary Clarifier We&ir ..."
         Index           =   90
      End
      Begin VB.Menu mnuViewItem 
         Caption         =   "-"
         Index           =   100
      End
      Begin VB.Menu mnuViewItem 
         Caption         =   "Change to &English Units"
         Index           =   110
      End
      Begin VB.Menu mnuViewItem 
         Caption         =   "Change to &S.I. Units"
         Index           =   120
      End
   End
   Begin VB.Menu mnuRun 
      Caption         =   "&Run"
      Begin VB.Menu mnuRunItem 
         Caption         =   "&Model ..."
      End
   End
   Begin VB.Menu mnuHelp 
      Caption         =   "&Help"
      Begin VB.Menu mnuHelpItem 
         Caption         =   "&Online Help ..."
         Index           =   10
      End
      Begin VB.Menu mnuHelpItem 
         Caption         =   "Online Manual ..."
         Index           =   20
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
      End
      Begin VB.Menu mnuHelpItem 
         Caption         =   "View Development Notes ..."
         Index           =   90
      End
      Begin VB.Menu mnuHelpItem 
         Caption         =   "-"
         Index           =   98
      End
      Begin VB.Menu mnuHelpItem 
         Caption         =   "&About ..."
         Index           =   99
      End
   End
   Begin VB.Menu mnuMTU 
      Caption         =   "&MTU Internal"
      Begin VB.Menu mnuMTUItem 
         Caption         =   "&Keep temporary model I/O files"
         Index           =   40
      End
      Begin VB.Menu mnuMTUItem 
         Caption         =   "-"
         Index           =   197
      End
      Begin VB.Menu mnuMTUItem 
         Caption         =   "&Make menu invisible"
         Index           =   198
      End
      Begin VB.Menu mnuMTUItem 
         Caption         =   "&Read me"
         Index           =   199
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


Sub Avoid_Weird_Focus_Problem()
  'Call unitsys_control_MostRecent_Force_lostfocus
  'frmMain.SetFocus
  '
  ' NOTE: IT IS VERY IMPORTANT TO SET FOCUS HERE
  ' TO SOME NON-UNITTEXTBOX CONTROL, I.E. DON'T
  ' SET IT TO txtData(0...3), BUT cboUnits(0)
  ' IS OKAY.
  'cboUnits(0).SetFocus
  'Text1.SetFocus
End Sub


Sub Populate_frmMain_Units()
Dim Frm As Form
Set Frm = frmMain
  'Call unitsys_register(Frm, lblData(0), txtData(0), cboUnits(0), "length", _
      "m", "m", "", "", 100#, True)
  'Call unitsys_register(Frm, lblData(1), txtData(1), cboUnits(1), "length", _
      "m", "m", "", "", 100#, True)
  'Call unitsys_register(Frm, lblData(2), txtData(2), cboUnits(2), "mass", _
      "kg", "kg", "", "", 100#, True)
  'Call unitsys_register(Frm, lblData(3), txtData(3), cboUnits(3), "flow_volumetric", _
      "m³/s", "m³/s", "", "", 100#, True)
  '
  ' MAIN BLOCK OF UNITS.
  '
  Call unitsys_register(Frm, lblData(4), txtData(4), cboUnits(4), "flow_volumetric", _
      "L/d", "L/d", "", "", 100#, True)
  Call unitsys_register(Frm, lblData(5), txtData(5), cboUnits(5), "concentration", _
      "mg/L", "mg/L", "", "", 100#, True)
  '
  ' OUTPUT UNITS.
  '
  Call unitsys_register(Frm, lblOutput(1), txtOutput(1), cboUnitsOutput(1), "flow_mass", _
      "kg/d", "kg/d", "", "", 100#, True)
  Call unitsys_register(Frm, lblOutput(2), txtOutput(2), cboUnitsOutput(2), "flow_mass", _
      "kg/d", "kg/d", "", "", 100#, True)
  Call unitsys_register(Frm, lblOutput(3), txtOutput(3), cboUnitsOutput(3), "flow_mass", _
      "kg/d", "kg/d", "", "", 100#, True)
End Sub


Sub Local_DirtyStatus_Set(DirtyFlag As Boolean, NewSetting As Boolean)
  Call Global_DirtyStatus_Set(Me, DirtyFlag, NewSetting)
End Sub
Sub Local_GenericStatus_Set(NewString As String)
  Call Global_GenericStatus_Set(Me, NewString)
End Sub


Private Sub cboUnitsOutput_Click(Index As Integer)
Dim Ctl As Control
Set Ctl = cboUnitsOutput(Index)
  Call unitsys_control_cbox_click(Ctl)
End Sub
Private Sub cboUnitsOutput_KeyPress(Index As Integer, KeyAscii As Integer)
  KeyAscii = Global_TextKeyPress(KeyAscii)
End Sub


Private Sub cmdMainButton_Click(Index As Integer)
Dim OUTPUT_Raise_Dirty_Flag As Boolean
Dim NEED_REFRESH As Boolean
  NEED_REFRESH = False
  OUTPUT_Raise_Dirty_Flag = False
  Select Case Index
    Case 0: Call frmD0_Props.frmD0_Props_Edit(OUTPUT_Raise_Dirty_Flag)
    Case 1: Call frmD1_InfluentWeir.frmD1_InfluentWeir_Edit(OUTPUT_Raise_Dirty_Flag)
    Case 2: Call frmD2_GritChamber.frmD2_GritChamber_Edit(OUTPUT_Raise_Dirty_Flag)
    Case 3: Call frmD3_PrimaryClarifier.frmD3_PrimaryClarifier_Edit(OUTPUT_Raise_Dirty_Flag)
    Case 4: Call frmD4_PrimaryWeir.frmD4_PrimaryWeir_Edit(OUTPUT_Raise_Dirty_Flag)
    Case 5: Call frmD5_AerationBasin.frmD5_AerationBasin_Edit(OUTPUT_Raise_Dirty_Flag)
    Case 6: Call frmD6_SecondaryClarifier.frmD6_SecondaryClarifier_Edit(OUTPUT_Raise_Dirty_Flag)
    Case 7: Call frmD7_SecondaryWeir.frmD7_SecondaryWeir_Edit(OUTPUT_Raise_Dirty_Flag)
    Case 8:
      Call ModelFAVOR_Go
      NEED_REFRESH = True
  End Select
  If (OUTPUT_Raise_Dirty_Flag = True) Then
    '
    ' THROW DIRTY FLAG AND REFRESH WINDOW.
    '
    Call Local_DirtyStatus_Set(Project_Is_Dirty, True)
    NEED_REFRESH = True
  End If
  If (NEED_REFRESH = True) Then
    Call frmMain_Refresh
  End If
End Sub


Private Sub cmdUnitChange_Click(Index As Integer)
  Select Case Index
    Case 0: Call Handle_Change_Of_Units(UnitType___ENGLISH)
    Case 1: Call Handle_Change_Of_Units(UnitType___SI)
  End Select
End Sub


Private Sub F1Book1_Click(ByVal nRow As Long, ByVal nCol As Long)
Dim Ctl As Control
Set Ctl = F1Book1
  '
  'DO _NOT_ ALLOW USER TO SCROLL OFF THE GRID.
  '
  Ctl.LeftCol = 1
  Ctl.TopRow = 1
End Sub
Private Sub F1Book1_KeyPress(KeyAscii As Integer)
'Dim Ctl As Control
'Set Ctl = F1Book1
'  '
'  'DO _NOT_ ALLOW USER TO SCROLL OFF THE GRID.
'  '
'  Ctl.LeftCol = 1
'  Ctl.TopRow = 1
End Sub
Private Sub F1Book1_KeyUp(KeyCode As Integer, Shift As Integer)
Dim Ctl As Control
Set Ctl = F1Book1
  '
  'DO _NOT_ ALLOW USER TO SCROLL OFF THE GRID.
  '
  Ctl.LeftCol = 1
  Ctl.TopRow = 1
End Sub


Private Sub Form_Load()
Dim is_internal_mtu As Boolean
Dim Ctl As Control
  '
  ' MISC INITS.
  '
  Call Local_DirtyStatus_Set(Project_Is_Dirty, False)
  Call Local_GenericStatus_Set("")
  Me.Caption = Name_App_Short
  ''''Me.Width = 9600
  Me.Width = 11250
  Me.Height = 7400
  Call CenterOnScreen(Me)
  ''''CommonDialog1.filename = MAIN_APP_PATH & "\examples\*.dat"
  CommonDialog1.FileName = _
      MAIN_APP_PATH & "\examples\*." & FileExt_App
  '
  ' STORE THE BUTTON POSITIONS FOR LATER RE-USE.
  '
Dim i As Integer
  For i = 1 To 7
    frmMain_OriginalButtonPos(i) = cmdMainButton(i).Left
  Next i
  '
  ' RESIZE A FEW CONTROLS.
  '
  Set Ctl = sspanel_ButtonHolder
  Ctl.BorderWidth = 0
  Ctl.BevelWidth = 0
  Set Ctl = sspanel_GridHolder
  Ctl.BorderWidth = 0
  Ctl.BevelWidth = 0
  'Ctl.Left = -1000
  'Ctl.Width = Me.ScaleWidth + 1000 * 2
  '
  ' CHECK FOR FILE THAT INDICATES THIS IS INTERNAL TO MTU:
  '
  is_internal_mtu = False
  If (check_internal_to_mtu()) Then is_internal_mtu = True
  mnuMTU.Visible = is_internal_mtu
  '
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
  Dim fn As String
  If (file_query_unload() = False) Then
    Cancel = True
  End If
  Call ModelFAVOR_RemoveLinkFiles
  
  
End Sub
Private Sub Form_Unload(Cancel As Integer)
  Call frmMain_Close_All_Windows
  Call unitsys_unregister_all_on_form(Me)
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
      If (Current_Filename = "") Then
        Call Avoid_Weird_Focus_Problem
        Call File_SaveAs("")
      Else
        Call Avoid_Weird_Focus_Problem
        Call File_SaveAs(Current_Filename)
      End If
    Case 3:      'Save As ...
      Call Avoid_Weird_Focus_Problem
      Call File_SaveAs("")
    'Case 6:       'Select Printer ...
     ' CommonDialog1.ShowPrinter
    Case 85:      'Print ...
      If Calculated_OK = False Then
        Show_Message ("File has changed, please calculate before printing")
      Else
        frmPrint_DO_INPUTS = True
        frmPrint_DO_OUTPUTS = False
        frmPrint_DO_PLOTS = False
        frmPrint.Show 1
      End If
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
      SendKeys "{F1}"
      ''NOTE: We currently do NOT have the resources to
      ''create an online help file for AdDesignS (1/7/98)
      ''therefore no online help is available.
      'Call Show_Message("Online help is currently unavailable.  " & _
      '    "Please refer to the printed manual or the Acrobat-format FAVOR.PDF file.")
      'Exit Sub
      ''Call LaunchFile_General("", MAIN_APP_PATH & "\help\favor.hlp")
    Case 20:      'ONLINE MANUAL.
      fn_This = MAIN_APP_PATH & "\help\favor.doc"
      If (FileExists(fn_This) = False) Then
        Call Show_Message("The file `" & fn_This & "` is missing.")
        Exit Sub
      End If
      Call ShellExecute_LocalFile(fn_This)
    Case 80:
      fn_This = MAIN_APP_PATH & "\dbase\readme.txt"
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
    Case 90:    'DEVELOPMENT NOTES.
      frmDevelopNotes.Show 1
    Case 99:    'ABOUT.
      frmAbout.Show 1
  End Select
End Sub
Private Sub mnuMTUItem_Click(Index As Integer)
  Select Case Index
    Case 40:    'KEEP TEMPORARY MODEL FILES.
      mnuMTUItem(40).Checked = Not mnuMTUItem(40).Checked
    Case 198:   'MAKE INVISIBLE.
      mnuMTU.Visible = False
    Case 199:   'READ ME.
      Call Show_Message("This menu should only appear on internal " & _
          "testing machines at MTU.  To remove the `MTU Internal` " & _
          "menu, select `Make menu invisible`.  This will make " & _
          "the menu invisible until the program is closed and reloaded.")
  End Select
End Sub


Private Sub cboUnits_Click(Index As Integer)
Dim Ctl As Control
Set Ctl = cboUnits(Index)
  Call unitsys_control_cbox_click(Ctl)
End Sub
Private Sub cboUnits_KeyPress(Index As Integer, KeyAscii As Integer)
  KeyAscii = Global_TextKeyPress(KeyAscii)
End Sub


Private Sub mnuPlantItem_Click(Index As Integer)
Dim OUTPUT_Raise_Dirty_Flag As Boolean
  OUTPUT_Raise_Dirty_Flag = False
  With NowProj.Plant
    Select Case Index
      Case 10:      'INFLUENT WEIR.
        .en_InfluentWeir = Not .en_InfluentWeir
        OUTPUT_Raise_Dirty_Flag = True
      Case 20:      'AERATED GRIT CHAMBER.
        .en_GritChamber = Not .en_GritChamber
        OUTPUT_Raise_Dirty_Flag = True
      Case 30:      'PRIMARY CLARIFIER WEIR.
        .en_PrimaryWeir = Not .en_PrimaryWeir
        OUTPUT_Raise_Dirty_Flag = True
      Case 40:      'SECONDARY CLARIFIER WEIR.
        .en_SecondaryWeir = Not .en_SecondaryWeir
        OUTPUT_Raise_Dirty_Flag = True
    End Select
  End With
  If (OUTPUT_Raise_Dirty_Flag = True) Then
    '
    ' THROW DIRTY FLAG AND REFRESH WINDOW.
    '
    Call Local_DirtyStatus_Set(Project_Is_Dirty, True)
    Call frmMain_Refresh
  End If
End Sub
Private Sub mnuRunItem_Click()
  Call cmdMainButton_Click(8)
End Sub
Private Sub mnuViewItem_Click(Index As Integer)
Dim idx_Button As Integer
  idx_Button = -1
  Select Case Index
    Case 10: idx_Button = 0
    Case 30: idx_Button = 1
    Case 40: idx_Button = 2
    Case 50: idx_Button = 3
    Case 60: idx_Button = 4
    Case 70: idx_Button = 5
    Case 80: idx_Button = 6
    Case 90: idx_Button = 7
    Case 110: Call cmdUnitChange_Click(0)     'ENGLISH UNITS.
    Case 120: Call cmdUnitChange_Click(1)     'S.I. UNITS.
  End Select
  If (idx_Button <> -1) Then
    Call cmdMainButton_Click(idx_Button)
  End If
End Sub


Private Sub txtData_GotFocus(Index As Integer)
Dim Ctl As Control
Set Ctl = txtData(Index)
Dim StatusMessagePanel As String
  Call unitsys_control_txtx_gotfocus(Ctl)
  Select Case Index
'    Case 0
'      StatusMessagePanel = "Type in the bed diameter"
'    Case 1
'      StatusMessagePanel = "Type in the bed length"
'    Case 2
'      StatusMessagePanel = "Type in the mass of adsorbent in the bed"
'    Case 3
'      StatusMessagePanel = "Type in the inlet flowrate"
  End Select
  Call Local_GenericStatus_Set(StatusMessagePanel)
End Sub
Private Sub txtData_KeyPress(Index As Integer, KeyAscii As Integer)
  KeyAscii = Global_NumericKeyPress(KeyAscii)
End Sub
Private Sub txtData_LostFocus(Index As Integer)
Dim NewValue_Okay As Integer
Dim NewValue As Double
Dim Ctl As Control
Set Ctl = txtData(Index)
Dim Val_Low As Double
Dim Val_High As Double
Dim Raise_Dirty_Flag As Boolean
Dim Too_Small As Integer
  'NOTE: LOW AND HIGH VALUES IN BASE UNITS
  Select Case Index
    Case 4: Val_Low = 1E-20: Val_High = 1E+20
    Case 5: Val_Low = 1E-20: Val_High = 1E+20
  End Select
  'If (Index = 4) Then
  '  Val_Low = 1E-20 * 60#
  '  Val_High = 1E+20 * 60#
  'Else
  '  Val_Low = 1E-20
  '  Val_High = 1E+20
  'End If
  NewValue_Okay = False
  If (unitsys_control_txtx_lostfocus_validate(Ctl, Val_Low, Val_High, NewValue, Raise_Dirty_Flag)) Then
    NewValue_Okay = True
  End If
  Call unitsys_control_txtx_lostfocus(Ctl, NewValue)
  Call Local_GenericStatus_Set("")
  If (NewValue_Okay) Then
    If (Raise_Dirty_Flag) Then
      'STORE TO MEMORY.
      Select Case Index
        'Case 0:     'BED LENGTH.
        '  NowProj.length = NewValue
        'Case 1:     'BED DIAMETER.
        '  NowProj.Diameter = NewValue
        'Case 2:     'BED MASS.
        '  NowProj.Mass = NewValue
        'Case 3:     'BED FLOW RATE.
        '  NowProj.FlowRate = NewValue
        Case 4: NowProj.Plant.Flow = NewValue
        Case 5: NowProj.Plant.SolidsConc = NewValue
      End Select
      If (Raise_Dirty_Flag) Then
        'THROW DIRTY FLAG.
        Call Local_DirtyStatus_Set(Project_Is_Dirty, True)
      End If
      'REFRESH WINDOW.
      Call frmMain_Refresh
    End If
  End If
End Sub



