VERSION 5.00
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "THREED32.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{0842D103-1E19-101B-9AAF-1A1626551E7C}#1.0#0"; "GRAPH32.OCX"
Begin VB.Form frmModelCPHSDMResults 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Results for the Constant Pattern Model (CPHSDM)"
   ClientHeight    =   6795
   ClientLeft      =   2355
   ClientTop       =   870
   ClientWidth     =   9525
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6795
   ScaleWidth      =   9525
   Begin VB.PictureBox Picture1 
      Height          =   855
      Left            =   9480
      ScaleHeight     =   795
      ScaleWidth      =   1275
      TabIndex        =   48
      Top             =   5880
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.CommandButton Command4 
      Caption         =   "Print Screen"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   7800
      TabIndex        =   47
      ToolTipText     =   "Click here to print current screen to selected printer"
      Top             =   6240
      Width           =   1455
   End
   Begin GraphLib.Graph grpBreak 
      Height          =   3555
      Left            =   60
      TabIndex        =   46
      Top             =   2490
      Width           =   7515
      _Version        =   65536
      _ExtentX        =   13256
      _ExtentY        =   6271
      _StockProps     =   96
      BorderStyle     =   1
      BottomTitle     =   "Testing"
      GraphStyle      =   4
      GraphType       =   6
      GridStyle       =   3
      LeftTitle       =   "Testing"
      RandomData      =   0
      ColorData       =   0
      ExtraData       =   0
      ExtraData[]     =   0
      FontFamily      =   4
      FontSize        =   4
      FontSize[0]     =   200
      FontSize[1]     =   150
      FontSize[2]     =   100
      FontSize[3]     =   100
      FontStyle       =   4
      GraphData       =   1
      GraphData[]     =   5
      GraphData[0,0]  =   0
      GraphData[0,1]  =   0
      GraphData[0,2]  =   0
      GraphData[0,3]  =   0
      GraphData[0,4]  =   0
      LabelText       =   0
      LegendText      =   0
      PatternData     =   0
      SymbolData      =   0
      XPosData        =   0
      XPosData[]      =   0
   End
   Begin VB.ComboBox cboGrid 
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
      Left            =   7830
      Style           =   2  'Dropdown List
      TabIndex        =   41
      Top             =   2640
      Width           =   1455
   End
   Begin Threed.SSFrame SSFrame2 
      Height          =   645
      Left            =   60
      TabIndex        =   37
      Top             =   6090
      Width           =   7485
      _Version        =   65536
      _ExtentX        =   13203
      _ExtentY        =   1138
      _StockProps     =   14
      Caption         =   "C/Co as a function of:"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Begin Threed.SSOption optType 
         Height          =   255
         Index           =   2
         Left            =   3990
         TabIndex        =   2
         TabStop         =   0   'False
         Top             =   270
         Width           =   2775
         _Version        =   65536
         _ExtentX        =   2646
         _ExtentY        =   1323
         _StockProps     =   78
         Caption         =   "Volume Treated by M&ass"
         ForeColor       =   -2147483640
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   24
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin Threed.SSOption optType 
         Height          =   255
         Index           =   1
         Left            =   2550
         TabIndex        =   1
         TabStop         =   0   'False
         Top             =   270
         Width           =   975
         _Version        =   65536
         _ExtentX        =   2646
         _ExtentY        =   1323
         _StockProps     =   78
         Caption         =   "&BVT"
         ForeColor       =   -2147483640
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   24
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin Threed.SSOption optType 
         Height          =   255
         Index           =   0
         Left            =   750
         TabIndex        =   0
         TabStop         =   0   'False
         Top             =   270
         Width           =   1215
         _Version        =   65536
         _ExtentX        =   2646
         _ExtentY        =   1323
         _StockProps     =   78
         Caption         =   "&Time"
         ForeColor       =   -2147483640
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   24
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
   End
   Begin Threed.SSFrame Frame3D1 
      Height          =   2385
      Left            =   60
      TabIndex        =   3
      Top             =   30
      Width           =   7515
      _Version        =   65536
      _ExtentX        =   13256
      _ExtentY        =   4207
      _StockProps     =   14
      Caption         =   "Results for {Component Name}"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Begin Threed.SSCommand cmdTreat 
         Height          =   255
         Left            =   90
         TabIndex        =   4
         TabStop         =   0   'False
         Top             =   2010
         Width           =   2535
         _Version        =   65536
         _ExtentX        =   2646
         _ExtentY        =   1323
         _StockProps     =   78
         Caption         =   "Command3D1"
         ForeColor       =   -2147483640
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
      Begin VB.Label lblLegend 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         Caption         =   "C (mg/L)"
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
         Left            =   6210
         TabIndex        =   36
         Top             =   1050
         Width           =   1215
      End
      Begin VB.Label lblData 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "999.99"
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
         Left            =   6210
         TabIndex        =   35
         Top             =   1770
         Width           =   1215
      End
      Begin VB.Label lblData 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "999.99"
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
         Left            =   6210
         TabIndex        =   34
         Top             =   1530
         Width           =   1215
      End
      Begin VB.Label lblData 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "999.99"
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
         Left            =   6210
         TabIndex        =   33
         Top             =   1290
         Width           =   1215
      End
      Begin VB.Label lblData 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "999.99"
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
         Left            =   5010
         TabIndex        =   32
         Top             =   1770
         Width           =   1215
      End
      Begin VB.Label lblData 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "999.99"
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
         Left            =   5010
         TabIndex        =   31
         Top             =   1530
         Width           =   1215
      End
      Begin VB.Label lblData 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "999.99"
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
         Left            =   5010
         TabIndex        =   30
         Top             =   1290
         Width           =   1215
      End
      Begin VB.Label lblData 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "999.99"
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
         Left            =   3810
         TabIndex        =   29
         Top             =   1770
         Width           =   1215
      End
      Begin VB.Label lblData 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "999.99"
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
         Left            =   3810
         TabIndex        =   28
         Top             =   1530
         Width           =   1215
      End
      Begin VB.Label lblData 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "999.99"
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
         Left            =   3810
         TabIndex        =   27
         Top             =   1290
         Width           =   1215
      End
      Begin VB.Label lblData 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "999.99"
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
         Left            =   2610
         TabIndex        =   26
         Top             =   1770
         Width           =   1215
      End
      Begin VB.Label lblData 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "999.99"
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
         Left            =   2610
         TabIndex        =   25
         Top             =   1530
         Width           =   1215
      End
      Begin VB.Label lblData 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "999.99"
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
         Left            =   2610
         TabIndex        =   24
         Top             =   1290
         Width           =   1215
      End
      Begin VB.Label lblLegend 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         Caption         =   "95% of influent conc."
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
         Left            =   90
         TabIndex        =   23
         Top             =   1770
         Width           =   2535
      End
      Begin VB.Label lblLegend 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         Caption         =   "50% of influent conc."
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
         Left            =   90
         TabIndex        =   22
         Top             =   1530
         Width           =   2535
      End
      Begin VB.Label lblLegend 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         Caption         =   "5% of influent conc."
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
         Left            =   90
         TabIndex        =   21
         Top             =   1290
         Width           =   2535
      End
      Begin VB.Label lblLegend 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Tr. Capacity"
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
         Left            =   5010
         TabIndex        =   20
         Top             =   1050
         Width           =   1215
      End
      Begin VB.Label lblLegend 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         Caption         =   "BVT"
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
         Left            =   3810
         TabIndex        =   19
         Top             =   1050
         Width           =   1215
      End
      Begin VB.Label lblLegend 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Time (days)"
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
         Left            =   2610
         TabIndex        =   18
         Top             =   1050
         Width           =   1215
      End
      Begin VB.Label lblLegend 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
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
         Left            =   90
         TabIndex        =   17
         Top             =   1050
         Width           =   2535
      End
      Begin VB.Label lblPara 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Minimum Stanton Number:"
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
         Left            =   390
         TabIndex        =   16
         Top             =   285
         Width           =   2295
      End
      Begin VB.Label lblPara 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Minimum EBCT (min):"
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
         Left            =   630
         TabIndex        =   15
         Top             =   525
         Width           =   2055
      End
      Begin VB.Label lblPara 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Minimum Column Length(cm):"
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
         Left            =   90
         TabIndex        =   14
         Top             =   765
         Width           =   2595
      End
      Begin VB.Label lblPara 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "MTZ Length(cm):"
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
         Left            =   4110
         TabIndex        =   13
         Top             =   760
         Width           =   2055
      End
      Begin VB.Label lblParaValue 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Label1"
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
         Left            =   2730
         TabIndex        =   12
         Top             =   270
         Width           =   1215
      End
      Begin VB.Label lblParaValue 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Label1"
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
         Left            =   6210
         TabIndex        =   11
         Top             =   750
         Width           =   1215
      End
      Begin VB.Label lblParaValue 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Label1"
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
         Left            =   2730
         TabIndex        =   10
         Top             =   510
         Width           =   1215
      End
      Begin VB.Label lblParaValue 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Label1"
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
         Left            =   2730
         TabIndex        =   9
         Top             =   750
         Width           =   1215
      End
      Begin VB.Label lblData 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "999.99"
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
         Index           =   15
         Left            =   2610
         TabIndex        =   8
         Top             =   2010
         Width           =   1215
      End
      Begin VB.Label lblData 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "999.99"
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
         Index           =   16
         Left            =   3810
         TabIndex        =   7
         Top             =   2010
         Width           =   1215
      End
      Begin VB.Label lblData 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "999.99"
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
         Index           =   17
         Left            =   5010
         TabIndex        =   6
         Top             =   2010
         Width           =   1215
      End
      Begin VB.Label lblData 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "999.99"
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
         Index           =   18
         Left            =   6210
         TabIndex        =   5
         Top             =   2010
         Width           =   1215
      End
   End
   Begin Threed.SSCommand cmdFile 
      Height          =   435
      Left            =   7830
      TabIndex        =   38
      TabStop         =   0   'False
      Top             =   5460
      Width           =   1455
      _Version        =   65536
      _ExtentX        =   2646
      _ExtentY        =   1323
      _StockProps     =   78
      Caption         =   "Print to &File"
      ForeColor       =   -2147483640
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
   Begin Threed.SSCommand cmdExcel 
      Height          =   435
      Left            =   7830
      TabIndex        =   39
      TabStop         =   0   'False
      Top             =   3300
      Width           =   1455
      _Version        =   65536
      _ExtentX        =   2646
      _ExtentY        =   1323
      _StockProps     =   78
      Caption         =   "&Excel..."
      ForeColor       =   -2147483640
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
   Begin MSComDlg.CommonDialog CMDialog1 
      Left            =   7950
      Top             =   1560
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      FontSize        =   0
      MaxFileSize     =   256
   End
   Begin Threed.SSCommand cmdSave 
      Height          =   435
      Left            =   7830
      TabIndex        =   40
      TabStop         =   0   'False
      Top             =   3840
      Width           =   1455
      _Version        =   65536
      _ExtentX        =   2646
      _ExtentY        =   1323
      _StockProps     =   78
      Caption         =   "S&ave Curve"
      ForeColor       =   -2147483640
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
   Begin Threed.SSCommand cmdSelect 
      Height          =   435
      Left            =   7830
      TabIndex        =   42
      TabStop         =   0   'False
      Top             =   4380
      Width           =   1455
      _Version        =   65536
      _ExtentX        =   2646
      _ExtentY        =   1323
      _StockProps     =   78
      Caption         =   "&Select Printer"
      ForeColor       =   -2147483640
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
   Begin Threed.SSCommand cmdPrint 
      Height          =   435
      Left            =   7830
      TabIndex        =   43
      TabStop         =   0   'False
      Top             =   4920
      Width           =   1455
      _Version        =   65536
      _ExtentX        =   2646
      _ExtentY        =   1323
      _StockProps     =   78
      Caption         =   "&Print"
      ForeColor       =   -2147483640
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
   Begin Threed.SSCommand cmdExit 
      Height          =   435
      Left            =   7830
      TabIndex        =   45
      TabStop         =   0   'False
      Top             =   120
      Width           =   1455
      _Version        =   65536
      _ExtentX        =   2646
      _ExtentY        =   1323
      _StockProps     =   78
      Caption         =   "&Close"
      ForeColor       =   -2147483640
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
   Begin VB.Label Label1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Grid Style:"
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
      Left            =   7830
      TabIndex        =   44
      Top             =   2400
      Width           =   1575
   End
End
Attribute VB_Name = "frmModelCPHSDMResults"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Option Base 1

Dim Treatment_Objective As Throughput
Dim Flag_TO As Integer

Dim PopulatingScrollboxes As Integer



Const frmModelCPHSDMResults_declarations_end = True


Private Sub Draw_CPM()
Dim J As Integer, i As Integer, f As Integer, FileNamebis As String
Dim Data_Max As Double, factor  As Double
Dim Bottom_Title As String
ReDim X_Values(CPM_Max_Points) As Double

  Screen.MousePointer = 11
  
  If optType(0) Then  'Time
    factor = 1#
    Bottom_Title = "Time(days)"
  Else
    If optType(1) Then   'BVF
      factor = 24# * 3600# * CPM_Results.Bed.Flowrate / CPM_Results.Bed.length / PI / (CPM_Results.Bed.Diameter / 2#) ^ 2
      Bottom_Title = "Bed Volumes Treated"
    Else   'Treatment Capacity
      factor = 24# * 3600# * CPM_Results.Bed.Flowrate / CPM_Results.Bed.Weight
      Bottom_Title = "m" & Chr$(179) & " treated per kg of adsorbent"
    End If
  End If
  For i = 1 To CPM_Max_Points
    X_Values(i) = CPM_Results.T(i) * factor
  Next i
  
  'Define Graph
  grpBreak.NumSets = 1
  grpBreak.GraphType = 6      'SCATTER
  grpBreak.GraphStyle = 4
  grpBreak.NumPoints = 100
    
  
  'TEMP BEGINS.
  'grpBreak.ThisSet = 1
  'grpBreak.NumPoints = 100
  'grpBreak.AutoInc = 0
  'For i = 1 To 100
  '  grpBreak.ThisPoint = i
  '  grpBreak.GraphData = CDbl(i)
  '  grpBreak.XPosData = CDbl(i)
  'Next i
  'TEMP ENDS.
  
  
  grpBreak.AutoInc = 0
  grpBreak.GridStyle = cboGrid.ListIndex
  grpBreak.ThisSet = 1
  For i = 1 To grpBreak.NumPoints
    grpBreak.ThisPoint = i
    If CPM_Results.C_Over_C0(i) < 0 Then
      grpBreak.GraphData = 0#
    Else
      grpBreak.GraphData = CPM_Results.C_Over_C0(i)
    End If
    ''''grpBreak.ThisPoint = i
    ''''grpBreak.LabelText = ""
    ''''grpBreak.ThisPoint = i
    grpBreak.XPosData = X_Values(i)
  Next i
  grpBreak.ThisPoint = 1
  grpBreak.PatternData = 0
  grpBreak.PatternedLines = 0
  grpBreak.YAxisStyle = 2
  grpBreak.YAxisMin = 0#
  Data_Max = 0
  For i = 1 To grpBreak.NumPoints
    If CPM_Results.C_Over_C0(i) > Data_Max Then
      Data_Max = CPM_Results.C_Over_C0(i)
    End If
  Next i
  grpBreak.YAxisMax = (Int(Data_Max * 10# + 1)) / 10#
  grpBreak.YAxisTicks = 4
  grpBreak.BottomTitle = Bottom_Title
  ''''grpBreak.BottomTitle = "Testing"
  grpBreak.LeftTitle = "C/Co"
    
  Screen.MousePointer = 0
    
  grpBreak.DrawMode = 2

End Sub


Private Sub cboGrid_Click()
  If (Not PopulatingScrollboxes) Then
    grpBreak.GridStyle = cboGrid.ListIndex
    grpBreak.DrawMode = 2
  End If
End Sub


Private Sub cmdExcel_Click()
  PFPSDM_Excel = False
  CPHSDM_Excel = True
  frmExcelCurves.Show 1
End Sub


Private Sub cmdExit_Click()
   Unload Me
End Sub


Private Sub cmdFile_Click()
Dim f As Integer, Error_Code As Integer, temp As String
Dim i As Integer, J As Integer, k As Integer
Dim Eq1 As String
Dim Filename_Input As String

  On Error GoTo File_Error
  CMDialog1.Filename = ""
  CMDialog1.DialogTitle = "Print to File"
  CMDialog1.Filter = "All Files (*.*)|*.*|Text Files (*.txt)|*.txt|Data Files (*.dat)|*.dat"
  CMDialog1.FilterIndex = 2
  CMDialog1.flags = _
      cdlOFNOverwritePrompt + _
      cdlOFNPathMustExist
  CMDialog1.Action = 2
      
   'f = FileNameIsValid(Filename_Input, CMDialog1)
   'If Not (f) Then Exit Sub
  Filename_Input = CMDialog1.Filename
      
      f = FreeFile
      Open Filename_Input For Output As f

      Print #f, "Input data for the Constant Pattern Model"
    '-- Print Filename

      Print #f,
      Print #f, "From Data File :", Filename
      

      Print #f,
      Print #f, "Chemical:"; Tab(10); Trim$(CPM_Results.Component.Name)
      Print #f, Tab(5); "Molecular weight: "; Tab(28); Format(CPM_Results.Component.MW, "0.00") & " g/mol"
      Print #f, Tab(5); "Normal Boiling Point: "; Tab(28); Format(CPM_Results.Component.BP, "0.00") & " C"
      Print #f, Tab(5); "Molar Volume @ NBP: "; Tab(28); Format_It(CPM_Results.Component.MolarVolume, 2) & " cm" & Chr$(179) & "/mol"
      Print #f, Tab(5); "Initial Concentration: "; Tab(28); Format_It(CPM_Results.Component.InitialConcentration, 2) & " mg/L"
      Print #f, Tab(5); "K: "; Tab(28); Format(CPM_Results.Component.Use_K, "0.000") & " (mg/g)(L/mg)^(1/n)"
      Print #f, Tab(5); "1/n: "; Tab(28); Format(CPM_Results.Component.Use_OneOverN, "0.000")
      Print #f,

      '-----------------------Bed Data ----------------------
      Print #f, "Bed Data:"

      Print #f, Tab(5); "Bed Length: "; Tab(28); Format$(CPM_Results.Bed.length, "0.000E+00") & " m"
      Print #f, Tab(5); "Bed Diameter: "; Tab(28); Format$(CPM_Results.Bed.Diameter, "0.000E+00") & " m"
      Print #f, Tab(5); "Weight of GAC: "; Tab(28); Format$(CPM_Results.Bed.Weight, "0.000E+00") & " kg"
      Print #f, Tab(5); "Inlet Flowrate: "; Tab(28); Format$(CPM_Results.Bed.Flowrate, "0.000E+00") & " m" & Chr$(179) & "/s"
      Print #f, Tab(5); "EBCT: "; Tab(28); Format$(CPM_Results.Bed.length * PI * CPM_Results.Bed.Diameter * CPM_Results.Bed.Diameter / 4# / CPM_Results.Bed.Flowrate / 60#, "0.000E+00") & " mn"
      Print #f,
      Print #f, Tab(5); "Temperature:"; Tab(28); Format$(CPM_Results.Bed.Temperature, "0.00") & " C"
      If CPM_Results.Bed.Phase = 0 Then
        Print #f, Tab(5); "Water Density:"; Tab(28); Format$(CPM_Results.Bed.WaterDensity, "0.0000") & " g/cm" & Chr$(179)
        Print #f, Tab(5); "Water Viscosity:"; Tab(28); Format$(CPM_Results.Bed.WaterViscosity, "0.00E+00") & " g/cm.s"
      Else
        Print #f, Tab(5); "Pressure:"; Tab(28); Format$(CPM_Results.Bed.Pressure, "0.00000") & " atm"
        Print #f, Tab(5); "Air Density:"; Tab(28); Format$(CPM_Results.Bed.WaterDensity, "0.0000") & " g/cm" & Chr$(179)
        Print #f, Tab(5); "Air Viscosity:"; Tab(28); Format$(CPM_Results.Bed.WaterViscosity, "0.00E+00") & " g/cm.s"
      End If
      Print #f,

      '-----------------Carbon Properties -------------------------------
      Print #f, "Carbon Properties:"

      Print #f, Tab(5); "Name: "; Tab(28); Trim$(CPM_Results.Carbon.Name)
      Print #f, Tab(5); "Apparent Density: "; Tab(28); Format$(CPM_Results.Carbon.Density, "0.000") & " g/cm" & Chr$(179)
      Print #f, Tab(5); "Particle Radius: "; Tab(28); Format$(CPM_Results.Carbon.ParticleRadius * 100#, "0.000000") & " cm"
      Print #f, Tab(5); "Porosity: "; Tab(28); Format$(CPM_Results.Carbon.Porosity, "0.000")
      Print #f, Tab(5); "Shape Factor: "; Tab(28); Format$(CPM_Results.Carbon.ShapeFactor, "0.000")
      'Print #f, Tab(5); "Tortuosity: "; Tab(28); Format$(CPM_Results.Carbon.Tortuosity, "0.000")
      Print #f,

      '---------------Kinetic Parameters -----------------------------------------
      Print #f, "Kinetic parameters:"
      Print #f, Tab(5); "kf"; Tab(28); Format_It(CPM_Results.Component.kf, 2) & " cm/s"
      Print #f, Tab(5); "Ds"; Tab(28); Format_It(CPM_Results.Component.Ds, 2) & " cm" & Chr$(178) & "/s"
      Print #f, Tab(5); "SPDFR"; Tab(28); Format_It(CPM_Results.Component.SPDFR, 2)

      Component(0) = CPM_Results.Component
      Print #f, Tab(5); "St"; Tab(28); Format_It(ST(0), 2)
      Print #f, Tab(5); "Eds"; Tab(28); Format_It(Eds(0), 2)
    
      Print #f,

      'Fouling-----------------------------------------
      Print #f, "Fouling correlations:"
      Print #f,
      Print #f,
      Print #f, " Water type : "; Trim$(CPM_Results.Bed.Water_Correlation.Name)
      Eq1 = Format$(CPM_Results.Bed.Water_Correlation.Coeff(1), "0.00")

      If CPM_Results.Bed.Water_Correlation.Coeff(2) > 0 Then
        Eq1 = Eq1 & " + " & Format$(CPM_Results.Bed.Water_Correlation.Coeff(2), "0.00E+00") & "* t "
      Else
        If CPM_Results.Bed.Water_Correlation.Coeff(2) < 0 Then
          Eq1 = Eq1 & " - " & Format$(Abs(CPM_Results.Bed.Water_Correlation.Coeff(2)), "0.00E+00") & "* t "
        End If
      End If
      If CPM_Results.Bed.Water_Correlation.Coeff(3) > 0 Then
        Eq1 = Eq1 & " + " & Format$(CPM_Results.Bed.Water_Correlation.Coeff(3), "0.00") & "* EXP("
      Else
        If CPM_Results.Bed.Water_Correlation.Coeff(3) < 0 Then
          Eq1 = Eq1 & " - " & Format$(Abs(CPM_Results.Bed.Water_Correlation.Coeff(3)), "0.00") & "* EXP("
        End If
      End If
      If CPM_Results.Bed.Water_Correlation.Coeff(3) <> 0 Then
        If CPM_Results.Bed.Water_Correlation.Coeff(4) > 0 Then
          Eq1 = Eq1 & Format$(CPM_Results.Bed.Water_Correlation.Coeff(4), "0.00E+00") & "* t)"
        Else
          If CPM_Results.Bed.Water_Correlation.Coeff(4) < 0 Then
            Eq1 = Eq1 & Format$(CPM_Results.Bed.Water_Correlation.Coeff(4), "0.00E+00") & "* t)"
          End If
        End If
      End If
      Print #f, "K(t)/K0 = " & Eq1
      Print #f, "(t in minutes)"
      Print #f,
     
      Eq1 = ""
      If CPM_Results.Component.Correlation.Coeff(1) = 1# Then
        Eq1 = "(K/K0) "
      Else
        If CPM_Results.Component.Correlation.Coeff(1) <> 0 Then Eq1 = Format$(CPM_Results.Component.Correlation.Coeff(1), "0.00") & " * (K/K0) "
      End If
      If CPM_Results.Component.Correlation.Coeff(2) > 0 Then
        Eq1 = Eq1 & "+ " & Format$(CPM_Results.Component.Correlation.Coeff(2), "0.00")
      Else
        If CPM_Results.Component.Correlation.Coeff(2) <> 0 Then Eq1 = Eq1 & "- " & Format$(Abs(CPM_Results.Component.Correlation.Coeff(2)), "0.00")
      End If
      If Trim$(Eq1) = "" Then
        Eq1 = "K/K0"
      End If
      Print #f, Trim$(CPM_Results.Component.Name) & ":"
      Print #f, Tab(10); "Correlation type: " & Trim$(CPM_Results.Component.Correlation.Name)
      Print #f, Tab(10); "K/K0 = " & Eq1
      Print #f,
                             
      If (CPM_Results.Component.Use_Tortuosity_Correlation) Then
        If (CPM_Results.Component.Constant_Tortuosity) Then
          Print #f, "Correlation used when SOC competition is important:"
          Print #f, " Tortuosity = 0.782 * EBCT^0.925 "
        Else
          Print #f, "Correlation used when NOM fouling is important:"
          Print #f, " Tortuosity = 1.0 if t< 70 days"
          Print #f, " Tortuosity = 0.334 + 6.610E-06 * EBCT"
        End If
      End If
      Print #f,

      '--------- CPM Results ----------------------------------
      Print #f, "Constant Pattern Model Results for " & Trim$(CPM_Results.Component.Name) & ":"
      Print #f,
      Print #f, "Minimum Stanton number:"; Tab(30); Format_It(CPM_Results.Par(1), 2)
      Print #f, "Minimum EBCT:"; Tab(30); Format_It(CPM_Results.Par(2), 2) & " min"
      Print #f, "Minimum Column Length:"; Tab(30); Format_It(CPM_Results.Par(3), 2) & " cm"
      Print #f, "Throughput at 95% of the MTZ:"; Tab(30); Format_It(CPM_Results.Par(4), 2)
      Print #f, "Throughput at 5% of the MTZ:"; Tab(30); Format_It(CPM_Results.Par(5), 2)
      Print #f, "EBCT of the MTZ:"; Tab(30); Format_It(CPM_Results.Par(6), 2) & " min"
      Print #f, "Length of the MTZ:"; Tab(30); Format_It(CPM_Results.Par(7), 2) & " cm"

      Print #f,
      Print #f, Tab(30); "Time(days)"; Tab(40); "BVT"; Tab(50); "TC"; Tab(60); "C (mg/L)"
      Print #f, "5% of the influent conc."; Tab(30); Format_It(CPM_Results.ThroughPut_05.T, 2); Tab(40); Format_It(CPM_Results.ThroughPut_05.T * 24# * 3600# * CPM_Results.Bed.Flowrate / CPM_Results.Bed.length / PI / (Bed.Diameter / 2) ^ 2, 2); Tab(50); Format_It(CPM_Results.ThroughPut_05.T * 24# * 3600# * CPM_Results.Bed.Flowrate / CPM_Results.Bed.Weight, 2); Tab(60); Format_It(CPM_Results.ThroughPut_05.C, 2)
      Print #f, "50% of the influent conc."; Tab(30); Format_It(CPM_Results.ThroughPut_50.T, 2); Tab(40); Format_It(CPM_Results.ThroughPut_50.T * 24# * 3600# * CPM_Results.Bed.Flowrate / CPM_Results.Bed.length / PI / (Bed.Diameter / 2) ^ 2, 2); Tab(50); Format_It(CPM_Results.ThroughPut_50.T * 24# * 3600# * CPM_Results.Bed.Flowrate / CPM_Results.Bed.Weight, 2); Tab(60); Format_It(CPM_Results.ThroughPut_50.C, 2)
      Print #f, "95% of the influent conc."; Tab(30); Format_It(CPM_Results.ThroughPut_95.T, 2); Tab(40); Format_It(CPM_Results.ThroughPut_95.T * 24# * 3600# * CPM_Results.Bed.Flowrate / CPM_Results.Bed.length / PI / (Bed.Diameter / 2) ^ 2, 2); Tab(50); Format_It(CPM_Results.ThroughPut_95.T * 24# * 3600# * CPM_Results.Bed.Flowrate / CPM_Results.Bed.Weight, 2); Tab(60); Format_It(CPM_Results.ThroughPut_95.C, 2)
      Print #f,
      Print #f, "TC (Treatment Capacity) is in m" & Chr$(179) & "  / kg of GAC"
      Print #f,

      If Flag_TO Then
        Print #f, "Treatment Objective: " & Format_It(Treatment_Objective.C, 2) & " mg/L"
        Print #f,
        Print #f, "Time (days):"; Tab(20); Format_It(Treatment_Objective.T, 2)
        Print #f, "BVT:"; Tab(20); Format_It(Treatment_Objective.T * 24# * 3600# * CPM_Results.Bed.Flowrate / CPM_Results.Bed.length / PI / (CPM_Results.Bed.Diameter / 2) ^ 2, 2)
        Print #f, "Tr. Capacity:"; Tab(20); Format_It(Treatment_Objective.T * 24# * 3600# * CPM_Results.Bed.Flowrate / CPM_Results.Bed.Weight, 2)
      Else
        Print #f, "The breakthrough time for the treatment objective (" & Format_It(Treatment_Objective.C, 2) & "mg/L) could not be calculated."
      End If
      Close (f)
    CMDialog1.Filename = ""
    Exit Sub

File_Error:
  If (Err.number = cdlCancel) Then
    'DO NOTHING.
  Else
    Call Show_Trapped_Error("cmdFile_Click")
  End If
  Resume Exit_Print_File
Exit_Print_File:
End Sub


Private Sub cmdPrint_Click()
Dim Error_Code As Integer, temp  As String
Dim H As Single, W As Single, Eq1 As String, i As Integer

    On Error GoTo Print_Error

'---Print Graph ---------------------------------------------------
'''    H = grpBreak.Height
'''    W = grpBreak.Width
'''
'''    grpBreak.Visible = False 'Hide it before printing

    '
    ' THIS CODE HAD TO BE REPLACED TODAY, 1999-MAY-11, EJOMAN.
    '
    '---- NEW CODE STARTS HERE:
'''    If Printer.Width < Printer.Height Then
'''      grpBreak.Height = CDbl(Printer.ScaleHeight) * 0.5
'''      grpBreak.Width = CDbl(Printer.ScaleWidth) * 0.75
'''    Else
'''      grpBreak.Height = CDbl(Printer.ScaleHeight) * 0.75
'''      grpBreak.Width = CDbl(Printer.ScaleWidth) * 0.75
'''    End If
    '---- NEW CODE ENDS.

'MsgBox _
"Printer.Height = " & Trim$(Str$(Printer.Height)) & ", " & _
"Printer.Width = " & Trim$(Str$(Printer.Width)) & ", " & _
"Printer.ScaleHeight = " & Trim$(Str$(Printer.ScaleHeight)) & ", " & _
"Printer.ScaleWidth = " & Trim$(Str$(Printer.ScaleWidth)) & ", " & _
"Printer.ScaleLeft = " & Trim$(Str$(Printer.ScaleLeft)) & ", " & _
"Printer.ScaleTop = " & Trim$(Str$(Printer.ScaleTop)) & ", "

    '
    ' THE PRINTING CODE HAD TO BE REPLACED TODAY, 1999-MAY-11, EJOMAN.
    ' REFER TO www.microsoft.com KNOWLEDGE BASE ARTICLE #Q150222.
    '
    '---- OLD CODE STARTS HERE:
    'grpBreak.PrintStyle = 2
    'grpBreak.DrawMode = 5
    '---- OLD CODE ENDS.
    '
    '---- NEW CODE STARTS HERE:
'''    Printer.ScaleLeft = -((Printer.Width - grpBreak.Width) / 2)
'''    Printer.ScaleTop = -((Printer.Height - grpBreak.Height) / 2)
'''    Printer.PaintPicture _
'''        grpBreak.Picture, _
'''        0, _
'''        0, _
'''        grpBreak.Width, _
'''        grpBreak.Height
'''    Printer.Line _
'''        (0, 0)- _
'''        (grpBreak.Width, grpBreak.Height), _
'''        QBColor(0), _
'''        B
    '---- NEW CODE ENDS.

'''    grpBreak.Height = H
'''    grpBreak.Width = W
'''
'''    grpBreak.Visible = True
'''
'''    grpBreak.PrintStyle = 2
'''    grpBreak.DrawMode = 2

    '
    ' A "SKIP TO NEXT PAGE" COMMAND HAD TO BE ADDED TO THE PRINTING
    ' CODE TODAY, 1999-MAY-11, EJOMAN.
    '
    '---- NEW CODE STARTS HERE:
'''    Printer.NewPage
    '---- NEW CODE ENDS.

'---Print other results------------------------------------------
  Printer.ScaleLeft = -1080  'Set a 3/4-inch margin
  Printer.ScaleTop = -1080
  Printer.CurrentX = 0
  Printer.CurrentY = 0

    '-- Print Filename

    Printer.FontSize = 10
    Printer.Print "From Data File: "; Filename
    Printer.Print

    Printer.FontSize = 12
    Printer.FontBold = True
    Printer.FontUnderline = True
    Printer.Print "Input data for the Constant Pattern Model"
    Printer.FontSize = 10
    Printer.FontBold = False
    Printer.FontUnderline = True
    Printer.Print
    Printer.Print "Chemical:"; Tab(10); Trim$(CPM_Results.Component.Name)
    Printer.FontUnderline = False
    Printer.Print Tab(5); "Molecular weight: "; Tab(28); Format(CPM_Results.Component.MW, "0.00") & " g/mol"
    Printer.Print Tab(5); "Normal Boiling Point: "; Tab(28); Format(CPM_Results.Component.BP, "0.00") & " C"
    Printer.Print Tab(5); "Molar Volume @ NBP: "; Tab(28); Format_It(CPM_Results.Component.MolarVolume, 2) & " cm" & Chr$(179) & "/mol"
    Printer.Print Tab(5); "Initial Concentration: "; Tab(28); Format_It(CPM_Results.Component.InitialConcentration, 2) & " mg/L"
    Printer.Print Tab(5); "K: "; Tab(28); Format(CPM_Results.Component.Use_K, "0.000") & " (mg/g)(L/mg)^(1/n)"
    Printer.Print Tab(5); "1/n: "; Tab(28); Format(CPM_Results.Component.Use_OneOverN, "0.000")
    Printer.Print

    '-----------------------Bed Data ----------------------
    Printer.FontUnderline = True
    Printer.Print "Bed Data:"
    Printer.FontUnderline = False

    Printer.Print Tab(5); "Bed Length: "; Tab(28); Format$(CPM_Results.Bed.length, "0.000E+00") & " m"
    Printer.Print Tab(5); "Bed Diameter: "; Tab(28); Format$(CPM_Results.Bed.Diameter, "0.000E+00") & " m"
    Printer.Print Tab(5); "Weight of GAC: "; Tab(28); Format$(CPM_Results.Bed.Weight, "0.000E+00") & " kg"
    Printer.Print Tab(5); "Inlet Flowrate: "; Tab(28); Format$(CPM_Results.Bed.Flowrate, "0.000E+00") & " m" & Chr$(179) & "/s"
    Printer.Print Tab(5); "EBCT: "; Tab(28); Format$(CPM_Results.Bed.length * PI * CPM_Results.Bed.Diameter * CPM_Results.Bed.Diameter / 4# / CPM_Results.Bed.Flowrate / 60#, "0.000E+00") & " mn"
    Printer.Print
    Printer.Print Tab(5); "Temperature:"; Tab(28); Format$(CPM_Results.Bed.Temperature, "0.00") & " C"
    If CPM_Results.Bed.Phase = 0 Then
      Printer.Print Tab(5); "Water Density:"; Tab(28); Format$(CPM_Results.Bed.WaterDensity, "0.0000") & " g/cm" & Chr$(179)
      Printer.Print Tab(5); "Water Viscosity:"; Tab(28); Format$(CPM_Results.Bed.WaterViscosity, "0.00E+00") & " g/cm.s"
    Else
      Printer.Print Tab(5); "Pressure:"; Tab(28); Format$(CPM_Results.Bed.Pressure, "0.00000") & " atm"
      Printer.Print Tab(5); "Air Density:"; Tab(28); Format$(CPM_Results.Bed.WaterDensity, "0.0000") & " g/cm" & Chr$(179)
      Printer.Print Tab(5); "Air Viscosity:"; Tab(28); Format$(CPM_Results.Bed.WaterViscosity, "0.00E+00") & " g/cm.s"
    End If
    Printer.Print

    '-----------------Carbon Properties -------------------------------
    Printer.FontUnderline = True
    Printer.Print "Carbon Properties:"
    Printer.FontUnderline = False

    Printer.Print Tab(5); "Name: "; Tab(28); Trim$(CPM_Results.Carbon.Name)
    Printer.Print Tab(5); "Apparent Density: "; Tab(28); Format$(CPM_Results.Carbon.Density, "0.000") & " g/cm" & Chr$(179)
    Printer.Print Tab(5); "Particle Radius: "; Tab(28); Format$(CPM_Results.Carbon.ParticleRadius * 100#, "0.000000") & " cm"
    Printer.Print Tab(5); "Porosity: "; Tab(28); Format$(CPM_Results.Carbon.Porosity, "0.000")
    Printer.Print Tab(5); "Shape Factor: "; Tab(28); Format$(CPM_Results.Carbon.ShapeFactor, "0.000")
    'Printer.Print Tab(5); "Tortuosity: "; Tab(28); Format$(CPM_Results.Carbon.Tortuosity, "0.000")
    Printer.Print

    '---------------Kinetic Parameters -----------------------------------------
    Printer.FontUnderline = True
    Printer.Print "Kinetic parameters:"
    Printer.FontUnderline = False
    Printer.Print Tab(5); "kf"; Tab(28); Format_It(CPM_Results.Component.kf, 2) & " cm/s"
    Printer.Print Tab(5); "Ds"; Tab(28); Format_It(CPM_Results.Component.Ds, 2) & " cm" & Chr$(178) & "/s"
    Printer.Print Tab(5); "SPDFR"; Tab(28); Format_It(CPM_Results.Component.SPDFR, 2)

    Component(0) = CPM_Results.Component
    Printer.Print Tab(5); "St"; Tab(28); Format_It(ST(0), 2)
    Printer.Print Tab(5); "Eds"; Tab(28); Format_It(Eds(0), 2)
    
    Printer.Print

    'Fouling-----------------------------------------
    Printer.FontUnderline = True
    Printer.Print "Fouling correlations:"
    Printer.FontUnderline = False
    Printer.Print

    Printer.Print
    Printer.Print " Water type : "; Trim$(CPM_Results.Bed.Water_Correlation.Name)
    Eq1 = Format$(CPM_Results.Bed.Water_Correlation.Coeff(1), "0.00")

    If CPM_Results.Bed.Water_Correlation.Coeff(2) > 0 Then
     Eq1 = Eq1 & " + " & Format$(CPM_Results.Bed.Water_Correlation.Coeff(2), "0.00E+00") & "* t "
    Else
     If CPM_Results.Bed.Water_Correlation.Coeff(2) < 0 Then
     Eq1 = Eq1 & " - " & Format$(Abs(CPM_Results.Bed.Water_Correlation.Coeff(2)), "0.00E+00") & "* t "
     End If
    End If
    If CPM_Results.Bed.Water_Correlation.Coeff(3) > 0 Then
     Eq1 = Eq1 & " + " & Format$(CPM_Results.Bed.Water_Correlation.Coeff(3), "0.00") & "* EXP("
    Else
     If CPM_Results.Bed.Water_Correlation.Coeff(3) < 0 Then
     Eq1 = Eq1 & " - " & Format$(Abs(CPM_Results.Bed.Water_Correlation.Coeff(3)), "0.00") & "* EXP("
     End If
    End If
    If CPM_Results.Bed.Water_Correlation.Coeff(3) <> 0 Then
      If CPM_Results.Bed.Water_Correlation.Coeff(4) > 0 Then
       Eq1 = Eq1 & Format$(CPM_Results.Bed.Water_Correlation.Coeff(4), "0.00E+00") & "* t)"
      Else
       If CPM_Results.Bed.Water_Correlation.Coeff(4) < 0 Then
        Eq1 = Eq1 & Format$(CPM_Results.Bed.Water_Correlation.Coeff(4), "0.00E+00") & "* t)"
       End If
      End If
    End If
    Printer.Print "K(t)/K0 = " & Eq1
    Printer.Print "(t in minutes)"
    Printer.Print

     
    Eq1 = ""
    If CPM_Results.Component.Correlation.Coeff(1) = 1# Then
      Eq1 = "(K/K0) "
    Else
      If CPM_Results.Component.Correlation.Coeff(1) <> 0 Then Eq1 = Format$(CPM_Results.Component.Correlation.Coeff(1), "0.00") & " * (K/K0) "
    End If
    If CPM_Results.Component.Correlation.Coeff(2) > 0 Then
      Eq1 = Eq1 & "+ " & Format$(CPM_Results.Component.Correlation.Coeff(2), "0.00")
    Else
      If CPM_Results.Component.Correlation.Coeff(2) <> 0 Then Eq1 = Eq1 & "- " & Format$(Abs(CPM_Results.Component.Correlation.Coeff(2)), "0.00")
    End If
    If Trim$(Eq1) = "" Then
       Eq1 = "K/K0"
    End If
    Printer.Print Trim$(CPM_Results.Component.Name) & ":"
    Printer.Print Tab(10); "Correlation type: " & Trim$(CPM_Results.Component.Correlation.Name)

    Printer.Print Tab(10); "K/K0 = " & Eq1

    Printer.Print
                            
    If (CPM_Results.Component.Use_Tortuosity_Correlation) Then
      If (CPM_Results.Component.Constant_Tortuosity) Then
        Printer.Print "Correlation used when SOC competition is important:"
        Printer.Print " Tortuosity = 0.782 * EBCT^0.925 "
      Else
        Printer.Print "Correlation used when NOM fouling is important:"
        Printer.Print " Tortuosity = 1.0 if t< 70 days"
        Printer.Print " Tortuosity = 0.334 + 6.610E-06 * EBCT"
      End If
    End If
    Printer.Print

    '--------- CPM Results ----------------------------------
    Printer.Print
    Printer.FontSize = 12
    Printer.FontBold = True
    Printer.FontUnderline = True
    Printer.Print "Constant Pattern Model Results for " & Trim$(CPM_Results.Component.Name) & ":"
    Printer.Print
    Printer.FontSize = 10
    Printer.FontBold = False
    Printer.FontUnderline = False
    Printer.Print "Minimum Stanton number:"; Tab(30); Format_It(CPM_Results.Par(1), 2)
    Printer.Print "Minimum EBCT:"; Tab(30); Format_It(CPM_Results.Par(2), 2) & " min"
    Printer.Print "Minimum Column Length:"; Tab(30); Format_It(CPM_Results.Par(3), 2) & " cm"
    Printer.Print "MTZ Length:"; Tab(30); Format_It(CPM_Results.Par(7), 2) & " cm"

    Printer.Print
    Printer.Print Tab(30); "Time(days)"; Tab(40); "BVT"; Tab(50); "TC"; Tab(60); "C (mg/L)"
    Printer.Print "5% of the influent conc."; Tab(30); Format_It(CPM_Results.ThroughPut_05.T, 2); Tab(40); Format_It(CPM_Results.ThroughPut_05.T * 24# * 3600# * CPM_Results.Bed.Flowrate / CPM_Results.Bed.length / PI / (Bed.Diameter / 2) ^ 2, 2); Tab(50); Format_It(CPM_Results.ThroughPut_05.T * 24# * 3600# * CPM_Results.Bed.Flowrate / CPM_Results.Bed.Weight, 2); Tab(60); Format_It(CPM_Results.ThroughPut_05.C, 2)
    Printer.Print "50% of the influent conc."; Tab(30); Format_It(CPM_Results.ThroughPut_50.T, 2); Tab(40); Format_It(CPM_Results.ThroughPut_50.T * 24# * 3600# * CPM_Results.Bed.Flowrate / CPM_Results.Bed.length / PI / (Bed.Diameter / 2) ^ 2, 2); Tab(50); Format_It(CPM_Results.ThroughPut_50.T * 24# * 3600# * CPM_Results.Bed.Flowrate / CPM_Results.Bed.Weight, 2); Tab(60); Format_It(CPM_Results.ThroughPut_50.C, 2)
    Printer.Print "95% of the influent conc."; Tab(30); Format_It(CPM_Results.ThroughPut_95.T, 2); Tab(40); Format_It(CPM_Results.ThroughPut_95.T * 24# * 3600# * CPM_Results.Bed.Flowrate / CPM_Results.Bed.length / PI / (Bed.Diameter / 2) ^ 2, 2); Tab(50); Format_It(CPM_Results.ThroughPut_95.T * 24# * 3600# * CPM_Results.Bed.Flowrate / CPM_Results.Bed.Weight, 2); Tab(60); Format_It(CPM_Results.ThroughPut_95.C, 2)
    Printer.Print
    Printer.Print "TC (Treatment Capacity) is in m" & Chr$(179) & "  / kg of GAC"
    Printer.Print

    If Flag_TO Then
     Printer.Print "Treatment Objective: " & Format_It(Treatment_Objective.C, 2) & " mg/L"
     Printer.Print
     Printer.Print "Time (days):"; Tab(20); Format_It(Treatment_Objective.T, 2)
     Printer.Print "BVT:"; Tab(20); Format_It(Treatment_Objective.T * 24# * 3600# * CPM_Results.Bed.Flowrate / CPM_Results.Bed.length / PI / (CPM_Results.Bed.Diameter / 2) ^ 2, 2)
     Printer.Print "Tr. Capacity:"; Tab(20); Format_It(Treatment_Objective.T * 24# * 3600# * CPM_Results.Bed.Flowrate / CPM_Results.Bed.Weight, 2)
    Else
     Printer.Print "The breakthrough time for the treatment objective (" & Format_It(Treatment_Objective.C, 2) & "mg/L) could not be calculated."
    End If
    Printer.EndDoc
    Exit Sub

Print_Error:
  Call Show_Trapped_Error("cmdPrint_Click")
  Resume Exit_Print
Exit_Print:

End Sub


Private Sub cmdSave_Click()
Dim f As Integer, i As Integer, temp As String
Dim Filename_CPM As String
On Error GoTo Save_Results_CPM_Error
  CMDialog1.CancelError = True
  CMDialog1.Filename = ""
  CMDialog1.Filter = "All Files (*.*)|*.*|Text Files (*.txt)|*.txt|Data Files (*.dat)|*.dat"
  CMDialog1.FilterIndex = 2
  CMDialog1.DialogTitle = "Save curve from Constant Pattern Model"
  CMDialog1.flags = _
      cdlOFNOverwritePrompt + _
      cdlOFNPathMustExist
  CMDialog1.Action = 2

   'f = FileNameIsValid(Filename_CPM, CMDialog1)
   'If Not (f) Then Exit Sub
  Filename_CPM = CMDialog1.Filename

    f = FreeFile
    Open Filename_CPM For Output As f
    Write #f, "Results file for Constant Pattern Model"
    temp = "Time(days)       "
     temp = temp & "BVT" & "        " & "Usage Rate " & "     " & Trim$(CPM_Results.Component.Name)
    Print #f, temp
    Print #f, " days             -         m" & Chr$(179) & "/kg GAC  "
    Write #f,
    temp = ""
    For i = 1 To 100
     temp = Format$(CPM_Results.T(i), "0.00")
     temp = temp & "       " & Format$(CPM_Results.T(i) * 24# * 3600# * CPM_Results.Bed.Flowrate / CPM_Results.Bed.length / PI / (CPM_Results.Bed.Diameter / 2#) ^ 2, "0.00")
     temp = temp & "       " & Format$(CPM_Results.T(i) * 24# * 3600# * CPM_Results.Bed.Flowrate / CPM_Results.Bed.Weight, "0.00")
     temp = temp & "       " & Format$(CPM_Results.C_Over_C0(i), "0.000")
     Print #f, temp
     temp = ""
    Next i
    Close f
    CMDialog1.Filename = ""
    Exit Sub
Save_Results_CPM_Error:
  If (Err.number = cdlCancel) Then
    'DO NOTHING.
  Else
    Call Show_Trapped_Error("cmdSave_Click")
  End If
  Resume Exit_Save_Results_CPM
Exit_Save_Results_CPM:
End Sub


Private Sub cmdSelect_Click()
Dim Error_Code As Integer
Dim temp As String
  On Error GoTo Select_Print_Error
  'CMDialog1.flags = PD_PRINTSETUP
  'CMDialog1.Action = 5
  CMDialog1.CancelError = False
  CMDialog1.ShowPrinter
  Exit Sub
Select_Print_Error:
  Call Show_Trapped_Error("cmdSelect_Click")
  Resume Exit_Select_Print
Exit_Select_Print:
End Sub


Private Sub cmdTreat_Click()
Dim Objective As String, temp As Double, Tr_Obj As Double, J  As Integer
  Objective = InputBox$("Enter your treatment objective in mg/L:", AppName_For_Display_Long, lblData(9))
On Error GoTo Bad_Treament_Objective
  temp = CDbl(Objective)
  Tr_Obj = temp / CPM_Results.Component.InitialConcentration
  For J = 1 To CPM_Max_Points
    If J > 2 Then
      If (CPM_Results.C_Over_C0(J) >= Tr_Obj) And (CPM_Results.C_Over_C0(J - 1) < Tr_Obj) Then
        Treatment_Objective.T = (CPM_Results.T(J) - CPM_Results.T(J - 1)) / (CPM_Results.C_Over_C0(J) - CPM_Results.C_Over_C0(J - 1)) * (Tr_Obj - CPM_Results.C_Over_C0(J - 1)) + CPM_Results.T(J - 1)
        Treatment_Objective.C = ((CPM_Results.C_Over_C0(J) - CPM_Results.C_Over_C0(J - 1)) / (CPM_Results.T(J) - CPM_Results.T(J - 1)) * (Treatment_Objective.T - CPM_Results.T(J - 1)) + CPM_Results.C_Over_C0(J - 1)) * CPM_Results.Component.InitialConcentration
        GoTo Exit_Loop
      End If
    End If
  Next J
  Flag_TO = False
  lblData(15) = "N/A"
  lblData(16) = "N/A"
  lblData(17) = "N/A"
  lblData(18) = "N/A"
  Exit Sub
Exit_Loop:
  Flag_TO = True
  lblData(15) = Format_It(Treatment_Objective.T, 2)
  lblData(16) = Format_It(Treatment_Objective.T * 24# * 3600# * CPM_Results.Bed.Flowrate / CPM_Results.Bed.length / PI / (CPM_Results.Bed.Diameter / 2) ^ 2, 2)
  lblData(17) = Format_It(Treatment_Objective.T * 24# * 3600# * CPM_Results.Bed.Flowrate / CPM_Results.Bed.Weight, 2)
  lblData(18) = Format_It(Treatment_Objective.C, 2)
  Exit Sub
Bad_Treament_Objective:
   Resume Exit_lblLegend_Click
Exit_lblLegend_Click:
End Sub


Private Sub Command4_Click()
    Set Picture1.Picture = CaptureActiveWindow()
    PrintPictureToFitPage Printer, Picture1.Picture
    Printer.EndDoc
    ' Set focus back to form.
    Me.SetFocus
End Sub

Private Sub Form_Load()
Dim i As Integer
   
  PopulatingScrollboxes = False
  Screen.MousePointer = 11
  
  'Move frmPFPSDM.Left + (frmPFPSDM.Width / 2) - (frmCPM.Width / 2), frmPFPSDM.Top + (frmPFPSDM.Height / 2) - (frmCPM.Height / 2)
  Call CenterOnForm(Me, frmMain)
  
  Frame3D1.Caption = "Results for " & Trim$(CPM_Results.Component.Name) & ":"
  lblParaValue(0) = Format_It(CPM_Results.Par(1), 2) 'Minimum Stanton
  lblParaValue(2) = Format_It(CPM_Results.Par(7), 2) 'MTZ Length
  lblParaValue(5) = Format_It(CPM_Results.Par(2), 2) 'Minimum EBCT
  lblParaValue(6) = Format_It(CPM_Results.Par(3), 2) 'Minimum Column Length
  
  lblLegend(2) = "BVT(m" & Chr$(179) & "/m" & Chr$(179) & ")"
  lblLegend(3) = "VTM(m" & Chr$(179) & "/kg)"
  
  Treatment_Objective = CPM_Results.ThroughPut_05
  
  Call Populate_Scrollboxes
  Call cboGrid_Click
  Call optType_Click(1, CInt(optType(1).Value))
  
  'cboGrid.AddItem "None"
  'cboGrid.AddItem "Horizontal"
  'cboGrid.AddItem "Vertical"
  'cboGrid.AddItem "Both"
  'cboGrid.ListIndex = 0
  
  '    optType(0) = True
  
  Flag_TO = True
  lblData(0) = Format_It(CPM_Results.ThroughPut_05.T, 2)
  lblData(1) = Format_It(CPM_Results.ThroughPut_50.T, 2)
  lblData(2) = Format_It(CPM_Results.ThroughPut_95.T, 2)
  '------ C -------
  lblData(9) = Format_It(CPM_Results.ThroughPut_05.C, 2)
  lblData(10) = Format_It(CPM_Results.ThroughPut_50.C, 2)
  lblData(11) = Format_It(CPM_Results.ThroughPut_95.C, 2)
  
  '----- BVF ---------
  lblData(3) = Format_It(CPM_Results.ThroughPut_05.T * 24# * 3600# * CPM_Results.Bed.Flowrate / CPM_Results.Bed.length / PI / (CPM_Results.Bed.Diameter / 2) ^ 2, 2)
  lblData(4) = Format_It(CPM_Results.ThroughPut_50.T * 24# * 3600# * CPM_Results.Bed.Flowrate / CPM_Results.Bed.length / PI / (CPM_Results.Bed.Diameter / 2) ^ 2, 2)
  lblData(5) = Format_It(CPM_Results.ThroughPut_95.T * 24# * 3600# * CPM_Results.Bed.Flowrate / CPM_Results.Bed.length / PI / (CPM_Results.Bed.Diameter / 2) ^ 2, 2)
  
  '-----Carbon Us. rate --------- m3 of water/kg of GAC
  lblData(6) = Format_It(CPM_Results.ThroughPut_05.T * 24# * 3600# * CPM_Results.Bed.Flowrate / CPM_Results.Bed.Weight, 2)
  lblData(7) = Format_It(CPM_Results.ThroughPut_50.T * 24# * 3600# * CPM_Results.Bed.Flowrate / CPM_Results.Bed.Weight, 2)
  lblData(8) = Format_It(CPM_Results.ThroughPut_95.T * 24# * 3600# * CPM_Results.Bed.Flowrate / CPM_Results.Bed.Weight, 2)
  
  '-----Treatment Objective------
  cmdTreat.Caption = "Treat. Objective"
  lblData(15) = lblData(0)
  lblData(16) = lblData(3)
  lblData(17) = lblData(6)
  lblData(18) = lblData(9)
  
  '     grpBreak.GridStyle = 0

End Sub
Private Sub Form_Unload(Cancel As Integer)
  Call UserPrefs_Save
End Sub


Private Sub optType_Click(Index As Integer, Value As Integer)
  If (Not PopulatingScrollboxes) Then
    Call Draw_CPM
  End If
End Sub


Private Sub Populate_Scrollboxes()
Dim i As Integer
  PopulatingScrollboxes = True
  cboGrid.AddItem "None"
  cboGrid.AddItem "Horizontal"
  cboGrid.AddItem "Vertical"
  cboGrid.AddItem "Both"
  '-- Read in INI settings
  cboGrid.ListIndex = 0
  Call UserPrefs_Load
  PopulatingScrollboxes = False
End Sub


Private Sub UserPrefs_Load()
Dim X As Long
  On Error GoTo err_FRMCPM_UserPrefs_Load
  X = CLng(INI_Getsetting("FRMCPM_cboGrid"))
  If ((X >= 0) And (X <= cboGrid.ListCount - 1)) Then
    cboGrid.ListIndex = X
  End If
  X = CLng(INI_Getsetting("FRMCPM_optType"))
  If ((X >= 0) And (X <= 2)) Then
    optType(X).Value = True
  End If
  Exit Sub
resume_err_FRMCPM_UserPrefs_Load:
  Call UserPrefs_Save
  Exit Sub
err_FRMCPM_UserPrefs_Load:
  Resume resume_err_FRMCPM_UserPrefs_Load
End Sub
Private Sub UserPrefs_Save()
Dim X As Long
  X = cboGrid.ListIndex
  Call INI_PutSetting("FRMCPM_cboGrid", Trim$(CStr(X)))
  If (optType(0)) Then X = 0
  If (optType(1)) Then X = 1
  If (optType(2)) Then X = 2
  Call INI_PutSetting("FRMCPM_optType", Trim$(CStr(X)))
End Sub


