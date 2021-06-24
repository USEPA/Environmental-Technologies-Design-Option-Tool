VERSION 5.00
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "Threed32.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{0842D103-1E19-101B-9AAF-1A1626551E7C}#1.0#0"; "graph32.ocx"
Begin VB.Form frmModelPSDMResults 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Results for the Pore and Surface Diffusion Model (PSDM)"
   ClientHeight    =   7305
   ClientLeft      =   3270
   ClientTop       =   3870
   ClientWidth     =   9480
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7305
   ScaleWidth      =   9480
   ShowInTaskbar   =   0   'False
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
      Left            =   7920
      TabIndex        =   50
      ToolTipText     =   "Click here to print current screen to selected printer"
      Top             =   6720
      Width           =   1455
   End
   Begin VB.PictureBox Picture1 
      Height          =   855
      Left            =   9240
      ScaleHeight     =   795
      ScaleWidth      =   1275
      TabIndex        =   49
      Top             =   6360
      Visible         =   0   'False
      Width           =   1335
   End
   Begin Threed.SSFrame ssframe_SSConc 
      Height          =   555
      Left            =   60
      TabIndex        =   42
      Top             =   2040
      Width           =   7755
      _Version        =   65536
      _ExtentX        =   13679
      _ExtentY        =   979
      _StockProps     =   14
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Begin VB.Label lblSSValueUnits 
         Caption         =   "{u}g/L"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   5400
         TabIndex        =   45
         Top             =   210
         Width           =   1065
      End
      Begin VB.Label lblSSValue 
         Alignment       =   2  'Center
         Caption         =   "lblSSValue"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   4140
         TabIndex        =   44
         Top             =   210
         Width           =   1065
      End
      Begin VB.Label lblMisc 
         Alignment       =   1  'Right Justify
         Caption         =   "Steady State Conc. at Saturation (Cr,ss)"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   0
         Left            =   120
         TabIndex        =   43
         Top             =   210
         Width           =   3825
      End
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
      Left            =   7920
      Style           =   2  'Dropdown List
      TabIndex        =   33
      Top             =   2760
      Width           =   1455
   End
   Begin Threed.SSFrame SSFrame1 
      Height          =   1995
      Left            =   60
      TabIndex        =   0
      Top             =   30
      Width           =   7755
      _Version        =   65536
      _ExtentX        =   13679
      _ExtentY        =   3519
      _StockProps     =   14
      Caption         =   "Results for:"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Begin VB.ComboBox cboCompo 
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
         Left            =   90
         Style           =   2  'Dropdown List
         TabIndex        =   3
         Top             =   300
         Width           =   3975
      End
      Begin Threed.SSCommand cmdTreat 
         Height          =   315
         Left            =   90
         TabIndex        =   2
         Top             =   1620
         Width           =   2655
         _Version        =   65536
         _ExtentX        =   2646
         _ExtentY        =   1323
         _StockProps     =   78
         Caption         =   "Treatment Objective"
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
         TabIndex        =   29
         Top             =   660
         Width           =   2655
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
         Left            =   2730
         TabIndex        =   28
         Top             =   660
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
         Left            =   3930
         TabIndex        =   27
         Top             =   660
         Width           =   1215
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
         Left            =   5130
         TabIndex        =   26
         Top             =   660
         Width           =   1215
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
         TabIndex        =   25
         Top             =   900
         Width           =   2655
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
         TabIndex        =   24
         Top             =   1140
         Width           =   2655
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
         Top             =   1380
         Width           =   2655
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
         Left            =   2730
         TabIndex        =   22
         Top             =   900
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
         Left            =   2730
         TabIndex        =   21
         Top             =   1140
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
         Left            =   2730
         TabIndex        =   20
         Top             =   1380
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
         Left            =   3930
         TabIndex        =   19
         Top             =   900
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
         Left            =   3930
         TabIndex        =   18
         Top             =   1140
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
         Left            =   3930
         TabIndex        =   17
         Top             =   1380
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
         Left            =   5130
         TabIndex        =   16
         Top             =   900
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
         Left            =   5130
         TabIndex        =   15
         Top             =   1140
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
         Left            =   5130
         TabIndex        =   14
         Top             =   1380
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
         Left            =   6330
         TabIndex        =   13
         Top             =   900
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
         Left            =   6330
         TabIndex        =   12
         Top             =   1140
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
         Left            =   6330
         TabIndex        =   11
         Top             =   1380
         Width           =   1215
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
         Left            =   6330
         TabIndex        =   10
         Top             =   660
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
         Index           =   12
         Left            =   2730
         TabIndex        =   9
         Top             =   1620
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
         Index           =   13
         Left            =   3930
         TabIndex        =   8
         Top             =   1620
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
         Index           =   14
         Left            =   5130
         TabIndex        =   7
         Top             =   1620
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
         Left            =   6330
         TabIndex        =   6
         Top             =   1620
         Width           =   1215
      End
      Begin VB.Label lblMTZ 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "lblMTZ"
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
         Left            =   6330
         TabIndex        =   5
         Top             =   285
         Width           =   1215
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Length of the MTZ (cm):"
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
         Left            =   4050
         TabIndex        =   4
         Top             =   300
         Width           =   2235
      End
   End
   Begin Threed.SSFrame SSFrame2 
      Height          =   645
      Left            =   90
      TabIndex        =   1
      Top             =   6600
      Width           =   4395
      _Version        =   65536
      _ExtentX        =   7752
      _ExtentY        =   1138
      _StockProps     =   14
      Caption         =   "X Axis Type:"
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
         Index           =   0
         Left            =   120
         TabIndex        =   39
         Top             =   270
         Width           =   795
         _Version        =   65536
         _ExtentX        =   1402
         _ExtentY        =   450
         _StockProps     =   78
         Caption         =   "&Time"
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
      Begin Threed.SSOption optType 
         Height          =   255
         Index           =   1
         Left            =   1020
         TabIndex        =   40
         Top             =   270
         Width           =   705
         _Version        =   65536
         _ExtentX        =   1244
         _ExtentY        =   450
         _StockProps     =   78
         Caption         =   "&BVT"
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
      Begin Threed.SSOption optType 
         Height          =   255
         Index           =   2
         Left            =   1860
         TabIndex        =   41
         Top             =   270
         Width           =   2415
         _Version        =   65536
         _ExtentX        =   2646
         _ExtentY        =   1323
         _StockProps     =   78
         Caption         =   "Volume Treated by M&ass"
         ForeColor       =   -2147483640
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   24.27
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
   End
   Begin Threed.SSCommand cmdFile 
      Height          =   435
      Left            =   7920
      TabIndex        =   30
      Top             =   5670
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
      Left            =   7920
      TabIndex        =   31
      Top             =   3990
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
   Begin Threed.SSCommand cmdSave 
      Height          =   435
      Left            =   7920
      TabIndex        =   32
      Top             =   4410
      Width           =   1455
      _Version        =   65536
      _ExtentX        =   2566
      _ExtentY        =   767
      _StockProps     =   78
      Caption         =   "S&ave Curves"
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
      Left            =   8310
      Top             =   930
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      FontSize        =   0
      MaxFileSize     =   256
   End
   Begin Threed.SSCommand cmdSelect 
      Height          =   435
      Left            =   7920
      TabIndex        =   34
      Top             =   4830
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
      Left            =   7920
      TabIndex        =   36
      Top             =   5250
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
      Left            =   7920
      TabIndex        =   37
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
   Begin GraphLib.Graph grpBreak 
      Height          =   3945
      Left            =   90
      TabIndex        =   38
      Top             =   2640
      Width           =   7755
      _Version        =   65536
      _ExtentX        =   13679
      _ExtentY        =   6959
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
   Begin Threed.SSFrame SSFrame3 
      Height          =   645
      Left            =   4470
      TabIndex        =   46
      Top             =   6600
      Width           =   3375
      _Version        =   65536
      _ExtentX        =   5953
      _ExtentY        =   1138
      _StockProps     =   14
      Caption         =   "Y Axis Type:"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Begin VB.ComboBox cboYAxisType 
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
         Left            =   150
         Style           =   2  'Dropdown List
         TabIndex        =   47
         Top             =   240
         Width           =   3075
      End
   End
   Begin Threed.SSCommand cmdViewProcessDiagram 
      Height          =   435
      Left            =   7920
      TabIndex        =   48
      Top             =   6090
      Visible         =   0   'False
      Width           =   1455
      _Version        =   65536
      _ExtentX        =   2646
      _ExtentY        =   1323
      _StockProps     =   78
      Caption         =   "&View Diagram"
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
      Enabled         =   0   'False
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
      Left            =   7920
      TabIndex        =   35
      Top             =   2430
      Width           =   1335
   End
End
Attribute VB_Name = "frmModelPSDMResults"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Option Base 1

Dim Flag_TO(Number_Compo_Max_PFPSDM) As Integer

Dim PopulatingScrollboxes As Integer
Dim HALT_ALL_CONTROLS As Boolean




Const frmModelPSDMResults_declarations_end = True


Sub Populate_cboYAxisType()
Dim Ctl As Control
  Set Ctl = cboYAxisType
  HALT_ALL_CONTROLS = True
  Ctl.Clear
  If (Results.is_psdm_in_room_model) Then
    If (Results.AnyCrCloseToZero = False) Then
      Ctl.AddItem "Cr/Cr,ss"
      Ctl.ItemData(Ctl.NewIndex) = CBOYAXISTYPE_C_CO
    End If
  Else
    Ctl.AddItem "C/Co"
    Ctl.ItemData(Ctl.NewIndex) = CBOYAXISTYPE_C_CO
  End If
  Ctl.AddItem "µg/L"
  Ctl.ItemData(Ctl.NewIndex) = CBOYAXISTYPE_UG_L
  Ctl.AddItem "mg/L"
  Ctl.ItemData(Ctl.NewIndex) = CBOYAXISTYPE_MG_L
  Ctl.AddItem "g/L"
  Ctl.ItemData(Ctl.NewIndex) = CBOYAXISTYPE_G_L
  Ctl.AddItem "ppb"
  Ctl.ItemData(Ctl.NewIndex) = CBOYAXISTYPE_PPB
  Ctl.AddItem "ppm"
  Ctl.ItemData(Ctl.NewIndex) = CBOYAXISTYPE_PPM
  Ctl.ListIndex = 0
  HALT_ALL_CONTROLS = False
End Sub


Private Sub cboCompo_Click()
Dim f As Double
  If (PopulatingScrollboxes) Then Exit Sub
  If (Results.is_psdm_in_room_model) Then
    lblSSValue.Caption = _
        NumberToMFBString(Results.psdmroom_Crss( _
        cboCompo.ListIndex + 1))
  End If
  If (Results.ThroughPut_50(cboCompo.ListIndex + 1).C <> -1#) And (Results.ThroughPut_50(cboCompo.ListIndex + 1).T <> -1#) And (Results.ThroughPut_05(cboCompo.ListIndex + 1).T <> -1#) And (Results.ThroughPut_05(cboCompo.ListIndex + 1).C <> -1#) And (Results.ThroughPut_95(cboCompo.ListIndex + 1).T <> -1#) And (Results.ThroughPut_95(cboCompo.ListIndex + 1).C <> -1#) Then
    f = 100 * Results.Bed.length / Results.ThroughPut_50(cboCompo.ListIndex + 1).T  'in cm/days
    lblMTZ = Format_It(f * (Results.ThroughPut_95(cboCompo.ListIndex + 1).T - Results.ThroughPut_05(cboCompo.ListIndex + 1).T), 3)
  Else
    lblMTZ = "N/A"
  End If
  If (Results.ThroughPut_05(cboCompo.ListIndex + 1).T <> -1#) And (Results.ThroughPut_05(cboCompo.ListIndex + 1).C <> -1#) Then
    lblData(0) = Format_It(Results.ThroughPut_05(cboCompo.ListIndex + 1).T / 24# / 60#, 2)
    lblData(9) = Format_It(Results.ThroughPut_05(cboCompo.ListIndex + 1).C, 2)
    lblData(3) = Format_It(Results.ThroughPut_05(cboCompo.ListIndex + 1).T * 60# * Results.Bed.Flowrate / Results.Bed.length / PI / (Results.Bed.Diameter / 2) ^ 2, 2)
    lblData(6) = Format_It(Results.ThroughPut_05(cboCompo.ListIndex + 1).T * 60# * Results.Bed.Flowrate / Results.Bed.Weight, 2)
  Else
    lblData(0) = "N/A"
    lblData(9) = "N/A"
    'lblData(12) = "N/A"
    lblData(3) = "N/A"
    lblData(6) = "N/A"
  End If
  If (Results.ThroughPut_50(cboCompo.ListIndex + 1).T <> -1#) And (Results.ThroughPut_50(cboCompo.ListIndex + 1).C <> -1#) Then
    lblData(1) = Format_It(Results.ThroughPut_50(cboCompo.ListIndex + 1).T / 24# / 60#, 2)
    lblData(10) = Format_It(Results.ThroughPut_50(cboCompo.ListIndex + 1).C, 2)
    lblData(4) = Format_It(Results.ThroughPut_50(cboCompo.ListIndex + 1).T * 60# * Results.Bed.Flowrate / Results.Bed.length / PI / (Results.Bed.Diameter / 2) ^ 2, 2)
    lblData(7) = Format_It(Results.ThroughPut_50(cboCompo.ListIndex + 1).T * 60# * Results.Bed.Flowrate / Results.Bed.Weight, 2)
  Else
    lblData(1) = "N/A"
    lblData(10) = "N/A"
    lblData(4) = "N/A"
    lblData(7) = "N/A"
  End If
  If (Results.ThroughPut_95(cboCompo.ListIndex + 1).T <> -1#) And (Results.ThroughPut_95(cboCompo.ListIndex + 1).C <> -1#) Then
    lblData(2) = Format_It(Results.ThroughPut_95(cboCompo.ListIndex + 1).T / 24# / 60#, 2)
    lblData(11) = Format_It(Results.ThroughPut_95(cboCompo.ListIndex + 1).C, 2)
    'lblData(14) = Format_It(Results.ThroughPut_95(cboCompo.ListIndex + 1).Q)
    'lblData(14) = "N/A"
    lblData(5) = Format_It(Results.ThroughPut_95(cboCompo.ListIndex + 1).T * 60# * Results.Bed.Flowrate / Results.Bed.length / PI / (Results.Bed.Diameter / 2) ^ 2, 2)
    lblData(8) = Format_It(Results.ThroughPut_95(cboCompo.ListIndex + 1).T * 60# * Results.Bed.Flowrate / Results.Bed.Weight, 2)
  Else
    lblData(2) = "N/A"
    lblData(11) = "N/A"
    'lblData(14) = "N/A"
    lblData(5) = "N/A"
    lblData(8) = "N/A"
  End If
  'cmdTreat.Caption = Format_It(Treament_Objective(cboCompo.ListIndex + 1).C, 2) & " mg/L"
  If Flag_TO(cboCompo.ListIndex + 1) Then
    lblData(12) = Format_It(Treatment_Objective(cboCompo.ListIndex + 1).T / 60# / 24#, 2)
    lblData(13) = Format_It(Treatment_Objective(cboCompo.ListIndex + 1).T * 60# * Results.Bed.Flowrate / Results.Bed.length / PI / (Results.Bed.Diameter / 2) ^ 2, 2)
    lblData(14) = Format_It(Treatment_Objective(cboCompo.ListIndex + 1).T * 60# * Results.Bed.Flowrate / Results.Bed.Weight, 2)
    lblData(15) = Format_It(Treatment_Objective(cboCompo.ListIndex + 1).C, 2)
  Else
    lblData(12) = "N/A"
    lblData(13) = "N/A"
    lblData(14) = "N/A"
    lblData(15) = "N/A"
  End If
End Sub


Private Sub cboGrid_Click()
  If (Not PopulatingScrollboxes) Then
    grpBreak.GridStyle = cboGrid.ListIndex
    grpBreak.DrawMode = 2
  End If
End Sub


Private Sub cboYAxisType_Click()
  If (HALT_ALL_CONTROLS = True) Then Exit Sub
  Call Draw_PFPSDM
End Sub

Private Sub cmdExcel_Click()
  PFPSDM_Excel = True
  CPHSDM_Excel = False
  frmExcelCurves.Show 1
End Sub


Private Sub cmdExit_Click()
  Unload Me
End Sub


Private Sub cmdFile_Click()
Dim f As Integer, Error_Code As Integer, temp As String
Dim i As Integer, J As Integer, k As Integer
Dim Eq1 As String, Filename_PFPSDM  As String

On Error GoTo File_Error
    CMDialog1.CancelError = True
    CMDialog1.Filename = ""
    CMDialog1.DialogTitle = "Print to File"
    CMDialog1.Filter = "All Files (*.*)|*.*|Text Files (*.txt)|*.txt|Data Files (*.dat)|*.dat"
    CMDialog1.FilterIndex = 2
    CMDialog1.flags = _
        cdlOFNOverwritePrompt + _
        cdlOFNPathMustExist
    CMDialog1.Action = 2

   'f = FileNameIsValid(Filename_PFPSDM, CMDialog1)
   'If Not (f) Then Exit Sub
   Filename_PFPSDM = CMDialog1.Filename

      f = FreeFile
      Open Filename_PFPSDM For Output As f
      Print #f, "Input data for the Plug-Flow Pore And Surface Diffusion Model"
    '-- Print Filename

      Print #f,
      Print #f, "From Data File :", Filename

      Print #f,
      Print #f, "Component"; Tab(30); "K*"; Tab(38); "1/n"; Tab(47); "C0"; Tab(57); "MW"; Tab(65); "Vm"; Tab(75); "NBP"
      Print #f, Tab(39); "-"; Tab(46); "mg/L"; Tab(56); "g/mol"; Tab(65); "cm" & Chr$(179) & "/mol"; Tab(76); "C"
                                                                     
      For i = 1 To Number_Component_PFPSDM
       Print #f, Trim$(Mid$(LTrim$(Results.Component(i).Name), 1, 25)); Tab(29); Format$(Results.Component(i).Use_K, "###,##0.000"); Tab(37); Format$(Results.Component(i).Use_OneOverN, "0.000"); Tab(46); Format_It(Results.Component(i).InitialConcentration, 2); Tab(55); Format_It(Results.Component(i).MW, 2); Tab(64); Format_It(Results.Component(i).MolarVolume, 2); Tab(73); Format_It(Results.Component(i).BP, 2)
      Next i
      Print #f,
      Print #f, "* K in (mg/g)*(L/mg)^(1/n) - Vm = Molar Volume at NBP"
      Print #f,

    '-----------------------Bed Data ----------------------
      Print #f, "Bed Data:"

      Print #f, Tab(5); "Bed Length: "; Tab(28); Format$(Results.Bed.length, "0.000E+00") & " m"
      Print #f, Tab(5); "Bed Diameter: "; Tab(28); Format$(Results.Bed.Diameter, "0.000E+00") & " m"
      Print #f, Tab(5); "Weight of GAC: "; Tab(28); Format$(Results.Bed.Weight, "0.000E+00") & " kg"
      Print #f, Tab(5); "Inlet Flowrate: "; Tab(28); Format$(Results.Bed.Flowrate, "0.000E+00") & " m" & Chr$(179) & "/s"
      Print #f, Tab(5); "EBCT: "; Tab(28); Format$(Results.Bed.length * PI * Results.Bed.Diameter * Results.Bed.Diameter / 4# / Results.Bed.Flowrate / 60#, "0.000E+00") & " mn"
      Print #f,
      Print #f, Tab(5); "Temperature:"; Tab(28); Format$(Results.Bed.Temperature, "0.00") & " C"
      If Results.Bed.Phase = 0 Then
        Print #f, Tab(5); "Water Density:"; Tab(28); Format$(Results.Bed.WaterDensity, "0.0000") & " g/cm" & Chr$(179)
        Print #f, Tab(5); "Water Viscosity:"; Tab(28); Format$(Results.Bed.WaterViscosity, "0.00E+00") & " g/cm.s"
      Else
        Print #f, Tab(5); "Pressure:"; Tab(28); Format$(Results.Bed.Pressure, "0.00000") & " atm"
        Print #f, Tab(5); "Air Density:"; Tab(28); Format$(Results.Bed.WaterDensity, "0.0000") & " g/cm" & Chr$(179)
        Print #f, Tab(5); "Air Viscosity:"; Tab(28); Format$(Results.Bed.WaterViscosity, "0.00E+00") & " g/cm.s"
      End If
      Print #f,

    '-----------------Carbon Properties -------------------------------
      Print #f, "Carbon Properties:"
      Print #f, Tab(5); "Name: "; Tab(28); Trim$(Results.Carbon.Name)
      Print #f, Tab(5); "Apparent Density: "; Tab(28); Format$(Results.Carbon.Density, "0.000") & " g/cm" & Chr$(179)
      Print #f, Tab(5); "Particle Radius: "; Tab(28); Format$(Results.Carbon.ParticleRadius * 100#, "0.000000") & " cm"
      Print #f, Tab(5); "Porosity: "; Tab(28); Format$(Results.Carbon.Porosity, "0.000")
      Print #f, Tab(5); "Shape Factor: "; Tab(28); Format$(Results.Carbon.ShapeFactor, "0.000")
      'Print #f, Tab(5); "Tortuosity: "; Tab(28); Format$(Results.Carbon.Tortuosity, "0.000")
      Print #f,

      '---------------Kinetic Parameters -----------------------------------------
      Print #f, "Kinetic parameters:"
      Print #f,
      Print #f, "Component"; Tab(24); "kf"; Tab(31); "Ds"; Tab(40); "Dp"; Tab(49); "St"; Tab(58); "Eds"; Tab(67); "Edp"; Tab(75); "SPDFR"
      Print #f, Tab(23); "cm/s"; Tab(32); "cm" & Chr$(178) & "/s"; Tab(41); "cm" & Chr$(178) & "/s"; Tab(50); "-"; Tab(59); "-"; Tab(68); "-"; Tab(76); "-"
      For i = 1 To Number_Component_PFPSDM
       Component(0) = Results.Component(i)
       Print #f, Trim$(Mid$(LTrim$(Results.Component(i).Name), 1, 20)); Tab(22); Format_It(Results.Component(i).kf, 2); Tab(35); Format_It(Results.Component(i).Ds, 2); Tab(44); Format_It(Results.Component(i).Dp, 2); Tab(54); Format_It(ST(0), 2); Tab(61); Format_It(Eds(0), 2); Tab(68); Format_It(Edp(0), 2); Tab(75); Format_It(Results.Component(i).SPDFR, 2)
      Next i
      Print #f,

      'Fouling-----------------------------------------
      Print #f, "Fouling correlations:"
      Print #f,
    
      Print #f, " Water type : "; Trim$(Results.Bed.Water_Correlation.Name)
      Eq1 = Format$(Results.Bed.Water_Correlation.Coeff(1), "0.00")

      If Results.Bed.Water_Correlation.Coeff(2) > 0 Then
        Eq1 = Eq1 & " + " & Format$(Results.Bed.Water_Correlation.Coeff(2), "0.00E+00") & "* t "
      Else
        If Results.Bed.Water_Correlation.Coeff(2) < 0 Then
        Eq1 = Eq1 & " - " & Format$(Abs(Results.Bed.Water_Correlation.Coeff(2)), "0.00E+00") & "* t "
       End If
      End If
      If Results.Bed.Water_Correlation.Coeff(3) > 0 Then
       Eq1 = Eq1 & " + " & Format$(Results.Bed.Water_Correlation.Coeff(3), "0.00") & "* EXP("
      Else
       If Results.Bed.Water_Correlation.Coeff(3) < 0 Then
         Eq1 = Eq1 & " - " & Format$(Abs(Results.Bed.Water_Correlation.Coeff(3)), "0.00") & "* EXP("
        End If
      End If
      If Results.Bed.Water_Correlation.Coeff(3) <> 0 Then
        If Results.Bed.Water_Correlation.Coeff(4) > 0 Then
         Eq1 = Eq1 & Format$(Results.Bed.Water_Correlation.Coeff(4), "0.00E+00") & "* t)"
        Else
         If Results.Bed.Water_Correlation.Coeff(4) < 0 Then
          Eq1 = Eq1 & Format$(Results.Bed.Water_Correlation.Coeff(4), "0.00E+00") & "* t)"
         End If
        End If
      End If
      Print #f, "K(t)/K0 = " & Eq1
      Print #f, "(t in minutes)"
      Print #f,

      For J = 1 To Number_Component_PFPSDM
       Eq1 = ""
       If Results.Component(J).Correlation.Coeff(1) = 1# Then
        Eq1 = "(K/K0) "
       Else
        If Results.Component(J).Correlation.Coeff(1) <> 0 Then Eq1 = Format$(Results.Component(J).Correlation.Coeff(1), "0.00") & " * (K/K0) "
      End If
      If Results.Component(J).Correlation.Coeff(2) > 0 Then
        Eq1 = Eq1 & "+ " & Format$(Results.Component(J).Correlation.Coeff(2), "0.00")
       Else
        If Results.Component(J).Correlation.Coeff(2) <> 0 Then Eq1 = Eq1 & "- " & Format$(Abs(Results.Component(J).Correlation.Coeff(2)), "0.00")
       End If
       If Trim$(Eq1) = "" Then
         Eq1 = "K/K0"
       End If
       Print #f, Trim$(Results.Component(J).Name) & ":"
       Print #f, Tab(10); "Correlation type: " & Trim$(Results.Component(J).Correlation.Name)
       Print #f, Tab(10); "K/K0 = " & Eq1
       
       If (Results.Component(J).Use_Tortuosity_Correlation) Then
         If (Results.Component(J).Constant_Tortuosity) Then
           Print #f, "Correlation used when SOC competition is important:"
           Print #f, " Tortuosity = 0.782 * EBCT^0.925 "
         Else
           Print #f, "Correlation used when NOM fouling is important:"
           Print #f, " Tortuosity = 1.0 if t< 70 days"
           Print #f, " Tortuosity = 0.334 + 6.610E-06 * EBCT"
         End If
       End If
       
       Print #f,
      Next J
                            
      'If Results.Use_Tortuosity_Correlation Then
      '  If Results.Constant_Tortuosity Then
      '    Print #f, "Correlation used when SOC competition is important:"
      '    Print #f, " Tortuosity = 0.782 * EBCT^0.925 "
      '  Else
      '    Print #f, "Correlation used when NOM fouling is important:"
      '    Print #f, " Tortuosity = 1.0 if t< 70 days"
      '    Print #f, " Tortuosity = 0.334 + 6.610E-06 * EBCT"
      '  End If
      'End If
      Print #f,

    '--- Print the results from the table
      Print #f, "Results for the Plug-Flow Pore And Surface Diffusion Model"
      Print #f,
      For i = 1 To Results.NComponent
        Print #f, Results.Component(i).Name
        Print #f, Tab(30); "Time(days)"; Tab(40); "BVT"; Tab(50); "TC"; Tab(60); "C (mg/L)"
        If (Results.ThroughPut_05(i).T <> -1#) And (Results.ThroughPut_05(i).C <> -1#) Then
          Print #f, "5% of the influent conc."; Tab(30); Format_It(Results.ThroughPut_05(i).T / 24# / 60#, 2); Tab(40); Format_It(Results.ThroughPut_05(i).T * 60# * Results.Bed.Flowrate / Results.Bed.length / PI / (Results.Bed.Diameter / 2) ^ 2, 2); Tab(50); Format_It(Results.ThroughPut_05(i).T * 60# * Results.Bed.Flowrate / Results.Bed.Weight, 2); Tab(60); Format_It(Results.ThroughPut_05(i).C, 2)
        Else
          Print #f, "5% of the influent conc."; Tab(30); "N/A"; Tab(40); "N/A"; Tab(50); "N/A"; Tab(60); "N/A"
        End If

        If (Results.ThroughPut_50(i).T <> -1#) And (Results.ThroughPut_50(i).C <> -1#) Then
          Print #f, "50% of the influent conc."; Tab(30); Format_It(Results.ThroughPut_50(i).T / 24# / 60#, 2); Tab(40); Format_It(Results.ThroughPut_50(i).T * 60# * Results.Bed.Flowrate / Results.Bed.length / PI / (Results.Bed.Diameter / 2) ^ 2, 2); Tab(50); Format_It(Results.ThroughPut_50(i).T * 60# * Results.Bed.Flowrate / Results.Bed.Weight, 2); Tab(60); Format_It(Results.ThroughPut_50(i).C, 2)
        Else
          Print #f, "50% of the influent conc."; Tab(30); "N/A"; Tab(40); "N/A"; Tab(50); "N/A"; Tab(60); "N/A"
        End If

        If (Results.ThroughPut_95(i).T <> -1#) And (Results.ThroughPut_95(i).C <> -1#) Then
          Print #f, "95% of the influent conc."; Tab(30); Format_It(Results.ThroughPut_95(i).T / 24# / 60#, 2); Tab(40); Format_It(Results.ThroughPut_95(i).T * 60# * Results.Bed.Flowrate / Results.Bed.length / PI / (Results.Bed.Diameter / 2) ^ 2, 2); Tab(50); Format_It(Results.ThroughPut_95(i).T * 60# * Results.Bed.Flowrate / Results.Bed.Weight, 2); Tab(60); Format_It(Results.ThroughPut_95(i).C, 2)
        Else
          Print #f, "95% of the influent conc."; Tab(30); "N/A"; Tab(40); "N/A"; Tab(50); "N/A"; Tab(60); "N/A"
        End If
        Print #f,
        If Flag_TO(i) Then
          Print #f, "Treatment Objective: " & Format_It(Treatment_Objective(i).C, 2) & " mg/L"
          Print #f,
          Print #f, Tab(10); "Time (days):"; Tab(25); Format_It(Treatment_Objective(i).T / 60# / 24#, 2)
          Print #f, Tab(10); "BVT:"; Tab(25); Format_It(Treatment_Objective(i).T * 60# * Results.Bed.Flowrate / Results.Bed.length / PI / (Results.Bed.Diameter / 2) ^ 2, 2)
          Print #f, Tab(10); "Tr. Capacity:"; Tab(25); Format_It(Treatment_Objective(i).T * 60# * Results.Bed.Flowrate / Results.Bed.Weight, 2)
        Else
          Print #f, "The breakthrough time for the treatment objective (" & Format_It(Treatment_Objective(i).C, 2) & "mg/L) could not be calculated."
        End If
        Print #f,
      Next i
      Print #f, "TC (Treatment Capacity) is in m" & Chr$(179) & "  / kg of GAC"
      
      '--- Print PSDM inputs/calculations that were returned from the FORTRAN routine.
      Print #f,
      Print #f, "PSDM Module Input Variables"
      Print #f, "Note: * designates a variable calculated in Visual BASIC"
      Print #f,

      Print #f, "Number of radial collocation points, NC            = " & Format$(PSDM_Inputs.VARS1(1), "0")
      Print #f, "Number of axial collocation points, MC             = " & Format$(PSDM_Inputs.VARS1(2), "0")
      Print #f, "Total no. of differential equations, NEQ           = " & Format$(PSDM_Inputs.VARS1(3), "0")
      Print #f, "Radius of adsorbent particle, RAD (cm)             = " & Format$(PSDM_Inputs.VARS1(4), "0.0000E+00")
      Print #f, "Apparent particle density, RHOP (g/cm^3)           = " & Format$(PSDM_Inputs.VARS1(5), "0.0000E+00")
      Print #f, "Void fraction of carbon, EPOR (-)                  = " & Format$(PSDM_Inputs.VARS1(6), "0.0000E+00")
      Print #f, "Void fraction of bed, EBED (-)                     = " & Format$(PSDM_Inputs.VARS1(7), "0.0000E+00")
      Print #f, "*Surface loading, SF (gpm/ft^2)                    = " & Format$(PSDM_Inputs.VARS1(8), "0.0000E+00")
      Print #f, "Packed bed contact time, TAU (sec)                 = " & Format$(PSDM_Inputs.VARS1(9), "0.0000E+00")
      Print #f, "Empty bed contact time, EBCT (min)                 = " & Format$(PSDM_Inputs.VARS1(10), "0.0000E+00")
      Print #f, "*Reynolds number, RE (-)                           = " & Format$(PSDM_Inputs.VARS1(11), "0.0000E+00")
      Print #f, "*Fluid density, DW (g/cm^3)                        = " & Format$(PSDM_Inputs.VARS1(12), "0.0000E+00")
      Print #f, "*Fluid viscosity, VW (g/cm-s)                      = " & Format$(PSDM_Inputs.VARS1(13), "0.0000E+00")
      Print #f, "Error flag, NFLAG                                  = " & Format$(PSDM_Inputs.VARS1(15), "0")
      Print #f,

      For i = 1 To Results.NComponent
        Print #f, Results.Component(i).Name
        Print #f, "Molal volume at the boiling pt., VB (cm^3/gmol)    = " & Format$(PSDM_Inputs.VARS2(i, 1), "0.0000E+00")
        Print #f, "Molecular weight of compound, XWT (g/gmol)         = " & Format$(PSDM_Inputs.VARS2(i, 2), "0.0000E+00")
        Print #f, "Initial bulk liquid-phase conc., CBO (umol/L)      = " & Format$(PSDM_Inputs.VARS2(i, 3), "0.0000E+00")
        Print #f, "Freundlich iso. cap., XK (umol/g)*(L/umol)^(1/n)   = " & Format$(PSDM_Inputs.VARS2(i, 4), "0.0000E+00")
        Print #f, "Freundlich isotherm constant, XN (-)               = " & Format$(PSDM_Inputs.VARS2(i, 5), "0.0000E+00")
        Print #f, "*Liquid diffusivity, DIFL (cm^2/sec)               = " & Format$(PSDM_Inputs.VARS2(i, 6), "0.0000E+00")
        Print #f, "Film transfer coefficient, KF (cm/sec)             = " & Format$(PSDM_Inputs.VARS2(i, 7), "0.0000E+00")
        Print #f, "Surface diffusion coefficient, DS (cm^2/s)         = " & Format$(PSDM_Inputs.VARS2(i, 8), "0.0000E+00")
        Print #f, "Stanton number, ST (-)                             = " & Format$(PSDM_Inputs.VARS2(i, 9), "0.0000E+00")
        Print #f, "Solute distribution parameter, DGS (-)             = " & Format$(PSDM_Inputs.VARS2(i, 10), "0.0000E+00")
        Print #f, "Biot number, BIS (-)                               = " & Format$(PSDM_Inputs.VARS2(i, 11), "0.0000E+00")
        Print #f, "Diffusivity modulus, EDS (-)                       = " & Format$(PSDM_Inputs.VARS2(i, 12), "0.0000E+00")
        Print #f, "Pore solute dist. parameter, DGP (-)               = " & Format$(PSDM_Inputs.VARS2(i, 13), "0.0000E+00")
        Print #f, "Pore diffusion coefficient, DP (cm^2/s)            = " & Format$(PSDM_Inputs.VARS2(i, 14), "0.0000E+00")
        Print #f, "Pore Biot number, BIP (-)                          = " & Format$(PSDM_Inputs.VARS2(i, 15), "0.0000E+00")
        Print #f, "Pore diffusion modulus, EDP (-)                    = " & Format$(PSDM_Inputs.VARS2(i, 16), "0.0000E+00")
        Print #f, "Surface to pore diffusivity ratio, D (-)           = " & Format$(PSDM_Inputs.VARS2(i, 17), "0.0000E+00")
        Print #f, "*Schmidt number, SC (-)                            = " & Format$(PSDM_Inputs.VARS2(i, 18), "0.0000E+00")
        Print #f, "*SPDFR (-)                                         = " & Format$(PSDM_Inputs.VARS2(i, 19), "0.0000E+00")
        Print #f,
      Next i
      
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
 
Dim Error_Code As Integer, temp As String
Dim i As Integer, H  As Single, W As Single
Dim Eq1 As String, J As Integer, MTZ As String, f As Double

On Error GoTo Print_Error

    '---- Print the graph ------------------------
'''    For i = 1 To Number_Component
'''      grpBreak.ThisPoint = i
'''      grpBreak.PatternData = i - 1
'''    Next i
'''
'''    H = grpBreak.Height
'''    W = grpBreak.Width
'''
'''    grpBreak.Visible = False 'Hide it before printing

    '
    ' THIS CODE HAD TO BE REPLACED TODAY, 1999-MAY-11, EJOMAN.
    '
    '---- OLD CODE STARTS HERE:
    'If Printer.Width < Printer.Height Then
    '  grpBreak.Height = CSng(Printer.Height / 2#)
    '  grpBreak.Width = Printer.Width
    'Else
    '  grpBreak.Height = Printer.Height
    '  grpBreak.Width = Printer.Width
    'End If
    '---- OLD CODE ENDS.
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

 'Print other results-----------------------------------------------
  Printer.ScaleLeft = -1080  'Set a 3/4-inch margin
  Printer.ScaleTop = -1080
  Printer.CurrentX = 0
  Printer.CurrentY = 0

    Printer.FontSize = 12
    Printer.FontBold = True
    Printer.FontUnderline = True
    Printer.Print "Input data for the Plug-Flow Pore And Surface Diffusion Model"
    Printer.FontSize = 10
    Printer.FontBold = False
    Printer.FontUnderline = False
    '-- Print Filename
    Printer.Print
    Printer.Print "From Data File: "; Filename
    

    Printer.Print
    Printer.Print "Component"; Tab(30); "K*"; Tab(38); "1/n"; Tab(47); "C0"; Tab(57); "MW"; Tab(65); "Vm"; Tab(75); "NBP"
    Printer.Print Tab(39); "-"; Tab(46); "mg/L"; Tab(56); "g/mol"; Tab(65); "cm" & Chr$(179) & "/mol"; Tab(76); "C"
                                                                     
    For i = 1 To Number_Component_PFPSDM
     Printer.Print Trim$(Mid$(LTrim$(Results.Component(i).Name), 1, 25)); Tab(29); Format$(Results.Component(i).Use_K, "###,##0.000"); Tab(37); Format$(Results.Component(i).Use_OneOverN, "0.000"); Tab(46); Format_It(Results.Component(i).InitialConcentration, 2); Tab(55); Format_It(Results.Component(i).MW, 2); Tab(64); Format_It(Results.Component(i).MolarVolume, 2); Tab(73); Format_It(Results.Component(i).BP, 2)
    Next i
    Printer.Print
    Printer.Print "* K in (mg/g)*(L/mg)^(1/n) - Vm = Molar Volume at NBP"
    Printer.Print

    '-----------------------Bed Data ----------------------
    Printer.FontUnderline = True
    Printer.Print "Bed Data:"
    Printer.FontUnderline = False

    Printer.Print Tab(5); "Bed Length: "; Tab(28); Format$(Results.Bed.length, "0.000E+00") & " m"
    Printer.Print Tab(5); "Bed Diameter: "; Tab(28); Format$(Results.Bed.Diameter, "0.000E+00") & " m"
    Printer.Print Tab(5); "Weight of GAC: "; Tab(28); Format$(Results.Bed.Weight, "0.000E+00") & " kg"
    Printer.Print Tab(5); "Inlet Flowrate: "; Tab(28); Format$(Results.Bed.Flowrate, "0.000E+00") & " m" & Chr$(179) & "/s"
    Printer.Print Tab(5); "EBCT: "; Tab(28); Format$(Results.Bed.length * PI * Results.Bed.Diameter * Results.Bed.Diameter / 4# / Results.Bed.Flowrate / 60#, "0.000E+00") & " mn"
    Printer.Print
    Printer.Print Tab(5); "Temperature:"; Tab(28); Format$(Results.Bed.Temperature, "0.00") & " C"
    If Results.Bed.Phase = 0 Then
      Printer.Print Tab(5); "Water Density:"; Tab(28); Format$(Results.Bed.WaterDensity, "0.0000") & " g/cm" & Chr$(179)
      Printer.Print Tab(5); "Water Viscosity:"; Tab(28); Format$(Results.Bed.WaterViscosity, "0.00E+00") & " g/cm.s"
    Else
      Printer.Print Tab(5); "Pressure:"; Tab(28); Format$(Results.Bed.Pressure, "0.00000") & " atm"
      Printer.Print Tab(5); "Air Density:"; Tab(28); Format$(Results.Bed.WaterDensity, "0.0000") & " g/cm" & Chr$(179)
      Printer.Print Tab(5); "Air Viscosity:"; Tab(28); Format$(Results.Bed.WaterViscosity, "0.00E+00") & " g/cm.s"
    End If
    Printer.Print

    '-----------------Carbon Properties -------------------------------
    Printer.FontUnderline = True
    Printer.Print "Carbon Properties:"
    Printer.FontUnderline = False

    Printer.Print Tab(5); "Name: "; Tab(28); Trim$(Results.Carbon.Name)
    Printer.Print Tab(5); "Apparent Density: "; Tab(28); Format$(Results.Carbon.Density, "0.000") & " g/cm" & Chr$(179)
    Printer.Print Tab(5); "Particle Radius: "; Tab(28); Format$(Results.Carbon.ParticleRadius * 100#, "0.000000") & " cm"
    Printer.Print Tab(5); "Porosity: "; Tab(28); Format$(Results.Carbon.Porosity, "0.000")
    Printer.Print Tab(5); "Shape Factor: "; Tab(28); Format$(Results.Carbon.ShapeFactor, "0.000")
    
    'Printer.Print Tab(5); "Tortuosity: "; Tab(28); Format$(Results.Carbon.Tortuosity, "0.000")
    Printer.Print

    '---------------Kinetic Parameters -----------------------------------------
    Printer.FontUnderline = True
    Printer.Print "Kinetic parameters:"
    Printer.FontUnderline = False
    
    Printer.Print
    Printer.Print "Component"; Tab(15); "kf"; Tab(22); "Ds"; Tab(29); "Dp"; Tab(36); "St"; Tab(43); "Eds"; Tab(50); "Edp"; Tab(57); "SPDFR"
    Printer.Print Tab(15); "cm/s"; Tab(22); "cm" & Chr$(178) & "/s"; Tab(29); "cm" & Chr$(178) & "/s"; Tab(36); "-"; Tab(43); "-"; Tab(50); "-"; Tab(57); "-"
    For i = 1 To Number_Component_PFPSDM
     Component(0) = Results.Component(i)
     Printer.Print Trim$(Mid$(LTrim$(Results.Component(i).Name), 1, 20)); Tab(15); Format_It(Results.Component(i).kf, 2); Tab(22); Format_It(Results.Component(i).Ds, 2); Tab(29); Format_It(Results.Component(i).Dp, 2); Tab(36); Format_It(ST(0), 2); Tab(43); Format_It(Eds(0), 2); Tab(50); Format_It(Edp(0), 2); Tab(57); Format_It(Results.Component(i).SPDFR, 2)
    Next i


    Printer.Print

    'Fouling-----------------------------------------
    Printer.FontUnderline = True
    Printer.Print "Fouling correlations:"
    Printer.FontUnderline = False
    Printer.Print
    
    Printer.Print " Water type : "; Trim$(Results.Bed.Water_Correlation.Name)
    Eq1 = Format$(Results.Bed.Water_Correlation.Coeff(1), "0.00")

    If Results.Bed.Water_Correlation.Coeff(2) > 0 Then
     Eq1 = Eq1 & " + " & Format$(Results.Bed.Water_Correlation.Coeff(2), "0.00E+00") & "* t "
    Else
     If Results.Bed.Water_Correlation.Coeff(2) < 0 Then
     Eq1 = Eq1 & " - " & Format$(Abs(Results.Bed.Water_Correlation.Coeff(2)), "0.00E+00") & "* t "
     End If
    End If
    If Results.Bed.Water_Correlation.Coeff(3) > 0 Then
     Eq1 = Eq1 & " + " & Format$(Results.Bed.Water_Correlation.Coeff(3), "0.00") & "* EXP("
    Else
     If Results.Bed.Water_Correlation.Coeff(3) < 0 Then
     Eq1 = Eq1 & " - " & Format$(Abs(Results.Bed.Water_Correlation.Coeff(3)), "0.00") & "* EXP("
     End If
    End If
    If Results.Bed.Water_Correlation.Coeff(3) <> 0 Then
      If Results.Bed.Water_Correlation.Coeff(4) > 0 Then
       Eq1 = Eq1 & Format$(Results.Bed.Water_Correlation.Coeff(4), "0.00E+00") & "* t)"
      Else
       If Results.Bed.Water_Correlation.Coeff(4) < 0 Then
        Eq1 = Eq1 & Format$(Results.Bed.Water_Correlation.Coeff(4), "0.00E+00") & "* t)"
       End If
      End If
    End If
    Printer.Print "K(t)/K0 = " & Eq1
    Printer.Print "(t in minutes)"
    Printer.Print

    For J = 1 To Number_Component_PFPSDM
     Eq1 = ""
     If Results.Component(J).Correlation.Coeff(1) = 1# Then
      Eq1 = "(K/K0) "
     Else
      If Results.Component(J).Correlation.Coeff(1) <> 0 Then Eq1 = Format$(Results.Component(J).Correlation.Coeff(1), "0.00") & " * (K/K0) "
     End If
     If Results.Component(J).Correlation.Coeff(2) > 0 Then
      Eq1 = Eq1 & "+ " & Format$(Results.Component(J).Correlation.Coeff(2), "0.00")
     Else
      If Results.Component(J).Correlation.Coeff(2) <> 0 Then Eq1 = Eq1 & "- " & Format$(Abs(Results.Component(J).Correlation.Coeff(2)), "0.00")
     End If
     If Trim$(Eq1) = "" Then
       Eq1 = "K/K0"
     End If
     Printer.Print Trim$(Results.Component(J).Name) & ":"
     Printer.Print Tab(10); "Correlation type: " & Trim$(Results.Component(J).Correlation.Name)

     Printer.Print Tab(10); "K/K0 = " & Eq1
    
     If (Results.Component(J).Use_Tortuosity_Correlation) Then
       If (Results.Component(J).Constant_Tortuosity) Then
         Printer.Print "Correlation used when SOC competition is important:"
         Printer.Print " Tortuosity = 0.782 * EBCT^0.925 "
       Else
         Printer.Print "Correlation used when NOM fouling is important:"
         Printer.Print " Tortuosity = 1.0 if t< 70 days"
         Printer.Print " Tortuosity = 0.334 + 6.610E-06 * EBCT"
       End If
     End If
     Printer.Print
    
    Next J
                            
    'If Results.Use_Tortuosity_Correlation Then
    '  If Results.Constant_Tortuosity Then
    '    Printer.Print "Correlation used when SOC competition is important:"
    '    Printer.Print " Tortuosity = 0.782 * EBCT^0.925 "
    '  Else
    '    Printer.Print "Correlation used when NOM fouling is important:"
    '    Printer.Print " Tortuosity = 1.0 if t< 70 days"
    '    Printer.Print " Tortuosity = 0.334 + 6.610E-06 * EBCT"
    '  End If
    'End If
    Printer.Print

'Model Parameters
    Printer.FontSize = 10
    Printer.FontBold = False
    Printer.FontUnderline = True
    Printer.Print "Model Parameters"
    Printer.FontUnderline = False
    Printer.FontSize = 10
    Printer.Print Tab(5); "Total Run Time:"; Tab(50); Format$(TimeP.End / 24# / 60#, "0.000"); " days"
    If (TimeP.Init / 60# / 24#) > 0.001 Then
       Printer.Print Tab(5); "First Point Displayed:"; Tab(50); Format$(TimeP.Init / 24# / 60#, "0.000"); " days"
    Else
       Printer.Print Tab(5); "First Point Displayed:"; Tab(50); Format$(TimeP.Init / 24# / 60#, "0.000E+00"); " days"
    End If
    Printer.Print Tab(5); "Time Step:"; Tab(50); Format$(TimeP.Step / 24# / 60#, "0.000"); " days"
    Printer.Print Tab(5); "Number of Axial Collocation Points:"; Tab(50); Format$(MC, "0")
    Printer.Print Tab(5); "Number of Radial Collocation Points:"; Tab(50); Format$(NC, "0")
    Printer.Print Tab(5); "Number of Axial Elements:"; Tab(50); Format$(Bed.NumberOfBeds, "0")
   
    '
    ' PAGE BREAK REMOVED ON 1999-MAY-11, EJOMAN.
    'Printer.NewPage

  '--- Print the results from the table
    Printer.FontSize = 10
    Printer.FontBold = True
    Printer.FontUnderline = True
    Printer.Print "Results for the Plug-Flow Pore And Surface Diffusion Model"
    Printer.FontUnderline = False
    Printer.Print
    For i = 1 To Results.NComponent
      Printer.FontSize = 12
      Printer.FontBold = True
      Printer.Print Results.Component(i).Name
      Printer.FontSize = 10
      Printer.FontBold = False
      Printer.Print Tab(30); "Time(days)"; Tab(40); "BVT"; Tab(50); "TC"; Tab(60); "C (mg/L)"
      If (Results.ThroughPut_05(i).T <> -1#) And (Results.ThroughPut_05(i).C <> -1#) Then
        Printer.Print "5% of the influent conc."; Tab(30); Format_It(Results.ThroughPut_05(i).T / 24# / 60#, 2); Tab(40); Format_It(Results.ThroughPut_05(i).T * 60# * Results.Bed.Flowrate / Results.Bed.length / PI / (Results.Bed.Diameter / 2) ^ 2, 2); Tab(50); Format_It(Results.ThroughPut_05(i).T * 60# * Results.Bed.Flowrate / Results.Bed.Weight, 2); Tab(60); Format_It(Results.ThroughPut_05(i).C, 2)
      Else
        Printer.Print "5% of the influent conc."; Tab(30); "N/A"; Tab(40); "N/A"; Tab(50); "N/A"; Tab(60); "N/A"
      End If

      If (Results.ThroughPut_50(i).T <> -1#) And (Results.ThroughPut_50(i).C <> -1#) Then
        Printer.Print "50% of the influent conc."; Tab(30); Format_It(Results.ThroughPut_50(i).T / 24# / 60#, 2); Tab(40); Format_It(Results.ThroughPut_50(i).T * 60# * Results.Bed.Flowrate / Results.Bed.length / PI / (Results.Bed.Diameter / 2) ^ 2, 2); Tab(50); Format_It(Results.ThroughPut_50(i).T * 60# * Results.Bed.Flowrate / Results.Bed.Weight, 2); Tab(60); Format_It(Results.ThroughPut_50(i).C, 2)
      Else
        Printer.Print "50% of the influent conc."; Tab(30); "N/A"; Tab(40); "N/A"; Tab(50); "N/A"; Tab(60); "N/A"
      End If

      If (Results.ThroughPut_95(i).T <> -1#) And (Results.ThroughPut_95(i).C <> -1#) Then
        Printer.Print "95% of the influent conc."; Tab(30); Format_It(Results.ThroughPut_95(i).T / 24# / 60#, 2); Tab(40); Format_It(Results.ThroughPut_95(i).T * 60# * Results.Bed.Flowrate / Results.Bed.length / PI / (Results.Bed.Diameter / 2) ^ 2, 2); Tab(50); Format_It(Results.ThroughPut_95(i).T * 60# * Results.Bed.Flowrate / Results.Bed.Weight, 2); Tab(60); Format_It(Results.ThroughPut_95(i).C, 2)
      Else
        Printer.Print "95% of the influent conc."; Tab(30); "N/A"; Tab(40); "N/A"; Tab(50); "N/A"; Tab(60); "N/A"
      End If
      Printer.Print
      If (Results.ThroughPut_50(i).C <> -1#) And (Results.ThroughPut_50(i).T <> -1#) And (Results.ThroughPut_05(i).T <> -1#) And (Results.ThroughPut_05(i).C <> -1#) And (Results.ThroughPut_95(i).T <> -1#) And (Results.ThroughPut_95(i).C <> -1#) Then
        f = 100# * Results.Bed.length / Results.ThroughPut_50(i).T   'in cm/dayss
        MTZ = Format$(f * (Results.ThroughPut_95(i).T - Results.ThroughPut_05(i).T), "0.00E+00")
      Else
        MTZ = "N/A"
      End If
      Printer.Print
      Printer.Print "MTZ Length 5%-95% (cm) :"; MTZ
      If Flag_TO(i) Then
        Printer.FontUnderline = True
        Printer.Print "Treatment Objective: " & Format_It(Treatment_Objective(i).C, 2) & " mg/L"
        Printer.FontUnderline = False
        Printer.Print
        Printer.Print Tab(10); "Time (days):"; Tab(25); Format_It(Treatment_Objective(i).T / 60# / 24#, 2)
        Printer.Print Tab(10); "BVT:"; Tab(25); Format_It(Treatment_Objective(i).T * 60# * Results.Bed.Flowrate / Results.Bed.length / PI / (Results.Bed.Diameter / 2) ^ 2, 2)
        Printer.Print Tab(10); "Tr. Capacity:"; Tab(25); Format_It(Treatment_Objective(i).T * 60# * Results.Bed.Flowrate / Results.Bed.Weight, 2)
      Else
        Printer.Print "The breakthrough time for the treatment objective (" & Format_It(Treatment_Objective(i).C, 2) & "mg/L) could not be calculated."
      End If
      Printer.Print
    Next i
    Printer.Print "TC (Treatment Capacity) is in m" & Chr$(179) & "  / kg of GAC"

      '--- Print PSDM inputs/calculations that were returned from the FORTRAN routine.
      Printer.Print
      Printer.Print "PSDM Module Input Variables"
      Printer.Print "Note: * designates a variable calculated in Visual BASIC"
      Printer.Print

      '---- OLD CODE MODIFIED 1999-MAY-11 (EJOMAN) STARTS HERE:
      'Printer.Print "Number of radial collocation points, NC            = " & Format$(PSDM_Inputs.VARS1(1), "0")
      'Printer.Print "Number of axial collocation points, MC             = " & Format$(PSDM_Inputs.VARS1(2), "0")
      'Printer.Print "Total no. of differential equations, NEQ           = " & Format$(PSDM_Inputs.VARS1(3), "0")
      'Printer.Print "Radius of adsorbent particle, RAD (cm)             = " & Format$(PSDM_Inputs.VARS1(4), "0.0000E+00")
      'Printer.Print "Apparent particle density, RHOP (g/cm^3)           = " & Format$(PSDM_Inputs.VARS1(5), "0.0000E+00")
      'Printer.Print "Void fraction of carbon, EPOR (-)                  = " & Format$(PSDM_Inputs.VARS1(6), "0.0000E+00")
      'Printer.Print "Void fraction of bed, EBED (-)                     = " & Format$(PSDM_Inputs.VARS1(7), "0.0000E+00")
      'Printer.Print "*Surface loading, SF (gpm/ft^2)                    = " & Format$(PSDM_Inputs.VARS1(8), "0.0000E+00")
      'Printer.Print "Packed bed contact time, TAU (sec)                 = " & Format$(PSDM_Inputs.VARS1(9), "0.0000E+00")
      'Printer.Print "Empty bed contact time, EBCT (min)                 = " & Format$(PSDM_Inputs.VARS1(10), "0.0000E+00")
      'Printer.Print "*Reynolds number, RE (-)                           = " & Format$(PSDM_Inputs.VARS1(11), "0.0000E+00")
      'Printer.Print "*Fluid density, DW (g/cm^3)                        = " & Format$(PSDM_Inputs.VARS1(12), "0.0000E+00")
      'Printer.Print "*Fluid viscosity, VW (g/cm-s)                      = " & Format$(PSDM_Inputs.VARS1(13), "0.0000E+00")
      'Printer.Print "Error flag, NFLAG                                  = " & Format$(PSDM_Inputs.VARS1(15), "0")
      'Printer.Print
      'For i = 1 To Results.NComponent
      '  Printer.Print Results.Component(i).Name
      '  Printer.Print "Molal volume at the boiling pt., VB (cm^3/gmol)    = " & Format$(PSDM_Inputs.VARS2(i, 1), "0.0000E+00")
      '  Printer.Print "Molecular weight of compound, XWT (g/gmol)         = " & Format$(PSDM_Inputs.VARS2(i, 2), "0.0000E+00")
      '  Printer.Print "Initial bulk liquid-phase conc., CBO (umol/L)      = " & Format$(PSDM_Inputs.VARS2(i, 3), "0.0000E+00")
      '  Printer.Print "Freundlich iso. cap., XK (umol/g)*(L/umol)^(1/n)   = " & Format$(PSDM_Inputs.VARS2(i, 4), "0.0000E+00")
      '  Printer.Print "Freundlich isotherm constant, XN (-)               = " & Format$(PSDM_Inputs.VARS2(i, 5), "0.0000E+00")
      '  Printer.Print "*Liquid diffusivity, DIFL (cm^2/sec)               = " & Format$(PSDM_Inputs.VARS2(i, 6), "0.0000E+00")
      '  Printer.Print "Film transfer coefficient, KF (cm/sec)             = " & Format$(PSDM_Inputs.VARS2(i, 7), "0.0000E+00")
      '  Printer.Print "Surface diffusion coefficient, DS (cm^2/s)         = " & Format$(PSDM_Inputs.VARS2(i, 8), "0.0000E+00")
      '  Printer.Print "Stanton number, ST (-)                             = " & Format$(PSDM_Inputs.VARS2(i, 9), "0.0000E+00")
      '  Printer.Print "Solute distribution parameter, DGS (-)             = " & Format$(PSDM_Inputs.VARS2(i, 10), "0.0000E+00")
      '  Printer.Print "Biot number, BIS (-)                               = " & Format$(PSDM_Inputs.VARS2(i, 11), "0.0000E+00")
      '  Printer.Print "Diffusivity modulus, EDS (-)                       = " & Format$(PSDM_Inputs.VARS2(i, 12), "0.0000E+00")
      '  Printer.Print "Pore solute dist. parameter, DGP (-)               = " & Format$(PSDM_Inputs.VARS2(i, 13), "0.0000E+00")
      '  Printer.Print "Pore diffusion coefficient, DP (cm^2/s)            = " & Format$(PSDM_Inputs.VARS2(i, 14), "0.0000E+00")
      '  Printer.Print "Pore Biot number, BIP (-)                          = " & Format$(PSDM_Inputs.VARS2(i, 15), "0.0000E+00")
      '  Printer.Print "Pore diffusion modulus, EDP (-)                    = " & Format$(PSDM_Inputs.VARS2(i, 16), "0.0000E+00")
      '  Printer.Print "Surface to pore diffusivity ratio, D (-)           = " & Format$(PSDM_Inputs.VARS2(i, 17), "0.0000E+00")
      '  Printer.Print "*Schmidt number, SC (-)                            = " & Format$(PSDM_Inputs.VARS2(i, 18), "0.0000E+00")
      '  Printer.Print "*SPDFR (-)                                         = " & Format$(PSDM_Inputs.VARS2(i, 19), "0.0000E+00")
      '  Printer.Print
      'Next i
      '---- OLD CODE ENDS.
      '
      '---- NEW CODE MODIFIED 1999-MAY-11 (EJOMAN) STARTS HERE:
      Printer.Print "Number of radial collocation points, NC = "; Tab(70); Format$(PSDM_Inputs.VARS1(1), "0")
      Printer.Print "Number of axial collocation points, MC = "; Tab(70); Format$(PSDM_Inputs.VARS1(2), "0")
      Printer.Print "Total no. of differential equations, NEQ = "; Tab(70); Format$(PSDM_Inputs.VARS1(3), "0")
      Printer.Print "Radius of adsorbent particle, RAD (cm) = "; Tab(70); Format$(PSDM_Inputs.VARS1(4), "0.0000E+00")
      Printer.Print "Apparent particle density, RHOP (g/cm^3) = "; Tab(70); Format$(PSDM_Inputs.VARS1(5), "0.0000E+00")
      Printer.Print "Void fraction of carbon, EPOR (-) = "; Tab(70); Format$(PSDM_Inputs.VARS1(6), "0.0000E+00")
      Printer.Print "Void fraction of bed, EBED (-) = "; Tab(70); Format$(PSDM_Inputs.VARS1(7), "0.0000E+00")
      Printer.Print "*Surface loading, SF (gpm/ft^2) = "; Tab(70); Format$(PSDM_Inputs.VARS1(8), "0.0000E+00")
      Printer.Print "Packed bed contact time, TAU (sec) = "; Tab(70); Format$(PSDM_Inputs.VARS1(9), "0.0000E+00")
      Printer.Print "Empty bed contact time, EBCT (min) = "; Tab(70); Format$(PSDM_Inputs.VARS1(10), "0.0000E+00")
      Printer.Print "*Reynolds number, RE (-) = "; Tab(70); Format$(PSDM_Inputs.VARS1(11), "0.0000E+00")
      Printer.Print "*Fluid density, DW (g/cm^3) = "; Tab(70); Format$(PSDM_Inputs.VARS1(12), "0.0000E+00")
      Printer.Print "*Fluid viscosity, VW (g/cm-s) = "; Tab(70); Format$(PSDM_Inputs.VARS1(13), "0.0000E+00")
      Printer.Print "Error flag, NFLAG = "; Tab(70); Format$(PSDM_Inputs.VARS1(15), "0")
      Printer.Print
      For i = 1 To Results.NComponent
        Printer.Print Results.Component(i).Name
        Printer.Print "Molal volume at the boiling pt., VB (cm^3/gmol) = "; Tab(70); Format$(PSDM_Inputs.VARS2(i, 1), "0.0000E+00")
        Printer.Print "Molecular weight of compound, XWT (g/gmol) = "; Tab(70); Format$(PSDM_Inputs.VARS2(i, 2), "0.0000E+00")
        Printer.Print "Initial bulk liquid-phase conc., CBO (umol/L) = "; Tab(70); Format$(PSDM_Inputs.VARS2(i, 3), "0.0000E+00")
        Printer.Print "Freundlich iso. cap., XK (umol/g)*(L/umol)^(1/n) = "; Tab(70); Format$(PSDM_Inputs.VARS2(i, 4), "0.0000E+00")
        Printer.Print "Freundlich isotherm constant, XN (-) = "; Tab(70); Format$(PSDM_Inputs.VARS2(i, 5), "0.0000E+00")
        Printer.Print "*Liquid diffusivity, DIFL (cm^2/sec) = "; Tab(70); Format$(PSDM_Inputs.VARS2(i, 6), "0.0000E+00")
        Printer.Print "Film transfer coefficient, KF (cm/sec) = "; Tab(70); Format$(PSDM_Inputs.VARS2(i, 7), "0.0000E+00")
        Printer.Print "Surface diffusion coefficient, DS (cm^2/s) = "; Tab(70); Format$(PSDM_Inputs.VARS2(i, 8), "0.0000E+00")
        Printer.Print "Stanton number, ST (-) = "; Tab(70); Format$(PSDM_Inputs.VARS2(i, 9), "0.0000E+00")
        Printer.Print "Solute distribution parameter, DGS (-) = "; Tab(70); Format$(PSDM_Inputs.VARS2(i, 10), "0.0000E+00")
        Printer.Print "Biot number, BIS (-) = "; Tab(70); Format$(PSDM_Inputs.VARS2(i, 11), "0.0000E+00")
        Printer.Print "Diffusivity modulus, EDS (-) = "; Tab(70); Format$(PSDM_Inputs.VARS2(i, 12), "0.0000E+00")
        Printer.Print "Pore solute dist. parameter, DGP (-) = "; Tab(70); Format$(PSDM_Inputs.VARS2(i, 13), "0.0000E+00")
        Printer.Print "Pore diffusion coefficient, DP (cm^2/s) = "; Tab(70); Format$(PSDM_Inputs.VARS2(i, 14), "0.0000E+00")
        Printer.Print "Pore Biot number, BIP (-) = "; Tab(70); Format$(PSDM_Inputs.VARS2(i, 15), "0.0000E+00")
        Printer.Print "Pore diffusion modulus, EDP (-) = "; Tab(70); Format$(PSDM_Inputs.VARS2(i, 16), "0.0000E+00")
        Printer.Print "Surface to pore diffusivity ratio, D (-) = "; Tab(70); Format$(PSDM_Inputs.VARS2(i, 17), "0.0000E+00")
        Printer.Print "*Schmidt number, SC (-) = "; Tab(70); Format$(PSDM_Inputs.VARS2(i, 18), "0.0000E+00")
        Printer.Print "*SPDFR (-) = "; Tab(70); Format$(PSDM_Inputs.VARS2(i, 19), "0.0000E+00")
        Printer.Print
      Next i
      '---- NEW CODE ENDS.
    
    Printer.EndDoc
    Exit Sub

Print_Error:
  Call Show_Trapped_Error("cmdPrint_Click")
  Resume Exit_Print
Exit_Print:

End Sub


Private Sub cmdSave_Click()
Dim f As Integer, i As Integer, J As Integer, temp As String
Dim Filename_PFS As String
 
On Error GoTo Save_Results_PF_Error

  CMDialog1.CancelError = True
  CMDialog1.Filename = ""
  CMDialog1.Filter = "All Files (*.*)|*.*|Text Files (*.txt)|*.txt|Data Files (*.dat)|*.dat"
  CMDialog1.FilterIndex = 2
  CMDialog1.DialogTitle = "Save curves from PSDM"
  CMDialog1.flags = _
      cdlOFNOverwritePrompt + _
      cdlOFNPathMustExist
  CMDialog1.Action = 2
   
   'f = FileNameIsValid(Filename_PFS, CMDialog1)
   'If Not (f) Then Exit Sub
   Filename_PFS = CMDialog1.Filename

   'Save, T, BVF, Usage rate, C/C0
    f = FreeFile
    Open Filename_PFS For Output As f
    Write #f, "Results file for PSDM - Windows - Version " & Format$(NVersion, "0.00")
    temp = "Time(min)    BVT(-)   VTM(m^3/kg)   "
    For i = 1 To Results.NComponent
     temp = temp & Trim$(Results.Component(i).Name) & "          "
    Next i
    Write #f, temp
    Write #f,

    temp = ""
    For i = 1 To Results.npoints
     temp = Format$(Results.T(i), "0.00")
     temp = temp & "       " & Format$(Results.T(i) * 60 * Results.Bed.Flowrate / Results.Bed.length / PI / (Results.Bed.Diameter / 2) ^ 2, "0.00")
     temp = temp & "       " & Format$(Results.T(i) * 60 * Results.Bed.Flowrate / Results.Bed.Weight, "0.00")
     For J = 1 To Results.NComponent
       temp = temp & "          " & Format$(Results.CP(J, i), "0.000")
     Next J
     Print #f, temp
     temp = ""
    Next i
    Close f
    CMDialog1.Filename = ""
 Exit Sub

Save_Results_PF_Error:
  If (Err.number = cdlCancel) Then
    'DO NOTHING.
  Else
    Call Show_Trapped_Error("cmdSave_Click")
  End If
  Resume Exit_Save_Results_PF
Exit_Save_Results_PF:
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
Dim Objective As String, temp As Double, Tr_Obj As Double, J  As Integer, i As Integer

  Objective = InputBox$("Enter your treatment objective in mg/L for " & Trim$(Results.Component(cboCompo.ListIndex + 1).Name) & ":", AppName_For_Display_Long, lblData(9))
On Error GoTo Bad_Treament_Objective
  temp = CDbl(Objective)
  i = cboCompo.ListIndex + 1
  Tr_Obj = temp / Results.Component(cboCompo.ListIndex + 1).InitialConcentration
  For J = 1 To Number_Points_Max
    If J > 2 Then
      If (Results.CP(i, J) >= Tr_Obj) And (Results.CP(i, J - 1) < Tr_Obj) Then
        Treatment_Objective(cboCompo.ListIndex + 1).T = (Results.T(J) - Results.T(J - 1)) / (Results.CP(i, J) - Results.CP(i, J - 1)) * (Tr_Obj - Results.CP(i, J - 1)) + Results.T(J - 1)
          Treatment_Objective(cboCompo.ListIndex + 1).C = ((Results.CP(i, J) - Results.CP(i, J - 1)) / (Results.T(J) - Results.T(J - 1)) * (Treatment_Objective(cboCompo.ListIndex + 1).T - Results.T(J - 1)) + Results.CP(i, J - 1)) * Results.Component(i).InitialConcentration
        GoTo Exit_Loop
      End If
    End If
  Next J
   Flag_TO(cboCompo.ListIndex + 1) = False
   lblData(12) = "N/A"
   lblData(13) = "N/A"
   lblData(14) = "N/A"
   lblData(15) = "N/A"
  Exit Sub
Exit_Loop:
  Flag_TO(cboCompo.ListIndex + 1) = True
  lblData(12) = Format_It(Treatment_Objective(cboCompo.ListIndex + 1).T / 60# / 24#, 2)
  lblData(13) = Format_It(Treatment_Objective(cboCompo.ListIndex + 1).T * 60# * Results.Bed.Flowrate / Results.Bed.length / PI / (Results.Bed.Diameter / 2) ^ 2, 2)
  lblData(14) = Format_It(Treatment_Objective(cboCompo.ListIndex + 1).T * 60# * Results.Bed.Flowrate / Results.Bed.Weight, 2)
  lblData(15) = Format_It(Treatment_Objective(cboCompo.ListIndex + 1).C, 2)
  Exit Sub

Bad_Treament_Objective:
   Resume Exit_lblLegend_Click
Exit_lblLegend_Click:


End Sub

Private Sub Draw_PFPSDM()
Dim i As Integer, J As Integer
Dim Data_Max As Double, factor As Double, Bottom_Title As String
ReDim X_Values(Number_Points_Max) As Double
Dim biggest_numpoints As Integer
Dim index_with_biggest_numpoints As Integer
Dim LastPointI As Integer
Dim SameX As Double
Dim SameY As Double

Dim most_recent_x As Double
Dim most_recent_y As Double
Dim end_the_plot As Integer

'Copy the results
  If optType(0) Then  'Time
    factor = 1# / 60# / 24#  'mn > days
    Bottom_Title = "Time(days)"
  Else
    If optType(1) Then   'BVF         mn * (mn/s) * (m3/s) / m / (m2) -> dimensionless
      factor = 60# * Results.Bed.Flowrate / (Results.Bed.length * PI * (Results.Bed.Diameter / 2#) ^ 2)
      Bottom_Title = "Bed Volumes Treated"
    Else   'Treatment Capacity
      factor = 60# * Results.Bed.Flowrate / Results.Bed.Weight   'mn * (s/mn) * (m3/s) / (kg) -> m3/kg
      'factor = 60# * Results.Bed.Flowrate / Results.Bed.Length / Pi / (Results.Bed.Diameter / 2#) ^ 2
      'factor = factor / (Bed.density * 1000)
      Bottom_Title = "m" & Chr$(179) & " treated per kg of adsorbent"
    End If
  End If
  'Results.T(I,1) time is in mn
  'Results.T(I,2) is BVF
  For i = 1 To Number_Points_Max
    X_Values(i) = Results.T(i) * factor
  Next i

   ' The following code is a rather
   ' unfortunate kludge, in my opinion.  I could find no other way to
   ' convince/force Visual Basic's graphical interface to accept two sets
   ' of data that were of two different sizes, so I determined which one
   ' was the smaller set and then filled the remainer of the smaller set
   ' with copies of the last data point in it (X,Y) (note, the default
   ' is for the data to hook back to the point (0,0) at the end of its
   ' plotting due to the fact that, by default, the (X,Y) data points
   ' that are unspecified are filled with 0's).
   ' -- If possible, it would be nice to replace this with something
   ' more elegant, but hey, it works. -- Eric J. Oman, 7/31/96
     
    'Define Graph
    grpBreak.NumSets = Results.NComponent
    grpBreak.GraphType = 6
    grpBreak.GraphStyle = 4

    ''Determine the set with the largest number of data points
    'biggest_numpoints = -1
    'index_with_biggest_numpoints = -1
    'For j = 1 To grpBreak.NumSets
    '  'If (biggest_numpoints < Results.NumPoints_Before_ThroughPut_100(j)) Then
    '  If (biggest_numpoints < Results.NPoints) Then
    '    index_with_biggest_numpoints = j
    '    'biggest_numpoints = Results.NumPoints_Before_ThroughPut_100(j)
    '    biggest_numpoints = Results.NPoints
    '  End If
    'Next j
    biggest_numpoints = Results.npoints
    
    For J = 1 To grpBreak.NumSets
     grpBreak.ThisSet = J
     'grpBreak.NumPoints = Results.NPoints
     grpBreak.NumPoints = biggest_numpoints
     grpBreak.PatternData = 1
    Next J

    grpBreak.AutoInc = 0
    
Dim dbl_CPConversionFactor As Double
Dim dblConvertedCP As Double
Dim OUT_strYAxisTitle As String
    For J = 1 To grpBreak.NumSets
      dbl_CPConversionFactor = _
          CBOYAXISTYPE_GetUnitConversion( _
          CInt(cboYAxisType.ItemData(cboYAxisType.ListIndex)), _
          Results.is_psdm_in_room_model, _
          Results.AnyCrCloseToZero, _
          J, _
          Results.Bed.Phase, _
          OUT_strYAxisTitle)
      grpBreak.ThisSet = J
      end_the_plot = False
      For i = 1 To grpBreak.NumPoints
         grpBreak.ThisPoint = i
         If (end_the_plot = False) Then
           dblConvertedCP = Results.CP(J, i) * dbl_CPConversionFactor
           If (Results.CP(J, i) < 0) Then
             If (Results.CP(J, i) = -10000#) Then
               end_the_plot = True
             Else
               grpBreak.GraphData = 0#
             End If
           Else                       'Results.CP(1,1)
             ''''grpBreak.GraphData = Results.CP(j, i)
             grpBreak.GraphData = dblConvertedCP
           End If
           If (end_the_plot = False) Then
             'grpBreak.ThisPoint = i
             'grpBreak.LabelText = ""
             'grpBreak.ThisPoint = i
             grpBreak.XPosData = X_Values(i)
             most_recent_x = X_Values(i)
             ''''most_recent_y = Results.CP(j, i)
             most_recent_y = dblConvertedCP
           End If
         End If
         If (end_the_plot = True) Then
           grpBreak.GraphData = most_recent_y
           'grpBreak.ThisPoint = i
           'grpBreak.LabelText = ""
           'grpBreak.ThisPoint = i
           grpBreak.XPosData = most_recent_x
         End If
       Next i
       grpBreak.ThisPoint = J
       grpBreak.LegendText = Trim$(Results.Component(J).Name)
    Next J

    ''Next, set values for remaining sets with # points < biggest_numpoints
    'For j = 1 To grpBreak.NumSets
    '  If (j <> index_with_biggest_numpoints) Then
    '    grpBreak.ThisSet = j
    '    LastPointI = Results.NumPoints_Before_ThroughPut_100(j)
    '    SameX = X_Values(LastPointI)
    '    SameY = Results.CP(j, LastPointI)
    '    For i = LastPointI + 1 To biggest_numpoints
    '      grpBreak.ThisPoint = i
    '      grpBreak.GraphData = SameY
    '      grpBreak.ThisPoint = i
    '      grpBreak.XPosData = SameX
    '    Next i
    '  End If
    'Next j
    
    'Other formatting
    grpBreak.PatternedLines = 0
    Data_Max = 0
    For J = 1 To grpBreak.NumSets
      grpBreak.ThisSet = J
      For i = 1 To grpBreak.NumPoints
        grpBreak.ThisPoint = i
        If grpBreak.GraphData > Data_Max Then
          Data_Max = grpBreak.GraphData
        End If
      Next i
    Next J
    grpBreak.YAxisMax = (Int(Data_Max * 10# + 1)) / 10#
    grpBreak.YAxisTicks = 4
    grpBreak.YAxisStyle = 2
    grpBreak.YAxisMin = 0#
    grpBreak.BottomTitle = Bottom_Title
    
    ''''If (Results.is_psdm_in_room_model) Then
    ''''  If (Results.AnyCrCloseToZero = True) Then
    ''''    grpBreak.LeftTitle = "Cr, " & Chr$(181) & "g/L!!"
    ''''  Else
    ''''    grpBreak.LeftTitle = "Cr/Cr,ss!!"
    ''''  End If
    ''''Else
    ''''  grpBreak.LeftTitle = "C/Co!!"
    ''''End If
    grpBreak.LeftTitle = OUT_strYAxisTitle
    
    grpBreak.DrawMode = 2

End Sub

Private Sub Command4_Click()
    Set Picture1.Picture = CaptureActiveWindow()
    PrintPictureToFitPage Printer, Picture1.Picture
    Printer.EndDoc
    ' Set focus back to form.
    Me.SetFocus
End Sub

Private Sub Form_Load()
Dim J As Integer, i As Integer
  'Set Window
  '
  ' MISC INITS.
  '
  Call Populate_cboYAxisType
'is_psdm_in_room_model As Integer
'int_Which_PSDMR_Model As Integer
'Global Const PSDMR_MODE_INROOM = 1
'Global Const PSDMR_MODE_ALONE = 2
  If (Results.is_psdm_in_room_model) Then
    Select Case Results.int_Which_PSDMR_Model
      Case PSDMR_MODE_INROOM:
        ''''Me.Caption = "Results for the PSDM in Room Model"
        Me.Caption = "Results for the PSDMR-in-Room Model (Reactions Present)"
        lblLegend(4).Caption = "5% of Cr,ss"
        lblLegend(5).Caption = "50% of Cr,ss"
        lblLegend(6).Caption = "95% of Cr,ss"
        ssframe_SSConc.Visible = True
      Case PSDMR_MODE_ALONE:
        ''''Me.Caption = "Results for the PSDM in Room Model"
        Me.Caption = "Results for the PSDMR-Alone Model (Reactions Present)"
        lblLegend(4).Caption = "5% of influent conc."
        lblLegend(5).Caption = "50% of influent conc."
        lblLegend(6).Caption = "95% of influent conc."
        ssframe_SSConc.Visible = False
    End Select
  Else
    Me.Caption = "Results for the PSDM (No Reactions Present)"
    lblLegend(4).Caption = "5% of influent conc."
    lblLegend(5).Caption = "50% of influent conc."
    lblLegend(6).Caption = "95% of influent conc."
    ssframe_SSConc.Visible = False
  End If
  lblSSValueUnits.Caption = "µg/L"
  
  Call CenterOnForm(Me, frmMain)
  
   PopulatingScrollboxes = False
   Screen.MousePointer = 11
   ''''Move frmPFPSDM.Left + (frmPFPSDM.Width / 2) - (frmbreak.Width / 2), frmPFPSDM.Top + (frmPFPSDM.Height / 2) - (frmbreak.Height / 2)
   ''''Me.HelpContextID = Hlp_Results_for
   ''''CMDialog1.CancelError = True
   lblLegend(2) = "BVT(m" & Chr$(179) & "/m" & Chr$(179) & ")"
   lblLegend(3) = "VTM(m" & Chr$(179) & "/kg)"
   Call Populate_Scrollboxes
   Call cboCompo_Click
   Call cboGrid_Click
   cboCompo.ListIndex = 0
   Call optType_Click(1, CInt(optType(1).Value))
'    optType(1) = True
   Screen.MousePointer = 0
   'grpBreak.GridStyle = 0
End Sub
Private Sub Form_Unload(Cancel As Integer)
  Call UserPrefs_Save
End Sub


Private Sub optType_Click(Index As Integer, Value As Integer)
  If (Not PopulatingScrollboxes) Then
    Call Draw_PFPSDM
  End If
End Sub


Private Sub Populate_Scrollboxes()
Dim i As Integer

    PopulatingScrollboxes = True
    
    cboGrid.AddItem "None"
    cboGrid.AddItem "Horizontal"
    cboGrid.AddItem "Vertical"
    cboGrid.AddItem "Both"
    
    For i = 1 To Results.NComponent
       cboCompo.AddItem Trim$(Results.Component(i).Name)
       Treatment_Objective(i) = Results.ThroughPut_05(i)
       If Treatment_Objective(i).C <> -1 Then
         Flag_TO(i) = True
       Else
         Flag_TO(i) = False
       End If
    Next i

    '-- Read in INI settings
    cboGrid.ListIndex = 0
    cboCompo.ListIndex = 0
    Call UserPrefs_Load
    
    PopulatingScrollboxes = False

End Sub

Private Sub UserPrefs_Load()
Dim X As Long

  On Error GoTo err_FRMBREAK_UserPrefs_Load

  X = CLng(INI_Getsetting("FRMBREAK_cboGrid"))
  If ((X >= 0) And (X <= cboGrid.ListCount - 1)) Then
    cboGrid.ListIndex = X
  End If
  X = CLng(INI_Getsetting("FRMBREAK_optType"))
  If ((X >= 0) And (X <= 2)) Then
    optType(X).Value = True
  End If
  
  Exit Sub

resume_err_FRMBREAK_UserPrefs_Load:
  Call UserPrefs_Save
  Exit Sub

err_FRMBREAK_UserPrefs_Load:
  Resume resume_err_FRMBREAK_UserPrefs_Load
           
End Sub

Private Sub UserPrefs_Save()
Dim X As Long

  X = cboGrid.ListIndex
  Call INI_PutSetting("FRMBREAK_cboGrid", Trim$(CStr(X)))
  If (optType(0)) Then X = 0
  If (optType(1)) Then X = 1
  If (optType(2)) Then X = 2
  Call INI_PutSetting("FRMBREAK_optType", Trim$(CStr(X)))

End Sub


