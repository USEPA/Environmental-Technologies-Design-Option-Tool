VERSION 5.00
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "Threed32.ocx"
Object = "{B16553C3-06DB-101B-85B2-0000C009BE81}#1.0#0"; "Spin32.ocx"
Begin VB.Form frmD5_AerationBasin 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Enter Parameters [Aeration Basin]"
   ClientHeight    =   6795
   ClientLeft      =   780
   ClientTop       =   1260
   ClientWidth     =   9480
   ControlBox      =   0   'False
   HelpContextID   =   6000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6795
   ScaleWidth      =   9480
   ShowInTaskbar   =   0   'False
   Begin Threed.SSFrame SSFrame3 
      Height          =   2955
      Left            =   6360
      TabIndex        =   26
      Top             =   6900
      Visible         =   0   'False
      Width           =   3135
      _Version        =   65536
      _ExtentX        =   5530
      _ExtentY        =   5212
      _StockProps     =   14
      Caption         =   "Unused -- Invisible"
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
         Index           =   9
         Left            =   120
         Style           =   2  'Dropdown List
         TabIndex        =   45
         TabStop         =   0   'False
         Top             =   1170
         Width           =   1500
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
         Index           =   8
         Left            =   1560
         Style           =   2  'Dropdown List
         TabIndex        =   40
         TabStop         =   0   'False
         Top             =   1170
         Width           =   1500
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
         Index           =   7
         Left            =   1560
         Style           =   2  'Dropdown List
         TabIndex        =   39
         TabStop         =   0   'False
         Top             =   810
         Width           =   1500
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
         Index           =   6
         Left            =   1560
         Style           =   2  'Dropdown List
         TabIndex        =   38
         TabStop         =   0   'False
         Top             =   420
         Width           =   1500
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
         Left            =   30
         Style           =   2  'Dropdown List
         TabIndex        =   28
         TabStop         =   0   'False
         Top             =   420
         Width           =   1500
      End
      Begin VB.Label lblData 
         Caption         =   "lblData(9).caption"
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
         Index           =   9
         Left            =   60
         TabIndex        =   46
         Top             =   1590
         Width           =   2805
      End
      Begin VB.Label lblData 
         Caption         =   "lblData(0).caption"
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
         Left            =   180
         TabIndex        =   27
         Top             =   150
         Width           =   2805
      End
   End
   Begin Threed.SSFrame SSFrame1 
      Height          =   3855
      Left            =   120
      TabIndex        =   4
      Top             =   2370
      Width           =   5475
      _Version        =   65536
      _ExtentX        =   9657
      _ExtentY        =   6800
      _StockProps     =   14
      Caption         =   "Basin Specifications:"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
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
         Height          =   360
         Index           =   10
         Left            =   2115
         TabIndex        =   52
         Text            =   "txtData(10)"
         Top             =   3390
         Width           =   1635
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
         Index           =   10
         Left            =   3780
         Style           =   2  'Dropdown List
         TabIndex        =   51
         TabStop         =   0   'False
         Top             =   3375
         Width           =   1500
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
         Height          =   360
         Index           =   5
         Left            =   2115
         TabIndex        =   24
         Text            =   "txtData(5)"
         Top             =   2775
         Width           =   1635
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
         Index           =   5
         Left            =   3780
         Style           =   2  'Dropdown List
         TabIndex        =   23
         TabStop         =   0   'False
         Top             =   2760
         Width           =   1500
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
         Height          =   360
         Index           =   4
         Left            =   2115
         TabIndex        =   21
         Text            =   "txtData(4)"
         Top             =   2385
         Width           =   1635
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
         Index           =   4
         Left            =   3780
         Style           =   2  'Dropdown List
         TabIndex        =   20
         TabStop         =   0   'False
         Top             =   2370
         Width           =   1500
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
         Height          =   360
         Index           =   3
         Left            =   2115
         TabIndex        =   18
         Text            =   "txtData(3)"
         Top             =   1995
         Width           =   1635
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
         Index           =   3
         Left            =   3780
         Style           =   2  'Dropdown List
         TabIndex        =   17
         TabStop         =   0   'False
         Top             =   1980
         Width           =   1500
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
         Height          =   360
         Index           =   2
         Left            =   2115
         TabIndex        =   15
         Text            =   "txtData(2)"
         Top             =   1365
         Width           =   1635
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
         Left            =   3780
         Style           =   2  'Dropdown List
         TabIndex        =   14
         TabStop         =   0   'False
         Top             =   1350
         Width           =   1500
      End
      Begin Threed.SSFrame SSFrame2 
         Height          =   1005
         Left            =   90
         TabIndex        =   8
         Top             =   270
         Width           =   5295
         _Version        =   65536
         _ExtentX        =   9340
         _ExtentY        =   1773
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
            Left            =   3690
            Style           =   2  'Dropdown List
            TabIndex        =   12
            TabStop         =   0   'False
            Top             =   480
            Width           =   1500
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
            Height          =   360
            Index           =   1
            Left            =   2025
            TabIndex        =   11
            Text            =   "txtData(1)"
            Top             =   495
            Width           =   1635
         End
         Begin Threed.SSOption opt_IsCovered 
            Height          =   345
            Index           =   0
            Left            =   120
            TabIndex        =   9
            TabStop         =   0   'False
            Top             =   180
            Width           =   1500
            _Version        =   65536
            _ExtentX        =   2646
            _ExtentY        =   609
            _StockProps     =   78
            Caption         =   "Uncovered"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   9.76
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Enabled         =   0   'False
         End
         Begin Threed.SSOption opt_IsCovered 
            Height          =   345
            Index           =   1
            Left            =   120
            TabIndex        =   10
            TabStop         =   0   'False
            Top             =   495
            Width           =   1500
            _Version        =   65536
            _ExtentX        =   2646
            _ExtentY        =   609
            _StockProps     =   78
            Caption         =   "Covered"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   9.76
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Enabled         =   0   'False
         End
         Begin VB.Label lblData 
            Caption         =   "Ventilation Rate:"
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
            Left            =   2050
            TabIndex        =   13
            Top             =   225
            Width           =   2805
         End
      End
      Begin VB.Label lblData 
         Alignment       =   1  'Right Justify
         Caption         =   "Effluent solids concentration:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   525
         Index           =   10
         Left            =   90
         TabIndex        =   53
         Top             =   3195
         Width           =   1905
      End
      Begin VB.Label lblData 
         Alignment       =   1  'Right Justify
         Caption         =   "SOTR:"
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
         Index           =   5
         Left            =   30
         TabIndex        =   25
         Top             =   2820
         Width           =   1965
      End
      Begin VB.Label lblData 
         Alignment       =   1  'Right Justify
         Caption         =   "Recycle Flow Rate:"
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
         Index           =   4
         Left            =   30
         TabIndex        =   22
         Top             =   2430
         Width           =   1965
      End
      Begin VB.Label lblData 
         Alignment       =   1  'Right Justify
         Caption         =   "Secondary Clarifier Wastage Flow Rate:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Index           =   3
         Left            =   30
         TabIndex        =   19
         Top             =   1800
         Width           =   1965
      End
      Begin VB.Label lblData 
         Alignment       =   1  'Right Justify
         Caption         =   "Depth:"
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
         Left            =   30
         TabIndex        =   16
         Top             =   1410
         Width           =   1965
      End
   End
   Begin Threed.SSPanel SSPanel2 
      Align           =   2  'Align Bottom
      Height          =   405
      Left            =   0
      TabIndex        =   0
      Top             =   6390
      Width           =   9480
      _Version        =   65536
      _ExtentX        =   16722
      _ExtentY        =   714
      _StockProps     =   15
      ForeColor       =   -2147483640
      BackColor       =   12632256
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
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
         ForeColor       =   -2147483640
         BackColor       =   12632256
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
      Begin Threed.SSPanel sspanel_Status 
         Height          =   285
         Left            =   2220
         TabIndex        =   2
         Top             =   60
         Width           =   5000
         _Version        =   65536
         _ExtentX        =   8819
         _ExtentY        =   503
         _StockProps     =   15
         Caption         =   "sspanel_Status"
         ForeColor       =   -2147483640
         BackColor       =   12632256
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
   Begin Threed.SSFrame SSFrame6 
      Height          =   915
      Left            =   120
      TabIndex        =   3
      Top             =   1380
      Width           =   2055
      _Version        =   65536
      _ExtentX        =   3625
      _ExtentY        =   1614
      _StockProps     =   14
      Caption         =   "Number of Basins:"
      ForeColor       =   -2147483640
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
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
         Left            =   180
         TabIndex        =   7
         Text            =   "txtData(0)"
         Top             =   390
         Width           =   1425
      End
      Begin Spin.SpinButton spnData 
         Height          =   300
         Index           =   0
         Left            =   1620
         TabIndex        =   6
         Top             =   390
         Width           =   255
         _Version        =   65536
         _ExtentX        =   450
         _ExtentY        =   529
         _StockProps     =   73
      End
   End
   Begin Threed.SSFrame SSFrame4 
      Height          =   915
      Left            =   2160
      TabIndex        =   29
      Top             =   1380
      Width           =   3435
      _Version        =   65536
      _ExtentX        =   6059
      _ExtentY        =   1614
      _StockProps     =   14
      Caption         =   "Select Aeration Mechanism:"
      ForeColor       =   -2147483640
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Begin VB.ComboBox cbo_Model_Type 
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
         Left            =   120
         Style           =   2  'Dropdown List
         TabIndex        =   30
         TabStop         =   0   'False
         Top             =   360
         Width           =   2925
      End
   End
   Begin Threed.SSFrame SSFrame5 
      Height          =   4845
      Left            =   5580
      TabIndex        =   31
      Top             =   1380
      Width           =   3735
      _Version        =   65536
      _ExtentX        =   6588
      _ExtentY        =   8546
      _StockProps     =   14
      Caption         =   "Additional Specifications:"
      ForeColor       =   -2147483640
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
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
         Height          =   360
         Index           =   8
         Left            =   795
         TabIndex        =   36
         Text            =   "txtData(8)"
         Top             =   2175
         Width           =   1635
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
         Height          =   360
         Index           =   7
         Left            =   795
         TabIndex        =   34
         Text            =   "txtData(7)"
         Top             =   1395
         Width           =   1635
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
         Height          =   360
         Index           =   6
         Left            =   795
         TabIndex        =   32
         Text            =   "txtData(6)"
         Top             =   615
         Width           =   1635
      End
      Begin Threed.SSFrame SSFrame7 
         Height          =   1485
         Left            =   90
         TabIndex        =   41
         Top             =   3180
         Width           =   3555
         _Version        =   65536
         _ExtentX        =   6271
         _ExtentY        =   2619
         _StockProps     =   14
         Caption         =   "Number of CSTRs:"
         ForeColor       =   -2147483640
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
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
            Index           =   9
            Left            =   180
            TabIndex        =   42
            Text            =   "txtData(9)"
            Top             =   390
            Width           =   1425
         End
         Begin Spin.SpinButton spnData 
            Height          =   300
            Index           =   9
            Left            =   1620
            TabIndex        =   43
            Top             =   390
            Width           =   255
            _Version        =   65536
            _ExtentX        =   450
            _ExtentY        =   529
            _StockProps     =   73
         End
         Begin Threed.SSCommand cmdNonuniformCSTRs 
            Height          =   495
            Left            =   180
            TabIndex        =   44
            TabStop         =   0   'False
            Top             =   810
            Width           =   3195
            _Version        =   65536
            _ExtentX        =   5636
            _ExtentY        =   873
            _StockProps     =   78
            Caption         =   "Non-Uniform CSTR Properties    >>"
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
      End
      Begin Threed.SSCommand cmdCalcBiomassConc 
         Height          =   495
         Left            =   270
         TabIndex        =   50
         TabStop         =   0   'False
         Top             =   2640
         Width           =   3195
         _Version        =   65536
         _ExtentX        =   5636
         _ExtentY        =   873
         _StockProps     =   78
         Caption         =   "Calculate Biomass Concentration"
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
      Begin VB.Label lblData 
         Caption         =   "Average Biomass Conc., mg/L:"
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
         Index           =   8
         Left            =   540
         TabIndex        =   37
         Top             =   1890
         Width           =   2925
      End
      Begin VB.Label lblData 
         Caption         =   "Total Gas Flow Rate, L/min:"
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
         Index           =   7
         Left            =   540
         TabIndex        =   35
         Top             =   1110
         Width           =   2925
      End
      Begin VB.Label lblData 
         Caption         =   "Total Volume, liters:"
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
         Index           =   6
         Left            =   540
         TabIndex        =   33
         Top             =   330
         Width           =   2925
      End
   End
   Begin Threed.SSCommand cmdCancelOK 
      Height          =   495
      Index           =   1
      Left            =   6720
      TabIndex        =   47
      TabStop         =   0   'False
      ToolTipText     =   "Click here to save the changes to this window"
      Top             =   60
      Width           =   1305
      _Version        =   65536
      _ExtentX        =   2302
      _ExtentY        =   873
      _StockProps     =   78
      Caption         =   "&OK"
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
   Begin Threed.SSCommand cmdCancelOK 
      Height          =   495
      Index           =   0
      Left            =   8010
      TabIndex        =   48
      TabStop         =   0   'False
      ToolTipText     =   "Click here to abandon any changes on this window"
      Top             =   60
      Width           =   1305
      _Version        =   65536
      _ExtentX        =   2302
      _ExtentY        =   873
      _StockProps     =   78
      Caption         =   "&Cancel"
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
   Begin Threed.SSCommand cmdCancelOK 
      Height          =   495
      Index           =   2
      Left            =   8010
      TabIndex        =   49
      TabStop         =   0   'False
      ToolTipText     =   "Click here for help"
      Top             =   600
      Width           =   1305
      _Version        =   65536
      _ExtentX        =   2302
      _ExtentY        =   873
      _StockProps     =   78
      Caption         =   "&Help"
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
      Caption         =   $"D5_AerationBasin.frx":0000
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1065
      Left            =   150
      TabIndex        =   5
      Top             =   90
      Width           =   5595
   End
End
Attribute VB_Name = "frmD5_AerationBasin"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim USER_HIT_CANCEL As Boolean
Dim USER_HIT_OK As Boolean
Dim frmD5_AerationBasin_Is_Dirty As Boolean

Dim Temp_Plant As TYPE_PlantDiagram

Public HALT_opt_IsCovered As Boolean
Public HALT_cbo_Model_Type As Boolean





Const frmD5_AerationBasin_declarations_end = True


Sub frmD5_AerationBasin_Edit( _
    OUTPUT_Raise_Dirty_Flag As Boolean)
  Temp_Plant = NowProj.Plant
  frmD5_AerationBasin.Show 1
  If (USER_HIT_OK) Then
    OUTPUT_Raise_Dirty_Flag = True
    NowProj.Plant = Temp_Plant
  Else
    OUTPUT_Raise_Dirty_Flag = False
  End If
End Sub


Sub frmD5_AerationBasin_PopulateUnits()
Dim Frm As Form
Set Frm = frmD5_AerationBasin
  '
  ' MAIN DATA BLOCK.
  '
  With Temp_Plant.AerationBasin
    Call unitsys_register(Frm, lblData(0), txtData(0), Nothing, "", _
        "", "", "0", "0", 100#, False)
    Call unitsys_register(Frm, lblData(1), txtData(1), cboUnits(1), "flow_volumetric", _
        .UnitsOfDisplay(1), "L/min", "", "", 100#, True)
    Call unitsys_register(Frm, lblData(2), txtData(2), cboUnits(2), "length", _
        .UnitsOfDisplay(2), "m", "", "", 100#, True)
    Call unitsys_register(Frm, lblData(3), txtData(3), cboUnits(3), "flow_volumetric", _
        .UnitsOfDisplay(3), "L/d", "", "", 100#, True)
    Call unitsys_register(Frm, lblData(4), txtData(4), cboUnits(4), "flow_volumetric", _
        .UnitsOfDisplay(4), "L/d", "", "", 100#, True)
    Call unitsys_register(Frm, lblData(5), txtData(5), cboUnits(5), "flow_mass", _
        .UnitsOfDisplay(5), "kg/hr", "", "", 100#, True)
    Call unitsys_register(Frm, lblData(6), txtData(6), Nothing, "", _
        "", "", "", "", 100#, False)
    Call unitsys_register(Frm, lblData(7), txtData(7), Nothing, "", _
        "", "", "", "", 100#, False)
    Call unitsys_register(Frm, lblData(8), txtData(8), Nothing, "", _
        "", "", "", "", 100#, False)
    Call unitsys_register(Frm, lblData(9), txtData(9), Nothing, "", _
        "", "", "0", "0", 100#, False)
  End With
  With Temp_Plant.SecondaryClarifier
    Call unitsys_register(Frm, lblData(10), txtData(10), cboUnits(10), "concentration", _
        .UnitsOfDisplay(5), "mg/L", "", "", 100#, True)
  End With
End Sub
Sub Store_Unit_Settings()
Dim i As Integer
  With Temp_Plant.AerationBasin
    For i = 1 To 5
      .UnitsOfDisplay(i) = unitsys_get_units(cboUnits(i))
    Next i
  End With
  With Temp_Plant.SecondaryClarifier
    .UnitsOfDisplay(5) = unitsys_get_units(cboUnits(10))
  End With
End Sub


Sub Populate_cbo_Model_Type()
Dim Ctl As Control
Set Ctl = cbo_Model_Type
  HALT_cbo_Model_Type = True
  Ctl.Clear
  Ctl.AddItem "Surface": Ctl.ItemData(Ctl.NewIndex) = BASIN_MODEL_TYPE_SURFACE
  Ctl.AddItem "Diff. Bubble": Ctl.ItemData(Ctl.NewIndex) = BASIN_MODEL_TYPE_DIFFBUBBLE
  HALT_cbo_Model_Type = False
End Sub


Sub Local_DirtyStatus_Set(DirtyFlag As Boolean, NewSetting As Boolean)
  Call Global_DirtyStatus_Set(Me, DirtyFlag, NewSetting)
End Sub
Sub Local_GenericStatus_Set(NewString As String)
  Call Global_GenericStatus_Set(Me, NewString)
End Sub


Private Sub cbo_Model_Type_Click()
Dim Ctl As Control
Set Ctl = cbo_Model_Type
  If (HALT_cbo_Model_Type) Then Exit Sub
  If (Val(Ctl.Tag) = Ctl.ListIndex) Then Exit Sub
  With Temp_Plant.AerationBasin
    .ModelingMechanism = Ctl.ItemData(Ctl.ListIndex)
  End With
  'RAISE DIRTY FLAG AND REFRESH WINDOW.
  Call Local_DirtyStatus_Set(frmD5_AerationBasin_Is_Dirty, True)
  Call frmD5_AerationBasin_Refresh(Temp_Plant)
End Sub


Private Sub cboUnits_Click(Index As Integer)
Dim Ctl As Control
Set Ctl = cboUnits(Index)
  Call unitsys_control_cbox_click(Ctl)
End Sub
Private Sub cboUnits_KeyPress(Index As Integer, KeyAscii As Integer)
  KeyAscii = Global_TextKeyPress(KeyAscii)
End Sub


Private Sub cmdCalcBiomassConc_Click()
Dim OUTPUT_Raise_Dirty_Flag As Boolean
Dim Temp_Total As Double
Dim i As Integer
  OUTPUT_Raise_Dirty_Flag = False
  frmD5_AerationBasin_Temp_Plant = Temp_Plant
  Call frmD5B_Biomass.frmD5B_Biomass_Edit( _
      INPUT_UseWhichStructure_D5, _
      OUTPUT_Raise_Dirty_Flag)
  If (OUTPUT_Raise_Dirty_Flag = True) Then
    '
    ' TRANSFER DATA.
    '
    Temp_Plant = frmD5_AerationBasin_Temp_Plant
    '
    ' TOTAL THE BIOMASS COLUMN.
    '
    Temp_Total = 0#
    For i = 0 To Temp_Plant.AerationBasin.CSTR.Count - 1
      Temp_Total = Temp_Total + Temp_Plant.AerationBasin.CSTR.BioMass(i)
    Next i
    Temp_Plant.AerationBasin.BioMass = Temp_Total
    ''''Call AssignTextAndTag(txtDBL(7), g_AerationBasin.BioMass)
    '
    ' THROW DIRTY FLAG AND REFRESH WINDOW.
    '
    Call Local_DirtyStatus_Set(frmD5_AerationBasin_Is_Dirty, True)
    Call frmD5_AerationBasin_Refresh(Temp_Plant)
  End If
End Sub
Private Sub cmdCancelOK_Click(Index As Integer)
Dim i As Integer
  Select Case Index
    Case 0:     'CANCEL.
      USER_HIT_CANCEL = True
      USER_HIT_OK = False
      Unload Me
      Exit Sub
    Case 1:     'OK.
      '
      ' STORE ALL UNIT SETTINGS.
      '
      Call Store_Unit_Settings
      '
      ' EXIT OUT OF HERE.
      '
      USER_HIT_CANCEL = False
      USER_HIT_OK = True
      Unload Me
      Exit Sub
    Case 2:     'HELP.
      SendKeys "{F1}"
  End Select
End Sub
Private Sub cmdNonuniformCSTRs_Click()
Dim OUTPUT_Raise_Dirty_Flag As Boolean
Dim Temp_Total As Double
Dim i As Integer
  OUTPUT_Raise_Dirty_Flag = False
  frmD5_AerationBasin_Temp_Plant = Temp_Plant
  Call frmD5A_CSTR.frmD5A_CSTR_Edit(OUTPUT_Raise_Dirty_Flag)
  If (OUTPUT_Raise_Dirty_Flag = True) Then
    '
    ' TRANSFER DATA.
    '
    Temp_Plant = frmD5_AerationBasin_Temp_Plant
    '
    ' THROW DIRTY FLAG AND REFRESH WINDOW.
    '
    Call Local_DirtyStatus_Set(frmD5_AerationBasin_Is_Dirty, True)
    Call frmD5_AerationBasin_Refresh(Temp_Plant)
  End If
End Sub


Private Sub Form_Load()
  '
  ' MISC INITS.
  '
  Call CenterOnForm(Me, frmMain)
  Call Local_DirtyStatus_Set(frmD5_AerationBasin_Is_Dirty, False)
  Call Local_GenericStatus_Set("")
  HALT_opt_IsCovered = False
  HALT_cbo_Model_Type = False
  Call Populate_cbo_Model_Type
  '
  ' POPULATE UNIT CONTROLS.
  '
  Call frmD5_AerationBasin_PopulateUnits
  '
  ' REFRESH DISPLAY.
  '
  Call frmD5_AerationBasin_Refresh(Temp_Plant)
End Sub
Private Sub Form_Unload(Cancel As Integer)
  Call unitsys_unregister_all_on_form(Me)
End Sub


Private Sub opt_IsCovered_Click(Index As Integer, Value As Integer)
Dim Ctl0 As Control
Dim Ctl1 As Control
Set Ctl0 = opt_IsCovered(0)
Set Ctl1 = opt_IsCovered(1)
Dim NewTag As Integer
Dim NewSetting As Integer
  If (HALT_opt_IsCovered) Then Exit Sub
  NewTag = Index
  If (CInt(Val(Ctl0.Tag)) <> NewTag) Then
    NewSetting = IIf(NewTag = 0, False, True)
    With Temp_Plant.AerationBasin
      .IsCovered = NewSetting
    End With
    'RAISE DIRTY FLAG AND REFRESH WINDOW.
    Call Local_DirtyStatus_Set(frmD5_AerationBasin_Is_Dirty, True)
    Call frmD5_AerationBasin_Refresh(Temp_Plant)
  End If
End Sub


Private Sub spnData_SpinDown(Index As Integer)
Dim Made_Dirty As Boolean
  Made_Dirty = False
  Select Case Index
    Case 0:
      With Temp_Plant.AerationBasin
        If (.Count > 1) Then
          .Count = .Count - 1
          Made_Dirty = True
        End If
      End With
    Case 9:
      With Temp_Plant.AerationBasin.CSTR
        If (.Count > 1) Then
          .Count = .Count - 1
          Made_Dirty = True
        End If
      End With
  End Select
  If (Made_Dirty) Then
    '
    ' THROW DIRTY FLAG AND REFRESH WINDOW.
    '
    Call Local_DirtyStatus_Set(frmD5_AerationBasin_Is_Dirty, True)
    Call frmD5_AerationBasin_Refresh(Temp_Plant)
  End If
End Sub
Private Sub spnData_SpinUp(Index As Integer)
Dim Made_Dirty As Boolean
  Made_Dirty = False
  Select Case Index
    Case 0:
      With Temp_Plant.AerationBasin
        If (.Count < AERATIONBASIN_MAX_BASIN) Then
          .Count = .Count + 1
          Made_Dirty = True
        End If
      End With
    Case 9:
      With Temp_Plant.AerationBasin.CSTR
        If (.Count < AERATIONBASIN_MAX_CSTR) Then
          .Count = .Count + 1
          Made_Dirty = True
        End If
      End With
  End Select
  If (Made_Dirty) Then
    '
    ' THROW DIRTY FLAG AND REFRESH WINDOW.
    '
    Call Local_DirtyStatus_Set(frmD5_AerationBasin_Is_Dirty, True)
    Call frmD5_AerationBasin_Refresh(Temp_Plant)
  End If
End Sub


Private Sub txtData_GotFocus(Index As Integer)
Dim Ctl As Control
Set Ctl = txtData(Index)
Dim StatusMessagePanel As String
  If (Ctl.Locked = True) Then
    Call Local_GenericStatus_Set("")
    Exit Sub
  End If
  Call unitsys_control_txtx_gotfocus(Ctl)
  Select Case Index
    '
    ' MAIN DATA BLOCK.
    '
    Case 0:
      StatusMessagePanel = ""
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
  If (Ctl.Locked = True) Then
    Call Local_GenericStatus_Set("")
    Exit Sub
  End If
  'NOTE: LOW AND HIGH VALUES IN BASE UNITS
  Select Case Index
    ' MAIN DATA BLOCK.
    Case 0: Val_Low = CDbl(1): Val_High = CDbl(AERATIONBASIN_MAX_BASIN)
    Case 1: Val_Low = 1E-20: Val_High = 1E+20
    Case 2: Val_Low = 1E-20: Val_High = 1E+20
    Case 3: Val_Low = 1E-20: Val_High = 1E+20
    Case 4: Val_Low = 1E-20: Val_High = 1E+20
    Case 5: Val_Low = 1E-20: Val_High = 1E+20
    Case 6: Val_Low = 1E-20: Val_High = 1E+20
    Case 7: Val_Low = 1E-20: Val_High = 1E+20
    Case 8: Val_Low = 1E-20: Val_High = 1E+20
    Case 9: Val_Low = CDbl(1): Val_High = CDbl(AERATIONBASIN_MAX_CSTR)
    Case 10: Val_Low = 0#: Val_High = 1E+20
  End Select
  NewValue_Okay = False
  If (unitsys_control_txtx_lostfocus_validate(Ctl, Val_Low, Val_High, NewValue, Raise_Dirty_Flag)) Then
    NewValue_Okay = True
  End If
  Call unitsys_control_txtx_lostfocus(Ctl, NewValue)
  Call Local_GenericStatus_Set("")
  If (NewValue_Okay) Then
    If (Raise_Dirty_Flag) Then
      'STORE TO MEMORY.
      With Temp_Plant.AerationBasin
        Select Case Index
          '
          ' MAIN DATA BLOCK.
          '
          Case 0: .Count = CInt(NewValue)
          Case 1: .VentilationRate = NewValue
          Case 2: .Depth = NewValue
          Case 3: .WastageFlow = NewValue
          Case 4: .RecycleFlow = NewValue
          Case 5: .SOTR = NewValue
          Case 6: .Volume = NewValue
          Case 7: .GasFlow = NewValue
          Case 8: .BioMass = NewValue
          Case 9: .CSTR.Count = CInt(NewValue)
          Case 10:
            With Temp_Plant.SecondaryClarifier
              .EffluentSolidsConc = NewValue
            End With
        End Select
      End With
      If (Raise_Dirty_Flag) Then
        'THROW DIRTY FLAG.
        Call Local_DirtyStatus_Set(frmD5_AerationBasin_Is_Dirty, True)
      End If
      'REFRESH WINDOW.
      Call frmD5_AerationBasin_Refresh(Temp_Plant)
    End If
  End If
End Sub

