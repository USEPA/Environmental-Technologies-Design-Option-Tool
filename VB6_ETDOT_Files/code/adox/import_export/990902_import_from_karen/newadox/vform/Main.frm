VERSION 5.00
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "THREED32.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmMain 
   Caption         =   "AdOx Process Software"
   ClientHeight    =   5475
   ClientLeft      =   1635
   ClientTop       =   2550
   ClientWidth     =   8865
   Icon            =   "Main.frx":0000
   LinkTopic       =   "Form1"
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   5475
   ScaleWidth      =   8865
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
      Left            =   2850
      Style           =   2  'Dropdown List
      TabIndex        =   51
      TabStop         =   0   'False
      Top             =   3090
      Width           =   1095
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
      Left            =   2850
      Style           =   2  'Dropdown List
      TabIndex        =   50
      TabStop         =   0   'False
      Top             =   3510
      Width           =   1095
   End
   Begin Threed.SSFrame SSFrame1 
      Height          =   1005
      Left            =   7920
      TabIndex        =   20
      Top             =   5250
      Visible         =   0   'False
      Width           =   2985
      _Version        =   65536
      _ExtentX        =   5265
      _ExtentY        =   1773
      _StockProps     =   14
      Caption         =   "Invisible"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Begin VB.PictureBox Picture1_nolongerused 
         Height          =   345
         Left            =   450
         ScaleHeight     =   285
         ScaleWidth      =   14805
         TabIndex        =   56
         Top             =   540
         Width           =   14865
      End
      Begin MSComDlg.CommonDialog CommonDialog1 
         Left            =   120
         Top             =   330
         _ExtentX        =   847
         _ExtentY        =   847
         _Version        =   393216
      End
      Begin Threed.SSPanel panDirty_to_be_deleted 
         Height          =   285
         Left            =   900
         TabIndex        =   55
         Top             =   270
         Width           =   1845
         _Version        =   65536
         _ExtentX        =   3254
         _ExtentY        =   503
         _StockProps     =   15
         Caption         =   "Data Unchanged"
         BackColor       =   12632256
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
   End
   Begin Threed.SSFrame SSFrame2 
      Height          =   2595
      Left            =   0
      TabIndex        =   21
      Top             =   120
      Width           =   4005
      _Version        =   65536
      _ExtentX        =   7064
      _ExtentY        =   4577
      _StockProps     =   14
      Caption         =   "Reactor Properties:"
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
         Index           =   0
         Left            =   2880
         Style           =   2  'Dropdown List
         TabIndex        =   49
         TabStop         =   0   'False
         Top             =   810
         Width           =   1095
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
         Left            =   2880
         Style           =   2  'Dropdown List
         TabIndex        =   48
         TabStop         =   0   'False
         Top             =   1230
         Width           =   1095
      End
      Begin VB.TextBox txtData 
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
         Left            =   1740
         TabIndex        =   4
         Text            =   "txtData(10)"
         Top             =   2130
         Width           =   1095
      End
      Begin VB.TextBox txtData 
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
         Left            =   1740
         TabIndex        =   3
         Text            =   "txtData(5)"
         Top             =   1680
         Width           =   1095
      End
      Begin VB.ComboBox cboReactorType 
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
         Left            =   1740
         Style           =   2  'Dropdown List
         TabIndex        =   0
         Top             =   330
         Width           =   1905
      End
      Begin VB.TextBox txtData 
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
         Left            =   1740
         TabIndex        =   2
         Text            =   "txtData(1)"
         Top             =   1230
         Width           =   1095
      End
      Begin VB.TextBox txtData 
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
         Left            =   1740
         TabIndex        =   1
         Text            =   "txtData(0)"
         Top             =   795
         Width           =   1095
      End
      Begin VB.Label lbldesc 
         Caption         =   "Number of Tanks:"
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
         Index           =   10
         Left            =   120
         TabIndex        =   47
         Top             =   2175
         Width           =   1755
      End
      Begin VB.Label lbldesc2 
         Caption         =   "tanks"
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
         Index           =   10
         Left            =   3120
         TabIndex        =   46
         Top             =   3990
         Width           =   825
      End
      Begin VB.Label lbldesc2 
         Caption         =   "gmol/L"
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
         Left            =   3300
         TabIndex        =   32
         Top             =   3600
         Width           =   825
      End
      Begin VB.Label lbldesc 
         Caption         =   "{Influent} H2O2:"
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
         Left            =   120
         TabIndex        =   31
         Top             =   1725
         Width           =   1755
      End
      Begin VB.Label lbldesc2 
         Caption         =   "minutes"
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
         Left            =   3120
         TabIndex        =   26
         Top             =   4020
         Width           =   825
      End
      Begin VB.Label lbldesc 
         Caption         =   "Retention Time:"
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
         Left            =   120
         TabIndex        =   25
         Top             =   1275
         Width           =   1755
      End
      Begin VB.Label lbldesc2 
         AutoSize        =   -1  'True
         Caption         =   "liters"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   0
         Left            =   3240
         TabIndex        =   24
         Top             =   4260
         Width           =   420
      End
      Begin VB.Label lbldesc 
         Caption         =   "Volume:"
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
         Left            =   120
         TabIndex        =   23
         Top             =   840
         Width           =   1755
      End
      Begin VB.Label lbldesc 
         Caption         =   "Reactor Type:"
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
         Index           =   100
         Left            =   120
         TabIndex        =   22
         Top             =   405
         Width           =   1755
      End
   End
   Begin Threed.SSFrame SSFrame3 
      Height          =   2115
      Left            =   0
      TabIndex        =   27
      Top             =   2730
      Width           =   4005
      _Version        =   65536
      _ExtentX        =   7064
      _ExtentY        =   3731
      _StockProps     =   14
      Caption         =   "Numerical Simulation Parameters:"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Begin VB.TextBox txtData 
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
         Left            =   1740
         TabIndex        =   7
         Text            =   "txtData(4)"
         Top             =   1470
         Width           =   1095
      End
      Begin VB.TextBox txtData 
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
         Left            =   1740
         TabIndex        =   6
         Text            =   "txtData(3)"
         Top             =   780
         Width           =   1095
      End
      Begin VB.TextBox txtData 
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
         Left            =   1740
         TabIndex        =   5
         Text            =   "txtData(2)"
         Top             =   345
         Width           =   1095
      End
      Begin VB.Label lbldesc 
         Caption         =   "# Retention Times to Simulate:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   540
         Index           =   4
         Left            =   120
         TabIndex        =   30
         Top             =   1275
         Width           =   1755
      End
      Begin VB.Label lbldesc 
         Caption         =   "Final Time:"
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
         Left            =   120
         TabIndex        =   29
         Top             =   840
         Width           =   1755
      End
      Begin VB.Label lbldesc 
         Caption         =   "Time Step:"
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
         Left            =   120
         TabIndex        =   28
         Top             =   420
         Width           =   1755
      End
   End
   Begin Threed.SSFrame SSFrame4 
      Height          =   2595
      Left            =   4020
      TabIndex        =   33
      Top             =   120
      Width           =   4815
      _Version        =   65536
      _ExtentX        =   8493
      _ExtentY        =   4577
      _StockProps     =   14
      Caption         =   "Water Quality Properties:"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Begin VB.TextBox txtData 
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
         Left            =   2670
         TabIndex        =   9
         Text            =   "txtData(7)"
         Top             =   780
         Width           =   1095
      End
      Begin VB.TextBox txtData 
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
         Left            =   2670
         TabIndex        =   8
         Text            =   "txtData(6)"
         Top             =   360
         Width           =   1095
      End
      Begin Threed.SSFrame ssframe_tic 
         Height          =   795
         Left            =   120
         TabIndex        =   38
         Top             =   1590
         Width           =   4485
         _Version        =   65536
         _ExtentX        =   7911
         _ExtentY        =   1402
         _StockProps     =   14
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
            Left            =   2550
            TabIndex        =   12
            Text            =   "txtData(8)"
            Top             =   285
            Width           =   1095
         End
         Begin VB.Label lbldesc 
            Caption         =   "{Influent} TIC Concentration:"
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
            Left            =   120
            TabIndex        =   40
            Top             =   330
            Width           =   2445
         End
         Begin VB.Label lbldesc2 
            Caption         =   "gmol/L"
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
            Left            =   3750
            TabIndex        =   39
            Top             =   330
            Width           =   675
         End
      End
      Begin Threed.SSFrame ssframe_alk 
         Height          =   795
         Left            =   4140
         TabIndex        =   41
         Top             =   1050
         Width           =   4485
         _Version        =   65536
         _ExtentX        =   7911
         _ExtentY        =   1402
         _StockProps     =   14
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
            Left            =   2400
            TabIndex        =   13
            Text            =   "txtData(9)"
            Top             =   285
            Width           =   1095
         End
         Begin VB.Label lbldesc2 
            Caption         =   "mg/L as CaCO3"
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
            Index           =   37
            Left            =   3600
            TabIndex        =   43
            Top             =   210
            Width           =   795
         End
         Begin VB.Label lbldesc 
            Caption         =   "{Influent} Alkalinity:"
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
            Left            =   120
            TabIndex        =   42
            Top             =   330
            Width           =   1785
         End
      End
      Begin Threed.SSOption optTICInput 
         Height          =   285
         Index           =   1
         Left            =   2340
         TabIndex        =   11
         Top             =   1260
         Width           =   1305
         _Version        =   65536
         _ExtentX        =   2302
         _ExtentY        =   503
         _StockProps     =   78
         Caption         =   "Alkalinity"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   9.76
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin Threed.SSOption optTICInput 
         Height          =   285
         Index           =   0
         Left            =   1470
         TabIndex        =   10
         Top             =   1260
         Width           =   765
         _Version        =   65536
         _ExtentX        =   1349
         _ExtentY        =   503
         _StockProps     =   78
         Caption         =   "TIC"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   9.76
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Value           =   -1  'True
      End
      Begin VB.Label lbldesc 
         Caption         =   "TIC Input As:"
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
         Index           =   49
         Left            =   120
         TabIndex        =   37
         Top             =   1275
         Width           =   1305
      End
      Begin VB.Label lbldesc 
         Caption         =   "{Influent} Phosphate Conc.:"
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
         Left            =   120
         TabIndex        =   36
         Top             =   825
         Width           =   2565
      End
      Begin VB.Label lbldesc2 
         Caption         =   "gmol/L"
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
         Left            =   3870
         TabIndex        =   35
         Top             =   825
         Width           =   795
      End
      Begin VB.Label lbldesc 
         Caption         =   "{Influent} pH:"
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
         Left            =   120
         TabIndex        =   34
         Top             =   405
         Width           =   2565
      End
   End
   Begin Threed.SSFrame SSFrame5 
      Height          =   1275
      Left            =   4020
      TabIndex        =   44
      Top             =   3570
      Width           =   4785
      _Version        =   65536
      _ExtentX        =   8440
      _ExtentY        =   2249
      _StockProps     =   14
      Caption         =   "NOM and Target Compounds:"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Begin VB.CommandButton cmdTarget 
         Caption         =   "Re&name"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   3
         Left            =   2640
         TabIndex        =   18
         Top             =   360
         Width           =   1485
      End
      Begin VB.CommandButton cmdTarget 
         Caption         =   "&Edit"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   2
         Left            =   1875
         TabIndex        =   17
         Top             =   360
         Width           =   615
      End
      Begin VB.CommandButton cmdTarget 
         Caption         =   "&Delete"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   1
         Left            =   915
         TabIndex        =   16
         Top             =   360
         Width           =   795
      End
      Begin VB.ComboBox cboTarget 
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
         TabIndex        =   19
         Top             =   720
         Width           =   4305
      End
      Begin VB.CommandButton cmdTarget 
         Caption         =   "&Add"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   0
         Left            =   150
         TabIndex        =   15
         Top             =   360
         Width           =   615
      End
   End
   Begin Threed.SSFrame SSFrame6 
      Height          =   855
      Left            =   4020
      TabIndex        =   45
      Top             =   2730
      Width           =   3105
      _Version        =   65536
      _ExtentX        =   5477
      _ExtentY        =   1508
      _StockProps     =   14
      Caption         =   "Photochemical Parameters:"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Begin VB.CommandButton cmdPhotochemical 
         Caption         =   "Edit"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   3
         Left            =   1140
         TabIndex        =   14
         Top             =   390
         Width           =   735
      End
   End
   Begin Threed.SSPanel SSPanel2 
      Align           =   2  'Align Bottom
      Height          =   405
      Left            =   0
      TabIndex        =   52
      Top             =   5070
      Width           =   8865
      _Version        =   65536
      _ExtentX        =   15637
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
         TabIndex        =   53
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
         TabIndex        =   54
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
   Begin Threed.SSFrame SSFrame7 
      Height          =   855
      Left            =   7110
      TabIndex        =   57
      Top             =   2730
      Width           =   1695
      _Version        =   65536
      _ExtentX        =   2990
      _ExtentY        =   1508
      _StockProps     =   14
      Caption         =   "Dye Study:"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Begin VB.CommandButton cmdDyeStudy 
         Caption         =   "Edit"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   0
         Left            =   360
         TabIndex        =   58
         Top             =   390
         Width           =   825
      End
   End
   Begin VB.Menu mnuFile 
      Caption         =   "&File"
      Begin VB.Menu mnuFileItem 
         Caption         =   "&New"
         Index           =   10
         Shortcut        =   ^N
      End
      Begin VB.Menu mnuFileItem 
         Caption         =   "&Open ..."
         Index           =   20
         Shortcut        =   ^O
      End
      Begin VB.Menu mnuFileItem 
         Caption         =   "&Save"
         Index           =   30
         Shortcut        =   ^S
      End
      Begin VB.Menu mnuFileItem 
         Caption         =   "Save &As ..."
         Index           =   40
      End
      Begin VB.Menu mnuFileItem 
         Caption         =   "-"
         Index           =   79
      End
      Begin VB.Menu mnuFileItem 
         Caption         =   "Old File #1"
         Index           =   191
      End
      Begin VB.Menu mnuFileItem 
         Caption         =   "Old File #2"
         Index           =   192
      End
      Begin VB.Menu mnuFileItem 
         Caption         =   "Old File #3"
         Index           =   193
      End
      Begin VB.Menu mnuFileItem 
         Caption         =   "Old File #4"
         Index           =   194
      End
      Begin VB.Menu mnuFileItem 
         Caption         =   "-"
         Index           =   198
      End
      Begin VB.Menu mnuFileItem 
         Caption         =   "E&xit"
         Index           =   199
      End
   End
   Begin VB.Menu mnuRun 
      Caption         =   "&Run"
      Begin VB.Menu mnuRunItem 
         Caption         =   "&Run Simulation ..."
         Index           =   10
         Shortcut        =   ^R
      End
   End
   Begin VB.Menu mnuResults 
      Caption         =   "Re&sults"
      Begin VB.Menu mnuResultsItem 
         Caption         =   "&View Low-Level Results ..."
         Index           =   10
      End
      Begin VB.Menu mnuResultsItem 
         Caption         =   "View Low-Level &Input File ..."
         Index           =   15
      End
      Begin VB.Menu mnuResultsItem 
         Caption         =   "-"
         Index           =   18
      End
      Begin VB.Menu mnuResultsItem 
         Caption         =   "Send Results to E&xcel ..."
         Index           =   20
         Shortcut        =   ^X
      End
      Begin VB.Menu mnuResultsItem 
         Caption         =   "Plo&t Results ..."
         Index           =   30
         Shortcut        =   ^T
      End
   End
   Begin VB.Menu mnuHelp 
      Caption         =   "&Help"
      Begin VB.Menu mnuHelpItem 
         Caption         =   "View Version History"
         Index           =   80
      End
      Begin VB.Menu mnuHelpItem 
         Caption         =   "View Disclaimer"
         Index           =   85
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
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Sub Avoid_Weird_Focus_Problem()
  Call unitsys_control_MostRecent_Force_lostfocus
  'frmMain.SetFocus
  '
  ' NOTE: IT IS VERY IMPORTANT TO SET FOCUS HERE
  ' TO SOME NON-UNITTEXTBOX CONTROL, I.E. DON'T
  ' SET IT TO txtData(0...3), BUT cboUnits(0)
  ' IS OKAY.
  cboReactorType.SetFocus
  'Text1.SetFocus
End Sub


Sub Populate_frmMain_Units()
  Call unitsys_register(frmMain, lbldesc(0), txtData(0), cboUnits(0), "volume", _
     "liter", "liter", "", "", 100#, True)
  Call unitsys_register(frmMain, lbldesc(1), txtData(1), cboUnits(1), "time", _
      "min", "min", "", "", 100#, True)
  Call unitsys_register(frmMain, lbldesc(2), txtData(2), cboUnits(2), "time", _
      "min", "min", "", "", 100#, True)
  Call unitsys_register(frmMain, lbldesc(3), txtData(3), cboUnits(3), "time", _
      "min", "min", "", "", 100#, True)
  Call unitsys_register(frmMain, lbldesc(4), txtData(4), Nothing, "", _
      "", "", "", "", 100#, False)
  Call unitsys_register(frmMain, lbldesc(5), txtData(5), Nothing, "", _
      "", "", "", "", 100#, False)
  Call unitsys_register(frmMain, lbldesc(6), txtData(6), Nothing, "", _
      "", "", "", "", 100#, False)
  Call unitsys_register(frmMain, lbldesc(7), txtData(7), Nothing, "", _
      "", "", "", "", 100#, False)
  Call unitsys_register(frmMain, lbldesc(8), txtData(8), Nothing, "", _
      "", "", "", "", 100#, False)
  Call unitsys_register(frmMain, lbldesc(9), txtData(9), Nothing, "", _
      "", "", "", "", 100#, False)
  Call unitsys_register(frmMain, lbldesc(10), txtData(10), Nothing, "", _
      "", "", "", "", 100#, False)
End Sub

Sub Local_DirtyStatus_Set(DirtyFlag As Boolean, NewSetting As Boolean)
  Call Global_DirtyStatus_Set(Me, DirtyFlag, NewSetting)
End Sub

Sub Local_GenericStatus_Set(NewString As String)
  Call Global_GenericStatus_Set(Me, NewString)
End Sub

Sub populate_cboReactorType()
  cboReactorType.Clear
  cboReactorType.AddItem "CMBR"
  cboReactorType.AddItem "CMFR"
End Sub


Private Sub cboReactorType_Click()
  If (cboReactorType.ListIndex <> val(cboReactorType.Tag)) Then
    Select Case cboReactorType.ListIndex
      Case 0: NowProj.idreact = IDREACT_CMBR
      Case 1: NowProj.idreact = IDREACT_CMFR
    End Select
    Call DirtyFlag_Throw(NowProj)
    Call refresh_frmMain
  End If
End Sub


Private Sub cboUnits_Click(Index As Integer)
Dim ctl As Control
Set ctl = cboUnits(Index)
  Call unitsys_control_cbox_click(ctl)
End Sub


Private Sub cmdDyeStudy_Click(Index As Integer)
Dim retval As Integer

  retval = frmDyeStudy.frmDyeStudy_DoEdit()
  If (retval) Then
    'USER HIT OK; ASSUME THEY MODIFIED SOMETHING.
    
    'REFRESH MAIN WINDOW, ALTHOUGH PROBABLY
    'NOTHING ON THE MAIN WINDOW NEEDS REFRESHING.
    Call refresh_frmMain
    
    'THROW DIRTY FLAG.
   
    If NowProj.dirty Then
       'THROW DIRTY FLAG.
       Call Local_DirtyStatus_Set( _
           Project_Is_Dirty, True)
    End If
    Call DirtyFlag_Refresh(NowProj)
  Else
    'RESTORE DIRTY FLAG DISPLAY IF NEEDED.
    Call DirtyFlag_Refresh(NowProj)
  End If
End Sub

Private Sub cmdPhotochemical_Click(Index As Integer)
Dim retval As Integer
  retval = frmPhotoChem.frmPhotoChem_DoEdit()
  If (retval) Then
    'USER HIT OK; ASSUME THEY MODIFIED SOMETHING.
    
    'REFRESH MAIN WINDOW, ALTHOUGH PROBABLY
    'NOTHING ON THE MAIN WINDOW NEEDS REFRESHING.
    Call refresh_frmMain
    
    'THROW DIRTY FLAG.
    Call DirtyFlag_Throw(NowProj)
  Else
    'RESTORE DIRTY FLAG DISPLAY IF NEEDED.
    Call DirtyFlag_Refresh(NowProj)
  End If
End Sub


Private Sub cmdTarget_Click(Index As Integer)
Dim name_original As String
Dim name_new As String
Dim is_aborted As Integer
Dim idx As Integer
Dim retval As Integer
Dim msg As String
Dim idx_del As Integer
Dim idx_max As Integer
Dim i As Integer
Dim j As Integer
Dim temp() As Double

  Select Case Index
    '- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
    Case 0:   'add
      If (frmTarget.frmTarget_DoAddNew()) Then
        'USER HIT OK.
        'NOTE THAT frmTarget_DoAddNew() TAKES CARE
        'OF ADDING THE ACTUAL COMPONENT TO MEMORY.
        
        'REFRESH MAIN WINDOW, ESPECIALLY THE
        'TARGET COMPOUND SCROLLBOX.
        Call refresh_frmMain
        
        'THROW DIRTY FLAG.
        Call DirtyFlag_Throw(NowProj)
      Else
        'USER HIT CANCEL; DO NOTHING.
      End If
    
    '- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
    Case 1:   'delete
      If (cboTarget.ListIndex < 0) Then Exit Sub
      idx = cboTarget.ListIndex
      name_original = cboTarget.List(idx)
      If (Trim$(UCase$(name_original)) = Trim$(UCase$("NOM"))) Then
        Call Show_Error("You cannot delete the NOM compound.")
        Exit Sub
      End If
      If (NowProj.TargetCompounds_Count <= 2) Then
        Call Show_Error("At a minimum, this list of compounds must contain " & _
            "the NOM compound and one target compound.  You cannot delete " & _
            "the last remaining target compound.")
        Exit Sub
      End If
      msg = "If you delete target compound " & Chr$(34) & _
            Trim$(name_original) & Chr$(34) & " you will not be able to " & _
            "undelete it.  Are you sure you want to delete it?"
      retval = MsgBox(msg, vbCritical + vbYesNo, App.title)
      If (retval = vbNo) Then Exit Sub
      
      'DELETE THIS TARGET COMPOUND STRUCTURE.
      idx_del = idx + 1
      idx_max = NowProj.TargetCompounds_Count
      For i = idx_del To idx_max - 1
        NowProj.TargetCompounds(i) = NowProj.TargetCompounds(i + 1)
      Next i
      ReDim Preserve NowProj.TargetCompounds(1 To idx_max - 1)
      
      'DELETE THIS SET OF EXTINCTION COEFFICIENTS (ONE PER WAVELENGTH).
      'NOTE: THIS STUPID TEMPORARY ARRAY IS NECESSARY TO GET AROUND THE
      'VISUAL BASIC STIPULATION THAT YOU CAN ONLY "REDIM PRESERVE" AN ARRAY
      'IF YOU ARE CHANGING THE *LAST* ARRAY INDEX.
      ReDim temp(1 To idx_max - 1, 1 To NowProj.Wavelength_Count)
      For i = 1 To idx_del - 1
        For j = 1 To NowProj.Wavelength_Count
          temp(i, j) = NowProj.extcoef(i, j)
        Next j
      Next i
      For i = idx_del To idx_max - 1
        For j = 1 To NowProj.Wavelength_Count
          temp(i, j) = NowProj.extcoef(i + 1, j)
        Next j
      Next i
      ReDim NowProj.extcoef(1 To idx_max - 1, 1 To NowProj.Wavelength_Count)
      For i = 1 To idx_max - 1
        For j = 1 To NowProj.Wavelength_Count
          NowProj.extcoef(i, j) = temp(i, j)
        Next j
      Next i
      
      'DELETE THIS SET OF QUANTUM YIELDS (ONE PER WAVELENGTH).
      'NOTE: THIS STUPID TEMPORARY ARRAY IS NECESSARY TO GET AROUND THE
      'VISUAL BASIC STIPULATION THAT YOU CAN ONLY "REDIM PRESERVE" AN ARRAY
      'IF YOU ARE CHANGING THE *LAST* ARRAY INDEX.
      ReDim temp(1 To idx_max - 1, 1 To NowProj.Wavelength_Count)
      For i = 1 To idx_del - 1
        For j = 1 To NowProj.Wavelength_Count
          temp(i, j) = NowProj.quatyd(i, j)
        Next j
      Next i
      For i = idx_del To idx_max - 1
        For j = 1 To NowProj.Wavelength_Count
          temp(i, j) = NowProj.quatyd(i + 1, j)
        Next j
      Next i
      ReDim NowProj.quatyd(1 To idx_max - 1, 1 To NowProj.Wavelength_Count)
      For i = 1 To idx_max - 1
        For j = 1 To NowProj.Wavelength_Count
          NowProj.quatyd(i, j) = temp(i, j)
        Next j
      Next i
      
      'UPDATE THE NUMBER OF TARGET COMPOUNDS.
      NowProj.TargetCompounds_Count = idx_max - 1
    
      'REFRESH MAIN WINDOW, ESPECIALLY THE
      'TARGET COMPOUND SCROLLBOX.
      Call refresh_frmMain
      
      'THROW DIRTY FLAG.
      Call DirtyFlag_Throw(NowProj)
    
    '- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
    Case 2:   'edit
      If (frmTarget.frmTarget_DoEdit(cboTarget.ListIndex + 1)) Then
        'USER HIT OK.
        'NOTE THAT frmTarget_DoAddNew() TAKES CARE
        'OF ADDING THE ACTUAL COMPONENT TO MEMORY.
        
        'REFRESH MAIN WINDOW (PROBABLY NOTHING NEEDS REFRESHING).
        Call refresh_frmMain
        
        'THROW DIRTY FLAG.
        Call DirtyFlag_Throw(NowProj)
      Else
        'USER HIT CANCEL; DO NOTHING.
      End If
    
    '- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
    Case 3:   'rename
      If (cboTarget.ListIndex < 0) Then Exit Sub
      idx = cboTarget.ListIndex
      name_original = cboTarget.List(idx)
      If (Trim$(UCase$(name_original)) = Trim$(UCase$("NOM"))) Then
        Call Show_Error("You cannot rename the NOM compound.")
        Exit Sub
      End If
      Do While (1 = 1)
        name_new = frmNewName.frmNewName_GetName( _
            "Enter New Name for Target Compound", _
            "Each target compound must have a unique name.", _
            name_original, _
            is_aborted)
        If (is_aborted) Then
          'EXIT OUTTA HERE.
          Exit Sub
        End If
        If (Not TargetCompound_IsKeyExist(NowProj, name_new)) Then
          Exit Do
        End If
        Call Show_Error("That name already exists.  Choose another name.")
      Loop
      NowProj.TargetCompounds(idx + 1).comname = name_new
      
      'REFRESH MAIN WINDOW, ESPECIALLY THE
      'TARGET COMPOUND SCROLLBOX.
      Call refresh_frmMain
      
      'THROW DIRTY FLAG.
      Call DirtyFlag_Throw(NowProj)
    
    '- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
  End Select
End Sub

'qstr

Private Sub Form_Load()
  'POPULATE SCROLLBOX.
  Call populate_cboReactorType
  
  'MISC WINDOW INITS.
  Me.Caption = App.title
  Call CenterOnScreen(Me)
  ''''Call PauseStatus_Set("")
  
  'CONSTRUCT A NEW PROJECT.
  Call file_new
  
  Call Populate_frmMain_Units
  Me.sspanel_Dirty = "Data Unchanged"
  Me.sspanel_Status = ""
  
  'POPULATE OLD FILE LIST.
  ''''fn_OldFileList = GetWindowsDir() & "\" & fn_INI_name
  Call OldFileList_Populate( _
      1, _
      frmMain.mnuFileItem(198), _
      frmMain.mnuFileItem(191), _
      frmMain.mnuFileItem(192), _
      frmMain.mnuFileItem(193), _
      frmMain.mnuFileItem(194))
      
  Call refresh_frmMain
  
End Sub


Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
  If (file_query_unload() = False) Then
    Cancel = True
  End If
End Sub
Private Sub Form_Unload(Cancel As Integer)
  Call unitsys_unregister_all_on_form(Me)
End Sub



Private Sub mnuFileItem_Click(Index As Integer)
  Select Case Index
    Case 10:      'New
      If (file_query_unload()) Then
        Call file_new
      End If
    Case 20:      'Open ...
      If (file_query_unload()) Then
        Call File_OpenAs("")
      End If
    Case 30:      'Save
      If (NowProj.Filename = "") Then
        Call File_SaveAs("")
      Else
        Call File_Save(NowProj.Filename)
      End If
    Case 40:      'Save As ...
      Call File_SaveAs("")
    Case 191 To 194:      'Last-few-files list
      If (file_query_unload()) Then
        If (mnuFileItem(Index).visible) Then
          Call File_OpenAs(OldFiles(1, Index - 190))
        End If
      End If
    Case 199:     'exit
     'NOTE: Form_QueryUnload() TAKES CAKE OF THIS.
      'If we do it here, _two_ message boxes will pop up
      'when the user has data which needs saving !
      'If (file_query_unload()) Then
      '  Unload Me
      'End If
      Unload Me
  End Select
End Sub


Private Sub mnuHelpItem_Click(Index As Integer)
Dim fn_this As String
  Select Case Index
    Case 10:      'ONLINE HELP.
      'NOTE: We currently do NOT have the resources to
      'create an online help file for AdDesignS (1/7/98)
      'therefore no online help is available.
      Call Show_Message("Online help is currently unavailable.  " & _
          "Please refer to the printed manual or the Acrobat-format ADS.PDF file.", _
          vbExclamation, App.title)
      Exit Sub
      'Call LaunchFile_General("", MAIN_APP_PATH & "\help\ads.hlp")
    Case 20:      'ONLINE MANUAL.
      fn_this = MAIN_APP_PATH & "\help\ads.pdf"
      If (FileExists(fn_this) = False) Then
        Call Show_Message("The file `" & fn_this & "` is missing.", _
        vbExclamation, App.title)
        Exit Sub
      End If
      Call LaunchFile_General("", fn_this)
    Case 80:
      fn_this = App.Path & "\dbase\readme.txt"
      If (FileExists(fn_this) = False) Then
        Call Show_Message("The file `" & fn_this & "` is missing.", _
        vbExclamation, App.title)
        Exit Sub
      End If
      Call Launch_Notepad(fn_this)
    Case 85:    'VIEW DISCLAIMER.
      'SHOW THE DISCLAIMER WINDOW.
      splash_mode = 101
      splash_button_pressed = 0
      frmSplash.Show 1
    Case 99:    'ABOUT.
      frmAbout.Show 1
  End Select
End Sub


Private Sub mnuResultsItem_Click(Index As Integer)
Dim need_to_check_results As Boolean
  need_to_check_results = False
  Select Case Index
    Case 10:      'View results
      need_to_check_results = True
    Case 15:      'View input file
      need_to_check_results = True
    Case 20:      'Send results to Excel
      need_to_check_results = True
    Case 30:      'Plot results
      need_to_check_results = True
  End Select
  If (need_to_check_results) Then
    Call FortranLink_SetFilenames
    If (FileExists(FortranLink_fn_MainOutput)) Then
      'DO NOTHING--CODE IS BELOW.
    Else
      Call Show_Error("There are no results to view.  Run the simulation first.")
      Exit Sub
    End If
  End If
  Select Case Index
    Case 10:      'View results
      Call Launch_Notepad(FortranLink_fn_MainOutput)
    Case 15:      'View input file
      Call Launch_Notepad(FortranLink_fn_MainInput)
    Case 20:      'Print results to Excel
      frmPrint_DO_INPUTS = True
      frmPrint_DO_OUTPUTS = True
      frmPrint_DO_PLOTS = False
      frmPrint.Show 1
    Case 30:      'Print and plot results to Excel
      frmPlot.Show 1
  End Select
End Sub


Private Sub mnuRunItem_Click(Index As Integer)
Dim msg As String
Dim retval As Integer
  Select Case Index
    Case 10:      'Run Simulation
      msg = "Running a simulation may take several minutes to an hour, " & _
          "depending on the speed of your CPU.  Proceed anyway?"
      retval = MsgBox(msg, vbCritical + vbYesNo, App.title)
      If (retval = vbNo) Then Exit Sub
      Call FortranLink_Run
  End Select
End Sub


Private Sub optTICInput_Click(Index As Integer, Value As Integer)
  If (Index <> val(optTICInput(0).Tag)) Then
    Select Case Index
      Case 0: NowProj.idcarbn = IDCARBN_TIC
      Case 1: NowProj.idcarbn = IDCARBN_ALKALINITY
    End Select
    Call DirtyFlag_Throw(NowProj)
    Call refresh_frmMain
  End If
End Sub


Private Sub txtdata_GotFocus(Index As Integer)
Dim ctl As Control
Set ctl = txtData(Index)
Dim StatusMessagePanel As String
  Call unitsys_control_txtx_gotfocus(ctl)
  Select Case Index
    Case 0
      StatusMessagePanel = "Type in the volume"
    Case 1
      StatusMessagePanel = "Type in the retention time"
    Case 2
      StatusMessagePanel = "Type in the time step"
    Case 3
      StatusMessagePanel = "Type in the final time"
  End Select
  Call Local_GenericStatus_Set(StatusMessagePanel)
End Sub
Private Sub txtdata_KeyPress(Index As Integer, KeyAscii As Integer)
  KeyAscii = Global_NumericKeyPress(KeyAscii)
End Sub
Private Sub txtdata_LostFocus(Index As Integer)
Dim NewValue_Okay As Integer
Dim NewValue As Double
Dim ctl As Control
Set ctl = txtData(Index)
Dim Val_Low As Double
Dim Val_High As Double
Dim Raise_Dirty_Flag As Boolean
Dim Too_Small As Integer
  'NOTE: LOW AND HIGH VALUES IN BASE UNITS
  If (Index = 4) Then
    Val_Low = 1E-20 * 60#
    Val_High = 1E+20 * 60#
  Else
    Val_Low = 1E-20     '0.00000000000000000001
    Val_High = 1E+20    '100000000000000000000
  End If
  NewValue_Okay = False
  If (unitsys_control_txtx_lostfocus_validate(ctl, Val_Low, Val_High, NewValue, Raise_Dirty_Flag)) Then
    NewValue_Okay = True
  End If
  Call unitsys_control_txtx_lostfocus(ctl, NewValue)
  Call Local_GenericStatus_Set("")
  If (NewValue_Okay) Then
    If (Raise_Dirty_Flag) Then
      'STORE TO MEMORY.
          Select Case Index
            Case 0: NowProj.volume = NewValue
            Case 1:
              NowProj.tau = NewValue
              'RECALCULATE FINAL TIME.
              NowProj.ttotal = NowProj.tau * NowProj.xntimes
            Case 2: NowProj.ssize = NewValue * 60#
                    NowProj.opsize = NewValue
            Case 3: NowProj.ttotal = NewValue
            Case 4:
              NowProj.xntimes = NewValue
              'RECALCULATE FINAL TIME.
              NowProj.ttotal = NowProj.tau * NowProj.xntimes
            Case 5: NowProj.inf_h2o2 = NewValue
            Case 6: NowProj.ph0 = NewValue
            Case 7: NowProj.phosph = NewValue
            Case 8: NowProj.ticarbn = NewValue
            Case 9: NowProj.alk = NewValue
            Case 10: NowProj.num_tanks = CInt(val(NewValue))
          End Select
     
      If (Raise_Dirty_Flag) Then
        'THROW DIRTY FLAG.
        Call Local_DirtyStatus_Set( _
            Project_Is_Dirty, True)
      End If
      'REFRESH WINDOW.
      Call refresh_frmMain
    End If
  End If
End Sub



