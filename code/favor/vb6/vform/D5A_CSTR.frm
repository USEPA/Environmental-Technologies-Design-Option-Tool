VERSION 5.00
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "Threed32.ocx"
Object = "{B16553C3-06DB-101B-85B2-0000C009BE81}#1.0#0"; "Spin32.ocx"
Begin VB.Form frmD5A_CSTR 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Define CSTR Parameters"
   ClientHeight    =   6900
   ClientLeft      =   5070
   ClientTop       =   1305
   ClientWidth     =   9075
   ControlBox      =   0   'False
   HelpContextID   =   7000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6900
   ScaleWidth      =   9075
   ShowInTaskbar   =   0   'False
   Begin Threed.SSFrame SSFrame1 
      Height          =   5055
      Left            =   120
      TabIndex        =   12
      Top             =   1320
      Width           =   8775
      _Version        =   65536
      _ExtentX        =   15478
      _ExtentY        =   8916
      _StockProps     =   14
      Caption         =   "CSTR Parameters:"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Begin Threed.SSPanel DisplayPanel 
         Height          =   4155
         Index           =   1
         Left            =   2970
         TabIndex        =   13
         Top             =   765
         Width           =   1815
         _Version        =   65536
         _ExtentX        =   3201
         _ExtentY        =   7329
         _StockProps     =   15
         ForeColor       =   12582912
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
         Begin VB.TextBox txtVolume 
            Height          =   315
            Index           =   9
            Left            =   120
            MousePointer    =   3  'I-Beam
            TabIndex        =   23
            Text            =   "1039823"
            Top             =   540
            Width           =   1575
         End
         Begin VB.TextBox txtVolume 
            BackColor       =   &H00FFFFFF&
            ForeColor       =   &H00FF0000&
            Height          =   315
            Index           =   8
            Left            =   120
            MousePointer    =   3  'I-Beam
            TabIndex        =   22
            Text            =   "0"
            Top             =   3720
            Width           =   1575
         End
         Begin VB.TextBox txtVolume 
            BackColor       =   &H00FFFFFF&
            ForeColor       =   &H00FF0000&
            Height          =   315
            Index           =   7
            Left            =   120
            MousePointer    =   3  'I-Beam
            TabIndex        =   21
            Text            =   "0"
            Top             =   3420
            Width           =   1575
         End
         Begin VB.TextBox txtVolume 
            BackColor       =   &H00FFFFFF&
            ForeColor       =   &H00FF0000&
            Height          =   315
            Index           =   6
            Left            =   120
            MousePointer    =   3  'I-Beam
            TabIndex        =   20
            Text            =   "0"
            Top             =   3120
            Width           =   1575
         End
         Begin VB.TextBox txtVolume 
            BackColor       =   &H00FFFFFF&
            ForeColor       =   &H00FF0000&
            Height          =   315
            Index           =   5
            Left            =   120
            MousePointer    =   3  'I-Beam
            TabIndex        =   19
            Text            =   "0"
            Top             =   2820
            Width           =   1575
         End
         Begin VB.TextBox txtVolume 
            BackColor       =   &H00FFFFFF&
            ForeColor       =   &H00FF0000&
            Height          =   315
            Index           =   4
            Left            =   120
            MousePointer    =   3  'I-Beam
            TabIndex        =   18
            Text            =   "0"
            Top             =   2520
            Width           =   1575
         End
         Begin VB.TextBox txtVolume 
            BackColor       =   &H00FFFFFF&
            ForeColor       =   &H00FF0000&
            Height          =   315
            Index           =   3
            Left            =   120
            MousePointer    =   3  'I-Beam
            TabIndex        =   17
            Text            =   "0"
            Top             =   2220
            Width           =   1575
         End
         Begin VB.TextBox txtVolume 
            BackColor       =   &H00FFFFFF&
            ForeColor       =   &H00FF0000&
            Height          =   315
            Index           =   2
            Left            =   120
            MousePointer    =   3  'I-Beam
            TabIndex        =   16
            Text            =   "0"
            Top             =   1920
            Width           =   1575
         End
         Begin VB.TextBox txtVolume 
            BackColor       =   &H00FFFFFF&
            ForeColor       =   &H00FF0000&
            Height          =   315
            Index           =   1
            Left            =   120
            MousePointer    =   3  'I-Beam
            TabIndex        =   15
            Text            =   "0"
            Top             =   1620
            Width           =   1575
         End
         Begin VB.TextBox txtVolume 
            BackColor       =   &H00FFFFFF&
            ForeColor       =   &H00FF0000&
            Height          =   315
            Index           =   0
            Left            =   120
            MousePointer    =   3  'I-Beam
            TabIndex        =   14
            Text            =   "0"
            Top             =   1320
            Width           =   1575
         End
         Begin Threed.SSCheck chkUniform 
            Height          =   255
            Index           =   1
            Left            =   120
            TabIndex        =   24
            Top             =   180
            Width           =   1575
            _Version        =   65536
            _ExtentX        =   2778
            _ExtentY        =   450
            _StockProps     =   78
            Caption         =   "&Uniform"
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
         Begin VB.Label SumLabel 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "(TOTAL)"
            ForeColor       =   &H00000000&
            Height          =   195
            Index           =   0
            Left            =   120
            TabIndex        =   25
            Top             =   900
            Width           =   1575
         End
      End
      Begin Threed.SSPanel DisplayPanel 
         Height          =   4155
         Index           =   0
         Left            =   1110
         TabIndex        =   26
         Top             =   765
         Width           =   1815
         _Version        =   65536
         _ExtentX        =   3201
         _ExtentY        =   7329
         _StockProps     =   15
         ForeColor       =   12582912
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
         Begin VB.TextBox txtFeed 
            Height          =   315
            Index           =   0
            Left            =   120
            MousePointer    =   3  'I-Beam
            TabIndex        =   35
            Text            =   "1039823"
            Top             =   1320
            Width           =   1575
         End
         Begin VB.TextBox txtFeed 
            Height          =   315
            Index           =   1
            Left            =   120
            MousePointer    =   3  'I-Beam
            TabIndex        =   34
            Text            =   "1039823"
            Top             =   1620
            Width           =   1575
         End
         Begin VB.TextBox txtFeed 
            Height          =   315
            Index           =   2
            Left            =   120
            MousePointer    =   3  'I-Beam
            TabIndex        =   33
            Text            =   "1039823"
            Top             =   1920
            Width           =   1575
         End
         Begin VB.TextBox txtFeed 
            Height          =   315
            Index           =   3
            Left            =   120
            MousePointer    =   3  'I-Beam
            TabIndex        =   32
            Text            =   "1039823"
            Top             =   2220
            Width           =   1575
         End
         Begin VB.TextBox txtFeed 
            Height          =   315
            Index           =   4
            Left            =   120
            MousePointer    =   3  'I-Beam
            TabIndex        =   31
            Text            =   "1039823"
            Top             =   2520
            Width           =   1575
         End
         Begin VB.TextBox txtFeed 
            Height          =   315
            Index           =   5
            Left            =   120
            MousePointer    =   3  'I-Beam
            TabIndex        =   30
            Text            =   "1039823"
            Top             =   2820
            Width           =   1575
         End
         Begin VB.TextBox txtFeed 
            Height          =   315
            Index           =   6
            Left            =   120
            MousePointer    =   3  'I-Beam
            TabIndex        =   29
            Text            =   "1039823"
            Top             =   3120
            Width           =   1575
         End
         Begin VB.TextBox txtFeed 
            Height          =   315
            Index           =   7
            Left            =   120
            MousePointer    =   3  'I-Beam
            TabIndex        =   28
            Text            =   "1039823"
            Top             =   3420
            Width           =   1575
         End
         Begin VB.TextBox txtFeed 
            Height          =   315
            Index           =   8
            Left            =   120
            MousePointer    =   3  'I-Beam
            TabIndex        =   27
            Text            =   "1039823"
            Top             =   3720
            Width           =   1575
         End
         Begin Threed.SSCheck chkUniform 
            Height          =   255
            Index           =   0
            Left            =   120
            TabIndex        =   36
            Top             =   180
            Width           =   1575
            _Version        =   65536
            _ExtentX        =   2778
            _ExtentY        =   450
            _StockProps     =   78
            Caption         =   "&Uniform"
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
         Begin Threed.SSCheck chkStepFeed 
            Height          =   255
            Left            =   120
            TabIndex        =   37
            Top             =   600
            Width           =   1575
            _Version        =   65536
            _ExtentX        =   2778
            _ExtentY        =   450
            _StockProps     =   78
            Caption         =   "&Step Feed"
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
      End
      Begin Threed.SSPanel DisplayPanel 
         Height          =   4155
         Index           =   2
         Left            =   4830
         TabIndex        =   38
         Top             =   765
         Width           =   1815
         _Version        =   65536
         _ExtentX        =   3201
         _ExtentY        =   7329
         _StockProps     =   15
         ForeColor       =   12582912
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
         Begin VB.TextBox txtGasFlow 
            Height          =   315
            Index           =   9
            Left            =   120
            MousePointer    =   3  'I-Beam
            TabIndex        =   48
            Text            =   "1039823"
            Top             =   540
            Width           =   1575
         End
         Begin VB.TextBox txtGasFlow 
            BackColor       =   &H00FFFFFF&
            Height          =   315
            Index           =   8
            Left            =   120
            MousePointer    =   3  'I-Beam
            TabIndex        =   47
            Text            =   "0"
            Top             =   3720
            Width           =   1575
         End
         Begin VB.TextBox txtGasFlow 
            BackColor       =   &H00FFFFFF&
            Height          =   315
            Index           =   7
            Left            =   120
            MousePointer    =   3  'I-Beam
            TabIndex        =   46
            Text            =   "0"
            Top             =   3420
            Width           =   1575
         End
         Begin VB.TextBox txtGasFlow 
            BackColor       =   &H00FFFFFF&
            Height          =   315
            Index           =   6
            Left            =   120
            MousePointer    =   3  'I-Beam
            TabIndex        =   45
            Text            =   "0"
            Top             =   3120
            Width           =   1575
         End
         Begin VB.TextBox txtGasFlow 
            BackColor       =   &H00FFFFFF&
            Height          =   315
            Index           =   5
            Left            =   120
            MousePointer    =   3  'I-Beam
            TabIndex        =   44
            Text            =   "0"
            Top             =   2820
            Width           =   1575
         End
         Begin VB.TextBox txtGasFlow 
            BackColor       =   &H00FFFFFF&
            Height          =   315
            Index           =   4
            Left            =   120
            MousePointer    =   3  'I-Beam
            TabIndex        =   43
            Text            =   "0"
            Top             =   2520
            Width           =   1575
         End
         Begin VB.TextBox txtGasFlow 
            BackColor       =   &H00FFFFFF&
            Height          =   315
            Index           =   3
            Left            =   120
            MousePointer    =   3  'I-Beam
            TabIndex        =   42
            Text            =   "0"
            Top             =   2220
            Width           =   1575
         End
         Begin VB.TextBox txtGasFlow 
            BackColor       =   &H00FFFFFF&
            Height          =   315
            Index           =   2
            Left            =   120
            MousePointer    =   3  'I-Beam
            TabIndex        =   41
            Text            =   "0"
            Top             =   1920
            Width           =   1575
         End
         Begin VB.TextBox txtGasFlow 
            BackColor       =   &H00FFFFFF&
            Height          =   315
            Index           =   1
            Left            =   120
            MousePointer    =   3  'I-Beam
            TabIndex        =   40
            Text            =   "0"
            Top             =   1620
            Width           =   1575
         End
         Begin VB.TextBox txtGasFlow 
            BackColor       =   &H00FFFFFF&
            Height          =   315
            Index           =   0
            Left            =   120
            MousePointer    =   3  'I-Beam
            TabIndex        =   39
            Text            =   "0"
            Top             =   1320
            Width           =   1575
         End
         Begin Threed.SSCheck chkUniform 
            Height          =   255
            Index           =   2
            Left            =   120
            TabIndex        =   49
            Top             =   180
            Width           =   1575
            _Version        =   65536
            _ExtentX        =   2778
            _ExtentY        =   450
            _StockProps     =   78
            Caption         =   "&Uniform"
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
         Begin VB.Label SumLabel 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "(TOTAL)"
            ForeColor       =   &H00000000&
            Height          =   195
            Index           =   1
            Left            =   120
            TabIndex        =   50
            Top             =   900
            Width           =   1575
         End
      End
      Begin Threed.SSPanel DisplayPanel 
         Height          =   4155
         Index           =   3
         Left            =   6690
         TabIndex        =   51
         Top             =   765
         Width           =   1815
         _Version        =   65536
         _ExtentX        =   3201
         _ExtentY        =   7329
         _StockProps     =   15
         ForeColor       =   12582912
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
         Begin VB.TextBox txtBioMass 
            Height          =   315
            Index           =   9
            Left            =   120
            MousePointer    =   3  'I-Beam
            TabIndex        =   61
            Text            =   "1039823"
            Top             =   540
            Width           =   1575
         End
         Begin VB.TextBox txtBioMass 
            BackColor       =   &H00FFFFFF&
            Height          =   315
            Index           =   8
            Left            =   120
            MousePointer    =   3  'I-Beam
            TabIndex        =   60
            Text            =   "0"
            Top             =   3720
            Width           =   1575
         End
         Begin VB.TextBox txtBioMass 
            BackColor       =   &H00FFFFFF&
            Height          =   315
            Index           =   7
            Left            =   120
            MousePointer    =   3  'I-Beam
            TabIndex        =   59
            Text            =   "0"
            Top             =   3420
            Width           =   1575
         End
         Begin VB.TextBox txtBioMass 
            BackColor       =   &H00FFFFFF&
            Height          =   315
            Index           =   6
            Left            =   120
            MousePointer    =   3  'I-Beam
            TabIndex        =   58
            Text            =   "0"
            Top             =   3120
            Width           =   1575
         End
         Begin VB.TextBox txtBioMass 
            BackColor       =   &H00FFFFFF&
            Height          =   315
            Index           =   5
            Left            =   120
            MousePointer    =   3  'I-Beam
            TabIndex        =   57
            Text            =   "0"
            Top             =   2820
            Width           =   1575
         End
         Begin VB.TextBox txtBioMass 
            BackColor       =   &H00FFFFFF&
            Height          =   315
            Index           =   4
            Left            =   120
            MousePointer    =   3  'I-Beam
            TabIndex        =   56
            Text            =   "0"
            Top             =   2520
            Width           =   1575
         End
         Begin VB.TextBox txtBioMass 
            BackColor       =   &H00FFFFFF&
            Height          =   315
            Index           =   3
            Left            =   120
            MousePointer    =   3  'I-Beam
            TabIndex        =   55
            Text            =   "0"
            Top             =   2220
            Width           =   1575
         End
         Begin VB.TextBox txtBioMass 
            BackColor       =   &H00FFFFFF&
            Height          =   315
            Index           =   2
            Left            =   120
            MousePointer    =   3  'I-Beam
            TabIndex        =   54
            Text            =   "0"
            Top             =   1920
            Width           =   1575
         End
         Begin VB.TextBox txtBioMass 
            BackColor       =   &H00FFFFFF&
            Height          =   315
            Index           =   1
            Left            =   120
            MousePointer    =   3  'I-Beam
            TabIndex        =   53
            Text            =   "0"
            Top             =   1620
            Width           =   1575
         End
         Begin VB.TextBox txtBioMass 
            BackColor       =   &H00FFFFFF&
            Height          =   315
            Index           =   0
            Left            =   120
            MousePointer    =   3  'I-Beam
            TabIndex        =   52
            Text            =   "0"
            Top             =   1320
            Width           =   1575
         End
         Begin Threed.SSCheck chkUniform 
            Height          =   255
            Index           =   3
            Left            =   120
            TabIndex        =   62
            Top             =   180
            Width           =   1575
            _Version        =   65536
            _ExtentX        =   2778
            _ExtentY        =   450
            _StockProps     =   78
            Caption         =   "&Uniform"
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
         Begin VB.Label SumLabel 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "(AVERAGE)"
            ForeColor       =   &H00000000&
            Height          =   195
            Index           =   2
            Left            =   120
            TabIndex        =   63
            Top             =   900
            Width           =   1575
         End
      End
      Begin VB.Label TopLabel 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Biomass Concentration  (mg/L)"
         ForeColor       =   &H00800000&
         Height          =   465
         Index           =   3
         Left            =   6750
         TabIndex        =   76
         Top             =   270
         Width           =   1755
      End
      Begin VB.Label TopLabel 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Volume (L)"
         ForeColor       =   &H00800000&
         Height          =   255
         Index           =   1
         Left            =   2970
         TabIndex        =   75
         Top             =   465
         Width           =   1815
      End
      Begin VB.Label TopLabel 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Gas Flowrate (L/min)"
         ForeColor       =   &H00800000&
         Height          =   255
         Index           =   2
         Left            =   4830
         TabIndex        =   74
         Top             =   465
         Width           =   1815
      End
      Begin VB.Label TopLabel 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Feed Fraction"
         ForeColor       =   &H00800000&
         Height          =   255
         Index           =   0
         Left            =   1110
         TabIndex        =   73
         Top             =   465
         Width           =   1815
      End
      Begin VB.Label LeftLabel 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "CSTR 9"
         ForeColor       =   &H00800000&
         Height          =   255
         Index           =   8
         Left            =   270
         TabIndex        =   72
         Top             =   4530
         Width           =   735
      End
      Begin VB.Label LeftLabel 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "CSTR 8"
         ForeColor       =   &H00800000&
         Height          =   255
         Index           =   7
         Left            =   270
         TabIndex        =   71
         Top             =   4230
         Width           =   735
      End
      Begin VB.Label LeftLabel 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "CSTR 7"
         ForeColor       =   &H00800000&
         Height          =   255
         Index           =   6
         Left            =   270
         TabIndex        =   70
         Top             =   3930
         Width           =   735
      End
      Begin VB.Label LeftLabel 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "CSTR 6"
         ForeColor       =   &H00800000&
         Height          =   255
         Index           =   5
         Left            =   270
         TabIndex        =   69
         Top             =   3630
         Width           =   735
      End
      Begin VB.Label LeftLabel 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "CSTR 5"
         ForeColor       =   &H00800000&
         Height          =   255
         Index           =   4
         Left            =   270
         TabIndex        =   68
         Top             =   3330
         Width           =   735
      End
      Begin VB.Label LeftLabel 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "CSTR 4"
         ForeColor       =   &H00800000&
         Height          =   255
         Index           =   3
         Left            =   270
         TabIndex        =   67
         Top             =   3030
         Width           =   735
      End
      Begin VB.Label LeftLabel 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "CSTR 3"
         ForeColor       =   &H00800000&
         Height          =   255
         Index           =   2
         Left            =   270
         TabIndex        =   66
         Top             =   2730
         Width           =   735
      End
      Begin VB.Label LeftLabel 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "CSTR 2"
         ForeColor       =   &H00800000&
         Height          =   255
         Index           =   1
         Left            =   270
         TabIndex        =   65
         Top             =   2430
         Width           =   735
      End
      Begin VB.Label LeftLabel 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "CSTR 1"
         ForeColor       =   &H00800000&
         Height          =   255
         Index           =   0
         Left            =   270
         TabIndex        =   64
         Top             =   2130
         Width           =   735
      End
   End
   Begin Threed.SSFrame SSFrame3 
      Height          =   2955
      Left            =   10200
      TabIndex        =   4
      Top             =   4350
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
         Left            =   300
         TabIndex        =   77
         Top             =   390
         Width           =   2805
      End
   End
   Begin Threed.SSPanel SSPanel2 
      Align           =   2  'Align Bottom
      Height          =   405
      Left            =   0
      TabIndex        =   0
      Top             =   6495
      Width           =   9075
      _Version        =   65536
      _ExtentX        =   16007
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
   Begin Threed.SSCommand cmdCancelOK 
      Height          =   495
      Index           =   1
      Left            =   6300
      TabIndex        =   5
      TabStop         =   0   'False
      ToolTipText     =   "Click here to save the changes to this window"
      Top             =   30
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
      Left            =   7590
      TabIndex        =   6
      TabStop         =   0   'False
      ToolTipText     =   "Click here to abandon any changes on this window"
      Top             =   30
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
      Left            =   7590
      TabIndex        =   7
      TabStop         =   0   'False
      ToolTipText     =   "Click here for help"
      Top             =   570
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
   Begin Threed.SSCommand cmdCalcBiomassConc 
      Height          =   495
      Left            =   2400
      TabIndex        =   8
      TabStop         =   0   'False
      Top             =   720
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
   Begin Threed.SSFrame SSFrame7 
      Height          =   855
      Left            =   120
      TabIndex        =   9
      Top             =   420
      Width           =   2085
      _Version        =   65536
      _ExtentX        =   3678
      _ExtentY        =   1508
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
         Index           =   0
         Left            =   180
         TabIndex        =   10
         Text            =   "txtData(0)"
         Top             =   390
         Width           =   1425
      End
      Begin Spin.SpinButton spnData 
         Height          =   300
         Index           =   0
         Left            =   1620
         TabIndex        =   11
         Top             =   390
         Width           =   255
         _Version        =   65536
         _ExtentX        =   450
         _ExtentY        =   529
         _StockProps     =   73
      End
   End
   Begin VB.Label Label1 
      Caption         =   "This dialog allows for modeling up to 9 CSTR stages."
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
      Left            =   150
      TabIndex        =   3
      Top             =   60
      Width           =   5865
   End
End
Attribute VB_Name = "frmD5A_CSTR"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim USER_HIT_CANCEL As Boolean
Dim USER_HIT_OK As Boolean
Dim frmD5A_CSTR_Is_Dirty As Boolean

Dim Temp_Plant As TYPE_PlantDiagram

Public HALT_chkUniform As Boolean
Public HALT_chkStepFeed As Boolean





Const frmD5A_CSTR_declarations_end = True


Sub frmD5A_CSTR_Edit( _
    OUTPUT_Raise_Dirty_Flag As Boolean)
  Temp_Plant = frmD5_AerationBasin_Temp_Plant
  frmD5A_CSTR.Show 1
  If (USER_HIT_OK) Then
    OUTPUT_Raise_Dirty_Flag = True
    frmD5_AerationBasin_Temp_Plant = Temp_Plant
  Else
    OUTPUT_Raise_Dirty_Flag = False
  End If
End Sub


Sub frmD5A_CSTR_PopulateUnits()
Dim Frm As Form
Set Frm = frmD5A_CSTR
Dim i As Integer
  '
  ' MAIN DATA BLOCK.
  '
  With Temp_Plant.AerationBasin.CSTR
    Call unitsys_register(Frm, lblData(0), txtData(0), Nothing, "", _
        "", "", "0", "0", 100#, False)
    For i = 0 To 9
      If (i <> 9) Then
        Call unitsys_register(Frm, lblData(0), txtFeed(i), Nothing, "", _
            "", "", "", "", 100#, False)
      End If
      Call unitsys_register(Frm, lblData(0), txtVolume(i), Nothing, "", _
          "", "", "", "", 100#, False)
      Call unitsys_register(Frm, lblData(0), txtGasFlow(i), Nothing, "", _
          "", "", "", "", 100#, False)
      Call unitsys_register(Frm, lblData(0), txtBioMass(i), Nothing, "", _
          "", "", "", "", 100#, False)
    Next i
  End With
End Sub
Sub Store_Unit_Settings()
  '
  ' NO UNIT SETTINGS TO STORE!
  '
'Dim i As Integer
'  With Temp_Plant.AerationBasin.BioTreat
'    For i = 0 To 4
'      .UnitsOfDisplay(i) = unitsys_get_units(cboUnits(i))
'    Next i
'  End With
End Sub


Sub Local_DirtyStatus_Set(DirtyFlag As Boolean, NewSetting As Boolean)
  Call Global_DirtyStatus_Set(Me, DirtyFlag, NewSetting)
End Sub
Sub Local_GenericStatus_Set(NewString As String)
  Call Global_GenericStatus_Set(Me, NewString)
End Sub


Private Sub chkStepFeed_Click(Value As Integer)
Dim Ctl As Control
Set Ctl = chkStepFeed
Dim NewTag As Integer
Dim NewSetting As Integer
  If (HALT_chkUniform) Then Exit Sub
  NewTag = CInt(Value)
  If (CInt(Val(Ctl.Tag)) <> NewTag) Then
    NewSetting = NewTag
    With Temp_Plant.AerationBasin.CSTR
      .UseStepFeed = NewSetting
      If (.UseStepFeed = True) Then
        .UniformBioMass = False
      End If
    End With
    'RAISE DIRTY FLAG AND REFRESH WINDOW.
    Call Local_DirtyStatus_Set(frmD5A_CSTR_Is_Dirty, True)
    Call frmD5A_CSTR_Refresh(Temp_Plant)
  End If
End Sub
Private Sub chkUniform_Click(Index As Integer, Value As Integer)
Dim Ctl As Control
Set Ctl = chkUniform(Index)
Dim NewTag As Integer
Dim NewSetting As Integer
  If (HALT_chkUniform) Then Exit Sub
  NewTag = CInt(Value)
  If (CInt(Val(Ctl.Tag)) <> NewTag) Then
    NewSetting = NewTag
    With Temp_Plant.AerationBasin.CSTR
      Select Case Index
        Case 0:
          .UniformFeed = NewSetting
          If (NewSetting = False) Then .UniformBioMass = False
        Case 1:
          .UniformVolume = NewSetting
          If (NewSetting = False) Then .UniformBioMass = False
        Case 2:
          .UniformGasFlow = NewSetting
          If (NewSetting = False) Then .UniformBioMass = False
        Case 3:
          .UniformBioMass = NewSetting
      End Select
    End With
    'RAISE DIRTY FLAG AND REFRESH WINDOW.
    Call Local_DirtyStatus_Set(frmD5A_CSTR_Is_Dirty, True)
    Call frmD5A_CSTR_Refresh(Temp_Plant)
  End If
End Sub


Private Sub cmdCalcBiomassConc_Click()
Dim OUTPUT_Raise_Dirty_Flag As Boolean
Dim Temp_Total As Double
Dim i As Integer
  OUTPUT_Raise_Dirty_Flag = False
  frmD5A_CSTR_Temp_Plant = Temp_Plant
  Call frmD5B_Biomass.frmD5B_Biomass_Edit( _
      INPUT_UseWhichStructure_D5A, _
      OUTPUT_Raise_Dirty_Flag)
  If (OUTPUT_Raise_Dirty_Flag = True) Then
    '
    ' TRANSFER DATA.
    '
    Temp_Plant = frmD5A_CSTR_Temp_Plant
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
    Call Local_DirtyStatus_Set(frmD5A_CSTR_Is_Dirty, True)
    Call frmD5A_CSTR_Refresh(Temp_Plant)
  End If
End Sub
Private Sub cmdCancelOK_Click(Index As Integer)
Dim i As Integer
Dim CalcSuccess As Boolean
Dim ThisSum As Double
Dim LastRow As Integer
  Select Case Index
    Case 0:     'CANCEL.
      USER_HIT_CANCEL = True
      USER_HIT_OK = False
      Unload Me
      Exit Sub
    Case 1:     'OK.
      '
      ' VERIFY THAT SETTINGS ARE VALID.
      '
      LastRow = Temp_Plant.AerationBasin.CSTR.Count - 1
      With Temp_Plant.AerationBasin
        ThisSum = 0#
        For i = 0 To LastRow - 1
          ThisSum = ThisSum + .CSTR.Feed(i)
        Next i
      End With
      If (ThisSum > 1#) Then
        Call Show_Error("The sum of feed fractions for CSTRs 1 " & _
            "through " & _
            Trim$(Str$(Temp_Plant.AerationBasin.CSTR.Count - 1)) & _
            " is equal to " & Trim$(Str$(ThisSum)) & ".  This sum " & _
            "must be less than or equal to 1.000 to be valid.  " & _
            "You must either correct this problem and hit OK again, " & _
            "or abandon your changes to this window by hitting Cancel.")
        Exit Sub
      End If
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


Private Sub Form_Load()
  '
  ' MISC INITS.
  '
  Call CenterOnForm(Me, frmMain)
  Call Local_DirtyStatus_Set(frmD5A_CSTR_Is_Dirty, False)
  Call Local_GenericStatus_Set("")
  HALT_chkUniform = False
  HALT_chkStepFeed = False
  '
  ' POPULATE UNIT CONTROLS.
  '
  Call frmD5A_CSTR_PopulateUnits
  '
  ' REFRESH DISPLAY.
  '
  Call frmD5A_CSTR_Refresh(Temp_Plant)
End Sub
Private Sub Form_Unload(Cancel As Integer)
  Call unitsys_unregister_all_on_form(Me)
End Sub


Private Sub spnData_SpinDown(Index As Integer)
Dim Made_Dirty As Boolean
  Made_Dirty = False
  Select Case Index
    Case 0:
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
    Call Local_DirtyStatus_Set(frmD5A_CSTR_Is_Dirty, True)
    Call frmD5A_CSTR_Refresh(Temp_Plant)
  End If
End Sub
Private Sub spnData_SpinUp(Index As Integer)
Dim Made_Dirty As Boolean
  Made_Dirty = False
  Select Case Index
    Case 0:
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
    Call Local_DirtyStatus_Set(frmD5A_CSTR_Is_Dirty, True)
    Call frmD5A_CSTR_Refresh(Temp_Plant)
  End If
End Sub


Private Sub txtBioMass_GotFocus(Index As Integer)
Dim Ctl As Control
Set Ctl = txtBioMass(Index)
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
    Case -1:
    Case 9:
      StatusMessagePanel = _
          "Enter the average biomass concentration"
    Case Else:
      StatusMessagePanel = _
          "Enter the biomass concentration for CSTR #" & _
          Trim$(Str$(Index + 1))
  End Select
  Call Local_GenericStatus_Set(StatusMessagePanel)
End Sub
Private Sub txtBioMass_KeyPress(Index As Integer, KeyAscii As Integer)
  KeyAscii = Global_NumericKeyPress(KeyAscii)
End Sub
Private Sub txtBioMass_LostFocus(Index As Integer)
Dim NewValue_Okay As Integer
Dim NewValue As Double
Dim Ctl As Control
Set Ctl = txtBioMass(Index)
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
    Case -1:
    Case Else:
      Val_Low = 0#: Val_High = 1E+20
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
          Case -1:
          Case 9:
            .BioMass = NewValue
          Case Else:
            .CSTR.BioMass(Index) = NewValue
        End Select
      End With
      If (Raise_Dirty_Flag) Then
        'THROW DIRTY FLAG.
        Call Local_DirtyStatus_Set(frmD5A_CSTR_Is_Dirty, True)
      End If
      'REFRESH WINDOW.
      Call frmD5A_CSTR_Refresh(Temp_Plant)
    End If
  End If
End Sub


Private Sub txtData_GotFocus(Index As Integer)
Dim Ctl As Control
Set Ctl = txtData(Index)
Dim StatusMessagePanel As String
  Call unitsys_control_txtx_gotfocus(Ctl)
  Select Case Index
    '
    ' MAIN DATA BLOCK.
    '
    Case 0:
      StatusMessagePanel = "Enter the number of CSTRs"
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
    ' MAIN DATA BLOCK.
    Case 0: Val_Low = CDbl(1): Val_High = CDbl(AERATIONBASIN_MAX_CSTR)
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
      With Temp_Plant.AerationBasin.CSTR
        Select Case Index
          '
          ' MAIN DATA BLOCK.
          '
          Case 0: .Count = CInt(NewValue)
        End Select
      End With
      If (Raise_Dirty_Flag) Then
        'THROW DIRTY FLAG.
        Call Local_DirtyStatus_Set(frmD5A_CSTR_Is_Dirty, True)
      End If
      'REFRESH WINDOW.
      Call frmD5A_CSTR_Refresh(Temp_Plant)
    End If
  End If
End Sub


Private Sub txtFeed_GotFocus(Index As Integer)
Dim Ctl As Control
Set Ctl = txtFeed(Index)
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
    Case -1:
    Case Else:
      StatusMessagePanel = _
          "Enter the feed fraction for CSTR #" & _
          Trim$(Str$(Index + 1))
  End Select
  Call Local_GenericStatus_Set(StatusMessagePanel)
End Sub
Private Sub txtFeed_KeyPress(Index As Integer, KeyAscii As Integer)
  KeyAscii = Global_NumericKeyPress(KeyAscii)
End Sub
Private Sub txtFeed_LostFocus(Index As Integer)
Dim NewValue_Okay As Integer
Dim NewValue As Double
Dim Ctl As Control
Set Ctl = txtFeed(Index)
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
    Case -1:
    Case Else:
      Val_Low = 0#: Val_High = 1#
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
      With Temp_Plant.AerationBasin.CSTR
        Select Case Index
          '
          ' MAIN DATA BLOCK.
          '
          Case -1:
          Case Else:
            .Feed(Index) = NewValue
        End Select
      End With
      If (Raise_Dirty_Flag) Then
        'THROW DIRTY FLAG.
        Call Local_DirtyStatus_Set(frmD5A_CSTR_Is_Dirty, True)
      End If
      'REFRESH WINDOW.
      Call frmD5A_CSTR_Refresh(Temp_Plant)
    End If
  End If
End Sub


Private Sub txtGasFlow_GotFocus(Index As Integer)
Dim Ctl As Control
Set Ctl = txtGasFlow(Index)
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
    Case -1:
    Case 9:
      StatusMessagePanel = _
          "Enter the total gas flow rate"
    Case Else:
      StatusMessagePanel = _
          "Enter the gas flow rate for CSTR #" & _
          Trim$(Str$(Index + 1))
  End Select
  Call Local_GenericStatus_Set(StatusMessagePanel)
End Sub
Private Sub txtGasFlow_KeyPress(Index As Integer, KeyAscii As Integer)
  KeyAscii = Global_NumericKeyPress(KeyAscii)
End Sub
Private Sub txtGasFlow_LostFocus(Index As Integer)
Dim NewValue_Okay As Integer
Dim NewValue As Double
Dim Ctl As Control
Set Ctl = txtGasFlow(Index)
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
    Case -1:
    Case Else:
      Val_Low = 0#: Val_High = 1E+20
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
          Case -1:
          Case 9:
            .GasFlow = NewValue
          Case Else:
            .CSTR.GasFlow(Index) = NewValue
        End Select
      End With
      If (Raise_Dirty_Flag) Then
        'THROW DIRTY FLAG.
        Call Local_DirtyStatus_Set(frmD5A_CSTR_Is_Dirty, True)
      End If
      'REFRESH WINDOW.
      Call frmD5A_CSTR_Refresh(Temp_Plant)
    End If
  End If
End Sub


Private Sub txtVolume_GotFocus(Index As Integer)
Dim Ctl As Control
Set Ctl = txtVolume(Index)
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
    Case -1:
    Case 9:
      StatusMessagePanel = _
          "Enter the total volume"
    Case Else:
      StatusMessagePanel = _
          "Enter the volume for CSTR #" & _
          Trim$(Str$(Index + 1))
  End Select
  Call Local_GenericStatus_Set(StatusMessagePanel)
End Sub
Private Sub txtVolume_KeyPress(Index As Integer, KeyAscii As Integer)
  KeyAscii = Global_NumericKeyPress(KeyAscii)
End Sub
Private Sub txtVolume_LostFocus(Index As Integer)
Dim NewValue_Okay As Integer
Dim NewValue As Double
Dim Ctl As Control
Set Ctl = txtVolume(Index)
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
    Case -1:
    Case Else:
      Val_Low = 0#: Val_High = 1E+20
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
          Case -1:
          Case 9:
            .Volume = NewValue
          Case Else:
            .CSTR.Volume(Index) = NewValue
        End Select
      End With
      If (Raise_Dirty_Flag) Then
        'THROW DIRTY FLAG.
        Call Local_DirtyStatus_Set(frmD5A_CSTR_Is_Dirty, True)
      End If
      'REFRESH WINDOW.
      Call frmD5A_CSTR_Refresh(Temp_Plant)
    End If
  End If
End Sub


