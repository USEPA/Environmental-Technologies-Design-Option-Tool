VERSION 5.00
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "THREED32.OCX"
Begin VB.Form frmInputParamsPSDMInRoom 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Parameters for PSDMR Model"
   ClientHeight    =   7545
   ClientLeft      =   1215
   ClientTop       =   1320
   ClientWidth     =   8175
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7545
   ScaleWidth      =   8175
   ShowInTaskbar   =   0   'False
   Begin Threed.SSFrame SSFrame1 
      Height          =   1335
      Left            =   90
      TabIndex        =   2
      Top             =   90
      Width           =   7965
      _Version        =   65536
      _ExtentX        =   14049
      _ExtentY        =   2355
      _StockProps     =   14
      Caption         =   "Main Set of Room Properties:"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Begin VB.ComboBox cboUnits 
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
         Left            =   5250
         Style           =   2  'Dropdown List
         TabIndex        =   5
         TabStop         =   0   'False
         Top             =   240
         Width           =   1275
      End
      Begin VB.TextBox txtData 
         Alignment       =   2  'Center
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
         Height          =   288
         Index           =   0
         Left            =   3990
         TabIndex        =   0
         Text            =   "txtData(0)"
         Top             =   270
         Width           =   1212
      End
      Begin VB.ComboBox cboUnits 
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
         Left            =   5250
         Style           =   2  'Dropdown List
         TabIndex        =   4
         TabStop         =   0   'False
         Top             =   600
         Width           =   1275
      End
      Begin VB.TextBox txtData 
         Alignment       =   2  'Center
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
         Height          =   288
         Index           =   1
         Left            =   3990
         TabIndex        =   1
         Text            =   "txtData(1)"
         Top             =   630
         Width           =   1212
      End
      Begin VB.Label lblData 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Volume of Room"
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
         Left            =   -30
         TabIndex        =   10
         Top             =   300
         Width           =   3885
      End
      Begin VB.Label lblData 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Volumetric Flow Rate of Air Through Room"
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
         Left            =   -30
         TabIndex        =   9
         Top             =   660
         Width           =   3885
      End
      Begin VB.Label lblDesc 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Air Change Rate"
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
         Left            =   -30
         TabIndex        =   8
         Top             =   1020
         Width           =   3885
      End
      Begin VB.Label lblAirRate 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         Caption         =   "lblAirRate"
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
         Left            =   3990
         TabIndex        =   7
         Top             =   1020
         Width           =   1215
      End
      Begin VB.Label lblAirRateUnits 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         Caption         =   "hour^(-1)"
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
         Left            =   5250
         TabIndex        =   6
         Top             =   1020
         Width           =   1275
      End
   End
   Begin Threed.SSFrame SSFrame2 
      Height          =   4995
      Left            =   90
      TabIndex        =   3
      Top             =   1500
      Width           =   7965
      _Version        =   65536
      _ExtentX        =   14049
      _ExtentY        =   8811
      _StockProps     =   14
      Caption         =   "Contaminant Properties:"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Begin VB.ComboBox cboChemical 
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
         Left            =   3300
         Style           =   2  'Dropdown List
         TabIndex        =   11
         TabStop         =   0   'False
         Top             =   210
         Width           =   3345
      End
      Begin Threed.SSPanel sspContaminantProps 
         Height          =   4335
         Left            =   150
         TabIndex        =   22
         Top             =   570
         Width           =   7665
         _Version        =   65536
         _ExtentX        =   13520
         _ExtentY        =   7646
         _StockProps     =   15
         Caption         =   "Properties of {ContaminantName}:"
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
         Alignment       =   6
         Begin VB.ComboBox cboUnits 
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
            Left            =   5190
            Style           =   2  'Dropdown List
            TabIndex        =   40
            TabStop         =   0   'False
            Top             =   2760
            Width           =   1275
         End
         Begin VB.TextBox txtData 
            Alignment       =   2  'Center
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
            Height          =   288
            Index           =   4
            Left            =   3930
            TabIndex        =   39
            Text            =   "txtData(4)"
            Top             =   2790
            Width           =   1212
         End
         Begin VB.TextBox txtData 
            Alignment       =   2  'Center
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
            Height          =   288
            Index           =   5
            Left            =   3930
            TabIndex        =   38
            Text            =   "txtData(5)"
            Top             =   3180
            Width           =   1212
         End
         Begin VB.ComboBox cboUnits 
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
            Left            =   5190
            Style           =   2  'Dropdown List
            TabIndex        =   37
            TabStop         =   0   'False
            Top             =   3150
            Width           =   1275
         End
         Begin VB.TextBox txtData 
            Alignment       =   2  'Center
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
            Height          =   288
            Index           =   6
            Left            =   3930
            TabIndex        =   36
            Text            =   "txtData(6)"
            Top             =   3570
            Width           =   1212
         End
         Begin VB.ComboBox cbo_RXN_PRODUCT 
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
            Left            =   3120
            Style           =   2  'Dropdown List
            TabIndex        =   35
            TabStop         =   0   'False
            Top             =   3960
            Width           =   3345
         End
         Begin Threed.SSFrame SSFrame5 
            Height          =   675
            Left            =   240
            TabIndex        =   23
            Top             =   930
            Width           =   7365
            _Version        =   65536
            _ExtentX        =   12991
            _ExtentY        =   1191
            _StockProps     =   14
            Caption         =   "Mass Emission Rate Within Room"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Alignment       =   2
            Begin VB.TextBox txtData 
               Alignment       =   2  'Center
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
               Height          =   288
               Index           =   3
               Left            =   1500
               TabIndex        =   25
               Text            =   "txtData(3)"
               Top             =   300
               Width           =   1212
            End
            Begin VB.ComboBox cboUnits 
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
               Left            =   2760
               Style           =   2  'Dropdown List
               TabIndex        =   24
               TabStop         =   0   'False
               Top             =   270
               Width           =   1275
            End
            Begin Threed.SSOption optTimeVarEmit 
               Height          =   255
               Index           =   0
               Left            =   120
               TabIndex        =   26
               Top             =   300
               Width           =   1335
               _Version        =   65536
               _ExtentX        =   2355
               _ExtentY        =   450
               _StockProps     =   78
               Caption         =   "Constant    -"
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
            Begin Threed.SSOption optTimeVarEmit 
               Height          =   255
               Index           =   1
               Left            =   4320
               TabIndex        =   27
               Top             =   300
               Width           =   1815
               _Version        =   65536
               _ExtentX        =   3201
               _ExtentY        =   450
               _StockProps     =   78
               Caption         =   "Time-Variable    -"
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
            Begin Threed.SSCommand cmdTimeVarEmit 
               Height          =   345
               Left            =   6210
               TabIndex        =   28
               TabStop         =   0   'False
               Top             =   270
               Width           =   1005
               _Version        =   65536
               _ExtentX        =   1773
               _ExtentY        =   609
               _StockProps     =   78
               Caption         =   "Set ..."
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
         Begin Threed.SSFrame SSFrame4 
            Height          =   675
            Left            =   240
            TabIndex        =   29
            Top             =   240
            Width           =   7365
            _Version        =   65536
            _ExtentX        =   12991
            _ExtentY        =   1191
            _StockProps     =   14
            Caption         =   "Concentration in Influent Stream to Room"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Alignment       =   2
            Begin VB.ComboBox cboUnits 
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
               Left            =   2760
               Style           =   2  'Dropdown List
               TabIndex        =   31
               TabStop         =   0   'False
               Top             =   270
               Width           =   1275
            End
            Begin VB.TextBox txtData 
               Alignment       =   2  'Center
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
               Height          =   288
               Index           =   2
               Left            =   1500
               TabIndex        =   30
               Text            =   "txtData(2)"
               Top             =   300
               Width           =   1212
            End
            Begin Threed.SSOption optTimeVarConc 
               Height          =   255
               Index           =   0
               Left            =   120
               TabIndex        =   32
               Top             =   300
               Width           =   1335
               _Version        =   65536
               _ExtentX        =   2355
               _ExtentY        =   450
               _StockProps     =   78
               Caption         =   "Constant    -"
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
            Begin Threed.SSOption optTimeVarConc 
               Height          =   255
               Index           =   1
               Left            =   4320
               TabIndex        =   33
               Top             =   300
               Width           =   1815
               _Version        =   65536
               _ExtentX        =   3201
               _ExtentY        =   450
               _StockProps     =   78
               Caption         =   "Time-Variable    -"
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
            Begin Threed.SSCommand cmdTimeVarConc 
               Height          =   345
               Left            =   6210
               TabIndex        =   34
               TabStop         =   0   'False
               Top             =   270
               Width           =   1005
               _Version        =   65536
               _ExtentX        =   1773
               _ExtentY        =   609
               _StockProps     =   78
               Caption         =   "Set ..."
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
         Begin Threed.SSFrame SSFrame6 
            Height          =   675
            Left            =   240
            TabIndex        =   49
            Top             =   1620
            Width           =   7365
            _Version        =   65536
            _ExtentX        =   12991
            _ExtentY        =   1191
            _StockProps     =   14
            Caption         =   "Freundlich K of Contaminant"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Alignment       =   2
            Begin VB.TextBox txtData 
               Alignment       =   2  'Center
               Appearance      =   0  'Flat
               BackColor       =   &H8000000F&
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   288
               Index           =   7
               Left            =   1500
               Locked          =   -1  'True
               TabIndex        =   51
               Text            =   "txtData(7)"
               Top             =   300
               Width           =   1212
            End
            Begin VB.ComboBox cboUnits 
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
               Left            =   2760
               Style           =   2  'Dropdown List
               TabIndex        =   50
               TabStop         =   0   'False
               Top             =   270
               Width           =   1275
            End
            Begin Threed.SSOption optTimeVarK 
               Height          =   255
               Index           =   0
               Left            =   120
               TabIndex        =   52
               Top             =   300
               Width           =   1335
               _Version        =   65536
               _ExtentX        =   2355
               _ExtentY        =   450
               _StockProps     =   78
               Caption         =   "Constant    -"
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
            Begin Threed.SSOption optTimeVarK 
               Height          =   255
               Index           =   1
               Left            =   4320
               TabIndex        =   53
               Top             =   300
               Width           =   1815
               _Version        =   65536
               _ExtentX        =   3201
               _ExtentY        =   450
               _StockProps     =   78
               Caption         =   "Time-Variable    -"
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
            Begin Threed.SSCommand cmdTimeVarK 
               Height          =   345
               Left            =   6210
               TabIndex        =   54
               TabStop         =   0   'False
               Top             =   270
               Width           =   1005
               _Version        =   65536
               _ExtentX        =   1773
               _ExtentY        =   609
               _StockProps     =   78
               Caption         =   "Set ..."
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
         Begin VB.Label lblData 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "Concentration in Room at Time = Zero"
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
            TabIndex        =   47
            Top             =   2820
            Width           =   3765
         End
         Begin VB.Label lblDesc 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
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
            ForeColor       =   &H80000008&
            Height          =   255
            Index           =   5
            Left            =   -30
            TabIndex        =   46
            Top             =   2460
            Width           =   3885
         End
         Begin VB.Label lblSSValue 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00C0C0C0&
            BackStyle       =   0  'Transparent
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
            ForeColor       =   &H80000008&
            Height          =   255
            Left            =   3930
            TabIndex        =   45
            Top             =   2460
            Width           =   1215
         End
         Begin VB.Label lblSSValueUnits 
            Appearance      =   0  'Flat
            BackColor       =   &H00C0C0C0&
            BackStyle       =   0  'Transparent
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
            ForeColor       =   &H80000008&
            Height          =   255
            Left            =   5190
            TabIndex        =   44
            Top             =   2460
            Width           =   1275
         End
         Begin VB.Label lblData 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "First-Order Destruction Rate"
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
            TabIndex        =   43
            Top             =   3210
            Width           =   3765
         End
         Begin VB.Label lblData 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "# Moles Of Product Per Mole Reactant"
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
            TabIndex        =   42
            Top             =   3600
            Width           =   3765
         End
         Begin VB.Label lbl_cbo_RXN_PRODUCT 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "Product For This Reactant"
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
            Left            =   540
            TabIndex        =   41
            Top             =   4020
            Width           =   2505
         End
      End
      Begin VB.Label lblDesc_cboChemical 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Select Contaminant Name:"
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
         Left            =   720
         TabIndex        =   12
         Top             =   270
         Width           =   2505
      End
   End
   Begin Threed.SSCommand cmdCancelOK 
      Height          =   495
      Index           =   1
      Left            =   3750
      TabIndex        =   13
      TabStop         =   0   'False
      ToolTipText     =   "Click here to save the changes you have made to the parameters on this window"
      Top             =   6570
      Width           =   2175
      _Version        =   65536
      _ExtentX        =   2646
      _ExtentY        =   1323
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
      Left            =   5910
      TabIndex        =   14
      TabStop         =   0   'False
      ToolTipText     =   "Click here to abandon any changes you have made to the parameters on this window"
      Top             =   6570
      Width           =   2175
      _Version        =   65536
      _ExtentX        =   2646
      _ExtentY        =   1323
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
   Begin Threed.SSPanel SSPanel2 
      Align           =   2  'Align Bottom
      Height          =   405
      Left            =   0
      TabIndex        =   15
      Top             =   7140
      Width           =   8175
      _Version        =   65536
      _ExtentX        =   14420
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
         TabIndex        =   16
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
         TabIndex        =   17
         Top             =   60
         Width           =   4720
         _Version        =   65536
         _ExtentX        =   8326
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
   Begin Threed.SSFrame SSFrame3 
      Height          =   2265
      Left            =   390
      TabIndex        =   18
      Top             =   5940
      Visible         =   0   'False
      Width           =   2355
      _Version        =   65536
      _ExtentX        =   4154
      _ExtentY        =   3995
      _StockProps     =   14
      Caption         =   "Invisible -- Do not delete"
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
         Left            =   300
         Style           =   2  'Dropdown List
         TabIndex        =   19
         TabStop         =   0   'False
         Top             =   390
         Width           =   1275
      End
      Begin Threed.SSFrame ssframe_ContaminantProps 
         Height          =   495
         Left            =   210
         TabIndex        =   48
         Top             =   1380
         Width           =   3135
         _Version        =   65536
         _ExtentX        =   5530
         _ExtentY        =   873
         _StockProps     =   14
         Caption         =   "Properties of {ContaminantName}:"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Alignment       =   2
         ShadowStyle     =   1
      End
      Begin VB.Label lblData 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Mass Emission Rate Within Room"
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
         Left            =   30
         TabIndex        =   21
         Top             =   1020
         Width           =   3765
      End
      Begin VB.Label lblData 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Concentration in Influent Stream To Room"
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
         Left            =   150
         TabIndex        =   20
         Top             =   810
         Width           =   3765
      End
   End
End
Attribute VB_Name = "frmInputParamsPSDMInRoom"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim TempData As RoomParam_Type
Dim NOW_CONTAMINANT As Integer

Dim USER_HIT_CANCEL As Boolean
Dim USER_HIT_OK As Boolean
Dim Temp_RP As RoomParam_Type

Dim frmInputParamsPSDMInRoom_Is_Dirty As Boolean
Dim HALT_cboChemical As Boolean
Public HALT_cbo_RXN_PRODUCT As Boolean
Public HALT_ALL_CONTROLS As Boolean

Const IN_cmdTimeVar_WhichButton___CO = 1
Const IN_cmdTimeVar_WhichButton___WA = 2
Const IN_cmdTimeVar_WhichButton___K = 3




Const frmInputParamsPSDMInRoom_declarations_end = True


Sub frmInputParamsPSDMInRoom_Edit( _
    OUTPUT_Raise_Dirty_Flag As Boolean)
  Temp_RP = RoomParams
  If (Temp_RP.COUNT_CONTAMINANT <> Number_Component) Then
    Temp_RP.COUNT_CONTAMINANT = Number_Component
  End If
  frmInputParamsPSDMInRoom.Show 1
  If (USER_HIT_OK) Then
    OUTPUT_Raise_Dirty_Flag = True
    RoomParams = Temp_RP
  Else
    OUTPUT_Raise_Dirty_Flag = False
  End If
End Sub
Sub LOCAL___Reset_DemoVersionDisablings()
  If (IsThisADemo() = True) Then
    cmdCancelOK(1).Enabled = False
  End If
End Sub


Sub frmInputParamsPSDMInRoom_PopulateUnits()
Dim Frm As Form
Set Frm = frmInputParamsPSDMInRoom
  'MAIN BLOCK OF UNITS.
  Call unitsys_register(Frm, lblData(0), _
      txtData(0), cboUnits(0), "volume", _
      Temp_RP.ROOM_VOL_Units, "m³", "", "", 100#, True)
  Call unitsys_register(Frm, lblData(1), _
      txtData(1), cboUnits(1), "flow_volumetric", _
      Temp_RP.ROOM_FLOWRATE_Units, "m³/s", "", "", 100#, True)
  Call unitsys_register(Frm, lblData(2), _
      txtData(2), cboUnits(2), "concentration", _
      Temp_RP.ROOM_C0_Units, "mg/L", "", "", 100#, True)
  Call unitsys_register(Frm, lblData(3), _
      txtData(3), cboUnits(3), "mass_emission_rate", _
      Temp_RP.ROOM_EMIT_Units, "µg/s", "", "", 100#, True)
  Call unitsys_register(Frm, lblData(4), _
      txtData(4), cboUnits(4), "concentration", _
      Temp_RP.INITIAL_ROOM_CONC_Units, "mg/L", "", "", 100#, True)
  Call unitsys_register(Frm, lblData(5), _
      txtData(5), cboUnits(5), "inverse_time", _
      "1/s", "1/s", "", "", 100#, True)
  Call unitsys_register(Frm, lblData(6), _
      txtData(6), Nothing, "", _
      "", "", "", "", 100#, False)
  Call unitsys_register(Frm, lblData(7), _
      txtData(7), cboUnits(7), "freundlich_k", _
      "(mg/g)*(L/mg)^(1/n)", "(mg/g)*(L/mg)^(1/n)", "", "", 100#, True)
End Sub
Sub Store_Unit_Settings()
  Temp_RP.ROOM_VOL_Units = unitsys_get_units(cboUnits(0))
  Temp_RP.ROOM_FLOWRATE_Units = unitsys_get_units(cboUnits(1))
  Temp_RP.ROOM_C0_Units = unitsys_get_units(cboUnits(2))
  Temp_RP.ROOM_EMIT_Units = unitsys_get_units(cboUnits(3))
  Temp_RP.INITIAL_ROOM_CONC_Units = unitsys_get_units(cboUnits(4))
  'Temp_RP.u_ROOM_KINI = unitsys_get_units(cboUnits(7))
End Sub


Sub Do_Refresh()
  Call frmInputParamsPSDMInRoom_Refresh(Temp_RP, NOW_CONTAMINANT)
End Sub


Private Sub populate_cboChemical()
Dim i As Integer
Dim Ctl As Control
Set Ctl = cboChemical
  HALT_cboChemical = True
  Ctl.Clear
  If (frmMain.cboSelectCompo.ListCount > 0) Then
    For i = 1 To frmMain.cboSelectCompo.ListCount
      Ctl.AddItem Trim$(frmMain.cboSelectCompo.List(i - 1))
    Next
    Ctl.ListIndex = 0
  End If
  HALT_cboChemical = False
End Sub
Private Sub populate_cbo_RXN_PRODUCT()
Dim i As Integer
Dim Ctl As Control
Set Ctl = cbo_RXN_PRODUCT
  HALT_cbo_RXN_PRODUCT = True
  Ctl.Clear
  If (frmMain.cboSelectCompo.ListCount > 0) Then
    For i = 1 To frmMain.cboSelectCompo.ListCount
      Ctl.AddItem Trim$(frmMain.cboSelectCompo.List(i - 1))
      Ctl.ItemData(Ctl.NewIndex) = i
    Next
    Ctl.ListIndex = 0
  End If
  HALT_cbo_RXN_PRODUCT = False
End Sub


Sub frmInputParamsPSDMInRoom_GenericStatus_Set(fn_Text As String)
  Me.sspanel_Status = fn_Text
End Sub
Sub frmInputParamsPSDMInRoom_DirtyStatus_Set(newVal As Boolean)
Dim Frm As Form
Set Frm = frmInputParamsPSDMInRoom
  If (newVal) Then
    Frm.sspanel_Dirty = "Data Changed"
    Frm.sspanel_Dirty.ForeColor = QBColor(12)
  Else
    Frm.sspanel_Dirty = "Unchanged"
    Frm.sspanel_Dirty.ForeColor = QBColor(0)
  End If
End Sub
Sub frmInputParamsPSDMInRoom_DirtyStatus_Set_Current()
  Call frmInputParamsPSDMInRoom_DirtyStatus_Set( _
      frmInputParamsPSDMInRoom_Is_Dirty)
End Sub
Sub frmInputParamsPSDMInRoom_DirtyStatus_Throw()
  frmInputParamsPSDMInRoom_Is_Dirty = True
  Call frmInputParamsPSDMInRoom_DirtyStatus_Set_Current
End Sub
Sub frmInputParamsPSDMInRoom_DirtyStatus_Clear()
  frmInputParamsPSDMInRoom_Is_Dirty = False
  Call frmInputParamsPSDMInRoom_DirtyStatus_Set_Current
End Sub


Private Sub cbo_RXN_PRODUCT_Click()
Dim Ctl As Control
Set Ctl = cbo_RXN_PRODUCT
  If (HALT_cbo_RXN_PRODUCT) Then Exit Sub
  Temp_RP.RXN_PRODUCT(NOW_CONTAMINANT) = Ctl.ItemData(Ctl.ListIndex)
  '
  ' THROW DIRTY FLAG AND REFRESH.
  Call frmInputParamsPSDMInRoom_DirtyStatus_Throw
  ''''Call RoomParam_Recalculate(Temp_RP)
  Call Do_Refresh
End Sub
Private Sub cboChemical_Click()
  If (HALT_cboChemical) Then Exit Sub
  NOW_CONTAMINANT = cboChemical.ListIndex + 1
  Call Do_Refresh
End Sub


Private Sub cboUnits_Click(Index As Integer)
Dim Ctl As Control
Set Ctl = cboUnits(Index)
  Call unitsys_control_cbox_click(Ctl)
End Sub
Private Sub cboUnits_KeyPress(Index As Integer, KeyAscii As Integer)
  KeyAscii = Global_TextKeyPress(KeyAscii)
End Sub


Private Sub cmdCancelOK_Click(Index As Integer)
Dim i As Integer
  Select Case Index
    Case 0:     'CANCEL.
      'If (frmCompoProp_Query_Unload() = False) Then
      '  'THE CANCEL WAS CANCELLED.
      '  Exit Sub
      'End If
      USER_HIT_CANCEL = True
      USER_HIT_OK = False
      Unload Me
      Exit Sub
    Case 1:     'OK.
      'STORE ALL UNIT SETTINGS.
      Call Store_Unit_Settings
      'EXIT OUT OF HERE.
      USER_HIT_CANCEL = False
      USER_HIT_OK = True
      Unload Me
      Exit Sub
  End Select
End Sub

Sub Do___cmdTimeVar___ButtonClick( _
    IN_cmdTimeVar_WhichButton As Integer)
Dim FormCaption As String
Dim UnitType(1 To 2) As String
Dim BaseUnits(1 To 2) As String
Dim CurrentUnits(1 To 2) As String
Dim lblUnitType(1 To 2) As String
Dim DataRowCount As Integer
Dim MaxRows As Integer
Dim ColumnCount As Integer
Dim ColumnNames() As String
Dim foStoreTo As Control
Dim USER_HIT_CANCEL As Boolean
  '
  ' EXTRA STEP:
  ' TRANSFER dbl_ROOM_COINI(), dbl_ROOM_TCOINI()
  ' DATA INTO frmMain.foVarConc.
  '
Dim i As Integer
Dim J As Integer
Dim Ctl As Control
  Set Ctl = frmMain.foVarConc
  Ctl.Sheet = 1
  Select Case IN_cmdTimeVar_WhichButton
    Case IN_cmdTimeVar_WhichButton___CO:
      If (Temp_RP.int_ROOM_NCOINI(NOW_CONTAMINANT) = 0) Then
        Ctl.MaxRow = 1
      Else
        Ctl.MaxRow = Temp_RP.int_ROOM_NCOINI(NOW_CONTAMINANT)
      End If
      For i = 1 To Temp_RP.int_ROOM_NCOINI(NOW_CONTAMINANT)
        Ctl.EntryRC(i, 1) = Temp_RP.dbl_ROOM_TCOINI(NOW_CONTAMINANT, i)
        Ctl.EntryRC(i, 2) = Temp_RP.dbl_ROOM_COINI(NOW_CONTAMINANT, i)
      Next i
    Case IN_cmdTimeVar_WhichButton___WA:
      If (Temp_RP.int_ROOM_NEMITINI(NOW_CONTAMINANT) = 0) Then
        Ctl.MaxRow = 1
      Else
        Ctl.MaxRow = Temp_RP.int_ROOM_NEMITINI(NOW_CONTAMINANT)
      End If
      For i = 1 To Temp_RP.int_ROOM_NEMITINI(NOW_CONTAMINANT)
        Ctl.EntryRC(i, 1) = Temp_RP.dbl_ROOM_TEMITINI(NOW_CONTAMINANT, i)
        Ctl.EntryRC(i, 2) = Temp_RP.dbl_ROOM_EMITINI(NOW_CONTAMINANT, i)
      Next i
    Case IN_cmdTimeVar_WhichButton___K:
      If (Temp_RP.int_ROOM_NKINI(NOW_CONTAMINANT) = 0) Then
        Ctl.MaxRow = 1
      Else
        Ctl.MaxRow = Temp_RP.int_ROOM_NKINI(NOW_CONTAMINANT)
      End If
      For i = 1 To Temp_RP.int_ROOM_NKINI(NOW_CONTAMINANT)
        Ctl.EntryRC(i, 1) = Temp_RP.dbl_ROOM_TKINI(NOW_CONTAMINANT, i)
        Ctl.EntryRC(i, 2) = Temp_RP.dbl_ROOM_KINI(NOW_CONTAMINANT, i)
      Next i
  End Select
  '
  ' NOW, PROCEED WITH NORMAL CODE.
  '
  ColumnCount = 2
  ReDim ColumnNames(1 To 2)
  ColumnNames(1) = "Time"
  Select Case IN_cmdTimeVar_WhichButton
    Case IN_cmdTimeVar_WhichButton___CO:
      FormCaption = _
          Me.cboChemical.List(Me.cboChemical.ListIndex) & _
          " Influent Concentrations To Room (Time-Variable)"
      UnitType(1) = "time"
      UnitType(2) = "concentration"
      BaseUnits(1) = "min"
      BaseUnits(2) = "µg/L"
      CurrentUnits(1) = Temp_RP.u_ROOM_TCOINI
      CurrentUnits(2) = Temp_RP.u_ROOM_COINI
      lblUnitType(1) = "Time Units:"
      lblUnitType(2) = "Concentration Units:"
      DataRowCount = Temp_RP.int_ROOM_NCOINI(NOW_CONTAMINANT)
      MaxRows = Max_int_ROOM_NCOINI
      ColumnNames(2) = "Concentration"
    Case IN_cmdTimeVar_WhichButton___WA:
      FormCaption = _
          Me.cboChemical.List(Me.cboChemical.ListIndex) & _
          " Mass Emission Rates (Time-Variable)"
      UnitType(1) = "time"
      UnitType(2) = "mass_emission_rate"
      BaseUnits(1) = "min"
      BaseUnits(2) = "µg/s"
      CurrentUnits(1) = Temp_RP.u_ROOM_TEMITINI
      CurrentUnits(2) = Temp_RP.u_ROOM_EMITINI
      lblUnitType(1) = "Time Units:"
      lblUnitType(2) = "Emission Rate Units:"
      DataRowCount = Temp_RP.int_ROOM_NEMITINI(NOW_CONTAMINANT)
      MaxRows = Max_int_ROOM_NEMITINI
      ColumnNames(2) = "Mass Emission Rate"
    Case IN_cmdTimeVar_WhichButton___K:
      FormCaption = _
          Me.cboChemical.List(Me.cboChemical.ListIndex) & _
          " Freundlich K (Time-Variable)"
      UnitType(1) = "time"
      UnitType(2) = "freundlich_k"
      BaseUnits(1) = "min"
      BaseUnits(2) = "(mg/g)*(L/mg)^(1/n)"
      CurrentUnits(1) = Temp_RP.u_ROOM_TKINI
      CurrentUnits(2) = Temp_RP.u_ROOM_KINI
      lblUnitType(1) = "Time Units:"
      lblUnitType(2) = "Freundlich K Units:"
      DataRowCount = Temp_RP.int_ROOM_NKINI(NOW_CONTAMINANT)
      MaxRows = Max_int_ROOM_NKINI
      ColumnNames(2) = "Freundlich K"
  End Select
  '
  ' DISPLAY THE USER INPUT WINDOW.
  '
  Set foStoreTo = frmMain.foVarConc
  Call frmTimeVarGrid.frmTimeVarGrid_Run( _
      FormCaption, _
      UnitType(), _
      BaseUnits(), _
      CurrentUnits(), _
      lblUnitType(), _
      DataRowCount, _
      MaxRows, _
      ColumnCount, _
      ColumnNames(), _
      foStoreTo, _
      USER_HIT_CANCEL)
  If (USER_HIT_CANCEL) Then
    Exit Sub
  End If
  Select Case IN_cmdTimeVar_WhichButton
    Case IN_cmdTimeVar_WhichButton___CO:
      Temp_RP.u_ROOM_TCOINI = CurrentUnits(1)
      Temp_RP.u_ROOM_COINI = CurrentUnits(2)
      Temp_RP.int_ROOM_NCOINI(NOW_CONTAMINANT) = DataRowCount
    Case IN_cmdTimeVar_WhichButton___WA:
      Temp_RP.u_ROOM_TEMITINI = CurrentUnits(1)
      Temp_RP.u_ROOM_EMITINI = CurrentUnits(2)
      Temp_RP.int_ROOM_NEMITINI(NOW_CONTAMINANT) = DataRowCount
    Case IN_cmdTimeVar_WhichButton___K:
      Temp_RP.u_ROOM_TKINI = CurrentUnits(1)
      Temp_RP.u_ROOM_KINI = CurrentUnits(2)
      Temp_RP.int_ROOM_NKINI(NOW_CONTAMINANT) = DataRowCount
  End Select
  '
  ' EXTRA STEP:
  ' TRANSFER frmMain.foVarConc DATA INTO
  ' dbl_ROOM_COINI(), dbl_ROOM_TCOINI().
  '
  Set Ctl = frmMain.foVarConc
  Ctl.Sheet = 1
  Select Case IN_cmdTimeVar_WhichButton
    Case IN_cmdTimeVar_WhichButton___CO:
      For i = 1 To Temp_RP.int_ROOM_NCOINI(NOW_CONTAMINANT)
        Temp_RP.dbl_ROOM_TCOINI(NOW_CONTAMINANT, i) = CDbl(Val(Ctl.EntryRC(i, 1)))
        Temp_RP.dbl_ROOM_COINI(NOW_CONTAMINANT, i) = CDbl(Val(Ctl.EntryRC(i, 2)))
      Next i
    Case IN_cmdTimeVar_WhichButton___WA:
      For i = 1 To Temp_RP.int_ROOM_NEMITINI(NOW_CONTAMINANT)
        Temp_RP.dbl_ROOM_TEMITINI(NOW_CONTAMINANT, i) = CDbl(Val(Ctl.EntryRC(i, 1)))
        Temp_RP.dbl_ROOM_EMITINI(NOW_CONTAMINANT, i) = CDbl(Val(Ctl.EntryRC(i, 2)))
      Next i
    Case IN_cmdTimeVar_WhichButton___K:
      For i = 1 To Temp_RP.int_ROOM_NKINI(NOW_CONTAMINANT)
        Temp_RP.dbl_ROOM_TKINI(NOW_CONTAMINANT, i) = CDbl(Val(Ctl.EntryRC(i, 1)))
        Temp_RP.dbl_ROOM_KINI(NOW_CONTAMINANT, i) = CDbl(Val(Ctl.EntryRC(i, 2)))
      Next i
  End Select
  '
  ''DO _NOT_ ALLOW USER TO HIT CANCEL FROM INFLUENT FORM.
  ''THEY MUST SAVE ALL INFLUENT DATA BECAUSE THEY MODIFIED
  ''AN INFLUENT GRID.
  'cmdCancelOK(0).Enabled = False
  '
  'RAISE DIRTY FLAG AND REFRESH WINDOW.
  Call frmInputParamsPSDMInRoom_DirtyStatus_Throw
  Call Do_Refresh
End Sub


Private Sub cmdTimeVarConc_Click()
  Call Do___cmdTimeVar___ButtonClick(IN_cmdTimeVar_WhichButton___CO)
End Sub
Private Sub cmdTimeVarEmit_Click()
  Call Do___cmdTimeVar___ButtonClick(IN_cmdTimeVar_WhichButton___WA)
End Sub
Private Sub cmdTimeVarK_Click()
  Call Do___cmdTimeVar___ButtonClick(IN_cmdTimeVar_WhichButton___K)
End Sub


Private Sub Form_Load()
  '
  ' MISC INITS.
  '
  HALT_cboChemical = False
  HALT_ALL_CONTROLS = False
  Call CenterOnForm(Me, frmMain)
  lblSSValueUnits.Caption = Chr$(181) & "g/L"
  NOW_CONTAMINANT = 1
  Component(0) = Component(NOW_CONTAMINANT)
  Call populate_cboChemical
  Call populate_cbo_RXN_PRODUCT
  '
  ' POPULATE UNIT CONTROLS.
  '
  Call frmInputParamsPSDMInRoom_PopulateUnits
  '
  ' REFRESH WINDOW.
  '
  Call Do_Refresh
  '
  ' DATA UNCHANGED AS YET.
  '
  Call frmInputParamsPSDMInRoom_DirtyStatus_Clear
  '
  ' DEMO SETTINGS.
  '
  Call LOCAL___Reset_DemoVersionDisablings
End Sub
Private Sub Form_Unload(Cancel As Integer)
  Call unitsys_unregister_all_on_form(Me)
End Sub


Private Sub optTimeVarConc_Click(Index As Integer, Value As Integer)
Dim Ctl0 As Control
Dim Ctl1 As Control
Set Ctl0 = optTimeVarConc(0)
Set Ctl1 = optTimeVarConc(1)
Dim NewTag As Integer
Dim NewSetting As Integer
  If (HALT_ALL_CONTROLS = True) Then Exit Sub
  NewTag = Index
  If (CInt(Val(Ctl0.Tag)) <> NewTag) Then
    NewSetting = IIf(NewTag = 0, False, True)
    Temp_RP.bool_ROOM_COINI_ISTIMEVAR(NOW_CONTAMINANT) = NewSetting
    'RAISE DIRTY FLAG AND REFRESH WINDOW.
    Call frmInputParamsPSDMInRoom_DirtyStatus_Throw
    Call Do_Refresh
  End If
End Sub
Private Sub optTimeVarEmit_Click(Index As Integer, Value As Integer)
Dim Ctl0 As Control
Dim Ctl1 As Control
Set Ctl0 = optTimeVarEmit(0)
Set Ctl1 = optTimeVarEmit(1)
Dim NewTag As Integer
Dim NewSetting As Integer
  If (HALT_ALL_CONTROLS = True) Then Exit Sub
  NewTag = Index
  If (CInt(Val(Ctl0.Tag)) <> NewTag) Then
    NewSetting = IIf(NewTag = 0, False, True)
    Temp_RP.bool_ROOM_EMITINI_ISTIMEVAR(NOW_CONTAMINANT) = NewSetting
    'RAISE DIRTY FLAG AND REFRESH WINDOW.
    Call frmInputParamsPSDMInRoom_DirtyStatus_Throw
    Call Do_Refresh
  End If
End Sub
Private Sub optTimeVarK_Click(Index As Integer, Value As Integer)
Dim Ctl0 As Control
Dim Ctl1 As Control
Set Ctl0 = optTimeVarK(0)
Set Ctl1 = optTimeVarK(1)
Dim NewTag As Integer
Dim NewSetting As Integer
  If (HALT_ALL_CONTROLS = True) Then Exit Sub
  NewTag = Index
  If (CInt(Val(Ctl0.Tag)) <> NewTag) Then
    NewSetting = IIf(NewTag = 0, False, True)
    Temp_RP.bool_ROOM_KINI_ISTIMEVAR(NOW_CONTAMINANT) = NewSetting
    'RAISE DIRTY FLAG AND REFRESH WINDOW.
    Call frmInputParamsPSDMInRoom_DirtyStatus_Throw
    Call Do_Refresh
  End If
End Sub


Private Sub txtData_GotFocus(Index As Integer)
Dim Ctl As Control
Set Ctl = txtData(Index)
Dim StatusMessagePanel As String
  If (Ctl.Locked = True) Then Exit Sub
  'If (Index = 0) Then
  '  Call Global_GotFocus(Ctl)
  'Else
    Call unitsys_control_txtx_gotfocus(Ctl)
  'End If
  Select Case Index
    Case 0:
      StatusMessagePanel = "Type in the volume of the room"
    Case 1:
      StatusMessagePanel = "Type in the volumetric flow rate of air"
    Case 2:
      StatusMessagePanel = "Type in the influent concentration to the room"
    Case 3:
      StatusMessagePanel = "Type in the mass emission rate within the room"
    Case 4:
      StatusMessagePanel = "Type in the concentration at time = zero"
  End Select
  Call frmInputParamsPSDMInRoom_GenericStatus_Set(StatusMessagePanel)
End Sub
Private Sub txtData_KeyPress(Index As Integer, KeyAscii As Integer)
  'If (Index = 0) Then
  '  KeyAscii = Global_TextKeyPress(KeyAscii)
  'Else
    KeyAscii = Global_NumericKeyPress(KeyAscii)
  'End If
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
  If (Ctl.Locked = True) Then Exit Sub
'  'HANDLE THE COMPONENT NAME TEXTBOX.
'  If (Index = 0) Then
'    If (Trim$(Ctl.Text) = "") Then
'      Ctl.Text = Component(0).Name
'      'Call Show_Error("You must enter a non-blank string for the component name.")
'      'NOTE: SHOWING THIS ERROR MESSAGE MESSES UP THE
'      'SUBSEQUENT GOTFOCUS IF THE USER HITS <Enter> OR <Tab>.
'    Else
'      If (Trim$(Component(0).Name) <> Trim$(Ctl.Text)) Then
'        Component(0).Name = Trim$(Ctl.Text)
'        'THROW DIRTY FLAG.
'        Call frmCompoProp_DirtyStatus_Throw
'      End If
'    End If
'    Call Global_LostFocus(Ctl)
'    Call frmCompoProp_GenericStatus_Set("")
'    Exit Sub
'  End If
  'NOTE: LOW AND HIGH VALUES IN BASE UNITS.
  Select Case Index
    Case 0: Val_Low = 1E-20: Val_High = 1E+20
    Case 1: Val_Low = 0#: Val_High = 1E+20
    Case 2: Val_Low = 0#: Val_High = 1E+20
    Case 3: Val_Low = 0#: Val_High = 1E+20
    Case 4: Val_Low = 0#: Val_High = 1E+20
    Case 5: Val_Low = 0#: Val_High = 1E+20
    Case 6: Val_Low = 0#: Val_High = 1E+20
  End Select
  NewValue_Okay = False
  If (unitsys_control_txtx_lostfocus_validate(Ctl, Val_Low, Val_High, NewValue, Raise_Dirty_Flag)) Then
    NewValue_Okay = True
  End If
  Call unitsys_control_txtx_lostfocus(Ctl, NewValue)
  Call frmInputParamsPSDMInRoom_GenericStatus_Set("")
  If (NewValue_Okay) Then
    If (Raise_Dirty_Flag) Then
      'STORE TO MEMORY.
      Select Case Index
        Case 0: Temp_RP.ROOM_VOL = NewValue
        Case 1: Temp_RP.ROOM_FLOWRATE = NewValue
        Case 2: Temp_RP.ROOM_C0(NOW_CONTAMINANT) = NewValue
        Case 3: Temp_RP.ROOM_EMIT(NOW_CONTAMINANT) = NewValue
        Case 4: Temp_RP.INITIAL_ROOM_CONC(NOW_CONTAMINANT) = NewValue
        Case 5: Temp_RP.RXN_RATE_CONSTANT(NOW_CONTAMINANT) = NewValue
        Case 6: Temp_RP.RXN_RATIO(NOW_CONTAMINANT) = NewValue
      End Select
      'RAISE DIRTY FLAG AND RECALCULATE IF NECESSARY.
      If (Raise_Dirty_Flag) Then
        'THROW DIRTY FLAG.
        Call frmInputParamsPSDMInRoom_DirtyStatus_Throw
        Call RoomParam_Recalculate(Temp_RP)
      End If
      'REFRESH WINDOW.
      Call Do_Refresh
    End If
  End If
End Sub

