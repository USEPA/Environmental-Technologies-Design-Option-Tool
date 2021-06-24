VERSION 5.00
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "THREED32.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form frmTarget 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Editing Target Compound ""{name}"""
   ClientHeight    =   7185
   ClientLeft      =   2595
   ClientTop       =   1230
   ClientWidth     =   11145
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   7185
   ScaleWidth      =   11145
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmdCancel 
      Caption         =   "&Cancel"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   9300
      TabIndex        =   1
      Top             =   30
      Width           =   1065
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "&OK"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   8190
      TabIndex        =   0
      Top             =   30
      Width           =   1065
   End
   Begin TabDlg.SSTab sstab_main 
      Height          =   7095
      Left            =   90
      TabIndex        =   2
      Top             =   30
      Width           =   10305
      _ExtentX        =   18177
      _ExtentY        =   12515
      _Version        =   393216
      Style           =   1
      TabHeight       =   520
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      TabCaption(0)   =   "Main"
      TabPicture(0)   =   "target.frx":0000
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "SSFrame6"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "SSFrame3"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "SSFrame1"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).ControlCount=   3
      TabCaption(1)   =   "CO3*- and HPO4*- Rate Constants"
      TabPicture(1)   =   "target.frx":001C
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "SSFrame9"
      Tab(1).Control(1)=   "SSFrame10"
      Tab(1).ControlCount=   2
      TabCaption(2)   =   "O2*- and HO2* Rate Constants"
      TabPicture(2)   =   "target.frx":0038
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "SSFrame12"
      Tab(2).Control(1)=   "SSFrame14"
      Tab(2).ControlCount=   2
      Begin Threed.SSFrame SSFrame1 
         Height          =   3015
         Left            =   120
         TabIndex        =   3
         Top             =   360
         Width           =   10065
         _Version        =   65536
         _ExtentX        =   17754
         _ExtentY        =   5318
         _StockProps     =   14
         Caption         =   "Properties of Protonated Form ""HR"":"
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
            Index           =   0
            Left            =   3180
            TabIndex        =   9
            Text            =   "txtData(0)"
            Top             =   420
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
            Index           =   1
            Left            =   3180
            TabIndex        =   8
            Text            =   "txtData(1)"
            Top             =   840
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
            Left            =   3180
            TabIndex        =   7
            Text            =   "txtData(2)"
            Top             =   1260
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
            Left            =   3180
            TabIndex        =   6
            Text            =   "txtData(3)"
            Top             =   1680
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
            Index           =   4
            Left            =   3180
            TabIndex        =   5
            Text            =   "txtData(4)"
            Top             =   2100
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
            Left            =   3180
            TabIndex        =   4
            Text            =   "txtData(5)"
            Top             =   2520
            Width           =   1095
         End
         Begin Threed.SSFrame SSFrame2 
            Height          =   1905
            Left            =   5730
            TabIndex        =   10
            Top             =   30
            Width           =   4335
            _Version        =   65536
            _ExtentX        =   7646
            _ExtentY        =   3360
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
            Begin Threed.SSPanel panEq1 
               Height          =   1635
               Left            =   120
               TabIndex        =   11
               Top             =   180
               Width           =   4125
               _Version        =   65536
               _ExtentX        =   7276
               _ExtentY        =   2884
               _StockProps     =   15
               ForeColor       =   0
               BackColor       =   16777215
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               BevelWidth      =   3
               Begin VB.Image Image1 
                  Height          =   1545
                  Left            =   90
                  Picture         =   "target.frx":0054
                  Top             =   30
                  Width           =   3915
               End
            End
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
            Index           =   0
            Left            =   4380
            TabIndex        =   23
            Top             =   465
            Width           =   795
         End
         Begin VB.Label lbldesc 
            Caption         =   "{Influent} Concentration"
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
            TabIndex        =   22
            Top             =   465
            Width           =   3000
         End
         Begin VB.Label lbldesc2 
            Caption         =   "-"
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
            Left            =   4380
            TabIndex        =   21
            Top             =   885
            Width           =   795
         End
         Begin VB.Label lbldesc 
            Caption         =   "Valence"
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
            TabIndex        =   20
            Top             =   885
            Width           =   3000
         End
         Begin VB.Label lbldesc2 
            Caption         =   "g/gmol"
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
            Left            =   4380
            TabIndex        =   19
            Top             =   1305
            Width           =   795
         End
         Begin VB.Label lbldesc 
            Caption         =   "Molecular Weight"
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
            TabIndex        =   18
            Top             =   1305
            Width           =   3000
         End
         Begin VB.Label lbldesc2 
            Caption         =   "-"
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
            Left            =   4380
            TabIndex        =   17
            Top             =   1725
            Width           =   795
         End
         Begin VB.Label lbldesc 
            Caption         =   "# C Atoms per Molecule"
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
            TabIndex        =   16
            Top             =   1725
            Width           =   3000
         End
         Begin VB.Label lbldesc2 
            Caption         =   "-"
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
            Left            =   4380
            TabIndex        =   15
            Top             =   2145
            Width           =   795
         End
         Begin VB.Label lbldesc 
            Caption         =   "# Halogen Atoms per Molecule"
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
            Left            =   120
            TabIndex        =   14
            Top             =   2145
            Width           =   3000
         End
         Begin VB.Label lbldesc2 
            Caption         =   "L/gmol-s"
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
            Left            =   4380
            TabIndex        =   13
            Top             =   2565
            Width           =   1400
         End
         Begin VB.Label lbldesc 
            Caption         =   "Second Order Rate Constant"
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
            TabIndex        =   12
            Top             =   2565
            Width           =   3000
         End
      End
      Begin Threed.SSFrame SSFrame3 
         Height          =   1635
         Left            =   120
         TabIndex        =   24
         Top             =   3450
         Width           =   10065
         _Version        =   65536
         _ExtentX        =   17754
         _ExtentY        =   2884
         _StockProps     =   14
         Caption         =   "Equilibrium Reaction:"
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
            Index           =   6
            Left            =   2100
            TabIndex        =   25
            Text            =   "txtData(6)"
            Top             =   420
            Width           =   1095
         End
         Begin Threed.SSFrame SSFrame4 
            Height          =   1605
            Left            =   4020
            TabIndex        =   26
            Top             =   30
            Width           =   3135
            _Version        =   65536
            _ExtentX        =   5530
            _ExtentY        =   2831
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
            Begin Threed.SSPanel SSPanel1 
               Height          =   1335
               Left            =   120
               TabIndex        =   27
               Top             =   180
               Width           =   2895
               _Version        =   65536
               _ExtentX        =   5106
               _ExtentY        =   2355
               _StockProps     =   15
               ForeColor       =   0
               BackColor       =   16777215
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               BevelWidth      =   3
               Begin VB.Image Image2 
                  Height          =   780
                  Left            =   90
                  Picture         =   "target.frx":13C06
                  Top             =   240
                  Width           =   2655
               End
            End
         End
         Begin Threed.SSFrame SSFrame5 
            Height          =   1605
            Left            =   7140
            TabIndex        =   28
            Top             =   30
            Width           =   2925
            _Version        =   65536
            _ExtentX        =   5159
            _ExtentY        =   2831
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
            Begin Threed.SSPanel SSPanel2 
               Height          =   1335
               Left            =   120
               TabIndex        =   29
               Top             =   180
               Width           =   2715
               _Version        =   65536
               _ExtentX        =   4789
               _ExtentY        =   2355
               _StockProps     =   15
               ForeColor       =   0
               BackColor       =   16777215
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               BevelWidth      =   3
               Begin VB.Image Image3 
                  Height          =   1200
                  Left            =   90
                  Picture         =   "target.frx":1A858
                  Top             =   30
                  Width           =   2475
               End
            End
         End
         Begin VB.Label lbldesc 
            Caption         =   "Equilibrium Constant:"
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
            TabIndex        =   30
            Top             =   465
            Width           =   1965
         End
      End
      Begin Threed.SSFrame SSFrame6 
         Height          =   1845
         Left            =   120
         TabIndex        =   31
         Top             =   5160
         Width           =   10065
         _Version        =   65536
         _ExtentX        =   17754
         _ExtentY        =   3254
         _StockProps     =   14
         Caption         =   "Properties of De-protonated Form ""R(-)"":"
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
            Index           =   9
            Left            =   3180
            TabIndex        =   34
            Text            =   "txtData(9)"
            Top             =   1200
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
            Index           =   8
            Left            =   3180
            TabIndex        =   33
            Text            =   "txtData(8)"
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
            Index           =   7
            Left            =   3180
            TabIndex        =   32
            Text            =   "txtData(7)"
            Top             =   360
            Width           =   1095
         End
         Begin Threed.SSFrame SSFrame7 
            Height          =   1815
            Left            =   5730
            TabIndex        =   35
            Top             =   30
            Width           =   4335
            _Version        =   65536
            _ExtentX        =   7646
            _ExtentY        =   3201
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
            Begin Threed.SSPanel SSPanel3 
               Height          =   1545
               Left            =   120
               TabIndex        =   36
               Top             =   180
               Width           =   4125
               _Version        =   65536
               _ExtentX        =   7276
               _ExtentY        =   2725
               _StockProps     =   15
               ForeColor       =   0
               BackColor       =   16777215
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               BevelWidth      =   3
               Begin VB.Image Image4 
                  Height          =   840
                  Left            =   90
                  Picture         =   "target.frx":2439A
                  Top             =   30
                  Width           =   3720
               End
               Begin VB.Image Image5 
                  Height          =   630
                  Left            =   270
                  Picture         =   "target.frx":2E69C
                  Top             =   810
                  Width           =   3390
               End
            End
         End
         Begin VB.Label lbldesc 
            Caption         =   "Second Order Rate Constant"
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
            Top             =   1245
            Width           =   3000
         End
         Begin VB.Label lbldesc2 
            Caption         =   "L/gmol-s"
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
            Left            =   4380
            TabIndex        =   41
            Top             =   1245
            Width           =   795
         End
         Begin VB.Label lbldesc 
            Caption         =   "Molecular Weight"
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
            Top             =   825
            Width           =   3000
         End
         Begin VB.Label lbldesc2 
            Caption         =   "g/gmol"
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
            Left            =   4380
            TabIndex        =   39
            Top             =   825
            Width           =   795
         End
         Begin VB.Label lbldesc 
            Caption         =   "Valence"
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
            TabIndex        =   38
            Top             =   405
            Width           =   3000
         End
         Begin VB.Label lbldesc2 
            Caption         =   "-"
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
            Left            =   4380
            TabIndex        =   37
            Top             =   405
            Width           =   795
         End
      End
      Begin Threed.SSFrame SSFrame9 
         Height          =   1935
         Left            =   -74910
         TabIndex        =   43
         Top             =   420
         Width           =   10095
         _Version        =   65536
         _ExtentX        =   17806
         _ExtentY        =   3413
         _StockProps     =   14
         Caption         =   "Reaction With CO3*-"
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
            Index           =   10
            Left            =   2940
            TabIndex        =   46
            Text            =   "txtData(10)"
            Top             =   390
            Width           =   1095
         End
         Begin Threed.SSFrame SSFrame8 
            Height          =   1905
            Left            =   5250
            TabIndex        =   44
            Top             =   30
            Width           =   4845
            _Version        =   65536
            _ExtentX        =   8546
            _ExtentY        =   3360
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
            Begin Threed.SSPanel SSPanel4 
               Height          =   1635
               Left            =   120
               TabIndex        =   45
               Top             =   180
               Width           =   4635
               _Version        =   65536
               _ExtentX        =   8176
               _ExtentY        =   2884
               _StockProps     =   15
               ForeColor       =   0
               BackColor       =   16777215
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               BevelWidth      =   3
               Begin VB.Image Image6 
                  Height          =   1440
                  Left            =   210
                  Picture         =   "target.frx":3566E
                  Top             =   60
                  Width           =   4170
               End
            End
         End
         Begin VB.Label lbldesc 
            Caption         =   "Second Order Rate Constant"
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
            TabIndex        =   48
            Top             =   435
            Width           =   3000
         End
         Begin VB.Label lbldesc2 
            Caption         =   "L/gmol-s"
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
            Left            =   4140
            TabIndex        =   47
            Top             =   435
            Width           =   795
         End
      End
      Begin Threed.SSFrame SSFrame10 
         Height          =   1935
         Left            =   -74910
         TabIndex        =   49
         Top             =   2520
         Width           =   10095
         _Version        =   65536
         _ExtentX        =   17806
         _ExtentY        =   3413
         _StockProps     =   14
         Caption         =   "Reaction With HPO4*-"
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
            Index           =   11
            Left            =   2940
            TabIndex        =   50
            Text            =   "txtData(11)"
            Top             =   390
            Width           =   1095
         End
         Begin Threed.SSFrame SSFrame11 
            Height          =   1905
            Left            =   5250
            TabIndex        =   51
            Top             =   30
            Width           =   4845
            _Version        =   65536
            _ExtentX        =   8546
            _ExtentY        =   3360
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
            Begin Threed.SSPanel SSPanel5 
               Height          =   1635
               Left            =   120
               TabIndex        =   52
               Top             =   180
               Width           =   4635
               _Version        =   65536
               _ExtentX        =   8176
               _ExtentY        =   2884
               _StockProps     =   15
               ForeColor       =   0
               BackColor       =   16777215
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               BevelWidth      =   3
               Begin VB.Image Image7 
                  Height          =   1395
                  Left            =   90
                  Picture         =   "target.frx":49030
                  Top             =   30
                  Width           =   4425
               End
            End
         End
         Begin VB.Label lbldesc2 
            Caption         =   "L/gmol-s"
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
            Index           =   11
            Left            =   4140
            TabIndex        =   54
            Top             =   435
            Width           =   795
         End
         Begin VB.Label lbldesc 
            Caption         =   "Second Order Rate Constant"
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
            Index           =   11
            Left            =   120
            TabIndex        =   53
            Top             =   435
            Width           =   3000
         End
      End
      Begin Threed.SSFrame SSFrame12 
         Height          =   1935
         Left            =   -74910
         TabIndex        =   55
         Top             =   420
         Width           =   10095
         _Version        =   65536
         _ExtentX        =   17806
         _ExtentY        =   3413
         _StockProps     =   14
         Caption         =   "Reaction With O2*-"
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
            Index           =   12
            Left            =   2940
            TabIndex        =   56
            Text            =   "txtData(12)"
            Top             =   390
            Width           =   1095
         End
         Begin Threed.SSFrame SSFrame13 
            Height          =   1905
            Left            =   5250
            TabIndex        =   57
            Top             =   30
            Width           =   4845
            _Version        =   65536
            _ExtentX        =   8546
            _ExtentY        =   3360
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
            Begin Threed.SSPanel SSPanel6 
               Height          =   1635
               Left            =   120
               TabIndex        =   58
               Top             =   180
               Width           =   4635
               _Version        =   65536
               _ExtentX        =   8176
               _ExtentY        =   2884
               _StockProps     =   15
               ForeColor       =   0
               BackColor       =   16777215
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               BevelWidth      =   3
               Begin VB.Image Image8 
                  Height          =   1365
                  Left            =   210
                  Picture         =   "target.frx":5D30A
                  Top             =   60
                  Width           =   3930
               End
            End
         End
         Begin VB.Label lbldesc2 
            Caption         =   "L/gmol-s"
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
            Index           =   12
            Left            =   4140
            TabIndex        =   60
            Top             =   435
            Width           =   795
         End
         Begin VB.Label lbldesc 
            Caption         =   "Second Order Rate Constant"
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
            Index           =   12
            Left            =   120
            TabIndex        =   59
            Top             =   435
            Width           =   3000
         End
      End
      Begin Threed.SSFrame SSFrame14 
         Height          =   1935
         Left            =   -74910
         TabIndex        =   61
         Top             =   2520
         Width           =   10095
         _Version        =   65536
         _ExtentX        =   17806
         _ExtentY        =   3413
         _StockProps     =   14
         Caption         =   "Reaction With HO2*"
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
            Index           =   13
            Left            =   2940
            TabIndex        =   62
            Text            =   "txtData(13)"
            Top             =   390
            Width           =   1095
         End
         Begin Threed.SSFrame SSFrame15 
            Height          =   1905
            Left            =   5250
            TabIndex        =   63
            Top             =   30
            Width           =   4845
            _Version        =   65536
            _ExtentX        =   8546
            _ExtentY        =   3360
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
            Begin Threed.SSPanel SSPanel7 
               Height          =   1635
               Left            =   120
               TabIndex        =   64
               Top             =   180
               Width           =   4635
               _Version        =   65536
               _ExtentX        =   8176
               _ExtentY        =   2884
               _StockProps     =   15
               ForeColor       =   0
               BackColor       =   16777215
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               BevelWidth      =   3
               Begin VB.Image Image9 
                  Height          =   1395
                  Left            =   210
                  Picture         =   "target.frx":6EB68
                  Top             =   60
                  Width           =   4095
               End
            End
         End
         Begin VB.Label lbldesc 
            Caption         =   "Second Order Rate Constant"
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
            Index           =   13
            Left            =   120
            TabIndex        =   66
            Top             =   435
            Width           =   3000
         End
         Begin VB.Label lbldesc2 
            Caption         =   "L/gmol-s"
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
            Index           =   13
            Left            =   4140
            TabIndex        =   65
            Top             =   435
            Width           =   795
         End
      End
   End
End
Attribute VB_Name = "frmTarget"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim USER_HIT_CANCEL As Integer
Dim FRMTARGET_MODE As Integer
Const FRMTARGET_ADDNEW = 1
Const FRMTARGET_EDIT = 2

Dim NowTargetCompound As TargetCompound_Type






Const frmTarget_Declarations_End = 1


'RETURNS:
'  TRUE = USER HIT OK
'  FALSE = USER HIT CANCEL
Public Function frmTarget_DoAddNew() As Integer
Dim is_aborted As Integer
Dim name_new As String
Dim i As Integer
Dim j As Integer
Dim idx_new As Integer
Dim temp() As Double

  'TELL FORM TO DO "ADD" MODE.
  FRMTARGET_MODE = FRMTARGET_ADDNEW
  
  'INIT THIS TARGET COMPOUND.
  Call TargetCompound_SetDefaults(NowTargetCompound)
  NowTargetCompound.comname = "new"
  
  'SHOW THE FORM.
  frmTarget.Show 1
  
  'UPDATE MEMORY.
  If (Not USER_HIT_CANCEL) Then
    Do While (1 = 1)
      name_new = frmNewName.frmNewName_GetName( _
          "Enter Name for New Target Compound", _
          "Each target compound must have a unique name.", _
          name_new, _
          is_aborted)
      If (is_aborted) Then
        'INFORM MAIN WINDOW THAT USER CANCELLED.
        frmTarget_DoAddNew = False
        Exit Function
      End If
      If (Not TargetCompound_IsKeyExist(NowProj, name_new)) Then
        Exit Do
      End If
      Call Show_Error("That name already exists.  Choose another name.")
    Loop
    NowTargetCompound.comname = name_new
    
    'ADD A NEW TARGET COMPOUND STRUCTURE.
    NowProj.TargetCompounds_Count = NowProj.TargetCompounds_Count + 1
    ReDim Preserve NowProj.TargetCompounds(1 To NowProj.TargetCompounds_Count)
    idx_new = NowProj.TargetCompounds_Count
    NowProj.TargetCompounds(idx_new) = NowTargetCompound
  
    'ADD A NEW SET OF EXTINCTION COEFFICIENTS (ONE PER WAVELENGTH).
    'NOTE: THIS STUPID TEMPORARY ARRAY IS NECESSARY TO GET AROUND THE
    'VISUAL BASIC STIPULATION THAT YOU CAN ONLY "REDIM PRESERVE" AN ARRAY
    'IF YOU ARE CHANGING THE *LAST* ARRAY INDEX.
    ReDim temp(1 To idx_new - 1, 1 To NowProj.Wavelength_Count)
    For i = 1 To idx_new - 1
      For j = 1 To NowProj.Wavelength_Count
        temp(i, j) = NowProj.extcoef(i, j)
      Next j
    Next i
    ReDim NowProj.extcoef(1 To idx_new, 1 To NowProj.Wavelength_Count)
    For i = 1 To idx_new - 1
      For j = 1 To NowProj.Wavelength_Count
        NowProj.extcoef(i, j) = temp(i, j)
      Next j
    Next i
    'DEFAULT VALUES FOR THE NEW TARGET COMPOUND.
    For i = 1 To NowProj.Wavelength_Count
      NowProj.extcoef(idx_new, i) = EXTCOEF_DEFAULT_VALUE
    Next i
    
    'ADD A NEW SET OF QUANTUM YIELDS (ONE PER WAVELENGTH).
    'NOTE: THIS STUPID TEMPORARY ARRAY IS NECESSARY TO GET AROUND THE
    'VISUAL BASIC STIPULATION THAT YOU CAN ONLY "REDIM PRESERVE" AN ARRAY
    'IF YOU ARE CHANGING THE *LAST* ARRAY INDEX.
    ReDim temp(1 To idx_new - 1, 1 To NowProj.Wavelength_Count)
    For i = 1 To idx_new - 1
      For j = 1 To NowProj.Wavelength_Count
        temp(i, j) = NowProj.quatyd(i, j)
      Next j
    Next i
    ReDim NowProj.quatyd(1 To idx_new, 1 To NowProj.Wavelength_Count)
    For i = 1 To idx_new - 1
      For j = 1 To NowProj.Wavelength_Count
        NowProj.quatyd(i, j) = temp(i, j)
      Next j
    Next i
    'DEFAULT VALUES FOR THE NEW TARGET COMPOUND.
    For i = 1 To NowProj.Wavelength_Count
      NowProj.quatyd(idx_new, i) = QUATYD_DEFAULT_VALUE
    Next i
  End If
  
  'RETURN TO MAIN WINDOW.
  frmTarget_DoAddNew = Not USER_HIT_CANCEL

End Function


'RETURNS:
'  TRUE = USER HIT OK
'  FALSE = USER HIT CANCEL
Public Function frmTarget_DoEdit(tcnum As Integer) As Integer
Dim is_aborted As Integer
Dim name_new As String
  
  'TELL FORM TO DO "EDIT" MODE.
  FRMTARGET_MODE = FRMTARGET_EDIT
  
  'IMPORT THIS TARGET COMPOUND FROM MEMORY TO THE FORM.
  NowTargetCompound = NowProj.TargetCompounds(tcnum)
  
  'SHOW THE FORM.
  frmTarget.Show 1
  
  'UPDATE MEMORY.
  If (Not USER_HIT_CANCEL) Then
    NowProj.TargetCompounds(tcnum) = NowTargetCompound
  End If
  
  'RETURN TO MAIN WINDOW.
  frmTarget_DoEdit = Not USER_HIT_CANCEL

End Function


Private Sub cmdCancel_Click()
  USER_HIT_CANCEL = True
  Unload Me
End Sub


Private Sub cmdOK_Click()
  USER_HIT_CANCEL = False
  Unload Me
End Sub


Private Sub Form_Load()
Dim MARGIN As Long
  'MISC INITS.
  MARGIN = sstab_main.left
  Me.width = (Me.width - Me.ScaleWidth) + MARGIN * 2 + sstab_main.width
  Call CenterOnForm(Me, frmMain)
  Call refresh_frmTarget(NowTargetCompound)
  sstab_main.Tab = 0

  'UPDATE WINDOW CAPTION.
  Select Case FRMTARGET_MODE
    Case FRMTARGET_ADDNEW:
      Me.Caption = "Adding New Target Compound"
    Case FRMTARGET_EDIT:
      Me.Caption = "Editing Target Compound " & Chr$(34) & NowTargetCompound.comname & Chr$(34)
  End Select
End Sub


Private Sub txtdata_GotFocus(Index As Integer)
  Dim txtctl As Control
  Set txtctl = txtData(Index)
  Call DisplayDataEntryError
  Call Global_GotFocus(txtctl)
End Sub
Private Sub txtdata_KeyPress(Index As Integer, KeyAscii As Integer)
  KeyAscii = Global_NumericKeyPress(KeyAscii)
End Sub
Private Sub txtdata_LostFocus(Index As Integer)
Dim newVal As Double
Dim txtctl As Control
Set txtctl = txtData(Index)
Dim ok_to_save As Integer
Dim xmin As Double
Dim xmax As Double
Dim refresh_type As Integer
  
  xmin = val(txtctl.LinkItem)
  xmax = val(txtctl.DataField)
  
  ok_to_save = False
  If (ValueHasChanged(txtctl)) Then
    If (IsValidNumber(txtctl, vbDouble)) Then
      newVal = CDbl(txtctl.Text)
      If (newVal < xmin) Or (newVal > xmax) Then
        txtctl.Text = txtctl.Tag
      Else
        If (txtctl.Text <> txtctl.Tag) Then
          ok_to_save = True
        End If
      End If
    Else
      txtctl.Text = txtctl.Tag
    End If
  End If
      
  If (ok_to_save) Then
    'DATA LOOKS OKAY, LET'S GO AHEAD AND SAVE IT.
    refresh_type = 1
    Select Case Index
      Case 0: NowTargetCompound.concini = newVal
      Case 1: NowTargetCompound.val = newVal
      Case 2: NowTargetCompound.mw = newVal
      Case 3: NowTargetCompound.ncarbn = newVal
      Case 4: NowTargetCompound.nsubstt = newVal
      Case 5: NowTargetCompound.xk = newVal
      Case 6: NowTargetCompound.dep_xke = newVal
      Case 7: NowTargetCompound.dep_val = newVal
      Case 8: NowTargetCompound.dep_mw = newVal
      Case 9: NowTargetCompound.dep_xk = newVal
      Case 10: NowTargetCompound.xk_co3XM = newVal
      Case 11: NowTargetCompound.xk_hpo4XM = newVal
      Case 12: NowTargetCompound.xk_o2XM = newVal
      Case 13: NowTargetCompound.xk_ho2X = newVal
    End Select
    
    Call AssignTextAndTag(txtctl, newVal)
    
    Select Case refresh_type
      Case 1:   'JUST THE MAIN WINDOW.
        Call refresh_frmTarget(NowTargetCompound)
    End Select
  End If
  Call Global_LostFocus(txtctl)
End Sub





