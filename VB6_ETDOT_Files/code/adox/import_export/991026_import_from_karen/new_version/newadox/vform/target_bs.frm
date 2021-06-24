VERSION 4.00
Begin VB.Form frmTarget 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Editting Target Compound ""{name}"""
   ClientHeight    =   7185
   ClientLeft      =   465
   ClientTop       =   510
   ClientWidth     =   11145
   ControlBox      =   0   'False
   Height          =   7590
   Left            =   405
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7185
   ScaleWidth      =   11145
   ShowInTaskbar   =   0   'False
   Top             =   165
   Width           =   11265
   Begin VB.CommandButton cmdCancel 
      Caption         =   "&Cancel"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
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
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
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
      Left            =   60
      TabIndex        =   2
      Top             =   30
      Width           =   10305
      _ExtentX        =   18177
      _ExtentY        =   12515
      _Version        =   262144
      Style           =   1
      TabHeight       =   520
      TabCaption(0)   =   "Tab 2"
      TabPicture(0)   =   "target_bs.frx":0000
      Tab(0).ControlCount=   0
      Tab(0).ControlEnabled=   -1  'True
      TabCaption(1)   =   "Tab 1"
      TabPicture(1)   =   "target_bs.frx":001C
      Tab(1).ControlCount=   0
      Tab(1).ControlEnabled=   0   'False
      TabCaption(2)   =   "Tab 2"
      TabPicture(2)   =   "target_bs.frx":0038
      Tab(2).ControlCount=   0
      Tab(2).ControlEnabled=   0   'False
      Begin Threed.SSFrame SSFrame1 
         Height          =   3015
         Left            =   -74880
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
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
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
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
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
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
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
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
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
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
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
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
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
                  Picture         =   "target_bs.frx":0054
                  Top             =   30
                  Width           =   3915
               End
            End
         End
         Begin VB.Label lbldesc2 
            Caption         =   "gmol/L"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
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
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
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
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
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
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
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
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
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
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
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
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
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
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
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
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
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
            Caption         =   "# H-Substit'd Atoms per Molecule"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
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
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
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
            Width           =   795
         End
         Begin VB.Label lbldesc 
            Caption         =   "Second Order Rate Constant"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
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
         Left            =   -74880
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
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
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
                  Picture         =   "target_bs.frx":13C06
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
                  Picture         =   "target_bs.frx":1A858
                  Top             =   30
                  Width           =   2475
               End
            End
         End
         Begin VB.Label lbldesc 
            Caption         =   "Equilibrium Constant:"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
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
            TabIndex        =   31
            Top             =   465
            Width           =   1965
         End
         Begin VB.Label lbldesc2 
            Caption         =   "gmol/L"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
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
            Left            =   3300
            TabIndex        =   30
            Top             =   465
            Width           =   795
         End
      End
      Begin Threed.SSFrame SSFrame6 
         Height          =   1845
         Left            =   -74880
         TabIndex        =   32
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
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
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
            TabIndex        =   35
            Text            =   "txtData(9)"
            Top             =   1200
            Width           =   1095
         End
         Begin VB.TextBox txtData 
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
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
            TabIndex        =   34
            Text            =   "txtData(8)"
            Top             =   780
            Width           =   1095
         End
         Begin VB.TextBox txtData 
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
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
            TabIndex        =   33
            Text            =   "txtData(7)"
            Top             =   360
            Width           =   1095
         End
         Begin Threed.SSFrame SSFrame7 
            Height          =   1815
            Left            =   5730
            TabIndex        =   36
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
               TabIndex        =   37
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
                  Picture         =   "target_bs.frx":2439A
                  Top             =   30
                  Width           =   3720
               End
               Begin VB.Image Image5 
                  Height          =   630
                  Left            =   270
                  Picture         =   "target_bs.frx":2E69C
                  Top             =   810
                  Width           =   3390
               End
            End
         End
         Begin VB.Label lbldesc 
            Caption         =   "Second Order Rate Constant"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
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
            TabIndex        =   43
            Top             =   1245
            Width           =   3000
         End
         Begin VB.Label lbldesc2 
            Caption         =   "L/gmol-s"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
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
            TabIndex        =   42
            Top             =   1245
            Width           =   795
         End
         Begin VB.Label lbldesc 
            Caption         =   "Molecular Weight"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
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
            TabIndex        =   41
            Top             =   825
            Width           =   3000
         End
         Begin VB.Label lbldesc2 
            Caption         =   "g/gmol"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
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
            TabIndex        =   40
            Top             =   825
            Width           =   795
         End
         Begin VB.Label lbldesc 
            Caption         =   "Valence"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
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
            TabIndex        =   39
            Top             =   405
            Width           =   3000
         End
         Begin VB.Label lbldesc2 
            Caption         =   "-"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
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
            TabIndex        =   38
            Top             =   405
            Width           =   795
         End
      End
   End
End
Attribute VB_Name = "frmTarget"
Attribute VB_Creatable = False
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
  MARGIN = sstab_main.Left
  Me.Width = (Me.Width - Me.ScaleWidth) + MARGIN * 2 + sstab_main.Width
  Call CenterOnForm(Me, frmMain)
  Call refresh_frmTarget(NowTargetCompound)

  'UPDATE WINDOW CAPTION.
  Select Case FRMTARGET_MODE
    Case FRMTARGET_ADDNEW:
      Me.Caption = "Adding New Target Compound"
    Case FRMTARGET_EDIT:
      Me.Caption = "Editting Target Compound " & Chr$(34) & NowTargetCompound.comname & Chr$(34)
  End Select
End Sub


Private Sub txtdata_GotFocus(Index As Integer)
  Dim txtctl As Control
  Set txtctl = txtdata(Index)
  Call DisplayDataEntryError
  Call Global_GotFocus(txtctl)
End Sub
Private Sub txtdata_KeyPress(Index As Integer, KeyAscii As Integer)
  KeyAscii = Global_NumericKeyPress(KeyAscii)
End Sub
Private Sub txtdata_LostFocus(Index As Integer)
Dim newVal As Double
Dim txtctl As Control
Set txtctl = txtdata(Index)
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
    End Select
    
    Call AssignTextAndTag(txtctl, newVal)
    
    Select Case refresh_type
      Case 1:   'JUST THE MAIN WINDOW.
        Call refresh_frmTarget(NowTargetCompound)
    End Select
  End If
  Call Global_LostFocus(txtctl)
End Sub





