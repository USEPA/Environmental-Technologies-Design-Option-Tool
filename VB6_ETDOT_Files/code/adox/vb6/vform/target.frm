VERSION 5.00
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "THREED32.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "tabctl32.ocx"
Begin VB.Form frmTarget 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Editing Target Compound"
   ClientHeight    =   6990
   ClientLeft      =   855
   ClientTop       =   1665
   ClientWidth     =   10125
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   6990
   ScaleWidth      =   10125
   ShowInTaskbar   =   0   'False
   Begin Threed.SSFrame SSFrame2 
      Height          =   1215
      Left            =   30
      TabIndex        =   16
      Top             =   0
      Width           =   6855
      _Version        =   65536
      _ExtentX        =   12091
      _ExtentY        =   2143
      _StockProps     =   14
      Caption         =   " "
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Begin VB.CommandButton cmdDataBase 
         Caption         =   "Use Database"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   4980
         TabIndex        =   56
         Top             =   360
         Width           =   1695
      End
      Begin VB.TextBox txtDataStr 
         Height          =   375
         Index           =   0
         Left            =   1920
         TabIndex        =   0
         Text            =   "txtDataStr(0)"
         Top             =   240
         Width           =   2775
      End
      Begin VB.TextBox txtData 
         Height          =   375
         Index           =   15
         Left            =   1920
         TabIndex        =   1
         Text            =   "txtData(15)"
         Top             =   660
         Width           =   2055
      End
      Begin VB.Label lblDesc 
         Caption         =   "Compound Name"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   14
         Left            =   120
         TabIndex        =   18
         Top             =   360
         Width           =   1935
      End
      Begin VB.Label lblDesc 
         Caption         =   "CAS Number"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   15
         Left            =   120
         TabIndex        =   17
         Top             =   690
         Width           =   1455
      End
   End
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
      Height          =   435
      Left            =   8040
      TabIndex        =   21
      Top             =   450
      Width           =   1185
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
      Height          =   435
      Left            =   7110
      TabIndex        =   20
      Top             =   450
      Width           =   705
   End
   Begin TabDlg.SSTab sstab_main 
      Height          =   5025
      Left            =   60
      TabIndex        =   19
      Top             =   1290
      Width           =   9315
      _ExtentX        =   16431
      _ExtentY        =   8864
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
      TabCaption(0)   =   "          Main         "
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
         Height          =   1695
         Left            =   30
         TabIndex        =   22
         Top             =   450
         Width           =   9225
         _Version        =   65536
         _ExtentX        =   16272
         _ExtentY        =   2990
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
            Index           =   5
            Left            =   7020
            TabIndex        =   7
            Text            =   "txtData(5)"
            Top             =   1125
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
            Left            =   6990
            TabIndex        =   6
            Text            =   "txtData(4)"
            Top             =   750
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
            Left            =   6990
            TabIndex        =   5
            Text            =   "txtData(3)"
            Top             =   330
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
            Left            =   2130
            TabIndex        =   2
            Text            =   "txtData(0)"
            Top             =   330
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
            Left            =   2130
            TabIndex        =   3
            Text            =   "txtData(1)"
            Top             =   750
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
            Left            =   2130
            TabIndex        =   4
            Text            =   "txtData(2)"
            Top             =   1125
            Width           =   1095
         End
         Begin VB.Label lblDesc 
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
            Left            =   4110
            TabIndex        =   34
            Top             =   1200
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
            Left            =   8160
            TabIndex        =   33
            Top             =   1200
            Width           =   795
         End
         Begin VB.Label lblDesc 
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
            Left            =   4110
            TabIndex        =   32
            Top             =   825
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
            Left            =   8160
            TabIndex        =   31
            Top             =   825
            Width           =   675
         End
         Begin VB.Label lblDesc 
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
            Left            =   4110
            TabIndex        =   30
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
            Index           =   3
            Left            =   8160
            TabIndex        =   29
            Top             =   405
            Width           =   795
         End
         Begin VB.Label lblDesc 
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
            TabIndex        =   28
            Top             =   1200
            Width           =   2070
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
            Left            =   3270
            TabIndex        =   27
            Top             =   1200
            Width           =   795
         End
         Begin VB.Label lblDesc 
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
            TabIndex        =   26
            Top             =   825
            Width           =   2010
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
            Left            =   3270
            TabIndex        =   25
            Top             =   825
            Width           =   795
         End
         Begin VB.Label lblDesc 
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
            TabIndex        =   24
            Top             =   405
            Width           =   2040
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
            Left            =   3270
            TabIndex        =   23
            Top             =   405
            Width           =   795
         End
      End
      Begin Threed.SSFrame SSFrame3 
         Height          =   945
         Left            =   30
         TabIndex        =   35
         Top             =   2190
         Width           =   9225
         _Version        =   65536
         _ExtentX        =   16272
         _ExtentY        =   1667
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
            Left            =   2280
            TabIndex        =   8
            Text            =   "txtData(6)"
            Top             =   420
            Width           =   1095
         End
         Begin VB.Label lblDesc 
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
            TabIndex        =   36
            Top             =   465
            Width           =   1965
         End
      End
      Begin Threed.SSFrame SSFrame6 
         Height          =   1695
         Left            =   30
         TabIndex        =   37
         Top             =   3210
         Width           =   9255
         _Version        =   65536
         _ExtentX        =   16325
         _ExtentY        =   2990
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
            Index           =   7
            Left            =   3180
            TabIndex        =   9
            Text            =   "txtData(7)"
            Top             =   360
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
            TabIndex        =   10
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
            Index           =   9
            Left            =   3180
            TabIndex        =   11
            Text            =   "txtData(9)"
            Top             =   1200
            Width           =   1095
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
            TabIndex        =   43
            Top             =   405
            Width           =   795
         End
         Begin VB.Label lblDesc 
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
            TabIndex        =   42
            Top             =   405
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
            TabIndex        =   41
            Top             =   825
            Width           =   795
         End
         Begin VB.Label lblDesc 
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
            TabIndex        =   39
            Top             =   1245
            Width           =   795
         End
         Begin VB.Label lblDesc 
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
            TabIndex        =   38
            Top             =   1245
            Width           =   3000
         End
      End
      Begin Threed.SSFrame SSFrame9 
         Height          =   1935
         Left            =   -74970
         TabIndex        =   44
         Top             =   570
         Width           =   9225
         _Version        =   65536
         _ExtentX        =   16272
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
            TabIndex        =   12
            Text            =   "txtData(10)"
            Top             =   390
            Width           =   1095
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
            TabIndex        =   46
            Top             =   435
            Width           =   795
         End
         Begin VB.Label lblDesc 
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
            TabIndex        =   45
            Top             =   435
            Width           =   3000
         End
      End
      Begin Threed.SSFrame SSFrame10 
         Height          =   1935
         Left            =   -74970
         TabIndex        =   47
         Top             =   2760
         Width           =   9225
         _Version        =   65536
         _ExtentX        =   16272
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
            Left            =   3000
            TabIndex        =   13
            Text            =   "txtData(11)"
            Top             =   390
            Width           =   1095
         End
         Begin VB.Label lblDesc 
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
            TabIndex        =   49
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
            Index           =   11
            Left            =   4140
            TabIndex        =   48
            Top             =   435
            Width           =   795
         End
      End
      Begin Threed.SSFrame SSFrame12 
         Height          =   1935
         Left            =   -74970
         TabIndex        =   50
         Top             =   570
         Width           =   9225
         _Version        =   65536
         _ExtentX        =   16272
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
            TabIndex        =   14
            Text            =   "txtData(12)"
            Top             =   360
            Width           =   1095
         End
         Begin VB.Label lblDesc 
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
            TabIndex        =   52
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
            Index           =   12
            Left            =   4140
            TabIndex        =   51
            Top             =   435
            Width           =   795
         End
      End
      Begin Threed.SSFrame SSFrame14 
         Height          =   1935
         Left            =   -74970
         TabIndex        =   53
         Top             =   2790
         Width           =   9225
         _Version        =   65536
         _ExtentX        =   16272
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
            TabIndex        =   15
            Text            =   "txtData(13)"
            Top             =   390
            Width           =   1095
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
            TabIndex        =   55
            Top             =   435
            Width           =   795
         End
         Begin VB.Label lblDesc 
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
            TabIndex        =   54
            Top             =   435
            Width           =   3000
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

   FRMTARGET_MODE = FRMTARGET_ADDNEW
  
    'INIT THIS TARGET COMPOUND.
    Call TargetCompound_SetDefaults(NowTargetCompound)
    NowTargetCompound.comname = "            "
    NowTargetCompound.cas = 0
  
    'SHOW THE FORM.
    frmTarget.Show 1

  'UPDATE MEMORY.
   If (Not USER_HIT_CANCEL) Then
    
    'ADD A NEW TARGET COMPOUND STRUCTURE.
    If NowProj.TargetCompounds(NowProj.TargetCompounds_Count).comname <> "" Then
      ReDim Preserve NowProj.TargetCompounds(1 To NowProj.TargetCompounds_Count)
      idx_new = NowProj.TargetCompounds_Count
      NowProj.TargetCompounds(idx_new) = NowTargetCompound
  
  '    'ADD A NEW SET OF EXTINCTION COEFFICIENTS (ONE PER WAVELENGTH).
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
  Else
    NowProj.TargetCompounds_Count = NowProj.TargetCompounds_Count - 1
  End If
  
  'RETURN TO MAIN WINDOW.
  frmTarget_DoAddNew = Not USER_HIT_CANCEL

End Function


'RETURNS:
'  TRUE = USER HIT OK
'  FALSE = USER HIT CANCEL
Public Function frmTarget_DoEdit(tcNum As Integer) As Integer
Dim is_aborted As Integer
Dim name_new As String
  
  'TELL FORM TO DO "EDIT" MODE.
  FRMTARGET_MODE = FRMTARGET_EDIT
  
  'IMPORT THIS TARGET COMPOUND FROM MEMORY TO THE FORM.
  NowTargetCompound = NowProj.TargetCompounds(tcNum)
  
  'SHOW THE FORM.
  frmTarget.Show 1
  
  'UPDATE MEMORY.
  If (Not USER_HIT_CANCEL) Then
    NowProj.TargetCompounds(tcNum) = NowTargetCompound
  End If
  
  'RETURN TO MAIN WINDOW.
  frmTarget_DoEdit = Not USER_HIT_CANCEL

End Function


Private Sub cmdCancel_Click()
  USER_HIT_CANCEL = True
  Unload Me
End Sub


Private Sub cmdDataBase_Click()
Dim RetVal As Boolean
Dim strOldComname As String
Dim txtctl As Control
    

'Call frmDatabase
'Display list of compounds with related CAS number,
    'moledular weight and second order rate constant
'Allow user to search by compound name, synonym, or CAS number
'When user selects compound and clicks on OK, fill fields
'   in frmTarget with values
'If user clicks Cancel, don't return values

Select Case FRMTARGET_MODE
Case FRMTARGET_ADDNEW
  Call frmChemDB.frmChemDB_IMPORT_MODE( _
    NowProj.TargetCompounds(NowProj.TargetCompounds_Count).comname)
  Set txtctl = txtDataStr(0)
  Call AssignTextAndTag(txtctl, txtctl.Text)
  NowTargetCompound.comname = txtctl.Text
  Set txtctl = txtData(15)
  Call AssignTextAndTag(txtctl, txtctl.Text)
  NowTargetCompound.cas = txtctl.Text
  Set txtctl = txtData(2)
  Call AssignTextAndTag(txtctl, txtctl.Text)
  NowTargetCompound.mw = txtctl.Text
  Set txtctl = txtData(5)
  Call AssignTextAndTag(txtctl, txtctl.Text)
  NowTargetCompound.xk = txtctl.Text
  Set txtctl = txtData(7)
  Call AssignTextAndTag(txtctl, txtctl.Text)
  NowTargetCompound.dep_mw = txtctl.Text
Case FRMTARGET_EDIT
  strOldComname = NowProj.TargetCompounds(frmMain.cboTarget.ListIndex + 1).comname
  Call frmChemDB.frmChemDB_EDIT_MODE( _
    NowProj.TargetCompounds(frmMain.cboTarget.ListIndex + 1). _
    comname)
  If strOldComname <> NowProj.TargetCompounds(frmMain.cboTarget.ListIndex + 1).comname Then
      Set txtctl = txtDataStr(0)
      Call AssignTextAndTag(txtctl, txtctl.Text)
      NowTargetCompound.comname = txtctl.Text
      Set txtctl = txtData(15)
      Call AssignTextAndTag(txtctl, txtctl.Text)
      NowTargetCompound.cas = txtctl.Text
      Set txtctl = txtData(2)
      Call AssignTextAndTag(txtctl, txtctl.Text)
      NowTargetCompound.mw = txtctl.Text
      Set txtctl = txtData(5)
      Call AssignTextAndTag(txtctl, txtctl.Text)
      NowTargetCompound.xk = txtctl.Text
      Set txtctl = txtData(7)
      Call AssignTextAndTag(txtctl, txtctl.Text)
      NowTargetCompound.dep_mw = txtctl.Text
  End If
End Select

End Sub

Private Sub cmdOK_Click()
  If Trim$(NowTargetCompound.comname) = "" Then
    Call Show_Message("Please enter new compound name", vbExclamation, App.title)
  Else
    USER_HIT_CANCEL = False
    Unload Me
  End If
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
      Call Show_Message("Please enter new compound name, CAS Number, " & _
        "or click on Use Database button", vbExclamation, App.title)
      txtDataStr(0).Enabled = True
      txtData(15).Enabled = True
      Call Local_AddNewCompound
    Case FRMTARGET_EDIT:
      Me.Caption = "Editing Target Compound "
      txtDataStr(0).Enabled = False
      txtData(15).Enabled = False
      If NowTargetCompound.comname = "NOM" Or _
        NowTargetCompound.comname = "R1" Then
          cmdDataBase.Enabled = False
      End If
      
     
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
              NowTargetCompound.dep_mw = newVal + NowTargetCompound.dep_val
      Case 3: NowTargetCompound.ncarbn = newVal
      Case 4: NowTargetCompound.nsubstt = newVal
      Case 5: NowTargetCompound.xk = newVal
      Case 6: NowTargetCompound.dep_xke = newVal
      Case 7: NowTargetCompound.dep_val = newVal
              NowTargetCompound.dep_mw = NowTargetCompound.mw + newVal
      Case 8: NowTargetCompound.dep_mw = newVal
      Case 9: NowTargetCompound.dep_xk = newVal
      Case 10: NowTargetCompound.xk_co3XM = newVal
      Case 11: NowTargetCompound.xk_hpo4XM = newVal
      Case 12: NowTargetCompound.xk_o2XM = newVal
      Case 13: NowTargetCompound.xk_ho2X = newVal
'      Case 14: NowTargetCompound.comname = newVal
      Case 15: NowTargetCompound.cas = newVal
    End Select
    
    If Index <> 7 Then
      Call AssignTextAndTag(txtctl, newVal)
    End If
    
    Select Case refresh_type
      Case 1:   'JUST THE MAIN WINDOW.
        Call refresh_frmTarget(NowTargetCompound)
    End Select
  End If
  Call Global_LostFocus(txtctl)
End Sub





Private Sub txtDataStr_GotFocus(Index As Integer)
  Dim txtctl As Control
  Set txtctl = txtDataStr(Index)
  Call DisplayDataEntryError
  Call Global_GotFocus(txtctl)
End Sub
Private Sub txtDataStr_KeyPress(Index As Integer, KeyAscii As Integer)
  KeyAscii = Global_TextKeyPress(KeyAscii)
End Sub
Private Sub txtDataStr_LostFocus(Index As Integer)
  Dim txtctl As Control
  Set txtctl = txtDataStr(Index)
  Dim ok_to_save As Integer
  Dim refresh_type As Integer
  ok_to_save = False
  If (txtctl.Text <> txtctl.Tag) Then
    ok_to_save = True
  End If
  If (ok_to_save) Then
    'DATA LOOKS OKAY, LET'S GO AHEAD AND SAVE IT.
      refresh_type = 1
      Select Case Index
        Case 0:
          Call AssignTextAndTag(txtctl, txtctl.Text)
          NowTargetCompound.comname = txtctl.Text
      End Select
    
    'THROW DIRTY FLAG, AND REFRESH EVERY WINDOWS.
'    Call DirtyFlag_Throw(TempProj)
    
    Select Case refresh_type
      Case 1:
        Call refresh_frmTarget(NowTargetCompound)
    End Select
  End If
  Call Global_LostFocus(txtctl)
End Sub
Sub Local_DirtyStatus_Set(DirtyFlag As Boolean, NewSetting As Boolean)
  Call Global_DirtyStatus_Set(Me, DirtyFlag, NewSetting)
End Sub

Sub Local_GenericStatus_Set(NewString As String)
  Call Global_GenericStatus_Set(Me, NewString)
End Sub


Sub Local_AddNewCompound()

Dim is_aborted As Integer
Dim name_new As String
Dim i As Integer
Dim j As Integer
Dim idx_new As Integer
Dim temp() As Double

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


End Sub
