VERSION 5.00
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "THREED32.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form frmD0_Props 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Chemical Data and Environmental Conditions (Physico-Chemical Properties)"
   ClientHeight    =   7290
   ClientLeft      =   1965
   ClientTop       =   1545
   ClientWidth     =   9480
   ControlBox      =   0   'False
   HelpContextID   =   11000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7290
   ScaleWidth      =   9480
   ShowInTaskbar   =   0   'False
   Begin TabDlg.SSTab SSTab1 
      Height          =   6705
      Left            =   90
      TabIndex        =   40
      Top             =   90
      Width           =   7695
      _ExtentX        =   13573
      _ExtentY        =   11827
      _Version        =   327681
      Style           =   1
      Tabs            =   2
      TabHeight       =   520
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      TabCaption(0)   =   "Environment and Contaminant"
      TabPicture(0)   =   "D0_Props.frx":0000
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "lblNote(3)"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "lblIndicateNote(3)"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "cmdImportSteppClipboard"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "SSFrame3"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "SSFrame4"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).ControlCount=   5
      TabCaption(1)   =   "Oxygen, Water, and Air"
      TabPicture(1)   =   "D0_Props.frx":001C
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "SSFrame6"
      Tab(1).Control(1)=   "SSFrame5"
      Tab(1).Control(2)=   "lblNote(2)"
      Tab(1).Control(3)=   "lblIndicateNote(2)"
      Tab(1).Control(4)=   "lblNote(1)"
      Tab(1).Control(5)=   "lblIndicateNote(1)"
      Tab(1).ControlCount=   6
      Begin Threed.SSFrame SSFrame6 
         Height          =   2895
         Left            =   -74880
         TabIndex        =   70
         Top             =   2190
         Width           =   7425
         _Version        =   65536
         _ExtentX        =   13097
         _ExtentY        =   5106
         _StockProps     =   14
         Caption         =   "Water and Air Properties:"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Begin VB.ComboBox cboSource 
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
            Index           =   18
            Left            =   5610
            Style           =   2  'Dropdown List
            TabIndex        =   119
            TabStop         =   0   'False
            Top             =   2340
            Width           =   1485
         End
         Begin VB.ComboBox cboSource 
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
            Index           =   17
            Left            =   5610
            Style           =   2  'Dropdown List
            TabIndex        =   117
            TabStop         =   0   'False
            Top             =   1950
            Width           =   1485
         End
         Begin VB.ComboBox cboSource 
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
            Index           =   16
            Left            =   5610
            Style           =   2  'Dropdown List
            TabIndex        =   115
            TabStop         =   0   'False
            Top             =   1560
            Width           =   1485
         End
         Begin VB.ComboBox cboSource 
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
            Index           =   15
            Left            =   5610
            Style           =   2  'Dropdown List
            TabIndex        =   113
            TabStop         =   0   'False
            Top             =   1170
            Width           =   1485
         End
         Begin VB.ComboBox cboSource 
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
            Index           =   14
            Left            =   5610
            Style           =   2  'Dropdown List
            TabIndex        =   111
            TabStop         =   0   'False
            Top             =   780
            Width           =   1485
         End
         Begin VB.ComboBox cboSource 
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
            Left            =   5610
            Style           =   2  'Dropdown List
            TabIndex        =   109
            TabStop         =   0   'False
            Top             =   390
            Width           =   1485
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
            Index           =   13
            Left            =   2250
            TabIndex        =   14
            Text            =   "txtData(13)"
            Top             =   405
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
            Index           =   13
            Left            =   3930
            Style           =   2  'Dropdown List
            TabIndex        =   75
            TabStop         =   0   'False
            Top             =   390
            Width           =   1290
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
            Index           =   14
            Left            =   2250
            TabIndex        =   15
            Text            =   "txtData(14)"
            Top             =   795
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
            Index           =   14
            Left            =   3930
            Style           =   2  'Dropdown List
            TabIndex        =   74
            TabStop         =   0   'False
            Top             =   780
            Width           =   1290
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
            Index           =   15
            Left            =   2250
            TabIndex        =   16
            Text            =   "txtData(15)"
            Top             =   1185
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
            Index           =   15
            Left            =   3930
            Style           =   2  'Dropdown List
            TabIndex        =   73
            TabStop         =   0   'False
            Top             =   1170
            Width           =   1290
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
            Index           =   16
            Left            =   2250
            TabIndex        =   17
            Text            =   "txtData(16)"
            Top             =   1575
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
            Index           =   17
            Left            =   2250
            TabIndex        =   18
            Text            =   "txtData(17)"
            Top             =   1965
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
            Index           =   17
            Left            =   3930
            Style           =   2  'Dropdown List
            TabIndex        =   72
            TabStop         =   0   'False
            Top             =   1950
            Width           =   1290
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
            Index           =   18
            Left            =   2250
            TabIndex        =   19
            Text            =   "txtData(18)"
            Top             =   2355
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
            Index           =   18
            Left            =   3930
            Style           =   2  'Dropdown List
            TabIndex        =   71
            TabStop         =   0   'False
            Top             =   2340
            Width           =   1290
         End
         Begin VB.Label lblIndicate 
            Alignment       =   2  'Center
            Caption         =   "*"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   24
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H0000C000&
            Height          =   315
            Index           =   18
            Left            =   5250
            TabIndex        =   120
            Top             =   2340
            Width           =   315
         End
         Begin VB.Label lblIndicate 
            Alignment       =   2  'Center
            Caption         =   "*"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   24
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H0000C000&
            Height          =   315
            Index           =   17
            Left            =   5250
            TabIndex        =   118
            Top             =   1950
            Width           =   315
         End
         Begin VB.Label lblIndicate 
            Alignment       =   2  'Center
            Caption         =   "*"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   24
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000FF&
            Height          =   315
            Index           =   16
            Left            =   5250
            TabIndex        =   116
            Top             =   1560
            Width           =   315
         End
         Begin VB.Label lblIndicate 
            Alignment       =   2  'Center
            Caption         =   "*"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   24
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H0000C000&
            Height          =   315
            Index           =   15
            Left            =   5250
            TabIndex        =   114
            Top             =   1170
            Width           =   315
         End
         Begin VB.Label lblIndicate 
            Alignment       =   2  'Center
            Caption         =   "*"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   24
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H0000C000&
            Height          =   315
            Index           =   14
            Left            =   5250
            TabIndex        =   112
            Top             =   780
            Width           =   315
         End
         Begin VB.Label lblIndicate 
            Alignment       =   2  'Center
            Caption         =   "*"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   24
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H0000C000&
            Height          =   315
            Index           =   13
            Left            =   5250
            TabIndex        =   110
            Top             =   390
            Width           =   315
         End
         Begin VB.Label lblDataUnits 
            Alignment       =   2  'Center
            Caption         =   "(dim'less)"
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
            Left            =   3930
            TabIndex        =   82
            Top             =   1620
            Width           =   1290
         End
         Begin VB.Label lblData 
            Alignment       =   1  'Right Justify
            Caption         =   "Density, H2O:"
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
            Left            =   180
            TabIndex        =   81
            Top             =   450
            Width           =   2025
         End
         Begin VB.Label lblData 
            Alignment       =   1  'Right Justify
            Caption         =   "Viscosity, H2O:"
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
            Index           =   14
            Left            =   180
            TabIndex        =   80
            Top             =   840
            Width           =   2025
         End
         Begin VB.Label lblData 
            Alignment       =   1  'Right Justify
            Caption         =   "Vapor Pressure, H2O:"
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
            Index           =   15
            Left            =   180
            TabIndex        =   79
            Top             =   1230
            Width           =   2025
         End
         Begin VB.Label lblData 
            Alignment       =   1  'Right Justify
            Caption         =   "Alpha:"
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
            Index           =   16
            Left            =   180
            TabIndex        =   78
            Top             =   1620
            Width           =   2025
         End
         Begin VB.Label lblData 
            Alignment       =   1  'Right Justify
            Caption         =   "Density, air:"
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
            Index           =   17
            Left            =   180
            TabIndex        =   77
            Top             =   2010
            Width           =   2025
         End
         Begin VB.Label lblData 
            Alignment       =   1  'Right Justify
            Caption         =   "Viscosity, air:"
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
            Index           =   18
            Left            =   180
            TabIndex        =   76
            Top             =   2400
            Width           =   2025
         End
      End
      Begin Threed.SSFrame SSFrame5 
         Height          =   1725
         Left            =   -74880
         TabIndex        =   63
         Top             =   420
         Width           =   7425
         _Version        =   65536
         _ExtentX        =   13097
         _ExtentY        =   3043
         _StockProps     =   14
         Caption         =   "Oxygen Properties:"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Begin VB.ComboBox cboSource 
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
            Left            =   5610
            Style           =   2  'Dropdown List
            TabIndex        =   107
            TabStop         =   0   'False
            Top             =   1170
            Width           =   1485
         End
         Begin VB.ComboBox cboSource 
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
            Left            =   5610
            Style           =   2  'Dropdown List
            TabIndex        =   105
            TabStop         =   0   'False
            Top             =   780
            Width           =   1485
         End
         Begin VB.ComboBox cboSource 
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
            Left            =   5610
            Style           =   2  'Dropdown List
            TabIndex        =   103
            TabStop         =   0   'False
            Top             =   390
            Width           =   1485
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
            Index           =   10
            Left            =   2250
            TabIndex        =   11
            Text            =   "txtData(10)"
            Top             =   405
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
            Left            =   3930
            Style           =   2  'Dropdown List
            TabIndex        =   65
            TabStop         =   0   'False
            Top             =   390
            Width           =   1290
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
            Index           =   11
            Left            =   2250
            TabIndex        =   12
            Text            =   "txtData(11)"
            Top             =   795
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
            Index           =   12
            Left            =   2250
            TabIndex        =   13
            Text            =   "txtData(12)"
            Top             =   1185
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
            Index           =   12
            Left            =   3930
            Style           =   2  'Dropdown List
            TabIndex        =   64
            TabStop         =   0   'False
            Top             =   1170
            Width           =   1290
         End
         Begin VB.Label lblIndicate 
            Alignment       =   2  'Center
            Caption         =   "*"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   24
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H0000C000&
            Height          =   315
            Index           =   12
            Left            =   5250
            TabIndex        =   108
            Top             =   1170
            Width           =   315
         End
         Begin VB.Label lblIndicate 
            Alignment       =   2  'Center
            Caption         =   "*"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   24
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H0000C000&
            Height          =   315
            Index           =   11
            Left            =   5250
            TabIndex        =   106
            Top             =   780
            Width           =   315
         End
         Begin VB.Label lblIndicate 
            Alignment       =   2  'Center
            Caption         =   "*"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   24
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H0000C000&
            Height          =   315
            Index           =   10
            Left            =   5250
            TabIndex        =   104
            Top             =   390
            Width           =   315
         End
         Begin VB.Label lblDataUnits 
            Alignment       =   2  'Center
            Caption         =   "(dim'less)"
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
            Left            =   3930
            TabIndex        =   69
            Top             =   840
            Width           =   1290
         End
         Begin VB.Label lblData 
            Alignment       =   1  'Right Justify
            Caption         =   "Sat'n Conc., O2:"
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
            Left            =   180
            TabIndex        =   68
            Top             =   450
            Width           =   2025
         End
         Begin VB.Label lblData 
            Alignment       =   1  'Right Justify
            Caption         =   "Henry's Const., O2:"
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
            Left            =   180
            TabIndex        =   67
            Top             =   840
            Width           =   2025
         End
         Begin VB.Label lblData 
            Alignment       =   1  'Right Justify
            Caption         =   "Diffusivity, O2:"
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
            Left            =   180
            TabIndex        =   66
            Top             =   1230
            Width           =   2025
         End
      End
      Begin Threed.SSFrame SSFrame4 
         Height          =   3645
         Left            =   120
         TabIndex        =   48
         Top             =   2190
         Width           =   7425
         _Version        =   65536
         _ExtentX        =   13097
         _ExtentY        =   6429
         _StockProps     =   14
         Caption         =   "Contaminant Properties:"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Begin VB.TextBox txtDataStr 
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
            Index           =   0
            Left            =   2250
            TabIndex        =   3
            Text            =   "txtDataStr(0)"
            Top             =   405
            Width           =   3195
         End
         Begin VB.ComboBox cboSource 
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
            Left            =   5610
            Style           =   2  'Dropdown List
            TabIndex        =   101
            TabStop         =   0   'False
            Top             =   3120
            Width           =   1485
         End
         Begin VB.ComboBox cboSource 
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
            Left            =   5610
            Style           =   2  'Dropdown List
            TabIndex        =   99
            TabStop         =   0   'False
            Top             =   2730
            Width           =   1485
         End
         Begin VB.ComboBox cboSource 
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
            Left            =   5610
            Style           =   2  'Dropdown List
            TabIndex        =   97
            TabStop         =   0   'False
            Top             =   2340
            Width           =   1485
         End
         Begin VB.ComboBox cboSource 
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
            Left            =   5610
            Style           =   2  'Dropdown List
            TabIndex        =   95
            TabStop         =   0   'False
            Top             =   1950
            Width           =   1485
         End
         Begin VB.ComboBox cboSource 
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
            Left            =   5610
            Style           =   2  'Dropdown List
            TabIndex        =   93
            TabStop         =   0   'False
            Top             =   1560
            Width           =   1485
         End
         Begin VB.ComboBox cboSource 
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
            Left            =   5610
            Style           =   2  'Dropdown List
            TabIndex        =   91
            TabStop         =   0   'False
            Top             =   1170
            Width           =   1485
         End
         Begin VB.ComboBox cboSource 
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
            Left            =   5610
            Style           =   2  'Dropdown List
            TabIndex        =   89
            TabStop         =   0   'False
            Top             =   780
            Width           =   1485
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
            Index           =   9
            Left            =   3930
            Style           =   2  'Dropdown List
            TabIndex        =   53
            TabStop         =   0   'False
            Top             =   3120
            Width           =   1290
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
            Index           =   9
            Left            =   2250
            TabIndex        =   10
            Text            =   "txtData(9)"
            Top             =   3135
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
            Index           =   8
            Left            =   3930
            Style           =   2  'Dropdown List
            TabIndex        =   52
            TabStop         =   0   'False
            Top             =   2730
            Width           =   1290
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
            Index           =   8
            Left            =   2250
            TabIndex        =   9
            Text            =   "txtData(8)"
            Top             =   2745
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
            Index           =   7
            Left            =   3930
            Style           =   2  'Dropdown List
            TabIndex        =   51
            TabStop         =   0   'False
            Top             =   2340
            Width           =   1290
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
            Left            =   2250
            TabIndex        =   8
            Text            =   "txtData(7)"
            Top             =   2355
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
            Left            =   2250
            TabIndex        =   7
            Text            =   "txtData(6)"
            Top             =   1965
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
            Index           =   5
            Left            =   2250
            TabIndex        =   6
            Text            =   "txtData(5)"
            Top             =   1575
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
            Left            =   3930
            Style           =   2  'Dropdown List
            TabIndex        =   50
            TabStop         =   0   'False
            Top             =   1170
            Width           =   1290
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
            Left            =   2250
            TabIndex        =   5
            Text            =   "txtData(4)"
            Top             =   1185
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
            Left            =   3930
            Style           =   2  'Dropdown List
            TabIndex        =   49
            TabStop         =   0   'False
            Top             =   780
            Width           =   1290
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
            Left            =   2250
            TabIndex        =   4
            Text            =   "txtData(3)"
            Top             =   795
            Width           =   1635
         End
         Begin VB.Label lblDataStr 
            Alignment       =   1  'Right Justify
            Caption         =   "Contaminant Name:"
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
            TabIndex        =   121
            Top             =   450
            Width           =   2025
         End
         Begin VB.Label lblIndicate 
            Alignment       =   2  'Center
            Caption         =   "*"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   24
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FF0000&
            Height          =   315
            Index           =   9
            Left            =   5250
            TabIndex        =   102
            Top             =   3120
            Width           =   315
         End
         Begin VB.Label lblIndicate 
            Alignment       =   2  'Center
            Caption         =   "*"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   24
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FF0000&
            Height          =   315
            Index           =   8
            Left            =   5250
            TabIndex        =   100
            Top             =   2730
            Width           =   315
         End
         Begin VB.Label lblIndicate 
            Alignment       =   2  'Center
            Caption         =   "*"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   24
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   315
            Index           =   7
            Left            =   5250
            TabIndex        =   98
            Top             =   2340
            Visible         =   0   'False
            Width           =   315
         End
         Begin VB.Label lblIndicate 
            Alignment       =   2  'Center
            Caption         =   "*"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   24
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FF0000&
            Height          =   315
            Index           =   6
            Left            =   5250
            TabIndex        =   96
            Top             =   1950
            Width           =   315
         End
         Begin VB.Label lblIndicate 
            Alignment       =   2  'Center
            Caption         =   "*"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   24
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FF0000&
            Height          =   315
            Index           =   5
            Left            =   5250
            TabIndex        =   94
            Top             =   1560
            Width           =   315
         End
         Begin VB.Label lblIndicate 
            Alignment       =   2  'Center
            Caption         =   "*"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   24
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   315
            Index           =   4
            Left            =   5250
            TabIndex        =   92
            Top             =   1170
            Visible         =   0   'False
            Width           =   315
         End
         Begin VB.Label lblIndicate 
            Alignment       =   2  'Center
            Caption         =   "*"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   24
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   315
            Index           =   3
            Left            =   5250
            TabIndex        =   90
            Top             =   780
            Visible         =   0   'False
            Width           =   315
         End
         Begin VB.Label lblDataUnits 
            Alignment       =   2  'Center
            Caption         =   "(dim'less)"
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
            Left            =   3930
            TabIndex        =   62
            Top             =   2010
            Width           =   1290
         End
         Begin VB.Label lblDataUnits 
            Alignment       =   2  'Center
            Caption         =   "(dim'less)"
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
            Left            =   3930
            TabIndex        =   61
            Top             =   1620
            Width           =   1290
         End
         Begin VB.Label lblData 
            Alignment       =   1  'Right Justify
            Caption         =   "Air Diffusivity, VOC:"
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
            Left            =   180
            TabIndex        =   60
            Top             =   3180
            Width           =   2025
         End
         Begin VB.Label lblData 
            Alignment       =   1  'Right Justify
            Caption         =   "Water Diffusivity, VOC:"
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
            Left            =   180
            TabIndex        =   59
            Top             =   2790
            Width           =   2025
         End
         Begin VB.Label lblData 
            Alignment       =   1  'Right Justify
            Caption         =   "Molec. Weight, VOC:"
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
            Left            =   180
            TabIndex        =   58
            Top             =   2400
            Width           =   2025
         End
         Begin VB.Label lblData 
            Alignment       =   1  'Right Justify
            Caption         =   "Henry's Const., VOC:"
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
            Left            =   180
            TabIndex        =   57
            Top             =   2010
            Width           =   2025
         End
         Begin VB.Label lblData 
            Alignment       =   1  'Right Justify
            Caption         =   "Log Kow, VOC:"
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
            Left            =   180
            TabIndex        =   56
            Top             =   1620
            Width           =   2025
         End
         Begin VB.Label lblData 
            Alignment       =   1  'Right Justify
            Caption         =   "Biodeg. Rate, VOC:"
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
            Left            =   180
            TabIndex        =   55
            Top             =   1230
            Width           =   2025
         End
         Begin VB.Label lblData 
            Alignment       =   1  'Right Justify
            Caption         =   "Influent Conc., VOC:"
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
            Left            =   180
            TabIndex        =   54
            Top             =   840
            Width           =   2025
         End
      End
      Begin Threed.SSFrame SSFrame3 
         Height          =   1725
         Left            =   120
         TabIndex        =   41
         Top             =   420
         Width           =   7425
         _Version        =   65536
         _ExtentX        =   13097
         _ExtentY        =   3043
         _StockProps     =   14
         Caption         =   "Environmental Conditions:"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Begin VB.ComboBox cboSource 
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
            Left            =   5610
            Style           =   2  'Dropdown List
            TabIndex        =   87
            TabStop         =   0   'False
            Top             =   1170
            Width           =   1485
         End
         Begin VB.ComboBox cboSource 
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
            Left            =   5610
            Style           =   2  'Dropdown List
            TabIndex        =   85
            TabStop         =   0   'False
            Top             =   780
            Width           =   1485
         End
         Begin VB.ComboBox cboSource 
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
            Left            =   5610
            Style           =   2  'Dropdown List
            TabIndex        =   83
            TabStop         =   0   'False
            Top             =   390
            Width           =   1485
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
            Left            =   3930
            Style           =   2  'Dropdown List
            TabIndex        =   44
            TabStop         =   0   'False
            Top             =   1170
            Width           =   1290
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
            Left            =   2250
            TabIndex        =   2
            Text            =   "txtData(2)"
            Top             =   1185
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
            Index           =   1
            Left            =   3930
            Style           =   2  'Dropdown List
            TabIndex        =   43
            TabStop         =   0   'False
            Top             =   780
            Width           =   1290
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
            Left            =   2250
            TabIndex        =   1
            Text            =   "txtData(1)"
            Top             =   795
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
            Index           =   0
            Left            =   3930
            Style           =   2  'Dropdown List
            TabIndex        =   42
            TabStop         =   0   'False
            Top             =   390
            Width           =   1290
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
            Index           =   0
            Left            =   2250
            TabIndex        =   0
            Text            =   "txtData(0)"
            Top             =   405
            Width           =   1635
         End
         Begin VB.Label lblIndicate 
            Alignment       =   2  'Center
            Caption         =   "*"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   24
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   315
            Index           =   2
            Left            =   5250
            TabIndex        =   88
            Top             =   1170
            Visible         =   0   'False
            Width           =   315
         End
         Begin VB.Label lblIndicate 
            Alignment       =   2  'Center
            Caption         =   "*"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   24
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   315
            Index           =   1
            Left            =   5250
            TabIndex        =   86
            Top             =   780
            Visible         =   0   'False
            Width           =   315
         End
         Begin VB.Label lblIndicate 
            Alignment       =   2  'Center
            Caption         =   "*"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   24
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   315
            Index           =   0
            Left            =   5250
            TabIndex        =   84
            Top             =   390
            Visible         =   0   'False
            Width           =   315
         End
         Begin VB.Label lblData 
            Alignment       =   1  'Right Justify
            Caption         =   "Wind Velocity:"
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
            Left            =   180
            TabIndex        =   47
            Top             =   1230
            Width           =   2025
         End
         Begin VB.Label lblData 
            Alignment       =   1  'Right Justify
            Caption         =   "Temperature:"
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
            Left            =   180
            TabIndex        =   46
            Top             =   840
            Width           =   2025
         End
         Begin VB.Label lblData 
            Alignment       =   1  'Right Justify
            Caption         =   "Barometric Pressure:"
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
            TabIndex        =   45
            Top             =   450
            Width           =   2025
         End
      End
      Begin Threed.SSCommand cmdImportSteppClipboard 
         Height          =   495
         Left            =   120
         TabIndex        =   126
         TabStop         =   0   'False
         Top             =   6000
         Width           =   3885
         _Version        =   65536
         _ExtentX        =   6853
         _ExtentY        =   873
         _StockProps     =   78
         Caption         =   "Import Properties from StEPP via Clipboard"
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
      Begin VB.Label lblIndicateNote 
         Alignment       =   2  'Center
         Caption         =   "*"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   24
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   315
         Index           =   3
         Left            =   4680
         TabIndex        =   128
         Top             =   6060
         Width           =   315
      End
      Begin VB.Label lblNote 
         Caption         =   "Temperature-dependent; re-import from StEPP if temperature changes."
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Index           =   3
         Left            =   5040
         TabIndex        =   127
         Top             =   5880
         Width           =   2490
      End
      Begin VB.Label lblNote 
         Caption         =   $"D0_Props.frx":0038
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   825
         Index           =   2
         Left            =   -74580
         TabIndex        =   125
         Top             =   5670
         Width           =   6960
      End
      Begin VB.Label lblIndicateNote 
         Alignment       =   2  'Center
         Caption         =   "*"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   24
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000C000&
         Height          =   315
         Index           =   2
         Left            =   -74880
         TabIndex        =   124
         Top             =   5640
         Width           =   315
      End
      Begin VB.Label lblNote 
         Caption         =   "Temperature-dependent; re-enter if temperature changes."
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
         Left            =   -74580
         TabIndex        =   123
         Top             =   5370
         Width           =   7200
      End
      Begin VB.Label lblIndicateNote 
         Alignment       =   2  'Center
         Caption         =   "*"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   24
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   315
         Index           =   1
         Left            =   -74880
         TabIndex        =   122
         Top             =   5340
         Width           =   315
      End
   End
   Begin Threed.SSFrame SSFrame1 
      Height          =   3795
      Left            =   3630
      TabIndex        =   26
      Top             =   6870
      Visible         =   0   'False
      Width           =   6975
      _Version        =   65536
      _ExtentX        =   12303
      _ExtentY        =   6694
      _StockProps     =   14
      Caption         =   "Old Stuff"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   9.75
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
         Index           =   16
         Left            =   600
         Style           =   2  'Dropdown List
         TabIndex        =   39
         TabStop         =   0   'False
         Top             =   3240
         Width           =   1290
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
         Index           =   11
         Left            =   570
         Style           =   2  'Dropdown List
         TabIndex        =   38
         TabStop         =   0   'False
         Top             =   2850
         Width           =   1290
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
         Left            =   540
         Style           =   2  'Dropdown List
         TabIndex        =   37
         TabStop         =   0   'False
         Top             =   2520
         Width           =   1290
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
         Left            =   540
         Style           =   2  'Dropdown List
         TabIndex        =   36
         TabStop         =   0   'False
         Top             =   2130
         Width           =   1290
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
         Index           =   102
         Left            =   5085
         Style           =   2  'Dropdown List
         TabIndex        =   34
         TabStop         =   0   'False
         Top             =   1140
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
         Height          =   300
         Index           =   102
         Left            =   3015
         TabIndex        =   33
         Text            =   "txtData(2)"
         Top             =   1170
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
         Index           =   100
         Left            =   3015
         TabIndex        =   30
         Text            =   "txtData(0)"
         Top             =   330
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
         Index           =   100
         Left            =   5085
         Style           =   2  'Dropdown List
         TabIndex        =   29
         TabStop         =   0   'False
         Top             =   300
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
         Index           =   101
         Left            =   5085
         Style           =   2  'Dropdown List
         TabIndex        =   28
         TabStop         =   0   'False
         Top             =   720
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
         Height          =   300
         Index           =   101
         Left            =   3015
         TabIndex        =   27
         Text            =   "txtData(1)"
         Top             =   750
         Width           =   1995
      End
      Begin VB.Label lblData 
         Alignment       =   1  'Right Justify
         Caption         =   "Gas Flow Rate:"
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
         Index           =   102
         Left            =   90
         TabIndex        =   35
         Top             =   1200
         Width           =   2805
      End
      Begin VB.Label lblData 
         Alignment       =   1  'Right Justify
         Caption         =   "Width:"
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
         Left            =   90
         TabIndex        =   32
         Top             =   360
         Width           =   2805
      End
      Begin VB.Label lblData 
         Alignment       =   1  'Right Justify
         Caption         =   "Distance betw. Water Levels:"
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
         Index           =   101
         Left            =   90
         TabIndex        =   31
         Top             =   780
         Width           =   2805
      End
   End
   Begin Threed.SSPanel SSPanel2 
      Align           =   2  'Align Bottom
      Height          =   405
      Left            =   0
      TabIndex        =   20
      Top             =   6885
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
         TabIndex        =   21
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
         TabIndex        =   22
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
      Left            =   8040
      TabIndex        =   23
      TabStop         =   0   'False
      ToolTipText     =   "Click here to save the changes to this window"
      Top             =   600
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
      Left            =   8040
      TabIndex        =   24
      TabStop         =   0   'False
      ToolTipText     =   "Click here to abandon any changes on this window"
      Top             =   120
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
      Left            =   8040
      TabIndex        =   25
      TabStop         =   0   'False
      ToolTipText     =   "Click here for help"
      Top             =   1260
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
End
Attribute VB_Name = "frmD0_Props"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim USER_HIT_CANCEL As Boolean
Dim USER_HIT_OK As Boolean
Dim frmD0_Props_Is_Dirty As Boolean

Dim Temp_Plant As TYPE_PlantDiagram

Public HALT_cboSource As Boolean



Const frmD0_Props_declarations_end = True


Sub frmD0_Props_Edit( _
    OUTPUT_Raise_Dirty_Flag As Boolean)
  Temp_Plant = NowProj.Plant
  frmD0_Props.Show 1
  If (USER_HIT_OK) Then
    OUTPUT_Raise_Dirty_Flag = True
    NowProj.Plant = Temp_Plant
  Else
    OUTPUT_Raise_Dirty_Flag = False
  End If
End Sub


Sub Populate_cboSource0( _
    idx_This As Integer, _
    SelType As String)
Dim Ctl As Control
Set Ctl = cboSource(idx_This)
Dim iThis As Integer
Dim i As Integer
Dim IsEnabled As Boolean
Dim IsLocked As Boolean
  Ctl.Clear
  For i = 1 To Len(SelType)
    Select Case UCase$(Mid$(SelType, i, 1))
      Case "U": Ctl.AddItem "User Entry": iThis = DATASOURCETYPE_USERINPUT
      Case "S": Ctl.AddItem "StEPP": iThis = DATASOURCETYPE_STEPP
      Case "C": Ctl.AddItem "Correlation": iThis = DATASOURCETYPE_CORR
    End Select
    Ctl.ItemData(Ctl.NewIndex) = iThis
  Next i
  IsEnabled = IIf(Len(SelType) = 1, False, True)
  Ctl.Enabled = IsEnabled
  IsLocked = Not IsEnabled
  ''''Ctl.Visible = IsEnabled
  Ctl.BackColor = _
      IIf(IsLocked, QBColor(7), QBColor(15))
  Ctl.ForeColor = _
      IIf(IsLocked, QBColor(8), QBColor(0))
End Sub
Sub Populate_cboSource()
  Call Populate_cboSource0(0, "U")
  Call Populate_cboSource0(1, "U")
  Call Populate_cboSource0(2, "U")
  Call Populate_cboSource0(3, "U")
  Call Populate_cboSource0(4, "U")
  Call Populate_cboSource0(5, "SU")
  Call Populate_cboSource0(6, "SU")
  Call Populate_cboSource0(7, "SU")
  Call Populate_cboSource0(8, "SU")
  Call Populate_cboSource0(9, "SU")
  Call Populate_cboSource0(10, "CU")
  Call Populate_cboSource0(11, "CU")
  Call Populate_cboSource0(12, "CU")
  Call Populate_cboSource0(13, "CU")
  Call Populate_cboSource0(14, "CU")
  Call Populate_cboSource0(15, "CU")
  Call Populate_cboSource0(16, "U")
  Call Populate_cboSource0(17, "CU")
  Call Populate_cboSource0(18, "CU")
End Sub


Sub frmD0_PopUnits0( _
    idx_This As Integer, _
    Base_Units As String, _
    Unit_Type As String, _
    Format_Entry As String, _
    Format_Display As String, _
    Has_Units As Boolean)
Dim Frm As Form
Set Frm = frmD0_Props
Dim CboX As Control
  '
  ' REGISTER THE APPROPRIATE UNIT CONTROL.
  '
  If (Has_Units) Then
    Set CboX = cboUnits(idx_This)
  Else
    Set CboX = Nothing
  End If
  Call unitsys_register( _
      Frm, _
      lblData(idx_This), _
      txtData(idx_This), _
      CboX, _
      Unit_Type, _
      Temp_Plant.ChemicalData.UnitsOfDisplay(idx_This), _
      Base_Units, _
      Format_Entry, _
      Format_Display, _
      100#, _
      CInt(Has_Units))
End Sub
Sub frmD0_Props_PopulateUnits()
Dim Frm As Form
Set Frm = frmD0_Props
Dim i As Integer
  '
  ' MAIN DATA BLOCK.
  '
  Call frmD0_PopUnits0(0, "kPa", "pressure", "", "", True)
  Call frmD0_PopUnits0(1, "C", "temperature", "", "", True)
  Call frmD0_PopUnits0(2, "m/s", "velocity", "", "", True)
  Call frmD0_PopUnits0(3, "µg/L", "concentration", "", "", True)
  Call frmD0_PopUnits0(4, "L/mg-d", "biodegredation_rate", "", "", True)
  Call frmD0_PopUnits0(5, "", "", "", "", False)
  Call frmD0_PopUnits0(6, "", "", "", "", False)
  Call frmD0_PopUnits0(7, "g/gmol", "molecular_weight", "", "", True)
  Call frmD0_PopUnits0(8, "cm²/s", "diffusivity", "", "", True)
  Call frmD0_PopUnits0(9, "cm²/s", "diffusivity", "", "", True)
  Call frmD0_PopUnits0(10, "mg/L", "concentration", "", "", True)
  Call frmD0_PopUnits0(11, "", "", "", "", False)
  Call frmD0_PopUnits0(12, "cm²/s", "diffusivity", "", "", True)
  Call frmD0_PopUnits0(13, "kg/m³", "density", "", "", True)
  Call frmD0_PopUnits0(14, "kg/m-s", "viscosity", "", "", True)
  Call frmD0_PopUnits0(15, "kPa", "pressure", "", "", True)
  Call frmD0_PopUnits0(16, "", "", "", "", False)
  Call frmD0_PopUnits0(17, "kg/m³", "density", "", "", True)
  Call frmD0_PopUnits0(18, "kg/m-s", "viscosity", "", "", True)
End Sub
Sub Store_Unit_Settings()
Dim i As Integer
  With Temp_Plant.ChemicalData
    For i = 0 To 18
      .UnitsOfDisplay(i) = unitsys_get_units(cboUnits(i))
    Next i
  End With
End Sub


Sub frmD0_TransProp( _
    idx_This As Integer, _
    OUT_SetVal As Double)
Dim ThisVal As Double
  With Temp_Plant.ChemicalData.DataSources(idx_This)
    Select Case .SourceType
      Case DATASOURCETYPE_USERINPUT: ThisVal = .Val_UserInput
      Case DATASOURCETYPE_STEPP: ThisVal = .Val_StEPP
      Case DATASOURCETYPE_CORR: ThisVal = .Val_Corr
    End Select
    OUT_SetVal = ThisVal
  End With
End Sub
Function Transfer_DataSources_Variables() As Boolean
Dim Ctl As Control
Dim ThisVal As Double
Dim i As Integer
Dim Is_Unavailable As Boolean
  '
  ' CHECK FOR "UNAVAILABLE" PROPERTIES.
  '
  For i = 0 To 18
    With Temp_Plant.ChemicalData.DataSources(i)
      Select Case .SourceType
        Case DATASOURCETYPE_USERINPUT: ThisVal = .Val_UserInput
        Case DATASOURCETYPE_STEPP: ThisVal = .Val_StEPP
        Case DATASOURCETYPE_CORR: ThisVal = .Val_Corr
      End Select
      Is_Unavailable = IIf(ThisVal <= -1E+20, True, False)
      If (Is_Unavailable = True) Then
        Call Show_Error("At least one property is marked " & _
            "as Unavailable.  Either enter a value for each " & _
            "Unavailable property and hit OK again, or " & _
            "abandon all of your changes by hitting Cancel.")
        Transfer_DataSources_Variables = False
        Exit Function
      End If
    End With
  Next i
  '
  ' TRANSFER THE ACTUAL VALUES.
  '
  With Temp_Plant.ChemicalData
    Call frmD0_TransProp(0, .env_Pressure)
    Call frmD0_TransProp(1, .env_Temperature)
    Call frmD0_TransProp(2, .env_WindVelocity)
    Call frmD0_TransProp(3, .InfluentConc)
    Call frmD0_TransProp(4, .BiodegredationRate)
    Call frmD0_TransProp(5, .LogKow)
    Call frmD0_TransProp(6, .VOC_HenrysConstant)
    Call frmD0_TransProp(7, .VOC_MolecularWeight)
    Call frmD0_TransProp(8, .VOC_DiffusivityInH2O)
    Call frmD0_TransProp(9, .VOC_DiffusivityInGas)
    Call frmD0_TransProp(10, .O2_SaturationConc)
    Call frmD0_TransProp(11, .O2_HenrysConstant)
    Call frmD0_TransProp(12, .O2_Diffusivity)
    Call frmD0_TransProp(13, .H2O_Density)
    Call frmD0_TransProp(14, .H2O_Viscosity)
    Call frmD0_TransProp(15, .H2O_VaporPressure)
    Call frmD0_TransProp(16, .H2O_Alpha)
    Call frmD0_TransProp(17, .AIR_Density)
    Call frmD0_TransProp(18, .AIR_Viscosity)
  End With
  Transfer_DataSources_Variables = True
End Function


Sub Local_DirtyStatus_Set(DirtyFlag As Boolean, NewSetting As Boolean)
  Call Global_DirtyStatus_Set(Me, DirtyFlag, NewSetting)
End Sub
Sub Local_GenericStatus_Set(NewString As String)
  Call Global_GenericStatus_Set(Me, NewString)
End Sub


Private Sub cboSource_Click(Index As Integer)
Dim Ctl As Control
Set Ctl = cboSource(Index)
  If (HALT_cboSource = True) Then Exit Sub
  If (Val(Ctl.Tag) = Ctl.ListIndex) Then Exit Sub
  With Temp_Plant.ChemicalData.DataSources(Index)
    .SourceType = Ctl.ItemData(Ctl.ListIndex)
  End With
  'RAISE DIRTY FLAG AND REFRESH WINDOW.
  Call Local_DirtyStatus_Set(frmD0_Props_Is_Dirty, True)
  Call frmD0_Props_Refresh(Temp_Plant)
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
      USER_HIT_CANCEL = True
      USER_HIT_OK = False
      Unload Me
      Exit Sub
    Case 1:     'OK.
      '
      ' TRANSFER ALL DataSources() DATA INTO
      ' VARIABLES USED BY REST OF PROGRAM.
      '
      If (Transfer_DataSources_Variables() = False) Then
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
Private Sub cmdImportSteppClipboard_Click()
Dim Was_Aborted As Boolean
  Call Do_ImportClipboard( _
    Was_Aborted, _
    Temp_Plant)
  If (Was_Aborted) Then
    Exit Sub
  Else
    '
    ' RAISE DIRTY FLAG AND REFRESH WINDOW.
    '
    Call Local_DirtyStatus_Set(frmD0_Props_Is_Dirty, True)
    Call frmD0_Props_Refresh(Temp_Plant)
  End If
End Sub


Private Sub Form_Load()
  '
  ' MISC INITS.
  '
  Call CenterOnForm(Me, frmMain)
  Call Local_DirtyStatus_Set(frmD0_Props_Is_Dirty, False)
  Call Local_GenericStatus_Set("")
  HALT_cboSource = False
  Call Populate_cboSource
  '
  ' POPULATE UNIT CONTROLS.
  '
  Call frmD0_Props_PopulateUnits
  '
  ' REFRESH DISPLAY.
  '
  Call frmD0_Props_Refresh(Temp_Plant)
End Sub
Private Sub Form_Unload(Cancel As Integer)
  Call unitsys_unregister_all_on_form(Me)
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
    Case 0: Val_Low = 1E-20: Val_High = 1E+20
    Case 1: Val_Low = 1E-20: Val_High = 1E+20
    Case 2: Val_Low = 1E-20: Val_High = 1E+20
    Case 3: Val_Low = 1E-20: Val_High = 1E+20
    Case 4: Val_Low = 1E-20: Val_High = 1E+20
    Case 5: Val_Low = 1E-20: Val_High = 1E+20
    Case 6: Val_Low = 1E-20: Val_High = 1E+20
    Case 7: Val_Low = 1E-20: Val_High = 1E+20
    Case 8: Val_Low = 1E-20: Val_High = 1E+20
    Case 9: Val_Low = 1E-20: Val_High = 1E+20
    Case 10: Val_Low = 1E-20: Val_High = 1E+20
    Case 11: Val_Low = 1E-20: Val_High = 1E+20
    Case 12: Val_Low = 1E-20: Val_High = 1E+20
    Case 13: Val_Low = 1E-20: Val_High = 1E+20
    Case 14: Val_Low = 1E-20: Val_High = 1E+20
    Case 15: Val_Low = 1E-20: Val_High = 1E+20
    Case 16: Val_Low = 1E-20: Val_High = 1E+20
    Case 17: Val_Low = 1E-20: Val_High = 1E+20
    Case 18: Val_Low = 1E-20: Val_High = 1E+20
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
      Select Case Index
        '
        ' MAIN DATA BLOCK.
        '
        Case -1:
        Case Else:
          With Temp_Plant.ChemicalData.DataSources(Index)
            Select Case .SourceType
              Case DATASOURCETYPE_USERINPUT: .Val_UserInput = NewValue
              Case DATASOURCETYPE_STEPP: .Val_StEPP = NewValue
              Case DATASOURCETYPE_CORR: .Val_Corr = NewValue
            End Select
          End With
      End Select
      If (Raise_Dirty_Flag) Then
        'THROW DIRTY FLAG.
        Call Local_DirtyStatus_Set(frmD0_Props_Is_Dirty, True)
      End If
      'REFRESH WINDOW.
      Call frmD0_Props_Refresh(Temp_Plant)
    End If
  End If
End Sub




Private Sub txtDataStr_GotFocus(Index As Integer)
Dim Ctl As Control
Set Ctl = txtDataStr(Index)
Dim StatusMessagePanel As String
  Call Global_GotFocus(Ctl)
  Select Case Index
    Case 0:
      StatusMessagePanel = "Type in the contaminant name"
  End Select
  Call Local_GenericStatus_Set(StatusMessagePanel)
End Sub
Private Sub txtDataStr_KeyPress(Index As Integer, KeyAscii As Integer)
  KeyAscii = Global_TextKeyPress(KeyAscii)
End Sub
Private Sub txtDataStr_LostFocus(Index As Integer)
Dim Ctl As Control
Set Ctl = txtDataStr(Index)
'''''
'Dim NewValue_Okay As Integer
'Dim NewValue As Double
'Dim Val_Low As Double
'Dim Val_High As Double
'Dim Raise_Dirty_Flag As Boolean
'Dim Too_Small As Integer
  With Temp_Plant.ChemicalData
    If (Trim$(Ctl.Text) = "") Then
      Ctl.Text = .ContaminantName
      'Call Show_Error("You must enter a non-blank string for the component name.")
      'NOTE: SHOWING THIS ERROR MESSAGE MESSES UP THE
      'SUBSEQUENT GOTFOCUS IF THE USER HITS <Enter> OR <Tab>.
    Else
      If (Trim$(.ContaminantName) <> Trim$(Ctl.Text)) Then
        .ContaminantName = Trim$(Ctl.Text)
        'THROW DIRTY FLAG.
        Call Local_DirtyStatus_Set( _
            frmD0_Props_Is_Dirty, True)
        '''''REFRESH WINDOW.
        ''''Call frmD0_Props_Refresh(Temp_Plant)
      End If
    End If
  End With
  Call Global_LostFocus(Ctl)
  Call Local_GenericStatus_Set("")
  Exit Sub
End Sub



