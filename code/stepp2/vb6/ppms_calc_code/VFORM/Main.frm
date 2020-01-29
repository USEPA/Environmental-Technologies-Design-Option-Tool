VERSION 5.00
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "THREED32.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{FAEEE763-117E-101B-8933-08002B2F4F5A}#1.1#0"; "DBLIST32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmMain 
   Appearance      =   0  'Flat
   BackColor       =   &H00C0C0C0&
   Caption         =   "{Generic Application -- Me.Caption set as Name_App_Short}"
   ClientHeight    =   8475
   ClientLeft      =   1410
   ClientTop       =   1305
   ClientWidth     =   12720
   ForeColor       =   &H80000008&
   Icon            =   "Main.frx":0000
   LockControls    =   -1  'True
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   8475
   ScaleWidth      =   12720
   Begin VB.PictureBox Picture1 
      Align           =   4  'Align Right
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   8070
      Left            =   9855
      ScaleHeight     =   8010
      ScaleWidth      =   2805
      TabIndex        =   24
      Top             =   0
      Visible         =   0   'False
      Width           =   2865
      Begin VB.PictureBox picUnvalidated 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   1050
         Picture         =   "Main.frx":030A
         ScaleHeight     =   285
         ScaleWidth      =   285
         TabIndex        =   29
         Top             =   120
         Width           =   345
      End
      Begin VB.PictureBox picValidated 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   660
         Picture         =   "Main.frx":03F4
         ScaleHeight     =   285
         ScaleWidth      =   285
         TabIndex        =   28
         Top             =   120
         Width           =   345
      End
      Begin VB.PictureBox picClosed 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   780
         Picture         =   "Main.frx":04DE
         ScaleHeight     =   285
         ScaleWidth      =   285
         TabIndex        =   27
         Top             =   6210
         Width           =   345
      End
      Begin VB.PictureBox picOpen 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   1170
         Picture         =   "Main.frx":05C8
         ScaleHeight     =   285
         ScaleWidth      =   285
         TabIndex        =   26
         Top             =   6210
         Width           =   345
      End
      Begin VB.PictureBox picLeaf 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   1560
         Picture         =   "Main.frx":06B2
         ScaleHeight     =   285
         ScaleWidth      =   285
         TabIndex        =   25
         Top             =   6210
         Width           =   345
      End
      Begin MSComctlLib.ImageList ilist_Valid 
         Left            =   30
         Top             =   120
         _ExtentX        =   1005
         _ExtentY        =   1005
         BackColor       =   -2147483643
         MaskColor       =   12632256
         _Version        =   393216
      End
      Begin MSComctlLib.ImageList ilist_Tree 
         Left            =   150
         Top             =   6210
         _ExtentX        =   1005
         _ExtentY        =   1005
         BackColor       =   -2147483643
         MaskColor       =   12632256
         _Version        =   393216
      End
      Begin Threed.SSFrame SSFrame2 
         Height          =   2385
         Left            =   90
         TabIndex        =   30
         Top             =   3540
         Visible         =   0   'False
         Width           =   2325
         _Version        =   65536
         _ExtentX        =   4101
         _ExtentY        =   4207
         _StockProps     =   14
         Caption         =   "Data Control -- Invisible"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Begin VB.Data datactlMaster 
            Caption         =   "datactlMaster"
            Connect         =   "Access"
            DatabaseName    =   ""
            DefaultCursorType=   0  'DefaultCursor
            DefaultType     =   2  'UseODBC
            Exclusive       =   0   'False
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   90
            Options         =   0
            ReadOnly        =   -1  'True
            RecordsetType   =   2  'Snapshot
            RecordSource    =   "PEARLS List"
            Top             =   1890
            Visible         =   0   'False
            Width           =   1935
         End
         Begin VB.Label Label1 
            Caption         =   $"Main.frx":079C
            Height          =   1635
            Left            =   90
            TabIndex        =   31
            Top             =   240
            Width           =   1995
         End
      End
      Begin Threed.SSFrame SSFrame4 
         Height          =   1035
         Left            =   180
         TabIndex        =   32
         Top             =   2430
         Visible         =   0   'False
         Width           =   3015
         _Version        =   65536
         _ExtentX        =   5318
         _ExtentY        =   1826
         _StockProps     =   14
         Caption         =   "Used -- Invisible"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Begin MSComDlg.CommonDialog CommonDialog1 
            Left            =   240
            Top             =   240
            _ExtentX        =   847
            _ExtentY        =   847
            _Version        =   393216
         End
      End
      Begin Threed.SSFrame SSFrame1 
         Height          =   2805
         Left            =   1410
         TabIndex        =   33
         Top             =   1830
         Visible         =   0   'False
         Width           =   7455
         _Version        =   65536
         _ExtentX        =   13150
         _ExtentY        =   4948
         _StockProps     =   14
         Caption         =   "Unused -- Invisible"
         ForeColor       =   255
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Begin VB.ComboBox cboUnitsX 
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
            Left            =   4320
            Style           =   2  'Dropdown List
            TabIndex        =   41
            TabStop         =   0   'False
            Top             =   360
            Width           =   1545
         End
         Begin VB.TextBox txtDataX 
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
            Left            =   2250
            TabIndex        =   40
            Text            =   "txtDataX()"
            Top             =   390
            Width           =   1995
         End
         Begin VB.TextBox txtDataX 
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
            Index           =   1
            Left            =   2250
            TabIndex        =   39
            Text            =   "txtDataX()"
            Top             =   810
            Width           =   1995
         End
         Begin VB.ComboBox cboUnitsX 
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
            Left            =   4320
            Style           =   2  'Dropdown List
            TabIndex        =   38
            TabStop         =   0   'False
            Top             =   780
            Width           =   1545
         End
         Begin VB.TextBox txtDataX 
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
            Index           =   2
            Left            =   2250
            TabIndex        =   37
            Text            =   "txtDataX()"
            Top             =   1230
            Width           =   1995
         End
         Begin VB.ComboBox cboUnitsX 
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
            Left            =   4320
            Style           =   2  'Dropdown List
            TabIndex        =   36
            TabStop         =   0   'False
            Top             =   1200
            Width           =   1545
         End
         Begin VB.TextBox txtDataX 
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
            Index           =   3
            Left            =   2250
            TabIndex        =   35
            Text            =   "txtDataX()"
            Top             =   1650
            Width           =   1995
         End
         Begin VB.ComboBox cboUnitsX 
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
            Left            =   4320
            Style           =   2  'Dropdown List
            TabIndex        =   34
            TabStop         =   0   'False
            Top             =   1620
            Width           =   1545
         End
         Begin VB.Label lblDataX 
            Alignment       =   1  'Right Justify
            Caption         =   "Length:"
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
            TabIndex        =   45
            Top             =   420
            Width           =   1845
         End
         Begin VB.Label lblDataX 
            Alignment       =   1  'Right Justify
            Caption         =   "Diameter:"
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
            Left            =   300
            TabIndex        =   44
            Top             =   840
            Width           =   1845
         End
         Begin VB.Label lblDataX 
            Alignment       =   1  'Right Justify
            Caption         =   "Mass:"
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
            Left            =   300
            TabIndex        =   43
            Top             =   1260
            Width           =   1845
         End
         Begin VB.Label lblDataX 
            Alignment       =   1  'Right Justify
            Caption         =   "Flow Rate:"
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
            Left            =   300
            TabIndex        =   42
            Top             =   1680
            Width           =   1845
         End
      End
      Begin VB.Label Label2 
         Caption         =   "Note: In calculation code, replace all msgbox() error-displaying with an alternate pass-back error system."
         ForeColor       =   &H000000FF&
         Height          =   1035
         Left            =   60
         TabIndex        =   46
         Top             =   750
         Width           =   2235
      End
   End
   Begin Threed.SSPanel sspAll 
      Height          =   7095
      Left            =   90
      TabIndex        =   3
      Top             =   90
      Width           =   9885
      _Version        =   65536
      _ExtentX        =   17436
      _ExtentY        =   12515
      _StockProps     =   15
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Begin Threed.SSPanel sspBottom 
         Height          =   4245
         Left            =   180
         TabIndex        =   14
         Top             =   2670
         Width           =   9495
         _Version        =   65536
         _ExtentX        =   16748
         _ExtentY        =   7488
         _StockProps     =   15
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Begin Threed.SSFrame ssfMain 
            Height          =   3975
            Left            =   2820
            TabIndex        =   18
            Top             =   60
            Width           =   6555
            _Version        =   65536
            _ExtentX        =   11562
            _ExtentY        =   7011
            _StockProps     =   14
            Caption         =   "{Selected Property Sheet:}"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Begin VB.Frame sspBasic 
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   3225
               Left            =   120
               TabIndex        =   47
               Top             =   510
               Width           =   6285
               Begin VB.TextBox txtDataStr 
                  BackColor       =   &H8000000F&
                  Height          =   315
                  Index           =   5
                  Left            =   2160
                  Locked          =   -1  'True
                  TabIndex        =   64
                  Text            =   "txtDataStr()"
                  Top             =   2790
                  Width           =   3200
               End
               Begin VB.TextBox txtDataStr 
                  BackColor       =   &H8000000F&
                  Height          =   315
                  Index           =   4
                  Left            =   2160
                  Locked          =   -1  'True
                  TabIndex        =   62
                  Text            =   "txtDataStr()"
                  Top             =   2430
                  Width           =   3200
               End
               Begin VB.TextBox txtDataStr 
                  BackColor       =   &H8000000F&
                  Height          =   315
                  Index           =   3
                  Left            =   2160
                  Locked          =   -1  'True
                  TabIndex        =   60
                  Text            =   "txtDataStr()"
                  Top             =   2070
                  Width           =   3200
               End
               Begin VB.TextBox txtDataStr 
                  BackColor       =   &H8000000F&
                  Height          =   315
                  Index           =   2
                  Left            =   2160
                  Locked          =   -1  'True
                  TabIndex        =   58
                  Text            =   "txtDataStr()"
                  Top             =   1710
                  Width           =   3200
               End
               Begin VB.TextBox txtDataStr 
                  BackColor       =   &H8000000F&
                  Height          =   315
                  Index           =   1
                  Left            =   2160
                  Locked          =   -1  'True
                  TabIndex        =   56
                  Text            =   "txtDataStr()"
                  Top             =   1350
                  Width           =   3200
               End
               Begin VB.TextBox txtDataStr 
                  BackColor       =   &H8000000F&
                  Height          =   315
                  Index           =   0
                  Left            =   2160
                  Locked          =   -1  'True
                  TabIndex        =   54
                  Text            =   "txtDataStr()"
                  Top             =   990
                  Width           =   3675
               End
               Begin VB.ComboBox cboUnits 
                  Height          =   315
                  Index           =   1
                  Left            =   3510
                  Style           =   2  'Dropdown List
                  TabIndex        =   53
                  Top             =   600
                  Width           =   1695
               End
               Begin VB.TextBox txtData 
                  Alignment       =   2  'Center
                  Height          =   315
                  Index           =   1
                  Left            =   2160
                  TabIndex        =   51
                  Text            =   "txtData()"
                  Top             =   615
                  Width           =   1275
               End
               Begin VB.ComboBox cboUnits 
                  Height          =   315
                  Index           =   0
                  Left            =   3510
                  Style           =   2  'Dropdown List
                  TabIndex        =   50
                  Top             =   225
                  Width           =   1695
               End
               Begin VB.TextBox txtData 
                  Alignment       =   2  'Center
                  Height          =   315
                  Index           =   0
                  Left            =   2160
                  TabIndex        =   48
                  Text            =   "txtData()"
                  Top             =   240
                  Width           =   1275
               End
               Begin VB.Label lblDataStr 
                  Alignment       =   1  'Right Justify
                  Caption         =   "Source:"
                  Height          =   225
                  Index           =   5
                  Left            =   180
                  TabIndex        =   65
                  Top             =   2835
                  Width           =   1905
               End
               Begin VB.Label lblDataStr 
                  Alignment       =   1  'Right Justify
                  Caption         =   "Family:"
                  Height          =   225
                  Index           =   4
                  Left            =   180
                  TabIndex        =   63
                  Top             =   2475
                  Width           =   1905
               End
               Begin VB.Label lblDataStr 
                  Alignment       =   1  'Right Justify
                  Caption         =   "Formula:"
                  Height          =   225
                  Index           =   3
                  Left            =   180
                  TabIndex        =   61
                  Top             =   2115
                  Width           =   1905
               End
               Begin VB.Label lblDataStr 
                  Alignment       =   1  'Right Justify
                  Caption         =   "SMILES:"
                  Height          =   225
                  Index           =   2
                  Left            =   180
                  TabIndex        =   59
                  Top             =   1755
                  Width           =   1905
               End
               Begin VB.Label lblDataStr 
                  Alignment       =   1  'Right Justify
                  Caption         =   "CAS:"
                  Height          =   225
                  Index           =   1
                  Left            =   180
                  TabIndex        =   57
                  Top             =   1395
                  Width           =   1905
               End
               Begin VB.Label lblDataStr 
                  Alignment       =   1  'Right Justify
                  Caption         =   "Name:"
                  Height          =   225
                  Index           =   0
                  Left            =   180
                  TabIndex        =   55
                  Top             =   1035
                  Width           =   1905
               End
               Begin VB.Label lblData 
                  Alignment       =   1  'Right Justify
                  Caption         =   "Operating Pressure:"
                  Height          =   225
                  Index           =   1
                  Left            =   180
                  TabIndex        =   52
                  Top             =   660
                  Width           =   1905
               End
               Begin VB.Label lblData 
                  Alignment       =   1  'Right Justify
                  Caption         =   "Operating Temperature:"
                  Height          =   225
                  Index           =   0
                  Left            =   180
                  TabIndex        =   49
                  Top             =   285
                  Width           =   1905
               End
            End
            Begin VB.TextBox txtChemNote 
               Height          =   555
               Left            =   2010
               MultiLine       =   -1  'True
               ScrollBars      =   2  'Vertical
               TabIndex        =   23
               Text            =   "Main.frx":0843
               Top             =   270
               Visible         =   0   'False
               Width           =   1755
            End
            Begin MSComctlLib.ListView lvMain 
               DragIcon        =   "Main.frx":0851
               Height          =   585
               Left            =   90
               TabIndex        =   21
               Top             =   240
               Width           =   1695
               _ExtentX        =   2990
               _ExtentY        =   1032
               View            =   3
               LabelEdit       =   1
               LabelWrap       =   -1  'True
               HideSelection   =   0   'False
               AllowReorder    =   -1  'True
               FullRowSelect   =   -1  'True
               GridLines       =   -1  'True
               _Version        =   393217
               ForeColor       =   -2147483640
               BackColor       =   -2147483643
               BorderStyle     =   1
               Appearance      =   1
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               NumItems        =   0
            End
         End
         Begin Threed.SSFrame ssfPropSheets 
            Height          =   2655
            Left            =   60
            TabIndex        =   16
            Top             =   60
            Width           =   2805
            _Version        =   65536
            _ExtentX        =   4948
            _ExtentY        =   4683
            _StockProps     =   14
            Caption         =   "Property Sheets:"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Begin VB.ListBox lstPropSheets 
               BackColor       =   &H8000000F&
               Height          =   1815
               Left            =   120
               TabIndex        =   17
               Top             =   300
               Width           =   2565
            End
         End
      End
      Begin Threed.SSPanel sspTop 
         Height          =   2115
         Left            =   180
         TabIndex        =   4
         Top             =   180
         Width           =   7155
         _Version        =   65536
         _ExtentX        =   12621
         _ExtentY        =   3731
         _StockProps     =   15
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Begin Threed.SSFrame ssfMasterList 
            Height          =   1875
            Left            =   60
            TabIndex        =   5
            Top             =   60
            Width           =   2205
            _Version        =   65536
            _ExtentX        =   3889
            _ExtentY        =   3307
            _StockProps     =   14
            Caption         =   "Master Chemical List:"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Begin MSDBCtls.DBList dblstMaster 
               Bindings        =   "Main.frx":0B5B
               DataSource      =   "datactlMaster"
               Height          =   1425
               Left            =   120
               TabIndex        =   19
               Top             =   300
               Width           =   1395
               _ExtentX        =   2461
               _ExtentY        =   2514
               _Version        =   393216
               MatchEntry      =   -1  'True
               ListField       =   "Name"
               BoundColumn     =   "Name"
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
         Begin Threed.SSFrame ssfButtons 
            Height          =   1875
            Left            =   2250
            TabIndex        =   6
            Top             =   60
            Width           =   1575
            _Version        =   65536
            _ExtentX        =   2778
            _ExtentY        =   3307
            _StockProps     =   14
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Begin VB.CommandButton cmdListButtons 
               Caption         =   "{File Note ...}"
               Enabled         =   0   'False
               Height          =   255
               Index           =   6
               Left            =   90
               TabIndex        =   20
               TabStop         =   0   'False
               ToolTipText     =   "View/edit file note"
               Top             =   1530
               Width           =   1395
            End
            Begin VB.CommandButton cmdListButtons 
               Caption         =   ">"
               Height          =   255
               Index           =   0
               Left            =   90
               TabIndex        =   12
               TabStop         =   0   'False
               ToolTipText     =   "Select chemical into user list"
               Top             =   180
               Width           =   1395
            End
            Begin VB.CommandButton cmdListButtons 
               Caption         =   "<<"
               Enabled         =   0   'False
               Height          =   255
               Index           =   5
               Left            =   810
               TabIndex        =   11
               TabStop         =   0   'False
               ToolTipText     =   "Deselect all chemicals from user list"
               Top             =   1260
               Width           =   675
            End
            Begin VB.CommandButton cmdListButtons 
               Caption         =   "<"
               Enabled         =   0   'False
               Height          =   255
               Index           =   4
               Left            =   90
               TabIndex        =   10
               TabStop         =   0   'False
               ToolTipText     =   "Deselect chemical from user list"
               Top             =   1260
               Width           =   690
            End
            Begin VB.CommandButton cmdListButtons 
               Caption         =   "Create New ..."
               Enabled         =   0   'False
               Height          =   255
               Index           =   3
               Left            =   90
               TabIndex        =   9
               TabStop         =   0   'False
               ToolTipText     =   "Create a new chemical"
               Top             =   990
               Width           =   1395
            End
            Begin VB.CommandButton cmdListButtons 
               Caption         =   "Find ..."
               Enabled         =   0   'False
               Height          =   255
               Index           =   2
               Left            =   90
               TabIndex        =   8
               TabStop         =   0   'False
               ToolTipText     =   "Search for a chemical"
               Top             =   720
               Width           =   1395
            End
            Begin VB.CommandButton cmdListButtons 
               Caption         =   "Recalculate All"
               Height          =   255
               Index           =   1
               Left            =   90
               TabIndex        =   7
               TabStop         =   0   'False
               ToolTipText     =   "Recalculate all properties for all chemicals"
               Top             =   450
               Width           =   1395
            End
         End
         Begin Threed.SSFrame ssfUserList 
            Height          =   1875
            Left            =   3810
            TabIndex        =   13
            Top             =   60
            Width           =   2205
            _Version        =   65536
            _ExtentX        =   3889
            _ExtentY        =   3307
            _StockProps     =   14
            Caption         =   "User Chemical List:"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Begin VB.ListBox lstUser 
               Height          =   1230
               Left            =   120
               TabIndex        =   22
               Top             =   300
               Width           =   1395
            End
         End
      End
      Begin Threed.SSPanel sspSep 
         Height          =   165
         Left            =   180
         TabIndex        =   15
         Top             =   2370
         Width           =   7575
         _Version        =   65536
         _ExtentX        =   13361
         _ExtentY        =   291
         _StockProps     =   15
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
   Begin Threed.SSPanel sspStatusBar 
      Align           =   2  'Align Bottom
      Height          =   405
      Left            =   0
      TabIndex        =   0
      Top             =   8070
      Width           =   12720
      _Version        =   65536
      _ExtentX        =   22437
      _ExtentY        =   714
      _StockProps     =   15
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
         Width           =   7185
         _Version        =   65536
         _ExtentX        =   12674
         _ExtentY        =   503
         _StockProps     =   15
         Caption         =   "sspanel_Status"
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
   Begin VB.Menu mnuFile 
      Caption         =   "&File"
      Begin VB.Menu mnuFileItem 
         Caption         =   "&New"
         Index           =   0
         Shortcut        =   ^N
      End
      Begin VB.Menu mnuFileItem 
         Caption         =   "&Open ..."
         Index           =   1
         Shortcut        =   ^O
      End
      Begin VB.Menu mnuFileItem 
         Caption         =   "&Save"
         Index           =   2
         Shortcut        =   ^S
      End
      Begin VB.Menu mnuFileItem 
         Caption         =   "Save &As ..."
         Index           =   3
      End
      Begin VB.Menu mnuFileItem 
         Caption         =   "-"
         Index           =   5
      End
      Begin VB.Menu mnuFileItem 
         Caption         =   "Select P&rinter ..."
         Index           =   6
      End
      Begin VB.Menu mnuFileItem 
         Caption         =   "&Print ..."
         Enabled         =   0   'False
         Index           =   7
      End
      Begin VB.Menu mnuFileItem 
         Caption         =   "-"
         Index           =   190
      End
      Begin VB.Menu mnuFileItem 
         Caption         =   "&1 Old File #1"
         Index           =   191
         Visible         =   0   'False
      End
      Begin VB.Menu mnuFileItem 
         Caption         =   "&2 Old File #2"
         Index           =   192
         Visible         =   0   'False
      End
      Begin VB.Menu mnuFileItem 
         Caption         =   "&3 Old File #3"
         Index           =   193
         Visible         =   0   'False
      End
      Begin VB.Menu mnuFileItem 
         Caption         =   "&4 Old File #4"
         Index           =   194
         Visible         =   0   'False
      End
      Begin VB.Menu mnuFileItem 
         Caption         =   "-"
         Index           =   199
         Visible         =   0   'False
      End
      Begin VB.Menu mnuFileItem 
         Caption         =   "E&xit"
         Index           =   200
      End
   End
   Begin VB.Menu mnuEdit 
      Caption         =   "&Edit"
      Begin VB.Menu mnuEditItem 
         Caption         =   "Select Chemical"
         Index           =   10
      End
      Begin VB.Menu mnuEditItem 
         Caption         =   "Recalculate All"
         Index           =   20
      End
      Begin VB.Menu mnuEditItem 
         Caption         =   "&Find ..."
         Enabled         =   0   'False
         Index           =   30
         Shortcut        =   {F3}
      End
      Begin VB.Menu mnuEditItem 
         Caption         =   "Create New Chemical ..."
         Enabled         =   0   'False
         Index           =   40
      End
      Begin VB.Menu mnuEditItem 
         Caption         =   "Unselect Chemical"
         Enabled         =   0   'False
         Index           =   50
      End
      Begin VB.Menu mnuEditItem 
         Caption         =   "Unselect All Chemicals"
         Enabled         =   0   'False
         Index           =   60
      End
      Begin VB.Menu mnuEditItem 
         Caption         =   "-"
         Index           =   70
      End
      Begin VB.Menu mnuEditItem 
         Caption         =   "View/Edit File Note ..."
         Enabled         =   0   'False
         Index           =   80
      End
      Begin VB.Menu mnuEditItem 
         Caption         =   "Refresh"
         Enabled         =   0   'False
         Index           =   90
      End
      Begin VB.Menu mnuEditItem 
         Caption         =   "-"
         Index           =   100
      End
      Begin VB.Menu mnuEditItem 
         Caption         =   "&Sort"
         Enabled         =   0   'False
         Index           =   110
         Begin VB.Menu mnuEditSortItem 
            Caption         =   "Sort Master Chemical List: Ascending"
            Index           =   10
         End
         Begin VB.Menu mnuEditSortItem 
            Caption         =   "Sort Master Chemical List: Descending"
            Index           =   20
         End
         Begin VB.Menu mnuEditSortItem 
            Caption         =   "-"
            Index           =   30
         End
         Begin VB.Menu mnuEditSortItem 
            Caption         =   "Sort User Chemical List: Ascending"
            Index           =   40
         End
         Begin VB.Menu mnuEditSortItem 
            Caption         =   "Sort User Chemical List: Descending"
            Index           =   50
         End
      End
   End
   Begin VB.Menu mnuProperty 
      Caption         =   "&Property"
      Begin VB.Menu mnuPropertyItem 
         Caption         =   "Change &Units ..."
         Index           =   10
         Shortcut        =   {F2}
      End
      Begin VB.Menu mnuPropertyItem 
         Caption         =   "View &Techniques Window ..."
         Index           =   20
         Shortcut        =   {F4}
      End
   End
   Begin VB.Menu mnuPlot 
      Caption         =   "Plo&t"
      Begin VB.Menu mnuPlotItem 
         Caption         =   "&Create Plot ..."
         Enabled         =   0   'False
         Index           =   10
         Shortcut        =   ^T
      End
   End
   Begin VB.Menu mnuOptions 
      Caption         =   "&Options"
      Begin VB.Menu mnuOptionsItem 
         Caption         =   "&Environment Preferences ..."
         Index           =   10
      End
      Begin VB.Menu mnuOptionsItem 
         Caption         =   "&File Preferences ..."
         Enabled         =   0   'False
         Index           =   20
      End
      Begin VB.Menu mnuOptionsItem 
         Caption         =   "-"
         Index           =   85
      End
      Begin VB.Menu mnuOptionsItem 
         Caption         =   "&Customize ..."
         Index           =   90
      End
   End
   Begin VB.Menu mnuHelp 
      Caption         =   "&Help"
      Begin VB.Menu mnuHelpItem 
         Caption         =   "&Online Help ..."
         Enabled         =   0   'False
         Index           =   10
      End
      Begin VB.Menu mnuHelpItem 
         Caption         =   "Online Manual ..."
         Enabled         =   0   'False
         Index           =   20
      End
      Begin VB.Menu mnuHelpItem 
         Caption         =   "-"
         Index           =   30
      End
      Begin VB.Menu mnuHelpItem 
         Caption         =   "View Version History ..."
         Index           =   80
      End
      Begin VB.Menu mnuHelpItem 
         Caption         =   "View Disclaimer ..."
         Index           =   85
      End
      Begin VB.Menu mnuHelpItem 
         Caption         =   "View Development Notes ..."
         Enabled         =   0   'False
         Index           =   90
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
   Begin VB.Menu mnuMTU 
      Caption         =   "&MTU Internal"
      Begin VB.Menu mnuMTUItem 
         Caption         =   "&Quick sig-fig test file generator"
         Index           =   10
      End
      Begin VB.Menu mnuMTUItem 
         Caption         =   "&Make menu invisible"
         Index           =   198
      End
      Begin VB.Menu mnuMTUItem 
         Caption         =   "&Read me"
         Index           =   199
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim sspSep_Setting As Double    'RANGE: 0 to 1.

Public HALT_Controls As Boolean





Const frmMain_declarations_end = True


Function frmMain_Go() _
    As Boolean
On Error GoTo err_ThisFunc
  '
  ' TRANSFER CHEMICAL LIST INTO dblstMaster.
  '
  With datactlMaster
    ''''.DatabaseName = PathMaster
    ''''.DatabaseName = "X:\pdt10\code\ppms\comm\990519_master_mdb_with_dippr801_data\MASTER.MDB"
    .databasename = fn_Master_MDB
    .RecordSource = "PEARLS List"
    .RecordsetType = 2
    .Refresh
    '.Recordset.FindFirst "Name = 'TOLUENE'"
    '''''frmMain!LSTSelList.Text = "TOLUENE"
    '.Refresh
  End With
  '
  ' SELECT THE TOLUENE CHEMICAL.
  '
  datactlMaster.Recordset.FindFirst "Name = 'TOLUENE'"
  '''''frmMain!LSTSelList.Text = "TOLUENE"
  datactlMaster.Refresh
  dblstMaster.Text = "TOLUENE"
  '
  ' DISPLAY THE MAIN WINDOW.
  '
  frmMain.Show 1
  '
  ' EXIT OUT OF HERE.
  '
exit_normally_ThisFunc:
  frmMain_Go = True
  Exit Function
exit_err_ThisFunc:
  frmMain_Go = False
  Exit Function
err_ThisFunc:
  Call Show_Trapped_Error("frmMain_Go")
  Resume exit_err_ThisFunc
End Function


Function frmMain_PopulateFirstTime_SeveralControls() _
    As Boolean
On Error GoTo err_ThisFunc
Dim Frm As Form
Set Frm = frmMain
Dim Ctl_IL As Control
Dim Ctl_LV As Control
Dim ImgI As ListImage
Dim ItmX As ListItem
  Frm.lstUser.Clear
  Frm.lstPropSheets.Clear
  '
  ' SET UP Frm.ilist_Valid.
  '
  Set Ctl_IL = Frm.ilist_Valid
  Ctl_IL.ListImages.Clear
  Set ImgI = Ctl_IL.ListImages.Add _
      (, "validated", Frm.picValidated.Picture)
  Set ImgI = Ctl_IL.ListImages.Add _
      (, "unvalidated", Frm.picUnvalidated.Picture)
  '
  ' SET UP Frm.lvMain.
  '
  Set Ctl_LV = Frm.lvMain
  Ctl_LV.View = lvwReport
  Ctl_LV.Icons = Ctl_IL
  Ctl_LV.SmallIcons = Ctl_IL
  Ctl_LV.ColumnHeaders.Clear
  Ctl_LV.ColumnHeaders.Add , , "Valid", 600, lvwColumnLeft
  Ctl_LV.ColumnHeaders.Add , , "Property", 3000, lvwColumnRight
  Ctl_LV.ColumnHeaders.Add , , "Value", 1200, lvwColumnRight
  Ctl_LV.ColumnHeaders.Add , , "Units", 1000, lvwColumnLeft
  Ctl_LV.ColumnHeaders.Add , , "Note", 600, lvwColumnLeft
  '
  ' ------------ TEMPORARY TEST DATA FOLLOWS: ------------
  '
  Set ItmX = Ctl_LV.ListItems.Add(, "x1", " ")
  ItmX.SubItems(1) = "Molecular Weight"
  ItmX.SubItems(2) = "60.023"
  ItmX.SubItems(3) = "g/gmol"
  ItmX.SubItems(4) = " "
  ItmX.Icon = 1: ItmX.SmallIcon = 1
  Set ItmX = Ctl_LV.ListItems.Add(, "x2", " ")
  ItmX.SubItems(1) = "Liquid density @ 298.15 K"
  ItmX.SubItems(2) = "Not Available"
  ItmX.SubItems(3) = "kg/m3"
  ItmX.SubItems(4) = " "
  ItmX.Icon = 2: ItmX.SmallIcon = 2
  '
  ' ------------ TEMPORARY TEST DATA ENDS. ------------
  '
  
  
'Sub populate_lvThis()
'Dim ImgI As ListImage
'Dim ItmX As ListItem
'Dim Ctl_LV As Control
'Dim Ctl_IL As Control
'  'SET UP THE lvThis CONTROL (IMAGES, STYLES, ETC).
'  Set Ctl_IL = frmMain.ilist_Valid
'  Ctl_IL.ListImages.Clear
'  Set ImgI = Ctl_IL.ListImages.Add _
'      (, "validated", frmMain.picValidated.Picture)
'  Set ImgI = Ctl_IL.ListImages.Add _
'      (, "unvalidated", frmMain.picUnvalidated.Picture)
'  Set Ctl_LV = lvThis
'  Ctl_LV.View = lvwReport
'  Ctl_LV.Icons = Ctl_IL
'  Ctl_LV.SmallIcons = Ctl_IL
'  'Ctl_LV.ColumnHeaders.Add , , "x2", 1000, lvwColumnLeft
'  'Ctl_LV.ColumnHeaders.Add , , "x3", 1000, lvwColumnLeft
'''''  '
'''''  ' PROTOTYPE DATA.
'''''  '
'''''  Set ItmX = Ctl_LV.ListItems.Add(, "x1", "Condenser")
'''''  ItmX.SubItems(1) = "($62,853)"
'''''  ItmX.SubItems(2) = "Heat Transfer Operations : Shell and Tube Exchanger : Fixed Tubesheet : Carbon Steel, 150 psig"
'''''  ItmX.SubItems(3) = "Validated(LE)"
'''''  ItmX.Icon = 1: ItmX.SmallIcon = 1
'''''  Set ItmX = Ctl_LV.ListItems.Add(, "x2", "DC-100")
'''''  ItmX.SubItems(1) = "($517,501)"
'''''  ItmX.SubItems(2) = "Vessels : Column : Tray : Carbon Steel, 150 psig"
'''''  ItmX.SubItems(3) = "Validated(LE)"
'''''  ItmX.Icon = 1: ItmX.SmallIcon = 1
'''''  Set ItmX = Ctl_LV.ListItems.Add(, "x3", "Preheater")
'''''  ItmX.SubItems(1) = "($242,201)"
'''''  ItmX.SubItems(2) = "Heat Transfer Operations : Shell and Tube Exchanger : Fixed Tubesheet : Carbon Steel, 150 psig"
'''''  ItmX.SubItems(3) = "Unvalidated(LE)"
'''''  ItmX.Icon = 2: ItmX.SmallIcon = 2
'''''  Set ItmX = Ctl_LV.ListItems.Add(, "x4", "Reboiler")
'''''  ItmX.SubItems(1) = "($361,671)"
'''''  ItmX.SubItems(2) = "Heat Transfer Operations : Shell and Tube Exchanger : Kettle Reboiler : Carbon Steel, 150 psig"
'''''  ItmX.SubItems(3) = "Unvalidated(LE)"
'''''  ItmX.Icon = 2: ItmX.SmallIcon = 2
'End Sub
  
  
  
  
  



exit_normally_ThisFunc:
  frmMain_PopulateFirstTime_SeveralControls = True
  Exit Function
exit_err_ThisFunc:
  frmMain_PopulateFirstTime_SeveralControls = False
  Exit Function
err_ThisFunc:
  Call Show_Trapped_Error("frmMain_PopulateFirstTime_SeveralControls")
  Resume exit_err_ThisFunc
End Function


Sub Avoid_Weird_Focus_Problem()
  Call unitsys_control_MostRecent_Force_lostfocus
  'frmMain.SetFocus
  '
  ' NOTE: IT IS VERY IMPORTANT TO SET FOCUS HERE
  ' TO SOME NON-UNITTEXTBOX CONTROL, I.E. DON'T
  ' SET IT TO txtData(0...3), BUT cboUnits(0)
  ' IS OKAY.
  ''''cboUnits(0).SetFocus
  cmdListButtons(0).SetFocus
  'Text1.SetFocus
End Sub


Sub frmMain_Populate_Units()
Dim Frm As Form
Set Frm = frmMain
  Call unitsys_register(Frm, lblData(0), txtdata(0), cboUnits(0), "temperature", _
      "K", "K", "", "", 100#, True)
  Call unitsys_register(Frm, lblData(1), txtdata(1), cboUnits(1), "pressure", _
      "Pa", "Pa", "", "", 100#, True)
''''  Call unitsys_register(Frm, lblData(2), txtData(2), cboUnits(2), "mass", _
''''      "kg", "kg", "", "", 100#, True)
''''  Call unitsys_register(Frm, lblData(3), txtData(3), cboUnits(3), "flow_volumetric", _
''''      "m³/s", "m³/s", "", "", 100#, True)
End Sub


Sub Local_DirtyStatus_Set(DirtyFlag As Boolean, NewSetting As Boolean)
  Call Global_DirtyStatus_Set(Me, DirtyFlag, NewSetting)
End Sub
Sub Local_GenericStatus_Set(NewString As String)
  Call Global_GenericStatus_Set(Me, NewString)
End Sub


Function frmMain_Resize( _
    ) _
    As Boolean
On Error GoTo err_ThisFunc
Const USE_MARGIN_sspSep = 50
Dim XX As Double
  '
  '////////// START OF MAIN RESIZING CODE. ///////////////////////////////////////
  '
  '
  ' RESIZE sspAll AND ALL CONTAINED CONTROLS.
  '
  XX = frmMain.ScaleHeight - frmMain.sspStatusBar.Height
  XX = IIf(XX > 10, XX, 10)
  sspAll.Move _
      0, _
      0, _
      frmMain.ScaleWidth, _
      XX
    '
    ' RESIZE sspSep AND ALL CONTAINED CONTROLS.
    '
    sspSep.Move _
        -1000, _
        CDbl(sspAll.Height) * sspSep_Setting, _
        sspAll.Width + 1000 + 1000, _
        sspSep.Height
    '
    ' RESIZE sspTop AND ALL CONTAINED CONTROLS.
    '
    XX = sspSep.Top - USE_MARGIN_sspSep
    XX = IIf(XX > 10, XX, 10)
    sspTop.Move _
        0, _
        0, _
        sspAll.Width, _
        XX
      '
      ' SOME CONTAINED CONTROLS ...
      '
Dim Width_of_Each_List As Double
Dim Height_of_Each_List As Double
Dim Width_of_This As Double
Dim Height_of_This As Double
      XX = CDbl(sspTop.Width - ssfButtons.Width - 60 - 60) / 2#
      Width_of_Each_List = IIf(XX > 100#, XX, 100#)
      XX = sspTop.Height - 60 - 60
      Height_of_Each_List = IIf(XX > 100#, XX, 100#)
      '
      ' RESIZE ssfMasterList AND ALL CONTAINED CONTROLS.
      '
      ssfMasterList.Move _
          60, _
          60, _
          Width_of_Each_List, _
          Height_of_Each_List
        '
        ' RESIZE dblstMaster AND ALL CONTAINED CONTROLS.
        '
        XX = ssfMasterList.Width - 120 - 120
        Width_of_This = IIf(XX > 100#, XX, 100#)
        XX = ssfMasterList.Height - 300 - 120
        Height_of_This = IIf(XX > 100#, XX, 100#)
        dblstMaster.Move _
            120, _
            300, _
            Width_of_This, _
            Height_of_This
      '
      ' RESIZE ssfButtons AND ALL CONTAINED CONTROLS.
      '
      ssfButtons.Move _
          ssfMasterList.Left + ssfMasterList.Width, _
          60, _
          ssfButtons.Width, _
          ssfButtons.Height
      '
      ' RESIZE ssfUserList AND ALL CONTAINED CONTROLS.
      '
      ssfUserList.Move _
          ssfButtons.Left + ssfButtons.Width, _
          60, _
          Width_of_Each_List, _
          Height_of_Each_List
        '
        ' RESIZE lstUser AND ALL CONTAINED CONTROLS.
        '
        XX = ssfUserList.Width - 120 - 120
        Width_of_This = IIf(XX > 100#, XX, 100#)
        XX = ssfUserList.Height - 300 - 120
        Height_of_This = IIf(XX > 100#, XX, 100#)
        lstUser.Move _
            120, _
            300, _
            Width_of_This, _
            Height_of_This
    '
    ' RESIZE sspBottom AND ALL CONTAINED CONTROLS.
    '
    XX = sspAll.Height - (sspSep.Top + sspSep.Height + USE_MARGIN_sspSep)
    XX = IIf(XX > 10, XX, 10)
    sspBottom.Move _
        0, _
        sspSep.Top + sspSep.Height + USE_MARGIN_sspSep, _
        sspAll.Width, _
        XX
      '
      ' SOME CONTAINED CONTROLS ...
      '
Dim Height_of_Each_Property_Frame As Double
      XX = sspBottom.Height - 60 - 60
      Height_of_Each_Property_Frame = IIf(XX > 10, XX, 10)
      '
      ' RESIZE ssfPropSheets AND ALL CONTAINED CONTROLS.
      '
      ssfPropSheets.Move _
          60, _
          60, _
          ssfPropSheets.Width, _
          Height_of_Each_Property_Frame
        '
        ' RESIZE lstPropSheets AND ALL CONTAINED CONTROLS.
        '
        XX = ssfPropSheets.Width - 120 - 120
        Width_of_This = IIf(XX > 100#, XX, 100#)
        XX = ssfPropSheets.Height - 300 - 120
        Height_of_This = IIf(XX > 100#, XX, 100#)
        lstPropSheets.Move _
            120, _
            300, _
            Width_of_This, _
            Height_of_This
      '
      ' RESIZE ssfMain AND ALL CONTAINED CONTROLS.
      '
      XX = sspBottom.Width - 60 - (ssfPropSheets.Left + ssfPropSheets.Width)
      XX = IIf(XX > 10, XX, 10)
      ssfMain.Move _
          ssfPropSheets.Left + ssfPropSheets.Width, _
          60, _
          XX, _
          Height_of_Each_Property_Frame
        '
        ' RESIZE lvMain AND ALL CONTAINED CONTROLS.
        '
        XX = ssfMain.Width - 120 - 120
        Width_of_This = IIf(XX > 100#, XX, 100#)
        XX = ssfMain.Height - 300 - 120
        Height_of_This = IIf(XX > 100#, XX, 100#)
        lvMain.Move _
            120, _
            300, _
            Width_of_This, _
            Height_of_This
        '
        ' RESIZE txtChemNote AND ALL CONTAINED CONTROLS.
        '
        txtChemNote.Move _
            lvMain.Left, _
            lvMain.Top, _
            lvMain.Width, _
            lvMain.Height
        '
        ' RESIZE sspBasic AND ALL CONTAINED CONTROLS.
        '
        sspBasic.Move _
            lvMain.Left, _
            lvMain.Top, _
            lvMain.Width, _
            lvMain.Height
  '
  '////////// END OF MAIN RESIZING CODE. ///////////////////////////////////////
  '
exit_normally_ThisFunc:
  frmMain_Resize = True
  Exit Function
exit_err_ThisFunc:
  frmMain_Resize = False
  Exit Function
err_ThisFunc:
  Call Show_Trapped_Error("frmMain_Resize")
  GoTo exit_err_ThisFunc
End Function


Private Sub cboUnits_Click(Index As Integer)
Dim Ctl As Control
Set Ctl = cboUnits(Index)
  Call unitsys_control_cbox_click(Ctl)
End Sub
Private Sub cboUnits_KeyPress(Index As Integer, KeyAscii As Integer)
  KeyAscii = Global_TextKeyPress(KeyAscii)
End Sub


Private Sub cmdListButtons_Click(Index As Integer)
  Select Case Index
    '
    '////////////////////////////////////////////////////////////////////
    Case 0:     'SELECT CHEMICAL.
      Call mnuEditItem_Click(10)
    '
    '////////////////////////////////////////////////////////////////////
    Case 1:     'RECALCULATE.
      Call mnuEditItem_Click(20)
    '
    '////////////////////////////////////////////////////////////////////
    Case 2:     'FIND.
      Call mnuEditItem_Click(30)
    '
    '////////////////////////////////////////////////////////////////////
    Case 3:     'CREATE NEW.
      Call mnuEditItem_Click(40)
    '
    '////////////////////////////////////////////////////////////////////
    Case 4:     'UNSELECT CHEMICAL.
      Call mnuEditItem_Click(50)
    '
    '////////////////////////////////////////////////////////////////////
    Case 5:     'UNSELECT ALL CHEMICALS.
      Call mnuEditItem_Click(60)
    '
    '////////////////////////////////////////////////////////////////////
    Case 6:     'VIEW/EDIT FILE NOTE.
      Call mnuEditItem_Click(80)
  End Select
End Sub


Private Sub Form_Load()
Dim is_internal_mtu As Boolean
  '
  ' MISC INITS.
  '
  Call Local_DirtyStatus_Set(Project_Is_Dirty, False)
  Call Local_GenericStatus_Set("")
  Me.Caption = Name_App_Short
  Me.Width = 9600
  Me.Height = 7600
  Call CenterOnScreen(Me)
  ''''CommonDialog1.filename = MAIN_APP_PATH & "\examples\*.dat"
  CommonDialog1.FileName = _
      MAIN_APP_PATH & "\examples\*." & FileExt_App
  sspSep_Setting = 0.35
  Call frmMain_PopulateFirstTime_SeveralControls
  ssfMain.Caption = ""
  '
  ' MISC CONTROL RESETTINGS.
  '
  sspTop.BevelWidth = 0
  sspBottom.BevelWidth = 0
  sspAll.BevelWidth = 0
  HALT_Controls = False
  '
  ' CHECK FOR FILE THAT INDICATES THIS IS INTERNAL TO MTU:
  '
  is_internal_mtu = False
  If (check_internal_to_mtu()) Then is_internal_mtu = True
  mnuMTU.Visible = is_internal_mtu
  '
  ' POPULATE UNITS INTO SCROLLBOX CONTROLS.
  '
  Call frmMain_Populate_Units
  '
  ' CREATE A NEW FILE IN MEMORY.
  '
  Call file_new
  '
  ' POPULATE LAST-FEW-FILES LIST.
  '
  Call OldFileList_Populate( _
      1, _
      frmMain.mnuFileItem(199), _
      frmMain.mnuFileItem(191), _
      frmMain.mnuFileItem(192), _
      frmMain.mnuFileItem(193), _
      frmMain.mnuFileItem(194))
  '
  ' UPDATE RESIZING.
  '
  Call frmMain_Resize
End Sub
Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
  If (file_query_unload() = False) Then
    Cancel = True
  End If
End Sub
Private Sub Form_Resize()
  If (Me.WindowState <> vbMinimized) Then
    '
    ' WARNING: RESIZING WHILE MINIMIZED CAN CAUSE ERRORS!
    '
    Call frmMain_Resize
  End If
End Sub
Private Sub Form_Unload(Cancel As Integer)
  Call frmMain_Close_All_Windows
  ''''Call unitsys_unregister_all_on_form(Me)
End Sub


Private Sub lstPropSheets_Click()
  Call frmMain_Populate_lvMain
  Call frmMain_Refresh
End Sub
Private Sub lstUser_Click()
  Call frmMain_Populate_lvMain
  Call frmMain_Refresh
End Sub


Private Sub lvMain_Click()
  ''''Call mnuPropertyItem_Click(20)
End Sub
'Private Sub lvMain_DblClick()
'  MsgBox Me.lvMain.SelectedItem.Key
'End Sub
Private Sub lvMain_MouseUp( _
    Button As Integer, _
    Shift As Integer, _
    X As Single, _
    y As Single)
  If ((Button And vbLeftButton) = vbLeftButton) Then
    ''''Call mnuPropertyItem_Click(20)
  End If
  If ((Button And vbRightButton) = vbRightButton) Then
    Me.PopupMenu mnuProperty
  End If
End Sub


Private Sub mnuEditItem_Click(Index As Integer)
On Error GoTo err_ThisFunc
Dim ThisName As String
Dim UB As Integer
  Select Case Index
    '
    '////////////////////////////////////////////////////////////////////
    Case 10:    'SELECT CHEMICAL.
      If (UBound(NowProj.UserChemicals) >= MAX_USERCHEMICALS) Then
        Call Show_Error("The maximum number of chemicals " & _
            "has been reached (" & Trim$(Str$(MAX_USERCHEMICALS)) & _
            ".  You cannot add any more chemicals.")
        GoTo exit_err_ThisFunc
      End If
      ThisName = Trim$(dblstMaster.Text)
      ''''MsgBox "ThisName = `" & ThisName & "`"
      If (UserChemical_IsKeyExist(ThisName) = True) Then
        Call Show_Error("You already have selected a chemical " & _
            "named `" & ThisName & "`.")
        GoTo exit_err_ThisFunc
      End If
      UB = UBound(NowProj.UserChemicals)
      UB = UB + 1
      If (UB = 1) Then
        ReDim NowProj.UserChemicals(1 To UB)
      Else
        ReDim Preserve NowProj.UserChemicals(1 To UB)
      End If
      Call UserChemical_SetDefaults(NowProj.UserChemicals(UB), ThisName)
      '
      ' TRANSFER SOME DATABASE VALUES INTO CHEMICAL VARIABLE.
      '
      datactlMaster.Recordset.FindFirst "Name = '" & ThisName & "'"
      'datactlMaster.Refresh
      With NowProj.UserChemicals(UB)
        .CAS = Trim$(Str$(Database_Get_Long(datactlMaster.Recordset, "CAS")))
        .SMILES = Database_Get_String(datactlMaster.Recordset, "Smiles")
        .Formula = Database_Get_String(datactlMaster.Recordset, "Formula")
        .Family = Database_Get_String(datactlMaster.Recordset, "Chemical Family")
        .Source = Database_Get_String(datactlMaster.Recordset, "Source")
      End With
      '
      ' RECALCULATE FOR THIS CHEMICAL ONLY.
      '
      Call PropertyData_InitializeAll_OneChemical(UB)
      Call Recalculate_OneChemical(UB)
      '
      ' THROW DIRTY FLAG AND REFRESH WINDOW.
      '
      Call Local_DirtyStatus_Set(Project_Is_Dirty, True)
      Call frmMain_Populate_lstUser
      Call frmMain_Refresh
    '
    '////////////////////////////////////////////////////////////////////
    Case 20:    'RECALCULATE ALL.
      Call Recalculate_All
      '
      ' THROW DIRTY FLAG AND REFRESH WINDOW.
      '
      Call Local_DirtyStatus_Set(Project_Is_Dirty, True)
      Call frmMain_Populate_lstUser
      Call frmMain_Populate_lstPropSheets
      Call frmMain_Populate_lvMain
      Call frmMain_Refresh
    '
    '////////////////////////////////////////////////////////////////////
    Case 30:    'FIND.
    '
    '////////////////////////////////////////////////////////////////////
    Case 40:    'CREATE NEW.
    '
    '////////////////////////////////////////////////////////////////////
    Case 50:    'UNSELECT CHEMICAL.
    '
    '////////////////////////////////////////////////////////////////////
    Case 60:    'UNSELECT ALL CHEMICALS.
    '
    '////////////////////////////////////////////////////////////////////
    Case 80:    'VIEW/EDIT FILE NOTE.
'MsgBox Me.lvMain.SelectedItem.Index
MsgBox Me.lvMain.SelectedItem.Key
    
    '
    '////////////////////////////////////////////////////////////////////
    Case 90:    'REFRESH.
  End Select
exit_normally_ThisFunc:
  'xxxxx = True
  Exit Sub
exit_err_ThisFunc:
  'xxxxx = False
  Exit Sub
err_ThisFunc:
  Call Show_Trapped_Error("mnuEditItem_Click")
  Resume exit_err_ThisFunc
End Sub
Private Sub mnuFileItem_Click(Index As Integer)
  Select Case Index
    Case 0:      'New
      If (file_query_unload()) Then
        Call Avoid_Weird_Focus_Problem
        Call file_new
      End If
    Case 1:      'Open ...
      If (file_query_unload()) Then
        Call Avoid_Weird_Focus_Problem
        Call File_OpenAs("")
      End If
    Case 2:      'Save
      If (Current_Filename = "") Then
        Call Avoid_Weird_Focus_Problem
        Call File_SaveAs("")
      Else
        Call Avoid_Weird_Focus_Problem
        Call File_SaveAs(Current_Filename)
      End If
    Case 3:      'Save As ...
      Call Avoid_Weird_Focus_Problem
      Call File_SaveAs("")
    Case 6:       'Select Printer ...
      CommonDialog1.ShowPrinter
    'Case 85:      'Print ...
    '  frmPrint.Show 1
    Case 191 To 194:      'Last-few-files list
      If (file_query_unload()) Then
        If (mnuFileItem(Index).Visible) Then
          Call File_OpenAs(OldFiles(1, Index - 190))
        End If
      End If
    Case 200:     'Exit
      'NOTE: MDIForm_QueryUnload() TAKES CAKE OF THIS.
      'If we do it here, _two_ message boxes will pop up
      'when the user has data which needs saving !
      'If (file_query_unload()) Then
      '  Unload Me
      'End If
      Unload Me
  End Select
End Sub
Private Sub mnuHelpItem_Click(Index As Integer)
Dim fn_This As String
  Select Case Index
    Case 10:      'ONLINE HELP.
      'NOTE: We currently do NOT have the resources to
      'create an online help file for AdDesignS (1/7/98)
      'therefore no online help is available.
      Call Show_Message("Online help is currently unavailable.  " & _
          "Please refer to the printed manual or the Acrobat-format ADS.PDF file.")
      Exit Sub
      'Call LaunchFile_General("", MAIN_APP_PATH & "\help\ads.hlp")
    Case 20:      'ONLINE MANUAL.
      fn_This = MAIN_APP_PATH & "\help\ads.pdf"
      If (FileExists(fn_This) = False) Then
        Call Show_Message("The file `" & fn_This & "` is missing.")
        Exit Sub
      End If
      Call LaunchFile_General("", fn_This)
    Case 80:
      fn_This = MAIN_APP_PATH & "\dbase\readme_ppms_calc.txt"
      If (FileExists(fn_This) = False) Then
        Call Show_Message("The file `" & fn_This & "` is missing.")
        Exit Sub
      End If
      Call Launch_Notepad(fn_This)
    Case 85:    'VIEW DISCLAIMER.
      'SHOW THE DISCLAIMER WINDOW.
      splash_mode = 101
      splash_button_pressed = 0
      frmSplash.Show 1
    Case 99:    'ABOUT.
      frmAbout.Show 1
  End Select
End Sub
Private Sub mnuMTUItem_Click(Index As Integer)
  Select Case Index
    Case 10:
      Dim F As Integer
      Dim i As Integer
      Dim j As Integer
      Dim X As Double
      Dim UseFormat As String
      F = FreeFile
      Open "c:\sfig.txt" For Output As #F
      For i = 3 To 6
        For j = -5 To 10
          X = 10# ^ j
          UseFormat = GetDoubleFormat_VarSigFigs(X, i)
          Write #F, X, i, UseFormat, Format(X, UseFormat)
        Next j
      Next i
      Close #F
    ''''Case 40:    'KEEP TEMPORARY MODEL FILES.
    ''''  mnuMTUItem(40).Checked = Not mnuMTUItem(40).Checked
    Case 198:   'MAKE INVISIBLE.
      mnuMTU.Visible = False
    Case 199:   'READ ME.
      Call Show_Message("This menu should only appear on internal " & _
          "testing machines at MTU.  To remove the `MTU Internal` " & _
          "menu, select `Make menu invisible`.  This will make " & _
          "the menu invisible until the program is closed and reloaded.")
  End Select
End Sub
Private Sub mnuOptionsItem_Click(Index As Integer)
On Error GoTo err_ThisFunc
Dim out_HitCancel As Boolean
  Select Case Index
    '
    '////////////////////////////////////////////////////////////////////
    Case 10:        'ENVIRONMENT PREFERENCES.
      If (False = frmPrefEnvironment.frmPrefEnvironment_Go( _
          out_HitCancel)) Then
        GoTo exit_err_ThisFunc
      End If
      If (out_HitCancel = True) Then GoTo exit_normally_ThisFunc
      '
      ' REFRESH WINDOW.
      '
      Call frmMain_Populate_lvMain
      Call frmMain_Refresh
    '
    '////////////////////////////////////////////////////////////////////
    Case 90:        'CUSTOMIZE.
      If (False = frmCustomProperties.frmCustomProperties_Go( _
          out_HitCancel)) Then
        GoTo exit_err_ThisFunc
      End If
      If (out_HitCancel = True) Then GoTo exit_normally_ThisFunc
      DoEvents
      '
      ' REORGANIZE AND RECALCULATE.
      '
      Call PropertyData_InitializeAll_AllChemicals
      Call Recalculate_All
      '
      ' REFRESH WINDOW.
      '
      Call frmMain_Populate_lstUser
      Call frmMain_Populate_lstPropSheets
      Call frmMain_Populate_lvMain
      Call frmMain_Refresh
  End Select
exit_normally_ThisFunc:
  'xxxxx = True
  Exit Sub
exit_err_ThisFunc:
  'xxxxx = False
  Exit Sub
err_ThisFunc:
  Call Show_Trapped_Error("mnuOptionsItem_Click")
  GoTo exit_err_ThisFunc
End Sub
Private Sub mnuPropertyItem_Click(Index As Integer)
On Error GoTo err_ThisFunc
Dim NeedExtraction As Boolean
Dim in_Key As String
Dim out_idx_PropertySheetOrder As Integer
Dim out_idx_PropertyOrder As Integer
Dim This_Property_Code As Long
Dim This_idx_Chemical As Integer
Dim out_HitCancel As Boolean
Dim i As Integer
  NeedExtraction = False
  Select Case Index
    Case 10:          'CHANGE UNITS.
      NeedExtraction = True
    Case 20:          'VIEW TECHNIQUES WINDOW.
      NeedExtraction = True
  End Select
  If (NeedExtraction = True) Then
    '
    ' EXTRACT INDEXES FROM KEY OF THE SELECTED LIST ITEM.
    '
    in_Key = ""
    On Error Resume Next
    in_Key = Me.lvMain.SelectedItem.Key
    On Error GoTo err_ThisFunc
    If (False = frmMain_lvMain_Extract_Key_Info( _
        in_Key, _
        out_idx_PropertySheetOrder, _
        out_idx_PropertyOrder)) Then
      Call Show_Error("You must first select a property.")
      GoTo exit_err_ThisFunc
    End If
    This_Property_Code = NowProj.UserHierarchy. _
        PropertySheetOrder(out_idx_PropertySheetOrder). _
        PropertyOrder(out_idx_PropertyOrder).Property_Code
  End If
  This_idx_Chemical = frmMain_lstUser_GetItemData()
  Select Case Index
    '
    '////////////////////////////////////////////////////////////////////
    Case 10:          'CHANGE UNITS.
Dim in_UnitType As String
Dim in_UnitBase As String
Dim inout_UnitDisplayed As String
Dim idx_PropertyData As Integer
      idx_PropertyData = PropertyData_GetIndex( _
          This_idx_Chemical, _
          This_Property_Code)
      If (idx_PropertyData = -1) Then
        Call Show_Error("Unable to find property code of " & _
            Trim$(Str$(This_Property_Code)) & " for this chemical; " & _
            "cancelling unit change operation.")
        GoTo exit_err_ThisFunc
      End If
      With NowProj.UserChemicals(This_idx_Chemical).PropertyData(idx_PropertyData)
        in_UnitType = .UnitType
        in_UnitBase = .UnitBase
        inout_UnitDisplayed = .UnitDisplayed
      End With
      If (False = frmUnitsAndOrValue.frmUnitsAndOrValue_GoUnitsOnly( _
          in_UnitType, _
          in_UnitBase, _
          inout_UnitDisplayed, _
          out_HitCancel)) Then
        GoTo exit_err_ThisFunc
      End If
      If (out_HitCancel = True) Then GoTo exit_normally_ThisFunc
      '
      ' UPDATE UNITS OF DISPLAY FOR THIS PROPERTY.
      '
      ''''For i = 1 To UBound(NowProj.UserChemicals)
      ''''  With NowProj.UserChemicals(i). _
      ''''      PropertyData(idx_PropertyData)
      ''''    .UnitDisplayed = inout_UnitDisplayed
      ''''  End With
      ''''Next i
      With NowProj.UserChemicals(This_idx_Chemical). _
          PropertyData(idx_PropertyData)
        .UnitDisplayed = inout_UnitDisplayed
      End With
      '
      ' THROW DIRTY FLAG AND REFRESH WINDOW.
      '
      Call Local_DirtyStatus_Set(Project_Is_Dirty, True)
      ''''Call frmMain_Populate_lstUser
      ''''Call frmMain_Refresh
      Call frmMain_Populate_lvMain
      lvMain.SetFocus
    '
    '////////////////////////////////////////////////////////////////////
    Case 20:          'VIEW TECHNIQUES WINDOW.
''''Call debug_output("mnuPropertyItem_Click: " & _
    "This_idx_Chemical = " & Trim$(Str$(This_idx_Chemical)) & ", " & _
    "This_Property_Code = " & Trim$(Str$(This_Property_Code)) & ".")
      If (frmTechniques.frmTechniques_Go( _
          This_idx_Chemical, _
          This_Property_Code, _
          out_HitCancel) = False) Then
        GoTo exit_err_ThisFunc
      End If
      If (out_HitCancel = True) Then GoTo exit_normally_ThisFunc
      '
      ' THROW DIRTY FLAG AND REFRESH WINDOW.
      '
      Call Local_DirtyStatus_Set(Project_Is_Dirty, True)
      ''''Call frmMain_Populate_lstUser
      ''''Call frmMain_Refresh
      Call frmMain_Populate_lvMain
      'lvMain.SetFocus
  End Select
exit_normally_ThisFunc:
  'xxxxx = True
  Exit Sub
exit_err_ThisFunc:
  'xxxxx = False
  Exit Sub
err_ThisFunc:
  Call Show_Trapped_Error("mnuPropertyItem_Click")
  GoTo exit_err_ThisFunc
End Sub


Private Sub txtData_GotFocus(Index As Integer)
Dim Ctl As Control
Set Ctl = txtdata(Index)
Dim StatusMessagePanel As String
  Call unitsys_control_txtx_gotfocus(Ctl)
  Select Case Index
''''    Case 0
''''      StatusMessagePanel = "Type in the bed diameter"
''''    Case 1
''''      StatusMessagePanel = "Type in the bed length"
''''    Case 2
''''      StatusMessagePanel = "Type in the mass of adsorbent in the bed"
''''    Case 3
''''      StatusMessagePanel = "Type in the inlet flowrate"
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
Set Ctl = txtdata(Index)
Dim Val_Low As Double
Dim Val_High As Double
Dim Raise_Dirty_Flag As Boolean
Dim Too_Small As Integer
Dim Recalc_Needed As Boolean
  '
  ' NOTE: LOW AND HIGH VALUES IN BASE UNITS
  '
  Select Case Index
    Case 0: Val_Low = 1E-20: Val_High = 1E+20
    Case 1: Val_Low = 1E-20: Val_High = 1E+20
  End Select
  NewValue_Okay = False
  If (unitsys_control_txtx_lostfocus_validate(Ctl, Val_Low, Val_High, NewValue, Raise_Dirty_Flag)) Then
    NewValue_Okay = True
  End If
  Call unitsys_control_txtx_lostfocus(Ctl, NewValue)
  Call Local_GenericStatus_Set("")
  If (NewValue_Okay) Then
    If (Raise_Dirty_Flag) Then
      '
      ' STORE TO MEMORY.
      '
      Recalc_Needed = False
      Select Case Index
        Case 0:
          NowProj.Op_T = NewValue
          Recalc_Needed = True
        Case 1:
          NowProj.Op_P = NewValue
          Recalc_Needed = True
      End Select
      If (Raise_Dirty_Flag) Then
        '
        ' THROW DIRTY FLAG.
        '
        Call Local_DirtyStatus_Set(Project_Is_Dirty, True)
      End If
      '
      ' REFRESH WINDOW.
      '
      Call frmMain_Refresh
      '
      ' IF RECALCULATION NEEDED (DUE TO T- OR P- CHANGE), DO IT.
      '
      If (Recalc_Needed = True) Then
        Call mnuEditItem_Click(20)
      End If
    End If
  End If
End Sub


Private Sub txtDataStr_GotFocus(Index As Integer)
Dim Ctl As Control
Set Ctl = txtDataStr(Index)
Dim StatusMessagePanel As String
  If (Ctl.Locked) Then Exit Sub
  Call Global_GotFocus(Ctl)
  Select Case Index
    Case 0:
      ''''StatusMessagePanel = "Type in the component name"
  End Select
  Call Local_GenericStatus_Set(StatusMessagePanel)
End Sub
Private Sub txtDataStr_KeyPress(Index As Integer, KeyAscii As Integer)
  KeyAscii = Global_TextKeyPress(KeyAscii)
End Sub
Private Sub txtDataStr_LostFocus(Index As Integer)
Dim Frm As Form
Set Frm = frmMain
Dim Ctl As Control
Set Ctl = txtDataStr(Index)
'''''
'Dim NewValue_Okay As Integer
'Dim NewValue As Double
'Dim Val_Low As Double
'Dim Val_High As Double
'Dim Raise_Dirty_Flag As Boolean
'Dim Too_Small As Integer
Dim Old_ItemData_lstUser As Integer
Dim DataChanged As Boolean
  If (Ctl.Locked) Then Exit Sub
  Old_ItemData_lstUser = -1
  If (Frm.lstUser.ListIndex >= 0) Then
    Old_ItemData_lstUser = _
        Frm.lstUser.ItemData(Frm.lstUser.ListIndex)
  End If
  If (Old_ItemData_lstUser = -1) Then
    Exit Sub
  End If
  If (Trim$(Ctl.Text) = "") Then
    With NowProj.UserChemicals(Old_ItemData_lstUser)
      Select Case Index
        Case 0: Ctl.Text = .Name
        Case 1: Ctl.Text = .CAS
        Case 2: Ctl.Text = .SMILES
        Case 3: Ctl.Text = .Formula
        Case 4: Ctl.Text = .Family
        Case 5: Ctl.Text = .Source
      End Select
    End With
    'Call Show_Error("You must enter a non-blank string for the component name.")
    'NOTE: SHOWING THIS ERROR MESSAGE MESSES UP THE
    'SUBSEQUENT GOTFOCUS IF THE USER HITS <Enter> OR <Tab>.
  Else
    With NowProj.UserChemicals(Old_ItemData_lstUser)
      DataChanged = False
      Select Case Index
        Case 0: DataChanged = IIf(.Name <> Trim$(Ctl.Text), True, False)
        Case 1: DataChanged = IIf(.CAS <> Trim$(Ctl.Text), True, False)
        Case 2: DataChanged = IIf(.SMILES <> Trim$(Ctl.Text), True, False)
        Case 3: DataChanged = IIf(.Formula <> Trim$(Ctl.Text), True, False)
        Case 4: DataChanged = IIf(.Family <> Trim$(Ctl.Text), True, False)
        Case 5: DataChanged = IIf(.Source <> Trim$(Ctl.Text), True, False)
      End Select
      If (DataChanged = True) Then
        Select Case Index
          Case 0: .Name = Trim$(Ctl.Text)
          Case 1: .CAS = Trim$(Ctl.Text)
          Case 2: .SMILES = Trim$(Ctl.Text)
          Case 3: .Formula = Trim$(Ctl.Text)
          Case 4: .Family = Trim$(Ctl.Text)
          Case 5: .Source = Trim$(Ctl.Text)
        End Select
        '
        ' THROW DIRTY FLAG.
        '
        Call Local_DirtyStatus_Set(Project_Is_Dirty, True)
      End If
    End With
  End If
  Call Global_LostFocus(Ctl)
  Call Local_GenericStatus_Set("")
  Exit Sub
End Sub




