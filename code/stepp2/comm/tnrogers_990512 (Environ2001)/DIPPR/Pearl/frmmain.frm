VERSION 5.00
Object = "{3B008041-905A-11D1-B4AE-444553540000}#1.0#0"; "VSOCX6.OCX"
Object = "{FAEEE763-117E-101B-8933-08002B2F4F5A}#1.1#0"; "DBLIST32.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.1#0"; "COMDLG32.OCX"
Begin VB.Form frmmain 
   Appearance      =   0  'Flat
   BackColor       =   &H00C0C0C0&
   Caption         =   "PEARLS"
   ClientHeight    =   6510
   ClientLeft      =   465
   ClientTop       =   735
   ClientWidth     =   9810
   FillStyle       =   0  'Solid
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
   Icon            =   "frmmain.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   6510
   ScaleWidth      =   9810
   Visible         =   0   'False
   WhatsThisButton =   -1  'True
   WhatsThisHelp   =   -1  'True
   WindowState     =   2  'Maximized
   Begin VB.Data Data2 
      Caption         =   "Data2"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   3120
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "User List"
      Top             =   1500
      Visible         =   0   'False
      Width           =   1980
   End
   Begin VB.Data Data1 
      Caption         =   "Data1"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   1080
      Options         =   0
      ReadOnly        =   -1  'True
      RecordsetType   =   2  'Snapshot
      RecordSource    =   "PEARLS List"
      Top             =   1500
      Visible         =   0   'False
      Width           =   1980
   End
   Begin MSComDlg.CommonDialog CODFilePath 
      Left            =   180
      Top             =   1440
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   327681
      CancelError     =   -1  'True
   End
   Begin VB.CommandButton Calculate_button 
      Caption         =   "Calculate"
      Height          =   375
      Left            =   6720
      TabIndex        =   25
      Top             =   1560
      Width           =   1095
   End
   Begin VB.CommandButton CMDRemoveAll 
      Caption         =   "<<"
      Height          =   255
      Left            =   4500
      TabIndex        =   24
      Top             =   1140
      Width           =   735
   End
   Begin VB.CommandButton CMDRemove 
      Caption         =   "<"
      Height          =   255
      Left            =   4500
      TabIndex        =   22
      Top             =   780
      Width           =   735
   End
   Begin VB.CommandButton CMDAdd 
      Caption         =   ">"
      Height          =   255
      Left            =   4500
      TabIndex        =   21
      Top             =   420
      Width           =   735
   End
   Begin MSDBCtls.DBList LSTUserList 
      Bindings        =   "frmmain.frx":030A
      Height          =   1035
      Left            =   5400
      TabIndex        =   19
      Top             =   360
      Width           =   4215
      _ExtentX        =   7435
      _ExtentY        =   1826
      _Version        =   327681
      MatchEntry      =   -1  'True
      BackColor       =   -2147483643
      ListField       =   "Name"
      BoundColumn     =   "Name"
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
   Begin MSDBCtls.DBList LSTSelList 
      Bindings        =   "frmmain.frx":031A
      DataSource      =   "Data1"
      Height          =   1035
      Left            =   120
      TabIndex        =   18
      Top             =   360
      Width           =   4215
      _ExtentX        =   7435
      _ExtentY        =   1826
      _Version        =   327681
      MatchEntry      =   -1  'True
      BackColor       =   -2147483643
      ListField       =   "Name"
      BoundColumn     =   "Name"
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
   Begin vsOcx6LibCtl.vsIndexTab TABViewProp 
      Height          =   4545
      Left            =   120
      TabIndex        =   1
      Top             =   1800
      Width           =   9495
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Enabled         =   -1  'True
      Appearance      =   1
      MousePointer    =   0
      _ConvInfo       =   1
      Version         =   600
      BackColor       =   -2147483633
      ForeColor       =   -2147483630
      FrontTabColor   =   12632256
      BackTabColor    =   -2147483633
      TabOutlineColor =   0
      FrontTabForeColor=   -2147483630
      Caption         =   $"frmmain.frx":032A
      Align           =   0
      Appearance      =   1
      CurrTab         =   6
      FirstTab        =   6
      Style           =   0
      Position        =   0
      AutoSwitch      =   -1  'True
      AutoScroll      =   -1  'True
      TabPreview      =   -1  'True
      ShowFocusRect   =   -1  'True
      TabsPerPage     =   2
      BorderWidth     =   0
      BoldCurrent     =   0   'False
      DogEars         =   -1  'True
      MultiRow        =   -1  'True
      MultiRowOffset  =   200
      CaptionStyle    =   0
      TabHeight       =   0
      Begin vsOcx6LibCtl.vsElastic VideoSoftElastic1 
         Height          =   3030
         Index           =   0
         Left            =   37545
         TabIndex        =   157
         TabStop         =   0   'False
         Top             =   1470
         Width           =   8625
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Enabled         =   -1  'True
         Appearance      =   1
         MousePointer    =   0
         _ConvInfo       =   1
         Version         =   600
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         FloodColor      =   192
         ForeColorDisabled=   -2147483631
         Caption         =   ""
         Align           =   0
         Appearance      =   1
         AutoSizeChildren=   0
         BorderWidth     =   0
         ChildSpacing    =   4
         Splitter        =   0   'False
         FloodDirection  =   0
         FloodPercent    =   0
         CaptionPos      =   1
         WordWrap        =   -1  'True
         MaxChildSize    =   0
         MinChildSize    =   0
         TagWidth        =   0
         TagPosition     =   0
         Style           =   0
         TagSplit        =   0   'False
         PicturePos      =   4
         CaptionStyle    =   0
         ResizeFonts     =   0   'False
         GridRows        =   0
         GridCols        =   0
         _GridInfo       =   ""
         Begin VB.Label PNLPropName 
            Alignment       =   1  'Right Justify
            Caption         =   "Fathead Minnow, 96h, EC50"
            Height          =   255
            Index           =   50
            Left            =   0
            TabIndex        =   190
            Top             =   480
            Width           =   3735
         End
         Begin VB.Label PNLPropUnits 
            Caption         =   "Label1"
            Height          =   255
            Index           =   42
            Left            =   5880
            TabIndex        =   189
            Top             =   480
            Width           =   2175
         End
         Begin VB.Label PNLPropVal 
            Alignment       =   2  'Center
            BackColor       =   &H80000009&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Label1"
            Height          =   255
            Index           =   42
            Left            =   3960
            TabIndex        =   188
            Top             =   480
            Width           =   1695
         End
         Begin VB.Label PNLPropName 
            Alignment       =   1  'Right Justify
            Caption         =   "Fathead Minnow, 48h, EC50"
            Height          =   255
            Index           =   49
            Left            =   0
            TabIndex        =   187
            Top             =   120
            Width           =   3735
         End
         Begin VB.Label PNLPropUnits 
            Caption         =   "Label1"
            Height          =   255
            Index           =   41
            Left            =   5880
            TabIndex        =   186
            Top             =   120
            Width           =   2175
         End
         Begin VB.Label PNLPropVal 
            Alignment       =   2  'Center
            BackColor       =   &H80000009&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Label1"
            Height          =   255
            Index           =   41
            Left            =   3960
            TabIndex        =   185
            Top             =   120
            Width           =   1695
         End
         Begin VB.Label PNLPropName 
            Alignment       =   1  'Right Justify
            Caption         =   "Fathead Minnow, 24h, LC50"
            Height          =   255
            Index           =   51
            Left            =   0
            TabIndex        =   184
            Top             =   840
            Width           =   3735
         End
         Begin VB.Label PNLPropUnits 
            Caption         =   "Label1"
            Height          =   255
            Index           =   43
            Left            =   5880
            TabIndex        =   183
            Top             =   840
            Width           =   2175
         End
         Begin VB.Label PNLPropVal 
            Alignment       =   2  'Center
            BackColor       =   &H80000009&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Label1"
            Height          =   255
            Index           =   43
            Left            =   3960
            TabIndex        =   182
            Top             =   840
            Width           =   1695
         End
         Begin VB.Label PNLPropVal 
            Alignment       =   2  'Center
            BackColor       =   &H80000009&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Label1"
            Height          =   255
            Index           =   44
            Left            =   3960
            TabIndex        =   181
            Top             =   1200
            Width           =   1695
         End
         Begin VB.Label PNLPropUnits 
            Caption         =   "Label1"
            Height          =   255
            Index           =   44
            Left            =   5880
            TabIndex        =   180
            Top             =   1200
            Width           =   2175
         End
         Begin VB.Label PNLPropName 
            Alignment       =   1  'Right Justify
            Caption         =   "Fathead Minnow, 48h, LC50"
            Height          =   255
            Index           =   52
            Left            =   0
            TabIndex        =   179
            Top             =   1200
            Width           =   3735
         End
         Begin VB.Label PNLPropVal 
            Alignment       =   2  'Center
            BackColor       =   &H80000009&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Label1"
            Height          =   255
            Index           =   45
            Left            =   3960
            TabIndex        =   178
            Top             =   1560
            Width           =   1695
         End
         Begin VB.Label PNLPropUnits 
            Caption         =   "Label1"
            Height          =   255
            Index           =   45
            Left            =   5880
            TabIndex        =   177
            Top             =   1560
            Width           =   2175
         End
         Begin VB.Label PNLPropName 
            Alignment       =   1  'Right Justify
            Caption         =   "Fathead Minnow, 96h, LC50"
            Height          =   255
            Index           =   53
            Left            =   0
            TabIndex        =   176
            Top             =   1560
            Width           =   3735
         End
         Begin VB.Label PNLPropVal 
            Alignment       =   2  'Center
            BackColor       =   &H80000009&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Label1"
            Height          =   255
            Index           =   46
            Left            =   3960
            TabIndex        =   175
            Top             =   1920
            Width           =   1695
         End
         Begin VB.Label PNLPropUnits 
            Caption         =   "Label1"
            Height          =   255
            Index           =   46
            Left            =   5880
            TabIndex        =   174
            Top             =   1920
            Width           =   2175
         End
         Begin VB.Label PNLPropName 
            Alignment       =   1  'Right Justify
            Caption         =   "Salmonidae, 24h, LC50"
            Height          =   255
            Index           =   54
            Left            =   0
            TabIndex        =   173
            Top             =   1920
            Width           =   3735
         End
         Begin VB.Label PNLPropVal 
            Alignment       =   2  'Center
            BackColor       =   &H80000009&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Label1"
            Height          =   255
            Index           =   47
            Left            =   3960
            TabIndex        =   172
            Top             =   2280
            Width           =   1695
         End
         Begin VB.Label PNLPropUnits 
            Caption         =   "Label1"
            Height          =   255
            Index           =   47
            Left            =   5880
            TabIndex        =   171
            Top             =   2280
            Width           =   2175
         End
         Begin VB.Label PNLPropName 
            Alignment       =   1  'Right Justify
            Caption         =   "Salmonidae, 48h, LC50"
            Height          =   255
            Index           =   55
            Left            =   0
            TabIndex        =   170
            Top             =   2280
            Width           =   3735
         End
         Begin VB.Label PNLPropVal 
            Alignment       =   2  'Center
            BackColor       =   &H80000009&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Label1"
            Height          =   255
            Index           =   48
            Left            =   3960
            TabIndex        =   169
            Top             =   2640
            Width           =   1695
         End
         Begin VB.Label PNLPropUnits 
            Caption         =   "Label1"
            Height          =   255
            Index           =   48
            Left            =   5880
            TabIndex        =   168
            Top             =   2640
            Width           =   2175
         End
         Begin VB.Label PNLPropName 
            Alignment       =   1  'Right Justify
            Caption         =   "Salmonidae, 96h, LC50"
            Height          =   255
            Index           =   56
            Left            =   0
            TabIndex        =   167
            Top             =   2640
            Width           =   3735
         End
      End
      Begin vsOcx6LibCtl.vsElastic VSElastic2 
         Height          =   3030
         Left            =   45
         TabIndex        =   8
         TabStop         =   0   'False
         Top             =   1470
         Width           =   8625
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Enabled         =   -1  'True
         Appearance      =   0
         MousePointer    =   0
         _ConvInfo       =   1
         Version         =   600
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         FloodColor      =   192
         ForeColorDisabled=   -2147483631
         Caption         =   ""
         Align           =   0
         Appearance      =   0
         AutoSizeChildren=   0
         BorderWidth     =   6
         ChildSpacing    =   4
         Splitter        =   0   'False
         FloodDirection  =   0
         FloodPercent    =   0
         CaptionPos      =   1
         WordWrap        =   -1  'True
         MaxChildSize    =   0
         MinChildSize    =   0
         TagWidth        =   0
         TagPosition     =   0
         Style           =   0
         TagSplit        =   0   'False
         PicturePos      =   4
         CaptionStyle    =   0
         ResizeFonts     =   0   'False
         GridRows        =   0
         GridCols        =   0
         _GridInfo       =   ""
         Begin VB.TextBox TXTFamily 
            Height          =   285
            Left            =   3960
            Locked          =   -1  'True
            TabIndex        =   23
            Text            =   "TXTFamily"
            Top             =   1920
            Width           =   3135
         End
         Begin VB.TextBox TXTSMILES 
            DataField       =   "Smiles"
            DataSource      =   "Data1"
            Height          =   285
            Left            =   3960
            Locked          =   -1  'True
            TabIndex        =   17
            Text            =   "TXTSMILES"
            Top             =   2640
            Width           =   3135
         End
         Begin VB.TextBox TXTOpP 
            Height          =   285
            Left            =   3960
            TabIndex        =   14
            Text            =   "TXTOpP"
            Top             =   480
            Width           =   1335
         End
         Begin VB.TextBox TXTOpT 
            Height          =   285
            Left            =   3960
            TabIndex        =   13
            Text            =   "TXTOpT"
            Top             =   120
            Width           =   1335
         End
         Begin VB.TextBox TXTSource 
            DataField       =   "Source"
            DataSource      =   "Data1"
            Height          =   285
            Left            =   3960
            Locked          =   -1  'True
            TabIndex        =   12
            Text            =   "TXTSource"
            Top             =   2280
            Width           =   3135
         End
         Begin VB.TextBox TXTFormula 
            DataField       =   "Formula"
            DataSource      =   "Data1"
            Height          =   285
            Left            =   3960
            Locked          =   -1  'True
            TabIndex        =   11
            Text            =   "TXTFormula"
            Top             =   1560
            Width           =   3135
         End
         Begin VB.TextBox TXTCAS 
            DataField       =   "CAS"
            DataSource      =   "Data1"
            Height          =   285
            Left            =   3960
            Locked          =   -1  'True
            TabIndex        =   10
            Text            =   "TXTCAS"
            Top             =   1200
            Width           =   3135
         End
         Begin VB.TextBox TXTName 
            DataField       =   "Name"
            DataSource      =   "Data1"
            Height          =   285
            Left            =   3960
            Locked          =   -1  'True
            TabIndex        =   9
            Text            =   "TXTName"
            Top             =   840
            Width           =   4575
         End
         Begin VB.Label PNLPropName 
            Alignment       =   1  'Right Justify
            Caption         =   "SMILES"
            Height          =   255
            Index           =   48
            Left            =   0
            TabIndex        =   33
            Top             =   2640
            Width           =   3735
         End
         Begin VB.Label PNLPropName 
            Alignment       =   1  'Right Justify
            Caption         =   "Source"
            Height          =   255
            Index           =   47
            Left            =   0
            TabIndex        =   32
            Top             =   2280
            Width           =   3735
         End
         Begin VB.Label PNLPropName 
            Alignment       =   1  'Right Justify
            Caption         =   "Family"
            Height          =   255
            Index           =   46
            Left            =   0
            TabIndex        =   31
            Top             =   1920
            Width           =   3735
         End
         Begin VB.Label PNLPropName 
            Alignment       =   1  'Right Justify
            Caption         =   "Formula"
            Height          =   255
            Index           =   45
            Left            =   0
            TabIndex        =   30
            Top             =   1560
            Width           =   3735
         End
         Begin VB.Label PNLPropName 
            Alignment       =   1  'Right Justify
            Caption         =   "Operating Pressure"
            Height          =   255
            Index           =   42
            Left            =   0
            TabIndex        =   29
            Top             =   480
            Width           =   3735
         End
         Begin VB.Label PNLPropName 
            Alignment       =   1  'Right Justify
            Caption         =   "Name"
            Height          =   255
            Index           =   43
            Left            =   0
            TabIndex        =   28
            Top             =   840
            Width           =   3735
         End
         Begin VB.Label PNLPropName 
            Alignment       =   1  'Right Justify
            Caption         =   "CAS"
            Height          =   255
            Index           =   44
            Left            =   0
            TabIndex        =   27
            Top             =   1200
            Width           =   3735
         End
         Begin VB.Label PNLPropName 
            Alignment       =   1  'Right Justify
            Caption         =   "Operating Temperature"
            Height          =   255
            Index           =   41
            Left            =   0
            TabIndex        =   26
            Top             =   120
            Width           =   3735
         End
         Begin VB.Label LBLOpPUnits 
            Appearance      =   0  'Flat
            BackColor       =   &H00C0C0C0&
            Caption         =   "LBLOpPUnits"
            ForeColor       =   &H80000008&
            Height          =   255
            Left            =   5400
            TabIndex        =   16
            Top             =   510
            Width           =   1215
         End
         Begin VB.Label LBLOpTUnits 
            Appearance      =   0  'Flat
            BackColor       =   &H00C0C0C0&
            Caption         =   "LBLOpTUnits"
            ForeColor       =   &H80000008&
            Height          =   255
            Left            =   5400
            TabIndex        =   15
            Top             =   150
            Width           =   1215
         End
      End
      Begin vsOcx6LibCtl.vsElastic VSElastic1 
         Height          =   3030
         Index           =   5
         Left            =   -35955
         TabIndex        =   7
         TabStop         =   0   'False
         Top             =   1470
         Width           =   8625
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Enabled         =   -1  'True
         Appearance      =   1
         MousePointer    =   0
         _ConvInfo       =   1
         Version         =   600
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         FloodColor      =   192
         ForeColorDisabled=   -2147483631
         Caption         =   ""
         Align           =   0
         Appearance      =   1
         AutoSizeChildren=   0
         BorderWidth     =   0
         ChildSpacing    =   4
         Splitter        =   0   'False
         FloodDirection  =   0
         FloodPercent    =   0
         CaptionPos      =   1
         WordWrap        =   -1  'True
         MaxChildSize    =   0
         MinChildSize    =   0
         TagWidth        =   0
         TagPosition     =   0
         Style           =   0
         TagSplit        =   0   'False
         PicturePos      =   4
         CaptionStyle    =   0
         ResizeFonts     =   0   'False
         GridRows        =   0
         GridCols        =   0
         _GridInfo       =   ""
         Begin VB.Label PNLPropVal 
            Alignment       =   2  'Center
            BackColor       =   &H80000009&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Label1"
            Height          =   255
            Index           =   40
            Left            =   3960
            TabIndex        =   156
            Top             =   1560
            Width           =   1695
         End
         Begin VB.Label PNLPropVal 
            Alignment       =   2  'Center
            BackColor       =   &H80000009&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "PNLPropVal"
            Height          =   255
            Index           =   39
            Left            =   3960
            TabIndex        =   155
            Top             =   1200
            Width           =   1695
         End
         Begin VB.Label PNLPropVal 
            Alignment       =   2  'Center
            BackColor       =   &H80000009&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Label1"
            Height          =   255
            Index           =   37
            Left            =   3960
            TabIndex        =   154
            Top             =   2640
            Width           =   1695
         End
         Begin VB.Label PNLPropVal 
            Alignment       =   2  'Center
            BackColor       =   &H80000009&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Label1"
            Height          =   255
            Index           =   36
            Left            =   3960
            TabIndex        =   153
            Top             =   2280
            Width           =   1695
         End
         Begin VB.Label PNLPropVal 
            Alignment       =   2  'Center
            BackColor       =   &H80000009&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Label1"
            Height          =   255
            Index           =   35
            Left            =   3960
            TabIndex        =   152
            Top             =   1920
            Width           =   1695
         End
         Begin VB.Label PNLPropVal 
            Alignment       =   2  'Center
            BackColor       =   &H80000009&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Label1"
            Height          =   255
            Index           =   34
            Left            =   3960
            TabIndex        =   151
            Top             =   120
            Width           =   1695
         End
         Begin VB.Label PNLPropVal 
            Alignment       =   2  'Center
            BackColor       =   &H80000009&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Label1"
            Height          =   255
            Index           =   33
            Left            =   3960
            TabIndex        =   150
            Top             =   840
            Width           =   1695
         End
         Begin VB.Label PNLPropVal 
            Alignment       =   2  'Center
            BackColor       =   &H80000009&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "PNLPropVal"
            Height          =   255
            Index           =   32
            Left            =   3960
            TabIndex        =   149
            Top             =   480
            Width           =   1695
         End
         Begin VB.Label PNLPropUnits 
            Caption         =   "Label1"
            Height          =   255
            Index           =   37
            Left            =   5880
            TabIndex        =   115
            Top             =   2640
            Width           =   2175
         End
         Begin VB.Label PNLPropUnits 
            Caption         =   "Label1"
            Height          =   255
            Index           =   36
            Left            =   5880
            TabIndex        =   114
            Top             =   2280
            Width           =   2175
         End
         Begin VB.Label PNLPropUnits 
            Caption         =   "Label1"
            Height          =   255
            Index           =   35
            Left            =   5880
            TabIndex        =   113
            Top             =   1920
            Width           =   2175
         End
         Begin VB.Label PNLPropUnits 
            Caption         =   "Label1"
            Height          =   255
            Index           =   40
            Left            =   5880
            TabIndex        =   112
            Top             =   1560
            Width           =   2175
         End
         Begin VB.Label PNLPropUnits 
            Caption         =   "Label1"
            Height          =   255
            Index           =   39
            Left            =   5880
            TabIndex        =   111
            Top             =   1200
            Width           =   2175
         End
         Begin VB.Label PNLPropUnits 
            Caption         =   "Label1"
            Height          =   255
            Index           =   33
            Left            =   5880
            TabIndex        =   110
            Top             =   840
            Width           =   2175
         End
         Begin VB.Label PNLPropUnits 
            Caption         =   "Label1"
            Height          =   255
            Index           =   34
            Left            =   5880
            TabIndex        =   109
            Top             =   120
            Width           =   2175
         End
         Begin VB.Label PNLPropUnits 
            Caption         =   "Label1"
            Height          =   255
            Index           =   32
            Left            =   5880
            TabIndex        =   108
            Top             =   480
            Width           =   2175
         End
         Begin VB.Label PNLPropName 
            Alignment       =   1  'Right Justify
            Caption         =   "Bioconcentration Factor"
            Height          =   255
            Index           =   37
            Left            =   0
            TabIndex        =   41
            Top             =   2640
            Width           =   3735
         End
         Begin VB.Label PNLPropName 
            Alignment       =   1  'Right Justify
            Caption         =   "Log Koc"
            Height          =   255
            Index           =   36
            Left            =   0
            TabIndex        =   40
            Top             =   2280
            Width           =   3735
         End
         Begin VB.Label PNLPropName 
            Alignment       =   1  'Right Justify
            Caption         =   "Log Kow"
            Height          =   255
            Index           =   35
            Left            =   0
            TabIndex        =   39
            Top             =   1920
            Width           =   3735
         End
         Begin VB.Label PNLPropName 
            Alignment       =   1  'Right Justify
            Caption         =   "Solubility Limit of Water in Chemical"
            Height          =   255
            Index           =   40
            Left            =   0
            TabIndex        =   38
            Top             =   1560
            Width           =   3735
         End
         Begin VB.Label PNLPropName 
            Alignment       =   1  'Right Justify
            Caption         =   "Solubility Limit of Chemical in Water"
            Height          =   255
            Index           =   39
            Left            =   0
            TabIndex        =   37
            Top             =   1200
            Width           =   3735
         End
         Begin VB.Label PNLPropName 
            Alignment       =   1  'Right Justify
            Caption         =   "Henry's Constant"
            Height          =   255
            Index           =   33
            Left            =   0
            TabIndex        =   36
            Top             =   840
            Width           =   3735
         End
         Begin VB.Label PNLPropName 
            Alignment       =   1  'Right Justify
            Caption         =   "Activity Coefficient of Water in Chemical"
            Height          =   255
            Index           =   32
            Left            =   0
            TabIndex        =   35
            Top             =   480
            Width           =   3735
         End
         Begin VB.Label PNLPropName 
            Alignment       =   1  'Right Justify
            Caption         =   "Activity Coefficient of Chemical in Water"
            Height          =   255
            Index           =   34
            Left            =   0
            TabIndex        =   34
            Top             =   120
            Width           =   3735
         End
      End
      Begin vsOcx6LibCtl.vsElastic VSElastic1 
         Height          =   3030
         Index           =   4
         Left            =   -36255
         TabIndex        =   6
         TabStop         =   0   'False
         Top             =   1470
         Width           =   8625
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Enabled         =   -1  'True
         Appearance      =   1
         MousePointer    =   0
         _ConvInfo       =   1
         Version         =   600
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         FloodColor      =   192
         ForeColorDisabled=   -2147483631
         Caption         =   ""
         Align           =   0
         Appearance      =   1
         AutoSizeChildren=   0
         BorderWidth     =   0
         ChildSpacing    =   4
         Splitter        =   0   'False
         FloodDirection  =   0
         FloodPercent    =   0
         CaptionPos      =   1
         WordWrap        =   -1  'True
         MaxChildSize    =   0
         MinChildSize    =   0
         TagWidth        =   0
         TagPosition     =   0
         Style           =   0
         TagSplit        =   0   'False
         PicturePos      =   4
         CaptionStyle    =   0
         ResizeFonts     =   0   'False
         GridRows        =   0
         GridCols        =   0
         _GridInfo       =   ""
         Begin VB.Label PNLPropVal 
            Alignment       =   2  'Center
            BackColor       =   &H80000009&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Label1"
            Height          =   255
            Index           =   31
            Left            =   3960
            TabIndex        =   148
            Top             =   1200
            Width           =   1695
         End
         Begin VB.Label PNLPropVal 
            Alignment       =   2  'Center
            BackColor       =   &H80000009&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Label1"
            Height          =   255
            Index           =   30
            Left            =   3960
            TabIndex        =   147
            Top             =   840
            Width           =   1695
         End
         Begin VB.Label PNLPropVal 
            Alignment       =   2  'Center
            BackColor       =   &H80000009&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Label1"
            Height          =   255
            Index           =   29
            Left            =   3960
            TabIndex        =   146
            Top             =   480
            Width           =   1695
         End
         Begin VB.Label PNLPropVal 
            Alignment       =   2  'Center
            BackColor       =   &H80000009&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Label1"
            Height          =   255
            Index           =   28
            Left            =   3960
            TabIndex        =   145
            Top             =   120
            Width           =   1695
         End
         Begin VB.Label PNLPropUnits 
            Caption         =   "Label1"
            Height          =   255
            Index           =   31
            Left            =   5880
            TabIndex        =   107
            Top             =   1200
            Width           =   2175
         End
         Begin VB.Label PNLPropUnits 
            Caption         =   "Label1"
            Height          =   255
            Index           =   30
            Left            =   5880
            TabIndex        =   106
            Top             =   840
            Width           =   2175
         End
         Begin VB.Label PNLPropUnits 
            Caption         =   "Label1"
            Height          =   255
            Index           =   29
            Left            =   5880
            TabIndex        =   105
            Top             =   480
            Width           =   2175
         End
         Begin VB.Label PNLPropUnits 
            Caption         =   "Label1"
            Height          =   255
            Index           =   28
            Left            =   5880
            TabIndex        =   104
            Top             =   120
            Width           =   2175
         End
         Begin VB.Label PNLPropName 
            Alignment       =   1  'Right Justify
            Caption         =   "Chemical Oxygen Demand"
            Height          =   255
            Index           =   30
            Left            =   0
            TabIndex        =   45
            Top             =   840
            Width           =   3735
         End
         Begin VB.Label PNLPropName 
            Alignment       =   1  'Right Justify
            Caption         =   "Carbonaceous ThOD"
            Height          =   255
            Index           =   28
            Left            =   0
            TabIndex        =   44
            Top             =   120
            Width           =   3735
         End
         Begin VB.Label PNLPropName 
            Alignment       =   1  'Right Justify
            Caption         =   "Biochemical Oxygen Demand"
            Height          =   255
            Index           =   31
            Left            =   0
            TabIndex        =   43
            Top             =   1200
            Width           =   3735
         End
         Begin VB.Label PNLPropName 
            Alignment       =   1  'Right Justify
            Caption         =   "Combined (C + N) ThOD"
            Height          =   255
            Index           =   29
            Left            =   0
            TabIndex        =   42
            Top             =   480
            Width           =   3735
         End
      End
      Begin vsOcx6LibCtl.vsElastic VSElastic1 
         Height          =   3030
         Index           =   3
         Left            =   -36555
         TabIndex        =   5
         TabStop         =   0   'False
         Top             =   1470
         Width           =   8625
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Enabled         =   -1  'True
         Appearance      =   1
         MousePointer    =   0
         _ConvInfo       =   1
         Version         =   600
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         FloodColor      =   192
         ForeColorDisabled=   -2147483631
         Caption         =   ""
         Align           =   0
         Appearance      =   1
         AutoSizeChildren=   0
         BorderWidth     =   0
         ChildSpacing    =   4
         Splitter        =   0   'False
         FloodDirection  =   0
         FloodPercent    =   0
         CaptionPos      =   1
         WordWrap        =   -1  'True
         MaxChildSize    =   0
         MinChildSize    =   0
         TagWidth        =   0
         TagPosition     =   0
         Style           =   0
         TagSplit        =   0   'False
         PicturePos      =   4
         CaptionStyle    =   0
         ResizeFonts     =   0   'False
         GridRows        =   0
         GridCols        =   0
         _GridInfo       =   ""
         Begin VB.Label PNLPropVal 
            Alignment       =   2  'Center
            BackColor       =   &H80000009&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Label1"
            Height          =   255
            Index           =   27
            Left            =   3960
            TabIndex        =   144
            Top             =   1560
            Width           =   1695
         End
         Begin VB.Label PNLPropVal 
            Alignment       =   2  'Center
            BackColor       =   &H80000009&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Label1"
            Height          =   255
            Index           =   26
            Left            =   3960
            TabIndex        =   143
            Top             =   1200
            Width           =   1695
         End
         Begin VB.Label PNLPropVal 
            Alignment       =   2  'Center
            BackColor       =   &H80000009&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Label1"
            Height          =   255
            Index           =   25
            Left            =   3960
            TabIndex        =   142
            Top             =   840
            Width           =   1695
         End
         Begin VB.Label PNLPropVal 
            Alignment       =   2  'Center
            BackColor       =   &H80000009&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Label1"
            Height          =   255
            Index           =   24
            Left            =   3960
            TabIndex        =   141
            Top             =   480
            Width           =   1695
         End
         Begin VB.Label PNLPropVal 
            Alignment       =   2  'Center
            BackColor       =   &H80000009&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Label1"
            Height          =   255
            Index           =   23
            Left            =   3960
            TabIndex        =   140
            Top             =   120
            Width           =   1695
         End
         Begin VB.Label PNLPropUnits 
            Caption         =   "Label1"
            Height          =   255
            Index           =   27
            Left            =   5880
            TabIndex        =   103
            Top             =   1560
            Width           =   2175
         End
         Begin VB.Label PNLPropUnits 
            Caption         =   "Label1"
            Height          =   255
            Index           =   26
            Left            =   5880
            TabIndex        =   102
            Top             =   1200
            Width           =   2175
         End
         Begin VB.Label PNLPropUnits 
            Caption         =   "Label1"
            Height          =   255
            Index           =   25
            Left            =   5880
            TabIndex        =   101
            Top             =   840
            Width           =   2175
         End
         Begin VB.Label PNLPropUnits 
            Caption         =   "Label1"
            Height          =   255
            Index           =   24
            Left            =   5880
            TabIndex        =   100
            Top             =   480
            Width           =   2175
         End
         Begin VB.Label PNLPropUnits 
            Caption         =   "Label1"
            Height          =   255
            Index           =   23
            Left            =   5880
            TabIndex        =   99
            Top             =   120
            Width           =   2175
         End
         Begin VB.Label PNLPropName 
            Alignment       =   1  'Right Justify
            Caption         =   "Heat of Combustion"
            Height          =   255
            Index           =   27
            Left            =   0
            TabIndex        =   50
            Top             =   1560
            Width           =   3735
         End
         Begin VB.Label PNLPropName 
            Alignment       =   1  'Right Justify
            Caption         =   "Lower Flamibility Limit"
            Height          =   255
            Index           =   24
            Left            =   0
            TabIndex        =   49
            Top             =   480
            Width           =   3735
         End
         Begin VB.Label PNLPropName 
            Alignment       =   1  'Right Justify
            Caption         =   "Upper Flamibility Limit"
            Height          =   255
            Index           =   23
            Left            =   0
            TabIndex        =   48
            Top             =   120
            Width           =   3735
         End
         Begin VB.Label PNLPropName 
            Alignment       =   1  'Right Justify
            Caption         =   "Flash Point"
            Height          =   255
            Index           =   25
            Left            =   0
            TabIndex        =   47
            Top             =   840
            Width           =   3735
         End
         Begin VB.Label PNLPropName 
            Alignment       =   1  'Right Justify
            Caption         =   "Autoignition Temperature"
            Height          =   255
            Index           =   26
            Left            =   0
            TabIndex        =   46
            Top             =   1200
            Width           =   3735
         End
      End
      Begin vsOcx6LibCtl.vsElastic VSElastic1 
         Height          =   3030
         Index           =   1
         Left            =   -37155
         TabIndex        =   3
         TabStop         =   0   'False
         Top             =   1470
         Width           =   8625
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Enabled         =   -1  'True
         Appearance      =   1
         MousePointer    =   0
         _ConvInfo       =   1
         Version         =   600
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         FloodColor      =   192
         ForeColorDisabled=   -2147483631
         Caption         =   ""
         Align           =   0
         Appearance      =   1
         AutoSizeChildren=   0
         BorderWidth     =   0
         ChildSpacing    =   4
         Splitter        =   0   'False
         FloodDirection  =   0
         FloodPercent    =   0
         CaptionPos      =   1
         WordWrap        =   -1  'True
         MaxChildSize    =   0
         MinChildSize    =   0
         TagWidth        =   0
         TagPosition     =   0
         Style           =   0
         TagSplit        =   0   'False
         PicturePos      =   4
         CaptionStyle    =   0
         ResizeFonts     =   0   'False
         GridRows        =   0
         GridCols        =   0
         _GridInfo       =   ""
         Begin VB.Label PNLPropVal 
            Alignment       =   2  'Center
            BackColor       =   &H80000009&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Label1"
            Height          =   255
            Index           =   38
            Left            =   3960
            TabIndex        =   131
            Top             =   2640
            Width           =   1695
         End
         Begin VB.Label PNLPropVal 
            Alignment       =   2  'Center
            BackColor       =   &H80000009&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Label1"
            Height          =   255
            Index           =   14
            Left            =   3960
            TabIndex        =   130
            Top             =   2280
            Width           =   1695
         End
         Begin VB.Label PNLPropVal 
            Alignment       =   2  'Center
            BackColor       =   &H80000009&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Label1"
            Height          =   255
            Index           =   13
            Left            =   3960
            TabIndex        =   129
            Top             =   1920
            Width           =   1695
         End
         Begin VB.Label PNLPropVal 
            Alignment       =   2  'Center
            BackColor       =   &H80000009&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Label1"
            Height          =   255
            Index           =   12
            Left            =   3960
            TabIndex        =   128
            Top             =   1560
            Width           =   1695
         End
         Begin VB.Label PNLPropVal 
            Alignment       =   2  'Center
            BackColor       =   &H80000009&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Label1"
            Height          =   255
            Index           =   11
            Left            =   3960
            TabIndex        =   127
            Top             =   1200
            Width           =   1695
         End
         Begin VB.Label PNLPropVal 
            Alignment       =   2  'Center
            BackColor       =   &H80000009&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Label1"
            Height          =   255
            Index           =   10
            Left            =   3960
            TabIndex        =   126
            Top             =   840
            Width           =   1695
         End
         Begin VB.Label PNLPropVal 
            Alignment       =   2  'Center
            BackColor       =   &H80000009&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Label1"
            Height          =   255
            Index           =   9
            Left            =   3960
            TabIndex        =   125
            Top             =   480
            Width           =   1695
         End
         Begin VB.Label PNLPropVal 
            Alignment       =   2  'Center
            BackColor       =   &H80000009&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Label1"
            Height          =   255
            Index           =   8
            Left            =   3960
            TabIndex        =   124
            Top             =   120
            Width           =   1695
         End
         Begin VB.Label PNLPropUnits 
            Caption         =   "Label1"
            Height          =   255
            Index           =   38
            Left            =   5880
            TabIndex        =   90
            Top             =   2640
            Width           =   2175
         End
         Begin VB.Label PNLPropUnits 
            Caption         =   "Label1"
            Height          =   255
            Index           =   14
            Left            =   5880
            TabIndex        =   89
            Top             =   2280
            Width           =   2175
         End
         Begin VB.Label PNLPropUnits 
            Caption         =   "Label1"
            Height          =   255
            Index           =   13
            Left            =   5880
            TabIndex        =   88
            Top             =   1920
            Width           =   2175
         End
         Begin VB.Label PNLPropUnits 
            Caption         =   "Label1"
            Height          =   255
            Index           =   12
            Left            =   5880
            TabIndex        =   87
            Top             =   1560
            Width           =   2175
         End
         Begin VB.Label PNLPropUnits 
            Caption         =   "Label1"
            Height          =   255
            Index           =   11
            Left            =   5880
            TabIndex        =   86
            Top             =   1200
            Width           =   2175
         End
         Begin VB.Label PNLPropUnits 
            Caption         =   "Label1"
            Height          =   255
            Index           =   10
            Left            =   5880
            TabIndex        =   85
            Top             =   840
            Width           =   2175
         End
         Begin VB.Label PNLPropUnits 
            Caption         =   "Label1"
            Height          =   255
            Index           =   9
            Left            =   5880
            TabIndex        =   84
            Top             =   480
            Width           =   2175
         End
         Begin VB.Label PNLPropUnits 
            Caption         =   "Label1"
            Height          =   255
            Index           =   8
            Left            =   5880
            TabIndex        =   83
            Top             =   120
            Width           =   2175
         End
         Begin VB.Label PNLPropName 
            Alignment       =   1  'Right Justify
            Caption         =   "Critical Volume"
            Height          =   255
            Index           =   38
            Left            =   0
            TabIndex        =   66
            Top             =   2640
            Width           =   3735
         End
         Begin VB.Label PNLPropName 
            Alignment       =   1  'Right Justify
            Caption         =   "Critical Pressure"
            Height          =   255
            Index           =   14
            Left            =   0
            TabIndex        =   65
            Top             =   2280
            Width           =   3735
         End
         Begin VB.Label PNLPropName 
            Alignment       =   1  'Right Justify
            Caption         =   "Critical Temperature"
            Height          =   255
            Index           =   13
            Left            =   0
            TabIndex        =   64
            Top             =   1920
            Width           =   3735
         End
         Begin VB.Label PNLPropName 
            Alignment       =   1  'Right Justify
            Caption         =   "Heat of Vaporization as f(T)"
            Height          =   255
            Index           =   12
            Left            =   0
            TabIndex        =   63
            Top             =   1560
            Width           =   3735
         End
         Begin VB.Label PNLPropName 
            Alignment       =   1  'Right Justify
            Caption         =   "Heat of Vaporization  @ NBP"
            Height          =   255
            Index           =   11
            Left            =   0
            TabIndex        =   62
            Top             =   1200
            Width           =   3735
         End
         Begin VB.Label PNLPropName 
            Alignment       =   1  'Right Justify
            Caption         =   "Heat of Vaporization @ 298.15 K"
            Height          =   255
            Index           =   10
            Left            =   0
            TabIndex        =   61
            Top             =   840
            Width           =   3735
         End
         Begin VB.Label PNLPropName 
            Alignment       =   1  'Right Justify
            Caption         =   "Vapor Heat Capacity as f(T)"
            Height          =   255
            Index           =   9
            Left            =   0
            TabIndex        =   60
            Top             =   480
            Width           =   3735
         End
         Begin VB.Label PNLPropName 
            Alignment       =   1  'Right Justify
            Caption         =   "Liquid Heat Capacity as f(T)"
            Height          =   255
            Index           =   8
            Left            =   0
            TabIndex        =   59
            Top             =   120
            Width           =   3735
         End
      End
      Begin vsOcx6LibCtl.vsElastic VSElastic1 
         Height          =   3030
         Index           =   0
         Left            =   -36855
         TabIndex        =   2
         TabStop         =   0   'False
         Top             =   1470
         Width           =   8625
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Enabled         =   -1  'True
         Appearance      =   1
         MousePointer    =   0
         _ConvInfo       =   1
         Version         =   600
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         FloodColor      =   192
         ForeColorDisabled=   -2147483631
         Caption         =   ""
         Align           =   0
         Appearance      =   1
         AutoSizeChildren=   0
         BorderWidth     =   0
         ChildSpacing    =   4
         Splitter        =   0   'False
         FloodDirection  =   0
         FloodPercent    =   0
         CaptionPos      =   1
         WordWrap        =   -1  'True
         MaxChildSize    =   0
         MinChildSize    =   0
         TagWidth        =   0
         TagPosition     =   0
         Style           =   0
         TagSplit        =   0   'False
         PicturePos      =   4
         CaptionStyle    =   0
         ResizeFonts     =   0   'False
         GridRows        =   0
         GridCols        =   0
         _GridInfo       =   ""
         Begin VB.Label PNLPropVal 
            Alignment       =   2  'Center
            BackColor       =   &H80000009&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Label1"
            Height          =   255
            Index           =   7
            Left            =   3960
            TabIndex        =   123
            Top             =   2640
            Width           =   1695
         End
         Begin VB.Label PNLPropVal 
            Alignment       =   2  'Center
            BackColor       =   &H80000009&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Label1"
            Height          =   255
            Index           =   6
            Left            =   3960
            TabIndex        =   122
            Top             =   2280
            Width           =   1695
         End
         Begin VB.Label PNLPropVal 
            Alignment       =   2  'Center
            BackColor       =   &H80000009&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Label1"
            Height          =   255
            Index           =   5
            Left            =   3960
            TabIndex        =   121
            Top             =   1920
            Width           =   1695
         End
         Begin VB.Label PNLPropVal 
            Alignment       =   2  'Center
            BackColor       =   &H80000009&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Label1"
            Height          =   255
            Index           =   4
            Left            =   3960
            TabIndex        =   120
            Top             =   1560
            Width           =   1695
         End
         Begin VB.Label PNLPropVal 
            Alignment       =   2  'Center
            BackColor       =   &H80000009&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Label1"
            Height          =   255
            Index           =   3
            Left            =   3960
            TabIndex        =   119
            Top             =   1200
            Width           =   1695
         End
         Begin VB.Label PNLPropVal 
            Alignment       =   2  'Center
            BackColor       =   &H80000009&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Label1"
            Height          =   255
            Index           =   2
            Left            =   3960
            TabIndex        =   118
            Top             =   840
            Width           =   1695
         End
         Begin VB.Label PNLPropVal 
            Alignment       =   2  'Center
            BackColor       =   &H80000009&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Label1"
            Height          =   255
            Index           =   1
            Left            =   3960
            TabIndex        =   117
            Top             =   480
            Width           =   1695
         End
         Begin VB.Label PNLPropVal 
            Alignment       =   2  'Center
            BackColor       =   &H80000009&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Label1"
            Height          =   255
            Index           =   0
            Left            =   3960
            TabIndex        =   116
            Top             =   120
            Width           =   1695
         End
         Begin VB.Label PNLPropUnits 
            Caption         =   "Label1"
            Height          =   255
            Index           =   7
            Left            =   5880
            TabIndex        =   82
            Top             =   2640
            Width           =   2175
         End
         Begin VB.Label PNLPropUnits 
            Caption         =   "Label1"
            Height          =   255
            Index           =   6
            Left            =   5880
            TabIndex        =   81
            Top             =   2280
            Width           =   2175
         End
         Begin VB.Label PNLPropUnits 
            Caption         =   "Label1"
            Height          =   255
            Index           =   5
            Left            =   5880
            TabIndex        =   80
            Top             =   1920
            Width           =   2175
         End
         Begin VB.Label PNLPropUnits 
            Caption         =   "Label1"
            Height          =   255
            Index           =   4
            Left            =   5880
            TabIndex        =   79
            Top             =   1560
            Width           =   2175
         End
         Begin VB.Label PNLPropUnits 
            Caption         =   "Label1"
            Height          =   255
            Index           =   3
            Left            =   5880
            TabIndex        =   78
            Top             =   1200
            Width           =   2175
         End
         Begin VB.Label PNLPropUnits 
            Caption         =   "Label1"
            Height          =   255
            Index           =   2
            Left            =   5880
            TabIndex        =   77
            Top             =   840
            Width           =   2175
         End
         Begin VB.Label PNLPropUnits 
            Caption         =   "Label1"
            Height          =   255
            Index           =   1
            Left            =   5880
            TabIndex        =   76
            Top             =   480
            Width           =   2175
         End
         Begin VB.Label PNLPropUnits 
            Caption         =   "Label1"
            Height          =   255
            Index           =   0
            Left            =   5880
            TabIndex        =   75
            Top             =   120
            Width           =   2175
         End
         Begin VB.Label PNLPropName 
            Alignment       =   1  'Right Justify
            Caption         =   "Heat of Formation"
            Height          =   255
            Index           =   7
            Left            =   0
            TabIndex        =   74
            Top             =   2640
            Width           =   3735
         End
         Begin VB.Label PNLPropName 
            Alignment       =   1  'Right Justify
            Caption         =   "Molecular Weight"
            Height          =   255
            Index           =   0
            Left            =   0
            TabIndex        =   73
            Top             =   120
            Width           =   3735
         End
         Begin VB.Label PNLPropName 
            Alignment       =   1  'Right Justify
            Caption         =   "Liquid Density @ 298.15 K"
            Height          =   255
            Index           =   1
            Left            =   0
            TabIndex        =   72
            Top             =   480
            Width           =   3735
         End
         Begin VB.Label PNLPropName 
            Alignment       =   1  'Right Justify
            Caption         =   "Liquid Density as f(T)"
            Height          =   255
            Index           =   2
            Left            =   0
            TabIndex        =   71
            Top             =   840
            Width           =   3735
         End
         Begin VB.Label PNLPropName 
            Alignment       =   1  'Right Justify
            Caption         =   "Vapor Pressure @ 298.15 K"
            Height          =   255
            Index           =   5
            Left            =   0
            TabIndex        =   70
            Top             =   1920
            Width           =   3735
         End
         Begin VB.Label PNLPropName 
            Alignment       =   1  'Right Justify
            Caption         =   "Normal Boiling Point (NBP)"
            Height          =   255
            Index           =   4
            Left            =   0
            TabIndex        =   69
            Top             =   1560
            Width           =   3735
         End
         Begin VB.Label PNLPropName 
            Alignment       =   1  'Right Justify
            Caption         =   "Melting Point"
            Height          =   255
            Index           =   3
            Left            =   0
            TabIndex        =   68
            Top             =   1200
            Width           =   3735
         End
         Begin VB.Label PNLPropName 
            Alignment       =   1  'Right Justify
            Caption         =   "Vapor Pressure as f(T)"
            Height          =   255
            Index           =   6
            Left            =   0
            TabIndex        =   67
            Top             =   2280
            Width           =   3735
         End
      End
      Begin vsOcx6LibCtl.vsElastic VideoSoftElastic1 
         Height          =   3030
         Index           =   1
         Left            =   37845
         TabIndex        =   158
         TabStop         =   0   'False
         Top             =   1470
         Width           =   8625
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Enabled         =   -1  'True
         Appearance      =   1
         MousePointer    =   0
         _ConvInfo       =   1
         Version         =   600
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         FloodColor      =   192
         ForeColorDisabled=   -2147483631
         Caption         =   ""
         Align           =   0
         Appearance      =   1
         AutoSizeChildren=   0
         BorderWidth     =   0
         ChildSpacing    =   4
         Splitter        =   0   'False
         FloodDirection  =   0
         FloodPercent    =   0
         CaptionPos      =   1
         WordWrap        =   -1  'True
         MaxChildSize    =   0
         MinChildSize    =   0
         TagWidth        =   0
         TagPosition     =   0
         Style           =   0
         TagSplit        =   0   'False
         PicturePos      =   4
         CaptionStyle    =   0
         ResizeFonts     =   0   'False
         GridRows        =   0
         GridCols        =   0
         _GridInfo       =   ""
         Begin VB.Label PNLPropVal 
            Alignment       =   2  'Center
            BackColor       =   &H80000009&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Label1"
            Height          =   255
            Index           =   54
            Left            =   3960
            TabIndex        =   200
            Top             =   1920
            Width           =   1695
         End
         Begin VB.Label PNLPropUnits 
            Caption         =   "Label1"
            Height          =   255
            Index           =   54
            Left            =   5880
            TabIndex        =   199
            Top             =   1920
            Width           =   2175
         End
         Begin VB.Label PNLPropVal 
            Alignment       =   2  'Center
            BackColor       =   &H80000009&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Label1"
            Height          =   255
            Index           =   53
            Left            =   3960
            TabIndex        =   198
            Top             =   1560
            Width           =   1695
         End
         Begin VB.Label PNLPropUnits 
            Caption         =   "Label1"
            Height          =   255
            Index           =   53
            Left            =   5880
            TabIndex        =   197
            Top             =   1560
            Width           =   2175
         End
         Begin VB.Label PNLPropVal 
            Alignment       =   2  'Center
            BackColor       =   &H80000009&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Label1"
            Height          =   255
            Index           =   52
            Left            =   3960
            TabIndex        =   196
            Top             =   1200
            Width           =   1695
         End
         Begin VB.Label PNLPropUnits 
            Caption         =   "Label1"
            Height          =   255
            Index           =   52
            Left            =   5880
            TabIndex        =   195
            Top             =   1200
            Width           =   2175
         End
         Begin VB.Label PNLPropVal 
            Alignment       =   2  'Center
            BackColor       =   &H80000009&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Label1"
            Height          =   255
            Index           =   51
            Left            =   3960
            TabIndex        =   194
            Top             =   840
            Width           =   1695
         End
         Begin VB.Label PNLPropUnits 
            Caption         =   "Label1"
            Height          =   255
            Index           =   51
            Left            =   5880
            TabIndex        =   193
            Top             =   840
            Width           =   2175
         End
         Begin VB.Label PNLPropVal 
            Alignment       =   2  'Center
            BackColor       =   &H80000009&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Label1"
            Height          =   255
            Index           =   50
            Left            =   3960
            TabIndex        =   192
            Top             =   480
            Width           =   1695
         End
         Begin VB.Label PNLPropUnits 
            Caption         =   "Label1"
            Height          =   255
            Index           =   50
            Left            =   5880
            TabIndex        =   191
            Top             =   480
            Width           =   2175
         End
         Begin VB.Label PNLPropName 
            Alignment       =   1  'Right Justify
            Caption         =   "Alternate Species"
            Height          =   255
            Index           =   62
            Left            =   0
            TabIndex        =   166
            Top             =   1920
            Width           =   3735
         End
         Begin VB.Label PNLPropUnits 
            Caption         =   "Label1"
            Height          =   255
            Index           =   49
            Left            =   5880
            TabIndex        =   165
            Top             =   120
            Width           =   2175
         End
         Begin VB.Label PNLPropVal 
            Alignment       =   2  'Center
            BackColor       =   &H80000009&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Label1"
            Height          =   255
            Index           =   49
            Left            =   3960
            TabIndex        =   164
            Top             =   120
            Width           =   1695
         End
         Begin VB.Label PNLPropName 
            Alignment       =   1  'Right Justify
            Caption         =   "Mysid, 96, LC50"
            Height          =   255
            Index           =   61
            Left            =   0
            TabIndex        =   163
            Top             =   1560
            Width           =   3735
         End
         Begin VB.Label PNLPropName 
            Alignment       =   1  'Right Justify
            Caption         =   "Daphnia Magna, 48h, LC50"
            Height          =   255
            Index           =   60
            Left            =   0
            TabIndex        =   162
            Top             =   1200
            Width           =   3735
         End
         Begin VB.Label PNLPropName 
            Alignment       =   1  'Right Justify
            Caption         =   "Daphnia Magna, 24h, LC50"
            Height          =   255
            Index           =   59
            Left            =   0
            TabIndex        =   161
            Top             =   840
            Width           =   3735
         End
         Begin VB.Label PNLPropName 
            Alignment       =   1  'Right Justify
            Caption         =   "Daphnia Magna, 48h, EC50"
            Height          =   255
            Index           =   58
            Left            =   0
            TabIndex        =   160
            Top             =   480
            Width           =   3735
         End
         Begin VB.Label PNLPropName 
            Alignment       =   1  'Right Justify
            Caption         =   "Daphnia Magna, 24h, EC50"
            Height          =   255
            Index           =   57
            Left            =   0
            TabIndex        =   159
            Top             =   120
            Width           =   3735
         End
      End
      Begin vsOcx6LibCtl.vsElastic VSElastic1 
         Height          =   3030
         Index           =   2
         Left            =   -37455
         TabIndex        =   4
         TabStop         =   0   'False
         Top             =   1470
         Width           =   8625
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Enabled         =   -1  'True
         Appearance      =   1
         MousePointer    =   0
         _ConvInfo       =   1
         Version         =   600
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         FloodColor      =   192
         ForeColorDisabled=   -2147483631
         Caption         =   ""
         Align           =   0
         Appearance      =   1
         AutoSizeChildren=   0
         BorderWidth     =   0
         ChildSpacing    =   4
         Splitter        =   0   'False
         FloodDirection  =   0
         FloodPercent    =   0
         CaptionPos      =   1
         WordWrap        =   -1  'True
         MaxChildSize    =   0
         MinChildSize    =   0
         TagWidth        =   0
         TagPosition     =   0
         Style           =   0
         TagSplit        =   0   'False
         PicturePos      =   4
         CaptionStyle    =   0
         ResizeFonts     =   0   'False
         GridRows        =   0
         GridCols        =   0
         _GridInfo       =   ""
         Begin VB.Label PNLPropVal 
            Alignment       =   2  'Center
            BackColor       =   &H80000009&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Label1"
            Height          =   255
            Index           =   22
            Left            =   3960
            TabIndex        =   139
            Top             =   2640
            Width           =   1695
         End
         Begin VB.Label PNLPropVal 
            Alignment       =   2  'Center
            BackColor       =   &H80000009&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Label1"
            Height          =   255
            Index           =   21
            Left            =   3960
            TabIndex        =   138
            Top             =   2280
            Width           =   1695
         End
         Begin VB.Label PNLPropVal 
            Alignment       =   2  'Center
            BackColor       =   &H80000009&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Label1"
            Height          =   255
            Index           =   20
            Left            =   3960
            TabIndex        =   137
            Top             =   1920
            Width           =   1695
         End
         Begin VB.Label PNLPropVal 
            Alignment       =   2  'Center
            BackColor       =   &H80000009&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Label1"
            Height          =   255
            Index           =   19
            Left            =   3960
            TabIndex        =   136
            Top             =   1560
            Width           =   1695
         End
         Begin VB.Label PNLPropVal 
            Alignment       =   2  'Center
            BackColor       =   &H80000009&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Label1"
            Height          =   255
            Index           =   18
            Left            =   3960
            TabIndex        =   135
            Top             =   1200
            Width           =   1695
         End
         Begin VB.Label PNLPropVal 
            Alignment       =   2  'Center
            BackColor       =   &H80000009&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Label1"
            Height          =   255
            Index           =   17
            Left            =   3960
            TabIndex        =   134
            Top             =   840
            Width           =   1695
         End
         Begin VB.Label PNLPropVal 
            Alignment       =   2  'Center
            BackColor       =   &H80000009&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Label1"
            Height          =   255
            Index           =   16
            Left            =   3960
            TabIndex        =   133
            Top             =   480
            Width           =   1695
         End
         Begin VB.Label PNLPropVal 
            Alignment       =   2  'Center
            BackColor       =   &H80000009&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Label1"
            Height          =   255
            Index           =   15
            Left            =   3960
            TabIndex        =   132
            Top             =   120
            Width           =   1695
         End
         Begin VB.Label PNLPropUnits 
            Caption         =   "Label1"
            Height          =   255
            Index           =   22
            Left            =   5880
            TabIndex        =   98
            Top             =   2640
            Width           =   2175
         End
         Begin VB.Label PNLPropUnits 
            Caption         =   "Label1"
            Height          =   255
            Index           =   21
            Left            =   5880
            TabIndex        =   97
            Top             =   2280
            Width           =   2175
         End
         Begin VB.Label PNLPropUnits 
            Caption         =   "Label1"
            Height          =   255
            Index           =   20
            Left            =   5880
            TabIndex        =   96
            Top             =   1920
            Width           =   2175
         End
         Begin VB.Label PNLPropUnits 
            Caption         =   "Label1"
            Height          =   255
            Index           =   19
            Left            =   5880
            TabIndex        =   95
            Top             =   1560
            Width           =   2175
         End
         Begin VB.Label PNLPropUnits 
            Caption         =   "Label1"
            Height          =   255
            Index           =   18
            Left            =   5880
            TabIndex        =   94
            Top             =   1200
            Width           =   2175
         End
         Begin VB.Label PNLPropUnits 
            Caption         =   "Label1"
            Height          =   255
            Index           =   17
            Left            =   5880
            TabIndex        =   93
            Top             =   840
            Width           =   2175
         End
         Begin VB.Label PNLPropUnits 
            Caption         =   "Label1"
            Height          =   255
            Index           =   16
            Left            =   5880
            TabIndex        =   92
            Top             =   480
            Width           =   2175
         End
         Begin VB.Label PNLPropUnits 
            Caption         =   "Label1"
            Height          =   255
            Index           =   15
            Left            =   5880
            TabIndex        =   91
            Top             =   120
            Width           =   2175
         End
         Begin VB.Label PNLPropName 
            Alignment       =   1  'Right Justify
            Caption         =   "Vapor Thermal Conductivity as f(T)"
            Height          =   255
            Index           =   22
            Left            =   0
            TabIndex        =   58
            Top             =   2640
            Width           =   3735
         End
         Begin VB.Label PNLPropName 
            Alignment       =   1  'Right Justify
            Caption         =   "Diffusivity in Water"
            Height          =   255
            Index           =   15
            Left            =   0
            TabIndex        =   57
            Top             =   120
            Width           =   3735
         End
         Begin VB.Label PNLPropName 
            Alignment       =   1  'Right Justify
            Caption         =   "Surface Tension @ 298.15 K"
            Height          =   255
            Index           =   17
            Left            =   0
            TabIndex        =   56
            Top             =   840
            Width           =   3735
         End
         Begin VB.Label PNLPropName 
            Alignment       =   1  'Right Justify
            Caption         =   "Diffusivity in Air"
            Height          =   255
            Index           =   16
            Left            =   0
            TabIndex        =   55
            Top             =   480
            Width           =   3735
         End
         Begin VB.Label PNLPropName 
            Alignment       =   1  'Right Justify
            Caption         =   "Liquid Thermal Conductivity as f(T)"
            Height          =   255
            Index           =   21
            Left            =   0
            TabIndex        =   54
            Top             =   2280
            Width           =   3735
         End
         Begin VB.Label PNLPropName 
            Alignment       =   1  'Right Justify
            Caption         =   "Liquid Viscosity as f(T)"
            Height          =   255
            Index           =   20
            Left            =   0
            TabIndex        =   53
            Top             =   1920
            Width           =   3735
         End
         Begin VB.Label PNLPropName 
            Alignment       =   1  'Right Justify
            Caption         =   "Vapor Viscosity as f(T)"
            Height          =   255
            Index           =   19
            Left            =   0
            TabIndex        =   52
            Top             =   1560
            Width           =   3735
         End
         Begin VB.Label PNLPropName 
            Alignment       =   1  'Right Justify
            Caption         =   "Surface Tension as f(T)"
            Height          =   255
            Index           =   18
            Left            =   0
            TabIndex        =   51
            Top             =   1200
            Width           =   3735
         End
      End
   End
   Begin VB.Label LBLUserList 
      Caption         =   "User List"
      Height          =   255
      Left            =   5400
      TabIndex        =   20
      Top             =   120
      Width           =   4215
   End
   Begin VB.Label LBLChemList 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      Caption         =   "Master Chemical List"
      ForeColor       =   &H80000008&
      Height          =   240
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   4215
   End
   Begin VB.Menu MNUfile 
      Caption         =   "&File"
      Begin VB.Menu MNUNew 
         Caption         =   "&New"
      End
      Begin VB.Menu MNULoad 
         Caption         =   "&Load"
      End
      Begin VB.Menu MNUSave 
         Caption         =   "&Save"
      End
      Begin VB.Menu MNUSaveAs 
         Caption         =   "Save &As"
      End
      Begin VB.Menu MNUPrintChem 
         Caption         =   "&Print"
      End
      Begin VB.Menu MNUExport 
         Caption         =   "&Export"
      End
      Begin VB.Menu MNUExitPEARLS 
         Caption         =   "E&xit"
      End
   End
   Begin VB.Menu MNUEdit 
      Caption         =   "&Edit"
      Begin VB.Menu MNUFind 
         Caption         =   "&Find..."
      End
      Begin VB.Menu MNUSort 
         Caption         =   "&Sort"
         Begin VB.Menu MNUChemList 
            Caption         =   "&Chemical List"
            Begin VB.Menu MNUCLAsc 
               Caption         =   "&Ascending"
            End
            Begin VB.Menu MNUCLDsc 
               Caption         =   "&Descending"
            End
         End
         Begin VB.Menu MNUUserList 
            Caption         =   "&User List"
            Begin VB.Menu MNUULAsc 
               Caption         =   "&Ascending"
            End
            Begin VB.Menu MNUULDsc 
               Caption         =   "&Descending"
            End
         End
      End
   End
   Begin VB.Menu MNUGrph 
      Caption         =   "&Graph"
      Begin VB.Menu MNUNewGrph 
         Caption         =   "&New Graph"
      End
   End
   Begin VB.Menu MNUOptions 
      Caption         =   "&Options"
      Begin VB.Menu MNUViewPref 
         Caption         =   "&Environment Preferences"
      End
      Begin VB.Menu MNUfilepref 
         Caption         =   "&File Preferences"
      End
   End
   Begin VB.Menu MNUHelp 
      Caption         =   "&Help"
      Begin VB.Menu MNUContents 
         Caption         =   "&Contents"
      End
      Begin VB.Menu MNUAbout 
         Caption         =   "&About"
      End
   End
End
Attribute VB_Name = "frmmain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Calculate_button_Click()
Dim TempCAS1 As Long
    Dim TempCAS2 As Long
    Dim DBTbl As Recordset
    
    On Error GoTo error_handler
    
    frmmain!TABViewProp.CurrTab = 6
    frmmain.Refresh
        
    TempCAS1 = Cur_Info.CAS
    frmmain!Data2.Recordset.FindFirst "Name = '" & frmmain!LSTUserList.Text & "'"
    Cur_Info.CAS = frmmain!Data2.Recordset("CAS")
    TempCAS2 = Cur_Info.CAS
    
    'Check to see if same CAS was selected
'    If TempCAS1 = Cur_Info.CAS Then Exit Sub
        
    Screen.MousePointer = 11
    
    Cur_Info.CAS = TempCAS1
    Call SaveUserData
    Cur_Info.CAS = TempCAS2
    
    frmmain!Data1.Recordset.FindFirst "CAS = " & Cur_Info.CAS
    TXTFamily.Text = GetFamilyGroup(frmmain!Data1.Recordset("Chemical Family"))
    
    'Store last CAS number viewed
    Set DBTbl = DBJetUser.OpenRecordset("Last CAS", dbOpenTable)
    DBTbl.Edit
    DBTbl("LastCAS") = Cur_Info.CAS
    DBTbl("ULCAS") = Cur_Info.CAS
    DBTbl("LastList") = 2
    DBTbl.Update
    DBTbl.Close

    Call Recalculate
    Call DisplayProps
    
    'Turn on all folders
    Call TabFolderEnable(True)

    frmmain.Refresh
    Screen.MousePointer = 1
        
error_handler:
        
    If Err = 3021 Then
        MsgBox "Error - No Chemical selected", 48, "Error"
    End If
    Screen.MousePointer = 1
        
End Sub



Private Sub CMDAdd_Click()

    Dim DBTbl As Recordset
    
    If frmmain!Data2.Recordset.RecordCount > 19 Then
        MsgBox "Maximum selection limit is 20 chemicals", 48, "Chemical Limit Reached"
        Exit Sub
    End If
    
    frmmain!Data2.Recordset.FindFirst "Name = " & Chr(34) & frmmain!LSTSelList.Text & Chr(34)
    
    If Not frmmain!Data2.Recordset.NoMatch Then
        MsgBox "Chemical already in user list", 48, "Chemical Previously Selected"
        Exit Sub
    End If
    
    'Set modified flag
    WorkModified = True
    frmmain.caption = "PEARLS:  " & SaveFileName & " modified"
    
    Set DBTbl = DBJetUser.OpenRecordset("User List", dbOpenDynaset)
    
    DBTbl.AddNew
    frmmain!Data1.Recordset.FindFirst "Name = " & Chr(34) & frmmain!LSTSelList.Text & Chr(34)
    DBTbl("CAS") = frmmain!Data1.Recordset("CAS")
    DBTbl("Name") = frmmain!LSTSelList.Text
    DBTbl.Update
    
    DBTbl.Close
    
    frmmain!Data2.Refresh

End Sub

Private Sub CMDRemove_Click()

    Dim DBTbl As Recordset
    
    If frmmain!Data2.Recordset.RecordCount < 1 Then
        MsgBox "No chemicals have been selected", 48, "No Chemicals"
        Exit Sub
    End If
    
    'Set modified flag
    WorkModified = True
    frmmain.caption = "PEARLS:  " & SaveFileName & " modified"
    Set DBTbl = DBJetUser.OpenRecordset("User List", dbOpenDynaset)
    
    DBTbl.FindFirst "Name = " & Chr(34) & frmmain!LSTUserList.Text & Chr(34)
    DBTbl.Delete
    
    frmmain!Data2.Refresh

End Sub

Private Sub CMDRemoveAll_Click()

    Dim i As Integer
    Dim J As Integer
    Dim DBTbl As Recordset
    
    'Set mousepointer to hourglass (wait mode)
    Screen.MousePointer = 11
                 
    If frmmain!Data2.Recordset.RecordCount < 1 Then
        MsgBox "No chemicals have been selected", 48, "No Chemicals"
        Exit Sub
    End If
    
    'Set modified flag
    WorkModified = True
    frmmain.caption = "PEARLS:  " & SaveFileName & " modified"
    
    Set DBTbl = DBJetUser.OpenRecordset("User List", dbOpenDynaset)
    
    DBTbl.MoveFirst
    Do While DBTbl.EOF = False
        DBTbl.Delete
        DBTbl.MoveNext
    Loop
    
    frmmain!Data2.Refresh
                                                   
    'Reset display and TFT
    For i = 0 To NumProperties
        For J = 1 To NumMethods
            InfoMethod(i).Enabled(J) = False
            InfoMethod(i).TFT = 298.15
        Next J
    Next i
    
    Call DisplayProps
    
    'Set to Chemical Information tab
    frmmain!TABViewProp.CurrTab = 6
    
    'Turn off all folders
    Call TabFolderEnable(False)
    
    'Set mousepointer to arrow (normal mode)
    Screen.MousePointer = 1
    
    'Reset CAS number to 0 (No chemical loaded in memory)
    Cur_Info.CAS = 0
    
    Exit Sub
           
End Sub

Private Sub Command1_Click()
Dim TempCAS1 As Long
    Dim TempCAS2 As Long
    Dim DBTbl As Recordset
    
    frmmain!TABViewProp.CurrTab = 6
    frmmain.Refresh
        
    TempCAS1 = Cur_Info.CAS
    frmmain!Data2.Recordset.FindFirst "Name = '" & frmmain!LSTUserList.Text & "'"
    Cur_Info.CAS = frmmain!Data2.Recordset("CAS")
    TempCAS2 = Cur_Info.CAS
    
    'Check to see if same CAS was selected
    If TempCAS1 = Cur_Info.CAS Then Exit Sub
        
    Screen.MousePointer = 11
    
    Cur_Info.CAS = TempCAS1
    Call SaveUserData
    Cur_Info.CAS = TempCAS2
    
    frmmain!Data1.Recordset.FindFirst "CAS = " & Cur_Info.CAS
    TXTFamily.Text = GetFamilyGroup(frmmain!Data1.Recordset("Chemical Family"))
    
    'Store last CAS number viewed
    Set DBTbl = DBJetUser.OpenRecordset("Last CAS", dbOpenTable)
    DBTbl.Edit
    DBTbl("LastCAS") = Cur_Info.CAS
    DBTbl("ULCAS") = Cur_Info.CAS
    DBTbl("LastList") = 2
    DBTbl.Update
    DBTbl.Close

    Call Recalculate
    Call DisplayProps
    
    'Turn on all folders
    Call TabFolderEnable(True)
    frmmain.Refresh
    Screen.MousePointer = 1

End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)

    If KeyAscii = &H70 Or KeyAscii = 43 Or KeyAscii = 46 Or KeyAscii = 13 Or KeyAscii = 8 Or KeyAscii = 45 Or KeyAscii = 69 Or KeyAscii = 101 Or KeyAscii = &H25 Or KeyAscii = &H27 Then Exit Sub
    If KeyAscii < 48 Or KeyAscii > 57 Then KeyAscii = 0

End Sub

Private Sub Form_Load()
    
    CenterForm Me
        
End Sub

Sub Die()

    Unload frm911DBInfo
    Unload frmantoine
    Unload frmexport
    Unload frmfind
    Unload frmgraph
    Unload frmgraphSet
    Unload frminfo
'    Unload FRM
    Unload frmmaster
    Unload frmmethod
    Unload frmpreferences
    Unload frmprint
    Unload frmtitle
    Unload frmunits
    Unload frmuser

End Sub

Private Sub Form_Resize()
    Call ReDrawForm
End Sub

Sub ReDrawForm()
' paul 4-23-99
'redraws the main form after a resize event allows the chemical
'list and user list to become readable on larger screens,
'finally utilizing the whole screen.

    If frmmain.Width >= 9930 And frmmain.Height >= 7200 Then
        LBLChemList.Left = 120
        LBLChemList.Top = 120
        
        LSTSelList.Left = 120
        LSTSelList.Top = 360
        LSTSelList.Width = frmmain.Width / 2 - 750
        LSTSelList.Height = frmmain.Height - 6165
        
        LBLUserList.Left = frmmain.Width / 2 - CMDAdd.Width / 2 + 800
        LBLUserList.Top = 120
        
        LSTUserList.Left = frmmain.Width / 2 - CMDAdd.Width / 2 + 800
        LSTUserList.Top = 360
        LSTUserList.Width = frmmain.Width / 2 - 750
        LSTUserList.Height = frmmain.Height - 6165
        
        TABViewProp.Left = frmmain.Width / 2 - TABViewProp.Width / 2 - 100
        TABViewProp.Top = LSTUserList.Height + 765
        
        CMDAdd.Left = frmmain.Width / 2 - CMDAdd.Width / 2 - 100
        CMDAdd.Top = LSTUserList.Height / 2 - 97
        CMDRemove.Left = frmmain.Width / 2 - CMDRemove.Width / 2 - 100
        CMDRemove.Top = LSTUserList.Height / 2 + 262
        CMDRemoveAll.Left = frmmain.Width / 2 - CMDRemoveAll.Width / 2 - 100
        CMDRemoveAll.Top = LSTUserList.Height / 2 + 662
        Calculate_button.Left = frmmain.Width / 2 - Calculate_button.Width / 2 + 2120
        Calculate_button.Top = LSTUserList.Height + 525
    Else
        frmmain.Width = 9930
        frmmain.Height = 7200
        Call ReDrawForm
    End If

    'CMDAdd.Left = 4500
    'CMDAdd.Top = 420
    'CMDRemove.Left = 4500
    'CMDRemove.Top = 780
    'CMDRemoveAll.Left = 4500
    'CMDRemoveAll.Top = 1140
    'Calculate_button.Left = 6720
    'Calculate_button.Top = 1560
    '
    'LBLChemList.Left = 120
    'LBLChemList.Top = 120
    '
    'LSTSelList.Left = 120
    'LSTSelList.Top = 360
    'LSTSelList.Width = 4215
    'LSTSelList.Height = 1035
    '
    'LBLUserList.Left = 5400
    'LBLUserList.Top = 120
    '
    'LSTUserList.Left = 5400
    'LSTUserList.Top = 360
    'LSTUserList.Width = 4215
    'LSTUserList.Height = 1035
    '
    'TABViewProp.Left = 120
    'TABViewProp.Top = 1800
End Sub

Private Sub form_Unload(Cancel As Integer)
    
Call Die
    
    ' REVISIONS 3/13/98  DMW  added writing def file here, should
    '                           be the only place it needs to be done
    Dim answer As Integer
    ' double check user wants to exit
    answer = MsgBox("Exit Pearls?", vbYesNo)
    If answer = vbNo Then
        Cancel = True
        Exit Sub
    End If
    
    ' if so, and user database has been changed, prompt to save it
    If WorkModified = True Then
        answer = MsgBox("Save Current Template?", vbYesNo)
        If answer = vbYes Then
            If Len(SaveFileName) > 3 Then
                Call MNUSave_Click
            Else
                Call MNUSaveAs_Click
            End If
        End If
    End If
    DBJetMaster.Close
    On Error GoTo after_close_user
    DBJetUser.Close
after_close_user:
    Call write_def_file
    Exit Sub
End Sub


Private Sub Label1_Click(Index As Integer)

End Sub

Private Sub LBLChemList_Click()

    Dim TempCAS As Long
    
    TempCAS = frmmain!Data1.Recordset("CAS")
    
    If SortChemListAsc = False Then
        SortChemListAsc = True
        frmmain!Data1.RecordSource = "SELECT * FROM [PEARLS List] ORDER BY [Name] ASC"
        frmmain!Data1.Refresh
    Else
        SortChemListAsc = False
        frmmain!Data1.RecordSource = "SELECT * FROM [PEARLS List] ORDER BY [Name] DESC"
        frmmain!Data1.Refresh
    End If
    
    frmmain!Data1.Recordset.FindFirst "CAS = " & TempCAS
    frmmain!LSTSelList.Text = frmmain!Data1.Recordset("Name")
    TXTFamily.Text = GetFamilyGroup(frmmain!Data1.Recordset("Chemical Family"))
    
End Sub

Private Sub LBLOpPUnits_Click()
    
    CurProp = OptPress
    TFTConvert = False
    
    Call LoadUnitsForm
        
    frmunits.Show

End Sub

Private Sub LBLOpTUnits_Click()
    
    CurProp = OptTemp
    TFTConvert = False
      
    Call LoadUnitsForm
    
    frmunits.Show

End Sub

Private Sub LBLUserList_Click()
    
    Dim TempCAS As Long
    On Error Resume Next
    TempCAS = frmmain!Data2.Recordset("CAS")

    If SortUserListAsc = False Then
        SortUserListAsc = True
        frmmain!Data2.RecordSource = "SELECT * FROM [User List] ORDER BY [Name] ASC"
        frmmain!Data2.Refresh
    Else
        SortUserListAsc = False
        frmmain!Data2.RecordSource = "SELECT * FROM [User List] ORDER BY [Name] DESC"
        frmmain!Data2.Refresh
    End If
    
    frmmain!Data1.Recordset.FindFirst "CAS = " & TempCAS
    frmmain!LSTUserList.Text = frmmain!Data1.Recordset("Name")
    TXTFamily.Text = GetFamilyGroup(frmmain!Data1.Recordset("Chemical Family"))
    
End Sub

Private Sub LSTSelList_Click()

    Dim TempCAS As Long
    Dim DBTbl As Recordset
    
    frmmain!TABViewProp.CurrTab = 6
        
    frmmain!Data1.Recordset.FindFirst "Name = " & Chr(34) & CStr(frmmain!LSTSelList.Text) & Chr(34)
    TXTFamily.Text = GetFamilyGroup(frmmain!Data1.Recordset("Chemical Family"))
    TempCAS = frmmain!Data1.Recordset("CAS")

    'Store last CAS number viewed
    Set DBTbl = DBJetUser.OpenRecordset("Last CAS", dbOpenTable)
    DBTbl.Edit
    DBTbl("CLCAS") = TempCAS
    DBTbl("LastList") = 1
    DBTbl.Update
    DBTbl.Close

    If TempCAS <> Cur_Info.CAS Then
        Call TabFolderEnable(False)
    Else
        Call TabFolderEnable(True)
    End If
    frmmain.Refresh
End Sub

Private Sub LSTSelList_DblClick()

    ' we want the double click of this list to update the front tab with info
    Dim DBTbl As Recordset
    
    If frmmain!Data2.Recordset.RecordCount > 19 Then
        MsgBox "Maximum selection limit is 20 chemicals", 48, "Chemical Limit Reached"
        Exit Sub
    End If
    
    frmmain!Data2.Recordset.FindFirst "Name = '" & frmmain!LSTSelList.Text & "'"
    
    If Not frmmain!Data2.Recordset.NoMatch Then
        MsgBox "Chemical already in user list", 48, "Chemical Previously Selected"
        Exit Sub
    End If
    
    Set DBTbl = DBJetUser.OpenRecordset("User List", dbOpenDynaset)
    
    DBTbl.AddNew
    frmmain!Data1.Recordset.FindFirst "Name = '" & frmmain!LSTSelList.Text & "'"
    DBTbl("CAS") = frmmain!Data1.Recordset("CAS")
    DBTbl("Name") = frmmain!LSTSelList.Text
    DBTbl.Update
    
    DBTbl.Close
        
    frmmain!Data2.Refresh

End Sub


Private Sub LSTUserList_Click()

    Dim TempCAS As Long
    Dim DBTbl As Recordset
    
    frmmain!TABViewProp.CurrTab = 6
        
    frmmain!Data2.Recordset.FindFirst "Name = " & Chr(34) & frmmain!LSTUserList.Text & Chr(34)
    TempCAS = frmmain!Data2.Recordset("CAS")
        
    frmmain!Data1.Recordset.FindFirst "CAS = " & TempCAS
    TXTFamily.Text = GetFamilyGroup(frmmain!Data1.Recordset("Chemical Family"))

    'Store last CAS number viewed
    Set DBTbl = DBJetUser.OpenRecordset("Last CAS", dbOpenTable)
    DBTbl.Edit
    DBTbl("ULCAS") = TempCAS
    DBTbl("LastList") = 2
    DBTbl.Update
    DBTbl.Close

    'Turn on all folders if CAS in user list matches that in other folders
    If TempCAS <> Cur_Info.CAS Then
        Call TabFolderEnable(False)
    Else
        Call TabFolderEnable(True)
    End If
        
End Sub

Private Sub LSTUserList_DblClick()
    
    ' we want the double click to calculate the chemical clicked on
    Dim TempCAS1 As Long
    Dim TempCAS2 As Long
    Dim DBTbl As Recordset
    
    frmmain!TABViewProp.CurrTab = 6
    frmmain.Refresh
        
    TempCAS1 = Cur_Info.CAS
    frmmain!Data2.Recordset.FindFirst "Name = " & Chr(34) & frmmain!LSTUserList.Text & Chr(34)
    Cur_Info.CAS = frmmain!Data2.Recordset("CAS")
    TempCAS2 = Cur_Info.CAS
    
    'Check to see if same CAS was selected
    If TempCAS1 = Cur_Info.CAS Then Exit Sub
        
    Screen.MousePointer = 11
    
    Cur_Info.CAS = TempCAS1
    Call SaveUserData
    Cur_Info.CAS = TempCAS2
    
    frmmain!Data1.Recordset.FindFirst "CAS = " & Cur_Info.CAS
    TXTFamily.Text = GetFamilyGroup(frmmain!Data1.Recordset("Chemical Family"))
    
    'Store last CAS number viewed
    Set DBTbl = DBJetUser.OpenRecordset("Last CAS", dbOpenTable)
    DBTbl.Edit
    DBTbl("LastCAS") = Cur_Info.CAS
    DBTbl("ULCAS") = Cur_Info.CAS
    DBTbl("LastList") = 2
    DBTbl.Update
    DBTbl.Close

    Call Recalculate
    Call DisplayProps
    
    'Turn on all folders
    Call TabFolderEnable(True)
    frmmain.Refresh
    Screen.MousePointer = 1

End Sub

Private Sub MNUAbout_Click()
    frmtitle.ProgressBar1.Visible = False
    frmtitle.Show 1

End Sub

Private Sub MNUCLAsc_Click()
    
    Dim TempCAS As Long
    
    TempCAS = frmmain!Data1.Recordset("CAS")

    frmmain!Data1.RecordSource = "SELECT * FROM [PEARLS List] ORDER BY [Name] ASC"
    frmmain!Data1.Refresh

    frmmain!Data1.Recordset.FindFirst "CAS = " & TempCAS
    frmmain!LSTSelList.Text = frmmain!Data1.Recordset("Name")
    TXTFamily.Text = GetFamilyGroup(frmmain!Data1.Recordset("Chemical Family"))

End Sub

Private Sub MNUCLDsc_Click()
    
    Dim TempCAS As Long
    On Error Resume Next
    TempCAS = frmmain!Data2.Recordset("CAS")

    frmmain!Data1.RecordSource = "SELECT * FROM [PEARLS List] ORDER BY [Name] DESC"
    frmmain!Data1.Refresh

    frmmain!Data1.Recordset.FindFirst "CAS = " & TempCAS
    frmmain!LSTSelList.Text = frmmain!Data1.Recordset("Name")
    TXTFamily.Text = GetFamilyGroup(frmmain!Data1.Recordset("Chemical Family"))

End Sub

Private Sub MNUContents_Click()

    App.HelpFile = App.path + "\pearls.hlp"
    CODFilePath.HelpFile = App.HelpFile
    CODFilePath.HelpCommand = &H3
    CODFilePath.ShowHelp
        
End Sub

Private Sub MNUExitPEARLS_Click()
Unload Me
End Sub


' MNUExport_Click():    This function checks that at least
'                       one chemical has been selected to
'                       export and, if so, shows the export
'                       form. Otherwise, a message box is
'                       displayed and control returned to mainfrm
'
Private Sub MNUExport_Click()
    
    If frmmain!Data2.Recordset.RecordCount < 1 Then
        MsgBox "No chemicals have been selected", 48, "No Chemicals"
        Exit Sub
    End If
    Call load_form_export
    frmexport.Show 1

End Sub

Private Sub MNUfilepref_Click()

    Dim answer As Integer
    Dim old_master As String
    Dim old_save As String
    Dim old_block5 As String
    Dim changes As Boolean
    changes = False
   
    old_master = PathMaster
    old_save = PathSave
    old_block5 = PathBlock5
    Call load_frm_master_info
    ' don't want the user to be able to exit this way
    frmmaster!cmdexit.Enabled = False
    ' ****** disable editing the master path
    frmmaster!txtpath(0).Enabled = False
    frmmaster!TXTName(0).Enabled = False
    Call CenterForm(frmmaster)
    frmmaster.Show 1
    ' ****** for now keep old master - can't change here
    PathMaster = old_master
    frmmaster!txtpath(0).Enabled = True
    frmmaster!TXTName(0).Enabled = True
    ' check for changes to file settings
    If Trim(old_master) <> Trim(PathMaster) Or Trim(old_save) <> Trim(PathSave) Or Trim(old_block5) <> Trim(PathBlock5) Then
        changes = True
    End If
    
    ' if changes have been made, need to reload master and/or user database
    If changes = True Then
        ' make sure the user really wants to do this, because it's going to
        ' reload PEARLS stuff
        answer = MsgBox("File Preferences have been made, save changes and update PEARLS?", vbYesNo)
        If answer = vbNo Then
            PathMaster = old_master
            PathBlock5 = old_block5
            PathSave = old_save
            Exit Sub
        End If
        ' reload new saved file if thats been changed
        If Trim(old_save) <> Trim(PathSave) Then
            If load_selected_save = False Then
                PathSave = old_save
            End If
        End If
        'Set up database security
        ' ******* commented this out for now, reloading master causes "out of stack space" error - needs to be fixed
        'Call SetUpSecurity
        ' reload new master if that's been changed
        'If Trim(old_master) <> Trim(PathMaster) Then
        '    DBJetMaster.Close
            'Initialize global variables
        '    If load_selected_master = False Then
         '       PathMaster = old_master
         '   End If
            'Set mousepointer to arrow
         '   Screen.MousePointer = 1
        
        'End If
    ' no need to do anything to block 5, that gets opened on calculation
        
    End If
    frmmain.Refresh
    Screen.MousePointer = 1
End Sub

Private Sub MNUFind_Click()
     
    frmfind.Show
    frmfind!CMBFindStr.SetFocus
    
End Sub

Public Sub MNULoad_Click()
       
    Dim DBTbl As Recordset
    Dim Response As Integer
    Dim CLCAS As Long
    Dim ULCAS As Long
    Dim LastList As Integer
    
    'Check for existing selections
    'On Error GoTo skip_save_warning
    If frmmain!Data2.Recordset.RecordCount >= 1 And WorkModified = True Then
        Response = MsgBox("Save current template?", 3, "Save")
        If Response = 6 Then
            Call MNUSave_Click
        ElseIf Response = 2 Then
            Exit Sub
        End If
    End If
   'DBJetUser.Close

    On Error GoTo cancel_error
    'Show user available files
    CODFilePath.Filter = "PEARLS (*.prl)|*.prl"
    CODFilePath.FilterIndex = 1
    CODFilePath.InitDir = App.path
    CODFilePath.DefaultExt = "prl"
    CODFilePath.Action = 1
            
skip_save_warning:
    On Error GoTo LoadError
    'Set mousepointer to hourglass (wait mode)
    Screen.MousePointer = 11
    
    'Refresh main form
    frmmain.Refresh
    On Error Resume Next
    DBJetUser.Close
    frmmain!Data2.databasename = App.path & "\temp.mdb"
    frmmain!Data2.RecordSource = "User List"
    frmmain!Data2.Refresh
    Kill PathUser
    ' set the savefile name and copy it to dbuser so we're not directly editing the saved file
    PathSave = CODFilePath.filename
    SaveFileName = PathSave
    FileCopy PathSave, PathUser
    Set DBJetUser = OpenDatabase(PathUser, False, False)
  
    UserDBName = PathUser
    frmmain!Data2.databasename = PathUser
    frmmain!Data2.RecordSource = "User List"
    frmmain!Data2.RecordsetType = 2
    frmmain!Data2.Refresh
    
    Set DBTbl = DBJetUser.OpenRecordset("Last CAS", dbOpenTable)
    On Error Resume Next
    Cur_Info.CAS = DBTbl("LastCAS")
    CLCAS = DBTbl("CLCAS")
    ULCAS = DBTbl("ULCAS")
    LastList = DBTbl("LastList")
    
    DBTbl.Close
    
    frmmain!Data1.Recordset.FindFirst "CAS =" & Cur_Info.CAS
    ' try moving this here
    Call LoadUserPreferences
    If GetUserData = True Then
        If LastList = 1 And CLCAS = Cur_Info.CAS Then
            'Call LoadUserPreferences
            Call Recalculate
            Call DisplayProps
            Call TabFolderEnable(True)
        End If
        If LastList = 2 And ULCAS = Cur_Info.CAS Then
            'Call LoadUserPreferences
            Call Recalculate
            Call DisplayProps
            Call TabFolderEnable(True)
        End If
    End If
        
    frmmain!Data1.Recordset.FindFirst "CAS =" & CLCAS
    frmmain!LSTSelList.Text = frmmain!Data1.Recordset("Name")
    
    frmmain!Data2.Recordset.FindFirst "CAS =" & ULCAS
    frmmain!LSTUserList.Text = frmmain!Data2.Recordset("Name")
        
    If LastList = 1 Then
        frmmain!LSTSelList.SetFocus
    Else
        frmmain!LSTUserList.SetFocus
    End If
        
    'Set mousepointer to arrow (normal mode)
    Screen.MousePointer = 1
    'Reset modified flag
    WorkModified = False
    ' make sure the file name fits on the main form
    If Len(SaveFileName) > 45 Then
        frmmain.caption = "PEARLS:  ..." & Right(SaveFileName, 45)
    Else
        frmmain.caption = "PEARLS:  " & SaveFileName
    End If
    Exit Sub
       
LoadError:
    frmmain!Data2.databasename = App.path & "\temp.mdb"
    frmmain!Data2.RecordSource = "User List"
    frmmain!Data2.Refresh
    Screen.MousePointer = 1
    If Err <> 32755 Then
        MsgBox "Error loading PEARLS file", 48, "Error"
    End If
    Exit Sub
cancel_error:
    
End Sub



Private Sub MNUNew_Click()

    Dim i As Integer
    Dim J As Integer
    Dim Response As Integer
             
    'Check for existing selections
    If frmmain!Data2.Recordset.RecordCount >= 1 And WorkModified = True Then
        Response = MsgBox("Save current template?", 3, "Save")
        If Response = 6 Then
            Call MNUSave_Click
        ElseIf Response = 2 Then
            Exit Sub
        End If
    End If
    Call CMDRemoveAll_Click
    
    'Set mousepointer to hourglass (wait mode)
    Screen.MousePointer = 11
    
    'Refresh main form
    WorkModified = False
    frmmain.caption = "PEARLS:  unmodified"
    frmmain.Refresh
    
    On Error Resume Next
    
    DBJetUser.Close
    frmmain!Data2.databasename = App.path & "\temp.mdb"
    frmmain!Data2.RecordSource = "User List"
    frmmain!Data2.Refresh
    Kill PathUser
    
    On Error GoTo NewError
    
    FileCopy PathSave, PathUser
    Set DBJetUser = OpenDatabase(PathUser, False, False)
               
    UserDBName = PathUser
    frmmain!Data2.databasename = PathUser
    frmmain!Data2.RecordSource = "User List"
    frmmain!Data2.RecordsetType = 2
    frmmain!Data2.Refresh
               
    Call LoadUserPreferences
                   
    'Reset display and TFT
    For i = 0 To NumProperties
        For J = 1 To NumMethods
            InfoMethod(i).Enabled(J) = False
            InfoMethod(i).TFT = 298.15
        Next J
    Next i
    Call update_DisplayData
    Call DisplayProps
    
    'Set to Chemical Information tab
    frmmain!TABViewProp.CurrTab = 6
    
    'Turn off all folders
    Call TabFolderEnable(False)
    
    'Set mousepointer to arrow (normal mode)
    Screen.MousePointer = 1
    
    'Reset CAS number to 0 (No chemical loaded in memory)
    Cur_Info.CAS = 0
    
    'Reset SaveFileName flag
    SaveFileName = ""
    
    Exit Sub
       
NewError:
    Screen.MousePointer = 1
    If Err <> 32755 Then
        MsgBox "Restore Error", 48, "Error"
    End If
    
End Sub

Private Sub MNUNewGrph_Click()
    
    If frmmain!Data2.Recordset.RecordCount >= 1 Then
        frmgraphSet.Show 1
    Else
        MsgBox "No chemicals have been selected", 48, "No Chemicals Selected"
    End If

End Sub

Private Sub MNUPrintChem_Click()

    If frmmain!Data2.Recordset.RecordCount < 1 Then
        MsgBox "No chemicals have been selected", 48, "No Chemicals"
        Exit Sub
    End If
    If PathReport = NULLPATH Then
        MsgBox ("Printing files don't exist or paths not setd")
        Exit Sub
    End If
        
    frmprint.Show 1
    
End Sub


Public Sub MNUSave_Click()

    ' this function saves what's in the dbuser to
    ' a separate file.  If the filename hasn't been set,
    ' it calls <saveas> so that the user can set
    ' file name
    
    ' savefilename should have the path and the filename
    If SaveFileName = "" Or Right(Trim(SaveFileName), 8) = "demo.prl" Or Right(Trim(SaveFileName), 8) = "DEMO.PRL" Then
        Call MNUSaveAs_Click
        Exit Sub
    End If
            
    On Error GoTo SaveError
                
    'Set mousepointer to hourglass (wait mode)
    Screen.MousePointer = 11
    
    'Refresh main form
    frmmain.Refresh
    
    'Save current chemical
    Call SaveUserData

    DBJetUser.Close
        ' need to set this temporarily because we're going to change it
    frmmain!Data2.databasename = App.path & "\temp.mdb"
    frmmain!Data2.RecordSource = "User List"
    frmmain!Data2.Refresh
    FileCopy PathUser, SaveFileName
    
    Set DBJetUser = OpenDatabase(PathUser, False, False)
       
    frmmain!Data2.databasename = PathUser
    frmmain!Data2.RecordSource = "User List"
    frmmain!Data2.RecordsetType = 2
    frmmain!Data2.Refresh

    'Set mousepointer to arrow (normal mode)
    Screen.MousePointer = 1
   
    WorkModified = False
    frmmain.caption = "PEARLS:  " & SaveFileName
    Exit Sub
    
SaveError:
    Screen.MousePointer = 1
    If Err <> 32755 Then
        MsgBox "Error saving PEARLS file", 16, "Error"
    End If
    
End Sub



Private Sub MNUSaveAs_Click()
     ' DENISE: may want to change this function to
     '          modify pathuser when the user saves
     '          a file??
     
    Dim answer As Integer
    If frmmain!Data2.Recordset.RecordCount < 1 Then
        MsgBox "No chemicals have been selected", 48, "No Chemicals"
        Exit Sub
    End If
      
    On Error GoTo SaveAsError
    
do_file_browser:
    'Show save dialog box
    CODFilePath.Filter = "PEARLS (*.prl)|*.prl"
    CODFilePath.FilterIndex = 1
    CODFilePath.InitDir = App.path
    CODFilePath.DefaultExt = "prl"
    CODFilePath.Action = 2
            
    'Set mousepointer to hourglass (wait mode)
    
    Screen.MousePointer = 11
    
    'Refresh main form
    frmmain.Refresh
    SaveFileName = Trim(CODFilePath.filename)
    'Save current chemical
    If Right(Trim(SaveFileName), 8) = "demo.prl" Or Right(Trim(SaveFileName), 8) = "DEMO.PRL" Then
        answer = MsgBox("Modify PEARLS demo file?", vbYesNo)
        If answer = vbNo Then
            Screen.MousePointer = 1
            GoTo do_file_browser
            Exit Sub
        End If
    End If
    

    DBJetUser.Close
    frmmain!Data2.databasename = App.path & "\temp.mdb"
    frmmain!Data2.RecordSource = "User List"
    frmmain!Data2.Refresh
    SaveFileName = Trim(CODFilePath.filename)
    FileCopy PathUser, CODFilePath.filename
    
    Set DBJetUser = OpenDatabase(PathUser, False, False)
       
    UserDBName = PathUser
    frmmain!Data2.databasename = PathUser
    frmmain!Data2.RecordSource = "User List"
    frmmain!Data2.RecordsetType = 2
    frmmain!Data2.Refresh

    'Set mousepointer to arrow (normal mode)
    Screen.MousePointer = 1
   
    WorkModified = False
    frmmain.caption = "PEARLS:  " & SaveFileName
    Exit Sub
    
SaveAsError:
    Screen.MousePointer = 1
    If Err <> 32755 Then
        MsgBox "Error saving PEARLS file", 16, "Error"
    End If

End Sub


Private Sub MNUULAsc_Click()
    
    Dim TempCAS As Long
    On Error Resume Next
    TempCAS = frmmain!Data2.Recordset("CAS")

    frmmain!Data2.RecordSource = "SELECT * FROM [User List] ORDER BY [Name] ASC"
    frmmain!Data2.Refresh

    frmmain!Data1.Recordset.FindFirst "CAS = " & TempCAS
    frmmain!LSTUserList.Text = frmmain!Data1.Recordset("Name")
    TXTFamily.Text = GetFamilyGroup(frmmain!Data1.Recordset("Chemical Family"))

End Sub

Private Sub MNUULDsc_Click()

    Dim TempCAS As Long
    On Error Resume Next
    TempCAS = frmmain!Data2.Recordset("CAS")

    frmmain!Data2.RecordSource = "SELECT * FROM [User List] ORDER BY [Name] DESC"
    frmmain!Data2.Refresh

    frmmain!Data1.Recordset.FindFirst "CAS = " & TempCAS
    frmmain!LSTUserList.Text = frmmain!Data1.Recordset("Name")
    TXTFamily.Text = GetFamilyGroup(frmmain!Data1.Recordset("Chemical Family"))

End Sub


Private Sub MNUViewPref_Click()
    
    Call LoadPreferences
    frmpreferences.Show 1
    
End Sub












Private Sub PNLPropName_Click(Index As Integer)

    If Index < 40 Then
        CurProp = Index
        Call CreateMethodForm
    End If
    
    If (Index >= 49 And Index <= 62) Then
        CurProp = Index - 8
        Call CreateMethodForm
    End If
    
End Sub



Private Sub PNLPropUnits_Click(Index As Integer)

    CurProp = Index
    TFTConvert = False
        
    Call LoadUnitsForm
        
    frmunits.Show 1
    
End Sub


Private Sub PNLPropVal_Click(Index As Integer)
      
    CurProp = Index
    
    Call CreateMethodForm
  
End Sub


Private Sub TABViewProp_Switch(OldTab As Integer, NewTab As Integer, Cancel As Integer)
 'slight glitch in control. Lets you select disabled tab 0 for some
 'reason.  this checks to see if its disabled and if it is then go
 'back to old tab
 If Not (TABViewProp.TabEnabled(0)) Then
    NewTab = OldTab
End If

End Sub

Private Sub TXTOpP_KeyPress(KeyAscii As Integer)

    Dim Response As Integer
    Dim TempCAS1 As Long
    Dim TempCAS2 As Long
    Dim DBTbl As Recordset
    
    On Error GoTo NotValid
    
    If frmmain!Data2.Recordset.RecordCount < 1 Then
        MsgBox "No chemicals have been selected", 48, "No Chemicals"
        KeyAscii = 0
        Exit Sub
    End If
    
    If KeyAscii <> 13 Then Exit Sub
    
    If FormatVal(Val(TXTOpP.Text)) = FormatVal(Cur_Disp.OpP) Then
        TXTOpP.Text = FormatVal(Cur_Disp.OpP)
        Exit Sub
    End If
            
    'Ask user to recalculate at new pressure
    Response = MsgBox("Recalculate using new pressure?", 3, "Recalculate")
    If Response = 6 Then
        Cur_Info.OpP = Convert(Val(TXTOpP.Text), OptPress, Cur_Info.OpPUnit, "Pa", False)
 
        frmmain!TABViewProp.CurrTab = 6
        frmmain.Refresh
        
        TempCAS1 = Cur_Info.CAS
        frmmain!Data2.Recordset.FindFirst "Name = '" & frmmain!LSTUserList.Text & "'"
        Cur_Info.CAS = frmmain!Data2.Recordset("CAS")
        TempCAS2 = Cur_Info.CAS
            
        Screen.MousePointer = 11
    
        Cur_Info.CAS = TempCAS1
        Call SaveUserData
        Cur_Info.CAS = TempCAS2
    
        frmmain!Data1.Recordset.FindFirst "CAS = " & Cur_Info.CAS
        TXTFamily.Text = GetFamilyGroup(frmmain!Data1.Recordset("Chemical Family"))
    
        'Store last CAS number viewed
        Set DBTbl = DBJetUser.OpenRecordset("Last CAS", dbOpenTable)
        DBTbl.Edit
        DBTbl("LastCAS") = Cur_Info.CAS
        DBTbl("ULCAS") = Cur_Info.CAS
        DBTbl("LastList") = 2
        DBTbl.Update
        DBTbl.Close

        Call Recalculate
        Call DisplayProps
        
        'Turn on all folders
        Call TabFolderEnable(True)
        frmmain.Refresh
        Screen.MousePointer = 1
    Else
        TXTOpP.Text = FormatVal(Cur_Disp.OpP)
    End If
    
    Exit Sub
    
NotValid:
    MsgBox "Not a valid number, please enter a smaller value", 0, "Value Not Valid"
    TXTOpP.Text = FormatVal(Cur_Disp.OpP)

End Sub

Private Sub TXTOpP_LostFocus()

    Dim Response As Integer
    Dim TempCAS1 As Long
    Dim TempCAS2 As Long
    Dim DBTbl As Recordset
    
    On Error GoTo NotValid
    
    If FormatVal(Val(TXTOpP.Text)) = FormatVal(Cur_Disp.OpP) Then
        TXTOpP.Text = FormatVal(Cur_Disp.OpP)
        Exit Sub
    End If
            
    'Ask user to recalculate at new pressure
    Response = MsgBox("Recalculate using new pressure?", 3, "Recalculate")
    If Response = 6 Then
        On Error GoTo no_calc
        Cur_Info.OpP = Convert(Val(TXTOpP.Text), OptPress, Cur_Info.OpPUnit, "Pa", False)
 
        frmmain!TABViewProp.CurrTab = 6
        frmmain.Refresh
        
        TempCAS1 = Cur_Info.CAS
        frmmain!Data2.Recordset.FindFirst "Name = '" & frmmain!LSTUserList.Text & "'"
        Cur_Info.CAS = frmmain!Data2.Recordset("CAS")
        TempCAS2 = Cur_Info.CAS
            
        Screen.MousePointer = 11
    
        Cur_Info.CAS = TempCAS1
        Call SaveUserData
        Cur_Info.CAS = TempCAS2
    
        frmmain!Data1.Recordset.FindFirst "CAS = " & Cur_Info.CAS
        TXTFamily.Text = GetFamilyGroup(frmmain!Data1.Recordset("Chemical Family"))
    
        'Store last CAS number viewed
        Set DBTbl = DBJetUser.OpenRecordset("Last CAS", dbOpenTable)
        DBTbl.Edit
        DBTbl("LastCAS") = Cur_Info.CAS
        DBTbl("ULCAS") = Cur_Info.CAS
        DBTbl("LastList") = 2
        DBTbl.Update
        DBTbl.Close

        Call Recalculate
        Call DisplayProps
        
        'Turn on all folders
        Call TabFolderEnable(True)
        frmmain.Refresh
        Screen.MousePointer = 1
no_calc:
    Else
        TXTOpP.Text = FormatVal(Cur_Disp.OpP)
    End If
    
    Exit Sub
    
NotValid:
    MsgBox "Not a valid number, please enter a smaller value", 0, "Value Not Valid"
    TXTOpP.Text = FormatVal(Cur_Disp.OpP)

End Sub


Private Sub TXTOpT_KeyPress(KeyAscii As Integer)

    Dim Response As Integer
    Dim TempCAS1 As Long
    Dim TempCAS2 As Long
    Dim DBTbl As Recordset
                    
    On Error GoTo NotValid
    
    If frmmain!Data2.Recordset.RecordCount < 1 Then
        MsgBox "No chemicals have been selected", 48, "No Chemicals"
        KeyAscii = 0
        Exit Sub
    End If
    
    If KeyAscii <> 13 Then Exit Sub
    
    If FormatVal(Val(TXTOpT.Text)) = FormatVal(Cur_Disp.OpT) Then
        TXTOpT.Text = FormatVal(Cur_Disp.OpT)
        Exit Sub
    End If
    
    'Ask user to recalculate at new temperature
    Response = MsgBox("Recalculate using new temperature?", 3, "Recalculate")
    If Response = 6 Then
        On Error GoTo no_calc
        Cur_Info.OpT = Convert(Val(TXTOpT.Text), OptTemp, Cur_Disp.OpTUnit, "K", False)
        
        frmmain!TABViewProp.CurrTab = 6
        frmmain.Refresh
        
        TempCAS1 = Cur_Info.CAS
        frmmain!Data2.Recordset.FindFirst "Name = '" & frmmain!LSTUserList.Text & "'"
        Cur_Info.CAS = frmmain!Data2.Recordset("CAS")
        TempCAS2 = Cur_Info.CAS
            
        Screen.MousePointer = 11
    
        Cur_Info.CAS = TempCAS1
        Call SaveUserData
        Cur_Info.CAS = TempCAS2
    
        frmmain!Data1.Recordset.FindFirst "CAS = " & Cur_Info.CAS
        TXTFamily.Text = GetFamilyGroup(frmmain!Data1.Recordset("Chemical Family"))
    
        'Store last CAS number viewed
        Set DBTbl = DBJetUser.OpenRecordset("Last CAS", dbOpenTable)
        DBTbl.Edit
        DBTbl("LastCAS") = Cur_Info.CAS
        DBTbl("ULCAS") = Cur_Info.CAS
        DBTbl("LastList") = 2
        DBTbl.Update
        DBTbl.Close

        Call Recalculate
        Call DisplayProps
    
        'Turn on all folders
        Call TabFolderEnable(True)
        frmmain.Refresh
        
        
no_calc:
        Screen.MousePointer = 1

    Else
        TXTOpT.Text = FormatVal(Cur_Disp.OpT)
    End If

    Exit Sub
    
NotValid:
    MsgBox "Not a valid number, please enter a smaller value", 0, "Value Not Valid"
    TXTOpT.Text = FormatVal(Cur_Disp.OpT)

End Sub


Private Sub TXTOpT_LostFocus()

    Dim Response As Integer
    Dim TempCAS1 As Long
    Dim TempCAS2 As Long
    Dim DBTbl As Recordset
    
    On Error GoTo NotValid
    
    If FormatVal(Val(TXTOpT.Text)) = FormatVal(Cur_Disp.OpT) Then
        TXTOpT.Text = FormatVal(Cur_Disp.OpT)
        Exit Sub
    End If
    
    'Ask user to recalculate at new temperature
    Response = MsgBox("Recalculate using new temperature?", 3, "Recalculate")
    If Response = 6 Then
        On Error GoTo no_calc   ' in case no current record
        Cur_Info.OpT = Convert(Val(TXTOpT.Text), OptTemp, Cur_Disp.OpTUnit, "K", False)
        
        frmmain!TABViewProp.CurrTab = 6
        frmmain.Refresh
        
        TempCAS1 = Cur_Info.CAS
        frmmain!Data2.Recordset.FindFirst "Name = '" & frmmain!LSTUserList.Text & "'"
        Cur_Info.CAS = frmmain!Data2.Recordset("CAS")
        TempCAS2 = Cur_Info.CAS
            
        Screen.MousePointer = 11
    
        Cur_Info.CAS = TempCAS1
        Call SaveUserData
        Cur_Info.CAS = TempCAS2
    
        frmmain!Data1.Recordset.FindFirst "CAS = " & Cur_Info.CAS
        TXTFamily.Text = GetFamilyGroup(frmmain!Data1.Recordset("Chemical Family"))
    
        'Store last CAS number viewed
        Set DBTbl = DBJetUser.OpenRecordset("Last CAS", dbOpenTable)
        DBTbl.Edit
        DBTbl("LastCAS") = Cur_Info.CAS
        DBTbl("ULCAS") = Cur_Info.CAS
        DBTbl("LastList") = 2
        DBTbl.Update
        DBTbl.Close

        Call Recalculate
        Call DisplayProps
    
        'Turn on all folders
        Call TabFolderEnable(True)
        frmmain.Refresh
        


no_calc:
        Screen.MousePointer = 1
        
    Else
        TXTOpT.Text = FormatVal(Cur_Disp.OpT)
    End If
    
    Exit Sub
    
NotValid:
    MsgBox "Not a valid number, please enter a smaller value", 0, "Value Not Valid"
    TXTOpT.Text = FormatVal(Cur_Disp.OpT)

End Sub

